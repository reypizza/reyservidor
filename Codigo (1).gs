/**
 * CAJA REY PIZZA — registro de gastos e ingresos
 * CAIX, S.A. · Centro · Almendras · Parque · Bodega Central
 *
 * La fecha y hora de registro las pone el servidor: el teléfono no puede
 * cambiarlas. Cada persona entra con su nombre y un PIN que se administra
 * desde la propia app, en el panel de administrador.
 *
 * PUBLICAR ASÍ:  Ejecutar como → YO (el dueño del proyecto)
 *                Quién tiene acceso → Cualquier usuario con una cuenta de Google
 * Con esa combinación nadie más necesita permisos sobre la hoja de cálculo.
 */

var ZONA = 'America/Guatemala';

/** La hoja de cálculo donde vive todo. */
var ID_HOJA = '1PH-Ww93TPPZfRbqk0mJXJ0_hOhgpA7j7b6lV_XLlM60';

var _SS = null;
/* ── lecturas de hoja sin repetir ──
 * Cada lectura a la hoja cuesta (decenas o cientos de milisegundos) y una misma pantalla leía la misma pestaña 6 o
 * 10 veces. Dentro de una sola ejecución, lo que ya se leyó se reutiliza hasta que alguien escriba en esa pestaña
 * (el envoltorio lo detecta solo) o se tome el candado (otra ejecución pudo haber escrito). */
var _LEE = {};
var _TOCADAS = {};
var _MEMO = {};                 // lo ya calculado en esta ejecución (pedidos, inventario, movimientos); se borra apenas algo se guarda              // las pestañas que esta ejecución cambió (se le avisa al servidor propio para que actualice su copia)
var ESCRIBE_HOJA = /^(append|set|clear|insert|delete|remove|copy|move|merge|sort|break|activate|check|uncheck|trim|randomize|create|add|protect|hide|show|resize|update|apply|split|flatten|fill|auto|unmerge|expand|collapse|group|ungroup|transpose|cut|paste)/i;
function olvidaLectura(nombre) { _MEMO = {}; if (nombre) { delete _LEE[nombre]; _TOCADAS[nombre] = 1; } else _LEE = {}; }
function envuelveRango(r, nombre) {
  if (typeof Proxy === 'undefined' || !r) return r;
  return new Proxy(r, { get: function (t, k) {
    var v = t[k];
    if (typeof v !== 'function') return v;
    if (typeof k === 'string' && ESCRIBE_HOJA.test(k)) return function () { olvidaLectura(nombre); return v.apply(t, arguments); };
    return function () { return v.apply(t, arguments); };
  } });
}
function envuelveHoja(h, nombre) {
  if (typeof Proxy === 'undefined' || !h) return h;
  return new Proxy(h, { get: function (t, k) {
    if (k === '__nombre') return nombre;
    var v = t[k];
    if (typeof v !== 'function') return v;
    if (typeof k === 'string' && /^get(Range|DataRange|RangeList|ActiveRange)/.test(k))
      return function () { return envuelveRango(v.apply(t, arguments), nombre); };
    if (typeof k === 'string' && ESCRIBE_HOJA.test(k)) return function () { olvidaLectura(nombre); return v.apply(t, arguments); };
    if (k === 'getLastColumn' || k === 'getLastRow' || k === 'getMaxColumns') return function () {
      var c = (_LEE[nombre] = _LEE[nombre] || {});
      if (c['#' + k] === undefined) c['#' + k] = v.call(t);
      return c['#' + k];
    };
    return function () { return v.apply(t, arguments); };
  } });
}
function envuelveLibro(ss) {
  if (typeof Proxy === 'undefined' || !ss) return ss;
  var hojas = {};
  return new Proxy(ss, { get: function (t, k) {
    var v = t[k];
    if (typeof v !== 'function') return v;
    if (k === 'getSheetByName') return function (n) {
      if (hojas[n]) return hojas[n];
      var h = v.call(t, n);
      return h ? (hojas[n] = envuelveHoja(h, String(n))) : h;
    };
    if (typeof k === 'string' && /^(insertSheet|deleteSheet|setActiveSheet|moveActiveSheet|duplicateActiveSheet|copy)/.test(k))
      return function () {
        hojas = {}; olvidaLectura(); _TOCADAS['*'] = 1;       // cambió la lista de pestañas
        var r = v.apply(t, arguments);
        try { if (r && typeof r.getName === 'function') return (hojas[r.getName()] = envuelveHoja(r, String(r.getName()))); } catch (e) {}
        return r;
      };
    return function () { return v.apply(t, arguments); };
  } });
}
function libro() {
  if (_SS) return _SS;                       // abrir la hoja cuesta casi un segundo
  _SS = envuelveLibro(ID_HOJA ? SpreadsheetApp.openById(ID_HOJA)
                : SpreadsheetApp.getActiveSpreadsheet());
  extiendeLugares();                         // los lugares creados en la app se suman a las sucursales
  return _SS;
}

var _USR = null;
/* Cada llamada del teléfono revisa nombre y PIN. Leer la hoja Usuarios cada vez
 * costaba varias idas a la hoja; ahora la lista vive 5 minutos en la memoria del
 * servidor y se borra en cuanto alguien cambia un usuario o un PIN desde la app.
 * (Si se cambia un PIN escribiendo directo en la hoja, tarda hasta 5 minutos.) */
var LLAVE_USR = 'usuarios_v2';
function usuariosCache() {
  if (_USR) return _USR;
  try {
    var c = CacheService.getScriptCache().get(LLAVE_USR);
    if (c) { _USR = JSON.parse(c); return _USR; }
  } catch (e) {}
  _USR = filasUsuarios();
  try { CacheService.getScriptCache().put(LLAVE_USR, JSON.stringify(_USR), 300); } catch (e) {}
  return _USR;
}
function olvidaUsuarios() {
  _USR = null;
  try { CacheService.getScriptCache().remove(LLAVE_USR); } catch (e) {}
}

/* ════════════ CATÁLOGOS ════════════ */

var UNIDADES = [
  { id: 'centro',    nombre: 'Centro',         marca: 'Rey Pizza',           vende: true,  tipo: 'sucursal', conInv: true },
  { id: 'almendras', nombre: 'Almendras',      marca: 'Rey Pizza I',         vende: true,  tipo: 'sucursal', conInv: true },
  { id: 'parque',    nombre: 'Parque',         marca: 'Rey Pizza II',        vende: true,  tipo: 'sucursal', conInv: true },
  { id: 'bodega',    nombre: 'Bodega Central', marca: 'Bodega y producción', vende: false, tipo: 'bodega',   conInv: true }
];
/* ── Lugares: además de las tres sucursales y la bodega, se pueden crear otros (oficinas de contabilidad
 *  y publicidad, por ejemplo). No venden, pero llevan sus propios gastos y presupuesto, y todo llega al
 *  mismo estado de resultados (columna propia y el total de CAIX, S.A.). ── */
var H_LUGARES = ['Id', 'Nombre', 'Descripción', 'Tipo', 'Activo', 'Creado en', 'Creado por'];
var TIPOS_LUGAR = { oficina: 'Oficinas', bodega: 'Bodega o almacén', otro: 'Otro lugar' };
var _LUG = false;
function leeLugares() {
  var h = _SS ? _SS.getSheetByName('Lugares') : null;
  if (!h) return [];
  return leeTodo(h, H_LUGARES.length).map(function (r) {
    var tipo = TIPOS_LUGAR[String(r[3]).trim()] ? String(r[3]).trim() : 'otro';
    return { id: String(r[0] || '').trim().toLowerCase(), nombre: String(r[1] || '').trim(),
      marca: String(r[2] || '').trim() || TIPOS_LUGAR[tipo], tipo: tipo, activo: String(r[4]).trim().toLowerCase() !== 'no' };
  }).filter(function (x) { return x.id && x.nombre; });
}
function extiendeLugares() {
  if (_LUG) return; _LUG = true;
  var extra = null, c = null, fallo = false;
  try { c = CacheService.getScriptCache(); var t = c.get('lugares_v1'); if (t) extra = JSON.parse(t); } catch (e) { c = null; }
  if (!extra) {
    extra = [];
    try { if (!_SS) libro(); extra = leeLugares(); } catch (e) { extra = []; fallo = true; _LUG = false; }
    try { if (c && !fallo) c.put('lugares_v1', JSON.stringify(extra), 600); } catch (e) {}   // un fallo nunca se guarda como «no hay lugares»
  }
  extra.forEach(function (x) {
    if (unidadPorId(x.id)) return;
    UNIDADES.push({ id: x.id, nombre: x.nombre, marca: x.marca, vende: false, tipo: x.tipo, conInv: false, activo: x.activo, extra: true });
  });
}
function recargaLugares() {
  for (var k = UNIDADES.length - 1; k >= 0; k--) if (UNIDADES[k].extra) UNIDADES.splice(k, 1);
  try { CacheService.getScriptCache().remove('lugares_v1'); } catch (e) {}
  _LUG = false; extiendeLugares();
}

var CATEGORIAS = [
  'Nómina y bonificaciones', 'IGSS y prestaciones', 'Insumos de cocina',
  'Bebidas y gaseosas', 'Empaque y desechables', 'Gas propano',
  'Energía eléctrica', 'Agua', 'Renta', 'Internet y teléfono',
  'Limpieza y químicos', 'Mantenimiento y reparaciones',
  'Combustible y transporte', 'Mercadeo y publicidad',
  'Comisiones de plataformas', 'Papelería y administración',
  'Impuestos y tasas', 'Otros gastos', 'Sistema punto de venta (Paladar POS)',
  'Equipo y mobiliario'
];

/* Las columnas A..N son las de siempre. Las tres últimas (O, P, Q) se agregaron
 * después, al final a propósito: así ninguna fórmula ni ninguna hoja que ya
 * estaba hecha se corre de lugar. */
var H_GASTOS = ['ID', 'Registrado en', 'Registrado por', 'Correo', 'Unidad',
  'Categoría', 'Fecha de pago', 'Mes', 'Monto', 'Nota', 'Anulado',
  'Motivo de anulación', 'Anulado por', 'Anulado en',
  'Proveedor', 'No. de factura', 'Artículos', 'Confirmado por', 'Confirmado en'];

/** Un renglón por artículo. El total de la factura vive en la hoja Gastos; aquí
 *  vive el desglose, para poder ver qué se le compra a cada proveedor. */
/* A..N son las de siempre; O..R (la presentación) se agregaron después, al final,
 * para no correr nada de lugar. */
var H_DETALLE = ['ID gasto', 'Registrado en', 'Registrado por', 'Unidad',
  'Proveedor', 'Fecha de pago', 'Mes', 'Categoría', 'Producto', 'Cantidad',
  'Medida', 'Precio unitario', 'Subtotal', 'Anulado',
  'Contenido', 'Unidad base', 'Cantidad base', 'Costo por unidad base'];

/* I y J (Contenido y Unidad base) también van al final a propósito. */
var H_PRODUCTOS = ['Producto', 'Proveedor', 'Medida', 'Categoría',
  'Último precio', 'Activo', 'Creado en', 'Creado por',
  'Contenido', 'Unidad base'];

/* Tipo: vacío = general (bodega y sucursales); «Sucursal» = solo de sucursales (la
 * bodega no lo ve); «Repartidor» = delivery que se paga cada lunes con lo cobrado
 * en envíos (tampoco lo ve la bodega; lo paga contabilidad). */
var H_PROVEEDORES = ['Proveedor', 'NIT', 'Teléfono', 'Activo', 'Creado en',
  'Creado por', 'Tipo', 'Sucursal'];
var TIPOS_PROV = ['', 'Sucursal', 'Repartidor'];
/* Lo que se cobra de envío en cada ticket y se le entrega al repartidor: no es venta
 * de CAIX. En el estado de resultados se resta de las ventas, no es un gasto. */
var CAT_ENVIOS = 'Envíos entregados a repartidores';

/** Cómo lo vende el proveedor. */
var MEDIDAS = ['unidad', 'caja', 'bolsa', 'saco', 'fardo', 'quintal', 'libra',
  'kilo', 'litro', 'galón', 'botella', 'docena', 'cubeta', 'tonel', 'servicio'];

/** En qué se mide de verdad lo que trae adentro. */
var BASES = ['libra', 'onza', 'kilo', 'gramo', 'unidad', 'litro', 'mililitro',
  'galón', 'docena', 'metro', 'rollo'];

var VARIAS = 'Varias categorías';

/** Cómo se agrupan las 18 categorías para el estado de resultados.
 *  Esto es solo el arranque: la hoja «Clasificación» manda, y ahí se puede
 *  mover una categoría de grupo sin tocar el programa. */
var GRUPOS = [
  { grupo: 'Costo de ventas', renglon: 'Materia prima y empaque',
    categorias: ['Insumos de cocina', 'Bebidas y gaseosas',
      'Empaque y desechables', 'Gas propano', 'Traslado a sucursales'] },
  { grupo: 'Gastos de operación', renglon: 'Personal',
    categorias: ['Nómina y bonificaciones', 'IGSS y prestaciones'] },
  { grupo: 'Gastos de operación', renglon: 'Local y servicios',
    categorias: ['Renta', 'Energía eléctrica', 'Agua', 'Internet y teléfono',
      'Sistema punto de venta (Paladar POS)'] },
  { grupo: 'Gastos de operación', renglon: 'Operación y ventas',
    categorias: ['Limpieza y químicos', 'Mantenimiento y reparaciones',
      'Combustible y transporte', 'Mercadeo y publicidad',
      'Comisiones de plataformas', 'Gastos de caja', 'Comisiones de tarjeta (NEONET)'] },
  { grupo: 'Gastos de operación', renglon: 'Administración',
    categorias: ['Papelería y administración', 'Impuestos y tasas',
      'Equipo y mobiliario', 'Otros gastos'] }
];

var H_CLASIF = ['Categoría', 'Grupo', 'Renglón'];

/** Un renglón por mes. Se actualiza cada vez que se vuelve a armar el PDF de ese
 *  mes: así el enlace nunca queda muerto y no se acumulan archivos repetidos. */
var H_REPORTES = ['Mes', 'Última vez generado', 'Generado por', 'Movimientos',
  'Archivo', 'Enlace', 'ID de Drive', 'Veces'];
var G_COSTO = 'Costo de ventas';
var G_OPER  = 'Gastos de operación';

/* A..N son las de siempre. O..R (el desglose del corte) van al final. */
var H_INGRESOS = ['ID', 'Registrado en', 'Registrado por', 'Correo', 'Sucursal',
  'Fecha del ingreso', 'Mes', 'Turno', 'Monto', 'Nota', 'Anulado',
  'Motivo de anulación', 'Anulado por', 'Anulado en',
  'Efectivo', 'Tarjeta', 'Gastos de caja', 'En qué se gastó',
  'Confirmado por', 'Confirmado en', 'Validado por', 'Validado en',
  'Recibido por', 'Recibido en', 'Comisión tarjeta', 'Tarjeta líquida', '% comisión'];

/** Lo que sale de la caja durante el turno. No es una categoría que se pueda
 *  escoger al registrar un gasto —se captura en el corte— pero sí entra al
 *  estado de resultados, y por eso vive en la hoja «Clasificación». */
var CAT_CAJA = 'Gastos de caja';
/** Lo que NEONET retiene de cada pago con tarjeta. Tampoco se escoge a mano: sale
 *  del corte (tarjeta × %) y entra al estado de resultados como gasto. */
var CAT_COMISION = 'Comisiones de tarjeta (NEONET)';
function tasaComision() { return ajustesInv().comision; }
function r2(n) { return Math.round((Number(n) || 0) * 100) / 100; }
function comisionDe(tarjeta, pct) { return r2((Number(tarjeta) || 0) * (pct == null ? tasaComision() : pct) / 100); }
/** La comisión guardada en la fila, o calculada si la fila es de antes. */
function comisionFila(s) {
  if (s[24] !== '' && s[24] != null && isFinite(Number(s[24]))) return Number(s[24]) || 0;
  return comisionDe(Number(s[15]) || 0);
}

/* ── TRASLADOS DE BODEGA ──
 * Samuel anota cuánto de cada producto mandó a cada sucursal. Cada día y cada
 * sucursal queda como UNA factura de «Bodega Central» en esa sucursal, al precio
 * de compra, sin recargo: así el producto se vuelve costo de quien lo recibe.
 * Para no contarlo dos veces, la bodega lleva la misma factura en negativo con la
 * categoría «Traslado a sucursales»: lo que compró menos lo que despachó. */
var PROV_BODEGA = 'Bodega Central';
var CAT_TRASLADO = 'Traslado a sucursales';
var H_CORRECCIONES = ['Fecha del traslado', 'Sucursal', 'Producto', 'Unidad',
  'Antes', 'Después', 'Diferencia', 'Motivo', 'Corregido por', 'Corregido en',
  'ID traslado', 'No. de traslado'];

/* G y H se agregaron después: la sucursal del gerente y si puede confirmar lo que
 * envían los gerentes. */
var H_USUARIOS = ['Nombre', 'PIN', 'Rol', 'Activo', 'Creado en', 'Creado por',
  'Sucursal', 'Confirma registros', 'Correo'];
/* operaciones = director operativo (valida los cierres de caja y define los
 * indicadores de cada sucursal); finanzas = supervisa los ingresos y confirma que
 * el dinero llegó. Los dos solo miran las cifras: no registran ni anulan. */
var ROLES = ['admin', 'registro', 'gerente', 'bodega', 'operaciones', 'finanzas', 'produccion', 'dueno', 'auxiliar', 'cmo'];
/** Una persona de acceso limitado (auxiliar de un lugar) solo pasa por `quien` cuando la función
 *  pidió permitirla con `permiteAux()`: entrar, datos, registrar un gasto, su PIN, su correo y avisos.
 *  Todo lo demás la rechaza, llegue como llegue la llamada. */
var _AUXOK = false;
function permiteAux() { _AUXOK = true; }
/** La CMO (responsable de las oficinas) tiene lo mismo que el acceso limitado y además su presupuesto, sus
 *  solicitudes de pago, sus indicadores, permisos, preguntas y sugerencias: esas funciones piden `permiteCmo()`. */
var _CMOOK = false;
function permiteCmo() { _CMOOK = true; }
var NOMBRE_ROL = { admin: 'Administra', registro: 'Registra', gerente: 'Gerente',
  bodega: 'Bodega', operaciones: 'Director operativo', finanzas: 'Finanzas', produccion: 'Producción',
  dueno: 'Dueños (solo reportes)', auxiliar: 'Acceso limitado', cmo: 'CMO · Chief Marketing Officer' };

var PIN_INICIAL = '1234';

/** Quiénes entran al principio. Se usa en instalar() y en usuariosDeArranque(). */
var USUARIOS_INICIALES = [
  ['Marcelino', '1234', 'admin']
];

/* ════════════ INSTALACIÓN ════════════ */

/** Ejecute esta función UNA VEZ desde el editor. Volver a correrla NO borra
 *  los movimientos ni los usuarios: solo rehace lo que falte. */
function instalar() {
  var ss = libro();
  ss.setSpreadsheetTimeZone(ZONA);

  if (!ss.getSheetByName('Gastos')) {
    var g = hojaLimpia(ss, 'Gastos', H_GASTOS);
    g.getRange('B:B').setNumberFormat('dd/mm/yyyy hh:mm');
    g.getRange('G:G').setNumberFormat('dd/mm/yyyy');
    g.getRange('H2:H').setNumberFormat('@');      // el mes es texto, no fecha
    g.getRange('I:I').setNumberFormat('"Q"#,##0.00');
  } else {
    // la hoja ya existía: solo se le agregan los encabezados nuevos al final
    encabezaAlFinal(ss.getSheetByName('Gastos'), 15,
      ['Proveedor', 'No. de factura', 'Artículos']);
  }
  encabezaAlFinal(ss.getSheetByName('Gastos'), 18, ['Confirmado por', 'Confirmado en']);
  ss.getSheetByName('Gastos').getRange('S:S').setNumberFormat('dd/mm/yyyy hh:mm');

  if (!ss.getSheetByName('Gastos detalle')) {
    var dd = hojaLimpia(ss, 'Gastos detalle', H_DETALLE);
    dd.getRange('B:B').setNumberFormat('dd/mm/yyyy hh:mm');
    dd.getRange('F:F').setNumberFormat('dd/mm/yyyy');
    dd.getRange('G2:G').setNumberFormat('@');     // el mes es texto, no fecha
    dd.getRange('L:M').setNumberFormat('"Q"#,##0.00');
    dd.getRange('R:R').setNumberFormat('"Q"#,##0.0000');
    dd.setColumnWidth(9, 230);
  } else {
    encabezaAlFinal(ss.getSheetByName('Gastos detalle'), 15,
      ['Contenido', 'Unidad base', 'Cantidad base', 'Costo por unidad base']);
    ss.getSheetByName('Gastos detalle').getRange('R:R')
      .setNumberFormat('"Q"#,##0.0000');
  }
  if (!ss.getSheetByName('Proveedores')) {
    var pv = hojaLimpia(ss, 'Proveedores', H_PROVEEDORES);
    pv.getRange('E:E').setNumberFormat('dd/mm/yyyy hh:mm');
    pv.setColumnWidth(1, 220);
  }
  if (!ss.getSheetByName('Productos')) {
    var pr = hojaLimpia(ss, 'Productos', H_PRODUCTOS);
    pr.getRange('E:E').setNumberFormat('"Q"#,##0.00');
    pr.getRange('G:G').setNumberFormat('dd/mm/yyyy hh:mm');
    pr.setColumnWidth(1, 230); pr.setColumnWidth(2, 190);
  } else {
    encabezaAlFinal(ss.getSheetByName('Productos'), 9,
      ['Contenido', 'Unidad base']);
  }

  if (!ss.getSheetByName('Ingresos')) {
    var i = hojaLimpia(ss, 'Ingresos', H_INGRESOS);
    i.getRange('B:B').setNumberFormat('dd/mm/yyyy hh:mm');
    i.getRange('F:F').setNumberFormat('dd/mm/yyyy');
    i.getRange('G2:G').setNumberFormat('@');      // el mes es texto, no fecha
    i.getRange('I:I').setNumberFormat('"Q"#,##0.00');
    i.getRange('O:Q').setNumberFormat('"Q"#,##0.00');
  } else {
    encabezaAlFinal(ss.getSheetByName('Ingresos'), 15,
      ['Efectivo', 'Tarjeta', 'Gastos de caja', 'En qué se gastó']);
    ss.getSheetByName('Ingresos').getRange('O:Q')
      .setNumberFormat('"Q"#,##0.00');
  }
  encabezaAlFinal(ss.getSheetByName('Ingresos'), 19, ['Confirmado por', 'Confirmado en']);
  ss.getSheetByName('Ingresos').getRange('T:T').setNumberFormat('dd/mm/yyyy hh:mm');
  encabezaAlFinal(ss.getSheetByName('Ingresos'), 21, ['Validado por', 'Validado en']);
  ss.getSheetByName('Ingresos').getRange('V:V').setNumberFormat('dd/mm/yyyy hh:mm');
  encabezaAlFinal(ss.getSheetByName('Ingresos'), 23, ['Recibido por', 'Recibido en',
    'Comisión tarjeta', 'Tarjeta líquida', '% comisión']);
  ss.getSheetByName('Ingresos').getRange('X:X').setNumberFormat('dd/mm/yyyy hh:mm');
  ss.getSheetByName('Ingresos').getRange('Y:Z').setNumberFormat('"Q"#,##0.00');
  comisionesViejas(ss);
  if (!ss.getSheetByName('Usuarios')) escribeUsuarios(ss);
  encabezaAlFinal(ss.getSheetByName('Usuarios'), 7, ['Sucursal', 'Confirma registros', 'Correo']);

  UNIDADES.forEach(function (un) {
    var h = hoja(ss, un.nombre);
    h.clear();
    h.getRange('A1').setValue(un.nombre + ' · ' + un.marca)
      .setFontSize(14).setFontWeight('bold');
    h.getRange('A2').setValue('Esta hoja se llena sola. No escriba en ella.')
      .setFontColor('#6c7079').setFontSize(9);
    h.getRange('A4').setValue('GASTOS').setFontWeight('bold');
    h.getRange('A5').setFormula(
      '=IFERROR(QUERY(Gastos!A:Q,"select G,O,F,I,Q,J,C,K where E = \'' +
      un.nombre + '\' order by G desc label G \'Fecha de pago\', O \'Proveedor\',' +
      ' F \'Categoría\', I \'Total\', Q \'Artículos\', J \'Nota\',' +
      ' C \'Registró\', K \'Anulado\'",1),"Sin gastos todavía")');
    if (un.vende) {
      h.getRange('J4').setValue('INGRESOS').setFontWeight('bold');
      h.getRange('J5').setFormula(
        '=IFERROR(QUERY(Ingresos!A:R,"select F,H,O,P,I,Q,C,K where E = \'' +
        un.nombre + '\' order by F desc label F \'Día\', H \'Turno\',' +
        ' O \'Efectivo\', P \'Tarjeta\', I \'Venta\', Q \'Gastos de caja\',' +
        ' C \'Registró\', K \'Anulado\'",1),"Sin ingresos todavía")');
    }
    h.setColumnWidth(1, 105); h.setColumnWidth(2, 190); h.setColumnWidth(3, 170);
    h.setColumnWidth(6, 210);
  });

  clasificacionHoja(ss);
  if (!ss.getSheetByName('Reportes')) {
    var hr = hojaLimpia(ss, 'Reportes', H_REPORTES);
    hr.getRange('A2:A').setNumberFormat('@');
    hr.getRange('B:B').setNumberFormat('dd/mm/yyyy hh:mm');
    hr.setColumnWidth(1, 90); hr.setColumnWidth(2, 150);
    hr.setColumnWidth(5, 230); hr.setColumnWidth(6, 300);
  }
  if (!ss.getSheetByName('Traslados correcciones')) hojaCorrecciones(ss);
  hojaBodega(ss);                      // solo la crea si no existe
  hojaTraslados(ss);
  hojaCorrecciones(ss);
  var hpc = hojaPorConfirmar(ss);
  if (hpc.getLastColumn() < 19)
    encabezaAlFinal(hpc, 16, ['Validado por', 'Validado en', 'Efectivo recibido',
      'Diferencia de efectivo']);
  if (hpc.getLastColumn() < 24)
    encabezaAlFinal(hpc, 20, ['Recibido por', 'Recibido en', 'Tarjeta NEONET',
      'Comisión tarjeta', 'Tarjeta líquida']);
  ajustesQueFaltan();
  var hpv = hojaCat('Proveedores', H_PROVEEDORES);
  if (hpv.getLastColumn() < 8) encabezaAlFinal(hpv, 7, ['Tipo', 'Sucursal']);
  marcaRepartidores();
  hojasImpuestos(ss);
  obligacionesDelRegimen(ss);
  hojasProduccion(ss);
  hojaOCP(ss);
  encabezaAlFinal(hojaOCP(), 16, ['Cotización (archivo)']);
  hojasMuestreo(ss);
  hojaPres(ss); hojaFondos(ss); hojaEmerg(ss); hojaPerm(ss);
  encabezaAlFinal(hojaFondos(), 23, ['Cotización (archivo)', 'Nombre del archivo']);
  hojaPresEstado(ss); hojaConsultas(ss); hojaCorreos(ss);
  instalaDisparadorCorreos();
  var hoc = hojaOC();
  if (hoc.getLastColumn() < 17) encabezaAlFinal(hoc, 16, ['Precio (cotización)', 'Subtotal']);
  var hk = hojasIndicadores(ss).lista;
  if (hk.getLastColumn() < 21) encabezaAlFinal(hk, 21, ['Contabilidad']);
  if (hk.getLastColumn() < 22) encabezaAlFinal(hk, 22, ['Marketing']);
  if (hk.getLastColumn() < 23) encabezaAlFinal(hk, 23, ['Tecnología']);
  if (hk.getLastColumn() < 24) encabezaAlFinal(hk, 24, ['Dirección operativa']);
  hojasInventario(ss);
  hojasIndicadores(ss);
  rellenaArticulos(ss);
  efectivoSinDesglose(ss);
  arreglaMeses(ss);
  resumen(ss);
  compras(ss);
  estadoResultados(ss);
  return 'Listo. Hojas: Gastos, Gastos detalle, Ingresos, Productos, ' +
    'Proveedores, Clasificación, Resumen, Compras por producto, ' +
    'Estado de resultados, Traslados, Traslados correcciones, Productos de bodega, ' +
    'Por confirmar, Inventario (conteos, movimientos, solicitudes, órdenes de compra, ajustes). Usuario inicial: ' + USUARIOS_INICIALES[0][0] +
    ', PIN ' + PIN_INICIAL + '. Hoja: ' + ss.getUrl();
}

/** Las facturas viejas no traían la cuenta de artículos. Se les pone 0 para que
 *  el estado de resultados las pueda separar de las que sí tienen desglose. */
function rellenaArticulos(ss) {
  var h = ss.getSheetByName('Gastos');
  if (!h || h.getLastRow() < 2) return;
  var n = h.getLastRow() - 1;
  var r = h.getRange(2, 17, n, 1);
  var v = r.getValues(), toco = false;
  for (var k = 0; k < v.length; k++) {
    if (v[k][0] === '' || v[k][0] === null) { v[k][0] = 0; toco = true; }
  }
  if (toco) r.setValues(v);
}

/** Los cortes de antes solo traían el total. Se toman como venta en efectivo:
 *  Efectivo = Venta y Tarjeta = 0. Solo toca los que tienen las dos vacías. */
function efectivoSinDesglose(ss) {
  var h = ss.getSheetByName('Ingresos');
  if (!h || h.getLastRow() < 2 || h.getMaxColumns() < 16) return 0;
  var n = h.getLastRow() - 1;
  var monto = h.getRange(2, 9, n, 1).getValues();
  var r = h.getRange(2, 15, n, 2), v = r.getValues(), cuantos = 0;
  for (var k = 0; k < n; k++) {
    var e = v[k][0], t = v[k][1], m = Number(monto[k][0]) || 0;
    if ((e === '' || e === null || Number(e) === 0) &&
        (t === '' || t === null || Number(t) === 0) && m > 0) {
      v[k][0] = m; v[k][1] = 0; cuantos++;
    }
  }
  if (cuantos) r.setValues(v);
  return cuantos;
}

/** LA COLUMNA «MES» TIENE QUE SER TEXTO.
 *  Al escribir «2026-09» en una celda sin formato, Google la entiende como una
 *  fecha (1 de septiembre de 2026) y la guarda así. Entonces ninguna fórmula que
 *  compare contra el texto «2026-09» encuentra nada: ni el Resumen, ni Compras
 *  por producto, ni el Estado de resultados, ni el desglose del reporte.
 *  Esta función pone la columna en formato texto y la vuelve a escribir a partir
 *  de la fecha que está al lado. Se puede correr las veces que sea. */
function arreglaMeses(ss) {
  [['Gastos', 7, 8], ['Gastos detalle', 6, 7], ['Ingresos', 6, 7]]
    .forEach(function (x) {
      var h = ss.getSheetByName(x[0]);
      if (!h) return;
      h.getRange(2, x[2], h.getMaxRows() - 1, 1).setNumberFormat('@');
      if (h.getLastRow() < 2) return;
      var n = h.getLastRow() - 1;
      var fechas = h.getRange(2, x[1], n, 1).getValues();
      h.getRange(2, x[2], n, 1).setValues(fechas.map(function (r) {
        return [fmtDia(r[0]).slice(0, 7)];
      }));
    });
}

/** La hoja donde se decide qué es costo de ventas y qué es gasto de operación.
 *  Si ya existe no se toca, solo se le agregan las categorías que falten. */
function clasificacionHoja(ss) {
  var h = ss.getSheetByName('Clasificación');
  if (!h) {
    h = hojaLimpia(ss, 'Clasificación', H_CLASIF);
    h.getRange('A1').setValue('Categoría');
    h.setColumnWidth(1, 240); h.setColumnWidth(2, 170); h.setColumnWidth(3, 210);
    var filas = [];
    GRUPOS.forEach(function (g) {
      g.categorias.forEach(function (c) { filas.push([c, g.grupo, g.renglon]); });
    });
    h.getRange(2, 1, filas.length, 3).setValues(filas);
    h.getRange(2, 2, filas.length, 1).setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireValueInList([G_COSTO, G_OPER], true).build());
    h.getRange(1, 5).setValue('Mueva una categoría de grupo aquí y el estado de ' +
      'resultados cambia solo. No borre renglones.')
      .setFontColor('#6c7079').setFontSize(9);
    return h;
  }
  var hay = {};
  leeTodo(h, 3).forEach(function (r) {
    if (String(r[0]).trim()) hay[String(r[0]).trim()] = true;
  });
  var todas = CATEGORIAS.slice();
  GRUPOS.forEach(function (x) {
    x.categorias.forEach(function (c) { if (todas.indexOf(c) < 0) todas.push(c); });
  });
  todas.forEach(function (c) {
    if (hay[c]) return;
    hay[c] = true;
    var g = null;
    GRUPOS.forEach(function (x) {
      if (x.categorias.indexOf(c) >= 0) g = x;
    });
    h.appendRow([c, g ? g.grupo : G_OPER, g ? g.renglon : 'Administración']);
  });
  return h;
}

/** Devuelve la clasificación viva: primero lo que diga la hoja, y lo que falte
 *  se completa con GRUPOS. */
function clasifica() {
  var mapa = {}, orden = [], idx = {};

  function mete(grupo, renglon, cat) {
    if (mapa[cat]) return;
    mapa[cat] = { grupo: grupo, renglon: renglon };
    var llave = grupo + '§' + renglon;
    if (!idx[llave]) {
      idx[llave] = { grupo: grupo, renglon: renglon, categorias: [] };
      orden.push(idx[llave]);
    }
    idx[llave].categorias.push(cat);
  }

  try {
    leeTodo(libro().getSheetByName('Clasificación'), 3).forEach(function (r) {
      var cat = String(r[0] || '').trim();
      if (!cat) return;
      var grupo = String(r[1] || '').trim();
      if (grupo !== G_COSTO) grupo = G_OPER;
      mete(grupo, String(r[2] || '').trim() || 'Otros', cat);
    });
  } catch (e) { /* si no hay hoja, mandan las constantes */ }

  GRUPOS.forEach(function (g) {
    g.categorias.forEach(function (c) { mete(g.grupo, g.renglon, c); });
  });
  CATEGORIAS.forEach(function (c) { mete(G_OPER, 'Administración', c); });

  // primero el costo de ventas, después la operación
  orden.sort(function (a, b) {
    if (a.grupo !== b.grupo) return a.grupo === G_COSTO ? -1 : 1;
    return 0;
  });
  return { de: mapa, orden: orden };
}

/** Las columnas del estado de resultados: cada sucursal, el consolidado de
 *  Rey Pizza, los centros de costo y el total de la sociedad. */
function columnasER() {
  extiendeLugares();
  var suc = [], otras = [];
  UNIDADES.forEach(function (u) {
    (u.vende ? suc : otras).push(u.nombre);
  });
  var cols = suc.map(function (n) { return { t: n, u: [n] }; });
  cols.push({ t: 'REY PIZZA', u: suc, fuerte: true });
  otras.forEach(function (n) { cols.push({ t: n, u: [n] }); });
  cols.push({ t: 'CAIX, S.A.', u: suc.concat(otras), fuerte: true });
  return cols;
}

/** Qué se compró, de qué proveedor y en qué se gastó. Se llena sola con lo que
 *  se registra artículo por artículo en la app. */
function compras(ss) {
  var h = hoja(ss, 'Compras por producto');
  h.clear();
  h.getRange('A1').setValue('Compras por producto y proveedor')
    .setFontSize(16).setFontWeight('bold').setFontColor('#14284B');
  h.getRange('A2').setValue('Sale del desglose de cada factura. No escriba aquí.')
    .setFontColor('#6c7079').setFontSize(9);

  h.getRange('A4').setValue('Mes (escriba así: 2026-09, o deje TODO)')
    .setFontWeight('bold');
  h.getRange('D4').setValue('TODO').setBackground('#fff2cc')
    .setFontColor('#0000ff').setHorizontalAlignment('center');

  // lo que la bodega manda a las sucursales no es una compra: no entra aquí
  var filtro = '(N = \'No\' and E <> \'' + PROV_BODEGA + '\')" & IF($D$4="TODO","", " and G = \'" & $D$4 & "\'")';

  h.getRange('A6').setValue('POR PRODUCTO').setFontWeight('bold');
  h.getRange('B6').setValue('«Contenido» son las libras, litros o unidades ' +
    'que de verdad entraron.').setFontColor('#6c7079').setFontSize(9);
  h.getRange('A7').setFormula(
    '=IFERROR(QUERY(\'Gastos detalle\'!A:R, "select I, sum(J), sum(Q), sum(M) where ' +
    filtro + ' & " group by I order by sum(M) desc ' +
    'label I \'Producto\', sum(J) \'Comprado\', sum(Q) \'Contenido\', ' +
    'sum(M) \'Total\'", 1), "Todavía no hay facturas con desglose")');

  h.getRange('F6').setValue('POR PROVEEDOR').setFontWeight('bold');
  h.getRange('F7').setFormula(
    '=IFERROR(QUERY(\'Gastos detalle\'!A:R, "select E, sum(M) where ' +
    filtro + ' & " group by E order by sum(M) desc ' +
    'label E \'Proveedor\', sum(M) \'Total\'", 1), "Sin datos")');

  h.getRange('I6').setValue('POR CATEGORÍA').setFontWeight('bold');
  h.getRange('I7').setFormula(
    '=IFERROR(QUERY(\'Gastos detalle\'!A:R, "select H, sum(M) where ' +
    filtro + ' & " group by H order by sum(M) desc ' +
    'label H \'Categoría\', sum(M) \'Total\'", 1), "Sin datos")');

  h.getRange('L6').setValue('COSTO POR UNIDAD BASE').setFontWeight('bold');
  h.getRange('L7').setFormula(
    '=IFERROR(QUERY(\'Gastos detalle\'!A:R, "select I, P, avg(R), min(R), max(R) ' +
    'where " & IF($D$4="TODO","","G = \'" & $D$4 & "\' and ") & ' +
    '"N = \'No\' and E <> \'' + PROV_BODEGA + '\' and R > 0 group by I, P order by I ' +
    'label I \'Producto\', P \'Por\', avg(R) \'Promedio\', min(R) \'Más barato\', ' +
    'max(R) \'Más caro\'", 1), "Sin presentaciones todavía")');

  h.setColumnWidth(1, 250); h.setColumnWidth(2, 95); h.setColumnWidth(3, 95);
  h.setColumnWidth(4, 110); h.setColumnWidth(5, 24);
  h.setColumnWidth(6, 210); h.setColumnWidth(7, 120); h.setColumnWidth(8, 24);
  h.setColumnWidth(9, 210); h.setColumnWidth(10, 120); h.setColumnWidth(11, 24);
  h.setColumnWidth(12, 230); h.setColumnWidth(13, 80);
  [14, 15, 16].forEach(function (c) { h.setColumnWidth(c, 105); });
  h.getRange('D:D').setNumberFormat('"Q"#,##0.00');
  h.getRange('G:G').setNumberFormat('"Q"#,##0.00');
  h.getRange('J:J').setNumberFormat('"Q"#,##0.00');
  h.getRange('N:P').setNumberFormat('"Q"#,##0.0000');
}

function hoja(ss, nombre) {
  return ss.getSheetByName(nombre) || ss.insertSheet(nombre);
}

/** Agrega encabezados nuevos al final de una hoja que ya tiene datos, sin mover
 *  ninguna columna de lugar. Se puede correr las veces que sea. */
function encabezaAlFinal(h, desde, titulos) {
  if (!h) return;
  var hasta = desde + titulos.length - 1;
  if (h.getMaxColumns() < hasta) h.insertColumnsAfter(h.getMaxColumns(),
    hasta - h.getMaxColumns());
  h.getRange(1, desde, 1, titulos.length).setValues([titulos])
    .setFontWeight('bold').setFontColor('#ffffff').setBackground('#14284B')
    .setFontSize(10);
}

/** Lee una hoja completa sin pasarse de las columnas que tiene de verdad. */
function leeTodo(h, ancho) {
  var nom = h ? h.__nombre : null, c = nom ? (_LEE[nom] = _LEE[nom] || {}) : null;
  var todo = c && c['#raw'] ? c['#raw'] : null;           // la hoja se lee UNA vez por ejecución, sin importar cuántas columnas pida cada quien
  if (!todo) {
    todo = h ? h.getDataRange().getValues() : [];
    if (c) c['#raw'] = todo;
  }
  var out = [];
  for (var i = 1; i < todo.length; i++) {
    var f = todo[i];
    f = f.length > ancho ? f.slice(0, ancho) : f.slice();
    while (f.length < ancho) f.push('');
    out.push(f);
  }
  return out;
}

function hojaLimpia(ss, nombre, encabezados) {
  var h = hoja(ss, nombre);
  h.clear();
  h.getRange(1, 1, 1, encabezados.length).setValues([encabezados])
    .setFontWeight('bold').setFontColor('#ffffff').setBackground('#14284B')
    .setFontSize(10);
  h.setFrozenRows(1);
  return h;
}

function resumen(ss) {
  var h = hoja(ss, 'Resumen');
  h.clear();
  h.getRange('A1').setValue('REY PIZZA — resumen por mes')
    .setFontSize(16).setFontWeight('bold').setFontColor('#14284B');
  h.getRange('A2').setValue('Se calcula solo con lo que se registra en la app.')
    .setFontColor('#6c7079').setFontSize(9);
  h.getRange('A4').setValue('Mes (escriba así: 2026-09)').setFontWeight('bold');
  h.getRange('B4').setValue(Utilities.formatDate(new Date(), ZONA, 'yyyy-MM'))
    .setBackground('#fff2cc').setFontColor('#0000ff').setHorizontalAlignment('center');
  h.getRange(6, 1, 1, 4)
    .setValues([['Unidad', 'Ingreso del mes', 'Gasto del mes', 'Resultado']])
    .setFontWeight('bold').setFontColor('#ffffff').setBackground('#14284B');
  UNIDADES.forEach(function (u, k) {
    var f = 7 + k;
    h.getRange(f, 1).setValue(u.nombre);
    h.getRange(f, 2).setFormula(u.vende
      ? '=SUMIFS(Ingresos!$I:$I,Ingresos!$E:$E,$A' + f +
        ',Ingresos!$G:$G,$B$4,Ingresos!$K:$K,"No")' : '=0');
    // el gasto incluye lo que salió de la caja durante los turnos
    h.getRange(f, 3).setFormula('=SUMIFS(Gastos!$I:$I,Gastos!$E:$E,$A' + f +
      ',Gastos!$H:$H,$B$4,Gastos!$K:$K,"No")' +
      '+SUMIFS(Ingresos!$Q:$Q,Ingresos!$E:$E,$A' + f +
      ',Ingresos!$G:$G,$B$4,Ingresos!$K:$K,"No")');
    h.getRange(f, 4).setFormula('=B' + f + '-C' + f);
  });
  var t = 7 + UNIDADES.length;
  h.getRange(t, 1).setValue('TOTAL').setFontWeight('bold');
  ['B', 'C', 'D'].forEach(function (c) {
    h.getRange(c + t).setFormula('=SUM(' + c + '7:' + c + (t - 1) + ')')
      .setFontWeight('bold');
  });
  h.getRange(7, 2, UNIDADES.length + 1, 3).setNumberFormat('"Q"#,##0.00');
  h.setColumnWidth(1, 160);
  [2, 3, 4].forEach(function (c) { h.setColumnWidth(c, 140); });
}

/* ════════════ ESTADO DE RESULTADOS (hoja) ════════════ */

/** Una hoja por mes, con una columna por sucursal, el consolidado de Rey Pizza,
 *  la bodega aparte y el total de la sociedad. Cada sucursal es su propia unidad
 *  de negocio: la bodega NO se reparte entre ellas. */
function estadoResultados(ss) {
  var h = hoja(ss, 'Estado de resultados');
  h.clear();

  var cols = columnasER();
  var letra = {};                       // nombre de unidad → letra de columna
  cols.forEach(function (c, i) {
    if (c.u.length === 1) letra[c.u[0]] = String.fromCharCode(66 + i);
  });
  function formula(c, fila, propia) {
    if (c.u.length === 1) return propia(String.fromCharCode(66 + cols.indexOf(c)));
    return '=' + c.u.map(function (u) { return letra[u] + fila; }).join('+');
  }

  h.getRange('A1').setValue('Estado de resultados — CAIX, S.A.')
    .setFontSize(16).setFontWeight('bold').setFontColor('#14284B');
  h.getRange('A2').setValue('Sale de lo registrado en la app. Lo único que se ' +
    'escribe aquí es el mes. Los grupos se cambian en la hoja «Clasificación».')
    .setFontColor('#6c7079').setFontSize(9);
  h.getRange('A4').setValue('Mes (escriba así: 2026-09)').setFontWeight('bold');
  h.getRange('B4').setValue(Utilities.formatDate(new Date(), ZONA, 'yyyy-MM'))
    .setBackground('#fff2cc').setFontColor('#0000ff')
    .setHorizontalAlignment('center').setNumberFormat('@');

  var enc = ['Concepto'].concat(cols.map(function (c) { return c.t; }));
  h.getRange(6, 1, 1, enc.length).setValues([enc])
    .setFontWeight('bold').setFontColor('#ffffff').setBackground('#14284B')
    .setHorizontalAlignment('center');
  h.getRange(6, 1).setHorizontalAlignment('left');

  var cl = clasifica();
  var f = 7;
  var seccion = [], subtot = [], fuerte = [], porc = [];

  function etiqueta(txt, negrita) {
    h.getRange(f, 1).setValue(txt).setFontWeight(negrita ? 'bold' : 'normal');
  }
  function linea(txt, propia, clase) {
    etiqueta(txt, clase !== 'cat');
    h.getRange(f, 2, 1, cols.length).setFormulas([cols.map(function (c) {
      return formula(c, f, propia);
    })]);
    if (clase === 'sec') seccion.push(f);
    if (clase === 'sub') subtot.push(f);
    if (clase === 'res') fuerte.push(f);
    f++;
  }

  // ── ventas ──
  h.getRange(f, 1).setValue('VENTAS').setFontWeight('bold')
    .setFontColor('#14284B');
  seccion.push(f); f++;
  linea('Ventas (cortes)', function (L) {
    return '=SUMIFS(Ingresos!$I:$I,Ingresos!$E:$E,' + L + '$6,' +
      'Ingresos!$G:$G,$B$4,Ingresos!$K:$K,"No")';
  }, 'cat');
  var fCortes = f - 1;
  linea('(−) Envíos entregados a repartidores', function (L) {
    return '=SUMIFS(Gastos!$I:$I,Gastos!$E:$E,' + L + '$6,Gastos!$H:$H,$B$4,' +
      'Gastos!$F:$F,"' + CAT_ENVIOS + '",Gastos!$K:$K,"No")';
  }, 'cat');
  var fEnv = f - 1;
  linea('Ventas netas', function (L) { return '=' + L + fCortes + '-' + L + fEnv; }, 'sub');
  var fVentas = f - 1;

  function fCat(cat) {
    // los gastos de caja no viven en Gastos: se capturan en el corte del turno
    if (cat === CAT_COMISION) {
      return function (L) {
        return '=SUMIFS(Ingresos!$Y:$Y,Ingresos!$E:$E,' + L +
          '$6,Ingresos!$G:$G,$B$4,Ingresos!$K:$K,"No")';
      };
    }
    if (cat === CAT_CAJA) {
      return function (L) {
        return '=SUMIFS(Ingresos!$Q:$Q,Ingresos!$E:$E,' + L +
          '$6,Ingresos!$G:$G,$B$4,Ingresos!$K:$K,"No")';
      };
    }
    return function (L) {
      return '=SUMIFS(\'Gastos detalle\'!$M:$M,\'Gastos detalle\'!$D:$D,' + L +
        '$6,\'Gastos detalle\'!$G:$G,$B$4,\'Gastos detalle\'!$H:$H,"' + cat +
        '",\'Gastos detalle\'!$N:$N,"No")' +
        '+SUMIFS(Gastos!$I:$I,Gastos!$E:$E,' + L + '$6,Gastos!$H:$H,$B$4,' +
        'Gastos!$F:$F,"' + cat + '",Gastos!$K:$K,"No",Gastos!$Q:$Q,0)';
    };
  }
  function sumaFilas(filas) {
    return function (L) {
      return '=' + filas.map(function (x) { return L + x; }).join('+');
    };
  }

  // ── costo de ventas ──
  f++;
  h.getRange(f, 1).setValue('COSTO DE VENTAS').setFontWeight('bold')
    .setFontColor('#14284B');
  seccion.push(f); f++;
  var fCosto = [];
  cl.orden.forEach(function (g) {
    if (g.grupo !== G_COSTO) return;
    g.categorias.forEach(function (c) {
      fCosto.push(f); linea('   ' + c, fCat(c), 'cat');
    });
  });
  linea('Total costo de ventas', sumaFilas(fCosto), 'sub');
  var fTotCosto = f - 1;

  linea('UTILIDAD BRUTA', function (L) {
    return '=' + L + fVentas + '-' + L + fTotCosto; }, 'res');
  var fBruta = f - 1;
  linea('Margen bruto', function (L) {
    return '=IF(' + L + fVentas + '=0,"",' + L + fBruta + '/' + L + fVentas + ')';
  }, 'cat');
  porc.push(f - 1);

  // ── gastos de operación ──
  f++;
  h.getRange(f, 1).setValue('GASTOS DE OPERACIÓN').setFontWeight('bold')
    .setFontColor('#14284B');
  seccion.push(f); f++;
  var fSub = [];
  cl.orden.forEach(function (g) {
    if (g.grupo !== G_OPER) return;
    var fg = [];
    g.categorias.forEach(function (c) {
      fg.push(f); linea('   ' + c, fCat(c), 'cat');
    });
    fSub.push(f);
    linea(g.renglon, sumaFilas(fg), 'sub');
  });
  linea('Total gastos de operación', sumaFilas(fSub), 'sub');
  var fTotOper = f - 1;

  linea('UTILIDAD DE OPERACIÓN', function (L) {
    return '=' + L + fBruta + '-' + L + fTotOper; }, 'res');
  var fOper = f - 1;
  linea('Margen de operación', function (L) {
    return '=IF(' + L + fVentas + '=0,"",' + L + fOper + '/' + L + fVentas + ')';
  }, 'cat');
  porc.push(f - 1);

  // ── formato ──
  var ancho = cols.length + 1;
  h.getRange(7, 2, f - 7, cols.length).setNumberFormat('"Q"#,##0.00');
  porc.forEach(function (x) {
    h.getRange(x, 2, 1, cols.length).setNumberFormat('0.0%');
  });
  // primero el tinte de las columnas consolidadas, para que los renglones ganen
  cols.forEach(function (c, i) {
    if (c.fuerte) h.getRange(6, i + 2, f - 6, 1).setBackground('#f7f9fd');
  });
  seccion.forEach(function (x) {
    h.getRange(x, 1, 1, ancho).setBackground('#eef2f8');
  });
  subtot.forEach(function (x) {
    h.getRange(x, 1, 1, ancho).setFontWeight('bold')
      .setBorder(true, null, null, null, null, null, '#c9c6c0',
        SpreadsheetApp.BorderStyle.SOLID);
  });
  fuerte.forEach(function (x) {
    h.getRange(x, 1, 1, ancho).setFontWeight('bold').setBackground('#F7F2E7')
      .setBorder(true, null, true, null, null, null, '#14284B',
        SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  });
  h.setColumnWidth(1, 260);
  for (var k = 0; k < cols.length; k++) h.setColumnWidth(k + 2, 118);
  h.setFrozenRows(6); h.setFrozenColumns(1);
  return h;
}

/* ════════════ LA PÁGINA ════════════ */

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('REY PIZZA · CAIX, S.A.')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/* ════════════ API PARA LA APP INSTALADA ════════════
 * La app instalable (app.caixsa.com) no corre dentro de Google: le pide los datos a este
 * mismo proyecto por POST. Solo se pueden llamar las funciones de esta lista (las mismas
 * que usa la pantalla), y cada una revisa el código de quien pregunta, igual que antes.
 * La lista la pone sola el armado; no hace falta tocarla. */
/*API_INI*/var API_PERMITIDAS = ['abrirSolicitud', 'abrirTraslado', 'aceptarSolicitud', 'aceptarYPreparar', 'aclararDiferencia', 'activarCorreos', 'activarObligacion', 'activarProducto', 'activarProductoSucursal', 'adjuntarCotizacion', 'agregarEquipo', 'agregarProducto', 'agregarRecomendados', 'agregarRecomendadosCmo', 'agregarRubro', 'anclarProducto', 'anular', 'anularOrden', 'anularPagoImpuesto', 'anularSolicitud', 'aprobarFondos', 'aprobarPresupuesto', 'aprobarPresupuestoArea', 'asignarDescanso', 'borrarKpi', 'borrarKpiCmo', 'borrarSolicitud', 'cambiarEquipo', 'cambiarMiPin', 'cambiarPermiso', 'cancelarPermiso', 'confirmarPlanilla', 'confirmarRegistro', 'consolidado', 'copiarPresupuesto', 'costosTraslados', 'cotizarOrden', 'darDeBaja', 'datos', 'despacharSolicitudes', 'detalleGasto', 'editarProducto', 'eliminarPersonaCal', 'entrar', 'enviarEmergencia', 'enviarKpis', 'enviarKpisCmo', 'enviarSolicitud', 'enviarSugerencia', 'estadoCorreos', 'evaluarMerma', 'existencias', 'existenciasEmergencia', 'firmarTraslado', 'guardarAjustesInv', 'guardarBonoGeneral', 'guardarBonosCal', 'guardarBorradorCal', 'guardarComision', 'guardarConfigCal', 'guardarConteo', 'guardarEstudio', 'guardarHorario', 'guardarHorasApoyo', 'guardarHorasReales', 'guardarKpi', 'guardarKpiCmo', 'guardarLugar', 'guardarMando', 'guardarMedida', 'guardarMiCorreo', 'guardarMiPresupuesto', 'guardarObligacion', 'guardarOrden', 'guardarPersonaCal', 'guardarPrecios', 'guardarPresupuesto', 'guardarProducto', 'guardarProductoBodega', 'guardarProveedor', 'guardarProveedorInv', 'guardarReceta', 'guardarRepartidor', 'guardarTraslado', 'guardarTurnos', 'guardarUsuario', 'habilita', 'habilitarConteo', 'historialTraslados', 'indicadoresConta', 'leerBonoGeneral', 'leerSugerencias', 'listaBodega', 'listaCatalogo', 'listaConsultas', 'listaDiferencias', 'listaEntrada', 'listaFondos', 'listaImpuestos', 'listaKpis', 'listaKpisCmo', 'listaLugares', 'listaOrdenes', 'listaPermisos', 'listaPrecios', 'listaRepartidores', 'listaReportes', 'listaSolicitudes', 'listaUsuarios', 'llegoSolicitud', 'marcarPreparadas', 'medirAceite', 'miAvance', 'miCorreo', 'miPanel', 'miPresupuesto', 'notificaciones', 'ordenesPago', 'pagarFondos', 'pagarOrden', 'pagarRepartidor', 'pantallaCalendarios', 'pantallaConteo', 'pantallaEmergencias', 'pantallaFirmas', 'pantallaHorarios', 'pantallaKpi', 'pantallaMando', 'pantallaMant', 'pantallaMerma', 'pantallaMuestreo', 'pantallaPedidosBodega', 'pantallaProduccion', 'pantallaSolicitud', 'pantallaSugerencias', 'pdfCalendario', 'pdfOrden', 'pedirDescanso', 'pedirEmergencia', 'pedirFondos', 'pedirPermiso', 'ping', 'planillasPorPagar', 'ponPreciosGerentes', 'porConfirmar', 'preguntar', 'presupuestoMes', 'presupuestosAreas', 'probarCorreo', 'productosConfig', 'productosParaAnclar', 'productosSucursal', 'reabrirPlanilla', 'rechazarRegistro', 'recibirEfectivo', 'recibirEmergencia', 'recibirOrden', 'recibirSolicitud', 'registrarGasto', 'registrarIngreso', 'registrarKpis', 'registrarLote', 'registrarMant', 'registrarMerma', 'registrarObservacion', 'registrarPagoImpuesto', 'renombrarCategoria', 'reporteDiario', 'reporteTraslados', 'resolverDescanso', 'responderConsulta', 'responderMerma', 'resultadosKpi', 'resultadosKpiCmo', 'resultadosKpiRango', 'resultadosMuestreo', 'resumenInventario', 'resumenMes', 'revisarPermiso', 'seguimiento', 'sugerenciaCompra', 'tableroFinanzas', 'tableroKpi', 'terminarFondos', 'validarRegistro', 'ventas', 'verTraslado', 'verificarFondos', 'proveedoresProductos', 'guardarProveedorProd', 'eliminarProveedor', 'detalleAbastecimiento', 'pdfHistorialInventario', 'pantallaCierreMes', 'guardarCierreMes', 'misMenus', 'menusRol', 'guardarMenusRol', 'historialInventario', 'reiniciarPresupuestos', 'autorizarFondos', 'hechaFondos', 'listaPendientes', 'llegaPendiente', 'cancelarPendiente', 'despacharPendientes', 'devolverFondo', 'corregirFondo', 'pantallaEquipos', 'guardarEquipoA', 'bajaEquipo', 'revisarEquipos', 'confirmarRevisionEq', 'pedirEquipo', 'resolverPedidoEq', 'entregarCompraEq', 'recibirEquipo', 'devolverOrden', 'corregirOrden', 'reiniciarPrecios', 'guardarMezcla', 'registrarLoteMezcla', 'etiquetasLote', 'guardarConfigEtiquetas', 'listaKpisCoo', 'guardarKpiCoo', 'borrarKpiCoo', 'enviarKpisCoo', 'agregarRecomendadosCoo', 'resultadosKpiCoo', 'pantallaKpiCoo', 'registrarKpisCoo', 'resumenAreas', 'resumenAreasSenales', 'pantallaZonas', 'guardarZonas', 'listaKpisTec', 'guardarKpiTec', 'borrarKpiTec', 'enviarKpisTec', 'agregarRecomendadosTec', 'resultadosKpiTec', 'pantallaKpiTec', 'registrarKpisTec'];/*API_FIN*/
/** Solo contesta «aquí estoy»: sirve para despertar el servidor y medir la velocidad. No lee ni guarda nada. */
function ping() { return { ok: true }; }
function doPost(e) {
  var out;
  try {
    var q = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    var t0 = Date.now();
    if (q.interno === true) out = { ok: true, v: interno(String(q.fn || ''), q) };       // el servidor propio, con su clave secreta
    else out = { ok: true, v: llamar(String(q.fn || ''), String(q.rid || ''), Array.isArray(q.args) ? q.args : []) };
    out.ms = Date.now() - t0;
  } catch (err) {
    out = { ok: false, error: String(err && err.message ? err.message : err) };
  }
  var tocadas = Object.keys(_TOCADAS); if (tocadas.length) out.tocadas = tocadas;
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}


/** Los lugares (sucursales, bodega, oficinas…). Solo el administrador crea o cambia. */
function listaLugares(cred) {
  var yo = quien(cred);
  return { lugares: UNIDADES.map(function (u) {
      return { id: u.id, nombre: u.nombre, marca: u.marca, tipo: u.tipo || 'otro', vende: !!u.vende, extra: !!u.extra, activo: u.activo !== false };
    }), tipos: TIPOS_LUGAR, puedeCambiar: yo.esAdmin };
}
function guardarLugar(cred, d) {
  var yo = exigeAdmin(cred);
  d = d || {};
  var id = String(d.id || '').trim().toLowerCase(), nombre = String(d.nombre || '').replace(/\s+/g, ' ').trim();
  var desc = String(d.descripcion || '').replace(/\s+/g, ' ').trim().slice(0, 80);
  var tipo = TIPOS_LUGAR[d.tipo] ? d.tipo : 'oficina';
  var activo = d.activo === false ? 'No' : 'Sí';
  var lock = LockService.getScriptLock(); lock.waitLock(15000); _LEE = {};
  try {
    var h = hojaCat('Lugares', H_LUGARES);
    var filas = leeLugares(), destino = null;
    if (id) {
      var base = unidadPorId(id);
      if (!base) throw new Error('No se encontró ese lugar.');
      if (!base.extra) throw new Error(base.nombre + ' es una sucursal o la bodega de siempre: no se cambia aquí.');
      filas.forEach(function (x, i) { if (x.id === id) destino = i; });
      if (destino == null) throw new Error('No se encontró ese lugar.');
      h.getRange(destino + 2, 3).setValue(desc);
      h.getRange(destino + 2, 4, 1, 2).setValues([[tipo, activo]]);
    } else {
      if (nombre.length < 3) throw new Error('Escriba el nombre del lugar (mínimo 3 letras).');
      if (nombre.length > 30) throw new Error('El nombre del lugar es muy largo (máximo 30).');
      if (UNIDADES.some(function (u) { return sinAcentoS(u.nombre) === sinAcentoS(nombre); }) || /^(rey pizza|caix)/i.test(sinAcentoS(nombre)))
        throw new Error('Ya hay un lugar con ese nombre.');
      var slug = sinAcentoS(nombre).toLowerCase().replace(/[^a-z0-9]/g, '').slice(0, 14) || 'lugar', nuevo = slug, n = 1;
      while (unidadPorId(nuevo)) { n++; nuevo = slug.slice(0, 12) + n; }
      h.appendRow([nuevo, nombre, desc, tipo, activo, new Date(), yo.nombre]);
      id = nuevo;
    }
    recargaLugares();
    if (destino == null) { try { hojaDeLugar(unidadPorId(id)); } catch (e) {} }
  } finally { lock.releaseLock(); }
  var r = listaLugares(cred); r.ok = true;
  r.mensaje = d.id ? 'Se guardó ' + unidadPorId(id).nombre + '.' : unidadPorId(id).nombre + ' ya es un lugar: lleva sus gastos y presupuesto y suma al estado de resultados.';
  return r;
}
/** La hoja que muestra los gastos de un lugar (igual a la de cada sucursal). */
function hojaDeLugar(un) {
  var ss = libro();
  if (!un || ss.getSheetByName(un.nombre)) return;
  var h = ss.insertSheet(un.nombre);
  h.getRange('A1').setValue(un.nombre + ' · ' + un.marca).setFontSize(14).setFontWeight('bold');
  h.getRange('A2').setValue('Esta hoja se llena sola. No escriba en ella.').setFontColor('#6c7079').setFontSize(9);
  h.getRange('A4').setValue('GASTOS').setFontWeight('bold');
  h.getRange('A5').setFormula('=IFERROR(QUERY(Gastos!A:Q,"select G,O,F,I,Q,J,C,K where E = \'' + un.nombre + '\' order by G desc label G \'Fecha de pago\', O \'Proveedor\', F \'Categoría\', I \'Total\', Q \'Artículos\', J \'Nota\', C \'Registró\', K \'Anulado\'",1),"Sin gastos todavía")');
}

/** La puerta única de los cambios: cada llamada trae una llave (`rid`). Si la conexión se cae y la
 *  pantalla repite la llamada, el servidor reconoce la llave y devuelve lo que ya hizo en vez de
 *  guardar dos veces. Así el reintento automático nunca duplica un gasto, un pedido ni un pago. */
function llamar(fn, rid, args) {
  fn = String(fn || '');
  if (API_PERMITIDAS.indexOf(fn) < 0 || typeof globalThis[fn] !== 'function')
    throw new Error('La app pidió algo que no existe (' + fn + '). Falta pegar el Codigo.gs nuevo y crear una versión nueva de la implementación.');
  args = Array.isArray(args) ? args : [];
  rid = String(rid || '').replace(/[^A-Za-z0-9_-]/g, '').slice(0, 40);
  if (!rid) return globalThis[fn].apply(null, args);
  var c = null, k = 'rq_' + rid;
  try { c = CacheService.getScriptCache(); } catch (e) { c = null; }
  if (c) {
    var previo = c.get(k);
    if (previo === 'EN') {                          // el primer intento sigue corriendo: se espera su resultado
      for (var i = 0; i < 24 && previo === 'EN'; i++) { Utilities.sleep(1000); previo = c.get(k); }
      if (previo === 'EN') throw new Error('El servidor sigue trabajando en eso. Espere un momento y revise antes de guardar de nuevo.');
    }
    if (previo) {
      if (previo === 'OK') throw new Error('Eso ya se guardó. Actualice la pantalla para verlo.');
      return rqLee(previo);
    }
    c.put(k, 'EN', 120);
  }
  try {
    var v = globalThis[fn].apply(null, args);
    if (c) rqGuarda(c, k, v);
    return v;
  } catch (err) {
    if (c) { try { c.remove(k); } catch (e) {} }   // si falló, el reintento sí debe volver a correr
    throw err;
  }
}
function rqGuarda(c, k, v) {
  try {
    var s = JSON.stringify(v === undefined ? null : v);
    if (s.length > 90000) {
      try { s = 'Z' + Utilities.base64Encode(Utilities.gzip(Utilities.newBlob(s, 'application/json')).getBytes()); } catch (e) { s = 'OK'; }
      if (s.length > 95000) s = 'OK';
    }
    c.put(k, s, 600);
  } catch (e) {}
}
function rqLee(s) {
  if (s.charAt(0) === 'Z')
    return JSON.parse(Utilities.newBlob(Utilities.ungzip(Utilities.newBlob(Utilities.base64Decode(s.slice(1)), 'application/x-gzip'))).getDataAsString());
  return JSON.parse(s);
}

/* ════════════ USUARIOS Y PIN ════════════ */

/** Devuelve la hoja Usuarios y, si hace falta, la repara: la crea si no existe
 *  y deja un administrador inicial si no quedó ninguno activo. Así la app nunca
 *  se queda con la lista de nombres vacía. */
function hojaUsuarios() {
  var ss = libro();
  var h = ss.getSheetByName('Usuarios');
  var reparada = false;

  if (!h) {
    h = hojaLimpia(ss, 'Usuarios', H_USUARIOS);
    h.getRange('B:B').setNumberFormat('@');
    h.getRange('E:E').setNumberFormat('dd/mm/yyyy hh:mm');
    h.setColumnWidth(1, 160);
    reparada = true;
  }

  var hayAdmin = false;
  if (h.getLastRow() > 1) {
    var v = h.getRange(2, 1, h.getLastRow() - 1, 4).getValues();
    v.forEach(function (r) {
      if (String(r[0]).trim() &&
          String(r[2]).trim().toLowerCase() === 'admin' &&
          String(r[3]).trim().toLowerCase() !== 'no') hayAdmin = true;
    });
  }
  if (!hayAdmin) {
    var u0 = USUARIOS_INICIALES[0];
    h.appendRow([u0[0], u0[1], u0[2], 'Sí', new Date(), 'reparación']);
    h.getRange(h.getLastRow(), 2).setNumberFormat('@').setValue(u0[1]);
    reparada = true;
  }
  if (reparada) olvidaUsuarios();
  return h;
}

/** Deja la hoja Usuarios exactamente con los usuarios iniciales. Borra los que
 *  hubiera. NO toca gastos ni ingresos. Córrala desde el editor cuando los PIN
 *  no estén sirviendo y quiera empezar limpio. */
function usuariosDeArranque() {
  var ss = libro();
  var vieja = ss.getSheetByName('Usuarios');
  if (vieja) ss.deleteSheet(vieja);
  escribeUsuarios(ss);
  olvidaUsuarios();
  return diagnostico();
}

function escribeUsuarios(ss) {
  var h = hojaLimpia(ss, 'Usuarios', H_USUARIOS);
  h.getRange('B:B').setNumberFormat('@');          // el PIN se guarda como texto
  h.getRange('E:E').setNumberFormat('dd/mm/yyyy hh:mm');
  h.setColumnWidth(1, 160);
  var filas = USUARIOS_INICIALES.map(function (u) {
    return [u[0], u[1], u[2], 'Sí', new Date(), 'instalación'];
  });
  h.getRange(2, 1, filas.length, 6).setValues(filas);
  return h;
}

/** Para revisar desde el editor qué hay realmente en la hoja Usuarios. */
function diagnostico() {
  var ss = libro();
  var nombres = ss.getSheets().map(function (h) { return h.getName(); });
  var u = filasUsuarios().map(function (x) {
    return x.nombre + ' · ' + x.rol + ' · PIN ' + x.pin +
      (x.activo ? ' · activo' : ' · DADO DE BAJA');
  });
  var txt = 'Hoja: ' + ss.getName() + '\n' +
    'Pestañas: ' + nombres.join(', ') + '\n' +
    'Usuarios (' + u.length + '):\n  ' + (u.join('\n  ') || '(ninguno)');
  Logger.log(txt);
  return txt;
}

function filasUsuarios() {
  var h = hojaUsuarios();
  if (h.getLastRow() < 2) return [];
  var v = h.getRange(2, 1, h.getLastRow() - 1,
    Math.min(H_USUARIOS.length, h.getMaxColumns())).getValues();
  return v.map(function (r, i) {
    return {
      fila: i + 2,
      nombre: String(r[0]).trim(),
      pin: String(r[1]).trim(),
      rol: ROLES.indexOf(String(r[2]).trim().toLowerCase()) >= 0
        ? String(r[2]).trim().toLowerCase() : 'registro',
      activo: String(r[3]).trim().toLowerCase() !== 'no',
      sucursal: unidadPorId(String(r[6] || '').trim().toLowerCase())
        ? String(r[6]).trim().toLowerCase() : '',
      confirma: siNo(r[7]),
      correo: String(r[8] == null ? '' : r[8]).trim()
    };
  }).filter(function (u) { return u.nombre; });
}

/** Lo único que la app entrega sin PIN: los nombres para el desplegable. */
function usuariosParaEntrar() {
  return usuariosCache()
    .filter(function (u) { return u.activo; })
    .map(function (u) { return u.nombre; });
}

/** `cuenta` solo va en true cuando la persona aprieta Entrar a mano. El arranque
 *  automático nunca suma intentos fallidos: si no, refrescar la página con un PIN
 *  viejo guardado en el teléfono bloqueaba la cuenta sola. */
/** Para todo lo de la caja: gastos, ingresos, reportes, catálogo. El usuario de
 *  bodega entra, pero solo a su pantalla de traslados; aquí se le cierra lo demás
 *  aunque alguien intente llamar las funciones a mano. */
function verifica(cred, cuenta) {
  var yo = quien(cred, cuenta);
  if (yo.rol === 'bodega')
    throw new Error('Su usuario solo puede registrar traslados de bodega.');
  if (yo.rol === 'produccion')
    throw new Error('Su usuario es de producción: no ve la caja.');
  return yo;
}

/** Solo revisa nombre y PIN. Lo usan la entrada, el cambio de PIN y los traslados. */
function quien(cred, cuenta) {
  cred = cred || {};
  extiendeCategorias();
  extiendeLugares();                       // los lugares creados en la app valen desde la primera línea
  var nombre = String(cred.usuario || '').trim();
  var pin = String(cred.pin || '').trim();
  if (!nombre || !pin) throw new Error('Escoja su nombre y escriba su PIN.');

  var cache = null, llave = '', fallos = 0;
  if (cuenta) {
    cache = CacheService.getScriptCache();
    llave = 'fallos_' + nombre.toLowerCase();
    fallos = Number(cache.get(llave) || 0);
    if (fallos >= 8) {
      throw new Error('Demasiados intentos con ese nombre. Espere 10 minutos ' +
        'y vuelva a probar.');
    }
  }

  var hallado = null;
  usuariosCache().forEach(function (u) {
    if (u.nombre.toLowerCase() === nombre.toLowerCase()) hallado = u;
  });

  if (!hallado || !hallado.activo || !mismoPin(hallado.pin, pin)) {
    if (cache) cache.put(llave, String(fallos + 1), 600);
    throw new Error('Nombre o PIN incorrectos.');
  }
  if (cache) cache.remove(llave);
  if (hallado.rol === 'auxiliar' && !_AUXOK)
    throw new Error('Su usuario tiene acceso limitado: solo registra los gastos de su lugar.');
  if (hallado.rol === 'cmo' && !_AUXOK && !_CMOOK)
    throw new Error('Su usuario no entra a esta parte. Tiene su presupuesto, gastos, solicitudes e indicadores.');
  if ((hallado.rol === 'gerente' || hallado.rol === 'auxiliar' || hallado.rol === 'cmo') && !hallado.sucursal)
    throw new Error('A su usuario le falta la sucursal o el lugar. Pida al administrador que lo ponga.');
  return { nombre: hallado.nombre, rol: hallado.rol, esAdmin: hallado.rol === 'admin',
    codigoCorto: String(hallado.pin).length < 6,
    sucursal: (hallado.rol === 'gerente' || hallado.rol === 'auxiliar' || hallado.rol === 'cmo') ? hallado.sucursal : '',
    limitado: hallado.rol === 'auxiliar' || hallado.rol === 'cmo',
    cmo: hallado.rol === 'cmo',
    // valida = primer paso del cierre de caja (el director operativo);
    // recibe = último paso, contabilidad confirma que recibió el dinero
    valida: hallado.rol === 'admin' || hallado.rol === 'operaciones',
    // el contador interno (Registra con permiso de confirmar) recibe el efectivo
    recibe: hallado.rol === 'admin' || (hallado.confirma && hallado.rol === 'registro'),
    // el contador interno también anota la tarjeta del cierre NEONET y confirma el banco
    banco: hallado.rol === 'admin' || (hallado.confirma && hallado.rol === 'registro'),
    correo: hallado.correo || '',
    confirma: ['admin', 'operaciones', 'finanzas'].indexOf(hallado.rol) >= 0 ||
      (hallado.confirma && hallado.rol === 'registro'),
    soloMira: hallado.rol === 'operaciones' || hallado.rol === 'finanzas' || hallado.rol === 'dueno',
    // representa a los dueños de CAIX: ve todos los reportes y no cambia nada
    soloReportes: hallado.rol === 'dueno' };
}

/** Compara el PIN con tolerancia: espacios de sobra, y el caso en que la hoja
 *  lo haya guardado como número (1234) en vez de como texto ('1234'). */
function mismoPin(guardado, escrito) {
  var a = String(guardado == null ? '' : guardado).trim();
  var b = String(escrito == null ? '' : escrito).trim();
  if (!a || !b) return false;
  if (a === b) return true;
  if (/^\d+$/.test(a) && /^\d+$/.test(b) && Number(a) === Number(b)) return true;
  return false;
}

/** Lo de ver cifras de la empresa: el gerente no entra aquí, solo envía datos. */
function verificaCaja(cred) {
  var yo = verifica(cred);
  if (yo.rol === 'gerente')
    throw new Error('Su usuario solo envía registros de su sucursal.');
  if (yo.limitado)
    throw new Error('Su usuario tiene acceso limitado: solo registra los gastos de su lugar.');
  return yo;
}

/** Operaciones y finanzas ven, pero no escriben en la caja. */
function noMira(yo) {
  if (yo.soloMira) throw new Error('Su usuario solo revisa: no registra ni cambia movimientos.');
  return yo;
}

function exigeAdmin(cred) {
  var yo = verifica(cred);
  if (!yo.esAdmin) throw new Error('Solo un administrador puede hacer eso.');
  return yo;
}

/* ════════════ PROVEEDORES Y PRODUCTOS ════════════ */

/** Igual que hojaUsuarios: si la pestaña no está, la crea. Así la app no se cae
 *  cuando alguien todavía no ha corrido instalar. */
function hojaCat(nombre, encabezados) {
  var ss = libro();
  var h = ss.getSheetByName(nombre);
  if (h) return h;
  h = hojaLimpia(ss, nombre, encabezados);
  h.setColumnWidth(1, 230);
  if (nombre === 'Gastos detalle') {
    h.getRange('B:B').setNumberFormat('dd/mm/yyyy hh:mm');
    h.getRange('F:F').setNumberFormat('dd/mm/yyyy');
    h.getRange('L:M').setNumberFormat('"Q"#,##0.00');
    h.setColumnWidth(9, 230);
  }
  if (nombre === 'Productos') h.getRange('E:E').setNumberFormat('"Q"#,##0.00');
  if (nombre === 'Reportes') {
    h.getRange('A2:A').setNumberFormat('@');
    h.getRange('B:B').setNumberFormat('dd/mm/yyyy hh:mm');
  }
  return h;
}

function filasProveedores() {
  var h = hojaCat('Proveedores', H_PROVEEDORES);
  return leeTodo(h, H_PROVEEDORES.length).map(function (r, i) {
    return {
      fila: i + 2,
      nombre: String(r[0]).trim(),
      nit: String(r[1] || '').trim(),
      tel: String(r[2] || '').trim(),
      activo: String(r[3]).trim().toLowerCase() !== 'no',
      tipo: TIPOS_PROV.indexOf(String(r[6] || '').trim()) >= 0 ? String(r[6] || '').trim() : '',
      sucursal: unidadPorId(String(r[7] || '').trim().toLowerCase()) ? String(r[7]).trim().toLowerCase() : ''
    };
  }).filter(function (p) { return p.nombre; });
}
function esDeBodega(p) { return !p.tipo; }

function filasProductos() {
  var h = hojaCat('Productos', H_PRODUCTOS);
  return leeTodo(h, H_PRODUCTOS.length).map(function (r, i) {
    return {
      fila: i + 2,
      nombre: String(r[0]).trim(),
      proveedor: String(r[1] || '').trim(),
      medida: String(r[2] || '').trim() || 'unidad',
      categoria: String(r[3] || '').trim(),
      precio: Number(r[4]) || 0,
      activo: String(r[5]).trim().toLowerCase() !== 'no',
      contenido: Number(r[8]) || 0,          // cuántas unidades base trae
      base: String(r[9] || '').trim()        // libra, litro, unidad…
    };
  }).filter(function (p) { return p.nombre; });
}

/** Lo que interesa de una presentación: cuánto trae y cuánto cuesta lo de adentro. */
function conPresentacion(p) {
  return { nombre: p.nombre, proveedor: p.proveedor, medida: p.medida,
    categoria: p.categoria, precio: p.precio,
    contenido: p.contenido, base: p.base,
    costoBase: (p.contenido > 0 && p.precio > 0)
      ? Math.round((p.precio / p.contenido) * 10000) / 10000 : 0 };
}

/** Lo que necesita el formulario de gasto: solo lo que está habilitado. */
function catalogo() {
  return {
    proveedores: filasProveedores()
      .filter(function (p) { return p.activo && p.tipo !== 'Repartidor'; })
      .map(function (p) { return p.nombre; })
      .sort(function (a, b) { return a.localeCompare(b, 'es'); }),
    // para agrupar el desplegable del gasto: a quién se le paga en cada lugar
    provInfo: filasProveedores()
      .filter(function (p) { return p.activo && p.tipo !== 'Repartidor'; })
      .map(function (p) { return { n: p.nombre, t: p.tipo, s: p.sucursal }; })
      .sort(function (a, b) { return a.n.localeCompare(b.n, 'es'); }),
    productos: filasProductos()
      .filter(function (p) { return p.activo; })
      .map(conPresentacion)
      .sort(function (a, b) { return a.nombre.localeCompare(b.nombre, 'es'); }),
    medidas: MEDIDAS,
    bases: BASES
  };
}

/** Todo, habilitado y deshabilitado, para la pantalla de administración. */
function listaCatalogo(cred) {
  verificaCaja(cred);
  // los productos que Samuel ya asignó a cada proveedor viven en «Productos de bodega»
  var anc = {};
  try {
    filasBodega().forEach(function (b) {
      if (b.activo && b.proveedor) { var k = b.proveedor.toLowerCase(); anc[k] = (anc[k] || 0) + 1; }
    });
  } catch (e) {}
  return {
    proveedores: filasProveedores().map(function (p) {
      return { nombre: p.nombre, nit: p.nit, tel: p.tel, activo: p.activo, tipo: p.tipo, sucursal: p.sucursal,
        anclados: anc[p.nombre.toLowerCase()] || 0 };
    }),
    productos: filasProductos().map(function (p) {
      var x = conPresentacion(p);
      x.activo = p.activo;
      return x;
    }),
    categorias: CATEGORIAS,
    medidas: MEDIDAS,
    bases: BASES
  };
}

function guardarProveedor(cred, p) {
  var yo = noMira(verifica(cred));
  p = p || {};
  var nombre = String(p.nombre || '').trim();
  var original = String(p.original || '').trim();
  if (yo.rol === 'gerente' && original)
    throw new Error('Solo el administrador cambia proveedores que ya existen.');
  if (nombre.length < 2) throw new Error('El nombre del proveedor es muy corto.');
  if (nombre.length > 60) throw new Error('El nombre del proveedor es muy largo.');

  var lock = LockService.getScriptLock();
  lock.waitLock(15000); _LEE = {};
  try {
    var h = hojaCat('Proveedores', H_PROVEEDORES);
    var todos = filasProveedores();
    var destino = null, choque = false;
    todos.forEach(function (x) {
      if (x.nombre.toLowerCase() === original.toLowerCase() && original) destino = x;
      else if (x.nombre.toLowerCase() === nombre.toLowerCase()) choque = true;
    });
    if (!original) {
      todos.forEach(function (x) {
        if (x.nombre.toLowerCase() === nombre.toLowerCase()) destino = x;
      });
      if (destino) {                            // ya existía: solo lo reactiva
        h.getRange(destino.fila, 4).setValue('Sí');
        return { ok: true, mensaje: nombre + ' ya estaba en la lista.',
                 catalogo: listaCatalogo(cred) };
      }
    }
    if (choque) throw new Error('Ya hay un proveedor con ese nombre.');

    var activo = p.activo === false ? 'No' : 'Sí';
    var tipo = TIPOS_PROV.indexOf(String(p.tipo || '')) >= 0 ? String(p.tipo || '') : '';
    var sucP = tipo && unidadPorId(String(p.sucursal || '')) ? String(p.sucursal) : '';
    if (tipo === 'Repartidor' && sucP && !unidadPorId(sucP).vende)
      throw new Error('El repartidor recibe el dinero de una sucursal que vende: escoja Centro, Almendras o Parque.');
    if (destino) {
      h.getRange(destino.fila, 1, 1, 4).setValues([[nombre,
        String(p.nit || '').slice(0, 20), String(p.tel || '').slice(0, 20), activo]]);
      if (yo.esAdmin) h.getRange(destino.fila, 7, 1, 2).setValues([[tipo, sucP]]);
      // si le cambiaron el nombre, se arrastra a los productos
      if (original && original.toLowerCase() !== nombre.toLowerCase()) {
        var hp = hojaCat('Productos', H_PRODUCTOS);
        filasProductos().forEach(function (x) {
          if (x.proveedor.toLowerCase() === original.toLowerCase())
            hp.getRange(x.fila, 2).setValue(nombre);
        });
      }
    } else {
      h.appendRow([nombre, String(p.nit || '').slice(0, 20),
        String(p.tel || '').slice(0, 20), activo, new Date(), yo.nombre, yo.esAdmin ? tipo : '', yo.esAdmin ? sucP : '']);
    }
    return { ok: true, mensaje: 'Proveedor guardado: ' + nombre + '.',
             catalogo: listaCatalogo(cred) };
  } finally {
    lock.releaseLock();
  }
}

function guardarProducto(cred, p) {
  var yo = noMira(verifica(cred));
  p = p || {};
  var nombre = String(p.nombre || '').trim();
  var original = String(p.original || '').trim();
  if (yo.rol === 'gerente' && original)
    throw new Error('Solo el administrador cambia productos que ya existen.');
  var proveedor = String(p.proveedor || '').trim();
  var medida = String(p.medida || '').trim() || 'unidad';
  var categoria = CATEGORIAS.indexOf(p.categoria) >= 0 ? p.categoria : 'Insumos de cocina';
  var precio = Math.round((Number(p.precio) || 0) * 100) / 100;
  var contenido = Math.round((Number(p.contenido) || 0) * 1000) / 1000;
  var base = String(p.base || '').trim().slice(0, 20);
  if (!(contenido > 0) || !base) { contenido = 0; base = ''; }
  if (contenido > 0 && base.toLowerCase() === medida.toLowerCase()) {
    // «una caja trae 15 cajas» no dice nada: se deja sin presentación
    contenido = 0; base = '';
  }

  if (nombre.length < 2) throw new Error('El nombre del producto es muy corto.');
  if (nombre.length > 70) throw new Error('El nombre del producto es muy largo.');

  var lock = LockService.getScriptLock();
  lock.waitLock(15000); _LEE = {};
  try {
    var h = hojaCat('Productos', H_PRODUCTOS);
    var todos = filasProductos();
    var origProv = String(p.originalProveedor == null ? proveedor
                                                      : p.originalProveedor).trim();
    var destino = null;
    if (original) {
      todos.forEach(function (x) {
        if (x.nombre.toLowerCase() === original.toLowerCase() &&
            x.proveedor.toLowerCase() === origProv.toLowerCase()) destino = x;
      });
      if (!destino) throw new Error('No se encontró ese producto.');
    } else {
      // producto nuevo: si ya existía uno igual, se reusa en vez de duplicarlo
      todos.forEach(function (x) {
        if (x.nombre.toLowerCase() === nombre.toLowerCase() &&
            x.proveedor.toLowerCase() === proveedor.toLowerCase()) destino = x;
      });
    }
    var choque = false;
    todos.forEach(function (x) {
      if (destino && x.fila === destino.fila) return;
      if (x.nombre.toLowerCase() === nombre.toLowerCase() &&
          x.proveedor.toLowerCase() === proveedor.toLowerCase()) choque = true;
    });
    if (choque) throw new Error('Ya hay un producto con ese nombre para ese proveedor.');

    var activo = p.activo === false ? 'No' : 'Sí';
    var fin = precio || (destino ? destino.precio : 0);
    if (destino) {
      h.getRange(destino.fila, 1, 1, 6).setValues([[nombre, proveedor, medida,
        categoria, fin, activo]]);
      h.getRange(destino.fila, 9, 1, 2).setValues([[contenido, base]]);
    } else {
      h.appendRow([nombre, proveedor, medida, categoria, precio, activo,
        new Date(), yo.nombre, contenido, base]);
    }
    if (proveedor) aseguraProveedor(proveedor, yo.nombre);
    return { ok: true, mensaje: 'Producto guardado: ' + nombre + '.',
             catalogo: listaCatalogo(cred),
             producto: conPresentacion({ nombre: nombre, proveedor: proveedor,
               medida: medida, categoria: categoria, precio: fin,
               contenido: contenido, base: base }) };
  } finally {
    lock.releaseLock();
  }
}

/** Si el proveedor que se escribió a mano no estaba, lo deja anotado. */
function aseguraProveedor(nombre, quien, unidadId) {
  nombre = String(nombre || '').trim();
  if (!nombre) return;
  var hay = false;
  filasProveedores().forEach(function (p) {
    if (p.nombre.toLowerCase() === nombre.toLowerCase()) hay = true;
  });
  if (hay) return;
  // escrito a mano al registrar un gasto de un lugar: queda como gasto de ese lugar
  // (el administrador puede cambiarlo a proveedor de bodega central al editarlo)
  var lugar = unidadId && unidadPorId(unidadId) ? unidadId : '';
  hojaCat('Proveedores', H_PROVEEDORES)
    .appendRow([nombre, '', '', 'Sí', new Date(), quien || 'app', lugar ? 'Sucursal' : '', lugar]);
}

/** Deshabilitar no borra nada: el producto deja de salir en la lista, pero el
 *  histórico de compras se queda completo. */
function habilita(cred, que, nombre, proveedor, activo) {
  noMira(verificaCaja(cred));
  nombre = String(nombre || '').trim();
  var si = activo === true || activo === 'Sí';
  var lock = LockService.getScriptLock();
  lock.waitLock(15000); _LEE = {};
  try {
    if (que === 'proveedor') {
      var hp = hojaCat('Proveedores', H_PROVEEDORES), dp = null;
      filasProveedores().forEach(function (x) {
        if (x.nombre.toLowerCase() === nombre.toLowerCase()) dp = x;
      });
      if (!dp) throw new Error('No se encontró ese proveedor.');
      hp.getRange(dp.fila, 4).setValue(si ? 'Sí' : 'No');
    } else {
      var hx = hojaCat('Productos', H_PRODUCTOS), dx = null;
      var pv = String(proveedor || '').trim().toLowerCase();
      filasProductos().forEach(function (x) {
        if (x.nombre.toLowerCase() === nombre.toLowerCase() &&
            x.proveedor.toLowerCase() === pv) dx = x;
      });
      if (!dx) throw new Error('No se encontró ese producto.');
      hx.getRange(dx.fila, 6).setValue(si ? 'Sí' : 'No');
    }
    return { ok: true, mensaje: nombre + (si ? ' quedó habilitado.' : ' quedó deshabilitado.'),
             catalogo: listaCatalogo(cred) };
  } finally {
    lock.releaseLock();
  }
}

/** Guarda el precio con el que se compró de último, para que la próxima vez
 *  venga puesto solo. */
function recuerdaPrecios(lineas) {
  try {
    var h = hojaCat('Productos', H_PRODUCTOS);
    var todos = filasProductos();
    lineas.forEach(function (l) {
      if (!(l.precio > 0)) return;
      todos.forEach(function (x) {
        if (x.nombre.toLowerCase() === l.producto.toLowerCase() &&
            x.proveedor.toLowerCase() === String(l.proveedor || '').toLowerCase())
          h.getRange(x.fila, 5).setValue(l.precio);
      });
    });
  } catch (e) { /* si falla, no importa: es una comodidad, no un dato */ }
}

/* ════════════ DATOS PARA LA PANTALLA ════════════ */

/** Sin PIN devuelve solo la lista de nombres; con PIN válido, todo. */
/** Llamada chiquita: solo los nombres para el desplegable. Es la que sostiene
 *  la pantalla de entrada, así que no depende de leer gastos ni ingresos. */
function listaEntrada() {
  var out = {
    usuarios: [],
    hoy: Utilities.formatDate(new Date(), ZONA, 'yyyy-MM-dd'),
    ahora: Utilities.formatDate(new Date(), ZONA, 'dd/MM/yyyy, HH:mm'),
    problema: ''
  };
  try {
    // ya no se mandan los nombres: se entra solo con el código
    out.hay = usuariosParaEntrar().length;
    if (!out.hay) out.problema = 'No hay ninguna persona activa.';
  } catch (e) {
    out.problema = 'No se pudo leer la hoja Usuarios: ' + e.message;
  }
  return out;
}

/** Los datos de la app. Pide PIN. Si falla la lectura de movimientos, igual
 *  deja entrar: mejor entrar sin historial que no entrar. */
function datos(cred) {
  permiteAux();
  var yo = quien(cred);
  corrigeMesPlanillas();
  if (yo.rol === 'bodega') {
    // la bodega no ve ventas, gastos ni precios: traslados, solicitudes,
    // inventario y compras (sin precios)
    return { entro: true, yo: yo, soloTraslados: true,
      hoy: hoyISO(), ahora: Utilities.formatDate(new Date(), ZONA, 'dd/MM/yyyy, HH:mm'),
      traslado: pantallaTraslado(yo), avisos: avisosInventario(yo),
      notificaciones: (function () { try { return notificaciones(cred).avisos; } catch (e) { return []; } })() };
  }
  if (yo.rol === 'produccion') {
    return { entro: true, yo: yo, soloProduccion: true, hoy: hoyISO(),
      ahora: Utilities.formatDate(new Date(), ZONA, 'dd/MM/yyyy, HH:mm'),
      notificaciones: (function () { try { return notificaciones(cred).avisos; } catch (e) { return []; } })() };
  }
  if (yo.limitado) {
    // acceso limitado: solo registra gastos de su lugar y ve lo que él mismo envió.
    // La CMO, en cambio, no ve «lo enviado»: ve el avance de su gasto contra su presupuesto.
    var a = { entro: true, yo: yo, soloAuxiliar: true, soloGerente: true,
      unidades: UNIDADES.filter(function (u) { return u.id === yo.sucursal; }),
      categorias: CATEGORIAS, medidas: MEDIDAS,
      hoy: hoyISO(), ahora: Utilities.formatDate(new Date(), ZONA, 'dd/MM/yyyy, HH:mm'),
      proveedores: [], productos: [], movimientos: [], ocupados: {}, avisoDatos: '', avisos: [], notificaciones: [] };
    try { var ca = catalogo(); a.proveedores = ca.proveedores; a.productos = ca.productos; }
    catch (e) { a.avisoDatos = 'No se pudo leer el catálogo de productos: ' + e.message; }
    if (yo.cmo) {
      a.esCmo = true; a.enviados = [];
      try { a.notificaciones = notificaciones(cred).avisos; } catch (e) { a.notificaciones = []; }
    } else {
      try { a.enviados = misEnviados(yo).filter(function (x) { return x.quien === yo.nombre; }); } catch (e) { a.enviados = []; }
    }
    return a;
  }
  if (yo.rol === 'gerente') {
    // el gerente envía sus datos, pero no ve resúmenes: solo lo que él envió
    var g = { entro: true, yo: yo, soloGerente: true,
      unidades: UNIDADES.filter(function (u) { return u.id === yo.sucursal; }),
      categorias: CATEGORIAS, medidas: MEDIDAS,
      hoy: hoyISO(), ahora: Utilities.formatDate(new Date(), ZONA, 'dd/MM/yyyy, HH:mm'),
      proveedores: [], productos: [], movimientos: [], ocupados: {}, avisoDatos: '' };
    try { var cg = catalogo(); g.proveedores = cg.proveedores; g.productos = cg.productos; }
    catch (e) { g.avisoDatos = 'No se pudo leer el catálogo de productos: ' + e.message; }
    try { g.enviados = misEnviados(yo); } catch (e) { g.enviados = []; }
    try { g.ocupados = ocupadosDe(nombreUnidad(yo.sucursal)); } catch (e) { g.ocupados = {}; }
    g.avisos = avisosInventario(yo);
    return g;                      // los avisos los pide la pantalla aparte

  }
  var r = {
    entro: true,
    yo: yo,
    unidades: UNIDADES,
    categorias: CATEGORIAS,
    hoy: Utilities.formatDate(new Date(), ZONA, 'yyyy-MM-dd'),
    ahora: Utilities.formatDate(new Date(), ZONA, 'dd/MM/yyyy, HH:mm'),
    medidas: MEDIDAS,
    proveedores: [],
    productos: [],
    movimientos: [],
    ocupados: {},
    avisoDatos: ''
  };
  try {
    var c = catalogo();
    r.proveedores = c.proveedores;
    r.productos = c.productos;
  } catch (e) {
    r.avisoDatos = 'No se pudo leer el catálogo de productos: ' + e.message;
  }
  try {
    r.ventas = ventasRango(r.hoy.slice(0, 8) + '01', r.hoy);
  } catch (e) {
    r.ventas = null;
  }
  try {
    r.movimientos = movimientos();
    r.ocupados = ocupados(r.movimientos);
  } catch (e) {
    r.avisoDatos = 'No se pudieron leer los movimientos: ' + e.message;
  }
  if (yo.confirma) {
    try { r.porConfirmar = listaPorConfirmar(yo); } catch (e) { r.porConfirmar = null; }
  }
  // un corte que un gerente ya envió (por confirmar) también ocupa su turno
  try {
    filasPorConfirmar().forEach(function (x) {
      if (x.tipo === 'ingreso' && enCamino(x))
        r.ocupados[x.sucursal + '|' + x.fecha + '|' + x.turno] = true;
    });
  } catch (e) {}
  if (yo.esAdmin) r.avisos = avisosInventario(yo);
  try { r.cierres = cierresCaja(7); } catch (e) { r.cierres = null; }
  if ((yo.soloMira && !yo.soloReportes) || yo.esAdmin || (yo.rol === 'registro' && yo.recibe)) {
    // los avisos NO se calculan aquí (son ~13 hojas): la pantalla los pide aparte y aparecen en la campana al llegar
  }
  return r;
}

/** Cada corte (sucursal y turno) entre dos fechas, sin anulados. Es lo que usa
 *  la venta por día y la gráfica. Lee la hoja Ingresos completa: no depende de los
 *  300 movimientos que carga la pantalla. */
function ventasRango(desde, hasta) {
  var h = libro().getSheetByName('Ingresos');
  var out = [];
  leeTodo(h, H_INGRESOS.length).forEach(function (s) {
    if (!s[0]) return;
    if (String(s[10]).toLowerCase() === 'sí') return;
    var f = fmtDia(s[5]);
    if (f < desde || f > hasta) return;
    var monto = Number(s[8]) || 0, e = Number(s[14]) || 0, t = Number(s[15]) || 0;
    if (!(e > 0 || t > 0)) e = monto;           // cortes de antes: todo en efectivo
    out.push({ id: String(s[0]), f: f, u: String(s[4] || ''), t: String(s[7] || ''), m: monto,
      e: e, j: t, c: Number(s[16]) || 0, k: t > 0 ? comisionFila(s) : 0 });
  });
  return out;
}

/** Para la ventana de venta por día, cuando se escoge otro periodo. */
function ventas(cred, desde, hasta) {
  verificaCaja(cred);
  if (!esFecha(desde) || !esFecha(hasta)) throw new Error('Escoja las fechas.');
  if (desde > hasta) { var x = desde; desde = hasta; hasta = x; }
  var dias = (aDate(hasta) - aDate(desde)) / 86400000;
  if (dias > 400) throw new Error('Escoja un periodo de un año o menos.');
  return { desde: desde, hasta: hasta, ventas: ventasRango(desde, hasta) };
}

/** La entrada a mano. Es la única que cuenta intentos fallidos. */
/** Se entra solo con el código (6 a 8 dígitos, uno distinto por persona).
 *  Si un código viejo lo comparten dos personas, se les pregunta quién es. */
function entrar(cred) {
  permiteAux();
  cred = cred || {};
  if (!String(cred.usuario || '').trim()) {
    var pin = String(cred.pin || '').trim();
    if (!/^\d{4,8}$/.test(pin)) throw new Error('Escriba su código de 6 a 8 números.');
    frenoIntentos();
    var suyos = usuariosCache().filter(function (u) { return u.activo && mismoPin(u.pin, pin); });
    if (!suyos.length) { cuentaFallo(); throw new Error('Ese código no es de nadie. Revíselo.'); }
    if (suyos.length > 1)
      return { entro: false, elegir: suyos.map(function (u) { return u.nombre; }) };
    cred = { usuario: suyos[0].nombre, pin: pin };
  }
  quien(cred, true);
  var d = datos(cred);
  d.cred = { usuario: cred.usuario, pin: String(cred.pin) };
  return d;
}

/* Sin nombre no se pueden contar los intentos por persona: se cuentan todos juntos
 * en ventanas de 10 minutos. 40 códigos equivocados en 10 minutos cierran la entrada
 * por el resto de esa ventana (con 6 dígitos hay un millón de combinaciones). */
function ventanaIntentos() {
  var d = new Date();
  return 'fallos_cod_' + Utilities.formatDate(d, ZONA, 'yyyyMMddHH') + Math.floor(d.getMinutes() / 10);
}
function frenoIntentos() {
  var n = Number(CacheService.getScriptCache().get(ventanaIntentos()) || 0);
  if (n >= 40) throw new Error('Demasiados códigos equivocados. Espere unos minutos.');
}
function cuentaFallo() {
  var c = CacheService.getScriptCache(), k = ventanaIntentos();
  c.put(k, String(Number(c.get(k) || 0) + 1), 900);
}

/** Código válido: 6 a 8 números y que no lo tenga otra persona activa. */
function revisaCodigo(pin, nombre) {
  if (!/^\d{6,8}$/.test(pin)) throw new Error('El código debe ser de 6 a 8 números.');
  if (/^(\d)\1+$/.test(pin) || '0123456789'.indexOf(pin) >= 0 || '9876543210'.indexOf(pin) >= 0)
    throw new Error('Ese código es muy fácil de adivinar. Escoja otro.');
  filasUsuarios().forEach(function (u) {
    if (u.activo && u.nombre.toLowerCase() !== String(nombre || '').toLowerCase() &&
        mismoPin(u.pin, pin))
      throw new Error('Ese código ya lo tiene otra persona. Escoja otro.');
  });
}

/** {"Centro|2026-09-26|AM": true, ...} para no ir al servidor a preguntarlo. */
function ocupados(movs) {
  var o = {};
  (movs || []).forEach(function (m) {
    if (m.clase === 'ingreso' && !m.anulado)
      o[m.unidad + '|' + m.fecha + '|' + m.turno] = true;
  });
  return o;
}

/** Los turnos ya tomados de una sucursal (oficiales y por confirmar), 40 días. */
function ocupadosDe(sucursal) {
  var o = {}, desde = Utilities.formatDate(new Date(Date.now() - 40 * 86400000), ZONA,
    'yyyy-MM-dd');
  ultimas(libro().getSheetByName('Ingresos'), 11, 3000).forEach(function (s) {
    if (String(s[4]) !== sucursal || String(s[10]).toLowerCase() === 'sí') return;
    var f = fmtDia(s[5]);
    if (f >= desde) o[sucursal + '|' + f + '|' + s[7]] = true;
  });
  filasPorConfirmar().forEach(function (x) {
    if (x.tipo === 'ingreso' && x.sucursal === sucursal && x.estado === EST_PEND)
      o[sucursal + '|' + x.fecha + '|' + x.turno] = true;
  });
  return o;
}

/** `max` es cuántos movimientos se devuelven. La app pide los últimos 300; el
 *  reporte del mes pide muchos más, porque tiene que traer el mes completo. */
function movimientosCalc(max) {
  max = Number(max) || 300;
  var ss = libro();
  var out = [];
  var TOPE = Math.max(600, max * 2);
  var g = ss.getSheetByName('Gastos');
  if (g && g.getLastRow() > 1) {
    var ig = Math.max(2, g.getLastRow() - TOPE + 1);
    // la hoja puede no tener todavía las columnas nuevas: se lee lo que haya
    var ancho = Math.min(H_GASTOS.length, Math.max(g.getLastColumn(), 14));
    var vg = g.getRange(ig, 1, g.getLastRow() - ig + 1, ancho).getValues();
    for (var a = vg.length - 1; a >= 0 && out.length < max * 2; a--) {
      var r = vg[a];
      if (!r[0]) continue;
      var fg = fmtDia(r[6]);
      out.push({ clase: 'gasto', id: String(r[0]), sello: fmtSello(r[1]),
        quien: String(r[2] || ''), unidad: String(r[4] || ''),
        titulo: String(r[5] || ''), fecha: fg, mes: fg.slice(0, 7),
        monto: Number(r[8]) || 0, nota: String(r[9] || ''),
        anulado: String(r[10]).toLowerCase() === 'sí',
        motivo: String(r[11] || ''),
        proveedor: String(r[14] || ''), factura: String(r[15] || ''),
        articulos: Number(r[16]) || 0 });
    }
  }
  var i = ss.getSheetByName('Ingresos');
  if (i && i.getLastRow() > 1) {
    var ii = Math.max(2, i.getLastRow() - TOPE + 1);
    var anchoI = Math.min(H_INGRESOS.length, Math.max(i.getLastColumn(), 14));
    var vi = i.getRange(ii, 1, i.getLastRow() - ii + 1, anchoI).getValues();
    for (var b = vi.length - 1; b >= 0 && out.length < max * 4; b--) {
      var s = vi[b];
      if (!s[0]) continue;
      var fi = fmtDia(s[5]);
      out.push({ clase: 'ingreso', id: String(s[0]), sello: fmtSello(s[1]),
        quien: String(s[2] || ''), unidad: String(s[4] || ''), fecha: fi,
        mes: fi.slice(0, 7), turno: String(s[7] || ''),
        titulo: 'Ingreso ' + String(s[7] || ''), monto: Number(s[8]) || 0,
        nota: String(s[9] || ''), anulado: String(s[10]).toLowerCase() === 'sí',
        motivo: String(s[11] || ''),
        // los cortes de antes, que no separaban el cobro, cuentan como efectivo
        efectivo: ((Number(s[14]) || 0) > 0 || (Number(s[15]) || 0) > 0)
          ? Number(s[14]) || 0 : Number(s[8]) || 0,
        tarjeta: Number(s[15]) || 0,
        caja: Number(s[16]) || 0, detalleCaja: String(s[17] || ''),
        comision: comisionFila(s), desglosado: true });
    }
  }
  out.sort(function (x, y) {
    if (x.fecha !== y.fecha) return x.fecha < y.fecha ? 1 : -1;
    return x.sello < y.sello ? 1 : -1;
  });
  return out.slice(0, max);
}
/** Los últimos ingresos y gastos: se arma una vez por consulta (varias pantallas lo piden muchas veces). */
function movimientos(max) {
  var k = 'mov:' + (Number(max) || 300);
  return _MEMO[k] || (_MEMO[k] = movimientosCalc(max));
}

function fmtSello(v) {
  if (v instanceof Date) return Utilities.formatDate(v, ZONA, 'yyyy-MM-dd HH:mm');
  return String(v || '');
}
function fmtDia(v) {
  if (v instanceof Date) return Utilities.formatDate(v, ZONA, 'yyyy-MM-dd');
  return String(v || '');
}
function nombreUnidad(id) {
  if (!_LUG) extiendeLugares();
  for (var k = 0; k < UNIDADES.length; k++)
    if (UNIDADES[k].id === id) return UNIDADES[k].nombre;
  return String(id);
}
function unidadPorId(id) {
  if (!_LUG) extiendeLugares();
  for (var k = 0; k < UNIDADES.length; k++)
    if (UNIDADES[k].id === id) return UNIDADES[k];
  return null;
}
function esFecha(s) { return /^\d{4}-\d{2}-\d{2}$/.test(String(s)); }
function aDate(s) {
  var p = String(s).split('-');
  return new Date(+p[0], +p[1] - 1, +p[2], 12, 0, 0);
}
function dinero(n) {
  n = Number(n) || 0;
  return (n < 0 ? '-' : '') + 'Q' +
    Math.abs(n).toFixed(2).replace(/\B(?=(\d{3})+(?!\d))/g, ',');
}
/** 1 caja, 6 cajas, 2 galones, 3 unidades. */
/** «unidades» → «unidad», «bolsas» → «bolsa»: las presentaciones se guardan en singular. */
var SINGULAR = { unidades: 'unidad', sobres: 'sobre', botes: 'bote', paquetes: 'paquete', galones: 'galón', cajones: 'cajón',
  latas: 'lata', piezas: 'pieza', docenas: 'docena', rollos: 'rollo', sacos: 'saco', fardos: 'fardo' };
function singular(u) {
  u = String(u || '').trim().toLowerCase();
  if (!u) return u;
  if (SINGULAR[u]) return SINGULAR[u];
  if (/ones$/.test(u)) return u.slice(0, -4) + 'ón';
  if (/ces$/.test(u)) return u.slice(0, -3) + 'z';
  if (/[^aeiouáéíóú](es)$/.test(u) && /[dlnrjy]$/.test(u.slice(0, -2))) return u.slice(0, -2);
  if (/[aeiouáéíóú]s$/.test(u) && u.length > 3) return u.slice(0, -1);
  return u;
}
function plur(n, u) {
  u = String(u || '');
  if (!u || Math.abs(Number(n) || 0) === 1) return u;
  if (/[^aeiouáéíóú]es$|[aeiouáéíóú]s$/i.test(u) && u.length > 3) return u;   // ya viene en plural
  if (/ón$/i.test(u)) return u.slice(0, -2) + 'ones';
  if (/z$/i.test(u)) return u.slice(0, -1) + 'ces';
  if (/[aeiouáéíóú]$/i.test(u)) return u + 's';
  return u + 'es';
}
/** Cantidades: 6, 4.5, 0.75 — sin ceros de sobra. */
function num(n) {
  n = Math.round((Number(n) || 0) * 1000) / 1000;
  return String(n).replace(/\B(?=(\d{3})+(?!\d)(?:\.|$))/g, ',');
}
/** Para que un nombre con & o < no rompa el PDF. */
function htm(s) {
  return String(s == null ? '' : s).replace(/[&<>]/g, function (c) {
    return c === '&' ? '&amp;' : c === '<' ? '&lt;' : '&gt;';
  });
}
/** Los artículos comprados en un mes, sin los de facturas anuladas. */
/** Los artículos de un grupo de facturas. Se busca por el ID del gasto, nunca
 *  por el mes: así el desglose no depende de cómo haya quedado guardada esa
 *  columna, y siempre calza exactamente con los movimientos del reporte. */
function detalleDeGastos(movs) {
  var ids = {};
  (movs || []).forEach(function (m) {
    if (m.clase === 'gasto' && !m.anulado) ids[String(m.id)] = true;
  });
  var out = [];
  leeTodo(libro().getSheetByName('Gastos detalle'), H_DETALLE.length)
    .forEach(function (r) {
      if (!r[0] || !ids[String(r[0])]) return;
      // «Sí» = anulado; «Reemplazado» = renglón viejo de un traslado corregido
      var est = String(r[13]).toLowerCase();
      if (est === 'sí' || est === 'reemplazado') return;
      out.push({ id: String(r[0]), unidad: String(r[3] || ''),
        proveedor: String(r[4] || ''),
        categoria: String(r[7] || ''), producto: String(r[8] || ''),
        cantidad: Number(r[9]) || 0, medida: String(r[10] || ''),
        precio: Number(r[11]) || 0, subtotal: Number(r[12]) || 0,
        contenido: Number(r[14]) || 0, base: String(r[15] || ''),
        cantidadBase: Number(r[16]) || 0, costoBase: Number(r[17]) || 0 });
    });
  return out;
}

/* ════════════ REGISTRAR UN GASTO ════════════ */

/** Un gasto puede venir de dos maneras:
 *   · artículo por artículo — d.lineas = [{producto, cantidad, medida, precio,
 *     categoría}] y el total lo suma el servidor;
 *   · solo el total — d.categoria y d.monto, como se hacía antes.
 *  La hoja Gastos siempre lleva un renglón por factura (el total). El desglose
 *  se guarda aparte, en «Gastos detalle». */
function registrarGasto(d) {
  permiteAux();
  d = d || {};
  var yo = noMira(verifica(d.cred));
  if (yo.rol === 'gerente') throw new Error('Su usuario solo registra ingresos.');
  if (yo.limitado) d.unidad = yo.sucursal;     // el auxiliar y la CMO, solo su lugar
  var v = preparaGasto(d);
  if (yo.limitado) return guardaPorConfirmar(yo, 'gasto', v, d);   // queda por validar
  var r = escribeGasto(v, yo.nombre, new Date(), null);
  if (v.proveedor) { try { aseguraProveedor(v.proveedor, yo.nombre, v.unidad); } catch (e) {} }
  recuerdaPrecios(v.lineas);
  return r;
}

/** Revisa y limpia un gasto. No escribe nada. */
function preparaGasto(d) {
  var u = unidadPorId(d.unidad);
  if (!u) throw new Error('Escoja la unidad.');
  if (!esFecha(d.fechaPago)) throw new Error('Escoja la fecha de pago.');

  var proveedor = String(d.proveedor || '').trim().slice(0, 60);
  var factura = String(d.factura || '').trim().slice(0, 30);
  var mes = String(d.fechaPago).slice(0, 7);

  var lineas = [];
  (d.lineas || []).forEach(function (l) {
    var nombre = String((l && l.producto) || '').trim();
    if (!nombre) return;
    var cant = Math.round((Number(l.cantidad) || 0) * 1000) / 1000;
    if (!(cant > 0)) cant = 1;
    var precio = Math.round((Number(l.precio) || 0) * 100) / 100;
    var sub = Math.round(cant * precio * 100) / 100;
    if (!(sub > 0))
      throw new Error('Al artículo «' + nombre + '» le falta el precio.');
    var contenido = Math.round((Number(l.contenido) || 0) * 1000) / 1000;
    var base = String(l.base || '').trim().slice(0, 20);
    if (!(contenido > 0) || !base) { contenido = 0; base = ''; }
    lineas.push({
      producto: nombre.slice(0, 70),
      cantidad: cant,
      medida: String(l.medida || 'unidad').trim().slice(0, 20),
      precio: precio,
      subtotal: sub,
      categoria: CATEGORIAS.indexOf(l.categoria) >= 0 ? l.categoria : '',
      proveedor: proveedor,
      contenido: contenido,
      base: base,
      // cuántas libras (o litros, o unidades) entraron de verdad, y a cómo
      cantidadBase: contenido > 0 ? Math.round(cant * contenido * 1000) / 1000 : 0,
      costoBase: contenido > 0 ? Math.round((precio / contenido) * 10000) / 10000 : 0
    });
  });

  var categoria = CATEGORIAS.indexOf(d.categoria) >= 0 ? d.categoria : '';
  var monto;

  if (lineas.length) {
    monto = 0;
    var cats = {}, cuantas = 0;
    lineas.forEach(function (l) {
      if (!l.categoria) l.categoria = categoria || 'Insumos de cocina';
      monto += l.subtotal;
      if (!cats[l.categoria]) { cats[l.categoria] = true; cuantas++; }
    });
    monto = Math.round(monto * 100) / 100;
    categoria = cuantas === 1 ? lineas[0].categoria : VARIAS;

    var declarado = Math.round((Number(d.total) || 0) * 100) / 100;
    if (declarado > 0 && Math.abs(declarado - monto) > 1) {
      throw new Error('Los artículos suman ' + dinero(monto) + ' y el total de ' +
        'la factura dice ' + dinero(declarado) + '. Revise antes de guardar, o ' +
        'deje el total de la factura en blanco.');
    }
  } else {
    if (!categoria) throw new Error('Escoja el tipo de gasto.');
    monto = Math.round((Number(d.monto) || 0) * 100) / 100;
  }

  if (!(monto > 0)) throw new Error('El monto debe ser mayor que cero.');
  if (monto > 500000) throw new Error('Ese monto parece equivocado. Revíselo.');
  return { unidad: u.id, fechaPago: d.fechaPago, proveedor: proveedor, factura: factura,
    mes: mes, lineas: lineas, categoria: categoria, monto: monto,
    nota: String(d.nota || '').slice(0, 140) };
}

/** Escribe un gasto ya revisado. `conf` = {por, en} cuando lo confirmó alguien. */
function escribeGasto(v, nombre, ahora, conf) {
  var u = unidadPorId(v.unidad), lineas = v.lineas, mes = v.mes;
  var lock = LockService.getScriptLock();
  lock.waitLock(15000); _LEE = {};
  try {
    var h = libro().getSheetByName('Gastos');
    if (!h) throw new Error('Falta la hoja Gastos. Corra la función instalar.');
    var id = 'G' + Utilities.formatDate(new Date(), ZONA, 'yyyyMMddHHmmss') +
      Math.floor(Math.random() * 900 + 100);
    var pago = aDate(v.fechaPago);

    h.appendRow([id, ahora, nombre, '', u.nombre, v.categoria, pago, mes, v.monto,
      v.nota, 'No', '', '', '', v.proveedor, v.factura, lineas.length,
      conf ? conf.por : '', conf ? conf.en : '']);
    // el mes va como TEXTO: si no, Google lo convierte en fecha y ninguna
    // fórmula que compare contra «2026-09» lo vuelve a encontrar
    h.getRange(h.getLastRow(), 8).setNumberFormat('@').setValue(mes);

    if (lineas.length) {
      var hd = hojaCat('Gastos detalle', H_DETALLE);
      var filas = lineas.map(function (l) {
        return [id, ahora, nombre, u.nombre, v.proveedor, pago, mes,
          l.categoria, l.producto, l.cantidad, l.medida, l.precio, l.subtotal,
          'No', l.contenido || '', l.base || '', l.cantidadBase || '',
          l.costoBase || ''];
      });
      var d0 = hd.getLastRow() + 1;
      hd.getRange(d0, 7, filas.length, 1).setNumberFormat('@');
      hd.getRange(d0, 1, filas.length, H_DETALLE.length).setValues(filas);
    }
  } finally {
    lock.releaseLock();
  }
  return { ok: true,
    mensaje: 'Gasto de ' + dinero(v.monto) + ' en ' + u.nombre +
      (lineas.length ? ' · ' + lineas.length +
        (lineas.length === 1 ? ' artículo' : ' artículos') : '') + '.',
    nuevo: { clase: 'gasto', id: id, sello: fmtSello(ahora), quien: nombre,
      unidad: u.nombre, titulo: v.categoria, fecha: v.fechaPago, mes: mes,
      monto: v.monto, nota: v.nota, anulado: false,
      proveedor: v.proveedor, factura: v.factura, articulos: lineas.length },
    lineas: lineas };
}

/** El desglose de una factura, para verlo desde el teléfono. */
function detalleGasto(cred, id) {
  verificaCaja(cred);
  id = String(id || '');
  var ss = libro();
  var out = [];
  leeTodo(ss.getSheetByName('Gastos detalle'), H_DETALLE.length)
    .forEach(function (r) {
      if (String(r[0]) !== id) return;
      if (String(r[13]).toLowerCase() === 'reemplazado') return;
      out.push({ producto: String(r[8] || ''), cantidad: Number(r[9]) || 0,
        medida: String(r[10] || ''), precio: Number(r[11]) || 0,
        subtotal: Number(r[12]) || 0, categoria: String(r[7] || ''),
        contenido: Number(r[14]) || 0, base: String(r[15] || ''),
        cantidadBase: Number(r[16]) || 0, costoBase: Number(r[17]) || 0 });
    });
  return out;
}

/* ════════════ REGISTRAR UN INGRESO ════════════ */

function registrarIngreso(d) {
  d = d || {};
  var yo = noMira(verifica(d.cred));
  if (yo.limitado) throw new Error('Su usuario tiene acceso limitado: solo registra gastos.');
  if (yo.rol === 'gerente') d.unidad = yo.sucursal;
  var v = preparaIngreso(d);
  if (yo.rol === 'gerente') return guardaPorConfirmar(yo, 'ingreso', v, d);
  return escribeIngreso(v, yo.nombre, new Date(), null);
}

/** Revisa y limpia un corte. No escribe nada. */
function preparaIngreso(d) {
  var u = unidadPorId(d.unidad);
  if (!u || !u.vende) throw new Error('Escoja una sucursal que venda al público.');
  if (!esFecha(d.fecha)) throw new Error('Escoja el día del ingreso.');
  if (d.turno !== 'AM' && d.turno !== 'PM') throw new Error('Escoja el turno.');

  /* El corte se captura desglosado: la venta del turno es efectivo + tarjeta, y
   * aparte lo que salió de la caja durante el turno. Los registros viejos que
   * solo traen el total se siguen aceptando. */
  var cen = function (x) { return Math.round((Number(x) || 0) * 100) / 100; };
  var efectivo = cen(d.efectivo), tarjeta = cen(d.tarjeta), caja = cen(d.gastoCaja);
  var desglosado = (efectivo > 0 || tarjeta > 0);
  var monto = desglosado ? Math.round((efectivo + tarjeta) * 100) / 100
                         : cen(d.monto);

  if (efectivo < 0 || tarjeta < 0 || caja < 0)
    throw new Error('Los montos no pueden ser negativos.');
  if (!(monto > 0)) throw new Error('El monto debe ser mayor que cero.');
  if (monto > 500000) throw new Error('Ese monto parece equivocado. Revíselo.');
  if (caja > monto)
    throw new Error('Los gastos de caja (' + dinero(caja) + ') no pueden ser ' +
      'mayores que la venta del turno (' + dinero(monto) + ').');
  if (desglosado && caja > efectivo)
    throw new Error('Salieron ' + dinero(caja) + ' de la caja pero solo entraron ' +
      dinero(efectivo) + ' en efectivo. Revise el corte.');
  var detCaja = String(d.detalleCaja || '').slice(0, 140);
  if (caja > 0 && detCaja.length < 3)
    throw new Error('Escriba en qué se gastó el dinero de la caja.');

  var hoy = Utilities.formatDate(new Date(), ZONA, 'yyyy-MM-dd');
  if (d.fecha > hoy)
    throw new Error('No se puede registrar el ingreso de un día que todavía no llega.');
  return { unidad: u.id, fecha: d.fecha, turno: d.turno, monto: monto,
    efectivo: efectivo, tarjeta: tarjeta, caja: caja, detalleCaja: detCaja,
    desglosado: desglosado, nota: String(d.nota || '').slice(0, 140) };
}

function repetidoTxt(v, u, rep) {
  return 'Ya hay un ingreso ' + v.turno + ' de ' + u.nombre + ' para el ' +
    v.fecha.split('-').reverse().join('/') + ', por ' + dinero(rep.monto) +
    ', anotado por ' + rep.quien + (rep.pendiente ? ' (está por confirmar)' : '') +
    '. Solo caben dos registros por día: AM y PM. Si está mal, pida que lo anulen ' +
    'y vuelva a hacerlo.';
}

/** Escribe un corte ya revisado. `conf` = {por, en} cuando lo confirmó alguien. */
function escribeIngreso(v, nombre, ahora, conf, excepto) {
  var u = unidadPorId(v.unidad);
  var lock = LockService.getScriptLock();
  lock.waitLock(15000); _LEE = {};
  try {
    var h = libro().getSheetByName('Ingresos');
    var repetido = buscaIngreso(h, u.nombre, v.fecha, v.turno, excepto);
    if (repetido) throw new Error(repetidoTxt(v, u, repetido));
    var id = 'I' + Utilities.formatDate(new Date(), ZONA, 'yyyyMMddHHmmss') +
      Math.floor(Math.random() * 900 + 100);
    // la comisión de NEONET se calcula sola con el % de ese día y queda grabada
    var pct = tasaComision(), com = v.desglosado ? comisionDe(v.tarjeta, pct) : 0;
    h.appendRow([id, ahora, nombre, '', u.nombre, aDate(v.fecha),
      String(v.fecha).slice(0, 7), v.turno, v.monto, v.nota, 'No', '', '', '',
      v.desglosado ? v.efectivo : '', v.desglosado ? v.tarjeta : '', v.caja,
      v.detalleCaja, conf ? conf.por : '', conf ? conf.en : '',
      conf && conf.valPor ? conf.valPor : '', conf && conf.valEn ? conf.valEn : '',
      conf && conf.recPor ? conf.recPor : '', conf && conf.recEn ? conf.recEn : '',
      com, v.desglosado ? r2(v.tarjeta - com) : '', v.desglosado ? pct : '']);
    h.getRange(h.getLastRow(), 7).setNumberFormat('@')
      .setValue(String(v.fecha).slice(0, 7));
    return { ok: true,
      mensaje: 'Ingreso ' + v.turno + ' de ' + u.nombre + ': ' + dinero(v.monto) +
        (v.caja > 0 ? ' · ' + dinero(v.caja) + ' de caja' : ''),
      nuevo: { clase: 'ingreso', id: id, sello: fmtSello(ahora), quien: nombre,
        unidad: u.nombre, titulo: 'Ingreso ' + v.turno, fecha: v.fecha,
        mes: String(v.fecha).slice(0, 7), turno: v.turno, monto: v.monto,
        nota: v.nota, anulado: false,
        efectivo: v.desglosado ? v.efectivo : 0, tarjeta: v.desglosado ? v.tarjeta : 0,
        caja: v.caja, detalleCaja: v.detalleCaja, desglosado: v.desglosado, comision: com } };
  } finally {
    lock.releaseLock();
  }
}

/** ¿Ya hay corte de ese turno? Revisa la hoja oficial y lo que los gerentes
 *  enviaron y todavía está por confirmar (menos `excepto`, el que se confirma). */
function buscaIngreso(h, sucursal, fecha, turno, excepto) {
  var v = leeTodo(h, 11);
  for (var k = 0; k < v.length; k++) {
    if (String(v[k][4]) !== sucursal) continue;
    if (fmtDia(v[k][5]) !== fecha) continue;
    if (String(v[k][7]) !== turno) continue;
    if (String(v[k][10]).toLowerCase() === 'sí') continue;
    return { monto: Number(v[k][8]) || 0, quien: v[k][2] };
  }
  var p = null;
  filasPorConfirmar().forEach(function (x) {
    if (p || !enCamino(x) || x.tipo !== 'ingreso' || x.id === excepto) return;
    if (x.sucursal === sucursal && x.fecha === fecha && x.turno === turno)
      p = { monto: x.monto, quien: x.enviadoPor, pendiente: true };
  });
  return p;
}

/* ════════════ ANULAR ════════════ */

function anular(d) {
  d = d || {};
  var yo = noMira(verificaCaja(d.cred));
  var motivo = String(d.motivo || '').trim();
  if (motivo.length < 4) throw new Error('Escriba el motivo de la anulación.');

  var nombreHoja = d.clase === 'gasto' ? 'Gastos' : 'Ingresos';
  var lock = LockService.getScriptLock();
  lock.waitLock(15000); _LEE = {};
  try {
    var h = libro().getSheetByName(nombreHoja);
    var ids = h.getRange(2, 1, Math.max(h.getLastRow() - 1, 1), 1).getValues();
    for (var k = 0; k < ids.length; k++) {
      if (String(ids[k][0]) === String(d.id)) {
        var fila = k + 2;
        if (String(h.getRange(fila, 11).getValue()).toLowerCase() === 'sí')
          throw new Error('Ese registro ya estaba anulado.');
        h.getRange(fila, 11, 1, 4).setValues([['Sí', motivo, yo.nombre, new Date()]]);
        var par = '';
        if (d.clase === 'gasto') {
          // un traslado es una pareja: lo que recibió la sucursal y lo que salió
          // de la bodega. Se anulan juntos o las cuentas quedan chuecas.
          par = parDeTraslado(String(d.id));
          // y el inventario regresa (antes de marcar el desglose como anulado)
          if (par) try { kardexDeshaceTraslado(/-B$/.test(String(d.id)) ? par : String(d.id),
            yo, motivo); } catch (e) { /* el gasto igual queda anulado */ }
          anulaDetalle(String(d.id));
          if (par) {
            for (var j = 0; j < ids.length; j++) {
              if (String(ids[j][0]) === par &&
                  String(h.getRange(j + 2, 11).getValue()).toLowerCase() !== 'sí') {
                h.getRange(j + 2, 11, 1, 4).setValues([['Sí', motivo, yo.nombre,
                  new Date()]]);
                anulaDetalle(par);
              }
            }
          }
        }
        return { ok: true, mensaje: par ? 'Traslado anulado (sucursal y bodega).'
          : 'Registro anulado.', id: d.id, par: par, motivo: motivo };
      }
    }
    throw new Error('No se encontró ese registro.');
  } finally {
    lock.releaseLock();
  }
}

/** Cuando se anula una factura, sus artículos dejan de contar en las compras. */
function anulaDetalle(id) {
  try {
    var h = libro().getSheetByName('Gastos detalle');
    if (!h || h.getLastRow() < 2) return;
    var n = h.getLastRow() - 1;
    var ids = h.getRange(2, 1, n, 1).getValues();
    var est = h.getRange(2, 14, n, 1).getValues();
    for (var k = 0; k < ids.length; k++) {
      // los renglones ya reemplazados por una corrección se quedan como están
      if (String(ids[k][0]) === id &&
          String(est[k][0]).toLowerCase() !== 'reemplazado')
        h.getRange(k + 2, 14).setValue('Sí');
    }
  } catch (e) { /* el total ya quedó anulado; esto es el desglose */ }
}

/* ════════════ TRASLADOS DE BODEGA ════════════ */

function hoyISO() { return Utilities.formatDate(new Date(), ZONA, 'yyyy-MM-dd'); }

/** Las sucursales que reciben producto (las que venden). */
function sucursalesDestino() {
  return UNIDADES.filter(function (u) { return u.vende; });
}
function nombreBodega() {
  var b = unidadPorId('bodega');
  return b ? b.nombre : 'Bodega Central';
}

/** «T…» es lo que recibió la sucursal y «T…-B» lo que salió de la bodega. */
function parDeTraslado(id) {
  id = String(id || '');
  if (!/^T\d/.test(id)) return '';
  return /-B$/.test(id) ? id.slice(0, -2) : id + '-B';
}

function exigeTraslado(cred) {
  var yo = quien(cred);
  if (yo.rol !== 'bodega' && !yo.esAdmin)
    throw new Error('Los traslados los registra la bodega o un administrador.');
  return yo;
}

function hojaCorrecciones(ss) {
  ss = ss || libro();
  var h = ss.getSheetByName('Traslados correcciones');
  if (h) {
    if (h.getMaxColumns() < H_CORRECCIONES.length || h.getLastColumn() < H_CORRECCIONES.length)
      encabezaAlFinal(h, 12, ['No. de traslado']);
    return h;
  }
  h = hojaLimpia(ss, 'Traslados correcciones', H_CORRECCIONES);
  h.getRange('A2:A').setNumberFormat('@');
  h.getRange('J:J').setNumberFormat('dd/mm/yyyy hh:mm');
  h.setColumnWidth(1, 120); h.setColumnWidth(3, 220); h.setColumnWidth(8, 260);
  return h;
}

/* ── CATÁLOGO DE BODEGA ──
 * Es la lista maestra de lo que sale de Bodega Central (la del PDF «Catálogo de
 * productos»): código, área, grupo y qué sucursales manejan cada producto. Aparte
 * dice CÓMO lo traslada Samuel:
 *   · «Se traslada por»: la unidad normal (caja, paquete, bolsa…).
 *   · «Suelto» y «Sueltos por unidad»: si a veces se manda abierto. La caja de papas
 *     trae 6 bolsas: se puede mandar la caja o bolsas sueltas. El té viene en
 *     cajitas de 20 y solo se manda la caja: sin «suelto».
 * El costo sale del catálogo de compras («Producto de compra» + «Rinde»: cuántas
 * unidades de traslado salen de cada unidad de compra; el té se compra por fardo y
 * cada fardo rinde 10 cajitas) o, si no está ligado, del «Costo manual». */
var H_BODEGA = ['Código', 'Producto', 'Área', 'Grupo', 'Centro', 'Almendras',
  'Parque', 'Categoría', 'Se traslada por', 'Suelto', 'Sueltos por unidad',
  'Producto de compra', 'Proveedor de compra', 'Rinde', 'Costo manual',
  'Activo', 'Actualizado en', 'Actualizado por', 'Almacenamiento', 'Estante',
  'Ideal Centro', 'Ideal Almendras', 'Ideal Parque', 'Ideal Bodega',
  'Se manda suelto', 'Se compra por', 'Tipo de medida',
  'Unidad menor', 'Menores por suelto', 'Peso de cada', 'Peso en',
  'Descripción', 'Desactivado por', 'Desactivado en',
  'Precio de la presentación', 'Precio cambiado en', 'Precio cambiado por', 'Proveedor secundario', 'Precio anterior'];
var PESOS = ['libra', 'onza', 'kilo', 'gramo'];
var AREAS_BODEGA = ['Barra', 'Cocina'];

/* Dónde vive cada producto en la bodega. «Almacenamiento» es la condición (frío o
 * seco) y «Estante» el lugar. Los químicos van en su propio estante, lejos de la
 * comida. Se pueden agregar estantes nuevos desde la app. */
var ALMACENES = ['Congelado', 'Refrigerado', 'Seco'];
var ESTANTES = ['Congelador', 'Cuarto frío', 'Estante 1 · Abarrotes',
  'Estante 2 · Empaques y desechables', 'Estante 3 · Bebidas',
  'Estante 4 · Limpieza y químicos', 'Estante 5 · Papelería'];

/** Lo que se le pone a un producto que todavía no tiene almacenamiento. */
function almacenDe(grupo) {
  if (/congel/i.test(grupo)) return 'Congelado';
  if (/refriger/i.test(grupo)) return 'Refrigerado';
  return 'Seco';
}
function estanteDe(almacen, categoria) {
  if (almacen === 'Congelado') return ESTANTES[0];
  if (almacen === 'Refrigerado') return ESTANTES[1];
  if (categoria === 'Empaque y desechables') return ESTANTES[3];
  if (categoria === 'Bebidas y gaseosas') return ESTANTES[4];
  if (categoria === 'Limpieza y químicos') return ESTANTES[5];
  if (categoria === 'Papelería y administración') return ESTANTES[6];
  return ESTANTES[2];
}
var UNIDADES_TRASLADO = ['caja', 'paquete', 'bolsa', 'fardo', 'unidad', 'lata',
  'bote', 'botella', 'galón', 'rollo', 'bandeja', 'barra', 'sobre', 'libra',
  'docena', 'cubeta', 'saco', 'onza', 'kilo', 'gramo'];

/** «cod|nombre|área|grupo|CAP|categoría|unidad|suelto|n». CAP = Centro, Almendras,
 *  Parque (1 = lo maneja). Solo se usa para llenar la hoja la primera vez. */
var CATALOGO_BODEGA = [
  "BOLXTJ|PAQUETE BOLSAS EXTRAJUMBO|Barra|Bolsas|111|Empaque y desechables|paquete||",
  "BOLJUM|PAQUETE BOLSAS JUMBO|Barra|Bolsas|111|Empaque y desechables|paquete||",
  "BOLMED|PAQUETE BOLSAS MEDIANA|Barra|Bolsas|111|Empaque y desechables|paquete||",
  "BOLROL|PAQUETE BOLSAS EN ROLLO|Barra|Bolsas|111|Empaque y desechables|paquete||",
  "BOLTOP|PAQUETE BOLSAS TOPO 500U|Barra|Bolsas|111|Empaque y desechables|paquete||",
  "BOLGRA|PAQUETE BOLSAS GRANDE|Barra|Bolsas|111|Empaque y desechables|paquete||",
  "PABOTO|PAQUETE BOLSAS PARA TONELERA|Barra|Bolsas|001|Empaque y desechables|paquete||",
  "AZPO6G|CAJA SOBRES AZÚCAR 6 GR|Barra|Consumible|111|Insumos de cocina|caja||",
  "CAFPOR|CAFÉ PORCIONADO 3 OZ|Barra|Consumible|111|Insumos de cocina|unidad||",
  "CHPO13|CHOCOLATE PORCIONADO 1.3 OZ|Barra|Consumible|110|Insumos de cocina|unidad||",
  "BOLCHO|BOLSA DE CHOCOLATE|Barra|Consumible|001|Insumos de cocina|bolsa||",
  "BOLLEC|BOLSA DE LECHE EN POLVO|Barra|Consumible|001|Insumos de cocina|bolsa||",
  "KETSOB|CAJA KETCHUP SOBRE 8 G 500 U|Barra|Consumible|111|Insumos de cocina|caja||",
  "PICSOB|CAJA PICAMÁS SOBRE 8 G 500 U|Barra|Consumible|111|Insumos de cocina|caja||",
  "KET1GL|KETCHUP 1 GAL|Barra|Consumible|101|Insumos de cocina|unidad||",
  "PIC1GL|PICAMÁS 1 GAL|Barra|Consumible|101|Insumos de cocina|unidad||",
  "SPLSOB|CAJA AZÚCAR SPLENDA SOBRE 1 G|Barra|Consumible|110|Insumos de cocina|caja||",
  "TECANE|CAJA TÉ CANELA 20 U|Barra|Consumible|111|Insumos de cocina|caja||",
  "TEMACA|CAJA TÉ MANZANA CANELA 20 U|Barra|Consumible|111|Insumos de cocina|caja||",
  "TEMANZ|CAJA TÉ MANZANILLA 20 U|Barra|Consumible|111|Insumos de cocina|caja||",
  "TEROJA|CAJA TÉ ROSA JAMAICA 20 U|Barra|Consumible|111|Insumos de cocina|caja||",
  "TEVERD|CAJA TÉ VERDE 20 U|Barra|Consumible|111|Insumos de cocina|caja||",
  "DOUMIN|CAJA CHICLE DOUBLEMINT 10 U|Barra|Consumible|111|Insumos de cocina|caja||",
  "CHOMMS|PAQUETE 18 U CHOCOLATE M&M'S 47.9G|Barra|Consumible|111|Insumos de cocina|paquete||",
  "VASCAF|VASO PAPEL 12 OZ|Barra|Empaques|111|Empaque y desechables|unidad||",
  "VASCA8|VASO PAPEL 8 OZ|Barra|Empaques|010|Empaque y desechables|unidad||",
  "TAVA12|TAPA VASO PAPEL 12 OZ|Barra|Empaques|111|Empaque y desechables|unidad||",
  "VA5ONZ|BOLSA VASO DESECHABLE 5 OZ|Barra|Empaques|111|Empaque y desechables|bolsa||",
  "BOVP12|BOLSA VASO PLÁSTICO 12 OZ|Barra|Empaques|001|Empaque y desechables|bolsa||",
  "TBVP12|BOLSA TAPA VASO PLÁSTICO 12 OZ|Barra|Empaques|001|Empaque y desechables|bolsa||",
  "PORPIZ|PAQUETE PORTAPIZZA DESECHABLE 100 U|Barra|Empaques|111|Empaque y desechables|paquete||",
  "PORVA4|PORTAVASOS 4 CAVIDADES|Barra|Empaques|111|Empaque y desechables|unidad||",
  "PORVA2|PORTAVASOS 2 CAVIDADES|Barra|Empaques|111|Empaque y desechables|unidad||",
  "BOLKR6|BOLSA PAPEL KRAFT 6 LB|Barra|Empaques|111|Empaque y desechables|bolsa||",
  "BOLKRD|BOLSA PAPEL KRAFT DELIVERY|Barra|Empaques|111|Empaque y desechables|bolsa||",
  "PORCUB|PAQUETE PORTACUBIERTOS PAPEL KRAFT 100 U|Barra|Empaques|100|Empaque y desechables|paquete||",
  "VASHE4|VASO PARA HELADO 4 OZ|Barra|Empaques|001|Empaque y desechables|unidad||",
  "VASHE6|VASO PARA HELADO 6 OZ|Barra|Empaques|001|Empaque y desechables|unidad||",
  "PAQCUC|PAQUETE DE CUCHARITAS|Barra|Empaques|001|Empaque y desechables|paquete||",
  "SERMES|PAQUETE SERVILLETAS DE PAPEL|Barra|Descartable|100|Empaque y desechables|paquete||",
  "SERCOT|BOLSA SERVILLETAS WHITE COTTON 100 U|Barra|Descartable|111|Empaque y desechables|bolsa||",
  "FILCAP|PAQUETE FILTRO PARA CAFÉ PEQUEÑO 50 U|Barra|Descartable|110|Empaque y desechables|paquete||",
  "FILCAG|PAQUETE FILTRO PARA CAFÉ GRANDE 50 U|Barra|Descartable|001|Empaque y desechables|paquete||",
  "PAJFOR|BOLSA PAJILLA FORRADA 70 U|Barra|Descartable|111|Empaque y desechables|bolsa||",
  "REVOLV|BOLSA REVOLVEDORES 100 U|Barra|Descartable|111|Empaque y desechables|bolsa||",
  "TENNEG|BOLSA TENEDOR DESECHABLE 50 U|Barra|Descartable|111|Empaque y desechables|bolsa||",
  "CUCNEG|BOLSA CUCHILLO DESECHABLE 50 U|Barra|Descartable|111|Empaque y desechables|bolsa||",
  "JABMAN|JABÓN PARA MANOS|Barra|Descartable|100|Limpieza y químicos|unidad||",
  "RETE79|RECIBO TÉRMICO 79 MM|Barra|Suministro|111|Papelería y administración|unidad||",
  "RETENE|RECIBO TÉRMICO NEONET|Barra|Suministro|111|Papelería y administración|unidad||",
  "ROLPAP|ROLLO PAPEL PARA BAÑO|Barra|Suministro|100|Limpieza y químicos|rollo||",
  "PAPMAN|TOALLAS DE PAPEL PARA MANO|Barra|Suministro|100|Limpieza y químicos|unidad||",
  "AROMAT|AROMATIZANTE PARA BAÑO|Barra|Suministro|100|Limpieza y químicos|unidad||",
  "STIDEL|ROLLO STICKER DELIVERY|Barra|Suministro|111|Papelería y administración|rollo||",
  "MOSQDA|MOSQUICIDA|Barra|Suministro|101|Limpieza y químicos|unidad||",
  "ODORIZ|ODORIZANTE|Barra|Suministro|100|Limpieza y químicos|unidad||",
  "COCLAT|COCA COLA LATA|Barra|Gaseosas|111|Bebidas y gaseosas|unidad||",
  "COC25L|COCA COLA 2.5 L|Barra|Gaseosas|111|Bebidas y gaseosas|unidad||",
  "PEPLAT|PEPSI LATA|Barra|Gaseosas|111|Bebidas y gaseosas|unidad||",
  "PEP3LT|PEPSI 3 L|Barra|Gaseosas|111|Bebidas y gaseosas|unidad||",
  "SEVLAT|SEVEN UP LATA|Barra|Gaseosas|111|Bebidas y gaseosas|unidad||",
  "SEV3LT|SEVEN UP 3 L|Barra|Gaseosas|111|Bebidas y gaseosas|unidad||",
  "ORALAT|ORANGE LATA|Barra|Gaseosas|111|Bebidas y gaseosas|unidad||",
  "TIKLAT|TIKY LATA|Barra|Gaseosas|111|Bebidas y gaseosas|unidad||",
  "MIRLAT|MIRINDA LATA|Barra|Gaseosas|111|Bebidas y gaseosas|unidad||",
  "AGUPUR|AGUA PURA|Barra|Gaseosas|111|Bebidas y gaseosas|unidad||",
  "CALMOR|CALIFORNIA MORA UVA|Barra|Gaseosas|111|Bebidas y gaseosas|unidad||",
  "CALPIN|CALIFORNIA PIÑA COCO|Barra|Gaseosas|111|Bebidas y gaseosas|unidad||",
  "CALKIW|CALIFORNIA KIWI FRESA|Barra|Gaseosas|111|Bebidas y gaseosas|unidad||",
  "TEFRIO|TÉ FRÍO|Barra|Gaseosas|111|Bebidas y gaseosas|unidad||",
  "AZMO5L|BOLSA AZÚCAR BLANCA 2.5 KG|Cocina|Ambiente|110|Insumos de cocina|bolsa||",
  "FRIDUC|FRIJOL ENTERO COCIDO 1.8 KG|Cocina|Ambiente|100|Insumos de cocina|unidad||",
  "HAPO12|BOLSA HARINA PORCIONADA 12 LB|Cocina|Ambiente|111|Insumos de cocina|bolsa||",
  "HAPO07|BOLSA HARINA PORCIONADA 7 LB|Cocina|Ambiente|111|Insumos de cocina|bolsa||",
  "MEZPAN|BOLSA MEZCLA PARA PANQUEQUES|Cocina|Ambiente|100|Insumos de cocina|bolsa||",
  "MAN15L|MANÍA PORCIONADA 1.5 LB|Cocina|Ambiente|110|Insumos de cocina|unidad||",
  "BOTMAY|MAYONESA 880 G|Cocina|Ambiente|111|Insumos de cocina|unidad||",
  "MOSANA|MOSTAZA 496 G|Cocina|Ambiente|111|Insumos de cocina|unidad||",
  "BOLNAC|BOLSA DE NACHOS|Cocina|Ambiente|111|Insumos de cocina|bolsa||",
  "CHONUT|CHOCOLATE TIPO NUTELLA 1 KG|Cocina|Ambiente|100|Insumos de cocina|unidad||",
  "PASLAS|CAJA PASTA LASAÑA 400 G|Cocina|Ambiente|100|Insumos de cocina|caja||",
  "SAL1LB|BOLSA SAL GRUESA 275 G|Cocina|Ambiente|111|Insumos de cocina|bolsa||",
  "SALFIN|BOLSA SAL FINA 1 KG|Cocina|Ambiente|111|Insumos de cocina|bolsa||",
  "SALTOM|SALSA DE TOMATE RANCHERA|Cocina|Ambiente|110|Insumos de cocina|unidad||",
  "SEM3LB|SÉMOLA PORCIONADO|Cocina|Ambiente|111|Insumos de cocina|unidad||",
  "ESECRE|ESENCIA DE CREMA 1 GAL|Cocina|Ambiente|110|Insumos de cocina|unidad||",
  "ESEFRE|ESENCIA DE FRESA 1 GAL|Cocina|Ambiente|110|Insumos de cocina|unidad||",
  "ADERAN|ADEREZO RANCH|Cocina|Ambiente|110|Insumos de cocina|unidad||",
  "JARMAP|JARABE DE MAPPLE|Cocina|Ambiente|100|Insumos de cocina|unidad||",
  "PIMNEG|BOTE PIMIENTA NEGRA 16 OZ|Cocina|Ambiente|111|Insumos de cocina|bote||",
  "CHOSYR|CHOCOLATE SYRUP|Cocina|Ambiente|100|Insumos de cocina|unidad||",
  "BOSAPI|BOLSA SALSA PARA PIZZA|Cocina|Ambiente|111|Insumos de cocina|bolsa||",
  "BARMAN|BARRA DE MANTEQUILLA|Cocina|Ambiente|111|Insumos de cocina|barra||",
  "BOPOMO|BOLSITA PORCIONADA DE MOSH|Cocina|Ambiente|100|Insumos de cocina|bolsa||",
  "PAHAM4|BOLSA PAN HAMBURGUESA 4\"|Cocina|Congelado|111|Insumos de cocina|bolsa||",
  "TOC1LB|TOCINO|Cocina|Congelado|111|Insumos de cocina|unidad||",
  "TOPJAM|TIRAS DE JAMÓN|Cocina|Congelado|111|Insumos de cocina|unidad||",
  "TOPPEP|BOLSA PEPPERONI|Cocina|Congelado|111|Insumos de cocina|bolsa||",
  "TOPSAL|SALAMI|Cocina|Congelado|111|Insumos de cocina|unidad||",
  "TOPCHU|BOLSA TOPPING CHUNK|Cocina|Congelado|111|Insumos de cocina|bolsa||",
  "PAPA38|CAJA PAPAS FRITAS 3/8\"|Cocina|Congelado|111|Insumos de cocina|caja|bolsa|6",
  "QUEAME|QUESO AMERICANO AMARILLO|Cocina|Congelado|111|Insumos de cocina|unidad||",
  "QUEBLA|QUESO AMERICANO BLANCO|Cocina|Congelado|100|Insumos de cocina|unidad||",
  "PAPIMA|PAQUETE PIE DE MANZANA|Cocina|Congelado|011|Insumos de cocina|paquete||",
  "HELNAP|HELADO NAPOLITANO 1 GAL|Cocina|Congelado|100|Insumos de cocina|unidad||",
  "ACECAP|ACEITE PARA MASA|Cocina|Suministro cocina|111|Insumos de cocina|unidad||",
  "ACEFRE|ACEITE PARA FREIDORA|Cocina|Suministro cocina|111|Insumos de cocina|unidad||",
  "ALUMP1|PAQUETE BANDEJA ALUMINIO PEQUEÑO|Cocina|Suministro cocina|100|Empaque y desechables|paquete||",
  "ALUMP2|PAQUETE BANDEJA ALUMINIO GRANDE|Cocina|Suministro cocina|100|Empaque y desechables|paquete||",
  "GUADES|CAJA GUANTES DESECHABLES 500 U|Cocina|Suministro cocina|111|Limpieza y químicos|caja||",
  "LAVPLA|LAVAPLATOS|Cocina|Suministro cocina|111|Limpieza y químicos|unidad||",
  "PASVER|PASHTE VERDE|Cocina|Suministro cocina|111|Limpieza y químicos|unidad||",
  "ESPTRA|ESPONJA PARA TRASTES|Cocina|Suministro cocina|111|Limpieza y químicos|unidad||",
  "ESPOLL|ESPONJA PARA OLLA|Cocina|Suministro cocina|111|Limpieza y químicos|unidad||",
  "BOLBAS|ROLLO BOLSA PARA BASURA|Cocina|Suministro cocina|111|Limpieza y químicos|rollo||",
  "PIEVOL|PIEDRA VOLCÁNICA|Cocina|Suministro cocina|111|Limpieza y químicos|unidad||",
  "PAPENC|PAQUETE PAPEL ENCERADO 500 U|Cocina|Suministro cocina|111|Empaque y desechables|paquete||",
  "REDCAB|REDECILLAS PARA EL CABELLO|Cocina|Suministro cocina|111|Limpieza y químicos|unidad||",
  "DETPOL|DETERGENTE EN POLVO|Cocina|Suministro cocina|111|Limpieza y químicos|unidad||",
  "ACENEG|LATA ACEITUNA NEGRA 1 GAL|Cocina|Enlatado|111|Insumos de cocina|lata||",
  "MEL1GL|LATA MELOCOTÓN EN ALMIBAR 1 GAL|Cocina|Enlatado|101|Insumos de cocina|lata||",
  "PIÑ1GL|LATA PIÑA EN ALMIBAR 1 GAL|Cocina|Enlatado|111|Insumos de cocina|lata||",
  "SALPIZ|LATA SALSA TOMATE PARA PIZZA|Cocina|Enlatado|111|Insumos de cocina|lata||",
  "CON1OZ|CONTENEDOR CON TAPA 1 OZ|Cocina|Empaques cocina|100|Empaque y desechables|unidad||",
  "CAJ8PU|CAJA PARA PIZZA 8\"|Cocina|Empaques cocina|111|Empaque y desechables|caja||",
  "CA13PU|CAJA PARA PIZZA 13\"|Cocina|Empaques cocina|111|Empaque y desechables|caja||",
  "CA15PU|CAJA PARA PIZZA 15\"|Cocina|Empaques cocina|111|Empaque y desechables|caja||",
  "CA18PU|CAJA PARA PIZZA 18\"|Cocina|Empaques cocina|111|Empaque y desechables|caja||",
  "CAJHAM|CAJA PARA HAMBURGUESA|Cocina|Empaques cocina|111|Empaque y desechables|caja||",
  "CAJPAP|BOLSA CAJA PAPAS 50 U|Cocina|Empaques cocina|111|Empaque y desechables|bolsa||",
  "CON6X6|BOLSA CONTENEDOR 6X6|Cocina|Empaques cocina|111|Empaque y desechables|bolsa||",
  "CON8X8|BOLSA CONTENEDOR 8X8|Cocina|Empaques cocina|110|Empaque y desechables|bolsa||",
  "ENSPEQ|BOLSA 25 U ENSALADERAS|Cocina|Empaques cocina|111|Empaque y desechables|bolsa||",
  "EN3DOF|ENVASE 3 OZ|Cocina|Empaques cocina|100|Empaque y desechables|unidad||",
  "TAP3OZ|TAPA ENVASE 3 OZ|Cocina|Empaques cocina|100|Empaque y desechables|unidad||",
  "FILPLA|FILM PLÁSTICO|Cocina|Empaques cocina|111|Empaque y desechables|unidad||",
  "MESPIZ|BOLSA GUARDAPIZZA|Cocina|Empaques cocina|110|Empaque y desechables|bolsa||",
  "VAPL16|VASO PLÁSTICO 16 OZ|Cocina|Empaques cocina|100|Empaque y desechables|unidad||",
  "TAPL16|TAPA VASO PLÁSTICO 16 OZ|Cocina|Empaques cocina|100|Empaque y desechables|unidad||",
  "VAPL12|VASO PLÁSTICO 12 OZ|Cocina|Empaques cocina|110|Empaque y desechables|unidad||",
  "TAPL12|TAPA VASO PLÁSTICO 12 OZ|Cocina|Empaques cocina|110|Empaque y desechables|unidad||",
  "TOMAN6|BANDEJA DE TOMATE 6U|Cocina|Refrigerado|111|Insumos de cocina|bandeja||",
  "LEC1GL|LECHE ENTERA 1 GAL|Cocina|Refrigerado|110|Insumos de cocina|unidad||",
  "CHAFRE|CHAMPIÑONES|Cocina|Refrigerado|111|Insumos de cocina|unidad||",
  "QUENAC|QUESO CHEDDAR|Cocina|Refrigerado|111|Insumos de cocina|unidad||",
  "QUEMOB|BARRA QUESO MOZARELLA|Cocina|Refrigerado|111|Insumos de cocina|barra||",
  "QUEMOA|BOLSA QUESO MOZARELLA|Cocina|Refrigerado|111|Insumos de cocina|bolsa||",
  "CRELAC|CREMA LACTOLAC|Cocina|Refrigerado|100|Insumos de cocina|unidad||",
  "BOLDEQ|BOLSA PALITOS DE QUESO|Cocina|Refrigerado|111|Insumos de cocina|bolsa||",
  "LEVADU|LEVADURA|Cocina|Refrigerado|111|Insumos de cocina|unidad||"
];

var _HB = null, _BOD = null;
function hojaBodega(ss) {
  if (_HB) return _HB;
  ss = ss || libro();
  var h = ss.getSheetByName('Productos de bodega');
  if (h) {
    _HB = h;
    // las hojas creadas antes no traían almacenamiento ni estante
    if (h.getLastColumn() < 20) encabezaAlFinal(h, 19, ['Almacenamiento', 'Estante']);
    // y tampoco los niveles ideales de inventario
    if (h.getLastColumn() < 24)
      encabezaAlFinal(h, 21, ['Ideal Centro', 'Ideal Almendras', 'Ideal Parque', 'Ideal Bodega']);
    // «se manda suelto» ahora es aparte de «cuántas trae» (que sirve para contar)
    if (h.getLastColumn() < 26) encabezaAlFinal(h, 25, ['Se manda suelto', 'Se compra por']);
    if (h.getLastColumn() < 27) encabezaAlFinal(h, 27, ['Tipo de medida']);   // unidad o peso
    // la presentación completa: caja → paquete → unidad, y cuánto pesa lo que se abre
    if (h.getLastColumn() < 31) encabezaAlFinal(h, 28, ['Unidad menor', 'Menores por suelto', 'Peso de cada', 'Peso en']);
    // una breve descripción del producto (la ven también los gerentes) y quién lo desactivó
    if (h.getLastColumn() < 34) { encabezaAlFinal(h, 32, ['Descripción', 'Desactivado por', 'Desactivado en']); h.getRange('AH:AH').setNumberFormat('dd/mm/yyyy hh:mm'); h.setColumnWidth(32, 300); }
    // el precio de lo que se compra (por caja, bolsa, fardo…); de ahí sale solo lo que cuesta cada unidad
    if (h.getLastColumn() < 37) { encabezaAlFinal(h, 35, ['Precio de la presentación', 'Precio cambiado en', 'Precio cambiado por']);
      h.getRange('AI:AI').setNumberFormat('"Q"#,##0.00'); h.getRange('AJ:AJ').setNumberFormat('dd/mm/yyyy hh:mm'); }
    if (h.getLastColumn() < 39) { encabezaAlFinal(h, 38, ['Proveedor secundario', 'Precio anterior']); h.getRange('AM:AM').setNumberFormat('"Q"#,##0.00'); }
    return h;
  }
  h = hojaLimpia(ss, 'Productos de bodega', H_BODEGA);
  var ahora = new Date();
  var filas = CATALOGO_BODEGA.map(function (x) {
    var c = x.split('|');
    return [c[0], c[1], c[2], c[3], c[4].charAt(0) === '1' ? 'Sí' : 'No',
      c[4].charAt(1) === '1' ? 'Sí' : 'No', c[4].charAt(2) === '1' ? 'Sí' : 'No',
      c[5], c[6], c[7] || '', c[8] ? Number(c[8]) : '', '', '', 1, '', 'Sí', ahora,
      'catálogo inicial', almacenDe(c[3]), estanteDe(almacenDe(c[3]), c[5]), '', '', '', '',
      c[7] ? 'Sí' : '', '', '', '', '', '', '', '', '', '', '', '', '', '', ''];
  });
  h.getRange(2, 1, filas.length, H_BODEGA.length).setValues(filas);
  h.getRange('O:O').setNumberFormat('"Q"#,##0.00');
  h.getRange('Q:Q').setNumberFormat('dd/mm/yyyy hh:mm');
  h.setColumnWidth(1, 80); h.setColumnWidth(2, 300); h.setColumnWidth(12, 220);
  h.setFrozenColumns(2);
  _HB = h;
  return h;
}

function siNo(v) { return String(v == null ? '' : v).trim().toLowerCase().charAt(0) === 's'; }

function filasBodega() {
  if (_BOD) return _BOD;
  _BOD = leeTodo(hojaBodega(), H_BODEGA.length).map(function (r, i) {
    return {
      fila: i + 2,
      codigo: String(r[0] || '').trim(),
      nombre: String(r[1] || '').trim(),
      area: String(r[2] || '').trim() || 'Cocina',
      grupo: String(r[3] || '').trim() || 'Otros',
      suc: { centro: siNo(r[4]), almendras: siNo(r[5]), parque: siNo(r[6]) },
      categoria: CATEGORIAS.indexOf(String(r[7]).trim()) >= 0 ? String(r[7]).trim()
        : 'Insumos de cocina',
      unidad: String(r[8] || '').trim() || 'unidad',
      suelto: String(r[9] || '').trim(),
      porUnidad: Number(r[10]) || 0,
      compra: String(r[11] || '').trim(),
      proveedor: String(r[12] || '').trim(),
      rinde: Number(r[13]) > 0 ? Number(r[13]) : 1,
      costoManual: Number(r[14]) || 0,
      activo: String(r[15]).trim().toLowerCase() !== 'no',
      almacen: ALMACENES.indexOf(String(r[18] || '').trim()) >= 0
        ? String(r[18]).trim() : almacenDe(String(r[3] || '')),
      estante: String(r[19] || '').trim(),
      ideal: { centro: Number(r[20]) || 0, almendras: Number(r[21]) || 0,
        parque: Number(r[22]) || 0, bodega: Number(r[23]) || 0 },
      // vacío = como antes: si tiene suelto configurado, se manda suelto
      mandaSuelto: String(r[24] || '').trim() ? siNo(r[24])
        : !!(String(r[9] || '').trim() && Number(r[10]) > 0),
      presCompra: String(r[25] || '').trim(),
      // «unidad» (cajas, bolsas…) o «peso» (libras, onzas…); vacío = se deduce de la unidad
      medida: /^(peso|unidad)$/i.test(String(r[26] || '').trim()) ? String(r[26]).trim().toLowerCase()
        : PESOS.indexOf(String(r[8] || '').trim().toLowerCase()) >= 0 ? 'peso' : 'unidad',
      menor: singular(r[27]),
      porSuelto: Number(r[28]) || 0,
      pesoCada: Number(r[29]) || 0,
      pesoEn: String(r[30] || '').trim().toLowerCase(),
      descripcion: String(r[31] || '').trim(),
      desactPor: String(r[32] || '').trim(),
      desactEn: r[33] instanceof Date ? r[33] : null,
      precioPres: Number(r[34]) > 0 ? Number(r[34]) : 0,
      precioEn: fmtSello(r[35]), precioEnRaw: r[35] instanceof Date ? r[35] : '', precioPor: String(r[36] || '').trim(),
      proveedor2: String(r[37] || '').trim(), precioAnt: Number(r[38]) > 0 ? Number(r[38]) : 0
    };
  }).map(function (b) {
    b.unidad = singular(b.unidad) || 'unidad'; b.suelto = singular(b.suelto);
    if (!b.estante) b.estante = estanteDe(b.almacen, b.categoria);
    // «CAJA TÉ CANELA 20 U» → trae 20 unidades (para contar sueltas), si nadie lo dijo
    if (!(b.suelto && b.porUnidad > 0) && b.unidad !== 'unidad') {
      var m = /(\d+)\s*(U|UN|UND|UNID|UNIDADES|PZ|PZS|PIEZAS)\b/i.exec(b.nombre);
      if (m && Number(m[1]) >= 2) { b.suelto = 'unidad'; b.porUnidad = Number(m[1]); b.deducido = true; }
    }
    if (b.medida === 'peso') { b.suelto = ''; b.porUnidad = 0; b.deducido = false; b.pesoCada = 0; }
    if (!(b.porUnidad > 0)) b.mandaSuelto = false;
    if (!(b.suelto && b.porUnidad > 0 && b.menor && b.porSuelto >= 2 && b.menor !== b.suelto)) { b.menor = ''; b.porSuelto = 0; }
    if (!(b.pesoCada > 0 && PESOS.indexOf(b.pesoEn) >= 0)) { b.pesoCada = 0; b.pesoEn = ''; }
    b.presentacion = presentacionTxt(b);
    return b;
  }).filter(function (b) { return b.nombre; });
  return _BOD;
}

function indiceCompras() {
  var idx = {};
  filasProductos().forEach(function (p) {
    idx[(p.nombre + '§' + p.proveedor).toLowerCase()] = p;
    if (!idx[p.nombre.toLowerCase()] || (p.precio > 0 && !(idx[p.nombre.toLowerCase()].precio > 0)))
      idx[p.nombre.toLowerCase()] = p;
  });
  return idx;
}

/** A cómo sale una unidad de traslado, y una suelta. */
function costoBodega(b, idx) {
  var c = 0, fuente = '', ligado = null;
  // el precio que puso la bodega (por caja, bolsa, fardo…) manda: de ahí sale el de cada unidad
  if (b.precioPres > 0) { c = b.precioPres / b.rinde; fuente = 'precio'; }
  if (b.compra) {
    ligado = idx[(b.compra + '§' + b.proveedor).toLowerCase()] ||
      idx[b.compra.toLowerCase()] || null;
    if (!(c > 0) && ligado && ligado.precio > 0) { c = ligado.precio / b.rinde; fuente = 'compra'; }
  }
  if (!(c > 0) && b.costoManual > 0) { c = b.costoManual; fuente = 'manual'; }
  c = Math.round(c * 10000) / 10000;
  return { costo: c, fuente: fuente, ligadoExiste: !!ligado,
    suelto: (b.suelto && b.porUnidad > 0 && c > 0)
      ? Math.round(c / b.porUnidad * 10000) / 10000 : 0 };
}

/** La llave de un renglón de traslado: el producto y la unidad en que se mandó. */
function llaveT(nombre, medida) {
  return (String(nombre || '').trim() + '|' + String(medida || '').trim()).toLowerCase();
}

function suelta(u) { return /(a|d|ón)$/i.test(String(u)) ? 'suelta' : 'suelto'; }

/** Los renglones de la tabla de Samuel, en el orden del catálogo. Un producto que
 *  se puede mandar suelto sale dos veces: por caja y por bolsa. */
function productosTraslado() {
  var idx = indiceCompras(), out = [];
  filasBodega().forEach(function (b) {
    if (!b.activo) return;
    var c = costoBodega(b, idx);
    var conSuelto = b.mandaSuelto && b.suelto && b.porUnidad > 0;
    out.push({ key: llaveT(b.nombre, b.unidad), codigo: b.codigo, nombre: b.nombre,
      area: b.area, grupo: b.grupo, categoria: b.categoria, unidad: b.unidad,
      costo: c.costo, suc: b.suc, factor: 1, porUnidad: b.porUnidad,
      nota: 'por ' + b.unidad + (conSuelto ? ' de ' + num(b.porUnidad) + ' ' +
        plur(b.porUnidad, b.suelto) : '') });
    if (conSuelto)
      out.push({ key: llaveT(b.nombre, b.suelto), codigo: b.codigo, nombre: b.nombre,
        area: b.area, grupo: b.grupo, categoria: b.categoria, unidad: b.suelto,
        costo: c.suelto, suc: b.suc, suelto: true, factor: 1 / b.porUnidad,
        porUnidad: b.porUnidad,
        nota: 'por ' + b.suelto + ' ' + suelta(b.suelto) });
  });
  return out;
}

/* ── lo que ve y usa el administrador ── */

function listaBodega(cred) {
  exigeAdmin(cred);
  var idx = indiceCompras();
  return {
    productos: filasBodega().map(function (b) {
      var c = costoBodega(b, idx);
      return { codigo: b.codigo, nombre: b.nombre, area: b.area, grupo: b.grupo,
        suc: b.suc, categoria: b.categoria, unidad: b.unidad, suelto: b.suelto,
        porUnidad: b.porUnidad, compra: b.compra, proveedor: b.proveedor,
        rinde: b.rinde, costoManual: b.costoManual, activo: b.activo,
        costo: c.costo, costoSuelto: c.suelto, fuente: c.fuente,
        ligadoExiste: c.ligadoExiste, almacen: b.almacen, estante: b.estante,
        ideal: b.ideal, mandaSuelto: b.mandaSuelto, presCompra: b.presCompra,
        deducido: !!b.deducido, medida: b.medida };
    }),
    compras: filasProductos().filter(function (p) { return p.activo; })
      .map(function (p) { return { nombre: p.nombre, proveedor: p.proveedor,
        medida: p.medida, precio: p.precio }; })
      .sort(function (a, b) { return a.nombre.localeCompare(b.nombre, 'es'); }),
    categorias: CATEGORIAS,
    unidades: UNIDADES_TRASLADO,
    areas: AREAS_BODEGA,
    almacenes: ALMACENES,
    estantes: ESTANTES
  };
}

/** Crea o cambia un producto de bodega. `p.original` es el código anterior. */
function guardarProductoBodega(cred, p) {
  var yo = exigeAdmin(cred);
  p = p || {};
  var codigo = String(p.codigo || '').trim().toUpperCase().slice(0, 12);
  var original = String(p.original || '').trim().toUpperCase();
  var nombre = String(p.nombre || '').trim().slice(0, 80);
  var area = AREAS_BODEGA.indexOf(p.area) >= 0 ? p.area : 'Cocina';
  var grupo = String(p.grupo || '').trim().slice(0, 40) || 'Otros';
  var suc = p.suc || {};
  var categoria = CATEGORIAS.indexOf(p.categoria) >= 0 ? p.categoria : 'Insumos de cocina';
  var unidad = String(p.unidad || '').trim().toLowerCase().slice(0, 20) || 'unidad';
  var suelto = String(p.suelto || '').trim().toLowerCase().slice(0, 20);
  var porUnidad = Math.round((Number(p.porUnidad) || 0) * 1000) / 1000;
  var compra = String(p.compra || '').trim().slice(0, 70);
  var proveedor = compra ? String(p.proveedor || '').trim().slice(0, 60) : '';
  var rinde = Number(p.rinde) > 0 ? Math.round(Number(p.rinde) * 1000) / 1000 : 1;
  var costoManual = Math.max(0, Math.round((Number(p.costoManual) || 0) * 100) / 100);
  var activo = p.activo === false ? 'No' : 'Sí';
  var precioNuevo = p.precioPres == null ? null : Math.max(0, Math.round((Number(String(p.precioPres).replace(/[Q,\s]/g, '')) || 0) * 100) / 100);
  var almacen = ALMACENES.indexOf(p.almacen) >= 0 ? p.almacen : almacenDe(grupo);
  var estante = String(p.estante || '').trim().slice(0, 40) || estanteDe(almacen, categoria);
  var ideal = {};
  ['centro', 'almendras', 'parque', 'bodega'].forEach(function (k) {
    var n = Number(String((p.ideal || {})[k] == null ? '' : p.ideal[k]).replace(',', '.'));
    ideal[k] = isFinite(n) && n > 0 ? Math.round(n * 1000) / 1000 : '';
  });

  if (!/^[A-Z0-9Ñ]{2,12}$/.test(codigo))
    throw new Error('El código debe tener de 2 a 12 letras o números, sin espacios.');
  if (nombre.length < 2) throw new Error('Escriba el nombre del producto.');
  if (!suc.centro && !suc.almendras && !suc.parque)
    throw new Error('Marque al menos una sucursal que maneje el producto.');
  if (!suelto) porUnidad = 0;
  if (suelto) {
    if (suelto === unidad)
      throw new Error('Lo suelto tiene que ser distinto de «' + unidad + '».');
    if (!(porUnidad >= 2))
      throw new Error('Diga cuántas ' + plur(2, suelto) + ' trae cada ' + unidad + '.');
  }
  var mandaSuelto = !!(p.mandaSuelto && suelto && porUnidad >= 2);
  var presCompra = String(p.presCompra || '').trim().toLowerCase().slice(0, 20);
  if (compra) {
    var hay = false;
    filasProductos().forEach(function (x) {
      if (x.nombre.toLowerCase() === compra.toLowerCase() &&
          x.proveedor.toLowerCase() === proveedor.toLowerCase()) hay = true;
    });
    if (!hay) throw new Error('No se encontró «' + compra + '» en el catálogo de compras.');
  }

  var lock = LockService.getScriptLock();
  lock.waitLock(15000); _LEE = {};
  try {
    var h = hojaBodega();
    var todos = filasBodega(), destino = null;
    todos.forEach(function (x) {
      if (original && x.codigo.toUpperCase() === original) destino = x;
    });
    if (original && !destino) throw new Error('No se encontró el producto ' + original + '.');
    todos.forEach(function (x) {
      if (destino && x.fila === destino.fila) return;
      if (x.codigo.toUpperCase() === codigo)
        throw new Error('Ya hay un producto con el código ' + codigo + ': ' + x.nombre + '.');
    });
    var fila = [codigo, nombre, area, grupo, suc.centro ? 'Sí' : 'No',
      suc.almendras ? 'Sí' : 'No', suc.parque ? 'Sí' : 'No', categoria, unidad,
      suelto, porUnidad || '', compra, proveedor, rinde, costoManual || '', activo,
      new Date(), yo.nombre, almacen, estante, ideal.centro, ideal.almendras, ideal.parque,
      ideal.bodega, mandaSuelto ? 'Sí' : 'No', presCompra, PESOS.indexOf(unidad) >= 0 ? 'peso' : 'unidad',
      destino && suelto === destino.suelto ? destino.menor : '', destino && suelto === destino.suelto ? (destino.porSuelto || '') : '',
      destino ? (destino.pesoCada || '') : '', destino ? destino.pesoEn : '',
      destino ? destino.descripcion : '',
      activo === 'No' ? (destino && !destino.activo ? destino.desactPor : yo.nombre) : '',
      activo === 'No' ? (destino && !destino.activo ? (destino.desactEn || '') : new Date()) : '',
      precioNuevo != null ? (precioNuevo || '') : (destino && destino.precioPres ? destino.precioPres : ''),
      precioNuevo != null && precioNuevo !== (destino ? destino.precioPres : 0) ? new Date() : (destino ? (destino.precioEnRaw || '') : ''),
      precioNuevo != null && precioNuevo !== (destino ? destino.precioPres : 0) ? yo.nombre : (destino ? destino.precioPor : ''),
      destino ? destino.proveedor2 : '', destino ? (destino.precioAnt || '') : ''];
    if (destino) h.getRange(destino.fila, 1, 1, H_BODEGA.length).setValues([fila]);
    else h.appendRow(fila);
    _BOD = null;
  } finally {
    lock.releaseLock();
  }
  var r = listaBodega(cred);
  r.ok = true;
  r.mensaje = nombre + (original ? ' quedó guardado.' : ' se agregó al catálogo.');
  return r;
}

/* ── PRECIOS DE BODEGA ──
 * Cada producto lleva el precio de lo que se COMPRA (por caja, bolsa, fardo…) y cuántas unidades de traslado
 * rinde; lo que cuesta cada unidad (y cada suelta) se calcula solo. Los precios los ven y los cambian la
 * bodega y el administrador; los ven también contabilidad, finanzas y el director operativo. Los gerentes
 * solo los ven si el director operativo o el administrador activan «precios para gerentes». */
function veCostos(yo) {
  return !!(yo && (yo.esAdmin || yo.rol === 'bodega' || yo.rol === 'finanzas' || yo.rol === 'operaciones' || esContador(yo)));
}
function veCostosTraslado(yo) {
  return veCostos(yo) || !!(yo && yo.rol === 'gerente' && ajustesInv().preciosGerentes);
}
function r4(n) { return Math.round((Number(n) || 0) * 10000) / 10000; }
function listaPrecios(cred) {
  var yo = quien(cred);
  if (!veCostos(yo)) throw new Error('Los precios los ven la bodega, contabilidad y los directores.');
  var idx = indiceCompras();
  var productos = filasBodega().filter(function (b) { return b.activo; }).map(function (b) {
    var c = costoBodega(b, idx);
    var costoMenor = (b.menor && b.porSuelto >= 2 && c.suelto > 0) ? r4(c.suelto / b.porSuelto) : 0;
    return { codigo: b.codigo, nombre: b.nombre, area: b.area, grupo: b.grupo, unidad: b.unidad,
      suelto: b.suelto, porUnidad: b.porUnidad, menor: b.menor, porSuelto: b.porSuelto,
      presCompra: b.presCompra || b.unidad, rinde: b.rinde, precio: b.precioPres,
      costo: c.costo, costoSuelto: c.suelto, costoMenor: costoMenor, fuente: c.fuente,
      compra: b.compra, precioEn: b.precioEn, precioPor: b.precioPor };
  });
  return { productos: productos, sinPrecio: productos.filter(function (p) { return !(p.costo > 0); }).length,
    puedeEditar: yo.rol === 'bodega' || yo.esAdmin, preciosGerentes: ajustesInv().preciosGerentes,
    puedeInterruptor: !!(yo.esAdmin || yo.rol === 'operaciones'), unidades: UNIDADES_TRASLADO, puedeReiniciar: !!yo.esAdmin };
}
/** cambios = [{codigo, precio, rinde, presCompra}] — solo lo que cambió. */
function guardarPrecios(cred, cambios) {
  var yo = quien(cred);
  if (!(yo.rol === 'bodega' || yo.esAdmin)) throw new Error('Los precios los pone la bodega o el administrador.');
  cambios = cambios || [];
  if (!cambios.length) throw new Error('No hay precios por guardar.');
  var lock = LockService.getScriptLock();
  lock.waitLock(15000); _LEE = {};
  var n = 0, hechos = [];
  try {
    var h = hojaBodega(), porCod = {};
    filasBodega().forEach(function (b) { porCod[b.codigo.toUpperCase()] = b; });
    var ahora = new Date();
    cambios.forEach(function (c) {
      var b = porCod[String(c.codigo || '').toUpperCase()];
      if (!b) throw new Error('No se encontró el producto ' + c.codigo + '.');
      var precio = Number(String(c.precio == null ? '' : c.precio).replace(/[Q,\s]/g, ''));
      if (c.precio !== '' && c.precio != null && (!isFinite(precio) || precio < 0))
        throw new Error('El precio de «' + b.nombre + '» no es válido.');
      precio = r2(precio || 0);
      var rinde = c.rinde == null || c.rinde === '' ? b.rinde : Number(String(c.rinde).replace(',', '.'));
      if (!(rinde > 0) || rinde > 100000) throw new Error('«' + b.nombre + '»: diga cuántas ' + plur(2, b.unidad) + ' rinde cada presentación de compra.');
      rinde = Math.round(rinde * 1000) / 1000;
      var pc = c.presCompra == null ? b.presCompra : String(c.presCompra || '').trim().toLowerCase().slice(0, 20);
      if (precio !== b.precioPres) hechos.push({ nombre: b.nombre, pres: pc || b.presCompra || b.unidad, antes: b.precioPres || 0, ahora: precio });
      if (rinde !== b.rinde) h.getRange(b.fila, 14).setValue(rinde);
      if (pc !== b.presCompra) h.getRange(b.fila, 26).setValue(pc);
      if (precio !== b.precioPres || rinde !== b.rinde)
        h.getRange(b.fila, 35, 1, 3).setValues([[precio || '', ahora, yo.nombre]]);
      n++;
    });
    _BOD = null;
  } finally { lock.releaseLock(); }
  var avisados = []; try { avisados = avisaCambioPrecios(yo, hechos, 'Cambios de precios en bodega'); } catch (e) {}
  var r = listaPrecios(cred); r.ok = true;
  r.mensaje = (n === 1 ? 'Precio guardado.' : n + ' precios guardados.') + (avisados.length ? ' Se avisó por correo a ' + avisados.join(', ') + '.' : '');
  return r;
}
/** Enciende o apaga que los gerentes vean los precios de lo que reciben. */
function ponPreciosGerentes(cred, activo) {
  var yo = quien(cred);
  if (!(yo.esAdmin || yo.rol === 'operaciones')) throw new Error('Esto lo cambian el director operativo y el administrador.');
  var v = activo ? 'Sí' : 'No';
  var h = hojaInv('Ajustes de inventario');
  var filas = leeTodo(h, 1).map(function (r) { return String(r[0]).trim(); });
  var i = filas.indexOf('precios_gerentes');
  if (i >= 0) h.getRange(i + 2, 2).setValue(v);
  else h.appendRow(['precios_gerentes', v, 'Si los gerentes ven los precios de lo que reciben de bodega']);
  _AJ = null;
  try { CacheService.getScriptCache().remove('ajustes_inv'); } catch (e) {}
  return { ok: true, preciosGerentes: !!activo,
    mensaje: activo ? 'Los gerentes ya ven los precios de lo que reciben.' : 'Los gerentes ya no ven los precios.' };
}

/* ── TRASLADOS NUMERADOS ──
 * Cada vez que Samuel guarda, nace un traslado nuevo con su número (TR-0001,
 * TR-0002…): un viaje en la mañana y otro en la tarde son dos traslados. El
 * número va en la columna «No. de factura» de cada gasto que genera, así que todo
 * lo que es de un mismo traslado se encuentra por ese número. La hoja «Traslados»
 * lleva el consecutivo y quién lo hizo. */
var H_TRASLADOS = ['No. de traslado', 'Fecha', 'Hora', 'Registrado por', 'Registrado en'];

function hojaTraslados(ss) {
  ss = ss || libro();
  var h = ss.getSheetByName('Traslados');
  if (h) return h;
  h = hojaLimpia(ss, 'Traslados', H_TRASLADOS);
  h.getRange('A:C').setNumberFormat('@');
  h.getRange('E:E').setNumberFormat('dd/mm/yyyy hh:mm');
  h.setColumnWidth(1, 120);
  return h;
}

/** El que sigue. Se pide con el candado puesto, para que no salgan dos iguales. */
function siguienteNumero() {
  var max = 0;
  leeTodo(hojaTraslados(), 1).forEach(function (r) {
    var m = /^TR-(\d+)$/.exec(String(r[0]).trim());
    if (m) max = Math.max(max, Number(m[1]));
  });
  var n = String(max + 1);
  return 'TR-' + (n.length < 4 ? ('0000' + n).slice(-4) : n);
}

/** Los gastos de un traslado, sucursal por sucursal, con sus renglones. */
function trasladoPorNumero(numero) {
  numero = String(numero || '').trim();
  var ss = libro(), suc = {}, meta = { numero: numero, fecha: '', sello: '', quien: '' };
  var g = ss.getSheetByName('Gastos');
  if (!numero || !g || g.getLastRow() < 2) return { suc: suc, meta: meta };
  var desde = Math.max(2, g.getLastRow() - 6000);
  var ancho = Math.min(H_GASTOS.length, g.getLastColumn());
  var ids = {};
  g.getRange(desde, 1, g.getLastRow() - desde + 1, ancho).getValues().forEach(function (r, i) {
    var id = String(r[0] || '');
    if (!/^T\d/.test(id) || /-B$/.test(id)) return;
    if (String(r[15] || '').trim() !== numero) return;
    var sello = fmtSello(r[1]);
    if (!meta.sello || sello < meta.sello) {
      meta.fecha = fmtDia(r[6]); meta.sello = sello; meta.quien = String(r[2] || '');
    }
    if (String(r[10]).toLowerCase() === 'sí') return;
    var s = String(r[4] || '');
    suc[s] = { id: id, fila: desde + i, sello: sello, quien: String(r[2] || ''),
      monto: Number(r[8]) || 0, lineas: {} };
    ids[id] = s;
  });
  if (Object.keys(ids).length) {
    leeTodo(ss.getSheetByName('Gastos detalle'), H_DETALLE.length).forEach(function (r) {
      var s = ids[String(r[0])];
      if (!s) return;
      if (String(r[13]).toLowerCase() !== 'no') return;
      var nombre = String(r[8] || ''), medida = String(r[10] || '');
      suc[s].lineas[llaveT(nombre, medida)] = { producto: nombre,
        cantidad: Number(r[9]) || 0, medida: medida,
        precio: Number(r[11]) || 0, categoria: String(r[7] || '') };
    });
  }
  return { suc: suc, meta: meta };
}

/** La tabla para un traslado nuevo: siempre vacía. Sin precios. */
function pantallaTraslado(yo) {
  var sucs = sucursalesDestino();
  return {
    fecha: hoyISO(),
    sucursales: sucs.map(function (s) {
      return { id: s.id, nombre: s.nombre, marca: s.marca }; }),
    productos: productosTraslado().map(function (p) {
      return { key: p.key, codigo: p.codigo, nombre: p.nombre, unidad: p.unidad,
        area: p.area, grupo: p.grupo, suc: p.suc, nota: p.nota, suelto: !!p.suelto };
    }),
    esAdmin: !!(yo && yo.esAdmin)
  };
}

/** Para el administrador, desde Configuración. */
function verTraslado(cred) {
  return pantallaTraslado(exigeTraslado(cred));
}

/** Un traslado ya guardado, para verlo o corregirlo. */
function abrirTraslado(cred, numero) {
  var yo = exigeTraslado(cred);
  var t = trasladoPorNumero(numero);
  if (!t.meta.fecha) throw new Error('No se encontró el traslado ' + numero + '.');
  var catal = {};
  productosTraslado().forEach(function (p) { catal[p.key] = p; });
  var guardado = {}, extras = [], vistos = {};
  sucursalesDestino().forEach(function (s) {
    var x = t.suc[s.nombre];
    if (!x) return;
    Object.keys(x.lineas).forEach(function (k) {
      var l = x.lineas[k];
      if (!guardado[k]) guardado[k] = {};
      guardado[k][s.id] = l.cantidad;
      // lo que ya no está en el catálogo igual tiene que salir para corregirlo
      if (!catal[k] && !vistos[k]) {
        vistos[k] = true;
        extras.push({ key: k, codigo: '', nombre: l.producto, unidad: l.medida,
          area: 'Otros', grupo: 'Ya no está en el catálogo', nota: 'por ' + l.medida,
          suc: { centro: true, almendras: true, parque: true }, fuera: true });
      }
    });
  });
  return { numero: t.meta.numero, fecha: t.meta.fecha, hora: t.meta.sello.slice(11),
    quien: t.meta.quien, guardado: guardado, extras: extras,
    editable: t.meta.fecha === hoyISO() || yo.esAdmin };
}

/** Guarda un traslado.
 *  Sin d.numero: es uno NUEVO y recibe el número que sigue.
 *  Con d.numero: es una CORRECCIÓN de ese traslado; solo el mismo día y con motivo.
 *  d.cantidades = { 'caja papas fritas 3/8"|caja': { centro: 2, almendras: 1 } } */
function guardarTraslado(cred, d) {
  var yo = exigeTraslado(cred);
  d = d || {};
  var motivo = String(d.motivo || '').trim().slice(0, 200);
  var editando = String(d.numero || '').trim();
  var hoy = hoyISO();
  var sucs = sucursalesDestino();

  var lock = LockService.getScriptLock();
  lock.waitLock(20000); _LEE = {};
  try {
    var catal = {};
    productosTraslado().forEach(function (p) { catal[p.key] = p; });

    var ya = {}, fecha = hoy;
    if (editando) {
      var t = trasladoPorNumero(editando);
      if (!t.meta.fecha) throw new Error('No se encontró el traslado ' + editando + '.');
      if (t.meta.fecha !== hoy && !yo.esAdmin)
        throw new Error('El traslado ' + editando + ' es del ' + dia(t.meta.fecha) +
          '. Solo se puede corregir el mismo día; después, solo un administrador.');
      ya = t.suc; fecha = t.meta.fecha;
    }
    var antesDe = {};
    Object.keys(ya).forEach(function (s) {
      Object.keys(ya[s].lineas).forEach(function (k) { antesDe[k] = ya[s].lineas[k]; });
    });

    // lo que viene de la pantalla, limpio
    var pedido = {};
    var cant = d.cantidades || {};
    Object.keys(cant).forEach(function (k0) {
      var k = String(k0 || '').toLowerCase();
      var p = catal[k], viejo = antesDe[k];
      if (!p && !viejo) return;
      var nombre = p ? p.nombre : viejo.producto;
      var x = { k: k, p: p, viejo: viejo, s: {} };
      sucs.forEach(function (s) {
        var crudo = (cant[k0] || {})[s.id];
        var n = Number(String(crudo == null ? '' : crudo).replace(',', '.'));
        if (!isFinite(n) || n < 0)
          throw new Error('La cantidad de «' + nombre + '» para ' + s.nombre +
            ' no es válida.');
        n = Math.round(n * 1000) / 1000;
        if (n > 100000) throw new Error('La cantidad de «' + nombre + '» parece equivocada.');
        if (n > 0) x.s[s.id] = n;
      });
      pedido[k] = x;
    });

    var plan = [], cambios = [], sinPrecio = {};
    sucs.forEach(function (s) {
      var antes = ya[s.nombre] ? ya[s.nombre].lineas : {};
      var nuevas = [];
      Object.keys(pedido).forEach(function (k) {
        var x = pedido[k], n = x.s[s.id];
        if (!(n > 0)) return;
        var p = x.p, viejo = antes[k];
        if (p && !p.suc[s.id] && !(viejo && viejo.cantidad > 0))
          throw new Error(p.nombre + ' no se maneja en ' + s.nombre + '. Si ahora sí ' +
            'lo lleva, márquelo en Configuración → Productos de bodega.');
        var costo = p ? p.costo : (viejo ? viejo.precio : 0);
        // al corregir se respeta el precio con que se guardó
        if (viejo && viejo.precio > 0) costo = viejo.precio;
        var base = viejo || x.viejo;
        var nombre = p ? p.nombre : base.producto;
        var unidad = p ? p.unidad : base.medida;
        if (!(costo > 0)) sinPrecio[nombre + (p && p.suelto ? ' (' + unidad + ')' : '')] = true;
        nuevas.push({ k: k, producto: nombre, cantidad: n, medida: unidad,
          precio: costo, subtotal: Math.round(n * costo * 100) / 100,
          categoria: p ? p.categoria : (base.categoria || 'Insumos de cocina') });
      });

      var dif = [], llaves = {};
      Object.keys(antes).forEach(function (k) { llaves[k] = true; });
      nuevas.forEach(function (l) { llaves[l.k] = true; });
      Object.keys(llaves).forEach(function (k) {
        var a = antes[k] ? antes[k].cantidad : 0;
        var nl = null;
        nuevas.forEach(function (l) { if (l.k === k) nl = l; });
        var b = nl ? nl.cantidad : 0;
        if (Math.abs(a - b) > 0.0005)
          dif.push({ producto: nl ? nl.producto : antes[k].producto,
            unidad: nl ? nl.medida : antes[k].medida, antes: a, despues: b });
      });

      if (!ya[s.nombre] && !nuevas.length) return;
      if (ya[s.nombre] && !dif.length) return;
      plan.push({ s: s, previo: ya[s.nombre] || null, lineas: nuevas });
      if (editando) dif.forEach(function (x) {
        x.sucursal = s.nombre; x.id = ya[s.nombre] ? ya[s.nombre].id : ''; cambios.push(x); });
    });

    if (!plan.length)
      throw new Error(editando ? 'No cambió ninguna cantidad.'
        : 'Escriba al menos una cantidad antes de guardar.');
    if (editando && motivo.length < 4)
      throw new Error('Escriba el motivo de la corrección.');

    var ahora = new Date();
    var numero = editando || siguienteNumero();
    if (!editando)
      hojaTraslados().appendRow([numero, hoy, Utilities.formatDate(ahora, ZONA, 'HH:mm'),
        yo.nombre, ahora]);
    plan.forEach(function (x) {
      if (x.previo) corrigeTraslado(x, yo, ahora, motivo, fecha);
      else creaTraslado(x, yo, ahora, fecha, numero);
    });
    // el inventario: lo que cambió entra a la sucursal y sale de la bodega
    var kx = [];
    plan.forEach(function (x) {
      var antes = x.previo ? x.previo.lineas : {}, ahoraL = {};
      x.lineas.forEach(function (l) { ahoraL[l.k] = l; });
      var llaves = {};
      Object.keys(antes).forEach(function (k) { llaves[k] = true; });
      Object.keys(ahoraL).forEach(function (k) { llaves[k] = true; });
      Object.keys(llaves).forEach(function (k) {
        var dlt = (ahoraL[k] ? ahoraL[k].cantidad : 0) - (antes[k] ? antes[k].cantidad : 0);
        if (Math.abs(dlt) < 0.0005) return;
        var p = catal[k], f = p ? p.factor : 1, q = r3(dlt * f);
        var nom = p ? p.nombre : (ahoraL[k] || antes[k]).producto || (antes[k] || {}).producto;
        if (!p) return;                         // ya no está en el catálogo: sin código
        kx.push([ahora, x.s.nombre, p.codigo, nom, MOV.ENT_TR, q, numero, yo.nombre,
          editando ? 'corrección' : '']);
        kx.push([ahora, nombreBodega(), p.codigo, nom, MOV.SAL_TR, -q, numero, yo.nombre,
          'a ' + x.s.nombre]);
      });
    });
    escribeKardex(kx);

    if (cambios.length) {
      var hc = hojaCorrecciones();
      var filas = cambios.map(function (c) {
        return [fecha, c.sucursal, c.producto, c.unidad, c.antes, c.despues,
          Math.round((c.despues - c.antes) * 1000) / 1000, motivo, yo.nombre,
          ahora, c.id, numero];
      });
      var f0 = hc.getLastRow() + 1;
      hc.getRange(f0, 1, filas.length, 1).setNumberFormat('@');
      hc.getRange(f0, 1, filas.length, H_CORRECCIONES.length).setValues(filas);
    }

    var nombres = plan.map(function (x) { return x.s.nombre; });
    return {
      ok: true,
      numero: numero,
      mensaje: editando ? 'Traslado ' + numero + ' corregido.'
        : 'Traslado ' + numero + ' guardado: ' + nombres.join(', ') + '.',
      correcciones: cambios.length,
      // el precio solo lo ve el administrador; a la bodega se le avisa sin cifras
      sinPrecio: Object.keys(sinPrecio)
    };
  } finally {
    lock.releaseLock();
  }
}

/** Resume los renglones en la categoría del gasto, como en una factura. */
function categoriaDe(lineas) {
  var cats = {}, n = 0;
  lineas.forEach(function (l) { if (!cats[l.categoria]) { cats[l.categoria] = 1; n++; } });
  return n === 1 ? lineas[0].categoria : VARIAS;
}

function notaTraslado(lineas, s, numero) {
  var sin = lineas.filter(function (l) { return !(l.precio > 0); })
    .map(function (l) { return l.producto; });
  return ('Traslado ' + (numero ? numero + ' ' : '') + 'de bodega a ' + s.nombre +
    (sin.length ? ' · SIN PRECIO: ' + sin.join(', ') : '')).slice(0, 140);
}

function filasDetalleTraslado(id, lineas, yo, ahora, s, pago, mes, bodega) {
  var out = [];
  lineas.forEach(function (l) {
    out.push([id, ahora, yo.nombre, s.nombre, PROV_BODEGA, pago, mes, l.categoria,
      l.producto, l.cantidad, l.medida, l.precio, l.subtotal, 'No', '', '', '', '']);
  });
  lineas.forEach(function (l) {
    out.push([id + '-B', ahora, yo.nombre, bodega, PROV_BODEGA, pago, mes,
      CAT_TRASLADO, l.producto, -l.cantidad, l.medida, l.precio, -l.subtotal, 'No',
      '', '', '', '']);
  });
  return out;
}

/** Completa cada renglón con celdas vacías hasta el ancho de la hoja (si la hoja
 *  creció con columnas nuevas, los renglones viejos del código no deben fallar). */
function aAncho(filas, n) {
  return filas.map(function (f) { var r = f.slice(0, n); while (r.length < n) r.push(''); return r; });
}

function escribeDetalle(filas) {
  if (!filas.length) return;
  var hd = hojaCat('Gastos detalle', H_DETALLE);
  var d0 = hd.getLastRow() + 1;
  hd.getRange(d0, 7, filas.length, 1).setNumberFormat('@');
  hd.getRange(d0, 1, filas.length, H_DETALLE.length).setValues(filas);
}

function creaTraslado(x, yo, ahora, hoy, numero, ref) {
  var h = libro().getSheetByName('Gastos');
  if (!h) throw new Error('Falta la hoja Gastos. Corra la función instalar.');
  var s = x.s, lineas = x.lineas, bodega = nombreBodega();
  var id = 'T' + Utilities.formatDate(ahora, ZONA, 'yyyyMMddHHmmss') +
    Math.floor(Math.random() * 900 + 100) + s.id.charAt(0).toUpperCase();
  var pago = aDate(hoy), mes = hoy.slice(0, 7);
  var monto = 0;
  lineas.forEach(function (l) { monto += l.subtotal; });
  monto = Math.round(monto * 100) / 100;
  var factura = numero;

  var f0 = h.getLastRow() + 1;
  h.getRange(f0, 8, 2, 1).setNumberFormat('@');
  h.getRange(f0, 1, 2, H_GASTOS.length).setValues(aAncho([
    [id, ahora, yo.nombre, '', s.nombre, categoriaDe(lineas), pago, mes, monto,
      (ref ? 'Solicitud ' + ref + '. ' : '') + notaTraslado(lineas, s, numero), 'No', '', '',
      '', PROV_BODEGA, factura, lineas.length],
    [id + '-B', ahora, yo.nombre, '', bodega, CAT_TRASLADO, pago, mes, -monto,
      'Traslado ' + numero + ': salida de bodega a ' + s.nombre, 'No', '', '', '', PROV_BODEGA, factura,
      lineas.length]
  ], H_GASTOS.length));
  escribeDetalle(filasDetalleTraslado(id, lineas, yo, ahora, s, pago, mes, bodega));
}

/** Corregir no borra: los renglones viejos quedan marcados «Reemplazado» y la
 *  diferencia queda escrita en «Traslados correcciones». */
function corrigeTraslado(x, yo, ahora, motivo, hoy) {
  var ss = libro();
  var h = ss.getSheetByName('Gastos');
  var id = x.previo.id, par = id + '-B';
  var s = x.s, lineas = x.lineas, bodega = nombreBodega();
  var pago = aDate(hoy), mes = hoy.slice(0, 7);

  // las dos filas del gasto (sucursal y bodega)
  var desde = Math.max(2, h.getLastRow() - 1500);
  var ids = h.getRange(desde, 1, h.getLastRow() - desde + 1, 1).getValues();
  var fS = 0, fB = 0;
  ids.forEach(function (r, i) {
    if (String(r[0]) === id) fS = desde + i;
    if (String(r[0]) === par) fB = desde + i;
  });
  if (!fS) throw new Error('No se encontró el traslado de ' + s.nombre + '.');

  // renglones viejos → «Reemplazado»
  var hd = ss.getSheetByName('Gastos detalle');
  if (hd && hd.getLastRow() > 1) {
    var n = hd.getLastRow() - 1;
    var vi = hd.getRange(2, 1, n, 1).getValues();
    var ve = hd.getRange(2, 14, n, 1).getValues();
    for (var k = 0; k < n; k++) {
      var i2 = String(vi[k][0]);
      if ((i2 === id || i2 === par) && String(ve[k][0]).toLowerCase() === 'no')
        hd.getRange(k + 2, 14).setValue('Reemplazado');
    }
  }

  var nota = 'Corregido ' + Utilities.formatDate(ahora, ZONA, 'HH:mm') + ': ' + motivo;
  if (!lineas.length) {
    // se quitó todo lo de esa sucursal: el traslado queda anulado
    h.getRange(fS, 11, 1, 4).setValues([['Sí', nota.slice(0, 140), yo.nombre, ahora]]);
    if (fB) h.getRange(fB, 11, 1, 4).setValues([['Sí', nota.slice(0, 140), yo.nombre, ahora]]);
    return;
  }

  var monto = 0;
  lineas.forEach(function (l) { monto += l.subtotal; });
  monto = Math.round(monto * 100) / 100;
  h.getRange(fS, 6).setValue(categoriaDe(lineas));
  h.getRange(fS, 9, 1, 2).setValues([[monto,
    (notaTraslado(lineas, s, String(h.getRange(fS, 16).getValue() || '')) + ' · ' +
      nota).slice(0, 140)]]);
  h.getRange(fS, 17).setValue(lineas.length);
  if (fB) {
    h.getRange(fB, 9).setValue(-monto);
    h.getRange(fB, 17).setValue(lineas.length);
  }
  escribeDetalle(filasDetalleTraslado(id, lineas, yo, ahora, s, pago, mes, bodega));
}

/** dd/mm/aaaa */
function dia(iso) {
  var p = String(iso).split('-');
  return p[2] + '/' + p[1] + '/' + p[0];
}

/** El historial: un renglón por traslado (número, fecha y hora). Adentro, qué se
 *  mandó a cada sucursal y las correcciones. La bodega ve cantidades; el
 *  administrador además ve el costo. No hay forma de exportarlo desde la app. */
var FIRMAS_H = null;
function historialTraslados(cred, dias) { return historialTrasladosDe(exigeTraslado(cred), dias); }
function historialTrasladosDe(yo, dias) {
  FIRMAS_H = null;
  var verCosto = veCostos(yo), puedeEd = yo.rol === 'bodega' || yo.esAdmin;
  dias = Math.min(Math.max(Number(dias) || 31, 1), 120);
  var ss = libro();
  var hoy = hoyISO();
  var limite = Utilities.formatDate(new Date(Date.now() - (dias - 1) * 86400000),
    ZONA, 'yyyy-MM-dd');
  var sucs = sucursalesDestino();
  var idDe = {};
  sucs.forEach(function (s) { idDe[s.nombre] = s.id; });

  var porNum = {}, numDeId = {};
  function elT(n) {
    if (!porNum[n]) porNum[n] = { numero: n, fecha: '', sello: '', quien: '',
      productos: {}, costo: {}, articulos: {}, activas: 0, correcciones: [] };
    return porNum[n];
  }

  var g = ss.getSheetByName('Gastos');
  if (g && g.getLastRow() > 1) {
    var desde = Math.max(2, g.getLastRow() - 6000);
    var ancho = Math.min(H_GASTOS.length, g.getLastColumn());
    g.getRange(desde, 1, g.getLastRow() - desde + 1, ancho).getValues()
      .forEach(function (r) {
        var id = String(r[0] || '');
        if (!/^T\d/.test(id) || /-B$/.test(id)) return;
        var f = fmtDia(r[6]);
        if (f < limite) return;
        // los de antes de numerar traían «Traslado dd/mm/aaaa»: se agrupan igual
        var n = String(r[15] || '').trim() || ('Traslado ' + dia(f));
        var T = elT(n), sello = fmtSello(r[1]);
        numDeId[id] = n;
        if (!T.sello || sello < T.sello) {
          T.sello = sello; T.fecha = f; T.quien = String(r[2] || '');
        }
        if (String(r[10]).toLowerCase() === 'sí') return;
        var sid = idDe[String(r[4] || '')] || String(r[4] || '');
        T.costo[sid] = (T.costo[sid] || 0) + (Number(r[8]) || 0);
        T.activas++;
        T._vivos = T._vivos || {};
        T._vivos[id] = sid;
      });
  }
  if (Object.keys(numDeId).length) {
    leeTodo(ss.getSheetByName('Gastos detalle'), H_DETALLE.length).forEach(function (r) {
      var n = numDeId[String(r[0])];
      if (!n) return;
      var T = porNum[n];
      var sid = T._vivos && T._vivos[String(r[0])];
      if (!sid) return;
      if (String(r[13]).toLowerCase() !== 'no') return;
      var nombre = String(r[8] || ''), medida = String(r[10] || '');
      var k = llaveT(nombre, medida);
      if (!T.productos[k]) T.productos[k] = { nombre: nombre, unidad: medida, c: {}, total: 0, m: {} };
      var q = Number(r[9]) || 0;
      T.productos[k].m[sid] = (T.productos[k].m[sid] || 0) + (Number(r[12]) || 0);
      T.productos[k].c[sid] = (T.productos[k].c[sid] || 0) + q;
      T.productos[k].total += q;
      T.articulos[sid] = (T.articulos[sid] || 0) + 1;
    });
  }

  leeTodo(hojaCorrecciones(), H_CORRECCIONES.length).forEach(function (r) {
    var n = String(r[11] || '').trim() || numDeId[String(r[10] || '')];
    if (!n || !porNum[n]) return;
    porNum[n].correcciones.push({ sucursal: String(r[1] || ''),
      producto: String(r[2] || ''), unidad: String(r[3] || ''),
      antes: Number(r[4]) || 0, despues: Number(r[5]) || 0,
      motivo: String(r[7] || ''), quien: String(r[8] || ''), sello: fmtSello(r[9]) });
  });

  var out = Object.keys(porNum).map(function (n) {
    var T = porNum[n];
    var x = { numero: n, fecha: T.fecha, hora: T.sello.slice(11), quien: T.quien,
      anulado: T.activas === 0, editable: puedeEd && (T.fecha === hoy || yo.esAdmin) && /^TR-/.test(n),
      articulos: T.articulos,
      productos: Object.keys(T.productos).map(function (k) { return T.productos[k]; })
        .sort(function (a, b) { return a.nombre.localeCompare(b.nombre, 'es') ||
          a.unidad.localeCompare(b.unidad, 'es'); }),
      correcciones: T.correcciones.sort(function (a, b) {
        return a.sello < b.sello ? -1 : 1; }) };
    if (verCosto) {
      var cc = {}, tt = 0;
      Object.keys(T.costo).forEach(function (s) {
        cc[s] = Math.round(T.costo[s] * 100) / 100; tt += cc[s]; });
      x.costo = cc; x.costoTotal = Math.round(tt * 100) / 100;
    } else x.productos.forEach(function (p) { delete p.m; });
    x._sello = T.sello;
    if (/^TR-/.test(n)) {
      x.firmas = {};
      sucs.forEach(function (s) {
        var ls = x.productos.filter(function (p) { return p.c[s.id] > 0; }).map(function (p) { return { producto: p.nombre, medida: p.unidad, cantidad: p.c[s.id] }; });
        if (!ls.length) return;
        var e = estadoFirma({ numero: n, sucursal: s.nombre, lineas: ls }, FIRMAS_H || (FIRMAS_H = firmasGuardadas()));
        x.firmas[s.id] = { estado: e.estado, por: e.por || '', en: e.en || '', nota: e.nota || '', firma: e.firma || '' };
      });
    }
    return x;
  }).sort(function (a, b) { return a._sello < b._sello ? 1 : -1; });
  out.forEach(function (x) { delete x._sello; });
  return { traslados: out, sucursales: sucs.map(function (s) {
    return { id: s.id, nombre: s.nombre }; }), esAdmin: yo.esAdmin, verCosto: verCosto, desde: limite };
}

/** Lo que costaron los traslados de bodega: para el administrador, contabilidad y finanzas (y la bodega y el
 *  director operativo). Por sucursal, por mes y traslado por traslado, con lo que cuesta cada producto. */
function costosTraslados(cred, dias) {
  var yo = quien(cred);
  if (!veCostos(yo)) throw new Error('Los costos de los traslados los ven el administrador, contabilidad, la bodega y el director operativo.');
  var r = historialTrasladosDe(yo, dias || 93);
  var meses = {}, sucs = r.sucursales;
  r.traslados.forEach(function (t) {
    if (t.anulado || !t.costo) return;
    var m = String(t.fecha || '').slice(0, 7); if (!m) return;
    if (!meses[m]) { meses[m] = { mes: m, porSucursal: {}, total: 0, traslados: 0 }; sucs.forEach(function (s) { meses[m].porSucursal[s.id] = 0; }); }
    Object.keys(t.costo).forEach(function (sid) { meses[m].porSucursal[sid] = (meses[m].porSucursal[sid] || 0) + t.costo[sid]; });
    meses[m].total += t.costoTotal || 0; meses[m].traslados++;
  });
  var lista = Object.keys(meses).sort().reverse().map(function (m) {
    var x = meses[m]; x.total = r2(x.total); Object.keys(x.porSucursal).forEach(function (k) { x.porSucursal[k] = r2(x.porSucursal[k]); }); return x; });
  r.meses = lista; r.preciosGerentes = ajustesInv().preciosGerentes;
  r.puedeInterruptor = !!(yo.esAdmin || yo.rol === 'operaciones');
  return r;
}

/* ════════════ PANEL DE ADMINISTRADOR ════════════ */

function listaUsuarios(cred) {
  exigeAdmin(cred);
  return filasUsuarios().map(function (u) {
    return { nombre: u.nombre, rol: u.rol, activo: u.activo, pin: u.pin,
      sucursal: u.sucursal, confirma: u.confirma };
  });
}

/** Crea o modifica. `original` es el nombre anterior cuando se está editando. */
function guardarUsuario(cred, u) {
  var yo = exigeAdmin(cred);
  u = u || {};
  var nombre = String(u.nombre || '').trim();
  var pin = String(u.pin || '').trim();
  var rol = ROLES.indexOf(u.rol) >= 0 ? u.rol : 'registro';
  var sucursal = (rol === 'gerente' || rol === 'auxiliar' || rol === 'cmo') ? String(u.sucursal || '').trim().toLowerCase() : '';
  var confirma = rol === 'registro' && u.confirma ? 'Sí' : '';
  if (rol === 'gerente') {
    var su = unidadPorId(sucursal);
    if (!su || !su.vende) throw new Error('Escoja la sucursal del gerente.');
  }
  if (rol === 'auxiliar' || rol === 'cmo') {
    var lu = unidadPorId(sucursal);
    if (!lu || lu.activo === false) throw new Error('Escoja el lugar donde trabaja esta persona.');
  }
  var activo = u.activo === false ? 'No' : 'Sí';
  var original = String(u.original || '').trim();

  if (nombre.length < 2) throw new Error('El nombre debe tener al menos 2 letras.');
  if (activo === 'Sí') revisaCodigo(pin, original || nombre);

  var lock = LockService.getScriptLock();
  lock.waitLock(15000); _LEE = {};
  try {
    var h = hojaUsuarios();
    var todos = filasUsuarios();
    var repetido = todos.filter(function (x) {
      return x.nombre.toLowerCase() === nombre.toLowerCase() &&
             x.nombre.toLowerCase() !== original.toLowerCase();
    });
    if (repetido.length) throw new Error('Ya hay alguien con ese nombre.');

    var destino = null;
    if (original) {
      todos.forEach(function (x) {
        if (x.nombre.toLowerCase() === original.toLowerCase()) destino = x;
      });
      if (!destino) throw new Error('No se encontró a esa persona.');
    }

    // no dejar la app sin ningún administrador activo
    var admins = todos.filter(function (x) {
      return x.rol === 'admin' && x.activo &&
             (!destino || x.nombre.toLowerCase() !== destino.nombre.toLowerCase());
    }).length;
    if (admins === 0 && (rol !== 'admin' || activo === 'No')) {
      throw new Error('Tiene que quedar al menos un administrador activo.');
    }

    if (destino) {
      h.getRange(destino.fila, 1, 1, 4).setValues([[nombre, pin, rol, activo]]);
      h.getRange(destino.fila, 7, 1, 2).setValues([[sucursal, confirma]]);
    } else {
      h.appendRow([nombre, pin, rol, activo, new Date(), yo.nombre, sucursal, confirma]);
      h.getRange(h.getLastRow(), 2).setNumberFormat('@').setValue(pin);
    }
    olvidaUsuarios();
    return { ok: true, mensaje: destino ? 'Se actualizó a ' + nombre + '.'
      : 'Se agregó a ' + nombre + '.', usuarios: listaUsuarios(cred) };
  } finally {
    lock.releaseLock();
  }
}

/** Da de baja: no borra el renglón, lo marca inactivo, para no perder el historial. */
function darDeBaja(cred, nombre) {
  var yo = exigeAdmin(cred);
  nombre = String(nombre || '').trim();
  if (nombre.toLowerCase() === yo.nombre.toLowerCase())
    throw new Error('No puede darse de baja usted mismo.');

  var lock = LockService.getScriptLock();
  lock.waitLock(15000); _LEE = {};
  try {
    var h = hojaUsuarios();
    var destino = null, admins = 0;
    filasUsuarios().forEach(function (x) {
      if (x.nombre.toLowerCase() === nombre.toLowerCase()) destino = x;
      else if (x.rol === 'admin' && x.activo) admins++;
    });
    if (!destino) throw new Error('No se encontró a esa persona.');
    if (destino.rol === 'admin' && admins === 0)
      throw new Error('Tiene que quedar al menos un administrador activo.');
    h.getRange(destino.fila, 4).setValue('No');
    olvidaUsuarios();
    return { ok: true, mensaje: nombre + ' ya no puede entrar.',
      usuarios: listaUsuarios(cred) };
  } finally {
    lock.releaseLock();
  }
}

/** Cada quien puede cambiar su propio PIN sin ser administrador. */
function cambiarMiPin(cred, nuevo) {
  permiteAux();
  var yo = quien(cred);
  nuevo = String(nuevo || '').trim();
  revisaCodigo(nuevo, yo.nombre);
  var h = hojaUsuarios(), destino = null;
  filasUsuarios().forEach(function (x) {
    if (x.nombre.toLowerCase() === yo.nombre.toLowerCase()) destino = x;
  });
  if (!destino) throw new Error('No se encontró su usuario.');
  h.getRange(destino.fila, 2).setNumberFormat('@').setValue(nuevo);
  olvidaUsuarios();
  return { ok: true, mensaje: 'Su código quedó cambiado.' };
}

/* ════════════ RESUMEN DEL MES EN PDF ════════════ */

function carpetaReportes() {
  var nombre = 'Caja Rey Pizza — reportes';
  var it = DriveApp.getFoldersByName(nombre);
  return it.hasNext() ? it.next() : DriveApp.createFolder(nombre);
}

/** El estado de resultados del mes, en HTML, para el PDF. Una columna por
 *  sucursal —cada una es su propia unidad de negocio—, una con el consolidado
 *  de Rey Pizza, la bodega aparte porque no se reparte, y el total de CAIX. */
function bloqueER(gUC, por, sinDet, tg) {
  var cols = columnasER();
  var cl = clasifica();
  var h = [];

  function cortes(c) {
    var t = 0;
    c.u.forEach(function (u) { t += (por[u] ? por[u].ing : 0); });
    return t;
  }
  function envios(c) {
    var t = 0;
    c.u.forEach(function (u) { t += (por[u] ? por[u].env || 0 : 0); });
    return t;
  }
  function ventas(c) { return cortes(c) - envios(c); }
  function gasto(c, cat) {
    var t = 0;
    c.u.forEach(function (u) { t += (gUC[u] && gUC[u][cat]) || 0; });
    return t;
  }
  function gastoDe(c, cats) {
    var t = 0;
    cats.forEach(function (cat) { t += gasto(c, cat); });
    return t;
  }
  function hay(cats) {
    for (var i = 0; i < cols.length; i++)
      if (Math.abs(gastoDe(cols[i], cats)) > 0.004) return true;
    return false;
  }

  // categorías que llegaron y no están clasificadas (registros viejos)
  var conocidas = {};
  cl.orden.forEach(function (g) {
    g.categorias.forEach(function (c) { conocidas[c] = true; });
  });
  var sueltas = [];
  Object.keys(gUC).forEach(function (u) {
    Object.keys(gUC[u]).forEach(function (c) {
      if (!conocidas[c] && c !== CAT_ENVIOS && sueltas.indexOf(c) < 0) sueltas.push(c);
    });
  });

  var costo = [], oper = [];
  cl.orden.forEach(function (g) {
    (g.grupo === G_COSTO ? costo : oper).push(g);
  });
  if (sueltas.length)
    oper.push({ grupo: G_OPER, renglon: 'Sin clasificar', categorias: sueltas });

  var catsCosto = [], catsOper = [];
  costo.forEach(function (g) { catsCosto = catsCosto.concat(g.categorias); });
  oper.forEach(function (g) { catsOper = catsOper.concat(g.categorias); });

  function fila(txt, vals, clase, sangra) {
    var t = '<tr' + (clase ? ' class="' + clase + '"' : '') + '><td' +
      (sangra ? ' class="i"' : '') + '>' + htm(txt) + '</td>';
    vals.forEach(function (v, i) {
      t += '<td class="n' + (cols[i].fuerte ? ' f' : '') + '">' + v + '</td>';
    });
    return t + '</tr>';
  }
  function money(v, ceroRaya) {
    if (ceroRaya && Math.abs(v) < 0.004) return '&mdash;';
    return (v < 0 ? '<span class="neg">' + dinero(v) + '</span>' : dinero(v));
  }
  function cifras(fn, ceroRaya) {
    return cols.map(function (c) { return money(fn(c), ceroRaya); });
  }

  h.push('<h2>Estado de resultados</h2>');
  h.push('<table class="er"><tr><th>Concepto</th>' +
    cols.map(function (c) {
      return '<th class="n' + (c.fuerte ? ' f' : '') + '">' + htm(c.t) + '</th>';
    }).join('') + '</tr>');

  h.push(fila('VENTAS', cols.map(function () { return ''; }), 'sec'));
  var hayEnv = cols.some(function (c) { return Math.abs(envios(c)) > 0.004; });
  if (hayEnv) {
    h.push(fila('Ventas (cortes)', cifras(cortes), '', true));
    h.push(fila('(−) Envíos entregados a repartidores', cifras(function (c) { return -envios(c); }, true), '', true));
  }
  h.push(fila('Ventas netas', cifras(ventas), 'sub'));

  h.push(fila('COSTO DE VENTAS', cols.map(function () { return ''; }), 'sec'));
  catsCosto.forEach(function (cat) {
    if (!hay([cat])) return;
    h.push(fila(cat, cifras(function (c) { return gasto(c, cat); }, true),
      '', true));
  });
  h.push(fila('Total costo de ventas',
    cifras(function (c) { return gastoDe(c, catsCosto); }), 'sub'));

  h.push(fila('UTILIDAD BRUTA', cifras(function (c) {
    return ventas(c) - gastoDe(c, catsCosto); }), 'res'));
  h.push(fila('Margen bruto', cols.map(function (c) {
    var v = ventas(c);
    if (!v) return '&mdash;';
    return ((ventas(c) - gastoDe(c, catsCosto)) / v * 100).toFixed(1) + '%';
  }), 'pct', true));

  h.push(fila('GASTOS DE OPERACIÓN', cols.map(function () { return ''; }), 'sec'));
  oper.forEach(function (g) {
    var vivas = g.categorias.filter(function (cat) { return hay([cat]); });
    if (!vivas.length) return;
    // con una sola categoría el subtotal repetiría la misma cifra
    if (vivas.length === 1) {
      h.push(fila(g.renglon + ' — ' + vivas[0],
        cifras(function (c) { return gasto(c, vivas[0]); }), 'sub'));
      return;
    }
    vivas.forEach(function (cat) {
      h.push(fila(cat, cifras(function (c) { return gasto(c, cat); }, true),
        '', true));
    });
    h.push(fila(g.renglon,
      cifras(function (c) { return gastoDe(c, g.categorias); }), 'sub'));
  });
  h.push(fila('Total gastos de operación',
    cifras(function (c) { return gastoDe(c, catsOper); }), 'sub'));

  h.push(fila('UTILIDAD DE OPERACIÓN', cifras(function (c) {
    return ventas(c) - gastoDe(c, catsCosto) - gastoDe(c, catsOper); }), 'res'));
  h.push(fila('Margen de operación', cols.map(function (c) {
    var v = ventas(c);
    if (!v) return '&mdash;';
    return ((v - gastoDe(c, catsCosto) - gastoDe(c, catsOper)) / v * 100)
      .toFixed(1) + '%';
  }), 'pct', true));

  h.push('</table>');
  h.push('<p class="pie2">Cada sucursal es una unidad de negocio aparte. ' +
    'El producto que despacha la bodega se carga a la sucursal que lo recibe, al ' +
    'precio de compra y sin recargo; en la columna de la bodega aparece restado ' +
    '(«Traslado a sucursales»), para no contarlo dos veces. ' +
    '<b>Los gastos propios de la bodega no se reparten</b>: van en su columna y ' +
    'solo entran en el total de CAIX, S.A. ' +
    (sinDet > 0
      ? 'De los ' + dinero(tg) + ' de gasto del mes, ' + dinero(sinDet) +
        ' viene de facturas anotadas solo con el total, sin desglose por artículo.'
      : 'Todas las facturas del mes traen su desglose por artículo.') +
    ' Los grupos se cambian en la hoja «Clasificación».</p>');
  return h.join('');
}

function resumenMes(cred, mes) {
  var yo = verificaCaja(cred);
  mes = String(mes || '').trim();
  if (!/^\d{4}-\d{2}$/.test(mes)) throw new Error('Mes inválido.');

  var movs = movimientos(4000).filter(function (m) {
    return m.mes === mes && !m.anulado;
  });

  var por = {};
  UNIDADES.forEach(function (u) {
    por[u.nombre] = { ing: 0, gas: 0, efe: 0, tar: 0, caja: 0, sinDesglose: 0, com: 0, env: 0 };
  });
  movs.forEach(function (m) {
    if (!por[m.unidad])
      por[m.unidad] = { ing: 0, gas: 0, efe: 0, tar: 0, caja: 0, sinDesglose: 0, com: 0, env: 0 };
    var p = por[m.unidad];
    if (m.clase !== 'ingreso' && m.titulo === CAT_ENVIOS) { p.env += m.monto; return; }   // no es gasto: se resta de la venta
    if (m.clase !== 'ingreso') { p.gas += m.monto; return; }
    p.ing += m.monto;
    if (m.desglosado) { p.efe += m.efectivo || 0; p.tar += m.tarjeta || 0; }
    else p.sinDesglose += m.monto;
    // lo que salió de la caja es gasto, aunque se capture en el corte
    if (m.caja > 0) { p.caja += m.caja; p.gas += m.caja; }
    if (m.comision > 0) { p.com += m.comision; p.gas += m.comision; }
  });

  var MES_N = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio',
    'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre'];
  var titulo = MES_N[parseInt(mes.slice(5), 10) - 1] + ' de ' + mes.slice(0, 4);

  var h = [];
  h.push('<html><head><meta charset="utf-8"><style>');
  h.push('body{font-family:Arial,Helvetica,sans-serif;font-size:11px;color:#16181d;' +
    'margin:28px}');
  h.push('table.band{width:100%;background:#14284B;color:#fff;border-bottom:4px solid #F2B705;margin:0 0 14px;border-collapse:collapse}');
  h.push('table.band td{border:0;padding:12px 14px;vertical-align:middle}');
  h.push('.band img{width:42px}.band td.r{text-align:right;font-size:13px;color:#F7F2E7;font-weight:bold}');
  h.push('h1{font-size:21px;color:#fff;margin:0;letter-spacing:1px}.band .co{color:#F2B705;font-size:10px;letter-spacing:2px;font-weight:bold}');
  h.push('h2{font-size:13px;color:#14284B;margin:18px 0 6px}');
  h.push('table{width:100%;border-collapse:collapse;font-size:10.5px}');
  h.push('th{background:#14284B;color:#fff;text-align:left;padding:6px 8px}');
  h.push('td{padding:5px 8px;border-bottom:1px solid #ececea}');
  h.push('td.n,th.n{text-align:right}');
  h.push('tr.t td{background:#F7F2E7;font-weight:bold;border-top:2px solid #14284B}');
  /* el renglón de la factura y, debajo, sus artículos */
  h.push('tr.cab td{border-bottom:0;border-top:1px solid #d9d6d1;padding-top:7px}');
  h.push('tr.cab b{font-size:11px}');
  h.push('span.d{display:block;color:#6c7079;font-size:9px;margin-top:1px}');
  h.push('tr.art td{border-bottom:0;padding:1px 8px;font-size:9.5px;color:#3d4046}');
  h.push('tr.art td.a{padding-left:24px}');
  h.push('tr.art td.a i{font-style:normal;color:#6c7079}');
  h.push('tr.fin td{padding-bottom:7px}');
  h.push('tr.sin td{border-bottom:0;border-top:1px solid #d9d6d1;padding:7px 8px}');
  /* estado de resultados */
  h.push('table.er{font-size:9.5px}');
  h.push('table.er th{padding:5px 5px;font-size:9px}');
  h.push('table.er td{padding:3px 5px;border-bottom:0}');
  h.push('table.er td.f{background:#FAF7F1}');
  h.push('table.er th.f{background:#1D3A6B}');
  h.push('table.er tr.sub td{font-weight:bold;border-top:1px solid #cfcdc8}');
  h.push('table.er tr.sec td{background:#F7F2E7;color:#14284B;font-weight:bold;' +
    'text-transform:uppercase;font-size:8.5px;letter-spacing:.05em;padding:4px 5px}');
  h.push('table.er tr.res td{font-weight:bold;background:#FFF4D1;' +
    'border-top:1.4px solid #14284B;border-bottom:1.4px solid #14284B}');
  h.push('table.er tr.pct td{color:#6c7079;font-size:9px;padding-bottom:5px}');
  h.push('table.er td.i{padding-left:16px}');
  h.push('.neg{color:#b3261e}');
  h.push('.pie2{font-size:8.5px;color:#6c7079;margin:5px 0 0;line-height:1.45}');
  h.push('.pie{margin-top:18px;border-top:1px solid #e3dfd8;padding-top:7px;' +
    'font-size:9px;color:#6c7079}');
  h.push('</style></head><body>');
  h.push('<table class="band"><tr><td style="width:54px"><img src="' + LOGO_PNG + '"></td><td><h1>REY PIZZA</h1><div class="co">CAIX, S.A.</div></td>' +
    '<td class="r">Resumen de ' + titulo + '</td></tr></table>');

  var ti = 0, tg = 0;
  UNIDADES.forEach(function (u) {
    ti += por[u.nombre].ing; tg += por[u.nombre].gas;
  });

  /* ── el desglose por artículo ── */
  var det = detalleDeGastos(movs);
  var porGasto = {};
  det.forEach(function (x) {
    if (!porGasto[x.id]) porGasto[x.id] = [];
    porGasto[x.id].push(x);
  });

  var orden = function (o, campo) {
    return Object.keys(o).sort(function (a, b) {
      var va = campo ? o[a][campo] : o[a], vb = campo ? o[b][campo] : o[b];
      return vb - va;
    });
  };

  /* ── gasto por unidad y categoría: el renglón del artículo cuando la factura
   *    trae desglose, el total de la factura cuando no lo trae. Sin contar dos
   *    veces, así que esto suma exactamente el gasto del mes. ── */
  var gUC = {}, sinDet = 0;
  function suma(unidad, cat, monto) {
    if (!gUC[unidad]) gUC[unidad] = {};
    cat = cat || '(sin categoría)';
    gUC[unidad][cat] = (gUC[unidad][cat] || 0) + monto;
  }
  det.forEach(function (x) { suma(x.unidad, x.categoria, x.subtotal); });
  movs.forEach(function (m) {
    if (m.clase === 'ingreso') {
      if (m.caja > 0) suma(m.unidad, CAT_CAJA, m.caja);
      if (m.comision > 0) suma(m.unidad, CAT_COMISION, m.comision);
      return;
    }
    if (m.titulo === CAT_ENVIOS) return;              // se resta de la venta, no es gasto
    if (porGasto[m.id]) return;
    suma(m.unidad, m.titulo, m.monto);
    sinDet += m.monto;
  });

  h.push(bloqueER(gUC, por, sinDet, tg));

  /* ── cómo cobró cada sucursal y qué salió de la caja ── */
  var hayCorte = false;
  UNIDADES.forEach(function (u) {
    if (u.vende && (por[u.nombre].ing > 0 || por[u.nombre].caja > 0)) hayCorte = true;
  });
  if (hayCorte) {
    h.push('<h2>Cobro y caja</h2><table><tr><th>Sucursal</th>' +
      '<th class="n">Efectivo</th><th class="n">Tarjeta</th>' +
      '<th class="n">Venta</th>' +
      '<th class="n">Gastos de caja</th><th class="n">Efectivo a depositar</th>' +
      '<th class="n">Comisión NEONET</th><th class="n">Tarjeta líquida</th></tr>');
    var te = 0, tt = 0, ts = 0, tv = 0, tc = 0, tk = 0;
    UNIDADES.forEach(function (u) {
      if (!u.vende) return;
      var p = por[u.nombre];
      te += p.efe; tt += p.tar; ts += p.sinDesglose; tv += p.ing; tc += p.caja; tk += p.com;
      h.push('<tr><td>' + htm(u.nombre) + '</td><td class="n">' + dinero(p.efe) +
        '</td><td class="n">' + dinero(p.tar) + '</td><td class="n">' +
        dinero(p.ing) + '</td><td class="n">' +
        (p.caja > 0 ? dinero(p.caja) : '&mdash;') + '</td><td class="n">' +
        dinero(p.efe - p.caja) + '</td><td class="n">' + (p.com > 0 ? dinero(p.com) : '&mdash;') +
        '</td><td class="n">' + dinero(p.tar - p.com) + '</td></tr>');
    });
    h.push('<tr class="t"><td>TOTAL</td><td class="n">' + dinero(te) +
      '</td><td class="n">' + dinero(tt) + '</td><td class="n">' + dinero(tv) +
      '</td><td class="n">' + dinero(tc) + '</td><td class="n">' +
      dinero(te - tc) + '</td><td class="n">' + dinero(tk) + '</td><td class="n">' +
      dinero(tt - tk) + '</td></tr></table>');
    h.push('<p class="pie2">«Efectivo a depositar» es el efectivo del turno menos ' +
      'lo que salió de la caja. Los cortes anotados antes de que la app separara ' +
      'efectivo y tarjeta cuentan como efectivo. «Tarjeta líquida» es lo que deposita NEONET en ' +
      'el banco después de su comisión (' + num(tasaComision()) + ' %).</p>');
  }

  /* ── lo que salió de bodega a cada sucursal (cantidades y costo) ── */
  var bodegaN = nombreBodega();
  var tras = det.filter(function (x) {
    return x.proveedor === PROV_BODEGA && x.unidad !== bodegaN; });
  if (tras.length) {
    var sucs = sucursalesDestino(), tp = {}, tc = {}, tt = 0;
    tras.forEach(function (x) {
      var k = x.producto + '§' + x.medida;
      if (!tp[k]) tp[k] = { p: x.producto, m: x.medida, c: {}, n: 0, q: 0 };
      tp[k].c[x.unidad] = (tp[k].c[x.unidad] || 0) + x.cantidad;
      tp[k].n += x.cantidad; tp[k].q += x.subtotal;
      tc[x.unidad] = (tc[x.unidad] || 0) + x.subtotal; tt += x.subtotal;
    });
    h.push('<h2>Traslados de bodega a sucursales</h2><table><tr><th>Producto</th>' +
      sucs.map(function (s) { return '<th class="n">' + htm(s.nombre) + '</th>'; }).join('') +
      '<th class="n">Total</th><th class="n">Costo</th></tr>');
    orden(tp, 'q').forEach(function (k) {
      var x = tp[k];
      h.push('<tr><td>' + htm(x.p) + '</td>' + sucs.map(function (s) {
        var n = x.c[s.nombre] || 0;
        return '<td class="n">' + (n ? num(n) : '&mdash;') + '</td>'; }).join('') +
        '<td class="n">' + num(x.n) + ' ' + htm(plur(x.n, x.m)) + '</td>' +
        '<td class="n">' + dinero(x.q) + '</td></tr>');
    });
    h.push('<tr class="t"><td>Costo cargado a cada sucursal</td>' + sucs.map(function (s) {
      return '<td class="n">' + dinero(tc[s.nombre] || 0) + '</td>'; }).join('') +
      '<td></td><td class="n">' + dinero(tt) + '</td></tr></table>');
    h.push('<p class="pie2">Al precio de compra, sin recargo. Ya está incluido en el ' +
      'costo de ventas de cada sucursal y restado en la columna de la bodega.</p>');
  }

  // lo trasladado no es una compra: no va en «Compras por proveedor»
  var compra = det.filter(function (x) { return x.proveedor !== PROV_BODEGA; });
  if (compra.length) {
    var porProv = {}, porProd = {};
    compra.forEach(function (x) {
      var pv = x.proveedor || '(sin proveedor)';
      if (!porProv[pv]) porProv[pv] = { q: 0, n: 0 };
      porProv[pv].q += x.subtotal; porProv[pv].n++;
      var pd = x.producto;
      if (!porProd[pd]) porProd[pd] = { q: 0, c: 0, m: x.medida, b: 0, bu: x.base };
      porProd[pd].q += x.subtotal; porProd[pd].c += x.cantidad;
      if (x.cantidadBase > 0) { porProd[pd].b += x.cantidadBase;
        porProd[pd].bu = x.base; }
    });

    h.push('<h2>Compras por proveedor</h2><table><tr><th>Proveedor</th>' +
      '<th class="n">Artículos</th><th class="n">Total</th></tr>');
    orden(porProv, 'q').forEach(function (k) {
      h.push('<tr><td>' + htm(k) + '</td><td class="n">' + porProv[k].n +
        '</td><td class="n">' + dinero(porProv[k].q) + '</td></tr>');
    });
    h.push('</table>');

    var claves = orden(porProd, 'q');
    h.push('<h2>Productos comprados (' + claves.length + ')</h2>' +
      '<table><tr><th>Producto</th><th class="n">Comprado</th>' +
      '<th class="n">Contenido</th><th class="n">Costo unitario</th>' +
      '<th class="n">Total</th></tr>');
    claves.forEach(function (k) {
      var x = porProd[k];
      var cu = x.b > 0 ? dinero(x.q / x.b) + ' / ' + htm(x.bu)
                       : (x.c > 0 ? dinero(x.q / x.c) + ' / ' + htm(x.m || '') : '—');
      h.push('<tr><td>' + htm(k) + '</td><td class="n">' +
        num(x.c) + ' ' + htm(plur(x.c, x.m)) +
        '</td><td class="n">' + (x.b > 0
          ? num(x.b) + ' ' + htm(plur(x.b, x.bu)) : '—') +
        '</td><td class="n">' + cu +
        '</td><td class="n">' + dinero(x.q) + '</td></tr>');
    });
    h.push('</table>');
  }

  /* cada factura con sus artículos debajo: es lo que se traslada al estado de
   * resultados, así que tiene que venir renglón por renglón */
  h.push('<h2>Detalle de movimientos (' + movs.length + ')</h2>');
  if (!movs.length) {
    h.push('<p>No hay movimientos registrados en este mes.</p>');
  } else {
    h.push('<table><tr><th>Fecha</th><th>Unidad</th><th>Concepto</th>' +
      '<th>Registró</th><th class="n">Monto</th></tr>');
    movs.forEach(function (m) {
      var arts = (m.clase === 'gasto' && porGasto[m.id]) ? porGasto[m.id] : [];

      var tit = (m.clase === 'gasto')
        ? (m.proveedor || m.titulo)
        : '+ ' + m.titulo;
      var bajo = [];
      if (m.clase === 'gasto' && m.proveedor) bajo.push(htm(m.titulo));
      if (m.factura) bajo.push('factura ' + htm(m.factura));
      var cuantos = m.articulos || arts.length;
      if (cuantos) bajo.push(cuantos +
        (cuantos === 1 ? ' artículo' : ' artículos'));
      if (m.articulos > 0 && !arts.length)
        bajo.push('<b>no se encontró el desglose: corra la función instalar</b>');
      if (m.nota) bajo.push(htm(m.nota));

      var conDetalle = arts.length ||
        (m.clase === 'ingreso' && (m.desglosado || m.caja > 0));
      h.push('<tr class="' + (conDetalle ? 'cab' : 'sin') + '"><td>' +
        m.fecha.split('-').reverse().join('/') + '</td><td>' + htm(m.unidad) +
        '</td><td><b>' + htm(tit) + '</b>' +
        (bajo.length ? '<span class="d">' + bajo.join(' &middot; ') + '</span>' : '') +
        '</td><td>' + htm(m.quien || '') + '</td><td class="n"><b>' +
        (m.clase === 'gasto' ? '-' : '') + dinero(m.monto) + '</b></td></tr>');

      if (m.clase === 'ingreso' && (m.desglosado || m.caja > 0)) {
        var cortes = [];
        if (m.desglosado) {
          cortes.push(['Efectivo', m.efectivo]);
          cortes.push(['Tarjeta', m.tarjeta]);
        }
        if (m.caja > 0)
          cortes.push(['Gastos de caja' +
            (m.detalleCaja ? ' — ' + htm(m.detalleCaja) : ''), -m.caja]);
        cortes.forEach(function (c, k) {
          h.push('<tr class="art' + (k === cortes.length - 1 ? ' fin' : '') +
            '"><td></td><td></td><td class="a">' + c[0] +
            '</td><td></td><td class="n">' + dinero(c[1]) + '</td></tr>');
        });
      }

      arts.forEach(function (a, k) {
        var pie = num(a.cantidad) + ' ' + htm(plur(a.cantidad, a.medida)) +
          ' &times; ' + dinero(a.precio);
        if (a.cantidadBase > 0) pie += ' &middot; ' + num(a.cantidadBase) + ' ' +
          htm(plur(a.cantidadBase, a.base)) + ' a ' + dinero(a.costoBase);
        if (a.categoria) pie += ' &middot; ' + htm(a.categoria);
        h.push('<tr class="art' + (k === arts.length - 1 ? ' fin' : '') +
          '"><td></td><td></td><td class="a">' + htm(a.producto) +
          ' <i>' + pie + '</i></td><td></td><td class="n">' +
          dinero(a.subtotal) + '</td></tr>');
      });
    });
    h.push('</table>');
  }

  h.push('<div class="pie">Generado el ' +
    Utilities.formatDate(new Date(), ZONA, 'dd/MM/yyyy HH:mm') +
    ' desde la app Caja Rey Pizza. Los registros anulados no aparecen ni suman. ' +
    'Este archivo se puede abrir con el enlace: guárdelo donde corresponda.</div>');
  h.push('</body></html>');

  var nombreArchivo = 'Caja Rey Pizza — ' + mes + '.pdf';
  var blob = Utilities.newBlob(h.join(''), 'text/html', 'r.html')
    .getAs('application/pdf').setName(nombreArchivo);

  var carpeta = carpetaReportes();
  // si ya existe uno de ese mes, se manda a la papelera para no acumular
  var viejos = carpeta.getFilesByName(nombreArchivo);
  while (viejos.hasNext()) viejos.next().setTrashed(true);

  var f = carpeta.createFile(blob);
  f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  var apunte = apuntaReporte(mes, yo.nombre, movs.length, nombreArchivo,
    f.getUrl(), f.getId());
  return { ok: true, url: f.getUrl(), nombre: nombreArchivo,
    movimientos: movs.length, reporte: apunte };
}

/* ════════════ HISTORIAL DE REPORTES ════════════ */

/** Deja constancia de cada vez que se arma un reporte. Un renglón por mes: si se
 *  vuelve a generar, se actualiza en vez de duplicarse. */
function apuntaReporte(mes, quien, cuantos, archivo, url, id) {
  var ahora = new Date();
  var fila = { mes: mes, cuando: fmtSello(ahora), quien: quien,
    movimientos: cuantos, archivo: archivo, url: url, veces: 1 };
  try {
    var h = hojaCat('Reportes', H_REPORTES);
    var v = leeTodo(h, H_REPORTES.length);
    var destino = 0;
    for (var k = 0; k < v.length; k++) {
      if (String(v[k][0]).trim() === mes) { destino = k + 2;
        fila.veces = (Number(v[k][7]) || 1) + 1; break; }
    }
    var datos = [mes, ahora, quien, cuantos, archivo, url, id, fila.veces];
    if (destino) {
      h.getRange(destino, 1, 1, H_REPORTES.length).setValues([datos]);
    } else {
      h.appendRow(datos);
      h.getRange(h.getLastRow(), 1).setNumberFormat('@').setValue(mes);
    }
  } catch (e) { /* el PDF ya quedó hecho; el apunte es lo secundario */ }
  return fila;
}

function listaReportes(cred) {
  verificaCaja(cred);
  var out = [];
  try {
    leeTodo(hojaCat('Reportes', H_REPORTES), H_REPORTES.length)
      .forEach(function (r) {
        if (!String(r[0]).trim()) return;
        out.push({ mes: String(r[0]).trim(), cuando: fmtSello(r[1]),
          quien: String(r[2] || ''), movimientos: Number(r[3]) || 0,
          archivo: String(r[4] || ''), url: String(r[5] || ''),
          veces: Number(r[7]) || 1 });
      });
  } catch (e) { return []; }
  out.sort(function (a, b) { return a.mes < b.mes ? 1 : -1; });
  return out.slice(0, 60);
}

/* ════════════ LO QUE ENVÍAN LOS GERENTES: POR CONFIRMAR ════════════
 * El gerente registra sus cortes y sus gastos, pero no entran a las hojas oficiales
 * (Ingresos, Gastos) hasta que el administrador, o alguien con permiso de confirmar,
 * los revisa. Mientras tanto viven aquí. Al confirmar, se escriben en la hoja
 * oficial con la hora en que el gerente los envió, y quedan grabados los dos
 * nombres: quién envió y quién confirmó. */
var H_PEND = ['ID', 'Tipo', 'Sucursal', 'Enviado por', 'Enviado en', 'Fecha', 'Turno',
  'Monto', 'Resumen', 'Datos', 'Estado', 'Revisado por', 'Revisado en',
  'Motivo de rechazo', 'ID oficial', 'Validado por', 'Validado en', 'Efectivo recibido',
  'Diferencia de efectivo', 'Recibido por', 'Recibido en', 'Tarjeta NEONET', 'Comisión tarjeta',
  'Tarjeta líquida'];
/* El cierre de caja pasa por cuatro manos:
 *   el gerente lo reporta («Por confirmar», se muestra «por validar»)
 *   → el director operativo lo valida («Validado»)
 *   → el contador interno recibe el efectivo («Efectivo recibido»)
 *   → el director financiero anota la tarjeta del cierre NEONET y confirma que llegó
 *     al banco («Confirmado»): ahí entra a Ingresos, con la comisión calculada.
 * Cualquiera de los tres puede rechazarlo con motivo y regresa al gerente. */
var EST_PEND = 'Por confirmar', EST_VAL = 'Validado', EST_REC = 'Efectivo recibido',
  EST_CONF = 'Confirmado', EST_RECH = 'Rechazado';
function enCamino(x) { return x.estado === EST_PEND || x.estado === EST_VAL || x.estado === EST_REC; }

function hojaPorConfirmar(ss) {
  ss = ss || libro();
  var h = ss.getSheetByName('Por confirmar');
  if (h) return h;
  h = hojaLimpia(ss, 'Por confirmar', H_PEND);
  h.getRange('E:E').setNumberFormat('dd/mm/yyyy hh:mm');
  h.getRange('F:F').setNumberFormat('@');
  h.getRange('H:H').setNumberFormat('"Q"#,##0.00');
  h.getRange('M:M').setNumberFormat('dd/mm/yyyy hh:mm');
  h.setColumnWidth(9, 320); h.setColumnWidth(10, 120);
  return h;
}

function filasPorConfirmar() {
  return leeTodo(hojaPorConfirmar(), H_PEND.length).map(function (r, i) {
    return { fila: i + 2, id: String(r[0] || ''), tipo: String(r[1] || ''),
      sucursal: String(r[2] || ''), enviadoPor: String(r[3] || ''),
      enviadoEn: r[4], sello: fmtSello(r[4]), fecha: fmtDia(r[5]),
      turno: String(r[6] || ''), monto: Number(r[7]) || 0, resumen: String(r[8] || ''),
      datos: String(r[9] || ''), estado: String(r[10] || '') || EST_PEND,
      revisadoPor: String(r[11] || ''), revisadoEn: fmtSello(r[12]),
      motivo: String(r[13] || ''), oficial: String(r[14] || ''),
      validadoPor: String(r[15] || ''), validadoEn: fmtSello(r[16]),
      recibido: r[17] === '' || r[17] == null ? null : Number(r[17]),
      diferencia: Number(r[18]) || 0,
      recibidoPor: String(r[19] || ''), recibidoEn: fmtSello(r[20]),
      tarjetaNeonet: r[21] === '' || r[21] == null ? null : Number(r[21]),
      comision: Number(r[22]) || 0, liquida: Number(r[23]) || 0 };
  }).filter(function (x) { return x.id; });
}

function resumenDe(tipo, v) {
  if (tipo === 'ingreso')
    return 'Corte ' + v.turno + ' · efectivo ' + dinero(v.efectivo) + ' · tarjeta ' +
      dinero(v.tarjeta) + (v.caja > 0 ? ' · gastos de caja ' + dinero(v.caja) +
      ' (' + v.detalleCaja + ')' : '') + (v.nota ? ' · ' + v.nota : '');
  return [v.proveedor || '', v.factura ? 'fact. ' + v.factura : '', v.categoria,
    v.lineas.length ? v.lineas.length + (v.lineas.length === 1 ? ' artículo'
      : ' artículos') : '', v.nota].filter(function (x) { return x; }).join(' · ');
}

function guardaPorConfirmar(yo, tipo, v, d) {
  var u = unidadPorId(v.unidad);
  var lock = LockService.getScriptLock();
  lock.waitLock(15000); _LEE = {};
  try {
    if (tipo === 'ingreso') {
      var rep = buscaIngreso(libro().getSheetByName('Ingresos'), u.nombre, v.fecha, v.turno);
      if (rep) throw new Error(repetidoTxt(v, u, rep));
    }
    var ahora = new Date();
    var id = 'P' + Utilities.formatDate(ahora, ZONA, 'yyyyMMddHHmmss') +
      Math.floor(Math.random() * 900 + 100);
    var fecha = tipo === 'ingreso' ? v.fecha : v.fechaPago;
    var h = hojaPorConfirmar();
    h.appendRow([id, tipo, u.nombre, yo.nombre, ahora, fecha,
      tipo === 'ingreso' ? v.turno : '', v.monto, resumenDe(tipo, v), JSON.stringify(v),
      EST_PEND, '', '', '', '', '', '', '', '', '', '', '', '', '']);
    h.getRange(h.getLastRow(), 6).setNumberFormat('@').setValue(fecha);
    return { ok: true, porConfirmar: true,
      mensaje: (tipo === 'ingreso' ? 'Corte ' + v.turno : 'Gasto de ' + dinero(v.monto)) +
        ' enviado. Queda por validar.',
      enviado: enviadoParaGerente({ id: id, tipo: tipo, sucursal: u.nombre,
        sello: fmtSello(ahora), fecha: fecha, turno: tipo === 'ingreso' ? v.turno : '',
        monto: v.monto, resumen: resumenDe(tipo, v), estado: EST_PEND,
        enviadoPor: yo.nombre }) };
  } finally {
    lock.releaseLock();
  }
}

/** Lo que el gerente ve de lo que envió: qué, cuándo y en qué quedó. Sin totales. */
function enviadoParaGerente(x) {
  return { id: x.id, tipo: x.tipo, sello: x.sello, fecha: x.fecha, turno: x.turno,
    monto: x.monto, resumen: x.resumen, estado: x.estado, quien: x.enviadoPor,
    revisadoPor: x.revisadoPor || '', motivo: x.motivo || '',
    validadoPor: x.validadoPor || '' };
}

function misEnviados(yo) {
  var suc = nombreUnidad(yo.sucursal);
  var desde = Utilities.formatDate(new Date(Date.now() - 30 * 86400000), ZONA, 'yyyy-MM-dd');
  return filasPorConfirmar().filter(function (x) {
    return x.sucursal === suc && x.sello.slice(0, 10) >= desde && x.estado !== EST_CONF;
  }).map(enviadoParaGerente).reverse();
}

function exigeConfirmar(cred) {
  var yo = verificaCaja(cred);
  if (!yo.confirma) throw new Error('Su usuario no puede confirmar registros.');
  return yo;
}

/** Para quien confirma: lo que le toca a él (validar o recibir), lo que está en
 *  manos del otro paso y lo último revisado. */
function porConfirmar(cred) {
  var yo = exigeConfirmar(cred);
  return listaPorConfirmar(yo);
}
function pasoDe(x) {
  return x.estado === EST_PEND ? 'validar' : x.estado === EST_VAL ? 'recibir'
    : x.estado === EST_REC ? 'banco' : '';
}
function puedeEnPaso(yo, x) {
  var p = pasoDe(x);
  return (p === 'validar' && !!yo.valida) || (p === 'recibir' && !!yo.recibe) ||
    (p === 'banco' && !!yo.banco);
}
var QUIEN_PASO = { validar: 'el director operativo', recibir: 'el contador interno',
  banco: 'el contador interno' };
function listaPorConfirmar(yo) {
  yo = yo || { valida: true, recibe: true, banco: true };
  var todos = filasPorConfirmar();
  var desde = Utilities.formatDate(new Date(Date.now() - 7 * 86400000), ZONA, 'yyyy-MM-dd');
  function detalle(x) {
    var v = {};
    try { v = JSON.parse(x.datos); } catch (e) {}
    return { id: x.id, tipo: x.tipo, sucursal: x.sucursal, quien: x.enviadoPor,
      sello: x.sello, fecha: x.fecha, turno: x.turno, monto: x.monto, resumen: x.resumen,
      paso: pasoDe(x), validadoPor: x.validadoPor, validadoEn: x.validadoEn,
      recibidoPor: x.recibidoPor, recibidoEn: x.recibidoEn, recibido: x.recibido,
      diferencia: x.diferencia,
      efectivo: v.efectivo || 0, tarjeta: v.tarjeta || 0, caja: v.caja || 0,
      detalleCaja: v.detalleCaja || '', proveedor: v.proveedor || '',
      factura: v.factura || '', categoria: v.categoria || '', nota: v.nota || '',
      lineas: (v.lineas || []).map(function (l) {
        return { producto: l.producto, cantidad: l.cantidad, medida: l.medida,
          precio: l.precio, subtotal: l.subtotal }; }) };
  }
  var vivos = todos.filter(enCamino);
  var pend = vivos.filter(function (x) { return puedeEnPaso(yo, x); }).map(detalle);
  var otros = vivos.filter(function (x) { return !puedeEnPaso(yo, x); }).map(detalle);
  var hechos = todos.filter(function (x) {
    return !enCamino(x) && x.revisadoEn.slice(0, 10) >= desde;
  }).map(function (x) {
    return { id: x.id, tipo: x.tipo, sucursal: x.sucursal, quien: x.enviadoPor,
      sello: x.sello, fecha: x.fecha, turno: x.turno, monto: x.monto, resumen: x.resumen,
      estado: x.estado, revisadoPor: x.revisadoPor, revisadoEn: x.revisadoEn,
      validadoPor: x.validadoPor, recibidoPor: x.recibidoPor, motivo: x.motivo,
      recibido: x.recibido, diferencia: x.diferencia, tarjetaNeonet: x.tarjetaNeonet,
      comision: x.comision };
  }).reverse().slice(0, 30);
  return { pendientes: pend, enProceso: otros, revisados: hechos,
    valida: !!yo.valida, recibe: !!yo.recibe, banco: !!yo.banco, tasa: tasaComision() };
}

function buscaPend(id) {
  var x = null;
  filasPorConfirmar().forEach(function (r) { if (r.id === String(id)) x = r; });
  if (!x) throw new Error('No se encontró ese registro.');
  return x;
}
function yaRevisado(x) {
  if (x.estado === EST_CONF) return 'Ese registro ya fue confirmado por ' + x.revisadoPor + '.';
  if (x.estado === EST_RECH) return 'Ese registro ya fue rechazado por ' + x.revisadoPor + '.';
  return '';
}
/** ¿Le toca a esta persona? Si no, el mensaje dice a quién le toca. */
function exigePaso(yo, x, paso) {
  if (yaRevisado(x)) throw new Error(yaRevisado(x));
  var p = pasoDe(x);
  if (p !== paso) {
    var orden = ['validar', 'recibir', 'banco'];
    if (orden.indexOf(p) > orden.indexOf(paso))
      throw new Error(paso === 'validar' ? 'Ya lo validó ' + x.validadoPor + '.'
        : 'Ya recibió el efectivo ' + x.recibidoPor + '.');
    throw new Error('Primero lo ' + (p === 'validar' ? 'valida ' : 'recibe ') + QUIEN_PASO[p] + '.');
  }
}

/** Paso 2: el director operativo (o un administrador) revisa que el cierre esté bien. */
function validarRegistro(cred, id) {
  var yo = exigeConfirmar(cred);
  if (!yo.valida) throw new Error('El cierre lo valida el director operativo.');
  var lock = LockService.getScriptLock();
  lock.waitLock(15000); _LEE = {};
  var x;
  try {
    x = buscaPend(id);
    exigePaso(yo, x, 'validar');
    hojaPorConfirmar().getRange(x.fila, 11).setValue(EST_VAL);
    hojaPorConfirmar().getRange(x.fila, 16, 1, 2).setValues([[yo.nombre, new Date()]]);
  } finally { lock.releaseLock(); }
  var l = listaPorConfirmar(yo);
  l.ok = true;
  l.mensaje = 'Validado: ' + x.sucursal + ' · ' + (x.tipo === 'ingreso' ? 'corte ' + x.turno :
    'gasto') + ' ' + dinero(x.monto) + '. Ahora el contador recibe el efectivo.';
  return l;
}

/** Paso 3: el contador interno recibe el efectivo. `recibido` vacío = llegó completo. */
function recibirEfectivo(cred, id, recibido) {
  var yo = exigeConfirmar(cred);
  if (!yo.recibe) throw new Error('El efectivo lo recibe el contador interno.');
  var lock = LockService.getScriptLock();
  lock.waitLock(15000); _LEE = {};
  var x, rec = null, dif = 0;
  try {
    x = buscaPend(id);
    exigePaso(yo, x, 'recibir');
    var v = {}; try { v = JSON.parse(x.datos); } catch (e) {}
    if (x.tipo === 'ingreso' && recibido !== '' && recibido != null) {
      rec = r2(recibido);
      if (!isFinite(rec) || rec < 0) throw new Error('Escriba cuánto efectivo recibió.');
      dif = r2(rec - ((Number(v.efectivo) || 0) - (Number(v.caja) || 0)));
    }
    var h = hojaPorConfirmar();
    h.getRange(x.fila, 11).setValue(EST_REC);
    h.getRange(x.fila, 18, 1, 4).setValues([[rec == null ? '' : rec, rec == null ? '' : dif,
      yo.nombre, new Date()]]);
  } finally { lock.releaseLock(); }
  var l = listaPorConfirmar(yo);
  l.ok = true;
  l.mensaje = 'Efectivo recibido: ' + x.sucursal + ' · corte ' + x.turno +
    (dif ? ' · ' + (dif > 0 ? 'sobran ' : 'faltan ') + dinero(Math.abs(dif)) : '') +
    '. Ahora el director financiero confirma la tarjeta en el banco.';
  return l;
}

/** Paso 4: el director financiero anota la tarjeta que dice el cierre del POS NEONET
 *  (obligatorio, 0 si no hubo) y confirma que llegó al banco. La comisión se calcula
 *  sola. Ahí entra a las cifras oficiales. */
function confirmarRegistro(cred, id, tarjeta) {
  var yo = exigeConfirmar(cred);
  if (!yo.banco) throw new Error('Esto lo confirma el contador interno, cuando el dinero llega al banco.');
  var x = buscaPend(id);
  exigePaso(yo, x, 'banco');
  var v = JSON.parse(x.datos);
  var ahora = new Date(), pct = tasaComision(), tn = null, com = 0, notas = [];
  if (x.tipo === 'ingreso') {
    if (tarjeta === '' || tarjeta == null)
      throw new Error('Escriba lo que dice el cierre del POS NEONET en tarjeta (0 si no hubo).');
    tn = r2(tarjeta);
    if (!isFinite(tn) || tn < 0) throw new Error('La tarjeta va en números.');
    if (Math.abs(tn - (Number(v.tarjeta) || 0)) > 0.004)
      notas.push('Tarjeta según NEONET ' + dinero(tn) + ' (el gerente reportó ' + dinero(v.tarjeta || 0) + ')');
    v.tarjeta = tn;
    v.monto = r2((Number(v.efectivo) || 0) + tn);
    v.desglosado = true;
    com = comisionDe(tn, pct);
    if (x.diferencia) notas.push('El contador recibió ' + dinero(x.recibido) + ' de efectivo (' +
      (x.diferencia > 0 ? 'sobran ' : 'faltan ') + dinero(Math.abs(x.diferencia)) + ')');
    if (notas.length) v.nota = [v.nota].concat(notas).filter(Boolean).join(' · ').slice(0, 300);
  }
  var conf = { por: yo.nombre, en: ahora, valPor: x.validadoPor || '', valEn: x.validadoEn || '',
    recPor: x.recibidoPor || '', recEn: x.recibidoEn || '' };
  var cuando = x.enviadoEn instanceof Date ? x.enviadoEn : ahora;
  var h = hojaPorConfirmar();
  var antes = [x.estado, x.revisadoPor, x.revisadoEn];
  // se marca antes de escribir: si alguien más aprieta al mismo tiempo, no se duplica
  h.getRange(x.fila, 11, 1, 3).setValues([[EST_CONF, yo.nombre, ahora]]);
  var r;
  try {
    if (x.tipo === 'ingreso') r = escribeIngreso(v, x.enviadoPor, cuando, conf, x.id);
    else {
      r = escribeGasto(v, x.enviadoPor, cuando, conf);
      if (v.proveedor) { try { aseguraProveedor(v.proveedor, x.enviadoPor, v.unidad); } catch (e) {} }
      recuerdaPrecios(v.lineas || []);
    }
  } catch (e) {
    h.getRange(x.fila, 11, 1, 3).setValues([[antes[0], antes[1], antes[2]]]);
    throw e;
  }
  h.getRange(x.fila, 15).setValue(r.nuevo.id);
  if (tn != null) h.getRange(x.fila, 22, 1, 3).setValues([[tn, com, r2(tn - com)]]);
  var l = listaPorConfirmar(yo);
  l.ok = true; l.nuevo = r.nuevo;
  l.mensaje = 'Confirmado: ' + x.sucursal + ' · ' + (x.tipo === 'ingreso' ? 'corte ' + x.turno
    : 'gasto') + ' ' + dinero(r.nuevo.monto) + (tn ? ' · tarjeta ' + dinero(tn) + ', comisión NEONET ' +
    dinero(com) + ', al banco ' + dinero(tn - com) : '') + '.';
  return l;
}

function rechazarRegistro(cred, id, motivo) {
  var yo = exigeConfirmar(cred);
  motivo = String(motivo || '').trim().slice(0, 200);
  if (motivo.length < 4) throw new Error('Escriba por qué se rechaza, para que el gerente lo corrija.');
  var x = buscaPend(id);
  if (!enCamino(x)) throw new Error('Ese registro ya fue revisado.');
  if (!puedeEnPaso(yo, x))
    throw new Error('Ahora lo revisa ' + QUIEN_PASO[pasoDe(x)] + '.');
  hojaPorConfirmar().getRange(x.fila, 11, 1, 4)
    .setValues([[EST_RECH, yo.nombre, new Date(), motivo]]);
  var l = listaPorConfirmar(yo);
  l.ok = true; l.mensaje = 'Rechazado. El gerente lo verá con el motivo.';
  return l;
}

/** Los cierres de caja de los últimos días: sucursal × día × turno, y en qué paso va
 *  cada uno. Lo ven administración, el director operativo y finanzas en Inicio. */
function cierresCaja(dias) {
  dias = dias || 7;
  var hoy = hoyISO(), desde = masDias(hoy, -(dias - 1)), celdas = {};
  ultimas(libro().getSheetByName('Ingresos'), H_INGRESOS.length, 1500).forEach(function (s) {
    if (!s[0] || String(s[10]).toLowerCase() === 'sí') return;
    var f = fmtDia(s[5]);
    if (f < desde || f > hoy) return;
    var e = Number(s[14]) || 0, t = Number(s[15]) || 0;
    if (!(e > 0 || t > 0)) e = Number(s[8]) || 0;
    var com = t > 0 ? comisionFila(s) : 0;
    celdas[s[4] + '|' + f + '|' + s[7]] = { e: s[18] ? 'Confirmado' : 'Registrado',
      m: Number(s[8]) || 0, ef: e, tj: t, cj: Number(s[16]) || 0, dc: String(s[17] || ''),
      com: com, liq: r2(t - com), nota: String(s[9] || ''),
      p: String(s[2] || ''), pEn: fmtSello(s[1]), c: String(s[18] || ''), cEn: fmtSello(s[19]),
      v: String(s[20] || ''), vEn: fmtSello(s[21]), r: String(s[22] || ''), rEn: fmtSello(s[23]) };
  });
  filasPorConfirmar().forEach(function (x) {
    if (x.tipo !== 'ingreso' || x.fecha < desde || x.fecha > hoy) return;
    var k = x.sucursal + '|' + x.fecha + '|' + x.turno;
    if (celdas[k] && celdas[k].e !== 'Rechazado') return;
    if (x.estado === EST_CONF) return;
    var v = {}; try { v = JSON.parse(x.datos); } catch (e) {}
    celdas[k] = { e: x.estado === EST_PEND ? 'Por validar' : x.estado === EST_RECH ? 'Rechazado' : x.estado,
      m: x.monto, ef: v.efectivo || 0, tj: v.tarjeta || 0, cj: v.caja || 0, dc: v.detalleCaja || '',
      nota: v.nota || '', p: x.enviadoPor, pEn: x.sello, v: x.validadoPor, vEn: x.validadoEn,
      r: x.recibidoPor, rEn: x.recibidoEn, rec: x.recibido, dif: x.diferencia, id: x.id,
      motivo: x.motivo, rechPor: x.estado === EST_RECH ? x.revisadoPor : '' };
  });
  var fechas = [];
  for (var k = 0; k < dias; k++) fechas.push(masDias(hoy, -k));
  return { hoy: hoy, fechas: fechas, tasa: tasaComision(),
    sucursales: sucursalesDestino().map(function (u) { return u.nombre; }), celdas: celdas };
}

/** Rellena la comisión de los cortes anotados antes de que existiera la columna. */
function comisionesViejas(ss) {
  var h = (ss || libro()).getSheetByName('Ingresos');
  if (!h || h.getLastRow() < 2) return;
  var v = h.getRange(2, 1, h.getLastRow() - 1, H_INGRESOS.length).getValues();
  var pct = tasaComision(), n = 0;
  var out = v.map(function (s) {
    var t = Number(s[15]) || 0;
    if (s[24] !== '' && s[24] != null) return [s[24], s[25], s[26]];
    if (!(t > 0)) return [s[24], s[25], s[26]];
    n++;
    var c = comisionDe(t, pct);
    return [c, r2(t - c), pct];
  });
  if (n) h.getRange(2, 25, out.length, 3).setValues(out);
}
/** Agrega a «Ajustes de inventario» las claves nuevas que todavía no tenga. */
function ajustesQueFaltan() {
  var h = hojaInv('Ajustes de inventario'), hay = {};
  leeTodo(h, 1).forEach(function (r) { hay[String(r[0]).trim()] = true; });
  AJUSTES_INICIALES.forEach(function (x) {
    if (hay[x[0]]) return;
    var val = x[0] === 'impuestos_desde'
      ? Utilities.formatDate(new Date(Date.now() - 31 * 86400000), ZONA, 'yyyy-MM') : x[1];
    h.appendRow([x[0], val, x[2]]);
    h.getRange(h.getLastRow(), 2).setNumberFormat('@').setValue(val);
  });
  _AJ = null;
  try { CacheService.getScriptCache().remove('ajustes_inv'); } catch (e) {}
}

/* ════════════ INVENTARIO, SOLICITUDES DE ABASTECIMIENTO Y COMPRAS ════════════
 *
 * Cada unidad (las tres sucursales y la bodega) cuenta su inventario en sus días
 * (domingo y jueves, se cambia en la hoja «Ajustes de inventario»). La hora la pone
 * el servidor y el conteo ya no se puede cambiar: si se equivocaron, se cuenta otra
 * vez y vale el último.
 *
 * Lo que hay en una unidad = el último conteo + lo que entró o salió después
 * (hoja «Inventario movimientos»). Todo se lleva en la unidad en que se traslada el
 * producto (caja, fardo…); lo suelto se convierte: 3 bolsas de una caja de 6 = 0.5.
 *
 * El ciclo:
 *   1. El gerente cuenta → el sistema le sugiere qué pedir → él sube o baja y envía
 *      la solicitud (SA-0001). Sin conteo del día no se puede pedir.
 *   2. Samuel ve las solicitudes, las acepta (puede ajustar lo que manda) e imprime
 *      la hoja de cada sucursal para que se la firmen.
 *   3. Al llegar, el gerente recibe artículo por artículo. Lo recibido entra a su
 *      inventario, sale del de bodega y se vuelve costo de la sucursal (traslado
 *      TR-…, igual que los de siempre). Si llegó menos, la diferencia se queda en
 *      bodega y queda «por aclarar»: regresó a bodega o fue pérdida.
 *   4. Samuel cuenta la bodega → el sistema sugiere la orden de compra por proveedor,
 *      en la presentación en que se compra. Al llegar, la recibe y entra a bodega.
 *
 * La sugerencia: al principio sale del «nivel ideal» que se pone a cada producto
 * (Configuración → Productos de bodega). Cuando ya hay dos semanas de conteos, sale
 * del consumo real: lo que había + lo que entró − lo que hay ahora. */

var H_CONTEOS = ['ID conteo', 'Contado en', 'Fecha', 'Unidad', 'Contado por', 'Código',
  'Producto', 'Se cuenta por', 'Enteros', 'Sueltos', 'Total', 'Menores sueltos', 'Peso de lo abierto'];
var H_KARDEX = ['Fecha y hora', 'Unidad', 'Código', 'Producto', 'Movimiento', 'Cantidad',
  'Referencia', 'Registrado por', 'Nota'];
var H_SOL = ['No. de solicitud', 'Sucursal', 'Enviada en', 'Enviada por', 'ID conteo',
  'Entrega', 'Estado', 'Aceptada en', 'Aceptada por', 'Recibida en', 'Recibida por',
  'No. de traslado', 'Nota', 'Anulada por', 'Motivo de anulación',
  'Vista en', 'Preparada en', 'Despachada en', 'Despachada por', 'Llegó en'];
var H_SOLD = ['No. de solicitud', 'Sucursal', 'Código', 'Producto', 'Medida', 'Clave',
  'Existencia al contar', 'Sugerido', 'Base de la sugerencia', 'Pedido', 'Enviado',
  'Recibido', 'Diferencia', 'Diferencia: estado', 'Resuelta por', 'Resuelta en',
  'Pendiente', 'Pendiente: llega', 'Pendiente: entregado', 'Pendiente: no se entrega (quién)', 'Pendiente: motivo'];
var H_OC = ['No. de orden', 'Creada en', 'Creada por', 'Proveedor', 'Producto de compra',
  'Presentación', 'Sugerido', 'Pedido', 'Código', 'Producto de bodega', 'Rinde',
  'Estado', 'Recibido', 'Recibido en', 'Recibido por', 'Precio (cotización)', 'Subtotal'];
var H_AJUSTES = ['Clave', 'Valor', 'Qué es'];
var AJUSTES_INICIALES = [
  ['dias_conteo_sucursales', 'domingo, jueves', 'Días en que cada sucursal hace inventario'],
  ['dias_conteo_bodega', 'domingo, jueves', 'Días en que la bodega hace inventario'],
  ['dias_entrega', 'lunes, viernes', 'Días en que la bodega abastece a las sucursales'],
  ['colchon_dias', '1', 'Días de más que se piden, por seguridad'],
  ['cobertura_compra_dias', '7', 'Para cuántos días de despacho compra la bodega'],
  ['comision_tarjeta_pct', '6.11', '% que retiene NEONET de cada pago con tarjeta'],
  ['impuestos_desde', '', 'Primer mes (aaaa-mm) desde el que se siguen los pagos de impuestos'],
  ['venta_mensual', '400000', 'Venta presupuestada de la red al mes (trae la utilidad bruta y la neta adentro)'],
  ['sueldo_cmo', '3000', 'Sueldo mensual de la CMO que viene puesto en su presupuesto'],
  ['precios_gerentes', 'No', 'Si los gerentes ven los precios de lo que reciben de bodega (lo cambian el director operativo y el administrador)']
];
var DIAS_SEM = ['domingo', 'lunes', 'martes', 'miércoles', 'jueves', 'viernes', 'sábado'];
var SOL = { ENV: 'Enviada', ACE: 'Aceptada', CAM: 'En camino', REC: 'Recibida',
  DIF: 'Recibida con diferencias', ANU: 'Anulada' };
var DIF = { PEND: 'Por aclarar', BOD: 'Regresó a bodega', MERMA: 'Pérdida',
  OK: 'Aclarado' };
var MOV = { CONTEO: 'Conteo', ENT_AB: 'Entrada por abastecimiento',
  SAL_AB: 'Salida por abastecimiento', ENT_TR: 'Entrada por traslado',
  SAL_TR: 'Salida por traslado', COMPRA: 'Compra recibida', MERMA: 'Pérdida' };
var DIAS_CONSUMO = 14;            // con menos historia, la sugerencia usa el nivel ideal
var VENTANA_CONSUMO = 35;

/* ── hojas ── */

var INV_HOJAS = [
  ['Inventario conteos', H_CONTEOS, function (h) {
    h.getRange('B:B').setNumberFormat('dd/mm/yyyy hh:mm:ss');
    h.getRange('C:C').setNumberFormat('@'); h.setColumnWidth(7, 280); }],
  ['Inventario movimientos', H_KARDEX, function (h) {
    h.getRange('A:A').setNumberFormat('dd/mm/yyyy hh:mm:ss'); h.setColumnWidth(4, 280); }],
  ['Solicitudes', H_SOL, function (h) {
    h.getRange('C:C').setNumberFormat('dd/mm/yyyy hh:mm');
    h.getRange('F:F').setNumberFormat('@');
    h.getRange('H:H').setNumberFormat('dd/mm/yyyy hh:mm');
    h.getRange('J:J').setNumberFormat('dd/mm/yyyy hh:mm'); }],
  ['Solicitudes detalle', H_SOLD, function (h) {
    h.getRange('P:P').setNumberFormat('dd/mm/yyyy hh:mm'); h.setColumnWidth(4, 280); }],
  ['Órdenes de compra', H_OC, function (h) {
    h.getRange('B:B').setNumberFormat('dd/mm/yyyy hh:mm');
    h.getRange('N:N').setNumberFormat('dd/mm/yyyy hh:mm'); h.setColumnWidth(5, 240); }],
  ['Ajustes de inventario', H_AJUSTES, function (h) {
    h.getRange('B:B').setNumberFormat('@');
    h.getRange(2, 1, AJUSTES_INICIALES.length, 3).setValues(AJUSTES_INICIALES);
    h.setColumnWidth(1, 190); h.setColumnWidth(2, 160); h.setColumnWidth(3, 360); }]
];

function hojasInventario(ss) {
  ss = ss || libro();
  INV_HOJAS.forEach(function (x) { hojaInv(x[0], ss); });
}
var _HINV = {};
function hojaInv(nombre, ss) {
  if (_HINV[nombre]) return _HINV[nombre];
  ss = ss || libro();
  var h = ss.getSheetByName(nombre);
  if (h) {
    if (nombre === 'Inventario conteos' && h.getLastColumn() < H_CONTEOS.length)
      encabezaAlFinal(h, 12, ['Menores sueltos', 'Peso de lo abierto']);
    if (nombre === 'Solicitudes detalle' && h.getLastColumn() < H_SOLD.length)
      encabezaAlFinal(h, 17, ['Pendiente', 'Pendiente: llega', 'Pendiente: entregado', 'Pendiente: no se entrega (quién)', 'Pendiente: motivo']);
    return (_HINV[nombre] = h);
  }
  var x = INV_HOJAS.filter(function (y) { return y[0] === nombre; })[0];
  h = hojaLimpia(ss, nombre, x[1]);
  x[2](h);
  return (_HINV[nombre] = h);
}
/** Las últimas `n` filas: el inventario solo necesita lo reciente. */
function ultimas(h, ancho, n) {
  if (!h || h.getLastRow() < 2) return [];
  var desde = Math.max(2, h.getLastRow() - n + 1);
  return h.getRange(desde, 1, h.getLastRow() - desde + 1,
    Math.min(ancho, h.getMaxColumns())).getValues();
}
function ms(v) { return v instanceof Date ? v.getTime() : (v ? new Date(v).getTime() : 0); }
function r3(n) { return Math.round((Number(n) || 0) * 1000) / 1000; }

/* ── ajustes y calendario ── */

var _AJ = null;
function ajustesInv() {
  if (_AJ) return _AJ;
  var a = null;
  try { var c = CacheService.getScriptCache().get('ajustes_inv'); if (c) a = JSON.parse(c); }
  catch (e) {}
  if (!a) {
    a = {};
    AJUSTES_INICIALES.forEach(function (x) { a[x[0]] = x[1]; });
    leeTodo(hojaInv('Ajustes de inventario'), 2).forEach(function (r) {
      if (String(r[0]).trim()) a[String(r[0]).trim()] = String(r[1] == null ? '' : r[1]).trim();
    });
    try { CacheService.getScriptCache().put('ajustes_inv', JSON.stringify(a), 600); } catch (e) {}
  }
  var dias = function (txt) {
    var out = [];
    String(txt || '').toLowerCase().split(/[,;\s]+/).forEach(function (w) {
      w = w.replace('miercoles', 'miércoles').replace('sabado', 'sábado');
      var i = DIAS_SEM.indexOf(w);
      if (i >= 0 && out.indexOf(i) < 0) out.push(i);
    });
    return out.sort();
  };
  _AJ = {
    conteoSuc: dias(a.dias_conteo_sucursales), conteoBod: dias(a.dias_conteo_bodega),
    entrega: dias(a.dias_entrega),
    colchon: Math.max(0, Number(a.colchon_dias) || 0),
    cobertura: Math.max(1, Number(a.cobertura_compra_dias) || 7),
    comision: String(a.comision_tarjeta_pct).trim() !== '' && isFinite(Number(a.comision_tarjeta_pct))
      ? Number(a.comision_tarjeta_pct) : 6.11,
    impuestosDesde: /^\d{4}-\d{2}$/.test(String(a.impuestos_desde || '')) ? String(a.impuestos_desde) : '',
    ventaMensual: Number(a.venta_mensual) > 0 ? Number(a.venta_mensual) : 400000,
    preciosGerentes: /^s/i.test(String(a.precios_gerentes || '').trim()),
    sueldoCmo: Number(a.sueldo_cmo) > 0 ? Number(a.sueldo_cmo) : 3000
  };
  return _AJ;
}
function diaSemana(iso) { return aDate(iso).getDay(); }
function masDias(iso, n) {
  var d = aDate(iso); d.setDate(d.getDate() + n);
  return Utilities.formatDate(d, ZONA, 'yyyy-MM-dd');
}
function diasEntre(a, b) { return Math.round((aDate(b) - aDate(a)) / 86400000); }
/** El primer día después de `iso` que cae en uno de `dias`. */
function proximoDia(iso, dias) {
  if (!dias.length) return masDias(iso, 1);
  for (var k = 1; k <= 7; k++) {
    var f = masDias(iso, k);
    if (dias.indexOf(diaSemana(f)) >= 0) return f;
  }
  return masDias(iso, 7);
}
function nombresDias(dias) { return dias.map(function (i) { return DIAS_SEM[i]; }); }

/** ¿Le toca contar hoy? ¿Se le pasó? */
function calendarioConteo(unidadId, ultimoFecha, aj, hoy) {
  var dias = unidadId === 'bodega' ? aj.conteoBod : aj.conteoSuc;
  var toca = dias.indexOf(diaSemana(hoy)) >= 0;
  var tocaba = '';                       // el último día programado, hoy incluido
  for (var k = 0; k < 7 && dias.length; k++) {
    var f = masDias(hoy, -k);
    if (dias.indexOf(diaSemana(f)) >= 0) { tocaba = f; break; }
  }
  var hecho = !!(ultimoFecha && tocaba && ultimoFecha >= tocaba);
  return { dias: nombresDias(dias), tocaHoy: toca, hechoHoy: ultimoFecha === hoy,
    tocaba: tocaba, alDia: hecho || !tocaba,
    atrasado: !!tocaba && !hecho && tocaba < hoy,
    proximo: toca && !hecho ? hoy : proximoDia(hoy, dias) };
}

/* ── productos que maneja cada unidad ── */

function productosInv() {
  return filasBodega().filter(function (b) { return b.activo && b.codigo; });
}
function maneja(b, unidadId) { return unidadId === 'bodega' || !!b.suc[unidadId]; }
function tieneSuelto(b) { return !!(b.suelto && b.porUnidad > 0); }       // se cuenta suelto
/** Lo que se pesa cuando está abierto: la unidad más pequeña, y cuántas hay en la principal. */
function nivelPeso(b) {
  if (b.menor && b.porSuelto > 0 && tieneSuelto(b)) return { nombre: b.menor, n: b.porUnidad * b.porSuelto };
  if (tieneSuelto(b)) return { nombre: b.suelto, n: b.porUnidad };
  return { nombre: b.unidad, n: 1 };
}
/** «Caja completa trae 6 paquetes de 10 unidades (60 unidades) · cada unidad pesa 2 onzas». */
function presentacionTxt(b) {
  if (b.medida === 'peso') return 'Se cuenta por peso, en ' + plur(2, b.unidad);
  var t = '';
  if (b.suelto && b.porUnidad > 0) {
    t = capital(b.unidad) + (/(a|dad)$/i.test(b.unidad) ? ' completa' : ' completo') + ' trae ' + num(b.porUnidad) + ' ' + plur(b.porUnidad, b.suelto);
    if (b.menor && b.porSuelto > 0) t += ' de ' + num(b.porSuelto) + ' ' + plur(b.porSuelto, b.menor) + ' (' + num(b.porUnidad * b.porSuelto) + ' ' + plur(2, b.menor) + ')';
  } else t = 'Se cuenta por ' + b.unidad;
  if (b.pesoCada > 0 && b.pesoEn) t += ' · cada ' + nivelPeso(b).nombre + ' pesa ' + num(b.pesoCada) + ' ' + plur(b.pesoCada, b.pesoEn) + '; lo abierto se pesa';
  return t;
}
function capital(s) { s = String(s || ''); return s.charAt(0).toUpperCase() + s.slice(1); }
function seMandaSuelto(b) { return !!(b.mandaSuelto && tieneSuelto(b)); }  // y se pide suelto
/** 2.5 cajas de 6 bolsas → «2 cajas y 3 bolsas». */
function textoCantidad(q, b) {
  q = r3(q);
  if (tieneSuelto(b) && Math.abs(q - Math.round(q)) > 0.0005) {
    var e = Math.floor(q + 1e-9), s = Math.round((q - e) * b.porUnidad * 100) / 100;
    return (e ? num(e) + ' ' + plur(e, b.unidad) + ' y ' : '') + num(s) + ' ' + plur(s, b.suelto);
  }
  return num(q) + ' ' + plur(q, b.unidad);
}

/* ── lo que hay: último conteo + movimientos ── */

function estadoInventarioCalc(unidadId) {
  var u = unidadPorId(unidadId);
  if (!u) throw new Error('No existe esa unidad.');
  var conteos = {}, orden = [];
  ultimas(hojaInv('Inventario conteos'), H_CONTEOS.length, 40000).forEach(function (r) {
    if (String(r[3]) !== u.nombre || !r[0]) return;
    var id = String(r[0]);
    if (!conteos[id]) {
      conteos[id] = { id: id, t: ms(r[1]), sello: fmtSello(r[1]), fecha: fmtDia(r[2]) ||
        fmtDia(r[1]), quien: String(r[4] || ''), q: {} };
      orden.push(conteos[id]);
    }
    conteos[id].q[String(r[5])] = Number(r[10]) || 0;
  });
  orden.sort(function (a, b) { return a.t - b.t; });
  var ultimo = orden.length ? orden[orden.length - 1] : null;
  var movs = [];
  ultimas(hojaInv('Inventario movimientos'), H_KARDEX.length, 40000).forEach(function (r) {
    if (String(r[1]) !== u.nombre || !r[2]) return;
    movs.push({ t: ms(r[0]), codigo: String(r[2]), tipo: String(r[4] || ''),
      q: Number(r[5]) || 0 });
  });
  var ex = {};
  if (ultimo) Object.keys(ultimo.q).forEach(function (c) { ex[c] = ultimo.q[c]; });
  movs.forEach(function (m) {
    if (ultimo && m.t <= ultimo.t) return;
    ex[m.codigo] = (ex[m.codigo] || 0) + m.q;
  });
  Object.keys(ex).forEach(function (c) { ex[c] = r3(ex[c]); });
  return { unidad: u, conteos: orden, ultimo: ultimo, movs: movs, existencia: ex };
}
/** El inventario de una unidad (conteos, movimientos y existencia): se calcula una vez por consulta. */
function estadoInventario(unidadId) {
  var k = 'inv:' + unidadId;
  return _MEMO[k] || (_MEMO[k] = estadoInventarioCalc(unidadId));
}

/** Consumo por día de cada producto, o nada si todavía no hay historia suficiente.
 *  Sucursal: lo que había + lo que entró − lo que hay, entre conteo y conteo.
 *  Bodega: lo que despachó a las sucursales en las últimas semanas. */
function consumoPorDia(est, ahoraMs) {
  var out = {}, limite = ahoraMs - VENTANA_CONSUMO * 86400000;
  if (est.unidad.id === 'bodega') {
    var primero = est.conteos.length ? est.conteos[0].t : 0;
    est.movs.forEach(function (m) { if (!primero || m.t < primero) primero = m.t; });
    if (!primero) return out;
    var desde = Math.max(primero, ahoraMs - 28 * 86400000);
    var dias = (ahoraMs - desde) / 86400000;
    if (dias < DIAS_CONSUMO) return out;
    est.movs.forEach(function (m) {
      if (m.t < desde || m.q >= 0) return;
      if (m.tipo !== MOV.SAL_AB && m.tipo !== MOV.SAL_TR) return;
      out[m.codigo] = (out[m.codigo] || 0) - m.q;
    });
    Object.keys(out).forEach(function (c) { out[c] = out[c] / dias; });
    out._conHistoria = true;
    return out;
  }
  var cs = est.conteos.filter(function (c) { return c.t >= limite; });
  if (cs.length < 2) return out;
  var suma = {}, dias2 = {};
  for (var i = 1; i < cs.length; i++) {
    var a = cs[i - 1], b = cs[i], d = (b.t - a.t) / 86400000;
    if (d < 0.5) continue;                          // dos conteos el mismo día
    var ent = {};
    est.movs.forEach(function (m) {
      if (m.t > a.t && m.t <= b.t) ent[m.codigo] = (ent[m.codigo] || 0) + m.q;
    });
    Object.keys(b.q).forEach(function (c) {
      if (a.q[c] == null) return;
      suma[c] = (suma[c] || 0) + Math.max(0, a.q[c] + (ent[c] || 0) - b.q[c]);
      dias2[c] = (dias2[c] || 0) + d;
    });
  }
  Object.keys(suma).forEach(function (c) {
    if (dias2[c] >= DIAS_CONSUMO) out[c] = suma[c] / dias2[c];
  });
  return out;
}

/** Parte una cantidad (en unidades enteras, con decimales) en lo que se pide:
 *  cajas completas y, si el producto se puede mandar suelto, bolsas. */
function redondeaPedido(q, b) {
  if (!(q > 0.0005)) return { e: 0, s: 0 };
  if (seMandaSuelto(b)) {
    var e = Math.floor(q + 1e-9), s = Math.max(0, Math.ceil((q - e) * b.porUnidad - 1e-6));
    if (s >= b.porUnidad) { e++; s = 0; }
    return { e: e, s: s };
  }
  return { e: Math.ceil(q - 1e-6), s: 0 };
}

/** La sugerencia de una sucursal, recién contada. */
function sugerenciaSucursal(unidadId, est, hoy) {
  var aj = ajustesInv();
  var e1 = proximoDia(hoy, aj.entrega), e2 = proximoDia(e1, aj.entrega);
  var dias = diasEntre(hoy, e2) + aj.colchon;
  var cons = consumoPorDia(est, Date.now());
  var lineas = [];
  productosInv().forEach(function (b) {
    if (!maneja(b, unidadId)) return;
    var q = est.existencia[b.codigo] || 0, falta = 0, base = '';
    if (cons[b.codigo] != null) {
      falta = cons[b.codigo] * dias - q;
      base = 'consumo ' + num(Math.round(cons[b.codigo] * 100) / 100) + '/día';
    } else if (b.ideal[unidadId] > 0) {
      falta = b.ideal[unidadId] - q; base = 'nivel ideal ' + num(b.ideal[unidadId]);
    } else base = 'sin nivel ideal';
    var p = redondeaPedido(falta, b);
    var fila = { codigo: b.codigo, nombre: b.nombre, area: b.area, grupo: b.grupo,
      estante: b.estante, existencia: q, existenciaTxt: textoCantidad(q, b), base: base };
    lineas.push(Object.assign({}, fila, { clave: llaveT(b.nombre, b.unidad),
      medida: b.unidad, suelto: false, sugerido: p.e }));
    if (seMandaSuelto(b))
      lineas.push(Object.assign({}, fila, { clave: llaveT(b.nombre, b.suelto),
        medida: b.suelto, suelto: true, porUnidad: b.porUnidad, sugerido: p.s }));
  });
  return { entrega: e1, siguiente: e2, dias: dias, lineas: lineas,
    conHistoria: Object.keys(cons).length > 0 };
}

/* ── permisos ── */

/** Quién puede tocar el inventario de qué unidad. */
function unidadDe(yo, pedida) {
  if (yo.rol === 'gerente') return yo.sucursal;
  if (yo.rol === 'bodega') return 'bodega';
  if (yo.esAdmin) return unidadPorId(pedida) ? pedida : 'bodega';
  throw new Error('Su usuario no maneja inventario.');
}
/** Ver el inventario completo sin tocarlo: el administrador y el director operativo. */
function veInv(yo) { return !!yo.esAdmin || yo.rol === 'operaciones'; }
function exigeVerBodega(cred) {
  var yo = quien(cred);
  if (yo.rol !== 'bodega' && !veInv(yo)) throw new Error('Eso lo ve la bodega o un administrador.');
  return yo;
}
function exigeBodega(cred) {                 // Samuel o un administrador
  var yo = quien(cred);
  if (yo.rol !== 'bodega' && !yo.esAdmin)
    throw new Error('Eso lo hace la bodega o un administrador.');
  return yo;
}

/* ── el conteo ── */

/** La pantalla de conteo: qué productos, cuándo fue el último y si toca hoy. */
/* Un conteo programado se bloquea después de hacerlo y se habilita el próximo día que
 * toca (o si está atrasado). El director operativo puede permitir uno más ese día. */
function ajusteCrudo(k) {
  var v = ''; leeTodo(hojaInv('Ajustes de inventario'), 2).forEach(function (r) { if (String(r[0]).trim() === k) v = String(r[1] == null ? '' : r[1]); });
  return v;
}
function ponAjusteCrudo(k, v, nota) {
  var h = hojaInv('Ajustes de inventario'), i = -1;
  leeTodo(h, 1).forEach(function (r, j) { if (String(r[0]).trim() === k) i = j; });
  if (i >= 0) h.getRange(i + 2, 2).setNumberFormat('@').setValue(v);
  else { h.appendRow([k, v, nota || '']); h.getRange(h.getLastRow(), 2).setNumberFormat('@').setValue(v); }
}
function estadoBloqueoConteo(unidadId, cal, ult, hoy) {
  var extra = ajusteCrudo('conteo_extra_' + unidadId);
  var extraHoy = extraVigente(extra, ult, hoy);
  var abierto = !ult || cal.atrasado || (cal.tocaHoy && !cal.hechoHoy) || extraHoy;
  return { bloqueado: !abierto, extraHoy: extraHoy };
}
function extraVigente(extra, ult, hoy) {
  return String(extra || '').slice(0, 10) === hoy && (!ult || ult.t < Number(String(extra).split('|')[1] || 0));
}
function habilitarConteo(cred, unidadId) {
  var yo = quien(cred);
  if (yo.rol !== 'operaciones' && !yo.esAdmin) throw new Error('Esto lo decide el director operativo.');
  var u = unidadPorId(unidadId);
  if (!u || !u.vende) throw new Error('Escoja la sucursal.');
  var ahora = new Date();
  ponAjusteCrudo('conteo_extra_' + u.id, fmtSello(ahora) + '|' + ahora.getTime(), 'Conteo extra permitido por ' + yo.nombre);
  var r = resumenInventario(cred); r.ok = true; r.mensaje = u.nombre + ' puede contar otra vez hoy. Se vuelve a bloquear cuando guarde.';
  return r;
}
function pantallaConteo(cred, unidadId) {
  var yo = quien(cred);
  unidadId = unidadDe(yo, unidadId);
  var est = estadoInventario(unidadId), hoy = hoyISO(), aj = ajustesInv();
  var ult = est.ultimo;
  var zs = [], zp = {};
  if (unidadId !== 'bodega') { try { zs = zonasDe(unidadId); } catch (e) { zs = []; } }
  zs.forEach(function (z) { z.codigos.forEach(function (c) { (zp[c] = zp[c] || []).push(z.id); }); });
  var cal0 = calendarioConteo(unidadId, ult ? ult.fecha : '', aj, hoy);
  var blq = yo.rol === 'gerente' ? estadoBloqueoConteo(unidadId, cal0, ult, hoy) : { bloqueado: false, extraHoy: false };
  return { bloqueado: blq.bloqueado, extraHoy: blq.extraHoy,
    unidad: { id: est.unidad.id, nombre: est.unidad.nombre, marca: est.unidad.marca },
    hoy: hoy,
    calendario: calendarioConteo(unidadId, ult ? ult.fecha : '', aj, hoy),
    ultimo: ult ? { id: ult.id, sello: ult.sello, quien: ult.quien, fecha: ult.fecha } : null,
    productos: productosInv().filter(function (b) { return maneja(b, unidadId); })
      .map(function (b) {
        return { codigo: b.codigo, nombre: b.nombre, area: b.area, grupo: b.grupo,
          estante: b.estante, almacen: b.almacen, unidad: b.unidad,
          suelto: tieneSuelto(b) ? b.suelto : '', porUnidad: tieneSuelto(b) ? b.porUnidad : 0,
          deducido: !!b.deducido, medida: b.medida, menor: b.menor, porSuelto: b.porSuelto,
          pesoCada: b.pesoCada, pesoEn: b.pesoEn, pesa: b.pesoCada ? nivelPeso(b).nombre : '', presentacion: b.presentacion, descripcion: b.descripcion, zonas: zp[b.codigo] || [] };
      }),
    zonas: zs.map(function (z) { return { id: z.id, nombre: z.nombre }; }),
    solicitud: unidadId === 'bodega' ? null : solicitudDelConteo(unidadId, ult, hoy),
    cierreRef: cierreRefDe(unidadId, ult, hoy)
  };
}

function guardarConteo(cred, d) {
  var yo = quien(cred);
  d = d || {};
  var unidadId = unidadDe(yo, d.unidad);
  if (yo.rol === 'gerente') {
    var est0 = estadoInventario(unidadId), u0 = est0.ultimo, h0 = hoyISO();
    var cal = calendarioConteo(unidadId, u0 ? u0.fecha : '', ajustesInv(), h0);
    if (estadoBloqueoConteo(unidadId, cal, u0, h0).bloqueado)
      throw new Error(cal.hechoHoy ? 'El inventario de hoy ya se guardó. Si hay que corregirlo, pida al director operativo que le permita contar otra vez.'
        : 'Hoy no toca inventario. El próximo es el ' + dia(cal.proximo).slice(0, 5) + '.');
  }
  var u = unidadPorId(unidadId), cant = d.cantidades || {};
  var prods = productosInv().filter(function (b) { return maneja(b, unidadId); });
  var filas = [], ahora = new Date(), hoy = hoyISO();
  var id = 'C' + Utilities.formatDate(ahora, ZONA, 'yyyyMMddHHmmss') +
    unidadId.charAt(0).toUpperCase();
  var contados = 0;
  prods.forEach(function (b) {
    var x = cant[b.codigo] || {};
    var lee = function (v, que) {
      var n = Number(String(v == null ? '' : v).replace(',', '.'));
      if (String(v == null ? '' : v).trim() === '') return 0;
      if (!isFinite(n) || n < 0) throw new Error('La cantidad de «' + b.nombre + '» (' +
        que + ') no es válida.');
      if (n > 100000) throw new Error('La cantidad de «' + b.nombre + '» parece equivocada.');
      return r3(n);
    };
    var e = lee(x.e, b.unidad), s = tieneSuelto(b) ? lee(x.s, b.suelto) : 0;
    var m = b.menor ? lee(x.m, b.menor) : 0, p = b.pesoCada ? lee(x.p, b.pesoEn) : 0;
    var hay = function (v) { return String(v == null ? '' : v).trim() !== ''; };
    if (hay(x.e) || hay(x.s) || (b.menor && hay(x.m)) || (b.pesoCada && hay(x.p))) contados++;
    var total = r3(e + (tieneSuelto(b) ? s / b.porUnidad : 0) + (b.menor ? m / (b.porUnidad * b.porSuelto) : 0) +
      (b.pesoCada ? p / (b.pesoCada * nivelPeso(b).n) : 0));
    filas.push([id, ahora, hoy, u.nombre, yo.nombre, b.codigo, b.nombre, b.unidad, e,
      tieneSuelto(b) ? s : '', total, b.menor ? m : '', b.pesoCada ? p : '']);
  });
  if (!filas.length) throw new Error('Esta unidad no tiene productos para contar.');
  if (!contados) throw new Error('No escribió ninguna cantidad.');
  var lock = LockService.getScriptLock();
  lock.waitLock(20000); _LEE = {};
  try {
    var h = hojaInv('Inventario conteos');
    var f0 = h.getLastRow() + 1;
    h.getRange(f0, 3, filas.length, 1).setNumberFormat('@');
    h.getRange(f0, 1, filas.length, H_CONTEOS.length).setValues(filas);
  } finally {
    lock.releaseLock();
  }
  var r = { ok: true, id: id, sello: fmtSello(ahora), unidad: unidadId,
    mensaje: 'Inventario de ' + u.nombre + ' guardado: ' + filas.length + ' productos, ' +
      Utilities.formatDate(ahora, ZONA, 'dd/MM/yyyy HH:mm') + '.' };
  if (unidadId !== 'bodega') r.solicitud = pantallaSolicitud(cred);
  return r;
}

/* ── la solicitud de abastecimiento ── */

var _SOLOK = false;
function hojaSol() {
  var h = hojaInv('Solicitudes');
  if (!_SOLOK) {                     // las hojas de antes no traían el avance
    if (h.getLastColumn() < H_SOL.length)
      encabezaAlFinal(h, 16, ['Vista en', 'Preparada en', 'Despachada en', 'Despachada por',
        'Llegó en']);
    _SOLOK = true;
  }
  return h;
}
function hojaSolD() { return hojaInv('Solicitudes detalle'); }

function filasSol() {
  return leeTodo(hojaSol(), H_SOL.length).map(function (r, i) {
    return { fila: i + 2, numero: String(r[0] || ''), sucursal: String(r[1] || ''),
      enviadaEn: fmtSello(r[2]), enviadaPor: String(r[3] || ''), conteo: String(r[4] || ''),
      entrega: fmtDia(r[5]), estado: String(r[6] || ''), aceptadaEn: fmtSello(r[7]),
      aceptadaPor: String(r[8] || ''), recibidaEn: fmtSello(r[9]),
      recibidaPor: String(r[10] || ''), traslado: String(r[11] || ''),
      nota: String(r[12] || ''), anuladaPor: String(r[13] || ''),
      motivo: String(r[14] || ''), vistaEn: fmtSello(r[15]), preparadaEn: fmtSello(r[16]),
      despachadaEn: fmtSello(r[17]), despachadaPor: String(r[18] || ''),
      llegoEn: fmtSello(r[19]) };
  }).filter(function (x) { return x.numero; });
}
function lineasSolTodas() {
  var out = [];
  leeTodo(hojaSolD(), H_SOLD.length).forEach(function (r, i) {
    out.push({ fila: i + 2, numero: String(r[0]), sucursal: String(r[1] || ''),
      codigo: String(r[2] || ''), nombre: String(r[3] || ''), medida: String(r[4] || ''),
      clave: String(r[5] || ''), existencia: Number(r[6]) || 0,
      sugerido: Number(r[7]) || 0, base: String(r[8] || ''), pedido: Number(r[9]) || 0,
      enviado: r[10] === '' || r[10] == null ? null : Number(r[10]),
      recibido: r[11] === '' || r[11] == null ? null : Number(r[11]),
      diferencia: Number(r[12]) || 0, difEstado: String(r[13] || ''),
      resueltaPor: String(r[14] || ''),
      pend: Number(r[16]) || 0, pendLlega: String(r[17] || ''), pendEntregado: Number(r[18]) || 0, pendCanPor: String(r[19] || ''), pendMotivo: String(r[20] || '') });
  });
  return out;
}
/** Las líneas de los pedidos (de uno solo, si se da el número). La hoja se recorre una vez por consulta. */
function lineasSol(numero) {
  var todas = _MEMO.lsol || (_MEMO.lsol = lineasSolTodas());
  if (!numero) return todas;
  numero = String(numero);
  return todas.filter(function (l) { return l.numero === numero; });
}
function siguienteSol() {
  var max = 0;
  leeTodo(hojaSol(), 1).forEach(function (r) {
    var m = /^SA-(\d+)$/.exec(String(r[0]).trim());
    if (m) max = Math.max(max, Number(m[1]));
  });
  var n = String(max + 1);
  return 'SA-' + (n.length < 4 ? ('0000' + n).slice(-4) : n);
}

/** ¿Ya se pidió con este conteo? */
function solicitudDelConteo(unidadId, ult, hoy) {
  var u = unidadPorId(unidadId);
  if (!ult || ult.fecha !== hoy) return { puede: false,
    porque: 'Para pedir, primero haga el inventario de hoy.' };
  var ya = null;
  filasSol().forEach(function (s) {
    if (s.sucursal === u.nombre && s.conteo === ult.id && s.estado !== SOL.ANU) ya = s;
  });
  if (ya) return { puede: false, numero: ya.numero, porque: 'Ya envió la solicitud ' +
    ya.numero + ' con este inventario.' };
  return { puede: true, conteo: ult.id };
}

/** Lo que el gerente ve después de contar: lo sugerido, para subirlo o bajarlo. */
function pantallaSolicitud(cred) {
  var yo = quien(cred);
  var unidadId = unidadDe(yo, null);
  if (unidadId === 'bodega') throw new Error('La bodega no envía solicitudes.');
  var est = estadoInventario(unidadId), hoy = hoyISO();
  var sol = solicitudDelConteo(unidadId, est.ultimo, hoy);
  var r = { unidad: { id: unidadId, nombre: est.unidad.nombre }, hoy: hoy, estado: sol };
  if (!sol.puede) return r;
  var sug = sugerenciaSucursal(unidadId, est, hoy);
  r.entrega = sug.entrega; r.siguiente = sug.siguiente; r.lineas = sug.lineas;
  try {          // para que el gerente sepa en qué pide (caja de 20 unidades, saco de 50 libras…) y cuánto hay en bodega
    var exB = estadoInventario('bodega').existencia, porC = {}; filasBodega().forEach(function (b) { porC[b.codigo] = b; });
    r.lineas.forEach(function (l) { var b = porC[l.codigo]; if (!b) return; var qb = exB[l.codigo] || 0;
      l.presentacion = presentacionTxt(b); l.enBodega = r3(qb); l.enBodegaTxt = textoCantidad(qb, b); l.compraTxt = compraTxtDe(b);
      l.porUnidad = !l.suelto && b.suelto && b.porUnidad > 0 ? b.porUnidad : 0; l.sueltoNom = b.suelto || ''; });
  } catch (e) {}
  r.conHistoria = sug.conHistoria; r.conteo = est.ultimo.id;
  r.contado = est.ultimo.sello;
  return r;
}

function enviarSolicitud(cred, d) {
  var yo = quien(cred);
  d = d || {};
  var unidadId = unidadDe(yo, null);
  if (unidadId === 'bodega') throw new Error('La bodega no envía solicitudes.');
  var u = unidadPorId(unidadId), hoy = hoyISO();
  var lock = LockService.getScriptLock();
  lock.waitLock(20000); _LEE = {};
  try {
    var est = estadoInventario(unidadId);
    var sol = solicitudDelConteo(unidadId, est.ultimo, hoy);
    if (!sol.puede) throw new Error(sol.porque);
    if (d.conteo && d.conteo !== est.ultimo.id)
      throw new Error('Hay un inventario más nuevo. Vuelva a abrir la solicitud.');
    var sug = sugerenciaSucursal(unidadId, est, hoy);
    var ped = d.pedidos || {}, filas = [], hay = 0;
    var numero = siguienteSol();
    sug.lineas.forEach(function (l) {
      var crudo = ped[l.clave];
      var n = crudo == null || String(crudo).trim() === '' ? 0
        : Number(String(crudo).replace(',', '.'));
      if (!isFinite(n) || n < 0) throw new Error('La cantidad de «' + l.nombre +
        '» no es válida.');
      if (n > 10000) throw new Error('La cantidad de «' + l.nombre + '» parece equivocada.');
      n = r3(n);
      if (!(n > 0) && !(l.sugerido > 0)) return;
      if (n > 0) hay++;
      filas.push([numero, u.nombre, l.codigo, l.nombre, l.medida, l.clave,
        l.suelto ? '' : l.existencia, l.sugerido, l.base, n, '', '', '', '', '', '', '', '', '', '', '']);
    });
    if (!hay) throw new Error('No pidió nada. Si no necesita nada, no hace falta enviar.');
    var ahora = new Date();
    hojaSol().appendRow([numero, u.nombre, ahora, yo.nombre, est.ultimo.id, sug.entrega,
      SOL.ENV, '', '', '', '', '', String(d.nota || '').slice(0, 200), '', '']);
    hojaSol().getRange(hojaSol().getLastRow(), 6).setNumberFormat('@').setValue(sug.entrega);
    var h = hojaSolD(), f0 = h.getLastRow() + 1;
    h.getRange(f0, 1, filas.length, H_SOLD.length).setValues(filas);
    return { ok: true, numero: numero, entrega: sug.entrega,
      mensaje: 'Solicitud ' + numero + ' enviada a bodega: ' + hay +
        (hay === 1 ? ' producto' : ' productos') + '. Entrega: ' + DIAS_SEM[diaSemana(sug.entrega)] +
        ' ' + dia(sug.entrega) + '.' };
  } finally {
    lock.releaseLock();
  }
}

/* ── las solicitudes: la bodega las ve todas; el gerente, las suyas ── */

/* El avance de una solicitud, de 0 a 100 %: ocho pasos, del inventario a recibido.
 * Si un paso posterior ya pasó, los de antes cuentan aunque no tengan hora. */
var PASOS_SOL = ['Inventario', 'Pedido enviado', 'Bodega lo recibió', 'Aceptado',
  'Preparado', 'Cargado y en camino', 'Descargado en la sucursal', 'Recibido'];
function selloDeConteo(id) {
  var m = /^C(\d{4})(\d{2})(\d{2})(\d{2})(\d{2})/.exec(String(id || ''));
  return m ? m[1] + '-' + m[2] + '-' + m[3] + ' ' + m[4] + ':' + m[5] : '';
}
function avanceSol(s) {
  var rec = s.estado === SOL.REC || s.estado === SOL.DIF;
  var cam = rec || s.estado === SOL.CAM;
  var etapas = [
    { n: PASOS_SOL[0], cuando: selloDeConteo(s.conteo), quien: s.enviadaPor, ok: true },
    { n: PASOS_SOL[1], cuando: s.enviadaEn, quien: s.enviadaPor, ok: true },
    { n: PASOS_SOL[2], cuando: s.vistaEn, quien: 'Bodega', ok: !!s.vistaEn },
    { n: PASOS_SOL[3], cuando: s.aceptadaEn, quien: s.aceptadaPor, ok: !!s.aceptadaEn },
    { n: PASOS_SOL[4], cuando: s.preparadaEn, quien: s.aceptadaPor, ok: !!s.preparadaEn },
    { n: PASOS_SOL[5], cuando: s.despachadaEn, quien: s.despachadaPor, ok: cam },
    { n: PASOS_SOL[6], cuando: s.llegoEn, quien: s.recibidaPor, ok: !!s.llegoEn || rec },
    { n: PASOS_SOL[7], cuando: s.recibidaEn, quien: s.recibidaPor, ok: rec }
  ];
  var ultimo = 0;
  etapas.forEach(function (e, i) { if (e.ok) ultimo = i; });
  etapas.forEach(function (e, i) { if (i <= ultimo) e.ok = true; if (!e.ok) e.quien = ''; });
  return { paso: ultimo + 1, pct: Math.round((ultimo + 1) / etapas.length * 100),
    etapa: s.estado === SOL.ANU ? 'Anulada' : etapas[ultimo].n, etapas: etapas };
}
function vistaSol(s, lineas) {
  var ls = lineas.filter(function (l) { return l.numero === s.numero; });
  var pend = ls.filter(function (l) { return l.difEstado === DIF.PEND; }).length;
  return { numero: s.numero, sucursal: s.sucursal, enviadaEn: s.enviadaEn,
    enviadaPor: s.enviadaPor, entrega: s.entrega, estado: s.estado,
    aceptadaEn: s.aceptadaEn, aceptadaPor: s.aceptadaPor, recibidaEn: s.recibidaEn,
    recibidaPor: s.recibidaPor, traslado: s.traslado, nota: s.nota, motivo: s.motivo,
    productos: ls.filter(function (l) { return l.pedido > 0 || l.enviado > 0; }).length,
    porAclarar: pend, despachadaEn: s.despachadaEn, avance: avanceSol(s),
    pendientes: ls.filter(function (l) { return pendDeLinea(l) > 0.0005; }).length };
}

function listaSolicitudes(cred) {
  var yo = quien(cred);
  var todas = filasSol(), lineas = lineasSol(), desde = masDias(hoyISO(), -45);
  var mias = yo.rol === 'gerente' ? nombreUnidad(yo.sucursal) : '';
  if (yo.rol !== 'gerente' && yo.rol !== 'bodega' && !veInv(yo))
    throw new Error('Su usuario no maneja solicitudes.');
  if (yo.rol === 'bodega') todas.forEach(function (s) {
    if (s.estado === SOL.ENV && !s.vistaEn) {
      var ah = new Date(); hojaSol().getRange(s.fila, 16).setValue(ah); s.vistaEn = fmtSello(ah); }
  });
  var out = todas.filter(function (s) {
    if (mias && s.sucursal !== mias) return false;
    var abierta = s.estado === SOL.ENV || s.estado === SOL.ACE || s.estado === SOL.CAM;
    return abierta || s.enviadaEn.slice(0, 10) >= desde;
  }).map(function (s) { return vistaSol(s, lineas); }).reverse();
  return { solicitudes: out, hoy: hoyISO() };
}

function abrirSolicitud(cred, numero) {
  var yo = quien(cred);
  var s = null;
  filasSol().forEach(function (x) { if (x.numero === String(numero)) s = x; });
  if (!s) throw new Error('No se encontró la solicitud ' + numero + '.');
  if (yo.rol === 'gerente' && s.sucursal !== nombreUnidad(yo.sucursal))
    throw new Error('Esa solicitud es de otra sucursal.');
  if (yo.rol !== 'gerente' && yo.rol !== 'bodega' && !veInv(yo))
    throw new Error('Su usuario no maneja solicitudes.');
  if (yo.rol === 'bodega' && s.estado === SOL.ENV && !s.vistaEn) {
    var ahv = new Date(); hojaSol().getRange(s.fila, 16).setValue(ahv); s.vistaEn = fmtSello(ahv); }
  var cat = {};
  productosInv().forEach(function (b) { cat[b.codigo] = b; });
  var exB = yo.rol === 'gerente' ? {} : estadoInventario('bodega').existencia;
  var verPrecio = veCostosTraslado(yo), ptr = {};
  if (verPrecio) productosTraslado().forEach(function (p) { ptr[p.key] = p; });
  var lineasS = lineasSol(s.numero), ctxP = lineasS.some(function (l) { return l.pend > 0; }) ? contextoPend() : null, piden = {};
  if (yo.rol === 'bodega' && (s.estado === SOL.ENV)) {          // lo que piden las otras sucursales del mismo producto (para repartir lo que hay)
    var abiertasO = {}; filasSol().forEach(function (o) { if (o.estado === SOL.ENV && o.numero !== s.numero) abiertasO[o.numero] = o.sucursal; });
    lineasSol().forEach(function (l) { if (abiertasO[l.numero] && l.pedido > 0) (piden[l.clave] = piden[l.clave] || []).push({ sucursal: abiertasO[l.numero], pedido: l.pedido }); });
  }
  var ls = lineasS.filter(function (l) {
    return l.pedido > 0 || l.enviado > 0 || l.sugerido > 0; })
    .map(function (l) {
      var b = cat[l.codigo] || {};
      var x = { clave: l.clave, codigo: l.codigo, nombre: l.nombre, medida: l.medida,
        sugerido: l.sugerido, base: l.base, pedido: l.pedido, enviado: l.enviado,
        recibido: l.recibido, diferencia: l.diferencia, difEstado: l.difEstado,
        resueltaPor: l.resueltaPor, estante: b.estante || '', area: b.area || '',
        grupo: b.grupo || '', suelto: !!(b.suelto && llaveT(b.nombre, b.suelto) === l.clave) };
      if (yo.rol !== 'gerente' && b.codigo) {
        x.enBodega = exB[b.codigo] || 0;
        x.enBodegaTxt = textoCantidad(x.enBodega, b);
      }
      if (piden[l.clave]) x.otrosPiden = piden[l.clave].map(function (o) { return o.sucursal + ' ' + num(o.pedido); }).join(', ');
      if (l.pend > 0 && ctxP) { var estP = estadoPend(l, ctxP); x.pendiente = pendDeLinea(l); x.pendTotal = l.pend; x.pendEntregado = l.pendEntregado; x.pendLlega = l.pendLlega;
        x.pendEstado = estP; x.pendTxt = textoPend(estP, l, ctxP); x.pendMotivo = l.pendMotivo; }
      if (verPrecio) {
        var pp = ptr[l.clave], cant = l.recibido > 0 ? l.recibido : l.enviado > 0 ? l.enviado : l.pedido;
        x.precio = pp ? pp.costo : 0; x.monto = r2((pp ? pp.costo : 0) * (Number(cant) || 0));
      }
      return x;
    }).sort(function (a, b) {
      return (a.estante || '').localeCompare(b.estante || '', 'es') ||
        a.nombre.localeCompare(b.nombre, 'es') || (a.suelto ? 1 : 0) - (b.suelto ? 1 : 0);
    });
  var v = vistaSol(s, lineasSol(s.numero));
  v.lineas = ls;
  if (verPrecio) { v.verPrecios = true; v.total = r2(ls.reduce(function (a, x) { return a + (x.monto || 0); }, 0));
    v.sinPrecio = ls.filter(function (x) { return !(x.precio > 0); }).length; }
  // el administrador ve y da seguimiento; aceptar y despachar es de la bodega,
  // recibir es del gerente de esa sucursal
  v.puedeAceptar = yo.rol === 'bodega' && (s.estado === SOL.ENV || s.estado === SOL.ACE);
  v.puedeDespachar = yo.rol === 'bodega' && s.estado === SOL.ACE;
  v.puedeRecibir = yo.rol === 'gerente' && (s.estado === SOL.ACE || s.estado === SOL.CAM);
  v.puedeAnular = yo.esAdmin && (s.estado === SOL.ENV || s.estado === SOL.ACE ||
    s.estado === SOL.CAM);
  v.puedeBorrar = yo.esAdmin;
  v.puedeAclarar = (yo.rol === 'bodega' || yo.esAdmin) && v.porAclarar > 0;
  return v;
}

/** Samuel acepta: puede mandar menos (o más) de lo pedido. */
function exigeSoloBodega(cred) {
  var yo = quien(cred);
  if (yo.rol !== 'bodega') throw new Error('Eso lo hace la bodega.');
  return yo;
}
function aceptarSolicitud(cred, numero, enviados, opciones) {
  var yo = exigeSoloBodega(cred);
  enviados = enviados || {}; opciones = opciones || {};
  var lock = LockService.getScriptLock();
  lock.waitLock(20000); _LEE = {};
  try {
    var s = null;
    filasSol().forEach(function (x) { if (x.numero === String(numero)) s = x; });
    if (!s) throw new Error('No se encontró la solicitud ' + numero + '.');
    if (s.estado !== SOL.ENV && s.estado !== SOL.ACE)
      throw new Error('La solicitud ' + numero + ' ya está ' + s.estado.toLowerCase() + '.');
    var h = hojaSolD(), total = 0;
    lineasSol(s.numero).forEach(function (l) {
      var crudo = enviados[l.clave];
      var n = crudo == null || String(crudo).trim() === '' ? l.pedido
        : Number(String(crudo).replace(',', '.'));
      if (!isFinite(n) || n < 0) throw new Error('La cantidad de «' + l.nombre + '» no es válida.');
      n = r3(n); total += n;
      h.getRange(l.fila, 11).setValue(n);
      var falta = r3(l.pedido - n), op = opciones[l.clave] || {};          // lo que falta queda pendiente de entrega (salvo que Samuel diga que no se entrega)
      if (falta > 0.0005 && op.pend !== 'no') h.getRange(l.fila, 17, 1, 2).setValues([[falta, String(op.llega || '').trim().slice(0, 40)]]);
      else h.getRange(l.fila, 17, 1, 2).setValues([['', '']]);
    });
    if (!(total > 0)) throw new Error('No va a mandar nada. Si no se puede surtir, pida ' +
      'al administrador que anule la solicitud.');
    hojaSol().getRange(s.fila, 7, 1, 3).setValues([[SOL.ACE, new Date(), yo.nombre]]);
  } finally {
    lock.releaseLock();
  }
  var r = abrirSolicitud(cred, numero);
  r.ok = true;
  var nPend = lineasSol(String(numero)).filter(function (l) { return pendDeLinea(l) > 0.0005; }).length;
  r.mensaje = 'Solicitud ' + numero + ' aceptada. Imprima la hoja para que se la firmen.' + (nPend ? ' ' + nPend + (nPend === 1 ? ' producto queda pendiente de entrega' : ' productos quedan pendientes de entrega') + ': el gerente lo verá.' : '');
  return r;
}

/** El gerente recibe, artículo por artículo. Entra a su inventario, sale de la
 *  bodega y se vuelve costo de la sucursal. */
function recibirSolicitud(cred, numero, recibidos, nota, firma) {
  var yo = quien(cred);
  recibidos = recibidos || {};
  var lock = LockService.getScriptLock();
  lock.waitLock(25000); _LEE = {};
  try {
    var s = null;
    filasSol().forEach(function (x) { if (x.numero === String(numero)) s = x; });
    if (!s) throw new Error('No se encontró la solicitud ' + numero + '.');
    if (yo.rol !== 'gerente' || s.sucursal !== nombreUnidad(yo.sucursal))
      throw new Error('La recibe el gerente de ' + s.sucursal + '.');
    if (s.estado !== SOL.ACE && s.estado !== SOL.CAM) throw new Error(s.estado === SOL.ENV
      ? 'La bodega todavía no ha aceptado la solicitud ' + numero + '.'
      : 'La solicitud ' + numero + ' ya está ' + s.estado.toLowerCase() + '.');
    var suc = null;
    sucursalesDestino().forEach(function (x) { if (x.nombre === s.sucursal) suc = x; });
    var catT = {};
    productosTraslado().forEach(function (p) { catT[p.key] = p; });
    var ahora = new Date(), hoy = hoyISO(), h = hojaSolD();
    var lineas = [], kardex = [], difs = 0;
    lineasSol(s.numero).forEach(function (l) {
      var env = l.enviado || 0;
      var crudo = recibidos[l.clave];
      if (!(env > 0) && (crudo == null || String(crudo).trim() === '')) return;
      if (crudo == null || String(crudo).trim() === '')
        throw new Error('Falta escribir cuánto llegó de «' + l.nombre + '».');
      var n = Number(String(crudo).replace(',', '.'));
      if (!isFinite(n) || n < 0) throw new Error('La cantidad de «' + l.nombre + '» no es válida.');
      n = r3(n);
      var dif = r3(env - n);
      h.getRange(l.fila, 12, 1, 3).setValues([[n, dif, Math.abs(dif) > 0.0005 ? DIF.PEND : '']]);
      if (Math.abs(dif) > 0.0005) difs++;
      if (!(n > 0)) return;
      var p = catT[l.clave];
      var costo = p ? p.costo : 0;
      lineas.push({ k: l.clave, producto: l.nombre, cantidad: n, medida: l.medida,
        precio: costo, subtotal: Math.round(n * costo * 100) / 100,
        categoria: p ? p.categoria : 'Insumos de cocina' });
      var f = p ? p.factor : 1, q = r3(n * f);
      kardex.push([ahora, s.sucursal, l.codigo, l.nombre, MOV.ENT_AB, q, s.numero, yo.nombre, '']);
      kardex.push([ahora, nombreBodega(), l.codigo, l.nombre, MOV.SAL_AB, -q, s.numero,
        yo.nombre, 'a ' + s.sucursal]);
    });
    var tr = '';
    if (lineas.length) {
      tr = siguienteNumero();
      var quienTr = { nombre: s.aceptadaPor || yo.nombre };
      hojaTraslados().appendRow([tr, hoy, Utilities.formatDate(ahora, ZONA, 'HH:mm'),
        quienTr.nombre, ahora]);
      creaTraslado({ s: suc, lineas: lineas }, quienTr, ahora, hoy, tr, s.numero);
      escribeKardex(kardex);
    }
    hojaSol().getRange(s.fila, 7).setValue(difs ? SOL.DIF : SOL.REC);
    hojaSol().getRange(s.fila, 10, 1, 3).setValues([[ahora, yo.nombre, tr]]);
    // la firma de recibido queda en el traslado (así no hay que firmarlo aparte)
    if (tr && firma && /^data:image\/(png|jpeg);base64,/.test(String(firma)) && String(firma).length > 900 && String(firma).length < 48000) {
      var hf = hojaFirmas(), hu = huellaLineas(lineas.map(function (l) { return { producto: l.producto, cantidad: l.cantidad, medida: l.medida }; }));
      var ff = hf.getLastRow() + 1;
      hf.getRange(ff, 1, 1, H_FIRMAS.length).setValues([[tr, s.sucursal, yo.nombre, ahora, String(firma), String(nota || '').slice(0, 300), hu, hoy]]);
      hf.getRange(ff, 8).setNumberFormat('@').setValue(hoy);
    }
    if (!s.llegoEn) hojaSol().getRange(s.fila, 20).setValue(ahora);
    if (nota) hojaSol().getRange(s.fila, 13).setValue(
      (s.nota ? s.nota + ' · ' : '') + 'Al recibir: ' + String(nota).slice(0, 200));
    return { ok: true, traslado: tr, diferencias: difs,
      mensaje: 'Solicitud ' + numero + ' recibida' + (difs ? ' con ' + difs +
        (difs === 1 ? ' diferencia' : ' diferencias') + ': la bodega las va a aclarar.'
        : ', todo completo.') };
  } finally {
    lock.releaseLock();
  }
}

function escribeKardex(filas) {
  if (!filas.length) return;
  var h = hojaInv('Inventario movimientos');
  h.getRange(h.getLastRow() + 1, 1, filas.length, H_KARDEX.length).setValues(filas);
}

/** Lo que llegó de menos: ¿regresó a bodega o se perdió? Lo que llegó de más
 *  (la bodega mandó de más) solo se marca como aclarado. */
function aclararDiferencia(cred, numero, clave, decision) {
  var yo = exigeBodega(cred);
  var l = null;
  lineasSol(String(numero)).forEach(function (x) { if (x.clave === clave) l = x; });
  if (!l) throw new Error('No se encontró ese producto en la solicitud ' + numero + '.');
  if (l.difEstado !== DIF.PEND) throw new Error('Esa diferencia ya se aclaró.');
  var est;
  if (decision === 'merma') {
    if (!(l.diferencia > 0)) throw new Error('Solo lo que llegó de menos puede ser pérdida.');
    var catT = {};
    productosTraslado().forEach(function (p) { catT[p.key] = p; });
    var f = catT[l.clave] ? catT[l.clave].factor : 1;
    escribeKardex([[new Date(), nombreBodega(), l.codigo, l.nombre, MOV.MERMA,
      -r3(l.diferencia * f), numero, yo.nombre, 'No llegó a ' + l.sucursal]]);
    est = DIF.MERMA;
  } else est = l.diferencia > 0 ? DIF.BOD : DIF.OK;
  hojaSolD().getRange(l.fila, 14, 1, 3).setValues([[est, yo.nombre, new Date()]]);
  var r = abrirSolicitud(cred, numero);
  r.ok = true; r.mensaje = l.nombre + ': ' + est.toLowerCase() + '.';
  return r;
}

function anularSolicitud(cred, numero, motivo) {
  var yo = exigeAdmin(cred);
  motivo = String(motivo || '').trim().slice(0, 200);
  if (motivo.length < 4) throw new Error('Escriba el motivo.');
  var s = null;
  filasSol().forEach(function (x) { if (x.numero === String(numero)) s = x; });
  if (!s) throw new Error('No se encontró la solicitud ' + numero + '.');
  if (s.estado !== SOL.ENV && s.estado !== SOL.ACE && s.estado !== SOL.CAM)
    throw new Error('Una solicitud ' + s.estado.toLowerCase() + ' ya no se anula.');
  hojaSol().getRange(s.fila, 7).setValue(SOL.ANU);
  hojaSol().getRange(s.fila, 14, 1, 2).setValues([[yo.nombre, motivo]]);
  return { ok: true, mensaje: 'Solicitud ' + numero + ' anulada. La sucursal puede ' +
    'volver a pedir con su inventario de hoy.' };
}

/** Lo que la bodega ya comprometió: solicitudes enviadas o aceptadas sin recibir. */
function comprometido() {
  var abiertas = {}, out = {};
  filasSol().forEach(function (s) {
    if (s.estado === SOL.ENV || s.estado === SOL.ACE || s.estado === SOL.CAM)
      abiertas[s.numero] = s.estado;
  });
  var catT = {};
  productosTraslado().forEach(function (p) { catT[p.key] = p; });
  lineasSol().forEach(function (l) {
    if (!abiertas[l.numero]) return;
    var n = abiertas[l.numero] !== SOL.ENV && l.enviado != null ? l.enviado : l.pedido;
    var f = catT[l.clave] ? catT[l.clave].factor : 1;
    out[l.codigo] = (out[l.codigo] || 0) + n * f;
  });
  return out;
}

/* ── compras de la bodega ── */

function hojaOC() { return hojaInv('Órdenes de compra'); }
function filasOC() {
  return leeTodo(hojaOC(), H_OC.length).map(function (r, i) {
    return { fila: i + 2, numero: String(r[0] || ''), creadaEn: fmtSello(r[1]),
      creadaPor: String(r[2] || ''), proveedor: String(r[3] || ''),
      compra: String(r[4] || ''), presentacion: String(r[5] || ''),
      sugerido: Number(r[6]) || 0, pedido: Number(r[7]) || 0, codigo: String(r[8] || ''),
      nombre: String(r[9] || ''), rinde: Number(r[10]) || 1, estado: String(r[11] || ''),
      recibido: r[12] === '' || r[12] == null ? null : Number(r[12]),
      recibidoEn: fmtSello(r[13]), recibidoPor: String(r[14] || ''),
      precio: Number(r[15]) || 0, subtotal: Number(r[16]) || 0 };
  }).filter(function (x) { return x.numero; });
}

/** Lo que se le puede comprar a cada proveedor. Samuel escoge el proveedor y ve sus
 *  productos (los que él ancló). Mientras no haya dos semanas de despachos, él decide
 *  cuánto pedir; con historia, el sistema sugiere. */
function sugerenciaCompra(cred, proveedor) {
  var yo = exigeBodega(cred);
  var est = estadoInventario('bodega'), aj = ajustesInv(), hoy = hoyISO();
  var cons = consumoPorDia(est, Date.now()), comp = comprometido();
  var enCamino = {};
  filasOC().forEach(function (o) {
    if (o.estado === 'Pedida') enCamino[o.codigo] = (enCamino[o.codigo] || 0) + o.pedido * o.rinde;
  });
  var idx = indiceCompras(), dias = aj.cobertura + aj.colchon, cuenta = {}, sinProv = 0;
  var todos = productosInv();
  todos.forEach(function (b) {
    if (b.proveedor) cuenta[b.proveedor] = (cuenta[b.proveedor] || 0) + 1; else sinProv++; });
  var noBodega = {};
  filasProveedores().forEach(function (p) { if (!esDeBodega(p)) noBodega[p.nombre] = true; });
  var provs = filasProveedores().filter(function (p) { return p.activo && esDeBodega(p); }).map(function (p) {
    return { nombre: p.nombre, tel: p.tel, productos: cuenta[p.nombre] || 0 }; });
  Object.keys(cuenta).forEach(function (n) {           // anclados a uno que no está en la lista
    if (noBodega[n]) return;
    if (!provs.some(function (p) { return p.nombre === n; }))
      provs.push({ nombre: n, tel: '', productos: cuenta[n] }); });
  provs.sort(function (a, b) { return (b.productos > 0) - (a.productos > 0) ||
    a.nombre.localeCompare(b.nombre, 'es'); });
  var lineas = [];
  if (proveedor) todos.forEach(function (b) {
    if (b.proveedor !== proveedor) return;
    var q = est.existencia[b.codigo] || 0, c = comp[b.codigo] || 0, llega = enCamino[b.codigo] || 0;
    var ligado = b.compra ? (idx[(b.compra + '§' + b.proveedor).toLowerCase()] ||
      idx[b.compra.toLowerCase()]) : null;
    var pres = b.presCompra || (ligado ? ligado.medida : b.unidad);
    var rinde = b.rinde > 0 ? b.rinde : 1, sug = null, base = '';
    if (cons._conHistoria && cons[b.codigo] != null) {
      var falta = cons[b.codigo] * dias + c - q - llega;
      sug = falta > 0.0005 ? Math.ceil(falta / rinde - 1e-6) : 0;
      base = 'sugerido con lo que se despacha: ' + num(Math.round(cons[b.codigo] * 100) / 100) + ' ' +
        plur(2, b.unidad) + '/día';
    } else base = b.ideal.bodega > 0 ? 'nivel ideal de bodega: ' + textoCantidad(b.ideal.bodega, b) : '';
    var x = { codigo: b.codigo, nombre: b.nombre, estante: b.estante, unidad: b.unidad,
      existencia: q, existenciaTxt: textoCantidad(q, b), comprometido: r3(c),
      comprometidoTxt: c > 0 ? textoCantidad(c, b) : '', enCamino: llega > 0 ? textoCantidad(llega, b) : '',
      presentacion: pres, rinde: rinde, sugerido: sug, base: base, compra: ligado ? ligado.nombre : '' };
    if (yo.esAdmin && ligado && ligado.precio > 0 && sug > 0)
      x.costoEstimado = Math.round(sug * ligado.precio * 100) / 100;
    lineas.push(x);
  });
  lineas.sort(function (a, b) { return (b.comprometido > 0) - (a.comprometido > 0) ||
    a.nombre.localeCompare(b.nombre, 'es'); });
  var ult = est.ultimo, fresco = !!ult && diasEntre(ult.fecha, hoy) <= 3;
  return { hoy: hoy, proveedor: proveedor || '', proveedores: provs, sinProveedor: sinProv,
    lineas: lineas, conHistoria: !!cons._conHistoria,
    ultimo: ult ? { sello: ult.sello, quien: ult.quien, fecha: ult.fecha } : null,
    puedeGuardar: fresco,
    aviso: fresco ? '' : 'Haga el inventario de bodega antes de pedir: la orden sale de lo que ' +
      'hay en bodega (hace falta uno de los últimos 3 días).' };
}

/** Una orden para un proveedor. d = { proveedor, lineas: [{codigo, pedido}] } */
function guardarOrden(cred, d) {
  var yo = exigeBodega(cred);
  d = d || {};
  var prov = String(d.proveedor || '').trim();
  if (!prov) throw new Error('Escoja el proveedor.');
  var sug = sugerenciaCompra(cred, prov);
  if (!sug.puedeGuardar) throw new Error(sug.aviso);
  var ped = {};
  (d.lineas || []).forEach(function (l) { ped[l.codigo] = l.pedido; });
  var lock = LockService.getScriptLock();
  lock.waitLock(20000); _LEE = {};
  try {
    var ahora = new Date(), filas = [];
    var numero = 'OC-' + ('0000' + (siguienteOC() + 1)).slice(-4);
    sug.lineas.forEach(function (l) {
      if (ped[l.codigo] == null || String(ped[l.codigo]).trim() === '') return;
      var q = r3(Number(String(ped[l.codigo]).replace(',', '.')));
      if (!isFinite(q) || q < 0) throw new Error('La cantidad de «' + l.nombre + '» no es válida.');
      if (!(q > 0)) return;
      filas.push([numero, ahora, yo.nombre, prov, l.compra || l.nombre, l.presentacion,
        l.sugerido == null ? '' : l.sugerido, q, l.codigo, l.nombre, l.rinde, 'Pedida', '', '', '', '', '']);
    });
    if (!filas.length) throw new Error('Escriba al menos una cantidad.');
    var h = hojaOC();
    h.getRange(h.getLastRow() + 1, 1, filas.length, H_OC.length).setValues(filas);
    return { ok: true, numero: numero, mensaje: 'Orden ' + numero + ' para ' + prov + ': ' +
      filas.length + (filas.length === 1 ? ' producto.' : ' productos.'),
      ordenes: listaOrdenes(cred).ordenes };
  } finally {
    lock.releaseLock();
  }
}

function siguienteOC() {
  var max = 0;
  leeTodo(hojaOC(), 1).forEach(function (r) {
    var m = /^OC-(\d+)$/.exec(String(r[0]).trim());
    if (m) max = Math.max(max, Number(m[1]));
  });
  return max;
}

/** Los montos de una orden los ven solo la bodega (Samuel) y contabilidad. */
function veMontosOC(yo) { return yo.rol === 'bodega' || yo.esAdmin || esContador(yo) || yo.rol === 'finanzas'; }
function listaOrdenes(cred) {
  var yo = exigeVerBodega(cred), montos = veMontosOC(yo);
  var desde = masDias(hoyISO(), -60), por = {}, orden = [], pagos = filasOCP();
  var ultimo = {}, idx = montos ? indiceCompras() : null;   // último precio pagado/cotizado por proveedor y producto
  filasOC().forEach(function (o) {
    var llave = (o.proveedor + '§' + o.codigo).toLowerCase(), anterior = 0;
    if (montos) {
      anterior = ultimo[llave] || 0;
      if (!anterior) {                                    // sin orden anterior: el precio del catálogo de compras
        var lig = o.compra ? (idx[(o.compra + '§' + o.proveedor).toLowerCase()] || idx[o.compra.toLowerCase()]) : null;
        if (lig && lig.precio > 0) anterior = lig.precio;
      }
      if (o.estado !== 'Anulada' && o.precio > 0) ultimo[llave] = o.precio;
    }
    if (o.estado !== 'Pedida' && o.creadaEn.slice(0, 10) < desde) return;
    if (!por[o.numero]) {
      var pg = pagos[o.numero];
      por[o.numero] = { numero: o.numero, proveedor: o.proveedor, creadaEn: o.creadaEn,
        creadaPor: o.creadaPor, estado: o.estado, recibidoEn: o.recibidoEn,
        recibidoPor: o.recibidoPor, lineas: [],
        pago: pg && pg.pagadaEn ? 'Pagada' : pg && pg.cotizadaEn ? 'Cotizada' : 'Por cotizar',
        pagadaPor: pg ? pg.pagadaPor : '', totalCotizado: montos && pg ? pg.total : 0,
        verMontos: montos, transporte: montos && pg ? pg.transporte : 0,
        puedeCotizar: montos && (yo.rol === 'bodega' || yo.esAdmin) && !(pg && pg.pagadaEn) && o.estado !== 'Anulada' && !(pg && pg.devueltaEn),
        archivo: pg ? pg.archivo : '',
        devuelta: pg && pg.devueltaEn ? { en: pg.devueltaEn, por: pg.devueltaPor, motivo: pg.motivoDev } : null,
        puedeCorregir: (yo.rol === 'bodega' || yo.esAdmin) && !!(pg && pg.devueltaEn) && !(pg && pg.pagadaEn) && o.estado === 'Pedida' };
      orden.push(por[o.numero]);
    }
    var ln = { codigo: o.codigo, nombre: o.nombre, compra: o.compra,
      presentacion: o.presentacion, sugerido: o.sugerido, pedido: o.pedido,
      recibido: o.recibido, rinde: o.rinde, unidad: (unidadDeCodigo(o.codigo) || {}).unidad || 'unidad' };
    if (montos) { ln.anterior = anterior; ln.precio = o.precio; ln.subtotal = o.subtotal; }
    por[o.numero].lineas.push(ln);
  });
  // cómo va avanzando cada orden: pedida → cotizada → pagada → recibida (igual que en contabilidad)
  var hoy = hoyISO(), res = { porCotizar: 0, porPagar: 0, porRecibir: 0, completas: 0, anuladas: 0 };
  orden.forEach(function (x) {
    var pg = pagos[x.numero] || {}, anul = x.estado === 'Anulada', rec = x.estado === 'Recibida';
    var cot = !!pg.cotizadaEn, pag = !!pg.pagadaEn;
    x.paso = pasoOC({ estado: x.estado, pago: x.pago, devueltaEn: x.devuelta ? x.devuelta.en : '' });
    x.diasAbierta = diasEntre(String(x.creadaEn).slice(0, 10), hoy);
    x.avance = [
      { n: 'Pedida', hecho: true, quien: x.creadaPor, en: x.creadaEn,
        detalle: x.lineas.length + (x.lineas.length === 1 ? ' producto' : ' productos') },
      { n: 'Cotizada', hecho: cot, quien: cot ? pg.cotizadaPor : '', en: cot ? pg.cotizadaEn : '',
        detalle: cot ? (montos ? 'Total ' + dinero(pg.total) : (pg.cotizacion ? 'Cotización ' + pg.cotizacion : '')) : 'Falta la cotización del proveedor' },
      { n: 'Pagada', hecho: pag, quien: pag ? pg.pagadaPor : '', en: pag ? pg.pagadaEn : '',
        detalle: pag ? (pg.forma || '') + (pg.referencia ? ' ' + pg.referencia : '') + (montos && pg.monto ? ' · ' + dinero(pg.monto) : '') : (cot ? 'La paga el contador' : '') },
      { n: 'Recibida', hecho: rec, quien: rec ? x.recibidoPor : '', en: rec ? x.recibidoEn : '',
        detalle: rec ? (x.lineas.some(function (l) { return l.recibido != null && l.recibido < l.pedido - 0.0005; }) ? 'Llegó con faltantes' : 'Completa') : (pag ? 'Ya está pagada: falta que llegue' : '') }
    ];
    if (anul) res.anuladas++;
    else if (!cot) res.porCotizar++;
    else if (!pag) res.porPagar++;
    else if (!rec) res.porRecibir++;
    else res.completas++;
  });
  return { ordenes: orden.reverse(), resumen: res };
}

/** La orden de compra en PDF. Sin montos es la que se le manda al proveedor; con montos
 *  (precio de cada presentación, subtotal, transporte y total) es solo para Samuel y contabilidad. */
function pdfOrden(cred, numero, conMontos) {
  var yo = quien(cred);
  if (yo.rol !== 'bodega' && !veInv(yo) && !veMontosOC(yo)) throw new Error('Su usuario no ve las órdenes de compra.');
  conMontos = conMontos === true || conMontos === 'true';
  if (conMontos && !veMontosOC(yo)) throw new Error('Los montos de la orden los ven solo la bodega y contabilidad.');
  var o = null;
  ordenesAgrupadas().forEach(function (x) { if (x.numero === String(numero)) o = x; });
  if (!o) throw new Error('No se encontró la orden ' + numero + '.');
  var nombre = o.numero + ' ' + o.proveedor.replace(/[\\\/:*?"<>|]/g, '') + (conMontos ? ' (con montos)' : '') + '.pdf';
  var blob = Utilities.newBlob(htmlOrden(o, conMontos), 'text/html', 'o.html').getAs('application/pdf').setName(nombre);
  return { ok: true, nombre: nombre, tipo: 'application/pdf', base64: Utilities.base64Encode(blob.getBytes()) };
}
function htmlOrden(o, conMontos) {
  var h = [], sub = 0;
  h.push('<html><head><meta charset="utf-8"><style>');
  h.push('body{font-family:Arial,Helvetica,sans-serif;font-size:11px;color:#16223A;margin:26px}');
  h.push('table.band{width:100%;background:#14284B;color:#fff;border-bottom:4px solid #F2B705;margin:0 0 12px;border-collapse:collapse}');
  h.push('table.band td{border:0;padding:10px 12px;vertical-align:middle}.band img{width:44px}');
  h.push('.band td.r{text-align:right;font-size:13px;color:#F7F2E7;font-weight:bold}h1{font-size:20px;color:#fff;margin:0;letter-spacing:1px}');
  h.push('.band .co{color:#F2B705;font-size:9px;letter-spacing:2px;font-weight:bold}');
  h.push('table.d{width:100%;border-collapse:collapse;margin-bottom:10px}table.d td{padding:2px 0;font-size:11px}');
  h.push('table.l{width:100%;border-collapse:collapse}table.l th{background:#14284B;color:#F2B705;padding:5px;font-size:10px;text-align:left}');
  h.push('table.l td{border-bottom:1px solid #E3DBCB;padding:5px;font-size:11px}td.n,th.n{text-align:right}');
  h.push('table.t{margin-top:10px;margin-left:auto;border-collapse:collapse}table.t td{padding:3px 8px;font-size:12px}table.t tr.g td{font-weight:bold;border-top:2px solid #14284B;font-size:14px}');
  h.push('.firmas{margin-top:46px}.firmas td{width:50%;text-align:center;border-top:1px solid #16223A;padding-top:4px;font-size:10px;color:#44506A}');
  h.push('.aviso{margin-top:14px;font-size:9.5px;color:#6F7686}');
  h.push('</style></head><body>');
  h.push('<table class="band"><tr><td style="width:56px"><img src="' + LOGO_PNG + '"></td><td><h1>REY PIZZA</h1><div class="co">CAIX, S.A. · ' + htm(nombreBodega()) + '</div></td>' +
    '<td class="r">Orden de compra<br>' + htm(o.numero) + '</td></tr></table>');
  h.push('<table class="d"><tr><td><b>Proveedor:</b> ' + htm(o.proveedor) + '</td><td><b>Fecha:</b> ' + htm(dia(o.creadaEn.slice(0, 10)) + ' ' + o.creadaEn.slice(11, 16)) + '</td></tr>' +
    '<tr><td><b>Pidió:</b> ' + htm(o.creadaPor) + '</td><td><b>Estado:</b> ' + htm(o.estado) + '</td></tr></table>');
  h.push('<table class="l"><tr><th>#</th><th>Producto</th><th>Presentación</th><th class="n">Cantidad</th>' +
    (conMontos ? '<th class="n">Precio</th><th class="n">Subtotal</th>' : '') + '</tr>');
  o.lineas.forEach(function (l, i) {
    var st = l.precio > 0 ? r2(l.precio * l.pedido) : 0; sub += st;
    h.push('<tr><td>' + (i + 1) + '</td><td>' + htm(l.compra || l.nombre) + '</td><td>' + htm(l.presentacion) + '</td><td class="n">' + l.pedido + '</td>' +
      (conMontos ? '<td class="n">' + (l.precio > 0 ? dinero(l.precio) : '—') + '</td><td class="n">' + (st > 0 ? dinero(st) : '—') + '</td>' : '') + '</tr>');
  });
  h.push('</table>');
  if (conMontos) {
    h.push('<table class="t"><tr><td>Subtotal</td><td class="n">' + dinero(sub) + '</td></tr>');
    if (o.transporte > 0) h.push('<tr><td>Costo de transporte</td><td class="n">' + dinero(o.transporte) + '</td></tr>');
    h.push('<tr class="g"><td>Total</td><td class="n">' + dinero(r2(sub + (o.transporte || 0))) + '</td></tr></table>');
    h.push('<div class="aviso">Documento interno con montos: solo bodega y contabilidad. Para el proveedor use el PDF sin montos.</div>');
  } else {
    h.push('<div class="aviso">Favor enviar su cotización con el precio de cada presentación' + '. Gracias.</div>');
  }
  h.push('<table class="firmas" style="width:100%;border-collapse:separate;border-spacing:24px 0"><tr><td>Pidió · firma</td><td>Recibió en bodega · firma y fecha</td></tr></table>');
  h.push('</body></html>');
  return h.join('');
}

/** La orden llegó: lo recibido entra a la bodega (en su unidad de traslado). */
function recibirOrden(cred, numero, recibidos) {
  var yo = exigeBodega(cred);
  recibidos = recibidos || {};
  var lock = LockService.getScriptLock();
  lock.waitLock(20000); _LEE = {};
  try {
    var ls = filasOC().filter(function (o) { return o.numero === String(numero); });
    if (!ls.length) throw new Error('No se encontró la orden ' + numero + '.');
    if (ls[0].estado !== 'Pedida') throw new Error('La orden ' + numero + ' ya está ' +
      ls[0].estado.toLowerCase() + '.');
    var ahora = new Date(), h = hojaOC(), kardex = [];
    ls.forEach(function (o) {
      var crudo = recibidos[o.codigo];
      var n = crudo == null || String(crudo).trim() === '' ? o.pedido
        : Number(String(crudo).replace(',', '.'));
      if (!isFinite(n) || n < 0) throw new Error('La cantidad de «' + o.nombre + '» no es válida.');
      n = r3(n);
      h.getRange(o.fila, 12, 1, 4).setValues([['Recibida', n, ahora, yo.nombre]]);
      if (n > 0) kardex.push([ahora, nombreBodega(), o.codigo, o.nombre, MOV.COMPRA,
        r3(n * o.rinde), numero, yo.nombre, n + ' ' + o.presentacion + ' de ' + o.proveedor]);
    });
    escribeKardex(kardex);
    return { ok: true, mensaje: 'Orden ' + numero + ' recibida: entró a bodega.',
      ordenes: listaOrdenes(cred).ordenes };
  } finally {
    lock.releaseLock();
  }
}

function anularOrden(cred, numero) {
  exigeBodega(cred);
  var ls = filasOC().filter(function (o) { return o.numero === String(numero); });
  if (!ls.length) throw new Error('No se encontró la orden ' + numero + '.');
  if (ls[0].estado !== 'Pedida') throw new Error('Esa orden ya no se puede anular.');
  ls.forEach(function (o) { hojaOC().getRange(o.fila, 12).setValue('Anulada'); });
  return { ok: true, mensaje: 'Orden ' + numero + ' anulada.', ordenes: listaOrdenes(cred).ordenes };
}

/* ── para el administrador: cómo va cada unidad ── */

function resumenInventario(cred) {
  var yo = quien(cred);
  if (!veInv(yo) && yo.rol !== 'bodega') throw new Error('Solo un administrador.');
  var aj = ajustesInv(), hoy = hoyISO(), sols = filasSol(), lineas = lineasSol();
  var unidades = UNIDADES.filter(function (u) { return u.conInv; }).map(function (u) {
    var est = estadoInventario(u.id);
    var ult = est.ultimo;
    return { id: u.id, nombre: u.nombre, marca: u.marca,
      ultimo: ult ? { sello: ult.sello, quien: ult.quien, fecha: ult.fecha } : null,
      conteos30: est.conteos.filter(function (c) {
        return c.fecha >= masDias(hoy, -30); }).length,
      calendario: calendarioConteo(u.id, ult ? ult.fecha : '', aj, hoy),
      extraHoy: u.id !== 'bodega' && extraVigente(ajusteCrudo('conteo_extra_' + u.id), ult, hoy) };
  });
  var abiertas = sols.filter(function (s) {
    return s.estado === SOL.ENV || s.estado === SOL.ACE; }).length;
  var dif = lineas.filter(function (l) { return l.difEstado === DIF.PEND; }).length;
  return { unidades: unidades, abiertas: abiertas, porAclarar: dif,
    ajustes: { dias_conteo_sucursales: nombresDias(aj.conteoSuc).join(', '),
      dias_conteo_bodega: nombresDias(aj.conteoBod).join(', '),
      dias_entrega: nombresDias(aj.entrega).join(', '),
      colchon_dias: aj.colchon, cobertura_compra_dias: aj.cobertura } };
}

/** Lo que hay en una unidad, producto por producto. */
function existencias(cred, unidadId) {
  var yo = quien(cred);
  if (!veInv(yo) && !((yo.rol === 'bodega' || yo.rol === 'produccion') && unidadId === 'bodega'))
    throw new Error('Solo un administrador.');
  var est = estadoInventario(unidadId), cons = consumoPorDia(est, Date.now());
  var comp = unidadId === 'bodega' ? comprometido() : {};
  return { unidad: { id: est.unidad.id, nombre: est.unidad.nombre },
    ultimo: est.ultimo ? { sello: est.ultimo.sello, quien: est.ultimo.quien } : null,
    productos: productosInv().filter(function (b) { return maneja(b, unidadId); })
      .map(function (b) {
        var q = est.existencia[b.codigo] || 0, c = cons[b.codigo];
        return { codigo: b.codigo, nombre: b.nombre, area: b.area, grupo: b.grupo,
          unidad: b.unidad, existencia: q, existenciaTxt: textoCantidad(q, b),
          ideal: b.ideal[unidadId] || 0,
          consumo: c != null ? Math.round(c * 100) / 100 : null,
          alcanza: c > 0 ? Math.floor(q / c * 10) / 10 : null,
          comprometido: r3(comp[b.codigo] || 0) };
      }) };
}

function listaDiferencias(cred) {
  var yoD = exigeVerBodega(cred);
  var sols = {};
  filasSol().forEach(function (s) { sols[s.numero] = s; });
  return { diferencias: lineasSol().filter(function (l) { return l.difEstado === DIF.PEND; })
    .map(function (l) {
      var s = sols[l.numero] || {};
      return { numero: l.numero, sucursal: l.sucursal, clave: l.clave, nombre: l.nombre,
        medida: l.medida, enviado: l.enviado, recibido: l.recibido, diferencia: l.diferencia,
        recibidaEn: s.recibidaEn || '', recibidaPor: s.recibidaPor || '' };
    }) };
}

function guardarAjustesInv(cred, a) {
  exigeAdmin(cred);
  a = a || {};
  var claves = ['dias_conteo_sucursales', 'dias_conteo_bodega', 'dias_entrega',
    'colchon_dias', 'cobertura_compra_dias'];
  var h = hojaInv('Ajustes de inventario');
  var filas = leeTodo(h, 1).map(function (r) { return String(r[0]).trim(); });
  claves.forEach(function (k) {
    if (a[k] == null) return;
    var v = String(a[k]).trim().slice(0, 80);
    if (/^dias_/.test(k) && k !== 'colchon_dias' && k !== 'cobertura_compra_dias') {
      var ok = v.toLowerCase().split(/[,;\s]+/).filter(function (w) {
        return DIAS_SEM.indexOf(w.replace('miercoles', 'miércoles')
          .replace('sabado', 'sábado')) >= 0; });
      if (!ok.length) throw new Error('Escoja al menos un día para «' + k.replace(/_/g, ' ') + '».');
    } else if (!(Number(v) >= 0)) throw new Error('Revise el número de «' + k.replace(/_/g, ' ') + '».');
    var i = filas.indexOf(k);
    if (i >= 0) h.getRange(i + 2, 2).setValue(v);
    else {
      var ini = AJUSTES_INICIALES.filter(function (x) { return x[0] === k; })[0];
      h.appendRow([k, v, ini ? ini[2] : '']);
      filas.push(k);
    }
  });
  _AJ = null;
  try { CacheService.getScriptCache().remove('ajustes_inv'); } catch (e) {}
  var r = resumenInventario(cred);
  r.ok = true; r.mensaje = 'Calendario guardado.';
  return r;
}

/** Lo que cuenta el tablero de cada rol al entrar: sin leer todo el inventario. */
function avisosInventario(yo) {
  try {
    var aj = ajustesInv(), hoy = hoyISO(), out = {};
    var unidadId = yo.rol === 'gerente' ? yo.sucursal : yo.rol === 'bodega' ? 'bodega' : '';
    if (unidadId) {
      var est = estadoInventario(unidadId);
      out.calendario = calendarioConteo(unidadId, est.ultimo ? est.ultimo.fecha : '', aj, hoy);
      out.ultimo = est.ultimo ? est.ultimo.sello : '';
    }
    var sols = filasSol();
    if (yo.rol === 'gerente') {
      var n = nombreUnidad(yo.sucursal);
      out.porRecibir = sols.filter(function (s) {
        return s.sucursal === n && (s.estado === SOL.ACE || s.estado === SOL.CAM); }).length;
    } else {
      out.nuevas = sols.filter(function (s) { return s.estado === SOL.ENV; }).length;
      out.aceptadas = sols.filter(function (s) { return s.estado === SOL.ACE; }).length;
      out.porAclarar = lineasSol().filter(function (l) {
        return l.difEstado === DIF.PEND; }).length;
    }
    return out;
  } catch (e) {
    return { error: e.message };
  }
}

/** Al anular un traslado, lo que entró a la sucursal regresa a la bodega. */
function kardexDeshaceTraslado(idS, yo, motivo) {
  var g = libro().getSheetByName('Gastos');
  var suc = '', numero = '';
  leeTodo(g, 16).forEach(function (r) {
    if (String(r[0]) === idS) { suc = String(r[4] || ''); numero = String(r[15] || ''); }
  });
  if (!suc) return;
  var catT = {};
  productosTraslado().forEach(function (p) { catT[p.key] = p; });
  var ahora = new Date(), kx = [];
  leeTodo(libro().getSheetByName('Gastos detalle'), H_DETALLE.length).forEach(function (r) {
    if (String(r[0]) !== idS) return;
    var est = String(r[13]).toLowerCase();
    if (est === 'sí' || est === 'reemplazado') return;
    var p = catT[llaveT(r[8], r[10])];
    if (!p) return;
    var q = r3((Number(r[9]) || 0) * p.factor);
    if (!q) return;
    kx.push([ahora, suc, p.codigo, p.nombre, MOV.ENT_TR, -q, numero, yo.nombre,
      'traslado anulado: ' + motivo]);
    kx.push([ahora, nombreBodega(), p.codigo, p.nombre, MOV.SAL_TR, q, numero, yo.nombre,
      'traslado anulado: ' + motivo]);
  });
  escribeKardex(kx);
}


/* ════════════ SEGUIMIENTO DE SOLICITUDES, AVISOS Y REPORTES DE TRASLADOS ════════════ */

function solPorNumero(numero) {
  var s = null;
  filasSol().forEach(function (x) { if (x.numero === String(numero)) s = x; });
  if (!s) throw new Error('No se encontró la solicitud ' + numero + '.');
  return s;
}

/** Al imprimir la hoja: quedan «preparadas» (sin que Samuel apriete nada más). */
function marcarPreparadas(cred, numeros) {
  exigeSoloBodega(cred);
  var ahora = new Date(), h = hojaSol();
  filasSol().forEach(function (s) {
    if ((numeros || []).indexOf(s.numero) < 0) return;
    if (s.estado === SOL.ACE && !s.preparadaEn) h.getRange(s.fila, 17).setValue(ahora);
  });
  return { ok: true };
}

/** Salió el camión: todas las aceptadas que se escojan pasan a «en camino». */
function despacharSolicitudes(cred, numeros) {
  var yo = exigeSoloBodega(cred);
  var ahora = new Date(), h = hojaSol(), hechas = [];
  var lock = LockService.getScriptLock();
  lock.waitLock(20000); _LEE = {};
  try {
    filasSol().forEach(function (s) {
      if ((numeros || []).indexOf(s.numero) < 0 || s.estado !== SOL.ACE) return;
      h.getRange(s.fila, 7).setValue(SOL.CAM);
      if (!s.preparadaEn) h.getRange(s.fila, 17).setValue(ahora);
      h.getRange(s.fila, 18, 1, 2).setValues([[ahora, yo.nombre]]);
      hechas.push(s.numero + ' (' + s.sucursal + ')');
    });
  } finally {
    lock.releaseLock();
  }
  if (!hechas.length) throw new Error('No hay solicitudes aceptadas para despachar.');
  return { ok: true, mensaje: 'En camino: ' + hechas.join(', ') + '.' };
}

/** El gerente abre «Recibir»: el pedido ya se descargó en la sucursal. */
function llegoSolicitud(cred, numero) {
  var yo = quien(cred), s = solPorNumero(numero);
  if (yo.rol !== 'gerente' || s.sucursal !== nombreUnidad(yo.sucursal)) return { ok: false };
  if ((s.estado === SOL.CAM || s.estado === SOL.ACE) && !s.llegoEn)
    hojaSol().getRange(s.fila, 20).setValue(new Date());
  return { ok: true };
}

/** Borra una solicitud por completo (para pruebas). Si ya se había recibido, anula su
 *  traslado y quita sus movimientos de inventario. */
function borrarSolicitud(cred, numero) {
  var yo = exigeAdmin(cred);
  var r = borraSol(String(numero), yo.nombre);
  r.ok = true;
  return r;
}
function borraSol(numero, quien) {
  var s = solPorNumero(numero), hechos = [];
  var lock = LockService.getScriptLock();
  lock.waitLock(25000); _LEE = {};
  try {
    if (s.traslado) {                                // su traslado queda anulado
      var g = libro().getSheetByName('Gastos'), ids = [];
      leeTodo(g, 16).forEach(function (r, i) {
        if (String(r[15]) === s.traslado && String(r[10]).toLowerCase() !== 'sí') {
          g.getRange(i + 2, 11, 1, 4).setValues([['Sí', 'Solicitud ' + numero + ' borrada', quien,
            new Date()]]);
          ids.push(String(r[0]));
        }
      });
      ids.forEach(function (id) { anulaDetalle(id); });
      if (ids.length) hechos.push('traslado ' + s.traslado + ' anulado');
    }
    var hk = hojaInv('Inventario movimientos'), filas = [];
    leeTodo(hk, 7).forEach(function (r, i) { if (String(r[6]) === numero) filas.push(i + 2); });
    filas.reverse().forEach(function (f) { hk.deleteRow(f); });
    if (filas.length) hechos.push(filas.length + ' movimientos de inventario quitados');
    var hd = hojaSolD(), fd = [];
    leeTodo(hd, 1).forEach(function (r, i) { if (String(r[0]) === numero) fd.push(i + 2); });
    fd.reverse().forEach(function (f) { hd.deleteRow(f); });
    hojaSol().deleteRow(s.fila);
  } finally {
    lock.releaseLock();
  }
  return { mensaje: 'Solicitud ' + numero + ' de ' + s.sucursal + ' borrada' +
    (hechos.length ? ' (' + hechos.join(', ') + ')' : '') + '.' };
}
/** La hoja consolidada: lo de todas las solicitudes escogidas, producto por producto. */
function consolidado(cred, numeros) {
  var yo = exigeSoloBodega(cred);
  numeros = numeros || [];
  var sols = filasSol().filter(function (s) { return numeros.indexOf(s.numero) >= 0; });
  if (!sols.length) throw new Error('Escoja las solicitudes.');
  var cat = {};
  productosInv().forEach(function (b) { cat[b.codigo] = b; });
  var suc = {}, por = {};
  sols.forEach(function (s) { suc[s.sucursal] = s.numero; });
  lineasSol().forEach(function (l) {
    if (numeros.indexOf(l.numero) < 0) return;
    var n = l.enviado != null ? l.enviado : l.pedido;
    if (!(n > 0)) return;
    if (!por[l.clave]) {
      var b = cat[l.codigo] || {};
      por[l.clave] = { nombre: l.nombre, medida: l.medida, estante: b.estante || '',
        suelto: !!(b.suelto && llaveT(b.nombre, b.suelto) === l.clave), c: {}, total: 0 };
    }
    por[l.clave].c[l.sucursal] = (por[l.clave].c[l.sucursal] || 0) + n;
    por[l.clave].total = r3(por[l.clave].total + n);
  });
  marcarPreparadas(cred, numeros);
  var lineas = Object.keys(por).map(function (k) { return por[k]; }).sort(function (a, b) {
    return a.estante.localeCompare(b.estante, 'es') || a.nombre.localeCompare(b.nombre, 'es'); });
  return { sucursales: sucursalesDestino().map(function (u) { return u.nombre; })
      .filter(function (n) { return suc[n]; }),
    solicitudes: sols.map(function (s) { return { numero: s.numero, sucursal: s.sucursal,
      entrega: s.entrega }; }), lineas: lineas, preparo: yo.nombre };
}

/* ── la campanita: avisos de la bodega y de los gerentes ──
 * Cada aviso tiene un id fijo (el teléfono recuerda cuáles ya se vieron). */
/** La campanita. Cuando el teléfono pregunta solo (cada 90 s) se usa lo calculado
 *  hace menos de un minuto; después de cualquier acción se calcula de nuevo. */
function notificaciones(cred, deCache) {
  permiteAux();
  var yo = quien(cred), k = 'notif_' + yo.nombre.toLowerCase(), cc = null;
  try { cc = CacheService.getScriptCache(); } catch (e) {}
  if (deCache && cc) { try { var t = cc.get(k); if (t) return JSON.parse(t); } catch (e) {} }
  var r = notificacionesCalc(cred);
  if (cc) { try { var s = JSON.stringify(r); if (s.length < 90000) cc.put(k, s, 60); } catch (e) {} }
  return r;
}
function notificacionesCalc(cred) {
  var yo = quien(cred), out = [], hoy = hoyISO(), aj = ajustesInv();
  var desde = masDias(hoy, -10);
  try { avisosEquipos(yo).forEach(function (a) { out.push(a); }); } catch (e) {}
  try { avisosPendientes(yo).forEach(function (a) { out.push(a); }); } catch (e) {}
  try { avisosCierre(yo).forEach(function (a) { out.push(a); }); } catch (e) {}
  try { avisosPrecioProducto(yo, out, desde); } catch (e) {}
  if (yo.rol === 'bodega' || yo.esAdmin) {          // órdenes de compra devueltas por contabilidad: se corrigen y se reenvían
    try { var pgD = filasOCP(); Object.keys(pgD).forEach(function (n) { var p = pgD[n];
      if (p.devueltaEn && !p.pagadaEn) out.push({ id: 'ocdev-' + n + '-' + p.devueltaEn, tipo: 'Compra devuelta', cuando: p.devueltaEn,
        titulo: n + ' devuelta por ' + (p.devueltaPor || 'contabilidad'), texto: p.motivoDev, abrir: { tab: 'compras' } }); }); } catch (e) {}
  }
  if (yo.rol === 'bodega') {
    filasSol().forEach(function (s) {
      if (s.estado === SOL.ENV)
        out.push({ id: 'nueva-' + s.numero, tipo: 'Solicitud nueva', cuando: s.enviadaEn,
          titulo: s.sucursal + ' pidió (' + s.numero + ')',
          texto: 'Entrega ' + DIAS_SEM[diaSemana(s.entrega)] + ' ' + dia(s.entrega) + ' · envió ' + s.enviadaPor,
          abrir: { sol: s.numero } });
      if (s.estado === SOL.ACE)
        out.push({ id: 'cargar-' + s.numero, tipo: 'Por despachar', cuando: s.aceptadaEn,
          titulo: s.numero + ' de ' + s.sucursal + ' aceptada', texto: 'Cuando salga el camión, márquela en camino.',
          abrir: { sol: s.numero } });
    });
    lineasSol().forEach(function (l) {
      if (l.difEstado === DIF.PEND)
        out.push({ id: 'dif-' + l.numero + '-' + l.clave, tipo: 'Diferencia', cuando: '',
          titulo: l.sucursal + ': ' + (l.diferencia > 0 ? 'faltó ' : 'llegó de más ') +
            num(Math.abs(l.diferencia)) + ' ' + plur(Math.abs(l.diferencia), l.medida),
          texto: l.nombre + ' · ' + l.numero, abrir: { tab: 'dif' } });
    });
  }
  if (yo.rol === 'gerente') {
    var mia = nombreUnidad(yo.sucursal);
    filasSol().forEach(function (s) {
      if (s.sucursal !== mia) return;
      if (s.estado === SOL.ACE)
        out.push({ id: (s.preparadaEn ? 'prep-' : 'acep-') + s.numero, tipo: s.preparadaEn ? 'Pedido preparado' : 'Pedido aceptado', cuando: s.preparadaEn || s.aceptadaEn,
          titulo: s.preparadaEn ? s.numero + ': su pedido ya está preparado' : 'Bodega aceptó ' + s.numero,
          texto: (s.preparadaEn ? 'Lo preparó ' : 'Lo están preparando · aceptó ') + s.aceptadaPor + ' · entrega ' + dia(s.entrega).slice(0, 5),
          abrir: { sol: s.numero } });
      if (s.estado === SOL.CAM)
        out.push({ id: 'cam-' + s.numero, tipo: 'Pedido en camino', cuando: s.despachadaEn,
          titulo: s.numero + ' va en camino', texto: 'Cuando llegue, recíbalo artículo por artículo.',
          abrir: { sol: s.numero, recibir: true } });
      if (s.estado === SOL.ANU && s.enviadaEn.slice(0, 10) >= desde)
        out.push({ id: 'anu-' + s.numero, tipo: 'Pedido anulado', cuando: '',
          titulo: s.numero + ' fue anulada', texto: s.motivo, abrir: { tab: 'pedidos' } });
    });
    filasPorConfirmar().forEach(function (x) {
      if (x.sucursal !== mia || x.estado !== EST_RECH || x.revisadoEn.slice(0, 10) < desde) return;
      out.push({ id: 'rech-' + x.id, tipo: 'Registro rechazado', cuando: x.revisadoEn,
        titulo: (x.tipo === 'ingreso' ? 'Corte ' + x.turno + ' del ' + dia(x.fecha) : 'Registro') +
          ' rechazado', texto: x.motivo + ' · ' + x.revisadoPor, abrir: { pag: 'registrar' } });
    });
  }
  if (yo.rol === 'gerente' || yo.rol === 'bodega') {
    var uid = yo.rol === 'gerente' ? yo.sucursal : 'bodega';
    var est = estadoInventario(uid);
    var c = calendarioConteo(uid, est.ultimo ? est.ultimo.fecha : '', aj, hoy);
    if (!c.hechoHoy && (c.tocaHoy || c.atrasado))
      out.push({ id: 'contar-' + hoy, tipo: 'Inventario', cuando: '',
        titulo: c.tocaHoy ? 'Hoy toca inventario' : 'Inventario atrasado',
        texto: c.tocaHoy ? 'Cuente ' + (yo.rol === 'gerente' ? 'y envíe su solicitud.' : 'la bodega.')
          : 'Tocaba el ' + dia(c.tocaba) + '.', abrir: { tab: 'contar' } });
  }
  try { avisosCajaKpi(yo, out, hoy); } catch (e) {}
  try { avisosMuestreoCompras(yo, out); } catch (e) {}
  try { avisosNuevos(yo, out, cred); } catch (e) {}
  out.forEach(function (a) { var t = tareaDeAviso(a); if (t) { a.tarea = true; a.accion = t.accion; a.prio = t.prio; a.urgente = /URGENTE/.test(String(a.texto || '')); } });
  out.sort(function (a, b) { return (b.cuando || '9') < (a.cuando || '9') ? -1 : 1; });
  return { avisos: out };
}

/* ── los gerentes: qué productos maneja su sucursal ── */
function productosSucursal(cred) {
  var yo = quien(cred);
  var sid = yo.rol === 'gerente' ? yo.sucursal : null;
  if (!sid) throw new Error('Esto es para los gerentes.');
  return { sucursal: nombreUnidad(sid), productos: productosInv().map(function (b) {
    return { codigo: b.codigo, nombre: b.nombre, area: b.area, grupo: b.grupo, unidad: b.unidad, descripcion: b.descripcion,
      activo: !!b.suc[sid] }; }) };
}
function activarProductoSucursal(cred, codigo, activo) {
  var yo = quien(cred);
  if (yo.rol !== 'gerente') throw new Error('Esto es para los gerentes.');
  var col = { centro: 5, almendras: 6, parque: 7 }[yo.sucursal];
  var b = null;
  filasBodega().forEach(function (x) { if (x.codigo === String(codigo)) b = x; });
  if (!b) throw new Error('No se encontró ese producto.');
  hojaBodega().getRange(b.fila, col).setValue(activo ? 'Sí' : 'No');
  _BOD = null;
  return { ok: true, mensaje: b.nombre + (activo ? ' se maneja en ' : ' ya no se maneja en ') +
    nombreUnidad(yo.sucursal) + '.' };
}

/* ── la bodega: proveedores y a quién se le compra cada producto ── */
function guardarProveedorInv(cred, p) {
  var yo = exigeBodega(cred);
  p = p || {};
  var nombre = String(p.nombre || '').trim().slice(0, 60), tel = String(p.tel || '').trim().slice(0, 30);
  if (nombre.length < 2) throw new Error('Escriba el nombre del proveedor.');
  var ya = null;
  filasProveedores().forEach(function (x) { if (x.nombre.toLowerCase() === nombre.toLowerCase()) ya = x; });
  if (ya && !esDeBodega(ya)) throw new Error('Ese nombre ya lo usa un proveedor de sucursal. Use otro.');
  var h = hojaCat('Proveedores', H_PROVEEDORES);
  if (ya) {
    if (tel) h.getRange(ya.fila, 3).setValue(tel);
    if (!ya.activo) h.getRange(ya.fila, 4).setValue('Sí');
  } else h.appendRow([nombre, '', tel, 'Sí', new Date(), yo.nombre]);
  var r = sugerenciaCompra(cred, nombre);
  r.ok = true; r.mensaje = ya ? nombre + ' actualizado.' : nombre + ' agregado.';
  return r;
}
/** Ancla (o suelta) un producto a un proveedor, con la presentación en que se compra. */
function anclarProducto(cred, d) {
  var yo = exigeBodega(cred);
  d = d || {};
  var b = null;
  filasBodega().forEach(function (x) { if (x.codigo === String(d.codigo)) b = x; });
  if (!b) throw new Error('No se encontró ese producto.');
  var prov = String(d.proveedor || '').trim().slice(0, 60);
  var pres = String(d.presCompra || '').trim().toLowerCase().slice(0, 20);
  var rinde = Number(String(d.rinde || '').replace(',', '.'));
  if (!(rinde > 0)) rinde = 1;
  var h = hojaBodega();
  h.getRange(b.fila, 13, 1, 2).setValues([[prov, Math.round(rinde * 1000) / 1000]]);
  h.getRange(b.fila, 26).setValue(prov ? pres : '');
  h.getRange(b.fila, 17, 1, 2).setValues([[new Date(), yo.nombre]]);
  _BOD = null;
  return { ok: true, mensaje: b.nombre + (prov ? ' queda con ' + prov + '.' : ' ya no tiene proveedor.') };
}
/** Todos los productos, para anclarlos (sin precios). */
function productosParaAnclar(cred) {
  exigeBodega(cred);
  return { productos: productosInv().map(function (b) {
    return { codigo: b.codigo, nombre: b.nombre, grupo: b.grupo, unidad: b.unidad,
      proveedor: b.proveedor, presCompra: b.presCompra, rinde: b.rinde }; }) };
}

/* ── reporte de traslados para el administrador ──
 * Lo que de verdad llegó a cada sucursal (traslados sin anular), por día y producto. */
function reporteTraslados(cred, desde, hasta) {
  var yoR = quien(cred);
  if (!veInv(yoR)) throw new Error('Solo un administrador.');
  if (!esFecha(desde) || !esFecha(hasta)) throw new Error('Escoja las fechas.');
  if (desde > hasta) { var x = desde; desde = hasta; hasta = x; }
  var g = libro().getSheetByName('Gastos'), vivos = {};
  leeTodo(g, 16).forEach(function (r) {
    var id = String(r[0] || '');
    if (!/^T\d/.test(id) || /-B$/.test(id) || String(r[10]).toLowerCase() === 'sí') return;
    var f = fmtDia(r[6]);
    if (f < desde || f > hasta) return;
    vivos[id] = { f: f, s: String(r[4] || ''), tr: String(r[15] || '') };
  });
  var catT = {};
  productosTraslado().forEach(function (p) { catT[p.key] = p; });
  var filas = [];
  leeTodo(libro().getSheetByName('Gastos detalle'), H_DETALLE.length).forEach(function (r) {
    var v = vivos[String(r[0])];
    if (!v) return;
    var est = String(r[13]).toLowerCase();
    if (est === 'sí' || est === 'reemplazado') return;
    var p = catT[llaveT(r[8], r[10])];
    var cant = Number(r[9]) || 0;
    filas.push({ f: v.f, s: v.s, tr: v.tr, c: p ? p.codigo : '', n: String(r[8] || ''),
      m: String(r[10] || ''), q: cant, u: p ? r3(cant * p.factor) : cant,
      un: p ? (p.suelto ? '' : p.unidad) : String(r[10] || ''), $: Number(r[12]) || 0 });
  });
  return { desde: desde, hasta: hasta, filas: filas,
    sucursales: sucursalesDestino().map(function (u) { return u.nombre; }) };
}

/* ════════════ CARGA DE SEPTIEMBRE DESDE EL PUNTO DE VENTA ════════════
 * Los «Cortes de caja» del 1 al 27 de septiembre de 2026 (Páldar POS), por sucursal:
 * turno 1 = AM, turno 2 = PM; «Gastos» del corte = gastos de caja.
 * El POS solo da la tarjeta del mes completo («Resumen de ventas»), así que se
 * reparte entre los cortes en proporción a su venta; el resto es efectivo.
 *
 * Qué hace con lo que ya estaba anotado en la app:
 *   · mismo turno y misma venta → lo deja como está (lo anotado a mano trae el
 *     efectivo y la tarjeta reales) y se descuenta su tarjeta del reparto;
 *   · mismo turno y venta distinta → lo anula («corregido con el POS») y escribe el
 *     del POS;
 *   · un corte de gerente «por confirmar» de ese turno → queda rechazado con el motivo.
 * Se puede correr las veces que sea: la segunda vez ya no hace nada.
 *
 * Primero corra revisarCortesSeptiembre() (no escribe nada, solo dice qué haría).
 * Después, cargarCortesSeptiembre(). */
var CORTES_SEPTIEMBRE = {
  centro: { tarjetaMes: 3923.00, filas: [
    '1|1|1635|217|08:03|13:00|Feliciano Tambriz', '1|2|4031|0|13:00|20:27|Ofelia Tambriz', '2|1|1702|428.5|08:08|13:12|Hans Catinac', '2|2|2967|15|13:13|20:24|Débora López',
    '3|1|1277|9|08:05|12:58|Hans Catinac', '3|2|3473|55|12:59|20:30|Ofelia Tambriz', '4|1|1908|0|08:07|13:18|Débora López', '4|2|5391|122|13:18|20:24|Manuela Guachiac',
    '5|1|2204|183|07:59|13:04|Ofelia Tambriz', '5|2|5271|4523|13:06|20:16|Hans Catinac', '6|1|2442|2|08:06|13:04|Manuela Guachiac', '6|2|4900|82|13:05|20:20|Hans Catinac',
    '7|1|1607|259|08:20|13:09|Feliciano Tambriz', '7|2|3444|174|13:10|20:37|Ofelia Tambriz', '8|1|1791|0|08:11|13:07|Débora López', '8|2|2881|0|13:09|20:44|Feliciano Tambriz',
    '9|1|2177|205|08:09|13:38|Feliciano Tambriz', '9|2|2952|92.5|13:39|20:22|Débora López', '10|1|1836|125|08:06|13:10|Hans Catinac', '10|2|2753|35|13:15|20:21|Débora López',
    '11|1|1133|111|08:09|13:05|Manuela Guachiac', '11|2|4185|50|13:06|20:48|Ofelia Tambriz', '12|1|1505|539|08:08|13:00|Ofelia Tambriz', '12|2|4586|796|13:01|20:28|Hans Catinac',
    '13|1|2482|33|08:08|13:00|Débora López', '13|2|6902|1572|13:01|20:44|Hans Catinac', '14|1|4768|43|07:59|14:05|Ofelia Tambriz', '14|2|7314|17|14:05|21:10|Ofelia Tambriz',
    '15|1|6183|53|08:01|14:55|Ofelia Tambriz', '15|2|11939|46|14:56|22:28|Ofelia Tambriz', '16|1|2356|695|08:49|13:19|Hans Catinac', '16|2|4746|0|13:24|20:30|Manuela Guachiac',
    '17|1|1718|16|08:11|13:13|Hans Catinac', '17|2|2642|49|13:14|20:26|Débora López', '18|1|1966|570|08:21|13:20|Feliciano Tambriz', '18|2|4241|0|13:31|20:28|Manuela Guachiac',
    '19|1|2290|60|08:20|12:59|Feliciano Tambriz', '19|2|3782|38|13:03|20:20|Hans Catinac', '20|1|1586|0|08:00|13:05|Ofelia Tambriz', '20|2|5725|0|13:05|20:24|Hans Catinac',
    '21|1|673|375|08:14|13:20|Débora López', '21|2|2870|104|13:21|20:46|Feliciano Tambriz', '22|1|464|0|08:05|13:11|Hans Catinac', '22|2|3710|0|13:13|20:39|Feliciano Tambriz',
    '23|1|991|90.5|08:01|13:26|Hans Catinac', '23|2|3011|0|13:30|20:36|Débora López', '24|1|989|190|08:02|13:07|Hans Catinac', '24|2|3538|0|13:08|20:17|Manuela Guachiac',
    '25|1|1396|0|07:59|12:57|Hans Catinac', '25|2|3477|143|12:58|20:14|Ofelia Tambriz', '26|1|1910|148|08:11|13:07|Débora López', '26|2|4700|2343|13:09|20:39|Ofelia Tambriz',
    '27|1|1520|150|09:56|13:18|Débora López', '27|2|4143|2100|13:18|20:19|Feliciano Tambriz'
  ] },
  almendras: { tarjetaMes: 4097.00, filas: [
    '1|1|1467|0|08:35|13:43|Jazmin Ixmatá', '1|2|2963|0|13:43|20:29|Fabiola Xocol', '2|1|1699|0|08:54|13:39|Wendy Guarchaj', '2|2|2277|0|13:40|20:20|Jazmin Ixmatá',
    '3|1|1435|0|09:15|13:38|Wendy Guarchaj', '3|2|2814|53|13:39|20:32|Fabiola Xocol', '4|1|538|0|08:42|13:32|Fabiola Xocol', '4|2|3702|8|13:33|20:30|Jazmin Ixmatá',
    '5|1|1592|0|08:42|13:50|Jazmin Ixmatá', '5|2|4102|0|13:51|20:37|Feliciano Tambriz', '6|1|1176|0|08:40|13:46|Jazmin Ixmatá', '6|2|2865|0|13:57|20:25|Wendy Guarchaj',
    '7|1|1424|0|08:45|13:45|Wendy Guarchaj', '7|2|2486|0|13:45|20:42|Fabiola Xocol', '8|1|1427|0|08:50|13:30|Fabiola Xocol', '8|2|2997|0|13:31|20:36|Jazmin Ixmatá',
    '9|1|1052|0|08:56|13:44|Wendy Guarchaj', '9|2|2176|0|13:44|20:28|Fabiola Xocol', '10|1|1823|0|08:33|13:38|Jazmin Ixmatá', '10|2|2507|0|13:39|20:00|Fabiola Xocol',
    '11|1|856|0|09:06|13:41|Wendy Guarchaj', '11|2|2163|0|13:42|20:22|Jazmin Ixmatá', '12|1|1180|0|08:53|13:32|Feliciano Tambriz', '12|2|2915|0|13:33|20:23|Jazmin Ixmatá',
    '13|1|2090|0|08:38|13:39|Jazmin Ixmatá', '13|2|4401|0|13:39|20:28|Fabiola Xocol', '14|1|6580|295|08:44|14:56|Jazmin Ixmatá', '14|2|4065|0|14:57|20:51|Jazmin Ixmatá',
    '15|1|6989|200|08:45|14:38|Fabiola Xocol', '15|2|3535|0|14:39|20:52|Fabiola Xocol', '16|1|1463|0|08:43|13:39|Jazmin Ixmatá', '16|2|2810|0|13:40|20:26|Fabiola Xocol',
    '17|1|1152|0|08:39|13:37|Fabiola Xocol', '17|2|1715|0|13:44|20:21|Wendy Guarchaj', '18|1|1496|0|08:42|13:39|Fabiola Xocol', '18|2|1601|0|13:39|20:18|Jazmin Ixmatá',
    '19|1|578|0|08:38|13:41|Jazmin Ixmatá', '19|2|2413|0|13:44|20:17|Jazmin Ixmatá', '20|1|1115|0|08:44|13:27|Wendy Guarchaj', '20|2|3883|0|13:28|20:25|Jazmin Ixmatá',
    '21|1|1197|0|08:35|13:52|Fabiola Xocol', '21|2|1881|0|13:56|20:21|Wendy Guarchaj', '22|1|974|0|08:46|13:29|Fabiola Xocol', '22|2|2930|0|13:30|20:33|Jazmin Ixmatá',
    '23|1|1093|0|08:57|13:31|Wendy Guarchaj', '23|2|2189|0|13:32|20:42|Jazmin Ixmatá', '24|1|664|0|08:38|13:27|Fabiola Xocol', '24|2|1394|0|13:27|20:17|Jazmin Ixmatá',
    '25|1|1023|0|08:33|13:33|Jazmin Ixmatá', '25|2|2100|0|13:34|21:00|Fabiola Xocol', '26|1|1124|0|08:37|13:39|Jazmin Ixmatá', '26|2|2447|0|13:39|20:15|Jazmin Ixmatá',
    '27|1|1910|0|08:47|13:51|Fabiola Xocol', '27|2|3042|0|13:52|20:30|Wendy Guarchaj'
  ] },
  parque: { tarjetaMes: 3226.00, filas: [
    '1|1|253|0|09:07|13:34|Lucrecia Chox', '1|2|2707|0|13:37|20:32|Jenny Catinac', '2|1|351|3|09:09|13:36|Shelda Xocol', '2|2|1653|0|13:39|20:19|Sandra López',
    '3|1|304|35|09:15|13:40|Shelda Xocol', '3|2|2217|0|13:42|20:23|Jenny Catinac', '4|1|602|5|09:14|13:33|Jenny Catinac', '4|2|1674|0|13:37|20:20|Lucrecia Chox',
    '5|1|568|0|09:10|13:33|Shelda Xocol', '5|2|1849|1136.5|13:40|20:30|Sandra López', '6|1|1400|0|09:17|13:31|Jenny Catinac', '6|2|3238|0|13:33|20:41|Shelda Xocol',
    '7|1|330|0|09:15|13:38|Jenny Catinac', '7|2|2131|28|13:40|20:25|Sandra López', '8|1|679|0|09:04|13:35|Lucrecia Chox', '8|2|1954|0|13:36|20:18|Shelda Xocol',
    '9|1|523|6|09:05|13:34|Sandra López', '9|2|2134|0|13:34|20:27|Lucrecia Chox', '10|1|683|0|09:08|13:36|Jenny Catinac', '10|2|1822|0|13:37|20:41|Sandra López',
    '11|1|683|0|09:10|13:36|Lucrecia Chox', '11|2|1469|125|13:39|20:21|Jenny Catinac', '12|1|464|54|09:10|13:37|Sandra López', '12|2|3001|0|13:43|20:31|Jenny Catinac',
    '13|1|1779|0|08:43|13:59|Jenny Catinac', '13|2|5488|0|14:05|20:51|Jenny Catinac', '14|1|5064|6|09:08|13:45|Shelda Xocol', '14|2|6419|575|13:47|22:17|Shelda Xocol',
    '15|1|8077|663|09:20|14:53|Jenny Catinac', '15|2|8818.75|23|14:55|22:13|Jenny Catinac', '16|1|553|0|09:18|13:48|Sandra López', '16|2|2136|0|13:49|20:20|Lucrecia Chox',
    '17|1|330|0|09:12|13:37|Jenny Catinac', '17|2|1202|0|13:39|20:24|Shelda Xocol', '18|1|583|0|09:13|13:33|Shelda Xocol', '18|2|1660|1|13:34|20:19|Lucrecia Chox',
    '19|1|709|0|09:12|13:39|Jenny Catinac', '19|2|1513|0|13:42|20:24|Sandra López', '20|1|994|0|09:15|13:25|Lucrecia Chox', '20|2|1978|1424|13:26|20:18|Shelda Xocol',
    '21|1|521|0|09:05|13:34|Sandra López', '21|2|1541|0|13:38|20:28|Lucrecia Chox', '22|1|229|0|09:09|13:41|Shelda Xocol', '22|2|1686|0|13:42|20:27|Jenny Catinac',
    '23|1|940|63.5|09:14|13:41|Lucrecia Chox', '23|2|940|152|13:42|20:14|Jenny Catinac', '24|1|740|158|09:08|13:35|Jenny Catinac', '24|2|1607|0|13:39|20:17|Lucrecia Chox',
    '25|1|708|0|09:07|13:57|Sandra López', '25|2|981|0|13:58|20:15|Shelda Xocol', '26|1|1262|0|09:09|13:45|Shelda Xocol', '26|2|1617|0|13:47|20:31|Sandra López',
    '27|1|522|8|10:05|13:42|Sandra López', '27|2|2509|0|13:45|20:28|Jenny Catinac'
  ] }
};

function revisarCortesSeptiembre() { return cortesSeptiembre(false); }
function cargarCortesSeptiembre() { return cortesSeptiembre(true); }

function cortesSeptiembre(escribir) {
  var ss = libro(), h = ss.getSheetByName('Ingresos');
  if (!h) throw new Error('Falta la hoja Ingresos. Corra instalar().');
  var lock = LockService.getScriptLock();
  lock.waitLock(30000); _LEE = {};
  try {
    var ahora = new Date(), quien = 'Punto de venta (carga)';
    var filasH = leeTodo(h, H_INGRESOS.length);
    var pend = filasPorConfirmar();
    var informe = [], nuevas = [], anular = [], rechazar = [];
    var tot = { nuevas: 0, iguales: 0, corregidas: 0, rechazadas: 0 };

    Object.keys(CORTES_SEPTIEMBRE).forEach(function (sid) {
      var u = unidadPorId(sid), cfg = CORTES_SEPTIEMBRE[sid];
      var cortes = cfg.filas.map(function (x) {
        var c = x.split('|');
        var d = ('0' + c[0]).slice(-2);
        return { fecha: '2026-09-' + d, turno: c[1] === '1' ? 'AM' : 'PM', venta: Number(c[2]),
          caja: Number(c[3]), abre: c[4], cierra: c[5], por: c[6] };
      });
      // lo que ya está en la app, por turno
      var ya = {};
      filasH.forEach(function (r, i) {
        if (String(r[4]) !== u.nombre || String(r[10]).toLowerCase() === 'sí') return;
        var f = fmtDia(r[5]);
        if (f < '2026-09-01' || f > '2026-09-27') return;
        ya[f + '|' + r[7]] = { fila: i + 2, id: String(r[0]), monto: Number(r[8]) || 0,
          tarjeta: Number(r[15]) || 0, quien: String(r[2] || '') };
      });
      var queda = [], tarjetaYa = 0, iguales = 0, corregidas = 0;
      cortes.forEach(function (c) {
        var x = ya[c.fecha + '|' + c.turno];
        if (x && Math.abs(x.monto - c.venta) < 0.01) { iguales++; tarjetaYa += x.tarjeta; return; }
        if (x) { anular.push({ fila: x.fila, motivo: 'Corregido con el corte del POS: venta ' +
          dinero(c.venta) + ' (estaba ' + dinero(x.monto) + ')' }); corregidas++;
          informe.push(u.nombre + ' ' + dia(c.fecha) + ' ' + c.turno + ': estaba ' +
            dinero(x.monto) + ' (' + x.quien + '), el POS dice ' + dinero(c.venta) + ' → se corrige'); }
        queda.push(c);
      });
      // la tarjeta del mes que falta, repartida en proporción a la venta
      var tRest = Math.max(0, Math.round((cfg.tarjetaMes - tarjetaYa) * 100) / 100);
      var vRest = queda.reduce(function (a, c) { return a + c.venta; }, 0);
      var suma = 0, mayor = null;
      queda.forEach(function (c) {
        c.tarjeta = vRest > 0 ? Math.round(tRest * c.venta / vRest * 100) / 100 : 0;
        suma += c.tarjeta;
        if (!mayor || c.venta > mayor.venta) mayor = c;
      });
      if (mayor) mayor.tarjeta = Math.round((mayor.tarjeta + tRest - suma) * 100) / 100;
      queda.forEach(function (c) {
        c.efectivo = Math.round((c.venta - c.tarjeta) * 100) / 100;
        if (c.caja > c.efectivo) {                       // no debería pasar; por si acaso
          c.tarjeta = Math.max(0, Math.round((c.venta - c.caja) * 100) / 100);
          c.efectivo = Math.round((c.venta - c.tarjeta) * 100) / 100;
        }
        var partes = cortesFecha(c.fecha, c.cierra);
        var id = 'I' + Utilities.formatDate(partes, ZONA, 'yyyyMMddHHmm') + 'POS' +
          sid.charAt(0).toUpperCase() + c.turno.charAt(0);
        nuevas.push([id, partes, quien, '', u.nombre, aDate(c.fecha), c.fecha.slice(0, 7), c.turno,
          c.venta, 'Corte POS turno ' + (c.turno === 'AM' ? 1 : 2) + ' · ' + c.abre + '–' + c.cierra +
          ' · abrió ' + c.por + ' · tarjeta repartida del total del mes', 'No', '', '', '',
          c.efectivo, c.tarjeta, c.caja, c.caja > 0 ? 'Gastos anotados en el punto de venta' : '',
          'Carga POS septiembre', ahora]);
      });
      // cortes de gerentes por confirmar de esos turnos
      pend.forEach(function (p) {
        if (p.estado !== EST_PEND || p.tipo !== 'ingreso' || p.sucursal !== u.nombre) return;
        if (p.fecha < '2026-09-01' || p.fecha > '2026-09-27') return;
        rechazar.push(p);
      });
      tot.nuevas += queda.length; tot.iguales += iguales; tot.corregidas += corregidas;
      informe.push(u.nombre + ': ' + cortes.length + ' cortes del POS · ' + iguales +
        ' ya estaban iguales · ' + corregidas + ' se corrigen · ' + queda.length +
        ' se escriben · tarjeta a repartir ' + dinero(tRest) + ' de ' + dinero(cfg.tarjetaMes));
    });
    tot.rechazadas = rechazar.length;

    if (escribir) {
      anular.forEach(function (a) {
        h.getRange(a.fila, 11, 1, 4).setValues([['Sí', a.motivo, 'Carga POS septiembre', ahora]]);
      });
      if (nuevas.length) {
        var f0 = h.getLastRow() + 1;
        h.getRange(f0, 7, nuevas.length, 1).setNumberFormat('@');
        h.getRange(f0, 1, nuevas.length, nuevas[0].length).setValues(nuevas);
      }
      var hp = hojaPorConfirmar();
      rechazar.forEach(function (p) {
        hp.getRange(p.fila, 11, 1, 4).setValues([[EST_RECH, 'Carga POS septiembre', ahora,
          'Ese turno ya se cargó del reporte de cortes del punto de venta']]);
      });
    }
    var txt = (escribir ? 'LISTO. ' : 'REVISIÓN (no se escribió nada). ') +
      tot.nuevas + ' cortes nuevos, ' + tot.iguales + ' ya estaban, ' + tot.corregidas +
      ' corregidos, ' + tot.rechazadas + ' por confirmar rechazados.\n' + informe.join('\n');
    Logger.log(txt);
    return txt;
  } finally {
    lock.releaseLock();
  }
}
/** «2026-09-05» y «20:16» → la fecha y hora de ese cierre (hora de Guatemala). */
function cortesFecha(iso, hhmm) {
  var p = iso.split('-'), t = hhmm.split(':');
  return new Date(Date.UTC(+p[0], +p[1] - 1, +p[2], +t[0] + 6, +t[1]));   // UTC−6
}

/* ════════════ INDICADORES (KPIs) DE LAS SUCURSALES ════════════
 * El director operativo arma la lista de lo que cada sucursal tiene que medir
 * (temperaturas, tiempos, merma…), con su rango aceptable y cada cuánto. Al
 * «enviarla», le aparece a cada gerente en la campanita y en su página
 * «Indicadores», donde anota el valor. La hora la pone el servidor. Si el valor sale
 * del rango, el gerente tiene que escribir qué hizo, y al director le llega el aviso.
 *
 * Hojas: «Indicadores» (la lista) e «Indicadores registros» (lo que anotaron). */
var H_KPI = ['ID', 'Indicador', 'Tipo', 'Unidad', 'Mínimo', 'Máximo', 'Frecuencia',
  'Centro', 'Almendras', 'Parque', 'Instrucciones', 'Estado', 'Categoría',
  'Creado por', 'Creado en', 'Actualizado por', 'Actualizado en', 'Enviado por', 'Enviado en',
  'Primer envío', 'Contabilidad', 'Marketing', 'Tecnología', 'Dirección operativa'];
var H_KPIREG = ['Registrado en', 'Fecha', 'Turno', 'Sucursal', 'ID indicador', 'Indicador',
  'Valor', 'Unidad', 'Mínimo', 'Máximo', 'En rango', 'Registrado por', 'Nota', 'Periodo'];
var KPI_TIPOS = ['Número', 'Sí / No', 'Texto'];
var KPI_FREC = ['Cada turno', 'Diario', 'Apertura', 'Cierre', 'Semanal', 'Mensual'];
var KPI_CATS = ['Apertura', 'Cierre', 'Inocuidad', 'Calidad', 'Servicio', 'Operación', 'Costos'];
var KPI_CATS_CONTA = ['Caja y bancos', 'Proveedores', 'Impuestos', 'Cierre contable', 'Control'];
var KPI_CATS_MKT = ['Redes sociales', 'Campañas', 'Clientes y ventas', 'Marca y diseño', 'Presupuesto'];
var KPI_EST = { BOR: 'Borrador', ENV: 'Enviado', DEL: 'Borrado' };
var SUC_KPI = ['centro', 'almendras', 'parque'];
/* A qué hora empieza el turno PM (así se sabe si ya toca medir el segundo turno). */
var INICIO_PM = { centro: '13:00', almendras: '13:30', parque: '13:30', contabilidad: '13:00', marketing: '13:00', tecnologia: '13:00', direccion: '13:00' };
/* Las «áreas» que llevan indicadores: las tres sucursales (las decide el director
 * operativo y las anota cada gerente) y contabilidad (las decide el director
 * financiero y las anota el contador interno). */
var AREA_CONTA = 'contabilidad';
/* Y marketing: la CMO (oficinas). Los indicadores de marketing los decide y envía el administrador. */
var AREA_MKT = 'marketing';
var _KPI_MKT = false;           // true cuando el administrador trabaja los indicadores de la CMO
/* Y tecnología: los decide y los anota el administrador. */
var AREA_TEC = 'tecnologia';
var _KPI_TEC = false;           // true cuando el administrador trabaja los indicadores de tecnología
/* Y la dirección operativa: los indicadores del director operativo. Los decide el administrador y los anota el director. */
var AREA_COO = 'direccion';
var _KPI_COO = false;           // true cuando se trabajan los indicadores del director operativo
function nombreArea(sid) { return sid === AREA_CONTA ? 'Contabilidad' : sid === AREA_MKT ? 'Marketing' : sid === AREA_TEC ? 'Tecnología' : sid === AREA_COO ? 'Dirección operativa' : nombreUnidad(sid); }
function areasDe(yo) {
  if (_KPI_COO && (yo.esAdmin || yo.rol === 'operaciones')) return [AREA_COO];
  if (_KPI_TEC && yo.esAdmin) return [AREA_TEC];
  if (_KPI_MKT && yo.esAdmin) return [AREA_MKT];
  if (yo.rol === 'operaciones') return SUC_KPI.slice();
  if (yo.rol === 'finanzas') return [AREA_CONTA];
  if (yo.esAdmin) return SUC_KPI.concat([AREA_CONTA]);
  return [];
}
/** Quién anota: el gerente su sucursal; el contador interno, contabilidad. */
function areaQueAnota(yo) {
  if (_KPI_COO && (yo.esAdmin || yo.rol === 'operaciones')) return AREA_COO;
  if (_KPI_TEC && yo.esAdmin) return AREA_TEC;
  if (yo.rol === 'gerente') return yo.sucursal;
  if (yo.rol === 'registro' && yo.recibe) return AREA_CONTA;
  if (yo.rol === 'cmo') return AREA_MKT;
  return '';
}
function kpiEnAreas(k, areas) { return areas.some(function (a) { return k.suc[a]; }); }

/* Los que Claude recomienda para comida rápida (pizzas y hamburguesas). Los rangos
 * son los de referencia de inocuidad (FDA Food Code / Codex); el director los puede
 * cambiar al agregarlos. */
var KPIS_RECOMENDADOS = [
  { c: 'Inocuidad', n: 'Temperatura del refrigerador / cuarto frío', t: 'Número', u: '°C', min: 0, max: 4, f: 'Cada turno',
    i: 'Termómetro en el centro del equipo, con la puerta cerrada. Si pasa de 5 °C: revise el empaque de la puerta, no lo llene de más y avise.',
    p: 'Arriba de 5 °C las bacterias se multiplican en queso, carnes y salsas.' },
  { c: 'Inocuidad', n: 'Temperatura del congelador', t: 'Número', u: '°C', min: -25, max: -18, f: 'Cada turno',
    i: 'Lea el termómetro del congelador. Si está arriba de −18 °C, no abra más de lo necesario y avise.',
    p: 'Carne de hamburguesa, papas y pollo deben mantenerse a −18 °C o menos.' },
  { c: 'Inocuidad', n: 'Temperatura de la mesa fría de preparación', t: 'Número', u: '°C', min: 0, max: 4, f: 'Cada turno',
    i: 'Mida un ingrediente del riel (queso o jamón), no el aire.',
    p: 'Es donde más tiempo pasan los ingredientes de la pizza fuera del refrigerador.' },
  { c: 'Inocuidad', n: 'Temperatura interna de la carne de hamburguesa', t: 'Número', u: '°C', min: 71, max: null, f: 'Cada turno',
    i: 'Termómetro de punta en la parte más gruesa de una torta recién salida de la plancha.',
    p: 'La carne molida tiene que llegar a 71 °C por dentro para eliminar E. coli.' },
  { c: 'Inocuidad', n: 'Temperatura del producto en mantenimiento caliente', t: 'Número', u: '°C', min: 60, max: null, f: 'Cada turno',
    i: 'Mida lo que está en el calentador o baño maría (carne, salsas, pollo).',
    p: 'Abajo de 60 °C la comida caliente entra a la zona de peligro.' },
  { c: 'Calidad', n: 'Temperatura del horno de pizza', t: 'Número', u: '°C', min: 230, max: 300, f: 'Cada turno',
    i: 'Lea el tablero o use termómetro infrarrojo en la piedra/banda. Ajuste el rango a su horno.',
    p: 'Si el horno está bajo, la pizza sale cruda o hay que dejarla más tiempo y se atrasan los pedidos.' },
  { c: 'Calidad', n: 'Temperatura del aceite de la freidora', t: 'Número', u: '°C', min: 170, max: 185, f: 'Cada turno',
    i: 'Termómetro de la freidora o de punta.',
    p: 'Aceite frío = papas grasosas; muy caliente = aceite quemado que dura menos.' },
  { c: 'Calidad', n: 'Calidad del aceite (compuestos polares)', t: 'Número', u: '%', min: null, max: 24, f: 'Diario',
    i: 'Con tira o medidor de aceite. Si pasa de 24 %, se cambia el aceite.',
    p: 'Cambiar el aceite a tiempo mantiene el sabor y evita cambiarlo antes de lo necesario.' },
  { c: 'Inocuidad', n: 'Producto vencido o sin etiqueta de fecha', t: 'Número', u: 'unidades', min: null, max: 0, f: 'Diario',
    i: 'Revise refrigerador, congelador y estantes. Cuente lo que encontró vencido o sin fecha y sáquelo.',
    p: 'Asegura que se use primero lo que entró primero (PEPS).' },
  { c: 'Inocuidad', n: 'Limpieza y desinfección de superficies y equipos al cierre', t: 'Sí / No', u: '', min: null, max: null, f: 'Diario',
    i: 'Mesas, plancha, cortadoras, mesa fría y pisos. Marque Sí solo si se completó todo.',
    p: 'Evita contaminación cruzada entre carne cruda y lo que ya está listo.' },
  { c: 'Operación', n: 'Limpieza de la campana y extractor', t: 'Sí / No', u: '', min: null, max: null, f: 'Semanal',
    i: 'Filtros de la campana lavados y grasa retirada.',
    p: 'La grasa acumulada es riesgo de incendio.' },
  { c: 'Costos', n: 'Masas de pizza descartadas', t: 'Número', u: 'unidades', min: null, max: 5, f: 'Diario',
    i: 'Cuente las masas que se botaron (pasadas, rotas, mal fermentadas).',
    p: 'Muestra si se está preparando más masa de la que se vende.' },
  { c: 'Costos', n: 'Merma del día (lo que se botó)', t: 'Número', u: 'Q', min: null, max: 150, f: 'Diario',
    i: 'Valor aproximado en quetzales de lo que se botó: pizzas mal hechas, hamburguesas devueltas, producto vencido.',
    p: 'La merma se come el margen; medirla en dinero la hace visible.' },
  { c: 'Servicio', n: 'Tiempo promedio de entrega a domicilio', t: 'Número', u: 'min', min: null, max: 30, f: 'Diario',
    i: 'Desde que entra el pedido hasta que se entrega, promedio del día (del punto de venta o de la plataforma).',
    p: 'Es lo que más pesa en si el cliente vuelve a pedir.' },
  { c: 'Servicio', n: 'Tiempo de preparación en cocina', t: 'Número', u: 'min', min: null, max: 12, f: 'Diario',
    i: 'Desde que entra el pedido hasta que sale de cocina, promedio del día.',
    p: 'Muestra si falta personal o si la cocina está mal organizada en horas pico.' },
  { c: 'Servicio', n: 'Quejas o pedidos devueltos', t: 'Número', u: 'pedidos', min: null, max: 1, f: 'Diario',
    i: 'Cuente las quejas del día (en tienda, teléfono, redes o plataformas) y anote de qué fueron.',
    p: 'Cada queja sin atender es un cliente que no regresa.' },
  { c: 'Operación', n: 'Personal que faltó o llegó tarde', t: 'Número', u: 'personas', min: null, max: 0, f: 'Diario',
    i: 'Cuente contra el turno programado.',
    p: 'Con turnos rotativos y un día de descanso, una falta descuadra el servicio.' },
  { c: 'Operación', n: 'Revisión de fechas y rotación de producto (PEPS)', t: 'Sí / No', u: '', min: null, max: null, f: 'Semanal',
    i: 'Lo más viejo adelante, lo nuevo atrás, todo con fecha.',
    p: 'Evita que el producto se venza en la bodega de la sucursal.' },
  { c: 'Apertura', n: 'Personal completo, con uniforme y gorra o redecilla', t: 'Sí / No', u: '', min: null, max: null, f: 'Apertura',
    i: 'Compare con el horario del día.', p: 'Un puesto vacío en la apertura atrasa todo el turno.' },
  { c: 'Apertura', n: 'Lavado de manos al entrar', t: 'Sí / No', u: '', min: null, max: null, f: 'Apertura',
    i: 'Todos, antes de tocar comida.', p: 'La primera barrera contra contaminación.' },
  { c: 'Apertura', n: 'Horno, freidora y plancha encendidos y a temperatura', t: 'Sí / No', u: '', min: null, max: null, f: 'Apertura',
    i: 'Antes de abrir la puerta al cliente.', p: 'Si el horno no está listo, los primeros pedidos salen tarde o crudos.' },
  { c: 'Apertura', n: 'Mise en place lista (masas, salsa, queso, verduras cortadas)', t: 'Sí / No', u: '', min: null, max: null, f: 'Apertura',
    i: 'Lo suficiente para las primeras horas.', p: 'Preparar durante la hora pico duplica el tiempo de entrega.' },
  { c: 'Apertura', n: 'Fondo de caja contado y cuadrado', t: 'Sí / No', u: '', min: null, max: null, f: 'Apertura',
    i: 'Cuente el fondo con quien entrega el turno.', p: 'Si no se cuadra al abrir, cualquier faltante del cierre no tiene responsable.' },
  { c: 'Apertura', n: 'Punto de venta y terminal NEONET funcionando', t: 'Sí / No', u: '', min: null, max: null, f: 'Apertura',
    i: 'Haga una prueba de conexión.', p: 'Sin terminal se pierden las ventas con tarjeta.' },
  { c: 'Apertura', n: 'Comedor, baños y entrada limpios', t: 'Sí / No', u: '', min: null, max: null, f: 'Apertura',
    i: 'Revise antes de abrir.', p: 'Es lo primero que ve el cliente.' },
  { c: 'Apertura', n: 'Refrigerador revisado: fechas y lo más viejo adelante', t: 'Sí / No', u: '', min: null, max: null, f: 'Apertura',
    i: 'Primero en entrar, primero en salir.', p: 'Evita usar producto vencido y reduce la merma.' },
  { c: 'Cierre', n: 'Producto guardado, tapado y etiquetado con fecha', t: 'Sí / No', u: '', min: null, max: null, f: 'Cierre',
    i: 'Todo lo abierto con la fecha del día.', p: 'Lo que queda destapado se contamina o se seca y se bota.' },
  { c: 'Cierre', n: 'Equipos apagados y limpios (horno, freidora, plancha, cortadora)', t: 'Sí / No', u: '', min: null, max: null, f: 'Cierre',
    i: 'Con el producto de limpieza indicado.', p: 'La grasa acumulada daña el equipo y es riesgo de incendio.' },
  { c: 'Cierre', n: 'Pisos, mesas y baños limpios; basura afuera', t: 'Sí / No', u: '', min: null, max: null, f: 'Cierre',
    i: 'Bolsas cerradas, fuera del local.', p: 'Evita plagas durante la noche.' },
  { c: 'Cierre', n: 'Gas y llaves de agua cerrados', t: 'Sí / No', u: '', min: null, max: null, f: 'Cierre',
    i: 'Revise la llave general de gas.', p: 'Seguridad del local y del personal.' },
  { c: 'Cierre', n: 'Refrigeradores y congeladores cerrados y funcionando', t: 'Sí / No', u: '', min: null, max: null, f: 'Cierre',
    i: 'Puertas bien cerradas; anote si algo suena raro.', p: 'Una puerta abierta de noche es producto perdido.' },
  { c: 'Cierre', n: 'Corte de caja hecho y enviado en la app', t: 'Sí / No', u: '', min: null, max: null, f: 'Cierre',
    i: 'El corte PM del día.', p: 'Sin el corte no se puede validar ni recibir el dinero.' },
  { c: 'Cierre', n: 'Puertas, ventanas y alarma aseguradas', t: 'Sí / No', u: '', min: null, max: null, f: 'Cierre',
    i: 'Último en salir revisa.', p: 'Seguridad del local.' }
];
/* Las tareas que el director financiero le puede pedir al contador interno. */
var KPIS_CONTA_RECOMENDADOS = [
  { c: 'Caja y bancos', n: 'Recibir y cuadrar el efectivo de los cierres del día', t: 'Sí / No', u: '', min: null, max: null, f: 'Diario',
    i: 'Cada cierre validado por el director operativo se recibe en la app el mismo día. Si falta dinero, se anota la cantidad real.',
    p: 'Un cierre sin recibir es dinero sin responsable.' },
  { c: 'Caja y bancos', n: 'Depositar en el banco el efectivo recibido', t: 'Sí / No', u: '', min: null, max: null, f: 'Diario',
    i: 'Guarde la boleta. El efectivo no pasa la noche en la oficina.',
    p: 'Reduce el riesgo de robo y deja rastro en el estado de cuenta.' },
  { c: 'Caja y bancos', n: 'Revisar que el depósito de NEONET llegó al banco', t: 'Sí / No', u: '', min: null, max: null, f: 'Diario',
    i: 'La tarjeta líquida de los cierres anteriores (tarjeta menos 6.11 %) debe aparecer en la cuenta de CAIX.',
    p: 'Si NEONET no deposita o retiene de más, se nota el mismo día y no a fin de mes.' },
  { c: 'Proveedores', n: 'Registrar las facturas de gastos del día', t: 'Sí / No', u: '', min: null, max: null, f: 'Diario',
    i: 'Las que no vienen de una orden de compra (servicios, mantenimiento, etc.).',
    p: 'Un gasto sin registrar no sale en el estado de resultados.' },
  { c: 'Caja y bancos', n: 'Conciliación bancaria de la semana', t: 'Sí / No', u: '', min: null, max: null, f: 'Semanal',
    i: 'Lo depositado y lo pagado en la app contra el estado de cuenta.',
    p: 'Encuentra errores y cobros del banco a tiempo.' },
  { c: 'Proveedores', n: 'Revisar órdenes de compra por pagar', t: 'Sí / No', u: '', min: null, max: null, f: 'Semanal',
    i: 'En Contabilidad → Pagos a proveedores: cotizadas sin pagar y recibidas sin pagar.',
    p: 'Pagar a tiempo mantiene el crédito y los precios con los proveedores.' },
  { c: 'Control', n: 'Faltantes de efectivo de la semana', t: 'Número', u: 'Q', min: null, max: 50, f: 'Semanal',
    i: 'Sume los faltantes de efectivo de los cierres de la semana (en positivo).',
    p: 'Más de Q50 a la semana hay que revisarlo con el director operativo.' },
  { c: 'Impuestos', n: 'Libro de compras y ventas (IVA) al día', t: 'Sí / No', u: '', min: null, max: null, f: 'Mensual',
    i: 'Todas las facturas FEL del mes cargadas antes de declarar.',
    p: 'Sin el libro al día el IVA sale mal y la SAT multa.' },
  { c: 'Impuestos', n: 'IVA del mes declarado y pagado antes del vencimiento', t: 'Sí / No', u: '', min: null, max: null, f: 'Mensual',
    i: 'Vence el último día hábil del mes siguiente. Anote el pago en Impuestos.',
    p: 'Evita multas y recargos.' },
  { c: 'Impuestos', n: 'Facturas de compras verificadas en FEL', t: 'Sí / No', u: '', min: null, max: null, f: 'Mensual',
    i: 'Que cada factura de proveedor exista en el portal de la SAT y esté a nombre de CAIX, S.A. con su NIT.',
    p: 'Una factura inválida no da crédito fiscal de IVA.' },
  { c: 'Cierre contable', n: 'Cierre del mes y estado de resultados', t: 'Sí / No', u: '', min: null, max: null, f: 'Mensual',
    i: 'Antes del día 10 del mes siguiente, con el PDF del mes de la app.',
    p: 'Las decisiones del mes siguiente se toman con números cerrados.' },
  { c: 'Cierre contable', n: 'Comisiones NEONET del mes cuadran con su estado de cuenta', t: 'Sí / No', u: '', min: null, max: null, f: 'Mensual',
    i: 'El total de comisiones de la app contra lo que cobró NEONET.',
    p: 'Si NEONET cobra de más, se reclama con el comprobante.' },
  { c: 'Proveedores', n: 'Resumen de pagos a proveedores revisado con el director financiero', t: 'Sí / No', u: '', min: null, max: null, f: 'Mensual',
    i: 'Contabilidad → Resumen del mes.',
    p: 'Permite negociar precios con quien más se le compra.' }
];

/* Los indicadores que Claude propone para la CMO (marketing de una cadena de pizzerías con delivery). Marcelino los revisa,
 * ajusta las metas a la realidad de Rey Pizza y los envía desde Configuración → Indicadores de la CMO. */
var KPIS_MKT_RECOMENDADOS = [
  { c: 'Redes sociales', n: 'Publicaciones de la semana en Facebook e Instagram', t: 'Número', u: 'publicaciones', min: 5, max: null, f: 'Semanal',
    i: 'Cuente las publicaciones y reels de las páginas de Rey Pizza (las 3 sucursales cuentan como una marca).', p: 'La constancia mantiene a la marca presente cuando el cliente decide qué pedir.' },
  { c: 'Redes sociales', n: 'Reels o videos cortos publicados en la semana', t: 'Número', u: 'videos', min: 2, max: null, f: 'Semanal',
    i: 'Preparación de pizzas, promociones o el equipo. Los videos cortos son lo que más alcance orgánico da.', p: 'El video corto llega a gente que todavía no nos sigue.' },
  { c: 'Redes sociales', n: 'Interacción de las publicaciones (reacciones + comentarios + compartidos entre el alcance)', t: 'Número', u: '%', min: 3, max: null, f: 'Mensual',
    i: 'Tome el dato de las estadísticas de la página del mes. Anote el porcentaje.', p: 'Mide si el contenido gusta, no solo cuánta gente lo vio.' },
  { c: 'Redes sociales', n: 'Seguidores nuevos del mes (neto, Facebook + Instagram)', t: 'Número', u: 'seguidores', min: 100, max: null, f: 'Mensual',
    i: 'Nuevos menos los que dejaron de seguirnos. Ajuste la meta al punto de partida.', p: 'La comunidad es la base para promociones sin pagar anuncios.' },
  { c: 'Clientes y ventas', n: 'Calificación promedio en Google Maps', t: 'Número', u: 'estrellas', min: 4.5, max: null, f: 'Mensual',
    i: 'La calificación de cada sucursal; anote la más baja y explique en la nota cuál es.', p: 'La mayoría de clientes nuevos nos encuentra en el mapa.' },
  { c: 'Clientes y ventas', n: 'Reseñas nuevas en Google Maps (las 3 sucursales)', t: 'Número', u: 'reseñas', min: 10, max: null, f: 'Mensual',
    i: 'Pídalas en el mostrador, en el empaque del delivery o con un código QR.', p: 'Más reseñas recientes suben la posición en el mapa.' },
  { c: 'Clientes y ventas', n: 'Reseñas negativas (1 o 2 estrellas) respondidas en menos de 24 horas', t: 'Sí / No', u: '', min: null, max: null, f: 'Semanal',
    i: 'Responda con calma y ofrezca solucionar por mensaje privado. Avise al director operativo si el problema es de una sucursal.', p: 'Una respuesta rápida recupera clientes y se ve a la vista de todos.' },
  { c: 'Clientes y ventas', n: 'Crecimiento de la venta del mes frente al mes anterior', t: 'Número', u: '%', min: 0, max: null, f: 'Mensual',
    i: 'Use la venta total del reporte del mes (contabilidad se la comparte). Anote el porcentaje.', p: 'La publicidad se paga sola si la venta crece.' },
  { c: 'Clientes y ventas', n: 'Pedidos a domicilio que llegaron por redes o WhatsApp en el mes', t: 'Número', u: 'pedidos', min: null, max: null, f: 'Mensual',
    i: 'Cuéntelos con los gerentes (preguntando cómo se enteró el cliente) o con las estadísticas de WhatsApp Business.', p: 'Permite saber si la publicidad trae pedidos y no solo likes.' },
  { c: 'Campañas', n: 'Promoción del mes planificada y avisada a las 3 sucursales antes del día 25', t: 'Sí / No', u: '', min: null, max: null, f: 'Mensual',
    i: 'Con precios, fechas y qué producto se promueve; coordínelo con el director operativo para que haya inventario.', p: 'Una promoción sin producto en bodega es mala publicidad.' },
  { c: 'Campañas', n: 'Resultado de la campaña del mes: pedidos o ventas atribuidas', t: 'Texto', u: '', min: null, max: null, f: 'Mensual',
    i: 'Anote la campaña, lo que se invirtió y lo que trajo (cupones canjeados, pedidos, mensajes).', p: 'Lo que se mide se puede repetir o corregir.' },
  { c: 'Marca y diseño', n: 'Calendario de contenido del mes siguiente listo antes del día 25', t: 'Sí / No', u: '', min: null, max: null, f: 'Mensual',
    i: 'Fechas, tema y quién produce cada publicación.', p: 'Evita publicar a última hora y con baja calidad.' },
  { c: 'Marca y diseño', n: 'Menús, rótulos y material impreso al día en las 3 sucursales', t: 'Sí / No', u: '', min: null, max: null, f: 'Mensual',
    i: 'Precios y promociones vigentes; que no haya material viejo a la vista del cliente.', p: 'Un precio distinto en el menú y en la caja genera reclamos.' },
  { c: 'Presupuesto', n: 'Gasto de marketing y oficina dentro del presupuesto aprobado', t: 'Sí / No', u: '', min: null, max: null, f: 'Mensual',
    i: 'Compare con «Avance de mi gasto». Si se pasó, explique en la nota por qué.', p: 'Cada quetzal de publicidad debe tener un resultado esperado.' }
];

/* Los indicadores de la CMO los maneja el administrador: estas funciones solo prenden esa marca y llaman a las de siempre. */
function exigeAdminMkt(cred) { var yo = exigeAdmin(cred); _KPI_MKT = true; return yo; }
function listaKpisCmo(cred) { exigeAdminMkt(cred); return listaKpis(cred); }
function guardarKpiCmo(cred, k) { exigeAdminMkt(cred); return guardarKpi(cred, k); }
function borrarKpiCmo(cred, id) { exigeAdminMkt(cred); return borrarKpi(cred, id); }
function enviarKpisCmo(cred, ids) { exigeAdminMkt(cred); return enviarKpis(cred, ids); }
function agregarRecomendadosCmo(cred, claves) { exigeAdminMkt(cred); return agregarRecomendados(cred, claves); }
function resultadosKpiCmo(cred, desde, hasta) { exigeAdminMkt(cred); return resultadosKpiRango(cred, desde, hasta, ''); }

function hojasIndicadores(ss) {
  ss = ss || libro();
  var h = ss.getSheetByName('Indicadores');
  if (!h) {
    h = hojaLimpia(ss, 'Indicadores', H_KPI);
    h.getRange('O:O').setNumberFormat('dd/mm/yyyy hh:mm');
    h.getRange('Q:Q').setNumberFormat('dd/mm/yyyy hh:mm');
    h.getRange('S:T').setNumberFormat('dd/mm/yyyy hh:mm');
    h.setColumnWidth(2, 280); h.setColumnWidth(11, 320);
  } else {
    if (h.getLastColumn() < 22) encabezaAlFinal(h, 22, ['Marketing']);      // el área de la CMO
    if (h.getLastColumn() < 23) encabezaAlFinal(h, 23, ['Tecnología']);     // el área de tecnología
    if (h.getLastColumn() < 24) encabezaAlFinal(h, 24, ['Dirección operativa']);   // el área del director operativo
  }
  var r = ss.getSheetByName('Indicadores registros');
  if (!r) {
    r = hojaLimpia(ss, 'Indicadores registros', H_KPIREG);
    r.getRange('A:A').setNumberFormat('dd/mm/yyyy hh:mm');
    r.getRange('B:B').setNumberFormat('@');
    r.setColumnWidth(6, 280); r.setColumnWidth(13, 260);
  }
  return { lista: h, reg: r };
}
function hojaKpi() { return hojasIndicadores().lista; }
function hojaKpiReg() { return hojasIndicadores().reg; }

function numONulo(v) {
  if (v === '' || v == null) return null;
  var n = Number(v);
  return isFinite(n) ? n : null;
}
function filasKpi() {
  return leeTodo(hojaKpi(), H_KPI.length).map(function (r, i) {
    return { fila: i + 2, id: String(r[0] || ''), nombre: String(r[1] || ''),
      tipo: KPI_TIPOS.indexOf(String(r[2])) >= 0 ? String(r[2]) : 'Número',
      unidad: String(r[3] || ''), min: numONulo(r[4]), max: numONulo(r[5]),
      frec: KPI_FREC.indexOf(String(r[6])) >= 0 ? String(r[6]) : 'Diario',
      suc: { centro: siNo(r[7]), almendras: siNo(r[8]), parque: siNo(r[9]), contabilidad: siNo(r[20]), marketing: siNo(r[21]), tecnologia: siNo(r[22]), direccion: siNo(r[23]) },
      instr: String(r[10] || ''), estado: String(r[11] || '') || KPI_EST.BOR,
      cat: String(r[12] || ''), creadoPor: String(r[13] || ''), creadoEn: fmtSello(r[14]),
      actPor: String(r[15] || ''), actEn: fmtSello(r[16]),
      envPor: String(r[17] || ''), envEn: fmtSello(r[18]), desde: fmtSello(r[19] || r[18]) };
  }).filter(function (k) { return k.id; });
}
function kpiParaFuera(k) {
  return { id: k.id, nombre: k.nombre, tipo: k.tipo, unidad: k.unidad, min: k.min, max: k.max,
    frec: k.frec, suc: k.suc, instr: k.instr, estado: k.estado, cat: k.cat,
    creadoPor: k.creadoPor, creadoEn: k.creadoEn, actPor: k.actPor, actEn: k.actEn,
    envPor: k.envPor, envEn: k.envEn,
    desde: k.desde };
}
function rangoTxt(k) {
  if (k.tipo === 'Sí / No') return 'Sí';
  if (k.tipo !== 'Número') return '';
  var c = function (n) { return conUnidad(n, k.unidad); };
  if (k.min != null && k.max != null)
    return k.unidad === 'Q' ? c(k.min) + ' a ' + c(k.max) : num(k.min) + ' a ' + c(k.max);
  if (k.min != null) return 'mínimo ' + c(k.min);
  if (k.max != null) return 'máximo ' + c(k.max);
  return '';
}
function conUnidad(n, u) { return u === 'Q' ? 'Q' + num(n) : num(n) + (u ? ' ' + u : ''); }

function exigeOperaciones(cred) {            // leer: los directores y el administrador
  var yo = quien(cred);
  if (yo.rol !== 'operaciones' && yo.rol !== 'finanzas' && !yo.esAdmin)
    throw new Error('Los indicadores los maneja el director operativo.');
  return yo;
}
function exigeDirector(cred) {               // cambiar: el director de esa área
  var yo = quien(cred);
  if ((_KPI_MKT || _KPI_TEC || _KPI_COO) && yo.esAdmin) return yo;      // los de la CMO y los de tecnología los decide el administrador
  if (yo.rol !== 'operaciones' && yo.rol !== 'finanzas')
    throw new Error('Los indicadores los decide el director operativo. Usted los puede consultar.');
  return yo;
}
function kpiMio(yo, k) {
  if (!k || k.estado === KPI_EST.DEL) throw new Error('No se encontró ese indicador.');
  if (!kpiEnAreas(k, areasDe(yo))) throw new Error('Ese indicador lo maneja otro director.');
  return k;
}

/** La ventana del director: la lista y los recomendados. */
function listaKpis(cred) {
  var yo = exigeOperaciones(cred), areas = areasDe(yo);
  var ks = filasKpi().filter(function (k) { return k.estado !== KPI_EST.DEL && kpiEnAreas(k, areas); });
  var nombres = {};
  ks.forEach(function (k) { nombres[sinAcentoS(k.nombre)] = true; });
  var conta = yo.rol === 'finanzas', mkt = _KPI_MKT && yo.esAdmin, tec = _KPI_TEC && yo.esAdmin, coo = _KPI_COO && yo.esAdmin;
  return { puedeCambiar: yo.rol === 'operaciones' || conta || mkt || tec || coo, conta: conta, mkt: mkt, tec: tec, coo: coo, areas: areas.map(function (a) {
      return { id: a, nombre: nombreArea(a) }; }),
    indicadores: ks.map(kpiParaFuera), tipos: KPI_TIPOS, frecuencias: KPI_FREC,
    categorias: coo ? KPI_CATS_COO : tec ? KPI_CATS_TEC : mkt ? KPI_CATS_MKT : conta ? KPI_CATS_CONTA : KPI_CATS,
    recomendados: (coo ? KPIS_COO_RECOMENDADOS : tec ? KPIS_TEC_RECOMENDADOS : mkt ? KPIS_MKT_RECOMENDADOS : conta ? KPIS_CONTA_RECOMENDADOS : KPIS_RECOMENDADOS).map(function (r, i) {
      return { clave: i, cat: r.c, nombre: r.n, tipo: r.t, unidad: r.u, min: r.min, max: r.max,
        frec: r.f, instr: r.i, porque: r.p, rango: rangoTxt({ tipo: r.t, unidad: r.u, min: r.min, max: r.max }),
        yaEsta: !!nombres[sinAcentoS(r.n)] }; }) };
}
function sinAcentoS(t) {
  return String(t || '').toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '')
    .replace(/\s+/g, ' ').trim();
}

function limpiaKpi(k, yo) {
  k = k || {};
  var nombre = String(k.nombre || '').trim().replace(/\s+/g, ' ').slice(0, 120);
  if (nombre.length < 3) throw new Error('Escriba el nombre del indicador.');
  var tipo = KPI_TIPOS.indexOf(k.tipo) >= 0 ? k.tipo : 'Número';
  var frec = KPI_FREC.indexOf(k.frec) >= 0 ? k.frec : 'Diario';
  // la apertura y el cierre son listas de chequeo: cada punto es Sí / No
  if (frec === 'Apertura' || frec === 'Cierre') tipo = 'Sí / No';
  var min = tipo === 'Número' ? numONulo(k.min) : null;
  var max = tipo === 'Número' ? numONulo(k.max) : null;
  if (min != null && max != null && min > max) throw new Error('El mínimo no puede ser mayor que el máximo.');
  var suc = {}, areas = areasDe(yo);
  SUC_KPI.concat([AREA_CONTA, AREA_MKT, AREA_TEC, AREA_COO]).forEach(function (s) { suc[s] = areas.indexOf(s) >= 0 && !!(k.suc && k.suc[s]); });
  if (yo.rol === 'finanzas') suc.contabilidad = true;
  if (_KPI_MKT && yo.esAdmin) suc.marketing = true;
  if (_KPI_TEC && yo.esAdmin) suc.tecnologia = true;
  if (_KPI_COO && yo.esAdmin) suc.direccion = true;
  if (!suc.centro && !suc.almendras && !suc.parque && !suc.contabilidad && !suc.marketing && !suc.tecnologia && !suc.direccion) throw new Error('Escoja al menos una sucursal.');
  return { nombre: nombre, tipo: tipo, unidad: tipo === 'Número' ? String(k.unidad || '').trim().slice(0, 20) : '',
    min: min, max: max, frec: frec, suc: suc, instr: String(k.instr || '').trim().slice(0, 400),
    cat: KPI_CATS.concat(KPI_CATS_CONTA, KPI_CATS_MKT, KPI_CATS_TEC, KPI_CATS_COO).indexOf(k.cat) >= 0 ? k.cat : '' };
}
function filaKpi(id, v, estado, yo, creadoPor, creadoEn, ahora, envPor, envEn, desde) {
  return [id, v.nombre, v.tipo, v.unidad, v.min == null ? '' : v.min, v.max == null ? '' : v.max, v.frec,
    v.suc.centro ? 'Sí' : 'No', v.suc.almendras ? 'Sí' : 'No', v.suc.parque ? 'Sí' : 'No',
    v.instr, estado, v.cat, creadoPor, creadoEn, yo.nombre, ahora, envPor || '', envEn || '',
    desde || '', v.suc.contabilidad ? 'Sí' : 'No', v.suc.marketing ? 'Sí' : 'No', v.suc.tecnologia ? 'Sí' : 'No', v.suc.direccion ? 'Sí' : 'No'];
}
function nuevoIdKpi() {
  return 'K' + Utilities.formatDate(new Date(), ZONA, 'yyyyMMddHHmmss') +
    Math.floor(Math.random() * 900 + 100);
}

/** Agrega (sin id) o cambia (con id). Uno nuevo queda en borrador hasta que se envía;
 *  si se cambia uno ya enviado, el gerente recibe el aviso del cambio. */
function guardarKpi(cred, k) {
  var yo = exigeDirector(cred);
  var v = limpiaKpi(k, yo);
  var lock = LockService.getScriptLock();
  lock.waitLock(15000); _LEE = {};
  var msj;
  try {
    var todos = filasKpi(), ahora = new Date();
    var otro = todos.filter(function (x) {
      return x.estado !== KPI_EST.DEL && x.id !== String(k.id || '') &&
        sinAcentoS(x.nombre) === sinAcentoS(v.nombre); });
    if (otro.length) throw new Error('Ya hay un indicador con ese nombre.');
    if (k.id) {
      var x = null;
      todos.forEach(function (t) { if (t.id === String(k.id)) x = t; });
      kpiMio(yo, x);
      // si ya estaba enviado, el cambio les llega de una vez a los gerentes
      var env = x.estado === KPI_EST.ENV;
      hojaKpi().getRange(x.fila, 1, 1, H_KPI.length).setValues([filaKpi(x.id, v, x.estado, yo,
        x.creadoPor, x.creadoEn, ahora, env ? yo.nombre : x.envPor, env ? ahora : x.envEn,
        env ? (x.desde || ahora) : '')]);
      msj = env ? 'Cambiado y enviado otra vez a los gerentes.'
        : 'Cambiado. Sigue en borrador hasta que lo envíe.';
      if (_KPI_TEC) msj = 'Cambiado.';
    } else if (_KPI_TEC) {
      hojaKpi().appendRow(filaKpi(nuevoIdKpi(), v, KPI_EST.ENV, yo, yo.nombre, ahora, ahora, yo.nombre, ahora, ahora));
      msj = 'Agregado. Ya aparece en «Anotar hoy».';
    } else {
      hojaKpi().appendRow(filaKpi(nuevoIdKpi(), v, KPI_EST.BOR, yo, yo.nombre, ahora, ahora, '', '', ''));
      msj = 'Agregado en borrador. Envíelo para que les aparezca a los gerentes.';
    }
  } finally { lock.releaseLock(); }
  var l = listaKpis(cred); l.ok = true; l.mensaje = msj;
  return l;
}

/** Borrar no quita lo que ya se anotó: el indicador queda «Borrado» y deja de pedirse. */
function borrarKpi(cred, id) {
  var yo = exigeDirector(cred);
  var x = null;
  filasKpi().forEach(function (t) { if (t.id === String(id)) x = t; });
  kpiMio(yo, x);
  hojaKpi().getRange(x.fila, 12).setValue(KPI_EST.DEL);
  hojaKpi().getRange(x.fila, 16, 1, 2).setValues([[yo.nombre, new Date()]]);
  var l = listaKpis(cred); l.ok = true;
  l.mensaje = '«' + x.nombre + '» borrado. Lo que ya anotaron se queda en la hoja.';
  return l;
}

/** Envía a los gerentes los que se indiquen (o todos los borradores y cambios). */
function enviarKpis(cred, ids) {
  var yo = exigeDirector(cred);
  var quiere = {};
  (ids || []).forEach(function (i) { quiere[String(i)] = true; });
  var todos = !ids || !ids.length, ahora = new Date(), n = 0, sucs = {};
  var h = hojaKpi();
  var misAreas = areasDe(yo);
  filasKpi().forEach(function (k) {
    if (k.estado === KPI_EST.DEL || !kpiEnAreas(k, misAreas)) return;
    if (k.estado !== KPI_EST.BOR) return;
    if (!todos && !quiere[k.id]) return;
    h.getRange(k.fila, 12).setValue(KPI_EST.ENV);
    h.getRange(k.fila, 18, 1, 3).setValues([[yo.nombre, ahora, ahora]]);
    SUC_KPI.concat([AREA_CONTA, AREA_MKT, AREA_TEC, AREA_COO]).forEach(function (s) { if (k.suc[s]) sucs[nombreArea(s)] = true; });
    n++;
  });
  if (!n) throw new Error('No hay nada nuevo por enviar.');
  var l = listaKpis(cred); l.ok = true;
  l.mensaje = (n === 1 ? 'Un indicador enviado' : n + ' indicadores enviados') + ' a ' +
    Object.keys(sucs).join(', ') + '. Les aparece en la campanita.';
  if (yo.rol === 'finanzas') l.mensaje = (n === 1 ? 'Una tarea enviada' : n + ' tareas enviadas') + ' al contador interno.';
  if (_KPI_MKT && yo.esAdmin) l.mensaje = (n === 1 ? 'Un indicador enviado' : n + ' indicadores enviados') + ' a la CMO. Le aparecen en la campanita y en Indicadores.';
  if (_KPI_COO && yo.esAdmin) l.mensaje = (n === 1 ? 'Un indicador enviado' : n + ' indicadores enviados') + ' al director operativo. Le aparecen en «Mis KPIs».';
  return l;
}

/** Agrega los recomendados escogidos, en borrador y para las tres sucursales. */
function agregarRecomendados(cred, claves) {
  var yo = exigeDirector(cred);
  var ya = {}, ahora = new Date(), n = 0, conta = yo.rol === 'finanzas', mkt = _KPI_MKT && yo.esAdmin, tec = _KPI_TEC && yo.esAdmin, coo = _KPI_COO && yo.esAdmin;
  filasKpi().forEach(function (k) { if (k.estado !== KPI_EST.DEL) ya[sinAcentoS(k.nombre)] = true; });
  var filas = [], lista = coo ? KPIS_COO_RECOMENDADOS : tec ? KPIS_TEC_RECOMENDADOS : mkt ? KPIS_MKT_RECOMENDADOS : conta ? KPIS_CONTA_RECOMENDADOS : KPIS_RECOMENDADOS;
  (claves || []).forEach(function (c) {
    var r = lista[Number(c)];
    if (!r || ya[sinAcentoS(r.n)]) return;
    ya[sinAcentoS(r.n)] = true;
    var v = { nombre: r.n, tipo: r.t, unidad: r.u, min: r.min, max: r.max, frec: r.f,
      suc: coo ? { direccion: true } : tec ? { tecnologia: true } : mkt ? { marketing: true } : conta ? { contabilidad: true } : { centro: true, almendras: true, parque: true }, instr: r.i, cat: r.c };
    if (tec) filas.push(filaKpi(nuevoIdKpi() + n, v, KPI_EST.ENV, yo, yo.nombre, ahora, ahora, yo.nombre, ahora, ahora));
    else filas.push(filaKpi(nuevoIdKpi() + n, v, KPI_EST.BOR, yo, yo.nombre, ahora, ahora, '', '', ''));
    n++;
  });
  if (!n) throw new Error('Escoja al menos uno que no tenga ya.');
  var h = hojaKpi();
  h.getRange(h.getLastRow() + 1, 1, filas.length, H_KPI.length).setValues(filas);
  var l = listaKpis(cred); l.ok = true;
  l.mensaje = (n === 1 ? 'Un indicador agregado' : n + ' indicadores agregados') +
    (tec ? '. Ya aparecen en «Anotar hoy»; ajuste los rangos si hace falta.' : ' en borrador. Revise los rangos y envíelos.');
  return l;
}

/* ── cuándo toca medir ── */
function lunesDe(iso) {
  var d = diaSemana(iso);
  return masDias(iso, d === 0 ? -6 : 1 - d);
}
/** Los periodos de un indicador en una fecha. `hora` (HH:mm) solo para hoy: el turno PM
 *  todavía no toca antes de que empiece. */
function periodosKpi(k, sid, fecha, hora) {
  if (k.frec === 'Cada turno') {
    var ps = [{ turno: 'AM', per: fecha + '|AM' }];
    if (!hora || hora >= INICIO_PM[sid]) ps.push({ turno: 'PM', per: fecha + '|PM' });
    return ps;
  }
  if (k.frec === 'Semanal') return [{ turno: '', per: 'S' + lunesDe(fecha) }];
  if (k.frec === 'Mensual') return [{ turno: '', per: 'M' + fecha.slice(0, 7) }];
  if (k.frec === 'Apertura') return [{ turno: 'Apertura', per: fecha + '|A' }];
  if (k.frec === 'Cierre') return [{ turno: 'Cierre', per: fecha + '|C' }];
  return [{ turno: '', per: fecha }];
}
function kpisDe(sid, alDia) {
  return filasKpi().filter(function (k) {
    return k.estado === KPI_EST.ENV && k.suc[sid] && (!alDia || k.desde.slice(0, 10) <= alDia);
  });
}
function registrosKpi(n, sinSello) {      // sinSello: quien no usa la hora exacta se ahorra formatear miles de fechas
  return ultimas(hojaKpiReg(), H_KPIREG.length, n || 5000).map(function (r) {
    return { en: sinSello ? '' : fmtSello(r[0]), fecha: fmtDia(r[1]), turno: String(r[2] || ''),
      sucursal: String(r[3] || ''), id: String(r[4] || ''), nombre: String(r[5] || ''),
      valor: r[6] === '' || r[6] == null ? '' : r[6], unidad: String(r[7] || ''),
      min: numONulo(r[8]), max: numONulo(r[9]), enRango: String(r[10] || ''),
      por: String(r[11] || ''), nota: String(r[12] || ''), per: String(r[13] || '') };
  }).filter(function (x) { return x.id; });
}
function valorTxt(x) {
  if (typeof x.valor === 'number') return conUnidad(x.valor, x.unidad);
  return String(x.valor);
}
function horaAhora() { return Utilities.formatDate(new Date(), ZONA, 'HH:mm'); }

/** Lo que le falta hoy a una sucursal. */
function kpiPendientes(sid, hoy, hora, regs) {
  var suc = nombreArea(sid), hechos = {};
  regs.forEach(function (x) { if (x.sucursal === suc) hechos[x.id + '|' + x.per] = x; });
  var pend = [];
  kpisDe(sid).forEach(function (k) {
    periodosKpi(k, sid, hoy, hora).forEach(function (p) {
      if (hechos[k.id + '|' + p.per]) return;
      pend.push({ id: k.id, nombre: k.nombre, tipo: k.tipo, unidad: k.unidad, min: k.min, max: k.max,
        rango: rangoTxt(k), instr: k.instr, cat: k.cat, frec: k.frec, turno: p.turno, per: p.per });
    });
  });
  return pend;
}

/** La página del gerente: lo que falta anotar hoy y lo que ya anotó. */
function pantallaKpi(cred) {
  permiteCmo();
  var yo = quien(cred);
  var sid = areaQueAnota(yo);
  if (!sid) throw new Error('Esto es para los gerentes y el contador interno.');
  var hoy = hoyISO(), hora = horaAhora(), regs = registrosKpi(3000);
  var suc = nombreArea(sid), semana = 'S' + lunesDe(hoy), mes = 'M' + hoy.slice(0, 7);
  var hechos = regs.filter(function (x) {
    return x.sucursal === suc && (x.fecha === hoy || x.per === semana || x.per === mes);
  }).map(function (x) {
    return { nombre: x.nombre, turno: x.turno, valor: valorTxt(x), enRango: x.enRango,
      por: x.por, hora: x.en.slice(11), nota: x.nota, semanal: x.per.charAt(0) === 'S',
      mensual: x.per.charAt(0) === 'M' };
  }).reverse();
  var tieneTurno = kpisDe(sid).some(function (k) { return k.frec === 'Cada turno'; });
  var mu = null;
  if (yo.rol === 'gerente') {
    try { mu = { categorias: MUE_CATS, pendientes: pendientesMuestreo(yo).filter(function (p) { return p.estado === 'Ahora'; }) }; }
    catch (e) { mu = null; }
  }
  return { sucursal: suc, conta: sid === AREA_CONTA, muestreo: mu, hoy: hoy, hora: hora, pendientes: kpiPendientes(sid, hoy, hora, regs),
    hechos: hechos, inicioPM: INICIO_PM[sid],
    faltaPM: tieneTurno && hora < INICIO_PM[sid] };
}

/** El gerente envía lo que midió. valores = [{id, per, valor, nota}] */
function registrarKpis(cred, valores) {
  permiteCmo();
  var yo = quien(cred);
  var sid = areaQueAnota(yo);
  if (!sid) throw new Error('Los indicadores los anota el gerente de la sucursal.');
  var suc = nombreArea(sid), hoy = hoyISO();
  var lock = LockService.getScriptLock();
  lock.waitLock(15000); _LEE = {};
  var n = 0, fuera = 0;
  try {
    var hora = horaAhora(), regs = registrosKpi(3000);
    var pend = {};
    kpiPendientes(sid, hoy, hora, regs).forEach(function (p) { pend[p.id + '|' + p.per] = p; });
    var filas = [], ahora = new Date();
    (valores || []).forEach(function (x) {
      var p = pend[String(x.id) + '|' + String(x.per)];
      if (!p) return;                         // ya se anotó o ya no toca
      var val = x.valor, ok = '';
      if (p.tipo === 'Número') {
        if (val === '' || val == null) return;
        val = Number(String(val).replace(',', '.'));
        if (!isFinite(val)) throw new Error('«' + p.nombre + '»: escriba un número.');
        ok = (p.min == null || val >= p.min) && (p.max == null || val <= p.max) ? 'Sí' : 'No';
      } else if (p.tipo === 'Sí / No') {
        if (val !== 'Sí' && val !== 'No') return;
        ok = val === 'Sí' ? 'Sí' : 'No';
      } else {
        val = String(val || '').trim().slice(0, 200);
        if (!val) return;
      }
      var nota = String(x.nota || '').trim().slice(0, 300);
      if (ok === 'No' && nota.length < 4)
        throw new Error(p.tipo === 'Sí / No' ? '«' + p.nombre + '» quedó en «No». Escriba qué pasó.'
          : '«' + p.nombre + '» está fuera de rango (' + p.rango + '). Escriba qué hizo.');
      if (ok === 'No') fuera++;
      filas.push([ahora, hoy, p.turno, suc, p.id, p.nombre, val, p.unidad,
        p.min == null ? '' : p.min, p.max == null ? '' : p.max, ok, yo.nombre, nota, p.per]);
      delete pend[p.id + '|' + p.per];
      n++;
    });
    if (!n) throw new Error('Anote al menos un valor.');
    var h = hojaKpiReg();
    var desde = h.getLastRow() + 1;
    h.getRange(desde, 1, filas.length, H_KPIREG.length).setValues(filas);
    h.getRange(desde, 2, filas.length, 1).setNumberFormat('@')
      .setValues(filas.map(function () { return [hoy]; }));
  } finally { lock.releaseLock(); }
  var r = pantallaKpi(cred); r.ok = true;
  r.mensaje = (sid === AREA_CONTA ? (n === 1 ? 'Una tarea marcada' : n + ' tareas marcadas')
    : (n === 1 ? 'Un indicador enviado' : n + ' indicadores enviados')) +
    (fuera ? ' · ' + fuera + (fuera === 1 ? ' con problema' : ' con problemas') +
      (sid === AREA_TEC ? '.' : ': le avisamos ' +
      (sid === AREA_CONTA ? 'al director financiero.' : (sid === AREA_MKT || sid === AREA_COO) ? 'al administrador.' : 'al director operativo.')) : '.');
  return r;
}

/** Para el director: cómo va cada sucursal en un día, y el cumplimiento de 7 días. */
function resultadosKpi(cred, fecha, suc) {
  return datosResultadosKpi(fecha, areasFiltro(cred, suc));
}
/** Las sucursales que puede ver quien pregunta, filtradas por la que escogió. */
function areasFiltro(cred, suc) {
  var yo = quien(cred);
  if (yo.rol === 'gerente') return [areaQueAnota(yo)];
  yo = exigeOperaciones(cred);
  var areas = areasDe(yo);
  return suc && areas.indexOf(suc) >= 0 ? [suc] : areas;
}
/** Resumen de un rango de fechas: por sucursal y por indicador, cuántos se anotaron,
 *  cuántos quedaron fuera de rango, el promedio y el cumplimiento de cada día. */
function resultadosKpiRango(cred, desde, hasta, suc) {
  var areas = areasFiltro(cred, suc), hoy = hoyISO();
  if (!esFecha(hasta) || hasta > hoy) hasta = hoy;
  if (!esFecha(desde) || desde > hasta) desde = masDias(hasta, -6);
  var recorte = false;
  if (diasEntre(desde, hasta) > 92) { desde = masDias(hasta, -92); recorte = true; }
  var todos = filasKpi().filter(function (k) { return k.estado === KPI_EST.ENV; });
  var regs = registrosKpi(20000), idx = {};
  regs.forEach(function (x) { idx[x.sucursal + '|' + x.id + '|' + x.per] = x; });
  var dias = [];
  for (var f = desde; f <= hasta; f = masDias(f, 1)) dias.push(f);
  var sucursales = areas.map(function (sid) {
    var suc = nombreArea(sid), por = {}, orden = [], vistos = {}, tot = { esperados: 0, hechos: 0, fuera: 0 }, serie = [];
    dias.forEach(function (f) {
      var hora = f === hoy ? horaAhora() : null, dd = { fecha: f, esperados: 0, hechos: 0, fuera: 0 };
      todos.forEach(function (k) {
        if (!k.suc[sid] || k.desde.slice(0, 10) > f) return;
        var s = por[k.id];
        if (!s) { s = por[k.id] = { id: k.id, nombre: k.nombre, rango: rangoTxt(k), frec: k.frec, cat: k.cat, tipo: k.tipo, unidad: k.unidad,
          esperados: 0, hechos: 0, fuera: 0, nums: [], fueraLista: [] }; orden.push(s); }
        periodosKpi(k, sid, f, hora).forEach(function (p) {
          if (vistos[k.id + '|' + p.per]) return;
          vistos[k.id + '|' + p.per] = true;
          s.esperados++; tot.esperados++; dd.esperados++;
          var x = idx[suc + '|' + k.id + '|' + p.per];
          if (!x) return;
          s.hechos++; tot.hechos++; dd.hechos++;
          if (typeof x.valor === 'number') s.nums.push(x.valor);
          if (x.enRango === 'No') { s.fuera++; tot.fuera++; dd.fuera++;
            s.fueraLista.push({ fecha: x.fecha, turno: p.turno, valor: valorTxt(x), nota: x.nota, por: x.por }); }
        });
      });
      dd.pct = dd.esperados ? Math.round(dd.hechos / dd.esperados * 100) : null;
      serie.push(dd);
    });
    var filas = orden.map(function (s) {
      var n = s.nums.length, prom = n ? s.nums.reduce(function (a, b) { return a + b; }, 0) / n : null;
      return { id: s.id, nombre: s.nombre, rango: s.rango, frec: s.frec, cat: s.cat, esperados: s.esperados, hechos: s.hechos, fuera: s.fuera,
        pct: s.esperados ? Math.round(s.hechos / s.esperados * 100) : null,
        promedio: prom == null ? '' : conUnidad(Math.round(prom * 10) / 10, s.unidad),
        minimo: n ? conUnidad(Math.min.apply(null, s.nums), s.unidad) : '', maximo: n ? conUnidad(Math.max.apply(null, s.nums), s.unidad) : '',
        fueraLista: s.fueraLista.slice(-8).reverse() };
    }).sort(function (a, b) { return (b.esperados - b.hechos + b.fuera * 2) - (a.esperados - a.hechos + a.fuera * 2); });
    return { id: sid, nombre: suc, filas: filas, esperados: tot.esperados, hechos: tot.hechos, fuera: tot.fuera,
      pct: tot.esperados ? Math.round(tot.hechos / tot.esperados * 100) : null, serie: serie };
  });
  return { desde: desde, hasta: hasta, dias: dias.length, recorte: recorte, sucursales: sucursales };
}
function datosResultadosKpi(fecha, areas) {
  areas = areas || SUC_KPI;
  var hoy = hoyISO();
  if (!esFecha(fecha) || fecha > hoy) fecha = hoy;
  var hora = fecha === hoy ? horaAhora() : null, regs = registrosKpi(6000);
  var idx = {};
  regs.forEach(function (x) { idx[x.sucursal + '|' + x.id + '|' + x.per] = x; });
  var sucursales = areas.map(function (sid) {
    var suc = nombreArea(sid), hechos = 0, esperados = 0, fuera = 0;
    var filas = kpisDe(sid, fecha).map(function (k) {
      return { id: k.id, nombre: k.nombre, rango: rangoTxt(k), frec: k.frec, cat: k.cat,
        celdas: periodosKpi(k, sid, fecha, hora).map(function (p) {
          var x = idx[suc + '|' + k.id + '|' + p.per];
          esperados++;
          if (!x) return { turno: p.turno, falta: true };
          hechos++; if (x.enRango === 'No') fuera++;
          return { turno: p.turno, valor: valorTxt(x), enRango: x.enRango, por: x.por,
            hora: x.en.slice(11), dia: x.fecha, nota: x.nota };
        }) };
    });
    // cumplimiento de los últimos 7 días (hasta la fecha escogida)
    var e7 = 0, h7 = 0, f7 = 0, semanas = {};
    for (var d = 0; d < 7; d++) {
      var f = masDias(fecha, -d);
      kpisDe(sid, f).forEach(function (k) {
        periodosKpi(k, sid, f, f === hoy ? horaAhora() : null).forEach(function (p) {
          if (k.frec === 'Semanal' || k.frec === 'Mensual') { if (semanas[k.id + p.per]) return; semanas[k.id + p.per] = true; }
          e7++;
          var x = idx[suc + '|' + k.id + '|' + p.per];
          if (x) { h7++; if (x.enRango === 'No') f7++; }
        });
      });
    }
    return { id: sid, nombre: suc, filas: filas, hechos: hechos, esperados: esperados, fuera: fuera,
      semana: { esperados: e7, hechos: h7, fuera: f7,
        pct: e7 ? Math.round(h7 / e7 * 100) : null } };
  });
  return { fecha: fecha, hoy: hoy, sucursales: sucursales };
}

/* ── avisos de la campanita para cierres de caja e indicadores ── */
function avisosCajaKpi(yo, out, hoy) {
  var ayer = masDias(hoy, -1), desde = masDias(hoy, -10);
  var etiqueta = function (x) {
    return x.sucursal + ' · corte ' + x.turno + ' del ' + dia(x.fecha).slice(0, 5); };
  if (yo.rol === 'operaciones') {
    filasPorConfirmar().forEach(function (x) {
      if (x.estado === EST_PEND)
        out.push({ id: 'val-' + x.id, tipo: 'Cierre por validar', cuando: x.sello,
          titulo: etiqueta(x) + ' · ' + dinero(x.monto), texto: 'Reportó ' + x.enviadoPor,
          abrir: { pc: true } });
      if (x.estado === EST_CONF && x.diferencia && x.revisadoEn.slice(0, 10) >= desde)
        out.push({ id: 'dif-caja-' + x.id, tipo: 'Diferencia de efectivo', cuando: x.revisadoEn,
          titulo: etiqueta(x) + ': ' + (x.diferencia > 0 ? 'sobran ' : 'faltan ') + dinero(Math.abs(x.diferencia)),
          texto: 'Lo recibió ' + x.revisadoPor, abrir: { pc: true } });
    });
    var regs = registrosKpi(3000);
    regs.forEach(function (x) {
      if (x.enRango !== 'No' || x.fecha < ayer || x.sucursal === nombreArea(AREA_CONTA) || x.sucursal === nombreArea(AREA_COO)) return;
      out.push({ id: 'kpif-' + x.sucursal + '-' + x.id + '-' + x.per, tipo: 'Fuera de rango', cuando: x.en,
        titulo: x.sucursal + ': ' + x.nombre + ' ' + valorTxt(x), texto: (x.nota ? 'Qué hizo: ' + x.nota + ' · ' : '') + x.por,
        abrir: { pag: 'kpi', kt: 'res', fecha: x.fecha } });
    });
    SUC_KPI.forEach(function (sid) {
      var falta = kpiFaltaDia(sid, ayer, regs);
      if (falta.length)
        out.push({ id: 'kpifalta-' + ayer + '-' + sid, tipo: 'Indicadores sin anotar', cuando: '',
          titulo: nombreUnidad(sid) + ' no anotó ' + falta.length + (falta.length === 1 ? ' indicador' : ' indicadores') + ' ayer',
          texto: falta.slice(0, 3).join(', ') + (falta.length > 3 ? '…' : ''),
          abrir: { pag: 'kpi', kt: 'res', fecha: ayer } });
    });
  }
  if (yo.recibe && !yo.esAdmin) {
    filasPorConfirmar().forEach(function (x) {
      if (x.estado !== EST_VAL) return;
      var v = {}; try { v = JSON.parse(x.datos); } catch (e) {}
      out.push({ id: 'rec-' + x.id, tipo: 'Efectivo por recibir', cuando: x.validadoEn,
        titulo: etiqueta(x) + ' · efectivo ' + dinero((v.efectivo || 0) - (v.caja || 0)),
        texto: 'Validó ' + x.validadoPor + '. Confirme cuando reciba el efectivo.', abrir: { pc: true } });
    });
  }
  if (yo.banco && !yo.esAdmin) {
    filasPorConfirmar().forEach(function (x) {
      if (x.estado !== EST_REC) return;
      var v = {}; try { v = JSON.parse(x.datos); } catch (e) {}
      out.push({ id: 'banco-' + x.id, tipo: 'Por confirmar en banco', cuando: x.recibidoEn,
        titulo: etiqueta(x) + ' · tarjeta reportada ' + dinero(v.tarjeta || 0),
        texto: 'Efectivo recibido por ' + x.recibidoPor + '. Anote la tarjeta del cierre NEONET.',
        abrir: { pc: true } });
    });
  }
  if (yo.rol === 'finanzas') {
    try { impuestosAvisos(out); } catch (e) {}
    var rc = registrosKpi(3000), nc = nombreArea(AREA_CONTA);
    rc.forEach(function (x) {
      if (x.enRango !== 'No' || x.fecha < ayer || x.sucursal !== nc) return;
      out.push({ id: 'kpif-' + x.sucursal + '-' + x.id + '-' + x.per, tipo: 'Tarea con problema', cuando: x.en,
        titulo: x.nombre + ': ' + valorTxt(x), texto: (x.nota ? x.nota + ' · ' : '') + x.por,
        abrir: { pag: 'conta', ct: 'res' } });
    });
    var fc = kpiFaltaDia(AREA_CONTA, ayer, rc);
    if (fc.length)
      out.push({ id: 'kpifalta-' + ayer + '-conta', tipo: 'Tareas sin hacer', cuando: '',
        titulo: 'El contador no marcó ' + fc.length + (fc.length === 1 ? ' tarea' : ' tareas') + ' ayer',
        texto: fc.slice(0, 3).join(', ') + (fc.length > 3 ? '…' : ''), abrir: { pag: 'conta', ct: 'res' } });
  }
  if (areaQueAnota(yo)) {
    var sid = areaQueAnota(yo), hora = horaAhora(), esConta = sid === AREA_CONTA;
    var pend = kpiPendientes(sid, hoy, hora, registrosKpi(3000));
    if (pend.length)
      out.push({ id: 'kpi-' + hoy + '-' + (hora >= INICIO_PM[sid] ? 'PM' : 'AM'), tipo: esConta ? 'Tareas' : 'Indicadores',
        cuando: '', titulo: 'Tiene ' + pend.length + (esConta ? (pend.length === 1 ? ' tarea' : ' tareas') + ' por hacer'
          : (pend.length === 1 ? ' indicador' : ' indicadores') + ' por anotar'),
        texto: pend.slice(0, 3).map(function (p) { return p.nombre + (p.turno ? ' (' + p.turno + ')' : ''); }).join(', ') +
          (pend.length > 3 ? '…' : ''), abrir: { pag: esConta ? 'conta' : 'kpi' } });
    // lo que el director envió junto (el mismo minuto) sale en un solo aviso
    var grupos = {};
    kpisDe(sid).forEach(function (k) {
      if (k.envEn.slice(0, 10) < desde) return;
      var g = k.envEn.slice(0, 16);
      (grupos[g] = grupos[g] || []).push(k);
    });
    Object.keys(grupos).forEach(function (g) {
      var ks = grupos[g], k = ks[0];
      var nuevo = k.desde.slice(0, 16) === k.envEn.slice(0, 16);
      out.push({ id: 'kpi-env-' + (ks.length > 1 ? '' : k.id + '-') + g,
        tipo: ks.length > 1 ? 'Indicadores nuevos' : nuevo ? 'Indicador nuevo' : 'Indicador cambiado',
        cuando: k.envEn, titulo: ks.length > 1 ? k.envPor + ' le envió ' + ks.length + ' indicadores' : k.nombre,
        texto: ks.length > 1 ? ks.slice(0, 3).map(function (x) { return x.nombre; }).join(', ') + (ks.length > 3 ? '…' : '')
          : (rangoTxt(k) ? 'Rango: ' + rangoTxt(k) + ' · ' : '') + k.frec.toLowerCase() + ' · envió ' + k.envPor,
        abrir: { pag: esConta ? 'conta' : 'kpi' } });
    });
  }
}
function kpiFaltaDia(sid, fecha, regs) {
  var suc = nombreArea(sid), hechos = {}, falta = [];
  regs.forEach(function (x) { if (x.sucursal === suc) hechos[x.id + '|' + x.per] = true; });
  kpisDe(sid, fecha).forEach(function (k) {
    if (k.frec === 'Semanal' || k.frec === 'Mensual') return;   // se cuentan completas
    periodosKpi(k, sid, fecha, null).forEach(function (p) {
      if (!hechos[k.id + '|' + p.per]) falta.push(k.nombre + (p.turno ? ' ' + p.turno : ''));
    });
  });
  return falta;
}

/* ════════════ IMPUESTOS: CALENDARIO Y PAGOS ════════════
 * El director financiero lleva aquí lo que se paga a la SAT y al IGSS. Cada
 * obligación tiene su frecuencia y su regla de vencimiento; la app arma los
 * periodos, dice cuál vence pronto o ya venció, y él anota cada pago con su boleta.
 * El administrador lo consulta. Las reglas son las del calendario tributario de la
 * SAT; si un vencimiento cae en día feriado, la SAT lo corre al siguiente día hábil
 * (la app solo corre sábados y domingos). */
var H_OBL = ['ID', 'Obligación', 'Formulario', 'Institución', 'Frecuencia', 'Vence', 'Activa', 'Nota'];
var H_PAGOIMP = ['ID', 'Registrado en', 'Registrado por', 'ID obligación', 'Obligación', 'Periodo',
  'Monto', 'Fecha de pago', 'No. de formulario o boleta', 'Nota', 'Anulado', 'Anulado por', 'Anulado en'];
var OBL_FREC = ['Mensual', 'Trimestral', 'Anual'];
/* Reglas: ultimo_habil = último día hábil del mes siguiente al periodo;
 * habiles:N = dentro de N días hábiles del mes siguiente; dia:N = día N del mes
 * siguiente; anual:MM-DD = esa fecha del año siguiente. */
/* CAIX, S.A. está en el régimen sobre las utilidades de actividades lucrativas (el
 * general). Los pagos de empleados (IGSS, retenciones) no se llevan aquí. */
var OBLIGACIONES_INICIALES = [
  ['IVA', 'IVA general mensual', 'SAT-2237', 'SAT', 'Mensual', 'ultimo_habil', 'Sí',
    'Declaración y pago del IVA del mes.'],
  ['ISRUT', 'ISR régimen sobre utilidades (pago trimestral)', 'SAT-1361', 'SAT', 'Trimestral', 'ultimo_habil', 'Sí',
    'Pago a cuenta del ISR del trimestre.'],
  ['ISO', 'ISO trimestral', 'SAT-1608', 'SAT', 'Trimestral', 'ultimo_habil', 'Sí',
    'Impuesto de solidaridad; se puede acreditar al ISR.'],
  ['ISRANUAL', 'ISR declaración anual', 'SAT-1411', 'SAT', 'Anual', 'anual:03-31', 'Sí',
    'Liquidación del año, con estados financieros.']
];
var OBLIGACIONES_QUITADAS = ['IGSS', 'ISRRET', 'ISRROSE'];
/** Deja las obligaciones como el régimen de CAIX: quita las de empleados y el ISR
 *  simplificado, y activa las del régimen sobre las utilidades. */
function obligacionesDelRegimen(ss) {
  var h = hojasImpuestos(ss).obl, filas = filasObligaciones();
  for (var i = filas.length - 1; i >= 0; i--)
    if (OBLIGACIONES_QUITADAS.indexOf(filas[i].id) >= 0) h.deleteRow(filas[i].fila);
  var hay = {};
  filasObligaciones().forEach(function (o) { hay[o.id] = o; });
  OBLIGACIONES_INICIALES.forEach(function (x) {
    if (hay[x[0]]) h.getRange(hay[x[0]].fila, 7).setValue('Sí');
    else h.appendRow(x);
  });
}

function hojasImpuestos(ss) {
  ss = ss || libro();
  var o = ss.getSheetByName('Obligaciones fiscales');
  if (!o) {
    o = hojaLimpia(ss, 'Obligaciones fiscales', H_OBL);
    o.getRange(2, 1, OBLIGACIONES_INICIALES.length, H_OBL.length).setValues(OBLIGACIONES_INICIALES);
    o.setColumnWidth(2, 300); o.setColumnWidth(8, 360);
  }
  var p = ss.getSheetByName('Pagos de impuestos');
  if (!p) {
    p = hojaLimpia(ss, 'Pagos de impuestos', H_PAGOIMP);
    p.getRange('B:B').setNumberFormat('dd/mm/yyyy hh:mm');
    p.getRange('F:F').setNumberFormat('@');
    p.getRange('G:G').setNumberFormat('"Q"#,##0.00');
    p.getRange('H:H').setNumberFormat('@');
    p.getRange('M:M').setNumberFormat('dd/mm/yyyy hh:mm');
    p.setColumnWidth(5, 280);
  }
  return { obl: o, pagos: p };
}
function filasObligaciones() {
  return leeTodo(hojasImpuestos().obl, H_OBL.length).map(function (r, i) {
    return { fila: i + 2, id: String(r[0] || '').trim(), nombre: String(r[1] || ''),
      formulario: String(r[2] || ''), institucion: String(r[3] || ''),
      frec: OBL_FREC.indexOf(String(r[4])) >= 0 ? String(r[4]) : 'Mensual',
      regla: String(r[5] || 'ultimo_habil').trim(), activa: siNo(r[6]), nota: String(r[7] || '') };
  }).filter(function (o) { return o.id && o.nombre; });
}
function filasPagosImp() {
  return leeTodo(hojasImpuestos().pagos, H_PAGOIMP.length).map(function (r, i) {
    return { fila: i + 2, id: String(r[0] || ''), en: fmtSello(r[1]), por: String(r[2] || ''),
      obl: String(r[3] || ''), nombre: String(r[4] || ''), periodo: String(r[5] || ''),
      monto: Number(r[6]) || 0, fecha: fmtDia(r[7]), boleta: String(r[8] || ''),
      nota: String(r[9] || ''), anulado: siNo(r[10]) };
  }).filter(function (p) { return p.id; });
}

/* ── fechas ── */
function esHabil(iso) { var d = diaSemana(iso); return d !== 0 && d !== 6; }
function habilDesde(iso) { while (!esHabil(iso)) iso = masDias(iso, 1); return iso; }
function finDeMes(ym) { var p = ym.split('-'); var d = new Date(+p[0], +p[1], 0); return ym + '-' + ('0' + d.getDate()).slice(-2); }
function mesSig(ym) { var p = ym.split('-'); var y = +p[0], m = +p[1] + 1; if (m > 12) { m = 1; y++; } return y + '-' + ('0' + m).slice(-2); }
function mesAnt(ym) { var p = ym.split('-'); var y = +p[0], m = +p[1] - 1; if (m < 1) { m = 12; y--; } return y + '-' + ('0' + m).slice(-2); }
/** El último mes de un periodo: 2026-09 → 2026-09; 2026-T3 → 2026-09; 2026 → 2026-12. */
function ultimoMesPeriodo(per) {
  if (/^\d{4}-\d{2}$/.test(per)) return per;
  var t = /^(\d{4})-T([1-4])$/.exec(per);
  if (t) return t[1] + '-' + ('0' + (Number(t[2]) * 3)).slice(-2);
  return per + '-12';
}
function periodoTxt(per) {
  var M = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto', 'septiembre',
    'octubre', 'noviembre', 'diciembre'];
  if (/^\d{4}-\d{2}$/.test(per)) return M[+per.slice(5) - 1] + ' ' + per.slice(0, 4);
  var t = /^(\d{4})-T([1-4])$/.exec(per);
  if (t) return ['', '1er', '2do', '3er', '4to'][+t[2]] + ' trimestre ' + t[1];
  return 'año ' + per;
}
function venceDe(o, per) {
  var sig = mesSig(ultimoMesPeriodo(per)), regla = o.regla;
  var m;
  if ((m = /^anual:(\d{2})-(\d{2})$/.exec(regla)))
    return habilDesde((Number(ultimoMesPeriodo(per).slice(0, 4)) + 1) + '-' + m[1] + '-' + m[2]);
  if ((m = /^dia:(\d{1,2})$/.exec(regla)))
    return habilDesde(sig + '-' + ('0' + Math.min(28, Number(m[1]))).slice(-2));
  if ((m = /^habiles:(\d{1,2})$/.exec(regla))) {
    var f = sig + '-01', n = 0;
    for (;;) { if (esHabil(f)) { n++; if (n >= Number(m[1])) return f; } f = masDias(f, 1); }
  }
  var fin = finDeMes(sig);                     // ultimo_habil
  while (!esHabil(fin)) fin = masDias(fin, -1);
  return fin;
}
function reglaTxt(o) {
  var m, cuando = o.frec === 'Trimestral' ? 'del mes siguiente al trimestre' : 'del mes siguiente';
  if ((m = /^anual:(\d{2})-(\d{2})$/.exec(o.regla))) return m[2] + '/' + m[1] + ' del año siguiente';
  if ((m = /^dia:(\d+)$/.exec(o.regla))) return 'día ' + m[1] + ' ' + cuando;
  if ((m = /^habiles:(\d+)$/.exec(o.regla))) return m[1] + ' días hábiles ' + cuando;
  return 'último día hábil ' + cuando;
}
/** Los periodos ya cerrados de una obligación, desde `desde` (aaaa-mm). El mes en
 *  curso todavía no se declara. */
function periodosObl(o, desde, hoy) {
  var out = [], ym = desde, actual = hoy.slice(0, 7);
  for (var k = 0; k < 60 && ym < actual; k++, ym = mesSig(ym)) {
    var mes = +ym.slice(5), per = null;
    if (o.frec === 'Mensual') per = ym;
    else if (o.frec === 'Trimestral' && mes % 3 === 0) per = ym.slice(0, 4) + '-T' + (mes / 3);
    else if (o.frec === 'Anual' && mes === 12) per = ym.slice(0, 4);
    if (per) out.push(per);
  }
  return out;
}

function exigeFinanzasVer(cred) {
  var yo = quien(cred);
  if (yo.rol !== 'finanzas' && !yo.esAdmin && yo.rol !== 'dueno') throw new Error('Esto lo lleva el director financiero.');
  return yo;
}
function exigeFinanzas(cred) {
  var yo = quien(cred);
  if (yo.rol !== 'finanzas') throw new Error('Los pagos de impuestos los anota el director financiero. Usted los puede consultar.');
  return yo;
}

function datosImpuestos() {
  var hoy = hoyISO(), aj = ajustesInv();
  var desde = aj.impuestosDesde || mesAnt(hoy.slice(0, 7));
  var pagos = filasPagosImp().filter(function (p) { return !p.anulado; });
  var pagado = {};
  pagos.forEach(function (p) { pagado[p.obl + '|' + p.periodo] = p; });
  var obls = filasObligaciones(), periodos = [];
  obls.forEach(function (o) {
    if (!o.activa) return;
    periodosObl(o, desde, hoy).forEach(function (per) {
      var vence = venceDe(o, per), p = pagado[o.id + '|' + per];
      var dias = diasEntre(hoy, vence);
      var estado = p ? 'Pagado' : dias < 0 ? 'Vencido' : dias <= 15 ? 'Por vencer' : 'Pendiente';
      periodos.push({ obl: o.id, nombre: o.nombre, formulario: o.formulario, institucion: o.institucion,
        periodo: per, periodoTxt: periodoTxt(per), vence: vence, dias: dias, estado: estado,
        pago: p ? { id: p.id, monto: p.monto, fecha: p.fecha, boleta: p.boleta, por: p.por, nota: p.nota,
          tarde: p.fecha > vence } : null });
    });
  });
  periodos.sort(function (a, b) { return a.vence < b.vence ? -1 : a.vence > b.vence ? 1 : 0; });
  var mes = hoy.slice(0, 7), pagadoMes = 0;
  pagos.forEach(function (p) { if (p.fecha.slice(0, 7) === mes) pagadoMes += p.monto; });
  var abiertos = periodos.filter(function (x) { return x.estado !== 'Pagado'; });
  return { hoy: hoy, desde: desde, periodos: periodos,
    obligaciones: obls.map(function (o) { return { id: o.id, nombre: o.nombre, formulario: o.formulario,
      institucion: o.institucion, frec: o.frec, regla: o.regla, reglaTxt: reglaTxt(o), activa: o.activa,
      nota: o.nota }; }),
    resumen: { vencidos: periodos.filter(function (x) { return x.estado === 'Vencido'; }).length,
      porVencer: periodos.filter(function (x) { return x.estado === 'Por vencer'; }).length,
      pagadosTarde: periodos.filter(function (x) { return x.pago && x.pago.tarde; }).length,
      pagadoMes: r2(pagadoMes), proximo: abiertos.length ? abiertos[0] : null,
      sinActivar: false },
    comision: aj.comision };
}
function listaImpuestos(cred) {
  var yo = exigeFinanzasVer(cred);
  var d = datosImpuestos();
  d.puedeCambiar = yo.rol === 'finanzas';
  return d;
}
function registrarPagoImpuesto(cred, p) {
  var yo = exigeFinanzas(cred);
  p = p || {};
  var o = null;
  filasObligaciones().forEach(function (x) { if (x.id === String(p.obl)) o = x; });
  if (!o) throw new Error('Escoja la obligación.');
  var per = String(p.periodo || '');
  if (!/^\d{4}(-\d{2}|-T[1-4])?$/.test(per)) throw new Error('Escoja el periodo.');
  var monto = r2(p.monto);
  if (!(monto >= 0) || String(p.monto).trim() === '') throw new Error('Escriba el monto pagado (0 si la declaración fue sin pago).');
  if (!esFecha(p.fecha) || p.fecha > hoyISO()) throw new Error('Escoja la fecha en que se pagó.');
  var boleta = String(p.boleta || '').trim().slice(0, 60);
  if (boleta.length < 3) throw new Error('Escriba el número de formulario o de boleta.');
  var lock = LockService.getScriptLock();
  lock.waitLock(15000); _LEE = {};
  try {
    var ya = filasPagosImp().filter(function (x) { return !x.anulado && x.obl === o.id && x.periodo === per; });
    if (ya.length) throw new Error('Ese periodo ya tiene pago anotado (' + ya[0].boleta + '). Si está mal, anúlelo primero.');
    var h = hojasImpuestos().pagos;
    var id = 'PI' + Utilities.formatDate(new Date(), ZONA, 'yyyyMMddHHmmss') + Math.floor(Math.random() * 90 + 10);
    h.appendRow([id, new Date(), yo.nombre, o.id, o.nombre, per, monto, p.fecha, boleta,
      String(p.nota || '').slice(0, 200), 'No', '', '']);
    h.getRange(h.getLastRow(), 6).setNumberFormat('@').setValue(per);
    h.getRange(h.getLastRow(), 8).setNumberFormat('@').setValue(p.fecha);
  } finally { lock.releaseLock(); }
  var d = listaImpuestos(cred); d.ok = true;
  d.mensaje = o.nombre + ' · ' + periodoTxt(per) + ': pago de ' + dinero(monto) + ' anotado.';
  return d;
}
function anularPagoImpuesto(cred, id, motivo) {
  var yo = exigeFinanzas(cred);
  if (String(motivo || '').trim().length < 4) throw new Error('Escriba por qué se anula.');
  var p = null;
  filasPagosImp().forEach(function (x) { if (x.id === String(id)) p = x; });
  if (!p || p.anulado) throw new Error('No se encontró ese pago.');
  var h = hojasImpuestos().pagos;
  h.getRange(p.fila, 10).setValue([p.nota, 'Anulado: ' + String(motivo).trim()].filter(Boolean).join(' · '));
  h.getRange(p.fila, 11, 1, 3).setValues([['Sí', yo.nombre, new Date()]]);
  var d = listaImpuestos(cred); d.ok = true; d.mensaje = 'Pago anulado.';
  return d;
}
/** Activar o desactivar, cambiar la regla, o agregar una obligación nueva. */
function guardarObligacion(cred, o) {
  exigeFinanzas(cred);
  o = o || {};
  var regla = String(o.regla || '').trim();
  if (!/^(ultimo_habil|habiles:\d{1,2}|dia:\d{1,2}|anual:\d{2}-\d{2})$/.test(regla)) throw new Error('Escoja cuándo vence.');
  var frec = OBL_FREC.indexOf(o.frec) >= 0 ? o.frec : 'Mensual';
  var nombre = String(o.nombre || '').trim().slice(0, 100);
  if (nombre.length < 3) throw new Error('Escriba el nombre de la obligación.');
  var h = hojasImpuestos().obl, x = null;
  filasObligaciones().forEach(function (y) { if (y.id === String(o.id || '')) x = y; });
  var fila = [x ? x.id : 'OB' + Utilities.formatDate(new Date(), ZONA, 'yyyyMMddHHmmss'), nombre,
    String(o.formulario || '').slice(0, 40), String(o.institucion || 'SAT').slice(0, 40), frec, regla,
    o.activa === false ? 'No' : 'Sí', String(o.nota || '').slice(0, 200)];
  if (x) h.getRange(x.fila, 1, 1, H_OBL.length).setValues([fila]);
  else h.appendRow(fila);
  var d = listaImpuestos(cred); d.ok = true; d.mensaje = x ? 'Obligación guardada.' : 'Obligación agregada.';
  return d;
}
function activarObligacion(cred, id, activa) {
  exigeFinanzas(cred);
  var x = null;
  filasObligaciones().forEach(function (y) { if (y.id === String(id)) x = y; });
  if (!x) throw new Error('No se encontró esa obligación.');
  hojasImpuestos().obl.getRange(x.fila, 7).setValue(activa ? 'Sí' : 'No');
  var d = listaImpuestos(cred); d.ok = true;
  d.mensaje = x.nombre + (activa ? ' activada.' : ' desactivada.');
  return d;
}
/** El % de NEONET: lo cambia el director financiero (o el administrador). Los cortes
 *  ya confirmados se quedan con el % que tenían. */
function guardarComision(cred, pct) {
  var yo = quien(cred);
  if (yo.rol !== 'finanzas' && !yo.esAdmin) throw new Error('Esto lo cambia el director financiero.');
  var n = Number(String(pct).replace(',', '.'));
  if (!isFinite(n) || n < 0 || n > 20) throw new Error('Escriba el % de comisión (por ejemplo 6.11).');
  ajustesQueFaltan();
  var h = hojaInv('Ajustes de inventario');
  leeTodo(h, 1).forEach(function (r, i) {
    if (String(r[0]).trim() === 'comision_tarjeta_pct') h.getRange(i + 2, 2).setNumberFormat('@').setValue(String(n));
  });
  _AJ = null;
  try { CacheService.getScriptCache().remove('ajustes_inv'); } catch (e) {}
  return { ok: true, comision: n, mensaje: 'Comisión de NEONET: ' + num(n) + ' %. Aplica a los cortes que se confirmen desde ahora.' };
}
function impuestosAvisos(out) {
  datosImpuestos().periodos.forEach(function (x) {
    if (x.estado === 'Vencido')
      out.push({ id: 'imp-venc-' + x.obl + '-' + x.periodo, tipo: 'Impuesto vencido', cuando: x.vence,
        titulo: x.nombre + ' · ' + x.periodoTxt, texto: 'Venció el ' + dia(x.vence) + '. Anote el pago cuando lo haga.',
        abrir: { pag: 'imp' } });
    else if (x.estado === 'Por vencer' && x.dias <= 7)
      out.push({ id: 'imp-pv-' + x.obl + '-' + x.periodo, tipo: 'Impuesto por vencer', cuando: '',
        titulo: x.nombre + ' · ' + x.periodoTxt, texto: 'Vence el ' + dia(x.vence) +
          (x.dias === 0 ? ' (hoy)' : ' (en ' + x.dias + (x.dias === 1 ? ' día)' : ' días)')),
        abrir: { pag: 'imp' } });
  });
}

/* ════════════ SEGUIMIENTO (administrador) ════════════
 * En una sola pantalla: cómo van los cierres de caja y cuánto tarda cada quien,
 * los indicadores que puso el director operativo y si las sucursales cumplen,
 * el inventario y los impuestos del director financiero. Solo consulta. */
/** El seguimiento lee ~30 hojas: se guarda 30 segundos para que abrir y cerrar la pestaña no lo repita. */
function seguimiento(cred) {
  exigeAdmin(cred);
  var cc = null; try { cc = CacheService.getScriptCache(); } catch (e) {}
  if (cc) { try { var t = cc.get('seguimiento_v1'); if (t) return JSON.parse(t); } catch (e) {} }
  var r = seguimientoCalc(cred);
  if (cc) { try { var s = JSON.stringify(r); if (s.length < 90000) cc.put('seguimiento_v1', s, 30); } catch (e) {} }
  return r;
}
function seguimientoCalc(cred) {
  exigeAdmin(cred);
  var hoy = hoyISO(), ahoraMs = Date.now(), out = { hoy: hoy };
  // ── cierres ──
  var pc = filasPorConfirmar(), porPaso = { validar: 0, recibir: 0, banco: 0 }, viejo = null;
  var h30 = masDias(hoy, -30), t = { validar: [], recibir: [], banco: [] };
  var horas = function (a, b) { var x = ms(aFechaHora(a)), y = ms(aFechaHora(b)); return x && y && y >= x ? (y - x) / 3600000 : null; };
  pc.forEach(function (x) {
    if (x.tipo !== 'ingreso') return;
    var p = pasoDe(x);
    if (p) {
      porPaso[p]++;
      var desde = p === 'validar' ? x.sello : p === 'recibir' ? x.validadoEn : x.recibidoEn;
      var hr = horas(desde, fmtSello(new Date()));
      if (hr != null && (!viejo || hr > viejo.horas))
        viejo = { horas: Math.round(hr), paso: p, quien: QUIEN_PASO[p], sucursal: x.sucursal, turno: x.turno, fecha: x.fecha };
    }
    if (x.sello.slice(0, 10) < h30) return;
    var a = horas(x.sello, x.validadoEn); if (a != null) t.validar.push(a);
    var b = horas(x.validadoEn, x.recibidoEn); if (b != null) t.recibir.push(b);
    var c = x.estado === EST_CONF && x.recibidoEn ? horas(x.recibidoEn, x.revisadoEn) : null; if (c != null) t.banco.push(c);
  });
  var prom = function (l) { if (!l.length) return null; var s = 0; l.forEach(function (x) { s += x; }); return Math.round(s / l.length * 10) / 10; };
  var ci = cierresCaja(8), faltan = 0;
  ci.fechas.forEach(function (f) {
    if (f >= hoy) return;
    ci.sucursales.forEach(function (s) { ['AM', 'PM'].forEach(function (tu) { if (!ci.celdas[s + '|' + f + '|' + tu]) faltan++; }); });
  });
  var mes = hoy.slice(0, 7), com = 0, tar = 0, difN = 0, difT = 0;
  ventasRango(mes + '-01', hoy).forEach(function (v) { com += v.k || 0; tar += v.j || 0; });
  pc.forEach(function (x) { if (x.diferencia && x.recibidoEn.slice(0, 7) === mes) { difN++; difT += x.diferencia; } });
  out.cierres = { porPaso: porPaso, masViejo: viejo, faltan7: faltan,
    tiempos: { validar: prom(t.validar), recibir: prom(t.recibir), banco: prom(t.banco) },
    comisionMes: r2(com), tarjetaMes: r2(tar), tasa: tasaComision(), diferenciasMes: { n: difN, total: r2(difT) } };
  // ── indicadores ──
  var ks = filasKpi().filter(function (k) { return k.estado !== KPI_EST.DEL; });
  var ult = null;
  ks.forEach(function (k) { [k.actEn, k.envEn].forEach(function (e) { if (e && (!ult || e > ult.en)) ult = { en: e, por: k.envEn === e ? k.envPor : k.actPor }; }); });
  var res = datosResultadosKpi(hoy);
  var fuera = registrosKpi(3000).filter(function (x) { return x.enRango === 'No' && x.fecha >= masDias(hoy, -6); });
  out.indicadores = { enviados: ks.filter(function (k) { return k.estado === KPI_EST.ENV; }).length,
    borradores: ks.filter(function (k) { return k.estado === KPI_EST.BOR; }).length, ultimo: ult,
    porCategoria: KPI_CATS.map(function (c) { return { cat: c, n: ks.filter(function (k) { return k.estado === KPI_EST.ENV && k.cat === c; }).length }; })
      .filter(function (x) { return x.n; }),
    sucursales: res.sucursales.map(function (s) { return { nombre: s.nombre, n: s.filas.length, hoy: s.hechos + '/' + s.esperados,
      pct: s.semana.pct, fuera: s.semana.fuera }; }),
    fuera: fuera.slice(-6).reverse().map(function (x) { return { sucursal: x.sucursal, nombre: x.nombre, valor: valorTxt(x),
      nota: x.nota, por: x.por, fecha: x.fecha }; }) };
  // ── inventario ──
  try {
    var aj = ajustesInv(), sols = filasSol();
    out.inventario = { unidades: UNIDADES.filter(function (u) { return u.conInv; }).map(function (u) {
        var est = estadoInventario(u.id), c = calendarioConteo(u.id, est.ultimo ? est.ultimo.fecha : '', aj, hoy);
        return { nombre: u.nombre, ultimo: est.ultimo ? est.ultimo.sello : '', atrasado: !!c.atrasado && !c.hechoHoy,
          tocaHoy: !!c.tocaHoy && !c.hechoHoy }; }),
      enviadas: sols.filter(function (s) { return s.estado === SOL.ENV; }).length,
      enCurso: sols.filter(function (s) { return s.estado === SOL.ACE || s.estado === SOL.CAM; }).length,
      porAclarar: lineasSol().filter(function (l) { return l.difEstado === DIF.PEND; }).length,
      ordenes: (function () { var n = {}; filasOC().forEach(function (o) { if (o.estado === 'Pedida') n[o.numero] = true; }); return Object.keys(n).length; })() };
  } catch (e) { out.inventario = null; }
  // ── impuestos ──
  try { var im = datosImpuestos(); out.impuestos = { resumen: im.resumen,
    lista: im.periodos.filter(function (x) { return x.estado !== 'Pendiente'; }).slice(-12) }; }
  catch (e) { out.impuestos = null; }
  try { var pr = pantallaProduccion(cred); out.produccion = { rendimiento: pr.rendimiento,
    lotes7: pr.lotes.filter(function (l) { return l.fecha >= masDias(hoy, -6); }).length }; } catch (e) { out.produccion = null; }
  try { out.compras = ordenesPago(cred).total; } catch (e) { out.compras = null; }
  try { out.contabilidad = datosIndicadoresConta(); } catch (e) { out.contabilidad = null; }
  try { out.muestreo = pantallaMuestreo(cred).estudios.slice(0, 3); } catch (e) { out.muestreo = null; }
  try { var ps = presupuestoMes(cred, hoy.slice(0, 7)); out.presupuesto = { total: ps.total, hay: ps.hay, diaHoy: ps.diaHoy, diasMes: ps.diasMes,
    pasados: [] };
    ps.unidades.forEach(function (u) { u.rubros.forEach(function (r) { if (r.presupuesto && r.pct >= 100) out.presupuesto.pasados.push(u.nombre + ' · ' + r.rubro + ' ' + r.pct + ' %'); }); });
  } catch (e) { out.presupuesto = null; }
  try { var fs = filasFondos(); out.fondos = { autorizar: fs.filter(function (x) { return x.estado === FONDO_ADM; }).length, aprobar: fs.filter(function (x) { return x.estado === FONDO_EST.PED; }).length,
    pagar: fs.filter(function (x) { return x.estado === FONDO_EST.APR; }).length,
    mes: r2(fs.filter(function (x) { return x.pagEn.slice(0, 7) === hoy.slice(0, 7); }).reduce(function (a, x) { return a + x.pagado; }, 0)) }; } catch (e) { out.fondos = null; }
  try { var pm = listaPermisos(cred); out.permisos = { porAprobar: pm.porAprobar.length, fuera: pm.fuera.filter(function (x) { return x.desde <= hoy; }).map(function (x) { return x.por; }) }; } catch (e) { out.permisos = null; }
  try { out.emergencias = emergAgrupadas().filter(function (e) { return e.en.slice(0, 7) === hoy.slice(0, 7); }).length; } catch (e) { out.emergencias = null; }
  try { out.personas = personasSeguimiento(cred, out); } catch (e) { out.personas = []; }
  try { out.consultas = listaConsultas(cred); } catch (e) { out.consultas = null; }
  try { out.mantenimiento = resumenMant(); } catch (e) { out.mantenimiento = null; }
  try { out.merma = resumenMerma(); } catch (e) { out.merma = null; }
  try { out.sugerencias = { sinLeer: filasSug().filter(function (x) { return !x.leidaPor; }).length }; } catch (e) { out.sugerencias = null; }
  try { out.descansos = { pedidos: filasDesc().filter(function (x) { return x.estado === DESC_EST.PED; }).length }; } catch (e) { out.descansos = null; }
  return out;
}
function aFechaHora(s) {
  if (!s) return null;
  if (s instanceof Date) return s;
  var m = /^(\d{4})-(\d{2})-(\d{2})(?: (\d{2}):(\d{2}))?/.exec(String(s));
  return m ? new Date(+m[1], +m[2] - 1, +m[3], +(m[4] || 0), +(m[5] || 0)) : null;
}

/* ════════════ PRODUCCIÓN (Paulino): SACAR DE BODEGA Y PORCIONAR ════════════
 * Paulino saca un insumo de la bodega (carne molida, queso en bloque…) y lo
 * convierte en porciones (tortas de 2 oz, bolsas de queso de 1 lb…). Cada receta
 * dice cuántas porciones deberían salir; cada lote anota cuántas salieron de verdad.
 * El insumo sale de la bodega, las porciones entran a la bodega (como cualquier
 * producto, y de ahí Samuel las manda a las sucursales), y el costo de cada porción
 * se calcula con lo que costó el insumo. */
var H_REC = ['ID', 'Insumo (código)', 'Insumo', 'Medida al pesar', 'Medida por unidad del insumo',
  'Porción (código)', 'Porción', 'Porciones esperadas por medida', 'Activa', 'Creada por', 'Creada en', 'Nota'];
var H_LOTE = ['No. de lote', 'Fecha y hora', 'Fecha', 'Receta', 'Insumo (código)', 'Insumo', 'Cantidad usada',
  'Medida', 'Salió de bodega (unidad del insumo)', 'Porción (código)', 'Porción', 'Porciones esperadas',
  'Porciones obtenidas', 'Rendimiento %', 'Costo del insumo', 'Costo por porción', 'Registrado por', 'Nota'];
MOV.SAL_PROD = 'Salida a producción';
MOV.ENT_PROD = 'Entrada de producción';

function hojasProduccion(ss) {
  ss = ss || libro();
  var r = ss.getSheetByName('Recetas de porcionado');
  if (!r) { r = hojaLimpia(ss, 'Recetas de porcionado', H_REC); r.getRange('K:K').setNumberFormat('dd/mm/yyyy hh:mm'); r.setColumnWidth(3, 240); r.setColumnWidth(7, 240); }
  var l = ss.getSheetByName('Producción lotes');
  if (!l) { l = hojaLimpia(ss, 'Producción lotes', H_LOTE); l.getRange('B:B').setNumberFormat('dd/mm/yyyy hh:mm'); l.getRange('C:C').setNumberFormat('@');
    l.getRange('O:P').setNumberFormat('"Q"#,##0.00'); l.setColumnWidth(6, 240); l.setColumnWidth(11, 240); }
  return { rec: r, lote: l };
}
function filasRecetas() {
  return leeTodo(hojasProduccion().rec, H_REC.length).map(function (r, i) {
    return { fila: i + 2, id: String(r[0] || ''), insumo: String(r[1] || ''), insumoNombre: String(r[2] || ''),
      medida: String(r[3] || '') || 'unidad', porUnidad: Number(r[4]) > 0 ? Number(r[4]) : 1,
      porcion: String(r[5] || ''), porcionNombre: String(r[6] || ''), esperadas: Number(r[7]) || 0,
      activa: String(r[8]).trim().toLowerCase() !== 'no', creadaPor: String(r[9] || ''), creadaEn: fmtSello(r[10]),
      nota: String(r[11] || '') };
  }).filter(function (x) { return x.id; });
}
function filasLotes() {
  return leeTodo(hojasProduccion().lote, H_LOTE.length).map(function (r) {
    return { numero: String(r[0] || ''), en: fmtSello(r[1]), fecha: fmtDia(r[2]), receta: String(r[3] || ''),
      insumo: String(r[4] || ''), insumoNombre: String(r[5] || ''), usada: Number(r[6]) || 0, medida: String(r[7] || ''),
      salio: Number(r[8]) || 0, porcion: String(r[9] || ''), porcionNombre: String(r[10] || ''),
      esperadas: Number(r[11]) || 0, obtenidas: Number(r[12]) || 0, rend: Number(r[13]) || 0,
      costoInsumo: Number(r[14]) || 0, costoPorcion: Number(r[15]) || 0, por: String(r[16] || ''), nota: String(r[17] || '') };
  }).filter(function (x) { return x.numero; });
}
function exigeProduccion(cred) {
  var yo = quien(cred);
  if (yo.rol !== 'produccion' && !yo.esAdmin) throw new Error('Esto lo hace producción.');
  return yo;
}
function veProduccion(yo) { return yo.rol === 'produccion' || veInv(yo) || yo.rol === 'bodega'; }

/** La pantalla de Paulino: recetas, lo que hay en bodega de cada insumo y los últimos lotes. */
function pantallaProduccion(cred) {
  var yo = quien(cred);
  if (!veProduccion(yo)) throw new Error('Su usuario no ve producción.');
  var ex = estadoInventario('bodega').existencia, idx = indiceCompras(), bod = {};
  filasBodega().forEach(function (b) { bod[b.codigo] = b; });
  var recetas = filasRecetas().filter(function (r) { return r.activa; }).map(function (r) {
    var b = bod[r.insumo] || {}, p = bod[r.porcion] || {};
    return { id: r.id, insumo: r.insumo, insumoNombre: b.nombre || r.insumoNombre, unidadInsumo: b.unidad || '',
      medida: r.medida, porUnidad: r.porUnidad, porcion: r.porcion, porcionNombre: p.nombre || r.porcionNombre,
      esperadas: r.esperadas, hay: r3(ex[r.insumo] || 0), hayTxt: b.codigo ? textoCantidad(ex[r.insumo] || 0, b) : '—',
      hayMedida: r3((ex[r.insumo] || 0) * r.porUnidad), porcionesHay: r3(ex[r.porcion] || 0),
      costoInsumo: b.codigo ? costoBodega(b, idx).costo : 0, nota: r.nota };
  });
  var lotes = filasLotes().slice(-60).reverse();
  var hace30 = masDias(hoyISO(), -30), porRec = {};
  lotes.forEach(function (l) {
    if (l.fecha < hace30) return;
    var k = l.receta; if (!porRec[k]) porRec[k] = { receta: k, nombre: l.porcionNombre, usada: 0, esp: 0, obt: 0, lotes: 0, medida: l.medida };
    porRec[k].usada += l.usada; porRec[k].esp += l.esperadas; porRec[k].obt += l.obtenidas; porRec[k].lotes++;
  });
  var eqL = filasLotesEtq();
  lotes.forEach(function (l) { var e = eqL[l.numero]; if (e) { l.vence = e.vence; l.vida = e.vida; l.recetaNombre = e.nombre; l.sobrante = e.sobrante;
    l.noUsado = e.ingredientes.filter(function (x, i) { return i > 0 && x.sobra > 0.05; }).map(function (x) { return { nombre: x.nombre, g: x.sobra }; }); } l.esMezcla = /^RM/.test(l.receta); });
  return { puede: yo.rol === 'produccion' || yo.esAdmin, recetas: recetas, lotes: lotes.slice(0, 40), mezclas: mezclasPantalla(ex, idx, bod), etiquetas: configEtiquetas(), hojasEtq: HOJAS_ETQ,
    rendimiento: Object.keys(porRec).map(function (k) { var x = porRec[k];
      return { receta: x.receta, nombre: x.nombre, lotes: x.lotes, usada: r3(x.usada), medida: x.medida,
        esperadas: r3(x.esp), obtenidas: r3(x.obt), pct: x.esp ? Math.round(x.obt / x.esp * 1000) / 10 : null }; }),
    insumos: productosInv().map(function (b) { var pu = pesoUnidadDe(b); return { codigo: b.codigo, nombre: b.nombre, unidad: b.unidad,
      hay: textoCantidad(ex[b.codigo] || 0, b), peso: pu ? pu.peso : 0, pesoEn: pu ? pu.medida : '', porPeso: b.medida === 'peso' }; }),
    medidas: ['libra', 'onza', 'kilo', 'gramo', 'unidad', 'litro'], hoy: hoyISO() };
}

/** Receta nueva o cambiada. Si la porción no existe en el catálogo de bodega, se crea. */
function guardarReceta(cred, r) {
  var yo = exigeProduccion(cred);
  r = r || {};
  var bod = {};
  filasBodega().forEach(function (b) { bod[b.codigo] = b; });
  var ins = bod[String(r.insumo || '')];
  if (!ins) throw new Error('Escoja el insumo que sale de bodega.');
  var medida = String(r.medida || 'libra').trim().slice(0, 20);
  var porUnidad = Number(String(r.porUnidad || '').replace(',', '.'));
  if (!(porUnidad > 0)) throw new Error('Diga cuántas ' + plur(2, medida) + ' trae cada ' + ins.unidad + ' de ' + ins.nombre + '.');
  var esperadas = Number(String(r.esperadas || '').replace(',', '.'));
  if (!(esperadas > 0)) throw new Error('Diga cuántas porciones deben salir de cada ' + medida + '.');
  var codigo = String(r.porcion || '').trim().toUpperCase();
  var nombreP = String(r.porcionNombre || '').trim().replace(/\s+/g, ' ').slice(0, 70);
  var lock = LockService.getScriptLock();
  lock.waitLock(15000); _LEE = {};
  try {
    if (!bod[codigo]) {
      if (nombreP.length < 3) throw new Error('Escriba el nombre de la porción (por ejemplo «Torta de hamburguesa 2 oz»).');
      if (!/^[A-Z0-9]{3,12}$/.test(codigo)) codigo = codigoNuevo(nombreP, bod);
      var h = hojaBodega(), ahora = new Date();
      h.appendRow([codigo, nombreP.toUpperCase(), 'Cocina', 'Producción', 'Sí', 'Sí', 'Sí', 'Insumos de cocina',
        'unidad', '', '', '', '', 1, '', 'Sí', ahora, yo.nombre, 'Congelado', 'Congelador', '', '', '', '', 'No', '']);
      _BOD = null;
    }
    var hr = hojasProduccion().rec, x = null;
    filasRecetas().forEach(function (y) { if (y.id === String(r.id || '')) x = y; });
    var fila = [x ? x.id : 'RP' + Utilities.formatDate(new Date(), ZONA, 'yyyyMMddHHmmss'), ins.codigo, ins.nombre,
      medida, porUnidad, codigo, bod[codigo] ? bod[codigo].nombre : nombreP.toUpperCase(), esperadas,
      r.activa === false ? 'No' : 'Sí', x ? x.creadaPor : yo.nombre, x ? x.creadaEn : new Date(), String(r.nota || '').slice(0, 200)];
    if (x) hr.getRange(x.fila, 1, 1, H_REC.length).setValues([fila]); else hr.appendRow(fila);
  } finally { lock.releaseLock(); }
  var d = pantallaProduccion(cred); d.ok = true; d.mensaje = 'Receta guardada.';
  return d;
}
function codigoNuevo(nombre, bod) {
  var base = sinAcentoS(nombre).toUpperCase().replace(/[^A-Z0-9 ]/g, '').split(' ').filter(Boolean)
    .map(function (w) { return w.slice(0, 3); }).join('').slice(0, 8) || 'PROD';
  var c = base, n = 1;
  while (bod[c]) c = base.slice(0, 7) + (n++);
  return c;
}

/** Un lote: se usaron `usada` (en la medida de la receta) y salieron `obtenidas` porciones. */
function registrarLote(cred, d) {
  var yo = exigeProduccion(cred);
  d = d || {};
  var rec = null;
  filasRecetas().forEach(function (y) { if (y.id === String(d.receta)) rec = y; });
  if (!rec || !rec.activa) throw new Error('Escoja la receta.');
  var usada = r3(Number(String(d.usada || '').replace(',', '.')));
  var obt = r3(Number(String(d.obtenidas || '').replace(',', '.')));
  if (!(usada > 0)) throw new Error('Escriba cuánto ' + rec.medida + ' usó.');
  if (!(obt >= 0) || String(d.obtenidas).trim() === '') throw new Error('Escriba cuántas porciones salieron.');
  var bod = {};
  filasBodega().forEach(function (b) { bod[b.codigo] = b; });
  var ins = bod[rec.insumo], por = bod[rec.porcion];
  if (!ins || !por) throw new Error('El insumo o la porción ya no están en el catálogo de bodega.');
  var salio = r3(usada / rec.porUnidad), esperadas = r3(usada * rec.esperadas);
  var rend = esperadas ? Math.round(obt / esperadas * 1000) / 10 : 0;
  var nota = String(d.nota || '').trim().slice(0, 200);
  if (rend < 90 && nota.length < 4) throw new Error('Salió el ' + rend + ' % de lo esperado (' + num(esperadas) +
    ' porciones). Escriba qué pasó (merma, grasa, producto dañado…).');
  var ex = estadoInventario('bodega').existencia;
  if ((ex[ins.codigo] || 0) < salio - 0.0005 && !d.sinExistencia)
    throw new Error('En bodega solo hay ' + textoCantidad(ex[ins.codigo] || 0, ins) + ' de ' + ins.nombre +
      ' y el lote usa ' + textoCantidad(salio, ins) + '. Revise la cantidad o pida que cuenten la bodega.');
  var costoU = costoBodega(ins, indiceCompras()).costo, costoIns = r2(costoU * salio);
  var costoPor = obt > 0 ? Math.round(costoIns / obt * 10000) / 10000 : 0;
  var lock = LockService.getScriptLock();
  lock.waitLock(15000); _LEE = {};
  var numero;
  try {
    var max = 0;
    leeTodo(hojasProduccion().lote, 1).forEach(function (r) { var m = /^PR-(\d+)$/.exec(String(r[0])); if (m) max = Math.max(max, +m[1]); });
    numero = 'PR-' + ('0000' + (max + 1)).slice(-4);
    var ahora = new Date(), hoy = hoyISO();
    hojasProduccion().lote.appendRow([numero, ahora, hoy, rec.id, ins.codigo, ins.nombre, usada, rec.medida, salio,
      por.codigo, por.nombre, esperadas, obt, rend, costoIns, costoPor, yo.nombre, nota]);
    hojasProduccion().lote.getRange(hojasProduccion().lote.getLastRow(), 3).setNumberFormat('@').setValue(hoy);
    escribeKardex([
      [ahora, nombreBodega(), ins.codigo, ins.nombre, MOV.SAL_PROD, -salio, numero, yo.nombre, num(usada) + ' ' + plur(usada, rec.medida)],
      [ahora, nombreBodega(), por.codigo, por.nombre, MOV.ENT_PROD, obt, numero, yo.nombre, 'Rendimiento ' + rend + ' %']]);
    // el costo de la porción queda como costo manual del producto de bodega
    if (costoPor > 0) { hojaBodega().getRange(por.fila, 15).setValue(costoPor); _BOD = null; }
  } finally { lock.releaseLock(); }
  var r = pantallaProduccion(cred); r.ok = true;
  r.mensaje = numero + ': ' + num(obt) + ' ' + plur(obt, 'porción') + ' de ' + por.nombre + ' (' + rend + ' % de lo esperado)' +
    (costoPor ? ' · ' + dinero(costoPor) + ' c/u' : '') + '.';
  return r;
}

/* ════════════ COMPRAS: COTIZACIÓN, PAGO Y RECEPCIÓN ════════════
 * Samuel hace la orden de compra → el proveedor manda la cotización (la anota el
 * contador o Samuel, precio por presentación) → el contador paga (se escribe sola la
 * factura en Gastos de la Bodega Central, con sus artículos) → Samuel recibe y entra
 * al inventario. Se puede recibir antes de pagar (crédito): queda «recibida sin pagar». */
var H_OCP = ['No. de orden', 'Proveedor', 'No. de cotización', 'Total cotizado', 'Cotizada en', 'Cotizada por',
  'Pagada en', 'Pagada por', 'Monto pagado', 'Forma de pago', 'Referencia', 'No. de factura', 'Fecha de pago',
  'ID gasto', 'Nota', 'Cotización (archivo)', 'Costo de transporte', 'Devuelta en', 'Devuelta por', 'Motivo de devolución', 'Factura impresa con NIT'];
var FORMAS_PAGO = ['Transferencia', 'Cheque', 'Depósito', 'Efectivo', 'Tarjeta'];
function hojaOCP(ss) {
  ss = ss || libro();
  var h = ss.getSheetByName('Órdenes de compra pagos');
  if (h && h.getLastColumn() < 17) { encabezaAlFinal(h, 17, ['Costo de transporte']); h.getRange('Q:Q').setNumberFormat('"Q"#,##0.00'); }
  if (h && h.getLastColumn() < 21) { encabezaAlFinal(h, 18, ['Devuelta en', 'Devuelta por', 'Motivo de devolución', 'Factura impresa con NIT']); h.getRange('R:R').setNumberFormat('dd/mm/yyyy hh:mm'); }
  if (!h) { h = hojaLimpia(ss, 'Órdenes de compra pagos', H_OCP); h.getRange('D:D').setNumberFormat('"Q"#,##0.00');
    h.getRange('E:E').setNumberFormat('dd/mm/yyyy hh:mm'); h.getRange('G:G').setNumberFormat('dd/mm/yyyy hh:mm');
    h.getRange('I:I').setNumberFormat('"Q"#,##0.00'); h.getRange('M:M').setNumberFormat('@'); }
  return h;
}
function filasOCP() {
  var o = {};
  leeTodo(hojaOCP(), H_OCP.length).forEach(function (r, i) {
    var n = String(r[0] || ''); if (!n) return;
    o[n] = { fila: i + 2, numero: n, proveedor: String(r[1] || ''), cotizacion: String(r[2] || ''),
      total: Number(r[3]) || 0, cotizadaEn: fmtSello(r[4]), cotizadaPor: String(r[5] || ''),
      pagadaEn: fmtSello(r[6]), pagadaPor: String(r[7] || ''), monto: Number(r[8]) || 0, forma: String(r[9] || ''),
      referencia: String(r[10] || ''), factura: String(r[11] || ''), fechaPago: fmtDia(r[12]), gasto: String(r[13] || ''),
      nota: String(r[14] || ''), archivo: String(r[15] || ''), transporte: Number(r[16]) || 0,
      devueltaEn: fmtSello(r[17]), devueltaPor: String(r[18] || ''), motivoDev: String(r[19] || ''), facturaNit: String(r[20] || '').trim().toLowerCase() === 'sí' };
  });
  return o;
}
function esContador(yo) { return yo.rol === 'registro' && !!yo.recibe; }
function veCompras(yo) { return yo.esAdmin || esContador(yo) || yo.rol === 'finanzas' || yo.rol === 'operaciones' || yo.rol === 'bodega' || yo.rol === 'dueno'; }

function ordenesAgrupadas() {
  var pagos = filasOCP(), por = {}, orden = [];
  filasOC().forEach(function (o) {
    if (!por[o.numero]) {
      var p = pagos[o.numero] || null;
      por[o.numero] = { numero: o.numero, proveedor: o.proveedor, creadaEn: o.creadaEn, creadaPor: o.creadaPor,
        estado: o.estado, recibidoEn: o.recibidoEn, recibidoPor: o.recibidoPor, lineas: [], total: 0,
        cotizacion: p ? p.cotizacion : '', cotizadaEn: p ? p.cotizadaEn : '', cotizadaPor: p ? p.cotizadaPor : '',
        totalCotizado: p ? p.total : 0, pagadaEn: p ? p.pagadaEn : '', pagadaPor: p ? p.pagadaPor : '',
        monto: p ? p.monto : 0, forma: p ? p.forma : '', referencia: p ? p.referencia : '', factura: p ? p.factura : '',
        fechaPago: p ? p.fechaPago : '', gasto: p ? p.gasto : '', archivo: p ? p.archivo : '', transporte: p ? p.transporte : 0,
        devueltaEn: p ? p.devueltaEn : '', devueltaPor: p ? p.devueltaPor : '', motivoDev: p ? p.motivoDev : '', facturaNit: p ? p.facturaNit : false };
      orden.push(por[o.numero]);
    }
    por[o.numero].lineas.push({ codigo: o.codigo, nombre: o.nombre, compra: o.compra, presentacion: o.presentacion,
      pedido: o.pedido, recibido: o.recibido, precio: o.precio, subtotal: o.subtotal, rinde: o.rinde });
  });
  orden.forEach(function (o) {
    o.pago = o.pagadaEn ? 'Pagada' : o.cotizadaEn ? 'Cotizada' : 'Por cotizar';
    o.deMenos = o.estado === 'Recibida' && o.lineas.some(function (l) { return l.recibido != null && l.recibido < l.pedido - 0.0005; });
  });
  return orden;
}
function pasoOC(o) {
  if (o.estado === 'Anulada') return 'Anulada';
  if (o.devueltaEn && o.pago !== 'Pagada' && o.estado !== 'Recibida') return 'Devuelta a bodega';
  if (o.pago === 'Pagada' && o.estado === 'Recibida') return 'Completa';
  if (o.estado === 'Recibida') return 'Recibida sin pagar';
  if (o.pago === 'Pagada') return 'Pagada, por recibir';
  if (o.pago === 'Cotizada') return 'Cotizada, por pagar';
  return 'Por cotizar';
}

/** Para el contador (y quien da seguimiento): las órdenes y el resumen del mes. */
function ordenesPago(cred, mes) {
  var yo = quien(cred);
  if (!veCompras(yo)) throw new Error('Su usuario no ve las compras.');
  var hoy = hoyISO();
  mes = /^\d{4}-\d{2}$/.test(String(mes || '')) ? mes : hoy.slice(0, 7);
  var desde = masDias(hoy, -90), puedeCot = yo.esAdmin || esContador(yo) || yo.rol === 'bodega', puedePag = yo.esAdmin || esContador(yo);
  var todas = ordenesAgrupadas();
  var lista = todas.filter(function (o) { var abierta = pasoOC(o) !== 'Completa' && o.estado !== 'Anulada';
    return abierta || o.creadaEn.slice(0, 10) >= desde; }).map(function (o) {
    o.paso = pasoOC(o);
    o.puedeCotizar = puedeCot && o.pago !== 'Pagada' && o.estado !== 'Anulada' && !o.devueltaEn;
    o.puedePagar = puedePag && o.pago === 'Cotizada' && o.estado !== 'Anulada' && !o.devueltaEn;
    o.puedeDevolver = (esContador(yo) || yo.esAdmin || yo.rol === 'finanzas') && o.pago !== 'Pagada' && o.estado === 'Pedida' && !o.devueltaEn;
    o.diasAbierta = diasEntre(o.creadaEn.slice(0, 10), hoy);
    return o; }).reverse();
  var prov = {}, tot = { pagado: 0, ordenes: 0, porPagar: 0, recibidoSinPagar: 0 };
  todas.forEach(function (o) {
    if (o.estado === 'Anulada') return;
    var p = prov[o.proveedor] || (prov[o.proveedor] = { proveedor: o.proveedor, pagado: 0, ordenes: 0, porPagar: 0, recibidoSinPagar: 0 });
    if (o.pagadaEn && o.fechaPago.slice(0, 7) === mes) { p.pagado += o.monto; p.ordenes++; tot.pagado += o.monto; tot.ordenes++; }
    if (o.pago === 'Cotizada') { p.porPagar += o.totalCotizado; tot.porPagar += o.totalCotizado; }
    if (o.estado === 'Recibida' && o.pago !== 'Pagada') { var t = o.totalCotizado; p.recibidoSinPagar += t; tot.recibidoSinPagar += t; }
  });
  var resumen = Object.keys(prov).map(function (k) { var p = prov[k];
    return { proveedor: p.proveedor, pagado: r2(p.pagado), ordenes: p.ordenes, porPagar: r2(p.porPagar), recibidoSinPagar: r2(p.recibidoSinPagar) }; })
    .filter(function (p) { return p.pagado || p.porPagar || p.recibidoSinPagar; })
    .sort(function (a, b) { return b.pagado - a.pagado; });
  // últimos 6 meses pagados, para la gráfica
  var meses = [], m = mes;
  for (var k = 0; k < 6; k++) { meses.unshift(m); m = mesAnt(m); }
  var porMes = meses.map(function (mm) { var t = 0; todas.forEach(function (o) { if (o.pagadaEn && o.fechaPago.slice(0, 7) === mm) t += o.monto; });
    return { mes: mm, pagado: r2(t) }; });
  return { mes: mes, ordenes: lista, resumen: resumen, total: { pagado: r2(tot.pagado), ordenes: tot.ordenes,
      porPagar: r2(tot.porPagar), recibidoSinPagar: r2(tot.recibidoSinPagar) }, porMes: porMes,
    formas: FORMAS_PAGO, puedeCotizar: puedeCot, puedePagar: puedePag, hoy: hoy };
}

/** La cotización del proveedor: precio de cada presentación. */
function cotizarOrden(cred, numero, d) {
  var yo = quien(cred);
  if (!(yo.esAdmin || esContador(yo) || yo.rol === 'bodega')) throw new Error('La cotización la anota el contador o la bodega.');
  d = d || {};
  var ls = filasOC().filter(function (o) { return o.numero === String(numero); });
  if (!ls.length) throw new Error('No se encontró la orden ' + numero + '.');
  if (ls[0].estado === 'Anulada') throw new Error('Esa orden está anulada.');
  var pagos = filasOCP();
  if (pagos[numero] && pagos[numero].pagadaEn) throw new Error('Esa orden ya se pagó; la cotización ya no se cambia.');
  if (pagos[numero] && pagos[numero].devueltaEn) throw new Error('Esa orden está devuelta a bodega: primero se corrige y se reenvía.');
  var precios = d.precios || {}, total = 0, h = hojaOC();
  var transp = Number(String(d.transporte == null ? '' : d.transporte).replace(/[Q,\s]/g, '')) || 0;
  if (transp < 0 || !isFinite(transp)) throw new Error('El costo de transporte no es válido.');
  transp = r2(transp);
  if (!(d.archivo && d.archivo.datos) && !(pagos[numero] && pagos[numero].archivo)) throw new Error('Adjunte la cotización del proveedor (PDF o foto): contabilidad la recibe en línea.');
  var arch0 = d.archivo && d.archivo.datos ? guardaArchivo(d.archivo, numero + ' cotización ' + ls[0].proveedor).url : '';   // fuera del candado
  var lock = LockService.getScriptLock();
  lock.waitLock(15000); _LEE = {};
  try {
    ls.forEach(function (o) {
      var p = Number(String(precios[o.codigo] == null ? '' : precios[o.codigo]).replace(/[Q,\s]/g, ''));
      if (!(p > 0)) throw new Error('Falta el precio de «' + o.nombre + '» (por ' + o.presentacion + ').');
      p = r2(p);
      var sub = r2(p * o.pedido); total += sub;
      h.getRange(o.fila, 16, 1, 2).setValues([[p, sub]]);
    });
    total = r2(total + transp);
    try { sincronizaPreciosBodega(ls, precios, numero); } catch (e) {}      // el precio nuevo queda en el producto de bodega
    var fila = [numero, ls[0].proveedor, String(d.cotizacion || '').trim().slice(0, 40), total, new Date(), yo.nombre,
      '', '', '', '', '', '', '', '', String(d.nota || '').slice(0, 200)];
    var hp = hojaOCP(), arch = pagos[numero] ? pagos[numero].archivo : '';
    if (arch0) arch = arch0;
    fila.push(arch); fila.push(transp);
    fila.push('', '', '', pagos[numero] && pagos[numero].facturaNit ? 'Sí' : '');      // sin devolución pendiente
    if (pagos[numero]) hp.getRange(pagos[numero].fila, 1, 1, H_OCP.length).setValues([fila]); else hp.appendRow(fila);
  } finally { lock.releaseLock(); }
  var r = ordenesPago(cred); r.ok = true; r.mensaje = numero + ' cotizada: ' + dinero(total) + '. Ahora el contador la paga.';
  return r;
}

/** Lo que cotizó el proveedor es el precio más fresco: pasa al producto de bodega ligado a esa compra. */
function sincronizaPreciosBodega(lineas, precios, numero) {
  var h = hojaBodega(), ahora = new Date(), hechos = 0;
  var bod = filasBodega();
  lineas.forEach(function (o) {
    var p = Number(String(precios[o.codigo] == null ? '' : precios[o.codigo]).replace(/[Q,\s]/g, ''));
    if (!(p > 0)) return;
    bod.forEach(function (b) {
      var misma = b.codigo === o.codigo;
      if (!misma || r2(p) === b.precioPres) return;
      h.getRange(b.fila, 35, 1, 3).setValues([[r2(p), ahora, 'Cotización ' + numero]]);
      hechos++;
    });
  });
  if (hechos) _BOD = null;
}

/** El contador paga. La factura entra sola a Gastos de la Bodega Central. */
function pagarOrden(cred, numero, d) {
  var yo = quien(cred);
  if (!(yo.esAdmin || esContador(yo))) throw new Error('Los pagos a proveedores los hace el contador.');
  d = d || {};
  var o = null;
  ordenesAgrupadas().forEach(function (x) { if (x.numero === String(numero)) o = x; });
  if (!o) throw new Error('No se encontró la orden ' + numero + '.');
  if (o.estado === 'Anulada') throw new Error('Esa orden está anulada.');
  if (o.pago === 'Pagada') throw new Error('Esa orden ya la pagó ' + o.pagadaPor + '.');
  if (o.pago !== 'Cotizada') throw new Error('Primero anote la cotización del proveedor.');
  if (o.devueltaEn) throw new Error('Esa orden está devuelta a bodega.');
  if (d.facturaNit !== true) throw new Error('Confirme que ya tiene la factura impresa con el NIT de CAIX, S.A.: sin ella no se paga.');
  var monto = r2(String(d.monto == null ? '' : d.monto).replace(/[Q,\s]/g, ''));
  if (!(monto > 0)) monto = o.totalCotizado;
  if (Math.abs(monto - o.totalCotizado) > 1)
    throw new Error('Lo pagado (' + dinero(monto) + ') no cuadra con la cotización (' + dinero(o.totalCotizado) +
      '). Si el proveedor cambió precios, corrija primero la cotización.');
  var fecha = esFecha(d.fecha) ? d.fecha : hoyISO();
  if (fecha > hoyISO()) throw new Error('La fecha de pago no puede ser de un día que no ha llegado.');
  var forma = FORMAS_PAGO.indexOf(d.forma) >= 0 ? d.forma : 'Transferencia';
  var ref = String(d.referencia || '').trim().slice(0, 40);
  if (forma !== 'Efectivo' && ref.length < 3) throw new Error('Escriba el número de ' + (forma === 'Cheque' ? 'cheque' : 'referencia') + '.');
  var factura = String(d.factura || '').trim().slice(0, 30) || numero;
  var cat = {};
  filasProductos().forEach(function (p) { cat[(p.nombre + '§' + p.proveedor).toLowerCase()] = p; if (!cat[p.nombre.toLowerCase()]) cat[p.nombre.toLowerCase()] = p; });
  var lineas = o.lineas.map(function (l) {
    var p = cat[(l.compra + '§' + o.proveedor).toLowerCase()] || cat[String(l.compra).toLowerCase()] || {};
    return { producto: l.compra || l.nombre, cantidad: l.pedido, medida: l.presentacion, precio: l.precio,
      categoria: p.categoria || 'Insumos de cocina', contenido: p.contenido || 0, base: p.base || '' };
  });
  if (o.transporte > 0) lineas.push({ producto: 'Costo de transporte', cantidad: 1, medida: 'servicio', precio: o.transporte,
    categoria: 'Combustible y transporte', contenido: 0, base: '' });
  var v = preparaGasto({ unidad: 'bodega', fechaPago: fecha, proveedor: o.proveedor, factura: factura, lineas: lineas,
    nota: 'Orden ' + numero + ' · ' + forma + (ref ? ' ' + ref : '') });
  var lock = LockService.getScriptLock();
  var hp = hojaOCP(), pagos = filasOCP(), fila = pagos[numero].fila;
  // se marca antes de escribir el gasto: si alguien aprieta dos veces, no se paga dos veces
  hp.getRange(fila, 7, 1, 2).setValues([[new Date(), yo.nombre]]);
  hp.getRange(fila, 21).setValue('Sí');                    // factura impresa con NIT de CAIX recibida
  var r;
  try {
    r = escribeGasto(v, yo.nombre, new Date(), null);
    try { aseguraProveedor(o.proveedor, yo.nombre); } catch (e) {}
    try { recuerdaPrecios(v.lineas); } catch (e) {}
  } catch (e) { hp.getRange(fila, 7, 1, 2).setValues([['', '']]); throw e; }
  hp.getRange(fila, 9, 1, 7).setValues([[monto, forma, ref, factura, fecha, r.nuevo.id, pagos[numero].nota]]);
  hp.getRange(fila, 13).setNumberFormat('@').setValue(fecha);
  var out = ordenesPago(cred); out.ok = true; out.nuevo = r.nuevo;
  out.mensaje = numero + ' pagada: ' + dinero(monto) + ' a ' + o.proveedor + '. La factura quedó en Gastos de ' + nombreBodega() +
    (o.estado === 'Recibida' ? '.' : '; Samuel la recibe cuando llegue.');
  return out;
}

/* ════════════ CONTABILIDAD: CÓMO VA EL CONTADOR ════════════
 * Lo que la app ya sabe del trabajo del contador, sin que nadie lo anote, y
 * sugerencias según lo que está pasando. Lo ven el contador y el director financiero. */
function indicadoresConta(cred) {
  var yo = quien(cred);
  if (!(yo.esAdmin || esContador(yo) || yo.rol === 'finanzas')) throw new Error('Esto es para contabilidad.');
  return datosIndicadoresConta();
}
function datosIndicadoresConta() {
  var hoy = hoyISO(), h30 = masDias(hoy, -30), ahoraS = fmtSello(new Date());
  var pc = filasPorConfirmar(), tiempos = [], espera = [], difN = 0, difT = 0;
  pc.forEach(function (x) {
    if (x.tipo !== 'ingreso') return;
    if (x.estado === EST_VAL) espera.push({ x: x, h: horasEntre(x.validadoEn, ahoraS) });
    if (x.recibidoEn && x.validadoEn && x.recibidoEn.slice(0, 10) >= h30) {
      var t = horasEntre(x.validadoEn, x.recibidoEn); if (t != null) tiempos.push(t);
      if (x.diferencia) { difN++; difT += x.diferencia; }
    }
  });
  var prom = tiempos.length ? Math.round(tiempos.reduce(function (a, b) { return a + b; }, 0) / tiempos.length * 10) / 10 : null;
  var viejo = espera.sort(function (a, b) { return (b.h || 0) - (a.h || 0); })[0] || null;
  var ocs = ordenesAgrupadas().filter(function (o) { return o.estado !== 'Anulada'; });
  var porPagar = ocs.filter(function (o) { return o.pago === 'Cotizada'; });
  var porCotizar = ocs.filter(function (o) { return o.pago === 'Por cotizar' && o.estado !== 'Recibida'; });
  var sinPagar = ocs.filter(function (o) { return o.estado === 'Recibida' && o.pago !== 'Pagada'; });
  var deMenos = ocs.filter(function (o) { return o.deMenos && o.pago === 'Pagada' && o.recibidoEn.slice(0, 10) >= h30; });
  var mes = hoy.slice(0, 7), pagadoMes = 0;
  ocs.forEach(function (o) { if (o.pagadaEn && o.fechaPago.slice(0, 7) === mes) pagadoMes += o.monto; });
  var tareas = null;
  try { var rs = datosResultadosKpi(hoy, [AREA_CONTA]).sucursales[0]; tareas = rs.filas.length ? rs.semana : null; } catch (e) {}
  var imp = null;
  try { imp = datosImpuestos().resumen; } catch (e) {}
  var sug = [];
  if (viejo && viejo.h > 24) sug.push('Hay un cierre validado esperando que reciba el efectivo desde hace ' + Math.round(viejo.h) +
    ' h (' + viejo.x.sucursal + ' ' + viejo.x.turno + ' del ' + dia(viejo.x.fecha).slice(0, 5) + '). El efectivo debe recibirse el mismo día.');
  if (difN) sug.push(difN + (difN === 1 ? ' cierre llegó' : ' cierres llegaron') + ' con diferencia de efectivo en 30 días (' +
    dinero(difT) + '). Revíselo con el director operativo y el gerente.');
  porPagar.forEach(function (o) { var d = diasEntre(o.cotizadaEn.slice(0, 10), hoy);
    if (d >= 3) sug.push(o.numero + ' de ' + o.proveedor + ' está cotizada hace ' + d + ' días y no se ha pagado (' + dinero(o.totalCotizado) + ').'); });
  sinPagar.forEach(function (o) { sug.push(o.numero + ' de ' + o.proveedor + ' ya se recibió en bodega y no está pagada' +
    (o.totalCotizado ? ' (' + dinero(o.totalCotizado) + ')' : ' ni cotizada') + '.'); });
  deMenos.forEach(function (o) { sug.push(o.numero + ' se pagó completa pero llegó de menos a bodega. Pida nota de crédito a ' + o.proveedor + '.'); });
  if (imp && imp.vencidos) sug.push('Hay ' + imp.vencidos + (imp.vencidos === 1 ? ' impuesto vencido' : ' impuestos vencidos') + ' sin pago anotado.');
  if (imp && imp.proximo && imp.proximo.dias >= 0 && imp.proximo.dias <= 7) sug.push(imp.proximo.nombre + ' (' + imp.proximo.periodoTxt +
    ') vence el ' + dia(imp.proximo.vence) + '.');
  var fijas = ['No pague a un proveedor sin orden de compra y cotización: así cada pago tiene quién lo pidió y a qué precio.',
    'Deposite el efectivo el mismo día que lo recibe y guarde la boleta.',
    'Verifique cada factura de compra en el portal FEL de la SAT antes de pagarla: una factura inválida no da crédito de IVA.',
    'Confirme el banco solo con el estado de cuenta en la mano: la tarjeta líquida del cierre NEONET debe cuadrar con el depósito.',
    'Cierre el mes antes del día 10 con el PDF de Reportes.'];
  return { hoy: hoy,
    efectivo: { promedioHoras: prom, esperando: espera.length, masViejoHoras: viejo ? Math.round(viejo.h || 0) : null,
      diferencias: { n: difN, total: r2(difT) } },
    compras: { porCotizar: porCotizar.length, porPagar: porPagar.length, porPagarQ: r2(porPagar.reduce(function (a, o) { return a + o.totalCotizado; }, 0)),
      recibidasSinPagar: sinPagar.length, pagadoMes: r2(pagadoMes), deMenos: deMenos.length },
    tareas: tareas, impuestos: imp, sugerencias: sug, buenasPracticas: fijas };
}
function horasEntre(a, b) {
  var x = aFechaHora(a), y = aFechaHora(b);
  if (!x || !y) return null;
  return Math.max(0, (y - x) / 3600000);
}

/* ════════════ TABLERO DE INDICADORES (gráficas del director) ════════════ */
function tableroKpi(cred, dias) {
  var yo = exigeOperaciones(cred), areas = areasDe(yo);
  if (yo.esAdmin) areas = SUC_KPI.slice();
  return tableroKpiAreas(areas, dias);
}
function tableroKpiAreas(areas, dias) {
  dias = Math.min(31, Math.max(7, Number(dias) || 14));
  var hoy = hoyISO(), fechas = [];
  for (var k = dias - 1; k >= 0; k--) fechas.push(masDias(hoy, -k));
  // la lista de indicadores se lee UNA vez (antes se releía y reformateaba por cada área y cada día: 6 × 14 = 84 veces)
  var todosK = filasKpi();
  var kpisDe = function (sid, alDia) {
    return todosK.filter(function (k) { return k.estado === KPI_EST.ENV && k.suc[sid] && (!alDia || k.desde.slice(0, 10) <= alDia); });
  };
  var regs = registrosKpi(8000, true), idx = {};
  regs.forEach(function (x) { idx[x.sucursal + '|' + x.id + '|' + x.per] = x; });
  var hora = horaAhora();
  var cumplimiento = {}, checklists = {}, numericos = {}, fallas = {};
  areas.forEach(function (sid) {
    var suc = nombreArea(sid);
    cumplimiento[sid] = []; checklists[sid] = { Apertura: [], Cierre: [] };
    fechas.forEach(function (f) {
      var esp = 0, hechos = 0, ch = { Apertura: { t: 0, si: 0, hecho: 0 }, Cierre: { t: 0, si: 0, hecho: 0 } };
      kpisDe(sid, f).forEach(function (k) {
        if (k.frec === 'Semanal' || k.frec === 'Mensual') return;
        periodosKpi(k, sid, f, f === hoy ? hora : null).forEach(function (p) {
          var x = idx[suc + '|' + k.id + '|' + p.per];
          esp++; if (x) hechos++;
          if (k.frec === 'Apertura' || k.frec === 'Cierre') { var c = ch[k.frec]; c.t++; if (x) { c.hecho++; if (x.enRango === 'Sí') c.si++; } }
        });
      });
      // hoy todavía no termina: no se grafica para no parecer una caída
      cumplimiento[sid].push(esp && f !== hoy ? Math.round(hechos / esp * 100) : null);
      ['Apertura', 'Cierre'].forEach(function (m) { var c = ch[m];
        checklists[sid][m].push(!c.t ? null : { t: c.t, hecho: c.hecho, si: c.si,
          e: c.hecho === 0 ? (f === hoy ? 'hoy' : 'falta') : c.si === c.t ? 'ok' : 'parcial' }); });
    });
  });
  // tendencias de los indicadores numéricos y los que más fallan (30 días)
  var ks = todosK.filter(function (k) { return k.estado === KPI_EST.ENV && kpiEnAreas(k, areas); });
  var desde = fechas[0], h30 = masDias(hoy, -29);
  ks.forEach(function (k) {
    if (k.tipo !== 'Número') return;
    var serie = {};
    areas.forEach(function (sid) { if (!k.suc[sid]) return;
      var suc = nombreArea(sid), porDia = {};
      regs.forEach(function (x) { if (x.id !== k.id || x.sucursal !== suc || x.fecha < desde) return;
        var v = Number(x.valor); if (!isFinite(v)) return;
        (porDia[x.fecha] = porDia[x.fecha] || []).push(v); });
      serie[sid] = fechas.map(function (f) { var l = porDia[f]; if (!l) return null;
        return Math.round(l.reduce(function (a, b) { return a + b; }, 0) / l.length * 100) / 100; });
    });
    numericos[k.id] = { id: k.id, nombre: k.nombre, unidad: k.unidad, min: k.min, max: k.max, frec: k.frec, serie: serie };
  });
  var nombresArea = areas.map(nombreArea);
  regs.forEach(function (x) {
    if (x.enRango !== 'No' || x.fecha < h30 || nombresArea.indexOf(x.sucursal) < 0) return;
    var f = fallas[x.id] || (fallas[x.id] = { nombre: x.nombre, n: 0, por: {} });
    f.n++; f.por[x.sucursal] = (f.por[x.sucursal] || 0) + 1;
  });
  return { fechas: fechas, areas: areas.map(function (a) { return { id: a, nombre: nombreArea(a) }; }),
    cumplimiento: cumplimiento, checklists: checklists,
    numericos: Object.keys(numericos).map(function (k) { return numericos[k]; }),
    fallas: Object.keys(fallas).map(function (k) { return fallas[k]; }).sort(function (a, b) { return b.n - a.n; }).slice(0, 6) };
}

/* ════════════ MUESTREO DE TRABAJO ════════════
 * Qué es: en vez de cronometrar a cada persona todo el día, se observa el local a
 * horas al azar y se anota cuántas personas están haciendo qué (preparando,
 * atendiendo, limpiando, esperando sin tarea…). Con suficientes observaciones, el
 * porcentaje de veces que alguien estaba en cada actividad es el porcentaje del
 * tiempo que el personal dedica a ella. Sirve para decidir cuánta gente poner en
 * cada turno y cada día. */
var H_MUE = ['ID', 'Nombre', 'Sucursal', 'Desde', 'Hasta', 'Observaciones por día', 'Categorías', 'Estado',
  'Creado por', 'Creado en', 'Nota'];
var H_MOB = ['ID estudio', 'Registrado en', 'Fecha', 'Hora programada', 'Sucursal', 'Observado por', 'Conteos',
  'Total personas', 'Nota'];
var MUE_CATS = [
  { k: 'prep', n: 'Preparando comida', t: 'P' }, { k: 'aten', n: 'Atendiendo o cobrando', t: 'P' },
  { k: 'emp', n: 'Empacando o entregando', t: 'P' }, { k: 'limp', n: 'Limpiando', t: 'A' },
  { k: 'rec', n: 'Recibiendo o acomodando producto', t: 'A' }, { k: 'esp', n: 'Esperando sin tarea', t: 'O' },
  { k: 'fuera', n: 'Fuera del puesto o en descanso', t: 'O' }, { k: 'tel', n: 'En el teléfono (personal)', t: 'O' }];
var MUE_HORARIO = { centro: ['08:00', '20:30'], almendras: ['09:00', '20:30'], parque: ['09:00', '20:30'] };
function hojasMuestreo(ss) {
  ss = ss || libro();
  var e = ss.getSheetByName('Muestreo estudios');
  if (!e) { e = hojaLimpia(ss, 'Muestreo estudios', H_MUE); e.getRange('D:E').setNumberFormat('@'); e.getRange('J:J').setNumberFormat('dd/mm/yyyy hh:mm'); }
  var o = ss.getSheetByName('Muestreo observaciones');
  if (!o) { o = hojaLimpia(ss, 'Muestreo observaciones', H_MOB); o.getRange('B:B').setNumberFormat('dd/mm/yyyy hh:mm'); o.getRange('C:D').setNumberFormat('@'); }
  return { est: e, obs: o };
}
function filasEstudios() {
  return leeTodo(hojasMuestreo().est, H_MUE.length).map(function (r, i) {
    var cats = MUE_CATS; try { var c = JSON.parse(String(r[6] || '')); if (c && c.length) cats = c; } catch (e) {}
    return { fila: i + 2, id: String(r[0] || ''), nombre: String(r[1] || ''), sid: String(r[2] || ''), desde: fmtDia(r[3]),
      hasta: fmtDia(r[4]), porDia: Number(r[5]) || 10, cats: cats, estado: String(r[7] || 'Activo'), por: String(r[8] || ''),
      en: fmtSello(r[9]), nota: String(r[10] || '') };
  }).filter(function (x) { return x.id; });
}
function filasObs() {
  return leeTodo(hojasMuestreo().obs, H_MOB.length).map(function (r) {
    var c = {}; try { c = JSON.parse(String(r[6] || '{}')) || {}; } catch (e) {}
    return { estudio: String(r[0] || ''), en: fmtSello(r[1]), fecha: fmtDia(r[2]), hora: String(r[3] || ''),
      sucursal: String(r[4] || ''), por: String(r[5] || ''), c: c, total: Number(r[7]) || 0, nota: String(r[8] || '') };
  }).filter(function (x) { return x.estudio; });
}
/* Horas al azar, las mismas cada vez para el mismo estudio y día (nadie las puede
 * «escoger»), repartidas en todo el horario de la sucursal. */
function horasMuestreo(e, fecha) {
  var h = MUE_HORARIO[e.sid] || ['09:00', '20:30'];
  var ini = +h[0].slice(0, 2) * 60 + +h[0].slice(3), fin = +h[1].slice(0, 2) * 60 + +h[1].slice(3) - 10;
  var semilla = 0, txt = e.id + fecha;
  for (var i = 0; i < txt.length; i++) semilla = (semilla * 31 + txt.charCodeAt(i)) % 2147483647;
  var azar = function () { semilla = (semilla * 48271) % 2147483647; return semilla / 2147483647; };
  // una hora al azar dentro de cada tramo igual: al azar, pero sin amontonarse
  var n = e.porDia, paso = (fin - ini) / n, out = [];
  for (var k = 0; k < n; k++) {
    var m = Math.floor(ini + paso * k + azar() * paso);
    out.push(('0' + Math.floor(m / 60)).slice(-2) + ':' + ('0' + (m % 60)).slice(-2));
  }
  return out;
}
function exigeMuestreo(cred, escribir) {
  var yo = quien(cred);
  if (yo.rol === 'operaciones' || yo.esAdmin) return yo;
  if (!escribir && yo.rol === 'gerente') return yo;
  throw new Error('El muestreo de trabajo lo maneja el director operativo.');
}
function guardarEstudio(cred, e) {
  var yo = quien(cred);
  if (yo.rol !== 'operaciones') throw new Error('Los estudios los arma el director operativo.');
  e = e || {};
  if (SUC_KPI.indexOf(e.sid) < 0) throw new Error('Escoja la sucursal.');
  if (!esFecha(e.desde) || !esFecha(e.hasta) || e.hasta < e.desde) throw new Error('Escoja las fechas del estudio.');
  if (diasEntre(e.desde, e.hasta) > 45) throw new Error('Un estudio dura máximo 45 días.');
  var n = Math.round(Number(e.porDia) || 0);
  if (n < 4 || n > 30) throw new Error('Entre 4 y 30 observaciones por día.');
  var nombre = String(e.nombre || '').trim().slice(0, 80) || 'Muestreo ' + nombreUnidad(e.sid);
  var h = hojasMuestreo().est, x = null;
  filasEstudios().forEach(function (y) { if (y.id === String(e.id || '')) x = y; });
  var fila = [x ? x.id : 'MU' + Utilities.formatDate(new Date(), ZONA, 'yyyyMMddHHmmss'), nombre, e.sid, e.desde, e.hasta, n,
    JSON.stringify(MUE_CATS), e.estado === 'Cerrado' ? 'Cerrado' : 'Activo', x ? x.por : yo.nombre, x ? x.en : new Date(),
    String(e.nota || '').slice(0, 200)];
  if (x) h.getRange(x.fila, 1, 1, H_MUE.length).setValues([fila]); else h.appendRow(fila);
  h.getRange(x ? x.fila : h.getLastRow(), 4, 1, 2).setNumberFormat('@').setValues([[e.desde, e.hasta]]);
  var d = pantallaMuestreo(cred); d.ok = true; d.mensaje = x ? 'Estudio guardado.' : 'Estudio creado. Las horas de observación salen solas.';
  return d;
}
function pendientesMuestreo(yo) {
  var hoy = hoyISO(), hora = horaAhora(), obs = filasObs(), out = [];
  filasEstudios().forEach(function (e) {
    if (e.estado !== 'Activo' || hoy < e.desde || hoy > e.hasta) return;
    if (yo.rol === 'gerente' && yo.sucursal !== e.sid) return;
    var hechas = {};
    obs.forEach(function (o) { if (o.estudio === e.id && o.fecha === hoy) hechas[o.hora] = true; });
    horasMuestreo(e, hoy).forEach(function (hh) {
      if (hechas[hh]) return;
      // toca desde su hora hasta 45 minutos después; más tarde ya no vale (se perdería el azar)
      var min = function (t) { return +t.slice(0, 2) * 60 + +t.slice(3); };
      var dif = min(hora) - min(hh);
      out.push({ estudio: e.id, nombre: e.nombre, sucursal: nombreUnidad(e.sid), hora: hh,
        estado: dif < 0 ? 'Próxima' : dif <= 45 ? 'Ahora' : 'Perdida' });
    });
  });
  return out;
}
function pantallaMuestreo(cred) {
  var yo = exigeMuestreo(cred, false);
  var ests = filasEstudios().filter(function (e) { return yo.rol !== 'gerente' || e.sid === yo.sucursal; });
  var obs = filasObs();
  return { puede: yo.rol === 'operaciones', categorias: MUE_CATS, hoy: hoyISO(), hora: horaAhora(),
    sucursales: SUC_KPI.map(function (s) { return { id: s, nombre: nombreUnidad(s) }; }),
    estudios: ests.map(function (e) {
      var mias = obs.filter(function (o) { return o.estudio === e.id; });
      return { id: e.id, nombre: e.nombre, sid: e.sid, sucursal: nombreUnidad(e.sid), desde: e.desde, hasta: e.hasta,
        porDia: e.porDia, estado: e.estado, por: e.por, observaciones: mias.length,
        personas: mias.reduce(function (a, o) { return a + o.total; }, 0),
        esperadas: Math.max(0, diasEntre(e.desde, e.hasta < hoyISO() ? e.hasta : hoyISO()) + 1) * e.porDia };
    }).reverse(),
    pendientes: pendientesMuestreo(yo) };
}
/** Una observación: cuántas personas había en cada actividad a esa hora. */
function registrarObservacion(cred, d) {
  var yo = quien(cred);
  if (yo.rol !== 'operaciones' && yo.rol !== 'gerente' && !yo.esAdmin) throw new Error('Su usuario no hace observaciones.');
  d = d || {};
  var e = null;
  filasEstudios().forEach(function (x) { if (x.id === String(d.estudio)) e = x; });
  if (!e || e.estado !== 'Activo') throw new Error('Ese estudio no está activo.');
  if (yo.rol === 'gerente' && yo.sucursal !== e.sid) throw new Error('Ese estudio es de otra sucursal.');
  var p = null;
  pendientesMuestreo(yo).forEach(function (x) { if (x.estudio === e.id && x.hora === String(d.hora)) p = x; });
  if (!p) throw new Error('Esa observación ya se anotó o no es de hoy.');
  if (p.estado === 'Próxima') throw new Error('Todavía no es hora de esa observación (' + p.hora + ').');
  if (p.estado === 'Perdida') throw new Error('Esa observación ya pasó (' + p.hora + '): ya no cuenta. Espere la siguiente.');
  var c = {}, total = 0;
  MUE_CATS.forEach(function (k) { var v = Math.round(Number((d.conteos || {})[k.k]) || 0); if (v < 0 || v > 40) throw new Error('Revise los números.');
    if (v) { c[k.k] = v; total += v; } });
  if (!total) throw new Error('Cuente al menos a una persona.');
  var hoy = hoyISO(), h = hojasMuestreo().obs;
  h.appendRow([e.id, new Date(), hoy, p.hora, nombreUnidad(e.sid), yo.nombre, JSON.stringify(c), total, String(d.nota || '').slice(0, 200)]);
  h.getRange(h.getLastRow(), 3, 1, 2).setNumberFormat('@').setValues([[hoy, p.hora]]);
  var r = pantallaMuestreo(cred); r.ok = true; r.mensaje = 'Observación de las ' + p.hora + ' anotada: ' + total + (total === 1 ? ' persona.' : ' personas.');
  return r;
}
/** Los resultados de un estudio: % del tiempo en cada actividad, con su margen de error. */
function resultadosMuestreo(cred, id) {
  exigeMuestreo(cred, false);
  var e = null;
  filasEstudios().forEach(function (x) { if (x.id === String(id)) e = x; });
  if (!e) throw new Error('No se encontró el estudio.');
  var obs = filasObs().filter(function (o) { return o.estudio === e.id; });
  var tot = {}, n = 0, porHora = {}, porDia = {}, porTurno = { AM: {}, PM: {} }, nT = { AM: 0, PM: 0 };
  var pm = INICIO_PM[e.sid] || '13:00';
  obs.forEach(function (o) {
    var hh = o.hora.slice(0, 2), wd = diaSemana(o.fecha), tu = o.hora >= pm ? 'PM' : 'AM';
    MUE_CATS.forEach(function (k) { var v = o.c[k.k] || 0; if (!v) return;
      tot[k.k] = (tot[k.k] || 0) + v;
      var ph = porHora[hh] || (porHora[hh] = { P: 0, A: 0, O: 0, n: 0 }); ph[k.t] += v;
      var pd = porDia[wd] || (porDia[wd] = { P: 0, A: 0, O: 0, n: 0, obs: 0 }); pd[k.t] += v;
      porTurno[tu][k.t] = (porTurno[tu][k.t] || 0) + v; });
    n += o.total;
    (porHora[hh] = porHora[hh] || { P: 0, A: 0, O: 0, n: 0 }).n += o.total;
    var pdd = porDia[wd] || (porDia[wd] = { P: 0, A: 0, O: 0, n: 0, obs: 0 }); pdd.n += o.total; pdd.obs++;
    nT[tu] += o.total;
  });
  var margen = function (p, nn) { return nn ? Math.round(1.96 * Math.sqrt(p * (1 - p) / nn) * 1000) / 10 : null; };
  var cats = MUE_CATS.map(function (k) { var p = n ? (tot[k.k] || 0) / n : 0;
    return { k: k.k, nombre: k.n, tipo: k.t, cuenta: tot[k.k] || 0, pct: Math.round(p * 1000) / 10, margen: margen(p, n) }; });
  var grupo = function (t) { var s = 0; MUE_CATS.forEach(function (k) { if (k.t === t) s += tot[k.k] || 0; }); return n ? s / n : 0; };
  var pO = grupo('O');
  // cuántas observaciones hacen falta para ±5 puntos en lo ocioso
  var necesarias = Math.ceil(1.96 * 1.96 * Math.max(pO, 0.1) * (1 - Math.max(pO, 0.1)) / 0.0025);
  var promPersonas = obs.length ? n / obs.length : 0;
  return { estudio: { id: e.id, nombre: e.nombre, sucursal: nombreUnidad(e.sid), desde: e.desde, hasta: e.hasta, porDia: e.porDia },
    observaciones: obs.length, personas: n, promedioPersonas: Math.round(promPersonas * 10) / 10,
    categorias: cats, productivo: Math.round(grupo('P') * 1000) / 10, apoyo: Math.round(grupo('A') * 1000) / 10,
    ocioso: Math.round(pO * 1000) / 10, margenOcioso: margen(pO, n), necesarias: necesarias,
    personasOciosas: Math.round(promPersonas * pO * 10) / 10,
    porHora: Object.keys(porHora).sort().map(function (h) { var x = porHora[h];
      return { hora: h, n: x.n, P: x.n ? Math.round(x.P / x.n * 100) : 0, A: x.n ? Math.round(x.A / x.n * 100) : 0, O: x.n ? Math.round(x.O / x.n * 100) : 0 }; }),
    porDia: [1, 2, 3, 4, 5, 6, 0].filter(function (w) { return porDia[w]; }).map(function (w) { var x = porDia[w];
      return { dia: DIAS_SEM[w], n: x.n, obs: x.obs, personas: Math.round(x.n / x.obs * 10) / 10,
        O: Math.round(x.O / x.n * 100), P: Math.round(x.P / x.n * 100) }; }),
    porTurno: ['AM', 'PM'].map(function (t) { var x = porTurno[t], nn = nT[t];
      return { turno: t, n: nn, P: nn ? Math.round((x.P || 0) / nn * 100) : 0, O: nn ? Math.round((x.O || 0) / nn * 100) : 0 }; }) };
}

/* ── avisos del muestreo y de las compras ── */
function avisosMuestreoCompras(yo, out) {
  if (yo.rol === 'gerente' || yo.rol === 'operaciones') {
    pendientesMuestreo(yo).forEach(function (p) {
      if (p.estado !== 'Ahora') return;
      out.push({ id: 'mue-' + p.estudio + '-' + hoyISO() + '-' + p.hora, tipo: 'Muestreo de trabajo', cuando: '',
        titulo: 'Toca observar ' + p.sucursal + ' (' + p.hora + ')', texto: 'Cuente qué está haciendo cada persona en este momento.',
        abrir: { pag: yo.rol === 'gerente' ? 'kpi' : 'mue' } });
    });
  }
  if (esContador(yo)) {
    ordenesAgrupadas().forEach(function (o) {
      if (o.estado === 'Anulada') return;
      if (o.pago === 'Cotizada') out.push({ id: 'oc-pagar-' + o.numero, tipo: 'Orden por pagar', cuando: o.cotizadaEn,
        titulo: o.numero + ' · ' + o.proveedor + ' · ' + dinero(o.totalCotizado), texto: 'Cotizada por ' + o.cotizadaPor + '.',
        abrir: { pag: 'conta', ct: 'pagos' } });
      if (o.pago === 'Por cotizar' && o.estado === 'Pedida') out.push({ id: 'oc-cot-' + o.numero, tipo: 'Orden sin cotización', cuando: o.creadaEn,
        titulo: o.numero + ' · ' + o.proveedor, texto: 'Samuel la pidió. Anote la cotización cuando llegue.', abrir: { pag: 'conta', ct: 'pagos' } });
    });
  }
  if (esContador(yo)) {
    var hoy = hoyISO(), semana = semanaAPagar(hoy), pagos = pagosRepartidores();
    filasProveedores().forEach(function (p) {
      if (p.tipo !== 'Repartidor' || !p.activo) return;
      if (pagos.some(function (x) { return x.repartidor === p.nombre && x.semana === semana; })) return;
      out.push({ id: 'rep-' + semana + '-' + p.nombre, tipo: 'Pago a repartidor', cuando: '',
        titulo: (diaSemana(hoy) === 1 ? 'Hoy toca pagar a ' : 'Falta pagar a ') + p.nombre,
        texto: 'Envíos de la semana ' + semanaTxt(semana) + '.', abrir: { pag: 'conta', ct: 'rep' } });
    });
  }
  if (yo.rol === 'bodega') {
    ordenesAgrupadas().forEach(function (o) {
      if (o.pago === 'Pagada' && o.estado === 'Pedida') out.push({ id: 'oc-pagada-' + o.numero, tipo: 'Orden pagada', cuando: o.pagadaEn,
        titulo: o.numero + ' de ' + o.proveedor + ' ya está pagada', texto: 'Recíbala cuando llegue el producto.', abrir: { tab: 'compras' } });
    });
  }
}

/* ════════════ REPARTIDORES (DELIVERY) ════════════
 * Edgar Delivery, Marvin Delivery… se les paga cada lunes lo que se cobró de envío en
 * los tickets de la semana anterior (lunes a domingo). Es variable y no es venta de
 * CAIX: se registra en Gastos de la sucursal con la categoría «Envíos entregados a
 * repartidores», y el estado de resultados lo resta de las ventas. Lo paga el
 * contador; lo ven el director financiero y el administrador. La bodega no lo ve. */
function marcaRepartidores() {
  var h = hojaCat('Proveedores', H_PROVEEDORES);
  filasProveedores().forEach(function (p) {
    if (!p.tipo && /delivery|repartid|motoris/i.test(p.nombre)) h.getRange(p.fila, 7).setValue('Repartidor');
  });
}
function exigeVerRepartidores(cred) {
  var yo = quien(cred);
  if (!(yo.esAdmin || esContador(yo) || yo.rol === 'finanzas' || yo.rol === 'dueno')) throw new Error('Esto lo ve contabilidad.');
  return yo;
}
/** El lunes de la semana que se paga: la anterior a `fecha`. */
function semanaAPagar(fecha) { return masDias(lunesDe(fecha), -7); }
function semanaTxt(lunes) { return dia(lunes).slice(0, 5) + ' al ' + dia(masDias(lunes, 6)).slice(0, 5); }
function pagosRepartidores() {
  var out = [];
  leeTodo(libro().getSheetByName('Gastos'), 16).forEach(function (r) {
    if (!r[0] || String(r[5]) !== CAT_ENVIOS || String(r[10]).toLowerCase() === 'sí') return;
    var fac = String(r[15] || ''), m = /^ENV-(\d{4}-\d{2}-\d{2})$/.exec(fac);
    out.push({ id: String(r[0]), en: fmtSello(r[1]), por: String(r[2] || ''), sucursal: String(r[4] || ''),
      fecha: fmtDia(r[6]), monto: Number(r[8]) || 0, nota: String(r[9] || ''), repartidor: String(r[14] || ''),
      semana: m ? m[1] : '' });
  });
  return out;
}
function listaRepartidores(cred) {
  var yo = exigeVerRepartidores(cred);
  var hoy = hoyISO(), semana = semanaAPagar(hoy), pagos = pagosRepartidores();
  var reps = filasProveedores().filter(function (p) { return p.tipo === 'Repartidor'; });
  var semanas = [];
  for (var k = 0; k < 8; k++) semanas.push(masDias(semana, -7 * k));
  var lista = reps.map(function (p) {
    var mios = pagos.filter(function (x) { return x.repartidor.toLowerCase() === p.nombre.toLowerCase(); });
    var estaSemana = mios.filter(function (x) { return x.semana === semana; });
    return { nombre: p.nombre, tel: p.tel, activo: p.activo, sucursal: p.sucursal, sucursalNombre: p.sucursal ? nombreUnidad(p.sucursal) : '',
      pagado: estaSemana.length ? estaSemana.reduce(function (a, x) { return a + x.monto; }, 0) : null,
      pagos: estaSemana,
      porSemana: semanas.map(function (w) { return r2(mios.filter(function (x) { return x.semana === w; })
        .reduce(function (a, x) { return a + x.monto; }, 0)); }),
      mes: r2(mios.filter(function (x) { return x.fecha.slice(0, 7) === hoy.slice(0, 7); }).reduce(function (a, x) { return a + x.monto; }, 0)) };
  });
  return { hoy: hoy, esLunes: diaSemana(hoy) === 1, semana: semana, semanaTxt: semanaTxt(semana),
    semanas: semanas.map(function (w) { return { lunes: w, txt: semanaTxt(w) }; }), repartidores: lista,
    sucursales: sucursalesDestino().map(function (u) { return { id: u.id, nombre: u.nombre }; }),
    formas: FORMAS_PAGO, puedePagar: yo.esAdmin || esContador(yo),
    historial: pagos.slice(-40).reverse() };
}
/** Pagar a un repartidor lo cobrado en envíos en una semana y sucursal. */
function pagarRepartidor(cred, d) {
  var yo = quien(cred);
  if (!(yo.esAdmin || esContador(yo))) throw new Error('A los repartidores les paga el contador.');
  d = d || {};
  var p = null;
  filasProveedores().forEach(function (x) { if (x.tipo === 'Repartidor' && x.nombre === String(d.repartidor || '')) p = x; });
  if (!p) throw new Error('Escoja el repartidor.');
  var u = unidadPorId(String(d.sucursal || ''));
  if (!u || !u.vende) throw new Error('Escoja la sucursal donde hizo los envíos.');
  var semana = esFecha(d.semana) ? lunesDe(d.semana) : semanaAPagar(hoyISO());
  if (semana > lunesDe(hoyISO())) throw new Error('Esa semana todavía no llega.');
  var monto = r2(String(d.monto == null ? '' : d.monto).replace(/[Q,\s]/g, ''));
  if (!(monto > 0)) throw new Error('Escriba lo que se cobró de envíos en los tickets de esa semana.');
  if (monto > 50000) throw new Error('Ese monto parece equivocado. Revíselo.');
  var envios = Math.round(Number(d.envios) || 0);
  var fecha = esFecha(d.fecha) ? d.fecha : hoyISO();
  if (fecha > hoyISO()) throw new Error('La fecha de pago no puede ser de un día que no ha llegado.');
  var forma = FORMAS_PAGO.indexOf(d.forma) >= 0 ? d.forma : 'Efectivo';
  var ref = String(d.referencia || '').trim().slice(0, 40);
  if (forma !== 'Efectivo' && ref.length < 3) throw new Error('Escriba el número de referencia.');
  var ya = pagosRepartidores().filter(function (x) { return x.repartidor === p.nombre && x.semana === semana && x.sucursal === u.nombre; });
  if (ya.length) throw new Error('A ' + p.nombre + ' ya se le pagó la semana del ' + semanaTxt(semana) + ' en ' + u.nombre +
    ' (' + dinero(ya[0].monto) + ', ' + ya[0].por + '). Si está mal, anúlelo en Reportes y páguelo otra vez.');
  var v = { unidad: u.id, fechaPago: fecha, proveedor: p.nombre, factura: 'ENV-' + semana, mes: fecha.slice(0, 7),
    lineas: [], categoria: CAT_ENVIOS, monto: monto,
    nota: ('Envíos de la semana ' + semanaTxt(semana) + (envios ? ' · ' + envios + ' envíos' : '') + ' · ' + forma +
      (ref ? ' ' + ref : '') + (d.nota ? ' · ' + String(d.nota).slice(0, 60) : '')).slice(0, 140) };
  var r = escribeGasto(v, yo.nombre, new Date(), null);
  var out = listaRepartidores(cred); out.ok = true; out.nuevo = r.nuevo;
  out.mensaje = p.nombre + ': ' + dinero(monto) + ' por los envíos del ' + semanaTxt(semana) + ' en ' + u.nombre + '.';
  return out;
}
/** Agregar un repartidor desde contabilidad. */
function guardarRepartidor(cred, d) {
  var yo = exigeVerRepartidores(cred);
  if (yo.rol === 'dueno') throw new Error('Su usuario solo consulta reportes.');
  d = d || {};
  var nombre = String(d.nombre || '').trim().replace(/\s+/g, ' ').slice(0, 60);
  if (nombre.length < 3) throw new Error('Escriba el nombre del repartidor.');
  var suc = unidadPorId(String(d.sucursal || '')) ? String(d.sucursal) : '';
  var h = hojaCat('Proveedores', H_PROVEEDORES), ya = null;
  filasProveedores().forEach(function (x) { if (x.nombre.toLowerCase() === nombre.toLowerCase()) ya = x; });
  if (ya) {
    h.getRange(ya.fila, 4).setValue('Sí');
    h.getRange(ya.fila, 7, 1, 2).setValues([['Repartidor', suc]]);
    if (d.tel) h.getRange(ya.fila, 3).setValue(String(d.tel).slice(0, 20));
  } else h.appendRow([nombre, '', String(d.tel || '').slice(0, 20), 'Sí', new Date(), yo.nombre, 'Repartidor', suc]);
  var out = listaRepartidores(cred); out.ok = true; out.mensaje = nombre + ' queda como repartidor.';
  return out;
}

/* ════════════ PRESUPUESTO MENSUAL POR SUCURSAL ════════════
 * La venta presupuestada de la red es Q400,000 al mes (ya trae la utilidad bruta y
 * la neta adentro). Se reparte entre las tres sucursales y, con eso, el director
 * financiero o el contador asignan cuánto puede gastar cada unidad en cada rubro
 * (planilla, luz, renta, internet, gas, Paladar POS…). Lo que sobra de la venta
 * después de los rubros es la utilidad que queda. Cada mes se compara con lo real. */
var H_PRES = ['Mes', 'Unidad', 'Rubro', 'Monto', 'Actualizado por', 'Actualizado en'];
var PRES_VENTA = '__venta';
var PRES_META = '__meta_neta_pct';
/* Rubros de arranque: los que pidió el propietario primero, luego el resto. */
var RUBROS_BASE = ['Nómina y bonificaciones', 'IGSS y prestaciones', 'Energía eléctrica', 'Renta',
  'Internet y teléfono', 'Gas propano', 'Sistema punto de venta (Paladar POS)', 'Agua',
  'Insumos de cocina', 'Bebidas y gaseosas', 'Empaque y desechables', 'Limpieza y químicos',
  'Mantenimiento y reparaciones', 'Equipo y mobiliario', 'Combustible y transporte', 'Mercadeo y publicidad',
  'Comisiones de plataformas', 'Papelería y administración', 'Impuestos y tasas', 'Otros gastos'];
var REPARTO_VENTA = { centro: 0.491, almendras: 0.330, parque: 0.179 };   // histórico jul–ago 2026

function hojaPres(ss) {
  ss = ss || libro();
  var h = ss.getSheetByName('Presupuesto mensual');
  if (!h) { h = hojaLimpia(ss, 'Presupuesto mensual', H_PRES); h.getRange('A:A').setNumberFormat('@');
    h.getRange('D:D').setNumberFormat('"Q"#,##0.00'); h.getRange('F:F').setNumberFormat('dd/mm/yyyy hh:mm'); h.setColumnWidth(3, 260); }
  return h;
}
function filasPres(mes) {
  return leeTodo(hojaPres(), H_PRES.length).map(function (r, i) {
    return { fila: i + 2, mes: String(r[0] || '').slice(0, 7), unidad: String(r[1] || ''), rubro: String(r[2] || ''),
      monto: Number(r[3]) || 0, por: String(r[4] || ''), en: fmtSello(r[5]) };
  }).filter(function (x) { return x.mes && (!mes || x.mes === mes); });
}
/** Rubros: los de arranque más los que agreguen (viven en «Clasificación»). */
var ESPECIALES = null;
function catsEspeciales() { return [CAT_CAJA, CAT_TRASLADO, CAT_COMISION, CAT_ENVIOS]; }
function rubros() {
  var out = RUBROS_BASE.slice(), esp = catsEspeciales();
  CATEGORIAS.forEach(function (c) { if (out.indexOf(c) < 0) out.push(c); });
  try { leeTodo(libro().getSheetByName('Clasificación'), 1).forEach(function (r) {
    var c = String(r[0] || '').trim(); if (c && esp.indexOf(c) < 0 && out.indexOf(c) < 0) out.push(c); }); } catch (e) {}
  return out;
}
/** Los rubros agregados también se pueden escoger al registrar un gasto. */
var _CATX = false;
function extiendeCategorias() {
  if (_CATX) return; _CATX = true;
  // los rubros agregados se guardan 5 minutos en memoria: así cada llamada no abre la hoja solo para leerlos
  var extra = null, c = null;
  try { c = CacheService.getScriptCache(); var t = c.get('rubros_x1'); if (t) extra = JSON.parse(t); } catch (e) { c = null; }
  if (!extra) {
    extra = [];
    try { rubros().forEach(function (x) { if (CATEGORIAS.indexOf(x) < 0) extra.push(x); }); } catch (e) {}
    try { if (c) c.put('rubros_x1', JSON.stringify(extra), 300); } catch (e) {}
  }
  extra.forEach(function (x) { if (CATEGORIAS.indexOf(x) < 0) CATEGORIAS.push(x); });
}
function exigePresupuesto(cred, escribir) {
  var yo = quien(cred);
  if (yo.rol === 'finanzas' || esContador(yo)) return yo;
  if (!escribir && (yo.esAdmin || yo.rol === 'dueno')) return yo;
  throw new Error(escribir ? 'El presupuesto lo asignan el director financiero o el contador.' : 'Esto lo ve contabilidad.');
}
/** Lo gastado de verdad en un mes, por unidad y rubro (como el estado de resultados). */
function gastoRealMes(mes) {
  var movs = movimientos(4000).filter(function (m) { return m.mes === mes && !m.anulado; });
  var det = detalleDeGastos(movs), conDet = {}, g = {};
  var suma = function (u, c, m) { if (!g[u]) g[u] = {}; g[u][c] = (g[u][c] || 0) + m; };
  det.forEach(function (x) { conDet[x.id] = true; suma(x.unidad, x.categoria, x.subtotal); });
  var venta = {};
  movs.forEach(function (m) {
    if (m.clase === 'ingreso') { venta[m.unidad] = (venta[m.unidad] || 0) + m.monto;
      if (m.caja > 0) suma(m.unidad, CAT_CAJA, m.caja);
      if (m.comision > 0) suma(m.unidad, CAT_COMISION, m.comision); return; }
    if (m.titulo === CAT_ENVIOS) { venta[m.unidad] = (venta[m.unidad] || 0) - m.monto; return; }
    if (conDet[m.id]) return;
    suma(m.unidad, m.titulo, m.monto);
  });
  return { gasto: g, venta: venta };
}
function presupuestoMes(cred, mes, interno) {
  var yo = interno ? { rol: 'admin', esAdmin: true } : exigePresupuesto(cred, false);
  var hoy = hoyISO();
  mes = /^\d{4}-\d{2}$/.test(String(mes || '')) ? mes : hoy.slice(0, 7);
  var filas = filasPres(mes), unidades = UNIDADES.map(function (u) { return { id: u.id, nombre: u.nombre, vende: u.vende }; });
  var asig = {}, venta = {}, meta = null, rs = rubros();
  filas.forEach(function (x) {
    if (x.rubro === PRES_VENTA) venta[x.unidad] = x.monto;
    else if (x.rubro === PRES_META) meta = x.monto;
    else { (asig[x.unidad] = asig[x.unidad] || {})[x.rubro] = x.monto; if (rs.indexOf(x.rubro) < 0) rs.push(x.rubro); }
  });
  var hay = filas.length > 0;
  var totalVenta = 400000;
  try { totalVenta = Number(ajustesInv().ventaMensual) || 400000; } catch (e) {}
  if (!Object.keys(venta).length) sucursalesDestino().forEach(function (u) { venta[u.nombre] = Math.round(totalVenta * (REPARTO_VENTA[u.id] || 0)); });
  var real = gastoRealMes(mes), cl = clasifica();
  var esCosto = function (c) { return cl.de[c] && cl.de[c].grupo === G_COSTO; };
  var porUnidad = unidades.map(function (u) {
    var a = asig[u.nombre] || {}, costo = 0, oper = 0, gr = real.gasto[u.nombre] || {};
    Object.keys(a).forEach(function (c) { if (esCosto(c)) costo += a[c]; else oper += a[c]; });
    var gc = 0, go = 0;
    Object.keys(gr).forEach(function (c) { if (c === CAT_TRASLADO) return; if (esCosto(c)) gc += gr[c]; else go += gr[c]; });
    // lo que la bodega despacha ya está en el costo de la sucursal que lo recibe
    var v = venta[u.nombre] || 0;
    return { id: u.id, nombre: u.nombre, vende: u.vende, venta: v, costo: r2(costo), operacion: r2(oper),
      utilidadBruta: r2(v - costo), utilidadNeta: r2(v - costo - oper),
      real: { venta: r2(real.venta[u.nombre] || 0), costo: r2(gc), operacion: r2(go) },
      rubros: rs.map(function (c) { var p = a[c] || 0, g = gr[c] || 0;
        var ib = RUBROS_BASE.indexOf(c);
        return { rubro: c, costo: esCosto(c), presupuesto: r2(p), gastado: r2(g), pct: p ? Math.round(g / p * 100) : (g ? null : 0),
          principal: ib >= 0 && ib < 7 }; }) };
  });
  var tot = { venta: 0, costo: 0, operacion: 0, realVenta: 0, realGasto: 0 };
  porUnidad.forEach(function (u) { tot.venta += u.venta; tot.costo += u.costo; tot.operacion += u.operacion;
    tot.realVenta += u.real.venta; tot.realGasto += u.real.costo + u.real.operacion; });
  tot.utilidadBruta = r2(tot.venta - tot.costo); tot.utilidadNeta = r2(tot.venta - tot.costo - tot.operacion);
  tot.metaNetaPct = meta == null ? 15 : meta;
  tot.metaNeta = r2(tot.venta * tot.metaNetaPct / 100);
  tot.disponible = r2(tot.venta - tot.metaNeta - tot.costo - tot.operacion);
  tot.realUtilidad = r2(tot.realVenta - tot.realGasto);
  ['venta', 'costo', 'operacion', 'realVenta', 'realGasto'].forEach(function (k) { tot[k] = r2(tot[k]); });
  var est = null; try { est = estadoPres(mes); } catch (e) {}
  return { mes: mes, hay: hay, puede: yo.rol === 'finanzas' || esContador(yo), rubros: rs,
    estado: est, puedeAprobar: yo.rol === 'finanzas' && !!est && est.estado === PRES_EST.POR, esFinanzas: yo.rol === 'finanzas',
    especiales: catsEspeciales(), unidades: porUnidad, total: tot,
    diasMes: Number(finDeMes(mes).slice(8)), diaHoy: mes === hoy.slice(0, 7) ? Number(hoy.slice(8)) : (mes < hoy.slice(0, 7) ? Number(finDeMes(mes).slice(8)) : 0) };
}
/** Guarda el presupuesto de un mes: {venta:{Centro:…}, rubros:{Centro:{Renta:…}}, metaNetaPct}. */
function guardarPresupuesto(cred, mes, d) {
  var yo = exigePresupuesto(cred, true);
  if (!/^\d{4}-\d{2}$/.test(String(mes || ''))) throw new Error('Escoja el mes.');
  d = d || {};
  var nums = function (v) { var n = Number(String(v == null ? '' : v).replace(/[Q,\s]/g, '')); return isFinite(n) && n > 0 ? r2(n) : 0; };
  var filas = [], ahora = new Date(), nombres = UNIDADES.map(function (u) { return u.nombre; });
  Object.keys(d.venta || {}).forEach(function (u) { if (nombres.indexOf(u) >= 0) filas.push([mes, u, PRES_VENTA, nums(d.venta[u]), yo.nombre, ahora]); });
  if (d.metaNetaPct != null && String(d.metaNetaPct).trim() !== '') {
    var mp = Number(String(d.metaNetaPct).replace(',', '.'));
    if (!isFinite(mp) || mp < 0 || mp > 80) throw new Error('La utilidad neta esperada va de 0 a 80 %.');
    filas.push([mes, 'Red', PRES_META, mp, yo.nombre, ahora]);
  }
  var rs = rubros();
  Object.keys(d.rubros || {}).forEach(function (u) {
    if (nombres.indexOf(u) < 0) return;
    Object.keys(d.rubros[u]).forEach(function (c) {
      if (rs.indexOf(c) < 0) throw new Error('El rubro «' + c + '» no existe. Agréguelo primero.');
      var n = nums(d.rubros[u][c]); if (n) filas.push([mes, u, c, n, yo.nombre, ahora]); });
  });
  var lock = LockService.getScriptLock();
  lock.waitLock(20000); _LEE = {};
  try {
    var h = hojaPres(), viejas = filasPres(mes);
    for (var i = viejas.length - 1; i >= 0; i--) h.deleteRow(viejas[i].fila);
    if (filas.length) {
      var f0 = h.getLastRow() + 1;
      h.getRange(f0, 1, filas.length, 1).setNumberFormat('@');
      h.getRange(f0, 1, filas.length, H_PRES.length).setValues(filas);
    }
  } finally { lock.releaseLock(); }
  var fin = yo.rol === 'finanzas', env = '';
  ponEstadoPres(mes, [mes, fin ? PRES_EST.APR : PRES_EST.POR, yo.nombre, new Date(), fin ? yo.nombre : '', fin ? new Date() : '', '']);
  if (fin && !d.sinCorreo) env = correoPresupuesto(mes, yo.nombre);
  var r = presupuestoMes(cred, mes); r.ok = true;
  r.mensaje = fin ? 'Presupuesto de ' + periodoTxt(mes) + ' guardado y aprobado.' + (env ? ' Resumen enviado a ' + env + '.' : '')
    : 'Presupuesto de ' + periodoTxt(mes) + ' enviado al director financiero para que lo apruebe.';
  return r;
}
function copiarPresupuesto(cred, mes) {
  exigePresupuesto(cred, true);
  var ant = mesAnt(mes), filas = filasPres(ant);
  if (!filas.length) throw new Error('El mes anterior no tiene presupuesto.');
  var d = { venta: {}, rubros: {}, metaNetaPct: null };
  filas.forEach(function (x) {
    if (x.rubro === PRES_VENTA) d.venta[x.unidad] = x.monto;
    else if (x.rubro === PRES_META) d.metaNetaPct = x.monto;
    else (d.rubros[x.unidad] = d.rubros[x.unidad] || {})[x.rubro] = x.monto;
  });
  d.sinCorreo = true;
  var r = guardarPresupuesto(cred, mes, d); r.mensaje = 'Copiado de ' + periodoTxt(ant) + '. Revíselo y ' + (r.esFinanzas ? 'guárdelo para aprobarlo.' : 'mándelo a aprobar.');
  return r;
}
/** Un rubro nuevo: queda en «Clasificación» (y se puede escoger al registrar un gasto). */
function agregarRubro(cred, nombre, grupo) {
  exigePresupuesto(cred, true);
  nombre = String(nombre || '').trim().replace(/\s+/g, ' ').slice(0, 50);
  if (nombre.length < 3) throw new Error('Escriba el nombre del rubro.');
  if (rubros().some(function (c) { return sinAcentoS(c) === sinAcentoS(nombre); })) throw new Error('Ese rubro ya existe.');
  var h = libro().getSheetByName('Clasificación') || clasificacionHoja(libro());
  h.appendRow([nombre, grupo === G_COSTO ? G_COSTO : G_OPER, grupo === G_COSTO ? 'Materia prima y empaque' : 'Otros rubros']);
  try { CacheService.getScriptCache().remove('rubros_x1'); } catch (e) {}
  _CATX = false; extiendeCategorias();
  return { ok: true, mensaje: 'Rubro «' + nombre + '» agregado. También sale al registrar un gasto.', rubros: rubros() };
}
function avisosPresupuestoDatos(out, p) {
  p.unidades.forEach(function (u) { u.rubros.forEach(function (r) {
    if (!r.presupuesto || r.pct == null || r.pct < 90) return;
    out.push({ id: 'pres-' + p.mes + '-' + u.id + '-' + sinAcentoS(r.rubro).replace(/ /g, '') + (r.pct >= 100 ? '-100' : '-90'),
      tipo: r.pct >= 100 ? 'Presupuesto pasado' : 'Presupuesto casi agotado', cuando: '',
      titulo: u.nombre + ' · ' + r.rubro + ': ' + r.pct + ' %', texto: dinero(r.gastado) + ' de ' + dinero(r.presupuesto) + ' en ' + periodoTxt(p.mes),
      abrir: { pag: 'dash' } }); }); });
}

/* ════════════ SOLICITUDES DE DINERO (director operativo → contabilidad) ════════════
 * Alvaro pide dinero para equipo, reparaciones o visitas técnicas → el director
 * financiero lo aprueba (o lo rechaza) → el contador entrega el dinero o paga al
 * proveedor, y se registra solo en Gastos de la sucursal → Alvaro marca que ya se
 * hizo el trabajo. */
var H_FONDOS = ['No.', 'Pedida en', 'Pedida por', 'Tipo', 'Unidad', 'Qué se necesita', 'Proveedor', 'Monto',
  'Urgencia', 'Estado', 'Aprobada por', 'Aprobada en', 'Comentario', 'Pagada por', 'Pagada en', 'Monto pagado',
  'Forma de pago', 'Referencia', 'Factura', 'ID gasto', 'Terminada en', 'Nota final', 'Cotización (archivo)', 'Nombre del archivo',
  'Verificada por', 'Verificada en', 'Comentario de verificación', 'Devuelta por', 'Devuelta en', 'Motivo de devolución', 'Monto antes de devolver',
  'Requiere dinero', 'Autorizó (administrador)', 'Autorizó en', 'Comentario del administrador', 'Etapa del rechazo'];
var FONDO_TIPOS = { 'Equipo': 'Equipo y mobiliario', 'Reparación': 'Mantenimiento y reparaciones',
  'Visita técnica': 'Mantenimiento y reparaciones', 'Otro': 'Otros gastos' };
/* Lo que la CMO puede pedir pagar (con su factura): cada tipo cae en su rubro del estado de resultados. */
var FONDO_TIPOS_CMO = { 'Sueldo': 'Nómina y bonificaciones', 'Renta': 'Renta', 'Internet': 'Internet y teléfono', 'Luz': 'Energía eléctrica',
  'Suscripción': 'Mercadeo y publicidad', 'Publicidad': 'Mercadeo y publicidad', 'Papelería e impresiones': 'Papelería y administración',
  'Servicio o proveedor': 'Otros gastos', 'Otro': 'Otros gastos' };
function rubroDeFondo(tipo) { return FONDO_TIPOS[tipo] || FONDO_TIPOS_CMO[tipo] || FONDO_TIPOS_BODEGA[tipo] || 'Otros gastos'; }
var FONDO_EST = { VER: 'Por verificar', PED: 'Pedida', APR: 'Aprobada', RECH: 'Rechazada', PAG: 'Pagada', FIN: 'Terminada' };
function hojaFondos(ss) {
  ss = ss || libro();
  var h = ss.getSheetByName('Requisiciones');
  if (!h && ss.getSheetByName('Solicitudes de dinero')) { h = ss.getSheetByName('Solicitudes de dinero'); h.setName('Requisiciones'); }
  if (!h) { h = hojaLimpia(ss, 'Requisiciones', H_FONDOS); h.getRange('B:B').setNumberFormat('dd/mm/yyyy hh:mm');
    h.getRange('H:H').setNumberFormat('"Q"#,##0.00'); h.getRange('P:P').setNumberFormat('"Q"#,##0.00'); h.setColumnWidth(6, 320); }
  if (h.getLastColumn() < 27) encabezaAlFinal(h, 25, ['Verificada por', 'Verificada en', 'Comentario de verificación']);
  if (h.getLastColumn() < 31) { encabezaAlFinal(h, 28, ['Devuelta por', 'Devuelta en', 'Motivo de devolución', 'Monto antes de devolver']); h.getRange('AC:AC').setNumberFormat('dd/mm/yyyy hh:mm'); }
  if (h.getLastColumn() < 36) { encabezaAlFinal(h, 32, ['Requiere dinero', 'Autorizó (administrador)', 'Autorizó en', 'Comentario del administrador', 'Etapa del rechazo']); h.getRange('AH:AH').setNumberFormat('dd/mm/yyyy hh:mm'); }
  return h;
}
function filasFondos() {
  return leeTodo(hojaFondos(), H_FONDOS.length).map(function (r, i) {
    return { fila: i + 2, numero: String(r[0] || ''), en: fmtSello(r[1]), por: String(r[2] || ''), tipo: String(r[3] || ''),
      unidad: String(r[4] || ''), que: String(r[5] || ''), proveedor: String(r[6] || ''), monto: Number(r[7]) || 0,
      urgencia: String(r[8] || ''), estado: String(r[9] || ''), aprPor: String(r[10] || ''), aprEn: fmtSello(r[11]),
      comentario: String(r[12] || ''), pagPor: String(r[13] || ''), pagEn: fmtSello(r[14]), pagado: Number(r[15]) || 0,
      forma: String(r[16] || ''), referencia: String(r[17] || ''), factura: String(r[18] || ''), gasto: String(r[19] || ''),
      finEn: fmtSello(r[20]), notaFin: String(r[21] || ''), archivo: String(r[22] || ''), archivoNombre: String(r[23] || ''),
      verPor: String(r[24] || ''), verEn: fmtSello(r[25]), verCom: String(r[26] || ''),
      devPor: String(r[27] || ''), devEn: fmtSello(r[28]), devMotivo: String(r[29] || ''), montoAntes: Number(r[30]) || 0,
      reqDinero: String(r[31] || '').trim().toLowerCase() !== 'no', autPor: String(r[32] || ''), autEn: fmtSello(r[33]), autCom: String(r[34] || ''), etapaRech: String(r[35] || '') };
  }).filter(function (x) { return x.numero; });
}
/* El equipo de bodega y producción (Samuel y Paulino) pide sus gastos: los aprueba el director operativo (sin pasar por el director financiero)
 * y los paga contabilidad. El pago queda como gasto de Bodega Central, en el rubro que corresponde. */
function esEquipoBodega(yo) { return !!yo && (yo.rol === 'bodega' || yo.rol === 'produccion'); }
var FONDO_TIPOS_BODEGA = { 'Combustible y transporte': 'Combustible y transporte', 'Insumos de cocina': 'Insumos de cocina', 'Empaque y desechables': 'Empaque y desechables',
  'Limpieza y químicos': 'Limpieza y químicos', 'Reparación': 'Mantenimiento y reparaciones', 'Equipo': 'Equipo y mobiliario', 'Otro': 'Otros gastos' };
function veFondos(yo) { return yo.cmo || yo.esAdmin || yo.rol === 'operaciones' || yo.rol === 'finanzas' || esContador(yo) || yo.rol === 'dueno' || yo.rol === 'gerente' || esEquipoBodega(yo); }
function rolDe(nombre) { var r = ''; usuariosCache().forEach(function (u) { if (u.nombre === nombre) r = u.rol; }); return r; }
function buscaFondo(numero) {
  var x = null; filasFondos().forEach(function (y) { if (y.numero === String(numero)) x = y; });
  if (!x) throw new Error('No se encontró la solicitud ' + numero + '.');
  return x;
}
/** El director operativo verifica la solicitud de un gerente: si la confirma pasa a contabilidad, si no se devuelve. */
function pagarFondos(cred, numero, d) {
  var yo = quien(cred);
  if (!(esContador(yo) || yo.esAdmin)) throw new Error('El dinero lo entrega el contador.');
  d = d || {};
  var x = buscaFondo(numero);
  if (x.estado !== FONDO_EST.APR) throw new Error(x.estado === FONDO_EST.PED ? 'Primero la aprueba el director financiero.' : (x.estado === FONDO_EST.VER || x.estado === FONDO_ADM) ? 'Todavía falta que la autoricen.' : 'Esa solicitud ya está ' + x.estado.toLowerCase() + '.');
  var monto = r2(String(d.monto == null || d.monto === '' ? x.monto : d.monto).replace(/[Q,\s]/g, ''));
  if (!(monto > 0)) throw new Error('Escriba cuánto se pagó.');
  var forma = FORMAS_PAGO.indexOf(d.forma) >= 0 ? d.forma : 'Transferencia';
  var ref = String(d.referencia || '').trim().slice(0, 40);
  if (forma !== 'Efectivo' && ref.length < 3) throw new Error('Escriba el número de referencia.');
  var u = null; UNIDADES.forEach(function (y) { if (y.nombre === x.unidad) u = y; });
  var fecha = esFecha(d.fecha) ? d.fecha : hoyISO();
  var v = { unidad: u.id, fechaPago: fecha, proveedor: String(d.proveedor || x.proveedor || '').slice(0, 60),
    factura: String(d.factura || '').trim().slice(0, 30) || numero, mes: fecha.slice(0, 7), lineas: [],
    categoria: rubroDeFondo(x.tipo), monto: monto,
    nota: (numero + ' · ' + x.tipo + ': ' + x.que).slice(0, 140) };
  var h = hojaFondos();
  h.getRange(x.fila, 10).setValue(FONDO_EST.PAG);
  var r;
  try { r = escribeGasto(v, yo.nombre, new Date(), null); }
  catch (e) { h.getRange(x.fila, 10).setValue(FONDO_EST.APR); throw e; }
  if (v.proveedor) { try { aseguraProveedor(v.proveedor, yo.nombre, v.unidad); } catch (e) {} }
  h.getRange(x.fila, 14, 1, 7).setValues([[yo.nombre, new Date(), monto, forma, ref, v.factura, r.nuevo.id]]);
  var out = listaFondos(cred); out.ok = true; out.nuevo = r.nuevo;
  out.mensaje = numero + ' pagada: ' + dinero(monto) + '. Quedó en Gastos de ' + x.unidad + ' (' + v.categoria + ').';
  return out;
}
/** La cotización (PDF o foto) se puede subir al pedir o después, mientras no esté terminada. */
function adjuntarCotizacion(cred, numero, archivo) {
  permiteCmo();
  var yo = quien(cred);
  var x = buscaFondo(numero);
  var suya = x.por === yo.nombre && (x.estado === FONDO_EST.VER || x.estado === FONDO_ADM || (yo.cmo && x.estado === FONDO_EST.PED));
  if (!(yo.rol === 'operaciones' || yo.esAdmin || esContador(yo) || suya)) throw new Error('La cotización la sube quien pidió, el director operativo o el contador.');
  if (x.estado === FONDO_EST.FIN || x.estado === FONDO_EST.RECH) throw new Error('Esa solicitud ya está ' + x.estado.toLowerCase() + '.');
  var a = guardaArchivo(archivo, numero + ' cotización');
  hojaFondos().getRange(x.fila, 23, 1, 2).setValues([[a.url, a.nombre]]);
  var r = listaFondos(cred); r.ok = true; r.mensaje = 'Cotización guardada en ' + numero + '.';
  return r;
}
function terminarFondos(cred, numero, nota) {
  var yo = quien(cred);
  if (yo.rol !== 'operaciones' && !yo.esAdmin) throw new Error('Eso lo marca el director operativo.');
  var x = buscaFondo(numero);
  if (x.estado !== FONDO_EST.PAG) throw new Error('Esa solicitud todavía no está pagada.');
  hojaFondos().getRange(x.fila, 10).setValue(FONDO_EST.FIN);
  hojaFondos().getRange(x.fila, 21, 1, 2).setValues([[new Date(), String(nota || '').slice(0, 200)]]);
  var r = listaFondos(cred); r.ok = true; r.mensaje = numero + ' terminada.';
  return r;
}

/* ════════════ LA CMO: PRESUPUESTO DE SU ÁREA, SOLICITUDES Y AVANCE ════════════
 * Amalia (CMO) trabaja en las oficinas. Cada mes arma el presupuesto de su área: el sueldo, la renta, el internet
 * y la luz ya vienen puestos (lo que ya se conoce), y las suscripciones y demás las escribe ella. Lo manda a
 * contabilidad (director financiero), que lo aprueba o lo devuelve con un comentario. Al aprobarse, entra al
 * presupuesto mensual del lugar y desde ahí se da seguimiento contra lo que gasta.
 * Hojas: «Presupuesto de área» (los renglones) y «Presupuesto de área estado» (en qué va cada mes). */
var H_PAREA = ['Mes', 'Lugar', 'Concepto', 'Rubro', 'Tipo', 'Monto', 'Nota', 'Orden'];
var H_PAREA_EST = ['Mes', 'Lugar', 'Estado', 'Enviado por', 'Enviado en', 'Resuelto por', 'Resuelto en', 'Comentario'];
var PA_EST = { BOR: 'Borrador', POR: 'Por aprobar', APR: 'Aprobado', DEV: 'Devuelto' };
var PA_TIPOS = ['Fijo', 'Suscripción', 'Otro'];
var PA_DIA_AVISO = 15, PA_DIA_LIMITE = 20;     // del mes anterior: desde el 15 se le recuerda; a más tardar el 20 lo manda
/** Hasta cuándo debe estar enviado el presupuesto de un mes: el 20 del mes anterior; el del mes en curso, los primeros 5 días. */
function limitePArea(mes) { return mes === hoyISO().slice(0, 7) ? mes + '-05' : mesMas(mes, -1) + '-' + PA_DIA_LIMITE; }
function mesMas(mes, n) { var y = +mes.slice(0, 4), m = +mes.slice(5, 7) - 1 + n; y += Math.floor(m / 12); m = ((m % 12) + 12) % 12; return y + '-' + ('0' + (m + 1)).slice(-2); }
function hojaMesTexto(nombre, enc) {
  var h = libro().getSheetByName(nombre);
  if (h) return h;
  h = hojaCat(nombre, enc); h.getRange('A:A').setNumberFormat('@');          // el mes «2026-10» se queda como texto
  return h;
}
function hojaPArea() { return hojaMesTexto('Presupuesto de área', H_PAREA); }
function hojaPAreaEst() { return hojaMesTexto('Presupuesto de área estado', H_PAREA_EST); }
function filasPArea(mes, lugar) {
  return leeTodo(hojaPArea(), H_PAREA.length).map(function (r, i) {
    return { fila: i + 2, mes: String(r[0] || '').slice(0, 7), lugar: String(r[1] || ''), concepto: String(r[2] || ''),
      rubro: String(r[3] || ''), tipo: String(r[4] || ''), monto: Number(r[5]) || 0, nota: String(r[6] || ''), orden: Number(r[7]) || 0 };
  }).filter(function (x) { return x.mes && (!mes || x.mes === mes) && (!lugar || x.lugar === lugar); })
    .sort(function (a, b) { return a.orden - b.orden; });
}
function filasPAreaEst() {
  return leeTodo(hojaPAreaEst(), H_PAREA_EST.length).map(function (r, i) {
    return { fila: i + 2, mes: String(r[0] || '').slice(0, 7), lugar: String(r[1] || ''), estado: String(r[2] || ''),
      por: String(r[3] || ''), en: fmtSello(r[4]), resPor: String(r[5] || ''), resEn: fmtSello(r[6]), comentario: String(r[7] || '') };
  }).filter(function (x) { return x.mes; });
}
function estadoPArea(mes, lugar) {
  var e = null; filasPAreaEst().forEach(function (x) { if (x.mes === mes && x.lugar === lugar) e = x; });
  return e;
}
function ponEstadoPArea(mes, lugar, fila) {
  var h = hojaPAreaEst(), e = estadoPArea(mes, lugar);
  if (e) h.getRange(e.fila, 1, 1, H_PAREA_EST.length).setValues([fila]);
  else h.appendRow(fila);
}
function sueldoCmo() { return ajustesInv().sueldoCmo; }
/** Lo último que se conoce de un rubro en un lugar: el presupuesto aprobado más reciente o, si no hay, lo gastado el mes pasado. */
function montoConocido(lugarNombre, rubro, mes) {
  var m = '', v = 0;
  filasPres().forEach(function (x) { if (x.unidad === lugarNombre && x.rubro === rubro && x.mes < mes && x.mes > m && x.monto > 0) { m = x.mes; v = x.monto; } });
  if (v > 0) return v;
  try { var g = gastoRealMes(mesAnt(mes)).gasto[lugarNombre]; if (g && g[rubro] > 0) return r2(g[rubro]); } catch (e) {}
  return 0;
}
function lineasBasePArea(un, mes) {
  var nom = un.nombre;
  return [
    { concepto: 'Sueldo de la CMO', rubro: 'Nómina y bonificaciones', tipo: 'Fijo', monto: sueldoCmo(), nota: 'Pago mensual', base: true },
    { concepto: 'Renta de la oficina', rubro: 'Renta', tipo: 'Fijo', monto: montoConocido(nom, 'Renta', mes), nota: '', base: true },
    { concepto: 'Internet', rubro: 'Internet y teléfono', tipo: 'Fijo', monto: montoConocido(nom, 'Internet y teléfono', mes), nota: '', base: true },
    { concepto: 'Luz', rubro: 'Energía eléctrica', tipo: 'Fijo', monto: montoConocido(nom, 'Energía eléctrica', mes), nota: '', base: true }
  ];
}
/** Qué mes le toca armar: el actual si todavía no está aprobado; si ya lo está, el que sigue. */
function mesDeTrabajoPArea(un) {
  var act = hoyISO().slice(0, 7), e = estadoPArea(act, un.nombre);
  return e && e.estado === PA_EST.APR ? mesMas(act, 1) : act;
}
function exigeCmo(cred) {
  permiteCmo();
  var yo = quien(cred);
  if (!yo.cmo) throw new Error('El presupuesto del área lo arma la CMO.');
  var un = unidadPorId(yo.sucursal);
  if (!un) throw new Error('Su usuario no tiene lugar de trabajo. Pídale al administrador que lo ponga.');
  return { yo: yo, un: un };
}
function miPresupuesto(cred, mes) {
  var c = exigeCmo(cred), un = c.un, hoy = hoyISO();
  mes = /^\d{4}-\d{2}$/.test(String(mes || '')) ? mes : mesDeTrabajoPArea(un);
  var guardadas = filasPArea(mes, un.nombre), est = estadoPArea(mes, un.nombre);
  var lineas = guardadas.length ? guardadas.map(function (x) { return { concepto: x.concepto, rubro: x.rubro, tipo: x.tipo, monto: x.monto, nota: x.nota }; })
    : lineasBasePArea(un, mes);
  var total = r2(lineas.reduce(function (a, x) { return a + (Number(x.monto) || 0); }, 0));
  var estado = est ? est.estado : (guardadas.length ? PA_EST.BOR : 'Sin armar');
  var limite = limitePArea(mes), aviso = mesMas(mes, -1) + '-' + PA_DIA_AVISO;
  var act = hoy.slice(0, 7);
  var tarde = estado !== PA_EST.POR && estado !== PA_EST.APR && hoy > limite;
  return { mes: mes, mesTxt: periodoTxt(mes), lugar: { id: un.id, nombre: un.nombre }, estado: estado, lineas: lineas, esBase: !guardadas.length,
    total: total, comentario: est ? est.comentario : '', enviadoPor: est ? est.por : '', enviadoEn: est ? est.en : '',
    resueltoPor: est ? est.resPor : '', resueltoEn: est ? est.resEn : '',
    puedeEditar: estado !== PA_EST.POR, limite: limite, diaLimite: PA_DIA_LIMITE, desdeAviso: aviso, tarde: tarde,
    cuando: 'Mándelo a contabilidad a más tardar el ' + dia(limite).slice(0, 5) + ' (el presupuesto de cada mes se manda el ' + PA_DIA_LIMITE + ' del mes anterior): así lo aprueban a tiempo y se programan los pagos fijos del día 1.',
    meses: [act, mesMas(act, 1)].map(function (m) { var e2 = estadoPArea(m, un.nombre);
      return { mes: m, txt: periodoTxt(m), estado: e2 ? e2.estado : (filasPArea(m, un.nombre).length ? PA_EST.BOR : 'Sin armar') }; }),
    tipos: PA_TIPOS, rubros: rubros(), sueldo: sueldoCmo() };
}
function limpiaLineasPArea(lineas) {
  var rs = rubros(), out = [];
  (lineas || []).forEach(function (l, i) {
    var concepto = String(l.concepto || '').trim().replace(/\s+/g, ' ').slice(0, 60);
    var monto = r2(String(l.monto == null ? '' : l.monto).replace(/[Q,\s]/g, ''));
    if (!concepto && !(monto > 0)) return;            // renglón vacío
    if (concepto.length < 3) throw new Error('Renglón ' + (i + 1) + ': escriba qué es (por ejemplo «Adobe Creative Cloud»).');
    if (!(monto > 0) || monto > 1000000) throw new Error('«' + concepto + '»: escriba el monto del mes.');
    var rubro = rs.indexOf(l.rubro) >= 0 ? l.rubro : 'Mercadeo y publicidad';
    out.push({ concepto: concepto, rubro: rubro, tipo: PA_TIPOS.indexOf(l.tipo) >= 0 ? l.tipo : 'Otro', monto: monto,
      nota: String(l.nota || '').trim().slice(0, 100) });
  });
  if (out.length > 40) throw new Error('Son muchos renglones (máximo 40).');
  return out;
}
/** Guarda el presupuesto del mes. Con enviar = true, además lo manda a contabilidad. */
function guardarMiPresupuesto(cred, mes, lineas, enviar) {
  var c = exigeCmo(cred), yo = c.yo, un = c.un;
  if (!/^\d{4}-\d{2}$/.test(String(mes || ''))) throw new Error('Escoja el mes.');
  var act = hoyISO().slice(0, 7);
  if (mes < act) throw new Error('Ese mes ya pasó. Arme el de este mes o el que sigue.');
  var est = estadoPArea(mes, un.nombre);
  if (est && est.estado === PA_EST.POR) throw new Error('Ese presupuesto ya está en contabilidad esperando su aprobación.');
  var ls = limpiaLineasPArea(lineas);
  if (!ls.length) throw new Error('Agregue al menos un renglón.');
  var lock = LockService.getScriptLock();
  lock.waitLock(15000); _LEE = {};
  try {
    var h = hojaPArea(), viejas = filasPArea(mes, un.nombre);
    for (var i = viejas.length - 1; i >= 0; i--) h.deleteRow(viejas[i].fila);
    var f0 = h.getLastRow() + 1;
    h.getRange(f0, 1, ls.length, 1).setNumberFormat('@');
    h.getRange(f0, 1, ls.length, H_PAREA.length).setValues(ls.map(function (x, k) { return [mes, un.nombre, x.concepto, x.rubro, x.tipo, x.monto, x.nota, k + 1]; }));
    ponEstadoPArea(mes, un.nombre, [mes, un.nombre, enviar ? PA_EST.POR : PA_EST.BOR, enviar ? yo.nombre : '', enviar ? new Date() : '', '', '', '']);
  } finally { lock.releaseLock(); }
  var enviado = false;
  if (enviar) {
    var total = r2(ls.reduce(function (a, x) { return a + x.monto; }, 0));
    nombresDe(function (u) { return u.rol === 'finanzas'; }).forEach(function (n) {
      try { correoAviso(n, { id: 'pa-' + un.id + '-' + mes + '-' + Date.now(), tipo: 'Presupuesto de área por aprobar',
        titulo: un.nombre + ' · ' + periodoTxt(mes) + ' · ' + dinero(total), texto: yo.nombre + ' mandó el presupuesto de su área para que lo apruebe.' }); } catch (e) {}
    });
    enviado = true;
  }
  var r = miPresupuesto(cred, mes); r.ok = true;
  r.mensaje = enviado ? 'Presupuesto de ' + periodoTxt(mes) + ' enviado a contabilidad. Le avisamos cuando lo aprueben.' : 'Borrador guardado. Cuando esté listo, mándelo a contabilidad.';
  return r;
}
/** El avance del gasto de un lugar en el mes: contra lo aprobado, lo ya validado y lo que falta validar. */
function avanceArea(lugarId, mes) {
  var un = unidadPorId(lugarId); if (!un) return null;
  mes = mes || hoyISO().slice(0, 7);
  var est = estadoPArea(mes, un.nombre), apr = est && est.estado === PA_EST.APR;
  var pres = {};
  if (apr) filasPArea(mes, un.nombre).forEach(function (x) { pres[x.rubro] = (pres[x.rubro] || 0) + x.monto; });
  var real = {}; try { real = gastoRealMes(mes).gasto[un.nombre] || {}; } catch (e) {}
  var pend = {}, nPend = 0, mPend = 0;
  filasPorConfirmar().forEach(function (x) {
    if (x.tipo !== 'gasto' || x.sucursal !== un.nombre || x.estado !== EST_PEND || String(x.fecha).slice(0, 7) !== mes) return;
    nPend++; mPend += x.monto;
  });
  var nombres = {}; Object.keys(pres).forEach(function (k) { nombres[k] = 1; }); Object.keys(real).forEach(function (k) { if (real[k] > 0) nombres[k] = 1; });
  var rubs = Object.keys(nombres).map(function (k) { var p = r2(pres[k] || 0), g = r2(real[k] || 0);
    return { rubro: k, presupuesto: p, gastado: g, pct: p ? Math.round(g / p * 100) : null }; })
    .sort(function (a, b) { return b.presupuesto - a.presupuesto || b.gastado - a.gastado; });
  var tp = r2(rubs.reduce(function (a, x) { return a + x.presupuesto; }, 0)), tg = r2(rubs.reduce(function (a, x) { return a + x.gastado; }, 0));
  return { mes: mes, mesTxt: periodoTxt(mes), lugar: un.nombre, aprobado: !!apr, estado: est ? est.estado : 'Sin armar',
    rubros: rubs, presupuesto: tp, gastado: tg, pct: tp ? Math.round(tg / tp * 100) : null,
    porValidar: nPend, porValidarMonto: r2(mPend) };
}
/** Lo que ve la CMO en vez de «lo enviado»: el avance de su gasto contra su presupuesto, y cómo va cada solicitud. */
function miAvance(cred, mes) {
  var c = exigeCmo(cred), act = hoyISO().slice(0, 7);
  mes = /^\d{4}-\d{2}$/.test(String(mes || '')) ? mes : act;
  var av = avanceArea(c.un.id, mes);
  var sol = filasFondos().filter(function (x) { return x.por === c.yo.nombre; }).reverse().slice(0, 12)
    .map(function (x) { return { numero: x.numero, tipo: x.tipo, que: x.que, monto: x.monto, estado: x.estado, en: x.en, pagadoEn: x.pagEn, pagado: x.pagado }; });
  var sig = mesMas(act, 1), es = estadoPArea(sig, c.un.nombre), ea = estadoPArea(act, c.un.nombre);
  return { avance: av, solicitudes: sol, presupuestos: [
    { mes: act, txt: periodoTxt(act), estado: ea ? ea.estado : 'Sin armar' }, { mes: sig, txt: periodoTxt(sig), estado: es ? es.estado : 'Sin armar' } ] };
}
/** Para contabilidad y el administrador: los presupuestos de áreas (oficinas) de un mes. */
function presupuestosAreas(cred, mes) {
  var yo = quien(cred);
  if (!(yo.esAdmin || yo.rol === 'finanzas' || esContador(yo) || yo.rol === 'dueno')) throw new Error('Esto lo ven contabilidad y el administrador.');
  var act = hoyISO().slice(0, 7), ests = filasPAreaEst();
  if (!/^\d{4}-\d{2}$/.test(String(mes || ''))) {
    var pend = ests.filter(function (x) { return x.estado === PA_EST.POR; }).map(function (x) { return x.mes; }).sort();
    mes = pend.length ? pend[0] : act;
  }
  var areas = UNIDADES.filter(function (u) { return u.extra && u.activo !== false; });
  var cmos = usuariosCache().filter(function (u) { return u.activo && u.rol === 'cmo'; });
  var lista = areas.map(function (u) {
    var e = estadoPArea(mes, u.nombre), ls = filasPArea(mes, u.nombre);
    var resp = cmos.filter(function (x) { return x.sucursal === u.id; }).map(function (x) { return x.nombre; });
    var tot = r2(ls.reduce(function (a, x) { return a + x.monto; }, 0));
    var av = null; try { av = avanceArea(u.id, mes); } catch (e2) {}
    return { id: u.id, nombre: u.nombre, responsables: resp, estado: e ? e.estado : (ls.length ? PA_EST.BOR : 'Sin armar'),
      total: tot, lineas: ls.map(function (x) { return { concepto: x.concepto, rubro: x.rubro, tipo: x.tipo, monto: x.monto, nota: x.nota }; }),
      enviadoPor: e ? e.por : '', enviadoEn: e ? e.en : '', resueltoPor: e ? e.resPor : '', resueltoEn: e ? e.resEn : '', comentario: e ? e.comentario : '',
      puedeAprobar: (yo.rol === 'finanzas' || yo.esAdmin) && !!e && e.estado === PA_EST.POR, avance: av };
  });
  return { mes: mes, mesTxt: periodoTxt(mes), areas: lista, hay: lista.length > 0,
    puedeAprobar: yo.rol === 'finanzas' || yo.esAdmin, antes: mesMas(mes, -1), despues: mesMas(mes, 1) };
}
function aprobarPresupuestoArea(cred, mes, lugarId, aprobar, comentario) {
  var yo = quien(cred);
  if (!(yo.rol === 'finanzas' || yo.esAdmin)) throw new Error('Los presupuestos de las áreas los aprueba el director financiero.');
  var un = unidadPorId(lugarId);
  if (!un || !un.extra) throw new Error('No se encontró ese lugar.');
  var e = estadoPArea(mes, un.nombre);
  if (!e || e.estado !== PA_EST.POR) throw new Error('Ese presupuesto no está esperando aprobación.');
  comentario = String(comentario || '').trim().slice(0, 200);
  if (!aprobar && comentario.length < 4) throw new Error('Escriba qué debe corregir (se lo mandamos a la CMO).');
  var ls = filasPArea(mes, un.nombre);
  var lock = LockService.getScriptLock();
  lock.waitLock(15000); _LEE = {};
  try {
    if (aprobar) {
      // pasa al presupuesto mensual del lugar, por rubro: ahí se da seguimiento contra lo gastado
      var hp = hojaPres(), viejas = filasPres(mes).filter(function (x) { return x.unidad === un.nombre; });
      for (var i = viejas.length - 1; i >= 0; i--) hp.deleteRow(viejas[i].fila);
      var por = {}; ls.forEach(function (x) { por[x.rubro] = (por[x.rubro] || 0) + x.monto; });
      var filas = Object.keys(por).map(function (k) { return [mes, un.nombre, k, r2(por[k]), yo.nombre, new Date()]; });
      if (filas.length) {
        var f0 = hp.getLastRow() + 1;
        hp.getRange(f0, 1, filas.length, 1).setNumberFormat('@');
        hp.getRange(f0, 1, filas.length, H_PRES.length).setValues(filas);
      }
    }
    ponEstadoPArea(mes, un.nombre, [mes, un.nombre, aprobar ? PA_EST.APR : PA_EST.DEV, e.por, e.en, yo.nombre, new Date(), comentario]);
  } finally { lock.releaseLock(); }
  if (e.por) { try { correoAviso(e.por, { id: 'pa-r-' + un.id + '-' + mes + '-' + Date.now(), tipo: aprobar ? 'Presupuesto aprobado' : 'Presupuesto devuelto',
    titulo: periodoTxt(mes), texto: aprobar ? 'Contabilidad aprobó el presupuesto de ' + un.nombre + '.' : 'Contabilidad lo devolvió: ' + comentario }); } catch (x) {} }
  var r = presupuestosAreas(cred, mes); r.ok = true;
  r.mensaje = aprobar ? 'Presupuesto de ' + un.nombre + ' aprobado. Ya cuenta en el presupuesto mensual.' : 'Devuelto a ' + (e.por || 'la CMO') + ' con su comentario.';
  return r;
}

/* ════════════ PEDIDOS DE EMERGENCIA (entre sucursales o a bodega) ════════════
 * Si a una sucursal se le acaba algo, le pide a otra sucursal o a la bodega. El que
 * da (el gerente de esa sucursal o Samuel) envía lo que puede, el que pide lo recibe,
 * y el inventario y el costo pasan de uno a otro. */
var H_EMERG = ['No.', 'Pedida en', 'Pedida por', 'Pide', 'Da', 'Código', 'Producto', 'Unidad', 'Pedido',
  'Enviado', 'Recibido', 'Costo unitario', 'Estado', 'Enviada por', 'Enviada en', 'Recibida por', 'Recibida en', 'Nota'];
var EM_EST = { PED: 'Pedida', ENV: 'Enviada', REC: 'Recibida', RECH: 'Rechazada' };
MOV.SAL_EM = 'Salida por emergencia';
MOV.ENT_EM = 'Entrada por emergencia';
function hojaEmerg(ss) {
  ss = ss || libro();
  var h = ss.getSheetByName('Pedidos de emergencia');
  if (!h) { h = hojaLimpia(ss, 'Pedidos de emergencia', H_EMERG); h.getRange('B:B').setNumberFormat('dd/mm/yyyy hh:mm');
    h.getRange('O:O').setNumberFormat('dd/mm/yyyy hh:mm'); h.getRange('Q:Q').setNumberFormat('dd/mm/yyyy hh:mm'); h.setColumnWidth(7, 260); }
  return h;
}
function filasEmerg() {
  return leeTodo(hojaEmerg(), H_EMERG.length).map(function (r, i) {
    return { fila: i + 2, numero: String(r[0] || ''), en: fmtSello(r[1]), por: String(r[2] || ''), pide: String(r[3] || ''),
      da: String(r[4] || ''), codigo: String(r[5] || ''), nombre: String(r[6] || ''), unidad: String(r[7] || ''),
      pedido: Number(r[8]) || 0, enviado: r[9] === '' || r[9] == null ? null : Number(r[9]),
      recibido: r[10] === '' || r[10] == null ? null : Number(r[10]), costo: Number(r[11]) || 0, estado: String(r[12] || ''),
      envPor: String(r[13] || ''), envEn: fmtSello(r[14]), recPor: String(r[15] || ''), recEn: fmtSello(r[16]), nota: String(r[17] || '') };
  }).filter(function (x) { return x.numero; });
}
function emergAgrupadas() {
  var por = {}, orden = [];
  filasEmerg().forEach(function (x) {
    if (!por[x.numero]) { por[x.numero] = { numero: x.numero, en: x.en, por: x.por, pide: x.pide, da: x.da, estado: x.estado,
      envPor: x.envPor, envEn: x.envEn, recPor: x.recPor, recEn: x.recEn, nota: x.nota, lineas: [] }; orden.push(por[x.numero]); }
    por[x.numero].lineas.push({ codigo: x.codigo, nombre: x.nombre, unidad: x.unidad, pedido: x.pedido, enviado: x.enviado,
      recibido: x.recibido, costo: x.costo, fila: x.fila });
  });
  return orden;
}
/** Qué unidad es la mía para emergencias: el gerente su sucursal; la bodega, la bodega. */
function unidadEmerg(yo) {
  if (yo.rol === 'gerente') return nombreUnidad(yo.sucursal);
  if (yo.rol === 'bodega') return nombreBodega();
  return '';
}
function pantallaEmergencias(cred) {
  var yo = quien(cred), mia = unidadEmerg(yo);
  if (!mia && !veInv(yo)) throw new Error('Su usuario no maneja pedidos de emergencia.');
  var desde = masDias(hoyISO(), -30);
  var ls = emergAgrupadas().filter(function (e) {
    var abierta = e.estado === EM_EST.PED || e.estado === EM_EST.ENV;
    return (!mia || e.pide === mia || e.da === mia) && (abierta || e.en.slice(0, 10) >= desde); }).reverse();
  ls.forEach(function (e) {
    e.puedeEnviar = !!mia && e.da === mia && e.estado === EM_EST.PED;
    e.puedeRecibir = !!mia && e.pide === mia && e.estado === EM_EST.ENV;
  });
  var out = { mia: mia, pedidos: ls, origenes: [] , productos: [] };
  if (yo.rol === 'gerente') {
    out.origenes = [nombreBodega()].concat(sucursalesDestino().map(function (u) { return u.nombre; }).filter(function (n) { return n !== mia; }));
    // La lista sale al instante: lo que hay en cada lugar se pide aparte (existenciasEmergencia) porque leer los conteos es lo lento.
    out.productos = productosInv().filter(function (b) { return maneja(b, yo.sucursal); }).map(function (b) {
      return { codigo: b.codigo, nombre: b.nombre, unidad: b.unidad, grupo: b.grupo }; });
  }
  return out;
}
/** Lo que hay en la bodega o en la otra sucursal a la que se le pide (se guarda 2 minutos para que no sea lento). */
function existenciasEmergencia(cred, origen) {
  var yo = quien(cred);
  if (yo.rol !== 'gerente') throw new Error('Los pedidos de emergencia los hace el gerente de la sucursal.');
  var mia = nombreUnidad(yo.sucursal), u = UNIDADES.filter(function (x) { return x.nombre === String(origen); })[0];
  if (!u || u.nombre === mia) throw new Error('Escoja la bodega u otra sucursal.');
  var llave = 'emx_' + u.id, cache = null, hay = null;
  try { cache = CacheService.getScriptCache(); var c = cache.get(llave); if (c) hay = JSON.parse(c); } catch (e) {}
  if (!hay) {
    var ex = {};
    try { ex = estadoInventario(u.id).existencia; } catch (e) { ex = {}; }
    hay = {};
    Object.keys(ex).forEach(function (k) { hay[k] = r3(ex[k] || 0); });
    try { if (cache) cache.put(llave, JSON.stringify(hay), 120); } catch (e) {}
  }
  return { ok: true, origen: u.nombre, hay: hay };
}
function pedirEmergencia(cred, d) {
  var yo = quien(cred);
  if (yo.rol !== 'gerente') throw new Error('Los pedidos de emergencia los hace el gerente de la sucursal.');
  d = d || {};
  var mia = nombreUnidad(yo.sucursal), da = String(d.da || '');
  var uDa = UNIDADES.filter(function (u) { return u.nombre === da; })[0];
  if (!uDa || da === mia) throw new Error('Escoja a quién le pide: la bodega u otra sucursal.');
  var bod = {}; productosInv().forEach(function (b) { bod[b.codigo] = b; });
  var filas = [], ahora = new Date(), idx = indiceCompras();
  Object.keys(d.cantidades || {}).forEach(function (c) {
    var q = r3(Number(String(d.cantidades[c]).replace(',', '.')));
    if (!(q > 0)) return;
    var b = bod[c]; if (!b) throw new Error('No se encontró el producto ' + c + '.');
    filas.push([null, ahora, yo.nombre, mia, da, b.codigo, b.nombre, b.unidad, q, '', '', costoBodega(b, idx).costo, EM_EST.PED, '', '', '', '', String(d.nota || '').slice(0, 200)]);
  });
  if (!filas.length) throw new Error('Escriba cuánto necesita de al menos un producto.');
  var lock = LockService.getScriptLock();
  lock.waitLock(15000); _LEE = {};
  var numero;
  try {
    var max = 0;
    filasEmerg().forEach(function (x) { var m = /^EM-(\d+)$/.exec(x.numero); if (m) max = Math.max(max, +m[1]); });
    numero = 'EM-' + ('0000' + (max + 1)).slice(-4);
    filas.forEach(function (f) { f[0] = numero; });
    var h = hojaEmerg();
    h.getRange(h.getLastRow() + 1, 1, filas.length, H_EMERG.length).setValues(filas);
  } finally { lock.releaseLock(); }
  var r = pantallaEmergencias(cred); r.ok = true;
  r.mensaje = numero + ' enviado a ' + da + ': ' + filas.length + (filas.length === 1 ? ' producto.' : ' productos.') + ' Le avisamos.';
  return r;
}
/** El que da manda lo que puede (o rechaza). Sale de su inventario. */
function enviarEmergencia(cred, numero, enviados, rechazar, motivo) {
  var yo = quien(cred), mia = unidadEmerg(yo);
  var e = emergAgrupadas().filter(function (x) { return x.numero === String(numero); })[0];
  if (!e) throw new Error('No se encontró el pedido ' + numero + '.');
  if (!mia || e.da !== mia) throw new Error('Ese pedido no es para usted.');
  if (e.estado !== EM_EST.PED) throw new Error('Ese pedido ya está ' + e.estado.toLowerCase() + '.');
  var h = hojaEmerg(), ahora = new Date();
  if (rechazar) {
    if (String(motivo || '').trim().length < 4) throw new Error('Escriba por qué no lo puede mandar.');
    e.lineas.forEach(function (l) { h.getRange(l.fila, 13, 1, 3).setValues([[EM_EST.RECH, yo.nombre, ahora]]);
      h.getRange(l.fila, 18).setValue([e.nota, 'No se mandó: ' + String(motivo).trim()].filter(Boolean).join(' · ').slice(0, 200)); });
    var rr = pantallaEmergencias(cred); rr.ok = true; rr.mensaje = numero + ' rechazado. ' + e.pide + ' ya lo sabe.';
    return rr;
  }
  enviados = enviados || {};
  var kardex = [], n = 0, uDa = UNIDADES.filter(function (u) { return u.nombre === e.da; })[0];
  e.lineas.forEach(function (l) {
    var crudo = enviados[l.codigo], q = crudo == null || String(crudo).trim() === '' ? l.pedido : r3(Number(String(crudo).replace(',', '.')));
    if (!isFinite(q) || q < 0) throw new Error('La cantidad de «' + l.nombre + '» no es válida.');
    h.getRange(l.fila, 10).setValue(q);
    h.getRange(l.fila, 13, 1, 3).setValues([[EM_EST.ENV, yo.nombre, ahora]]);
    if (q > 0) { n++; kardex.push([ahora, uDa.nombre, l.codigo, l.nombre, MOV.SAL_EM, -q, numero, yo.nombre, 'A ' + e.pide]); }
  });
  if (!n) throw new Error('Si no puede mandar nada, use «No puedo».');
  escribeKardex(kardex);
  var r = pantallaEmergencias(cred); r.ok = true; r.mensaje = numero + ' enviado a ' + e.pide + '. Cuando llegue, lo recibe.';
  return r;
}
/** El que pidió recibe: entra a su inventario y el costo pasa de una unidad a la otra. */
function recibirEmergencia(cred, numero, recibidos) {
  var yo = quien(cred), mia = unidadEmerg(yo);
  var e = emergAgrupadas().filter(function (x) { return x.numero === String(numero); })[0];
  if (!e) throw new Error('No se encontró el pedido ' + numero + '.');
  if (!mia || e.pide !== mia) throw new Error('Ese pedido lo recibe ' + e.pide + '.');
  if (e.estado !== EM_EST.ENV) throw new Error('Ese pedido todavía no se ha enviado.');
  recibidos = recibidos || {};
  var h = hojaEmerg(), ahora = new Date(), hoy = hoyISO(), kardex = [], lineas = [], bod = {};
  filasBodega().forEach(function (b) { bod[b.codigo] = b; });
  var uPide = UNIDADES.filter(function (u) { return u.nombre === e.pide; })[0];
  e.lineas.forEach(function (l) {
    var crudo = recibidos[l.codigo], q = crudo == null || String(crudo).trim() === '' ? (l.enviado || 0) : r3(Number(String(crudo).replace(',', '.')));
    if (!isFinite(q) || q < 0) throw new Error('La cantidad de «' + l.nombre + '» no es válida.');
    h.getRange(l.fila, 11).setValue(q);
    h.getRange(l.fila, 13).setValue(EM_EST.REC);
    h.getRange(l.fila, 16, 1, 2).setValues([[yo.nombre, ahora]]);
    if (q > 0) {
      kardex.push([ahora, uPide.nombre, l.codigo, l.nombre, MOV.ENT_EM, q, numero, yo.nombre, 'De ' + e.da]);
      var b = bod[l.codigo] || {};
      lineas.push({ producto: l.nombre, cantidad: q, medida: l.unidad, precio: l.costo, subtotal: r2(q * l.costo),
        categoria: b.categoria || 'Insumos de cocina' });
    }
  });
  escribeKardex(kardex);
  if (lineas.length && lineas.some(function (l) { return l.subtotal > 0; })) {
    if (e.da === nombreBodega()) creaTraslado({ s: uPide, lineas: lineas }, yo, ahora, hoy, numero, numero);
    else creaPrestamo(e, uPide, lineas, yo, ahora, hoy);
  }
  var r = pantallaEmergencias(cred); r.ok = true; r.mensaje = numero + ' recibido. Entró a su inventario.';
  return r;
}
/** Entre sucursales: el costo se suma a la que recibe y se resta de la que dio. */
function creaPrestamo(e, uPide, lineas, yo, ahora, hoy) {
  var h = libro().getSheetByName('Gastos');
  var id = 'E' + Utilities.formatDate(ahora, ZONA, 'yyyyMMddHHmmss') + Math.floor(Math.random() * 900 + 100);
  var pago = aDate(hoy), mes = hoy.slice(0, 7), monto = r2(lineas.reduce(function (a, l) { return a + l.subtotal; }, 0));
  var f0 = h.getLastRow() + 1;
  h.getRange(f0, 8, 2, 1).setNumberFormat('@');
  h.getRange(f0, 1, 2, H_GASTOS.length).setValues(aAncho([
    [id, ahora, yo.nombre, '', uPide.nombre, categoriaDe(lineas), pago, mes, monto, 'Emergencia ' + e.numero + ' de ' + e.da,
      'No', '', '', '', e.da, e.numero, lineas.length],
    [id + '-O', ahora, yo.nombre, '', e.da, categoriaDe(lineas), pago, mes, -monto, 'Emergencia ' + e.numero + ': salió a ' + uPide.nombre,
      'No', '', '', '', uPide.nombre, e.numero, lineas.length]], H_GASTOS.length));
  var det = [];
  lineas.forEach(function (l) { det.push([id, ahora, yo.nombre, uPide.nombre, e.da, pago, mes, l.categoria, l.producto, l.cantidad, l.medida, l.precio, l.subtotal, 'No', '', '', '', '']); });
  lineas.forEach(function (l) { det.push([id + '-O', ahora, yo.nombre, e.da, uPide.nombre, pago, mes, l.categoria, l.producto, -l.cantidad, l.medida, l.precio, -l.subtotal, 'No', '', '', '', '']); });
  escribeDetalle(det);
}

/** La corona, para los PDF y los correos. */
var LOGO_PNG = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAFIAAAA4CAYAAABuQ6+nAAAYRElEQVR42uWcebwdVZXvv2tX1ZnPPffeDCQQQsI8ZUCmSIAQbGxoVOAp8HzyWuVpt93aDo1Dq90t2gqvnW2HVmy7+wnqI7a2PmwhAZIwSYQACXNIQnIzD3e+55w6p6r2Xu+POvfmnpsE4/N9Png/XZ9PZahzqmrt317Db62194Hfg0OXLfPG/l3fOVujXTdpc+e92ti+S8NtTQ17Im1s36PRzlUa7fyEhttOHPu+Hrj31TzkVQdR1RMRqwNPdVKa8UlU30mQnQIWkhisTb/oGfAzgAdxcwgxtxMOf0Y6Tt4/+oz/tECOgVjbfC7Z4g/xSicSD0KS2JZs4hRBwIAyenqeR6YTkto24pG3SuGkX73aYMqrCKIREae1LYvI5FfgmzJhPQZ8B6LO4fk+BC3LjS02SRBjMIiCJuTyAZYa8fBlUjjp0dFnvhrjMa8WiHAzqvtm4md+hpEyYT3BmMA5FZMN8IoVkjhhz/Z97Nm+jySxeMUKJhvgnBOMCQjrFkMRk/+JDu+alj77U+Y/jUaqLvNErrNa3/rv5Luupt6f4Hm+TSxescT2LTv41m0/Y9WD69jfOwjAtKmdLL34LN77p1cxa84x2FoNz/fA2oTCFJ9G3/clP/ftr5aJy6vmF8Ptl5DLryKsWkQ8Zx2mWGLF3Y/wng98mT17+ykWcvh+atpJYqnVG8w8qptv/8NNXHb5BbhaFeMZELEY3xA1z6Y4dx0sMyLX2d8701ZVaZ1GVX9X8DX9034MjKLgrMMUCqxds5633fhZqtWQ6dO6yGUzeJ6H53nkshmmT+tipBryths/yxNrnsYUCjjrwDklKAgiHxNB4drfdbJF9VNmdNy/k0amD1jtwX491OyqrvLhEvfbOPexABP2HI8xL6Lq4xwqIhbDFW/6MI8/9RKVjmIrcB98+L7H0FCN884+hV/+/At4OERVMQbERDh7uuSPe/m3DTyp315t4BIrIjrRFcE0OdRnY3K9kvkBSev/WdjfDUEWvKpIR6/I0tHPPMAd7gUTLUBVlUbPFWQ6AsKBxDr1vWKBx1av5fGnXqKjXGgDUdrUODXxjo4Cjz25gcd//QyvXXIOtl4Tz7mEfGeWRt87gL9tWZs7MgChBbpLrw10QrUEmQSmD4hI8xDYHB7IUTVOCfKWTsq5N+P0TUQ75uPcNGgEiNS1uXMTxizH2ttFZMP4APIb5HYiolrvubo1R6KqgMc9960ljhNE2o3EtRAcf1lEiOOEu+9by2uXnI86BU8MGoLjOtVVnwHskfprAG1uX4AJrsclr6NZnYu6EjQtZkevRrvWo3IX9cZPRGRwHE56kGmrfsqIfDqdkWjHe8H7K4L8LHCgTUgSUAUzmmFkIa41UP6JxuDNUjm9T1et8mVpqqmHNev65tmYzAaEHNYqnpEoTljyhx9i46ad5PMZXAs9p1DIpkpVbxpMS1pjhDCMOPnEWTyw/CsEgYdah4ASZAWNz5PM7McPN7ktIEREnI5snk+28Deou4ZMyYMIbATWpbPneyDZVMGTxnbUfl4yx3xj/JjGgo1+KgVR+9dWNNr1c4LKNzA6i3AwIRy2NJrORona2KptxurqdUc4kOCSHJnS+yhMeVzDrZfK0qVJ6jtfIbCZzNVkO3JYm1inopkcax59jhc2bCOfz46B6BlluOZz5aJ+/mhRP8M1H8+0AHZKPp/l+Q09PLrmWTSTG73P4hfAumvSV14rrwhiY9tHyRUfI8i9BZd4hAOJq9ecbURq40RtFCuNpiMctoSDCeKOJej4ujZ3/kIHtnSKiBvlrWkUvvlmVPeWKMy4h6D8JsK+mChSRXzrnIcxxivmxCsWxCtkxRgx1qmPc0q9L8HTufi5e7W69V0iSxPVtcEhol0LheTtaJTaJ2n+98N/W0mS2HazFnAqXHVhL1ct7sWptIVGESFJLD/68ao0h0wfb3AhqFyvqj6rVzNejlZA8VLL6Pke2e6/16SZpT5oVRXrnG+MGK+QTcdazAnGGOucp4hPFClhX0ymfCWFzArt31xJEwsVHzAiYrX+8vfJT1tE2BchklGniCd4+Qq2PkLPCz2MVOtUKiXmHDcDr9gBjaoo+NJoOIwHxY7vam1rVmTON1vUwU9HuNEXkabWXr6KfPk1NGrWOfVMLse2jVv45fLH6CgXsK0ChQg0I8PcGQ3OO3UYFObMaNA/7BP4iio4a+koF/iPFb9m26atHDv3GFwzMqbZsOQrx1Pf8l5ZuvRrquqprhr1aQngtN7zTfJTbyTcHwv4KuKJCF6xAs0aWzfuYGioSrlU4LjjZoxdV6sixgTU+2IKU8/FJXeIyBtV1fNTELe9lXzXNTT64hREh2SzRI2Ib33rDpb9dBXbtu8jimJy2Qxz587khv96Ge96x5WIgMaJEWeVZs1SKH5DGzumwOpbRiM7YLW5bR7if5skdrRyaeNl+fY//4K+gWGmTqmMRWvPKAOhzw2X7aPclV674rx+vnPXTKZWYhIrKBAEPr19Q3zne7/gc7e+H+ca4HmGZtWRyX9em9vrIvLdAxo50ElU+zSZ4p8T9iaIBOocEgSowj9959+443/fy5atu2k0IzKZgNnHTuf6Ny/lz959FZlcBm02EWMCwr6Y/NQ3aLjlHSLyr6K6JUco68nmT6IZqipGMgH79g1yw42f5cGHn6ZUypPNBogIzjmazZharcEbrljEv3z34xQLOUgSjKCIKLkOQ7P2NM7eCdKD2gX4wXvw/TKNUB2IyeXZ9OIWllzxl6hOjMqpRq74wnrOOr0KwJPPlXn9R+eTyzhUx/u89PsP3vNlTjhlDq7RwIikQTFbhEZtHaprES0j5mKypZmEww4R45wDP6BWa/DOP7mVX9y9hmIpRzYTYIxBVWk2Y6rVOksuWsDt//zXTJ/WiUYxIjgyOSFqbmVoYJ5ovedacvllNOrpw4FEhWuu/TgrH1zPjKO6SZJkLAikUdPg+4Zdu/t585sW88MffAbbbCLOYYwBVUu+4EG2Rc0E4irEsaoxYp3Dzxe49vpPsOL+J6hUiljrxrRxqObz+nMH+PHnniMJU4T9vHLtJ85gxdouKsUE69LrnmcYGqrxh687h2V3fo4krOOJQVK4lXzBQCb1MC6EZpSmpM6BMZhMlre+7W/56V2PcPTMbpLE4ZwbN1bB93327O3ndRcv5Kc/vgVftBWlnSNfMoS1GwzCtYhRQBNrMbkSd9z+S+57YB0zjuoiiuI2ENOo6YiihJkzuvn5L9fw4Zu+hpfLY7KZFHTwXJhGdhcOWcKhhCRRB2ITi1/o4stfvJ1frniczkppDMTxGvkXV+9sK0Ji4C+u2ckEmom1js5Kif9Y8Rhf/tId+IUuEmtxqoKIIaw7wqGEcDhxjUgdeEmSYLIZTC7PTR/+Gv/n7jXMnNFNFCVtII4yhCiKmXFUF/c98BQ/uONuTK5EYi1gRsW71rv5kx/6PB5dWCsiRlDLx/7mNnr7hvA9Hx1nRxMH4ZyjWMzx0KPPsuGFLVywaB4dU6YhnoK14pwaVI2mEygml8VkS9z2jz/ik3/3r3R2FLHjBPc9ZaDq88ZF/fzlf99OUhc8oxgBGwlz5zR4+qUST79cpJhzaSQHnCqFfJZ7Vz7JlI4s515wDuKBixOcIqrOqGKMETHZDCZbZs+eXt77vi9w+533M6W7fFBKOnGsqX4r/YNVbrh+KUZBUMGIYG1JtN7TwEjWWYfJBuzbuZ8LLvsgtXqI53ltQCZWxqLm+MPzPAYGRzj2mGm8+8Y3cvUbFzN3zswWcW9JpAmbN+3gq1+7k+//6F7K5cIoJRkTXBVUhVVfWcfJx9exDcGMcUfByysbNhdY+qGFGNGxe0bpECgjIyF//N9ezwfffx0nnDgLxD/AvpKILVt387O7Hua737uL7bt66eosj7GF8SDGieB72ka3rLUUi3l+de9XmX70tNRXGgNOY9GwR0dV2OSzvLyhhyVX3IS1FmNSTyOkRH9Gd8Lu/gDfOxSYpuWYQ6ZNrXDqqccx97iZlEt54jhh6/a9PPnkS/QPDNPZWUa1PWgEvrKnP8MX3/My77thB/GQ4Hk6wYyFoKJ8/fZZfOQ7xzOjOyJOpA0AEcPg4Ajd3R285qyTmXPsUWQCn+FqyJatu3lhQw+9vUNjAfRQbiWxwszumD39Pl7LvYikGPmexwP3fJm5Jx+LCyNMK90SrW+t43l5Zy0m8OnfP8AFf/ABBodq+H6qkcYoQ1Wfqxb3MWtaky8um8UxUyOSRFDaSbJnDHGSEIYRcZKgmpJu3zcUWvXFicIHvrJvIOAtF/fyv25+gaQheIepgVgV/Kzyxzefxk8emsr0rrgNzNFJTRJLvd4gSRyKIiIEgU8+lyHwfaxz7W4L8DxlV1+Gj1y/g+37svz8kSlUSgnOSSsBSOislHn0vq/RNa0TjRPE88Da0ADb8f3UyuKE7ulTOPXk2TQa0VimoSrks441L3Tw0bdu4/3X7GJXbwYRHUvbRs00aWlyqZSnu6vM1O4K3d0ddJSLGCOHBHH/YMCFZw7zzY+8hCZg0FcooCpq4VsffYnFZw6zfzAg8PWgAGSM0FEu0t3dkcrQVaZUzGOMkFjbBqJnUjexuy/DB/7LTj5y/TbWPN9BIesYTYyMCI1GzGmnHEvXUVNwcYKCEvggss0Aj2AyGHDWOvCyLFk8nyhOxtRWFbKBY3dfhoefqfClv97Ep9/ZQ63hUw09fE/HfNno951zWOtIrMVa29IA2oQ3JjXnS88a5M7PPkc5b3HxwY5+oum5GMp5y7K/e56lZw2ypz+DMRMnFaxzWGtbMqS0ZrwMxii+p1RDj1rD8Jkbe/jiJzfz0DOd7O7PkBnHWcUIUZywZPECMJm0GA0OyYDqIwZj7kLTkpYYARIuXjyPQj7bpj2jg1v5ZBeE8PEbe/jZrc+y4IQavUMBtdBrFUt0LNLKuNNICp7vpbM/VPOphT4ffMtOfnLLc3QVE5JmyqN/Y1HTQNKErlLMT295jg+8ZSe10Geo5v9WMtRCj96hgAUn1Pj5rc/xV+/sgRBWPtV5UNXbWkehkOXixfOAhBQrBE1A9ReiumMKYbyBTHaKa0ZqfCPNZsxFl32QLT27yWYzqZ8UCCPDybNCVn91HSJKUFTihnDnyul8/+4ZPLWpRL1h8IwS+DpmMql2CHEiWCeU8paL5g/zweu2c+E5Q7g6OHtkILbTLzAemAI8vLaTryybxcNPd1ANPTxPCbzDy1DMORaeVOXtl+/hukv3EeSUpCZYFS75wEI27syTzzicpr6/0Yw4Yc5MHlzxVbLZAJc4NdmMEEW9xP6pvsisPq33rMHLX2kkcja2XrZc4fxzT+X5DT3k8zmstTiFXMaxaWeO53sKLDy9SnMkpUM3vHEvN/zBXp54qcyD6zt5amOJbXuyDNb8Fo2AciFh9lFNzjllhNedPcCCU6pgIB4RjOhvDeKoZqpLn3Hhawa5cOEg6zeUuO+JLp7YUGbb3iwjdZ/Epr64s5gwe0aTs06qsmTBIK85eQTy4KrQHBGyJeXZ54ps3pUn1wJxNJNrNCLOP/c0suUKtjaMZ4zDy3sQr5HKrD5fFaGpK8C7MnUt6d1LL1rIv9yxfFyRPzWLasPnoac7WTi/ms40jFGVs88Y4ewFI2DTbKze8IhtSqoLOYef17QmH4NtpEFsvF/7f2qDCniSapOIsuCUKgvOrEICSSjUGwbrhMBTCjmLyaerXkjANcAOCaZl6vjw0DOdVBuGaVlLYg80OqSFybi2h4IHIitUESOCEpuVxMOKiGfEgItYdO5pTJvaSRzbMf+oCoGnPLi+AnHqrEdpA0BSF6IhIamlN5QKlq6OhI6ixTN64PNQUp/1O4LYrp0pGEnYekcrK+oopjKUCinpTmoHPqclu7TuJ4YH1lcIxvFkEYhjy7SpnSw691RwEUYMiHhEw4rRlSK02qulY18gjjeSzYkRnDYjjp49g/lnzCVsNBExrVQspUHrNpfYtzeDH3DIKDgKkLVgk/TvtEuhB0X4/+/95XHvUG2X4VAyjiqIH8DevRnWbyqRzx5IP0UMYaPJvDOOZ+axM9BmhBEc2Zxgk5fIzH5RVcUAaQNI5H4kB+No0MWL5xFF7TQo4zv29gc8tqEMWcZeeDizGz1ftcVNRyCDU4EsPPZimb0DARn/AO0xRoiihCUXzgNvjMm4FCu5v9U88wysbt0h94x29g7QoPlpM2oCDbJOWL2u80AyP8mP0TGsXteFddIGurWOQj7LxYvnt9MeEjDc07oTA5ekSp/1HiEcGSbwPUGUqMm8M49n7nEzaTTjsSzHOSGfcTz6XAdRK2pP9iPwlWhEWPN8R2rWTsZS3mYzYu6cGcw7Yy5ETQRRgsCjUR2i6f0qfcIl1oiIpm3FWX3AGvw8RnA2tuQ6Oll0zqmEjWZasG21SHNZx6adeV7sKSIZxl48GQ/nBMnAC1uLbNyZO4j2hI2IReecRrajExsnqX/086CskcqsvlZLVk17q1SWt9YMjNGgSy5acNDLPZOmVQ89U4HgQBN/UgKpQAAPP1NJifwhAuEoBm20x7B8PHajQLpWQnkf0UgbDXrteaczdUqFOEnaaJDvKQ+u70xpkExeJI2M0p7OQ9CehGlTKrz2vNMm0J4RJU7uH4/d+HUvkNn0PDZpo0HHHDeTeafPJQwPpkFPbSqxf1+An5mcQUc1rT3v3xew7rC0Zy5Hz57ZTnuSeCNrtz2vmi40aFvWp6q+yNIEY1YeKQ3a0xfw+IYOyLwyDfr9NWuBDDz2Ygd7+g9Ney5ePP9g2uN596dLc9Q7eH3k6hYNsiyfSIOWLJ5PPpdpawy10SAzeTUSA6vXdR5Ee5xz5PMZlhyK9qDL2zBrA/KSFg2yjUdojIyM0aA4Yt6845lz3IyDaFCuRYPiSUqDAl+JR4RHn+9Io/U42tNoxMydPYN5Zx4PUdSiPb5HozpM5D/Shtl4IMdoUMfJ+1H59RgNihLylU7OP/tUwnACDco4Nu7Is6GnMOlo0CjteXFrgY07DpTMDtCeJuefcyq5SmUC7dFfS8cxvaO053BLnw9Dg+SQNMj3lJHQ46FnOkk8IUqExE6OM0qExBMefqZzrH55MO1ZmNZuU+LTKl2x4lDYTVyCd4AGxVVFGKNBF5x3OlO6O4jjhLRklPoYr1UN+rN37MQX/T3YS3ZkR0YVsmm1x59Ie5KEqd2HoD3xiOL0vjasDgXkARp0zLOEWzeTK5xoGqHTZmRmzUlp0CNrnqVUKqAu9SnFrGPthjK3/MNx+IFD3atbpDjSICMG4tjwxEtlCm1poaER1jl74UnMmtNGewyNcCPFOc+2YXW4NeSqq3wRSbTeswrJngChs9YZP5/jogvOZNWD6zBGcC5l+sYotYbhcz+Ync7qZHGTrcVXlWKSltxopz0XXTAPvBzWNvE945CsQGOViCQpRu0rkw9eXTsa0dXcA+7d42nQpRcv5H9+5c62VuaoOUytxJMyszlQBT/QUs5mAy69aOEE2uMkrZCNw2g8BgdrpIqIqI5smo4fbMbzSiSJOkHU+Lzh6o/yq8eep7NSOuwWjslbBfIYGKqy+PwzuOvf/x5xSboTyPcFm1TJZY4XOXr/KEYHR+l2P6m6bJkn5RP3oTxIUFAQi1M83+fWT7+bfDZDvd4kCHyMpKsQZELrc/Kc6TqoIPCp1Zvkcxlu+fS70g2lTgGxBAVF5YEUxGXeobbCyKGd8eg2t22XkyvdTWMkAfzRbW4r713Dn7zvS+za3UuhmCPwvUnkHA92lnFiqdcaHH30VG77+k1cetmiA9vzICFX9mlUL5f87OWH22fzSju/Wts5tiwn3/V66gPjNl4W2dWzm2/e9jNWPvAk+/YPHrSucNJUf4xh+rROXrfkNfz5n17N0bNnTtgw2uUTDi6XwpzLX2k32SsCCSjhltl4uSfwzRQaocXzPJdYTC6TRrXaMPt7B9OkXoDJkim2ZPU8w7SpnenmAtvANSJMCqIll/ewrpdEzyZ/9HY4UO05YiAh/a0Jue46q8ObLyRfuBvfLxHWYsBXENe2OX3ymvbopnpjDK1WfUK+GJDYEeLqFVI48ZHftLfxN45+zF8OvnguxcoP8MsnkQxCPO7nEtI8anIeE3/mIfA9/E5Iqi8RVd8mxRPXHske8CNSozEw+9dWKB/9cVT+B0FmavsPeCiT0rYR8Ly0sY0HcdSL6vfYsetWOeGcoSPdSH/E9jhq5gBa3zwbP38dcDnOnoFqN6reATAniUkj6aZ5kX7EPIfI3SRumRRmbR8fcI/kaf8XRS+cabHVo64AAAAASUVORK5CYII=';

/* ════════════ LÍNEAS DE MANDO Y COMUNICACIÓN ════════════
 * El administrador decide quién es el jefe directo de cada persona (Configuración →
 * Líneas de mando y comunicación). El jefe directo aprueba sus permisos y ve lo suyo;
 * el jefe del jefe ya no recibe esos avisos. Cada quien le puede preguntar a su jefe,
 * a su equipo y a quien el administrador agregue en «también se comunica con».
 * Sin configurar, queda como antes: gerentes, bodega y producción → director operativo;
 * contador → director financiero; directores → administrador. */
var H_MANDO = ['Usuario', 'Jefe directo', 'También se comunica con', 'Actualizado por', 'Actualizado en'];
var _MANDO = null;
function hojaMando(ss) {
  ss = ss || libro();
  var h = ss.getSheetByName('Líneas de mando');
  if (!h) { h = hojaLimpia(ss, 'Líneas de mando', H_MANDO); h.getRange('E:E').setNumberFormat('dd/mm/yyyy hh:mm'); h.setColumnWidth(3, 320); }
  return h;
}
function mandoCfg() {
  if (_MANDO) return _MANDO;
  _MANDO = {};
  leeTodo(hojaMando(), H_MANDO.length).forEach(function (r) {
    var n = String(r[0] || '').trim(); if (!n) return;
    _MANDO[n] = { jefe: String(r[1] || '').trim(), com: String(r[2] || '').split(' · ').map(function (s) { return s.trim(); }).filter(Boolean) };
  });
  return _MANDO;
}
function primeroDeRol(rol) { var n = ''; usuariosCache().forEach(function (u) { if (!n && u.activo && u.rol === rol) n = u.nombre; }); return n; }
function jefePorDefecto(u) {
  if (!u || u.rol === 'admin' || u.rol === 'dueno') return '';
  if (['gerente', 'bodega', 'produccion'].indexOf(u.rol) >= 0) return primeroDeRol('operaciones') || primeroDeRol('admin');
  if (esContadorU(u)) return primeroDeRol('finanzas') || primeroDeRol('admin');
  return primeroDeRol('admin');
}
function jefeDe(nombre) {
  var u = usuarioPorNombre(nombre); if (!u) return '';
  var c = mandoCfg()[u.nombre];
  if (c && c.jefe) { var j = usuarioPorNombre(c.jefe); if (j && j.activo && j.nombre !== u.nombre) return j.nombre; }
  if (c && c.jefe === '' && u.rol === 'admin') return '';
  return jefePorDefecto(u);
}
function equipoDirecto(nombre) {
  return usuariosCache().filter(function (u) { return u.activo && u.nombre !== nombre && jefeDe(u.nombre) === nombre; }).map(function (u) { return u.nombre; });
}
function equipoTotal(nombre) {
  var out = [], cola = [nombre], vistos = {}; vistos[nombre] = true;
  while (cola.length) { equipoDirecto(cola.shift()).forEach(function (n) { if (!vistos[n]) { vistos[n] = true; out.push(n); cola.push(n); } }); }
  return out;
}
function seComunican(a, b) {
  var ca = mandoCfg()[a], cb = mandoCfg()[b];
  return jefeDe(a) === b || jefeDe(b) === a || !!(ca && ca.com.indexOf(b) >= 0) || !!(cb && cb.com.indexOf(a) >= 0);
}
function pantallaMando(cred) {
  var yo = exigeAdmin(cred);
  var us = usuariosCache().filter(function (u) { return u.activo; });
  return { usuarios: us.map(function (u) {
    var c = mandoCfg()[u.nombre] || { jefe: '', com: [] };
    return { nombre: u.nombre, puesto: puestoDe(u), rol: u.rol, jefe: jefeDe(u.nombre), jefeDefecto: jefePorDefecto(u),
      configurado: !!c.jefe, com: c.com.filter(function (n) { return us.some(function (x) { return x.nombre === n; }); }),
      equipo: equipoDirecto(u.nombre), aprueba: equipoDirecto(u.nombre).length > 0 };
  }), yo: yo.nombre };
}
function guardarMando(cred, lista) {
  var yo = exigeAdmin(cred);
  var us = {}; usuariosCache().forEach(function (u) { if (u.activo) us[u.nombre] = u; });
  var nuevo = {};
  (lista || []).forEach(function (x) {
    var n = String(x.nombre || ''); if (!us[n]) return;
    var j = String(x.jefe || '');
    if (j && !us[j]) throw new Error('No existe el usuario «' + j + '».');
    if (j === n) throw new Error(n + ' no puede ser su propio jefe.');
    if (!j && us[n].rol !== 'admin' && us[n].rol !== 'dueno') throw new Error('Escoja el jefe directo de ' + n + '.');
    nuevo[n] = { jefe: j, com: (x.com || []).filter(function (k) { return us[k] && k !== n; }) };
  });
  // sin vueltas: nadie puede quedar como jefe de su propio jefe
  Object.keys(nuevo).forEach(function (n) {
    var p = n, pasos = 0;
    while (nuevo[p] && nuevo[p].jefe && pasos < 50) { p = nuevo[p].jefe; pasos++; if (p === n) throw new Error('La línea de mando de ' + n + ' da vuelta y regresa a él. Revise los jefes.'); }
  });
  var lock = LockService.getScriptLock(); lock.waitLock(15000); _LEE = {};
  try {
    var h = hojaMando(), ahora = new Date();
    if (h.getLastRow() > 1) h.getRange(2, 1, h.getLastRow() - 1, H_MANDO.length).clearContent();
    var filas = Object.keys(nuevo).map(function (n) { return [n, nuevo[n].jefe, nuevo[n].com.join(' · '), yo.nombre, ahora]; });
    if (filas.length) h.getRange(2, 1, filas.length, H_MANDO.length).setValues(filas);
    _MANDO = null;
  } finally { lock.releaseLock(); }
  var r = pantallaMando(cred); r.ok = true; r.mensaje = 'Líneas de mando guardadas. Los permisos y las preguntas siguen estas líneas desde ahora.';
  return r;
}

/* ════════════ PERMISOS ════════════
 * Cualquiera pide permiso desde Configuración; lo aprueba su jefe directo (según las
 * líneas de mando). Quien lo pidió lo puede cambiar o cancelar mientras no haya pasado;
 * si ya estaba aprobado y lo cambia, vuelve a su jefe para aprobarlo. */
var H_PERM = ['No.', 'Pedido en', 'Pedido por', 'Rol', 'Tipo', 'Desde', 'Hasta', 'Horas', 'Motivo', 'Quién cubre',
  'Estado', 'Revisado por', 'Revisado en', 'Comentario', 'Cambiado en'];
var PERM_TIPOS = ['Día libre', 'Vacaciones', 'Médico / IGSS', 'Asunto personal', 'Llegar tarde', 'Salir temprano'];
var PERM_HORA = { 'Llegar tarde': 'Llega a las ', 'Salir temprano': 'Sale a las ' };
function hojaPerm(ss) {
  ss = ss || libro();
  var h = ss.getSheetByName('Permisos');
  if (!h) { h = hojaLimpia(ss, 'Permisos', H_PERM); h.getRange('B:B').setNumberFormat('dd/mm/yyyy hh:mm'); h.getRange('F:G').setNumberFormat('@');
    h.getRange('M:M').setNumberFormat('dd/mm/yyyy hh:mm'); h.getRange('O:O').setNumberFormat('dd/mm/yyyy hh:mm'); h.setColumnWidth(9, 300); }
  else if (h.getLastColumn() < H_PERM.length) encabezaAlFinal(h, 15, ['Cambiado en']);
  return h;
}
function filasPerm() {
  return leeTodo(hojaPerm(), H_PERM.length).map(function (r, i) {
    return { fila: i + 2, numero: String(r[0] || ''), en: fmtSello(r[1]), por: String(r[2] || ''), rol: String(r[3] || ''),
      tipo: String(r[4] || ''), desde: fmtDia(r[5]), hasta: fmtDia(r[6]), horas: String(r[7] || ''), motivo: String(r[8] || ''),
      cubre: String(r[9] || ''), estado: String(r[10] || ''), revPor: String(r[11] || ''), revEn: fmtSello(r[12]),
      comentario: String(r[13] || ''), cambEn: fmtSello(r[14]) };
  }).filter(function (x) { return x.numero; });
}
function puedeAprobarPerm(yo, x) {
  return x.por !== yo.nombre && jefeDe(x.por) === yo.nombre;
}
function quienApruebaTxt(nombre) {
  var j = jefeDe(nombre), u = usuarioPorNombre(j);
  return j ? j + (u ? ' (' + puestoDe(u).toLowerCase() + ')' : '') : 'nadie';
}
function puedeCambiarPerm(yo, x, hoy) {
  return x.por === yo.nombre && (x.estado === 'Pedido' || (x.estado === 'Aprobado' && x.desde > hoy));
}
function listaPermisos(cred) {
  permiteCmo();
  var yo = quien(cred), hoy = hoyISO(), todos = filasPerm(), equipo = equipoDirecto(yo.nombre);
  var aprueba = equipo.length > 0;
  var mios = todos.filter(function (x) { return x.por === yo.nombre; }).reverse().slice(0, 20).map(function (x) {
    x.puedeCambiar = puedeCambiarPerm(yo, x, hoy); x.puedeCancelar = x.por === yo.nombre && (x.estado === 'Pedido' || (x.estado === 'Aprobado' && x.hasta >= hoy)); return x; });
  var porAprobar = todos.filter(function (x) { return x.estado === 'Pedido' && puedeAprobarPerm(yo, x); });
  var fuera = aprueba ? todos.filter(function (x) { return x.estado === 'Aprobado' && equipo.indexOf(x.por) >= 0 && x.hasta >= hoy && x.desde <= masDias(hoy, 13); })
    .sort(function (a, b) { return a.desde < b.desde ? -1 : 1; }) : [];
  return { tipos: PERM_TIPOS, conHora: PERM_HORA, mios: mios, porAprobar: porAprobar, fuera: fuera, aprueba: aprueba, hoy: hoy, equipo: equipo,
    puedePedir: !!jefeDe(yo.nombre) && yo.rol !== 'dueno', quienAprueba: quienApruebaTxt(yo.nombre),
    recientes: aprueba ? todos.filter(function (x) { return x.estado !== 'Pedido' && x.revPor === yo.nombre && x.por !== yo.nombre; }).reverse().slice(0, 15) : [] };
}
/** Revisa lo que se escribió en el formulario del permiso. */
function datosPerm(d) {
  d = d || {};
  if (PERM_TIPOS.indexOf(d.tipo) < 0) throw new Error('Escoja el tipo de permiso.');
  if (!esFecha(d.desde)) throw new Error('Escoja desde qué día.');
  var conHora = !!PERM_HORA[d.tipo];
  var hasta = conHora ? d.desde : esFecha(d.hasta) ? d.hasta : d.desde;
  if (hasta < d.desde) throw new Error('La fecha final no puede ser antes de la inicial.');
  if (diasEntre(d.desde, hasta) > 30) throw new Error('Un permiso de más de 30 días hay que hablarlo en persona.');
  var motivo = String(d.motivo || '').trim().slice(0, 300);
  if (motivo.length < 4) throw new Error('Escriba el motivo.');
  var horas = '';
  if (conHora) {
    var hh = String(d.hora || '').trim();
    if (!/^\d{1,2}:\d{2}$/.test(hh)) throw new Error(d.tipo === 'Llegar tarde' ? 'Escriba a qué hora va a llegar.' : 'Escriba a qué hora se va a ir.');
    horas = PERM_HORA[d.tipo] + ('0' + hh).slice(-5);
  }
  return { tipo: d.tipo, desde: d.desde, hasta: hasta, horas: horas, motivo: motivo, cubre: String(d.cubre || '').trim().slice(0, 60) };
}
function pedirPermiso(cred, d) {
  permiteCmo();
  var yo = quien(cred);
  if (yo.rol === 'dueno' || !jefeDe(yo.nombre)) throw new Error('Su usuario no pide permisos en la app.');
  var p = datosPerm(d);
  var lock = LockService.getScriptLock();
  lock.waitLock(15000); _LEE = {};
  var numero;
  try {
    var max = 0;
    filasPerm().forEach(function (x) { var m = /^PE-(\d+)$/.exec(x.numero); if (m) max = Math.max(max, +m[1]); });
    numero = 'PE-' + ('0000' + (max + 1)).slice(-4);
    var h = hojaPerm();
    h.appendRow([numero, new Date(), yo.nombre, yo.rol, p.tipo, p.desde, p.hasta, p.horas, p.motivo, p.cubre, 'Pedido', '', '', '', '']);
    h.getRange(h.getLastRow(), 6, 1, 2).setNumberFormat('@').setValues([[p.desde, p.hasta]]);
  } finally { lock.releaseLock(); }
  var r = listaPermisos(cred); r.ok = true; r.mensaje = numero + ' enviado. Lo revisa ' + r.quienAprueba + '.';
  return r;
}
function buscaPerm(numero) {
  var x = null; filasPerm().forEach(function (y) { if (y.numero === String(numero)) x = y; });
  if (!x) throw new Error('No se encontró ese permiso.');
  return x;
}
/** Quien lo pidió lo cambia. Si ya estaba aprobado, vuelve a su jefe. */
function cambiarPermiso(cred, numero, d) {
  permiteCmo();
  var yo = quien(cred), hoy = hoyISO();
  var p = datosPerm(d), x, antes;
  var lock = LockService.getScriptLock(); lock.waitLock(15000); _LEE = {};
  try {
    x = buscaPerm(numero); antes = x.estado;
    if (x.por !== yo.nombre) throw new Error('Solo quien pidió el permiso lo puede cambiar.');
    if (!puedeCambiarPerm(yo, x, hoy)) throw new Error(x.estado === 'Aprobado' ? 'Ese permiso ya empezó: ya no se puede cambiar. Hable con su jefe.' : 'Ese permiso ya fue ' + x.estado.toLowerCase() + '.');
    var h = hojaPerm();
    h.getRange(x.fila, 5, 1, 11).setValues([[p.tipo, p.desde, p.hasta, p.horas, p.motivo, p.cubre, 'Pedido', '', '', '', new Date()]]);
    h.getRange(x.fila, 6, 1, 2).setNumberFormat('@').setValues([[p.desde, p.hasta]]);
  } finally { lock.releaseLock(); }
  var r = listaPermisos(cred); r.ok = true;
  r.mensaje = numero + ' cambiado. ' + (antes === 'Aprobado' ? 'Como ya estaba aprobado, vuelve a ' : 'Lo revisa ') + r.quienAprueba + '.';
  return r;
}
function cancelarPermiso(cred, numero) {
  permiteCmo();
  var yo = quien(cred), hoy = hoyISO(), x;
  var lock = LockService.getScriptLock(); lock.waitLock(15000); _LEE = {};
  try {
    x = buscaPerm(numero);
    if (x.por !== yo.nombre) throw new Error('Solo quien pidió el permiso lo puede cancelar.');
    if (!(x.estado === 'Pedido' || (x.estado === 'Aprobado' && x.hasta >= hoy))) throw new Error('Ese permiso ya no se puede cancelar.');
    hojaPerm().getRange(x.fila, 11, 1, 5).setValues([['Cancelado', x.revPor, x.revEn ? x.revEn : '', x.comentario, new Date()]]);
  } finally { lock.releaseLock(); }
  var r = listaPermisos(cred); r.ok = true; r.mensaje = numero + ' cancelado.' + (x.estado === 'Aprobado' ? ' ' + jefeDe(yo.nombre) + ' lo verá en su campanita.' : '');
  return r;
}
function revisarPermiso(cred, numero, aprobar, comentario) {
  var yo = quien(cred);
  var x = buscaPerm(numero);
  if (!puedeAprobarPerm(yo, x)) throw new Error(x.por === yo.nombre ? 'Nadie aprueba su propio permiso.' :
    'Ese permiso lo aprueba ' + quienApruebaTxt(x.por) + ', el jefe directo de ' + x.por + '.');
  if (x.estado !== 'Pedido') throw new Error('Ese permiso ya fue ' + x.estado.toLowerCase() + '.');
  comentario = String(comentario || '').trim().slice(0, 200);
  if (!aprobar && comentario.length < 4) throw new Error('Escriba por qué no se aprueba.');
  hojaPerm().getRange(x.fila, 11, 1, 4).setValues([[aprobar ? 'Aprobado' : 'Rechazado', yo.nombre, new Date(), comentario]]);
  var r = listaPermisos(cred); r.ok = true; r.mensaje = numero + (aprobar ? ' aprobado.' : ' no aprobado.') + ' ' + x.por + ' lo verá en su campanita.';
  return r;
}

/* ── avisos de todo lo nuevo ── */
function avisosNuevos(yo, out, cred) {
  var hoy = hoyISO(), desde = masDias(hoy, -10);
  // permisos
  try {
    filasPerm().forEach(function (x) {
      if (x.estado === 'Pedido' && puedeAprobarPerm(yo, x))
        out.push({ id: 'perm-' + x.numero + (x.cambEn ? '-' + x.cambEn : ''), tipo: x.cambEn ? 'Permiso cambiado por aprobar' : 'Permiso por aprobar', cuando: x.en, titulo: x.por + ': ' + x.tipo,
          texto: dia(x.desde).slice(0, 5) + (x.hasta !== x.desde ? ' al ' + dia(x.hasta).slice(0, 5) : '') + ' · ' + x.motivo, abrir: { perm: true } });
      if (x.estado === 'Cancelado' && x.revPor && puedeAprobarPerm(yo, x) && x.cambEn.slice(0, 10) >= desde)
        out.push({ id: 'perm-c-' + x.numero, tipo: 'Permiso cancelado', cuando: x.cambEn, titulo: x.por + ': ' + x.tipo,
          texto: dia(x.desde).slice(0, 5) + ' · lo canceló quien lo pidió', abrir: { perm: true } });
      if (x.por === yo.nombre && x.estado !== 'Pedido' && x.estado !== 'Cancelado' && x.revEn.slice(0, 10) >= desde)
        out.push({ id: 'perm-r-' + x.numero, tipo: 'Permiso ' + x.estado.toLowerCase(), cuando: x.revEn,
          titulo: x.tipo + ' · ' + dia(x.desde).slice(0, 5) + ' ' + x.estado.toLowerCase(), texto: (x.comentario ? x.comentario + ' · ' : '') + x.revPor,
          abrir: { perm: true } });
    });
  } catch (e) {}
  // solicitudes de dinero
  try {
    filasFondos().forEach(function (x) {
      var cuando = x.pagEn || x.aprEn || x.autEn || x.verEn || x.en, urg = x.urgencia === 'Urgente' ? ' · URGENTE' : '';
      if (x.estado === FONDO_EST.VER && yo.rol === 'operaciones')
        out.push({ id: 'sd-ver-' + x.numero, tipo: 'Solicitud por verificar', cuando: x.en, titulo: x.numero + ' · ' + x.por + ' · ' + (x.reqDinero ? dinero(x.monto) : 'sin dinero'),
          texto: x.unidad + ' · ' + x.que.slice(0, 80) + urg, abrir: { pag: 'fondos' } });
      if (x.estado === FONDO_ADM && yo.esAdmin)
        out.push({ id: 'sd-aut-' + x.numero, tipo: 'Solicitud por autorizar', cuando: x.verEn || x.en, titulo: x.numero + ' · ' + x.por + ' · ' + (x.reqDinero ? dinero(x.monto) : 'sin dinero'),
          texto: x.unidad + ' · ' + x.que.slice(0, 80) + urg, abrir: { pag: 'fondos' } });
      if (x.estado === FONDO_EST.PED && yo.rol === 'finanzas')
        out.push({ id: 'sd-apr-' + x.numero, tipo: 'Solicitud por aprobar', cuando: x.autEn || x.en, titulo: x.numero + ' · ' + x.por + ' · ' + dinero(x.monto),
          texto: x.unidad + ' · ' + x.que.slice(0, 80) + urg, abrir: { pag: 'conta', ct: 'fondos' } });
      if (x.estado === FONDO_EST.APR && esContador(yo))
        out.push({ id: 'sd-pag-' + x.numero, tipo: 'Solicitud por pagar', cuando: x.aprEn || x.autEn, titulo: x.numero + ' · ' + dinero(x.monto) + ' · ' + x.unidad,
          texto: 'Aprobó ' + (x.aprPor || x.autPor || x.verPor) + '. ' + x.que.slice(0, 60), abrir: { pag: 'conta', ct: 'fondos' } });
      if (x.estado === FONDO_DEV && x.por === yo.nombre)
        out.push({ id: 'sd-dev-' + x.numero + '-' + x.devEn, tipo: 'Solicitud devuelta', cuando: x.devEn, titulo: x.numero + ' devuelta por ' + (x.devPor || 'contabilidad'), texto: x.devMotivo, abrir: { pag: 'fondos' } });
      if (x.por === yo.nombre) {
        if (x.estado === FONDO_AUT)
          out.push({ id: 'sd-ok-' + x.numero, tipo: 'Solicitud autorizada', cuando: x.autEn, titulo: x.numero + ' autorizada: ya puede proceder', texto: 'Autorizó ' + x.autPor + (x.reqDinero ? '' : ' · no necesitaba dinero') + '. Márquela como hecha cuando termine.', abrir: { pag: 'fondos' } });
        else if ([FONDO_EST.RECH, FONDO_EST.PAG, FONDO_EST.APR].indexOf(x.estado) >= 0 && cuando.slice(0, 10) >= desde)
          out.push({ id: 'sd-g-' + x.estado + '-' + x.numero, tipo: 'Solicitud ' + x.estado.toLowerCase(), cuando: cuando, titulo: x.numero + ' ' + x.estado.toLowerCase(),
            texto: x.estado === FONDO_EST.RECH ? ((x.aprPor || x.verPor) + ': ' + (x.comentario || x.verCom)) : x.estado === FONDO_EST.PAG ? 'Contabilidad entregó ' + dinero(x.pagado) + '.' : 'Aprobó ' + (x.aprPor || x.autPor) + '. Falta que contabilidad entregue el dinero.', abrir: { pag: 'fondos' } });
      }
    });
  } catch (e) {}
  // presupuestos de las áreas: contabilidad aprueba, la CMO arma y recibe la respuesta
  try {
    var actM = hoy.slice(0, 7), unC = yo.cmo ? unidadPorId(yo.sucursal) : null;
    filasPAreaEst().forEach(function (e) {
      if (e.estado === PA_EST.POR && (yo.rol === 'finanzas' || yo.esAdmin))
        out.push({ id: 'pa-apr-' + e.lugar + '-' + e.mes + '-' + e.en, tipo: 'Presupuesto de área por aprobar', cuando: e.en,
          titulo: e.lugar + ' · ' + periodoTxt(e.mes), texto: e.por + ' lo mandó para que lo apruebe.', abrir: { parea: true } });
      if (unC && e.lugar === unC.nombre && (e.estado === PA_EST.APR || e.estado === PA_EST.DEV) && e.resEn.slice(0, 10) >= desde)
        out.push({ id: 'pa-res-' + e.mes + '-' + e.estado + '-' + e.resEn, tipo: 'Presupuesto ' + e.estado.toLowerCase(), cuando: e.resEn,
          titulo: periodoTxt(e.mes) + (e.estado === PA_EST.APR ? ' aprobado' : ' devuelto'),
          texto: e.estado === PA_EST.APR ? 'Aprobó ' + e.resPor + '. Ya cuenta para su seguimiento.' : e.resPor + ': ' + e.comentario, abrir: { cfg: 'miPres' } });
    });
    if (unC) [actM, mesMas(actM, 1)].forEach(function (m) {
      var e = estadoPArea(m, unC.nombre), desdeAv = m === actM ? m + '-01' : mesMas(m, -1) + '-' + PA_DIA_AVISO;
      if (hoy >= desdeAv && (!e || e.estado === PA_EST.BOR || e.estado === PA_EST.DEV))
        out.push({ id: 'pa-rec-' + m + (hoy > limitePArea(m) ? '-tarde' : ''), tipo: hoy > limitePArea(m) ? 'Presupuesto atrasado' : 'Presupuesto por armar', cuando: desdeAv + ' 08:00',
          titulo: 'Presupuesto de ' + periodoTxt(m), texto: 'Mándelo a contabilidad a más tardar el ' + dia(limitePArea(m)).slice(0, 5) + '.', abrir: { cfg: 'miPres' } });
    });
  } catch (e) {}
  // emergencias
  try {
    var mia = unidadEmerg(yo);
    if (mia) emergAgrupadas().forEach(function (e) {
      if (e.da === mia && e.estado === EM_EST.PED)
        out.push({ id: 'em-' + e.numero, tipo: 'Pedido de emergencia', cuando: e.en, titulo: e.pide + ' le pide ' + e.lineas.length +
          (e.lineas.length === 1 ? ' producto' : ' productos') + ' (' + e.numero + ')', texto: e.lineas.slice(0, 3).map(function (l) { return num(l.pedido) + ' ' + l.nombre; }).join(', '),
          abrir: { tab: 'emerg' } });
      if (e.pide === mia && e.estado === EM_EST.ENV)
        out.push({ id: 'em-env-' + e.numero, tipo: 'Emergencia en camino', cuando: e.envEn, titulo: e.da + ' mandó ' + e.numero,
          texto: 'Recíbalo cuando llegue.', abrir: { tab: 'emerg' } });
      if (e.pide === mia && e.estado === EM_EST.RECH && e.envEn.slice(0, 10) >= desde)
        out.push({ id: 'em-rech-' + e.numero, tipo: 'Emergencia rechazada', cuando: e.envEn, titulo: e.da + ' no pudo mandar ' + e.numero,
          texto: e.nota, abrir: { tab: 'emerg' } });
    });
  } catch (e) {}
  // horario, descanso, mantenimiento y sugerencias
  try {
    var todosD = filasDesc();
    todosD.forEach(function (x) {
      if (x.estado === DESC_EST.PED && (yo.rol === 'operaciones'))
        out.push({ id: 'de-' + x.numero, tipo: 'Descanso por decidir', cuando: x.en, titulo: x.gerente + ' pide descansar los ' + x.dia.toLowerCase(),
          texto: nombreUnidad(x.sucursal), abrir: { hor: true } });
      if (yo.rol === 'gerente' && x.gerente === yo.nombre && x.estado !== DESC_EST.PED && x.resEn.slice(0, 10) >= desde)
        out.push({ id: 'de-r-' + x.numero + '-' + x.estado, tipo: x.estado === DESC_EST.ASIG ? 'Su día de descanso' : 'Descanso ' + x.estado.toLowerCase(),
          cuando: x.resEn, titulo: x.estado === DESC_EST.RECH ? 'No se aprobó ' + x.dia.toLowerCase() : 'Descansa los ' + x.dia.toLowerCase(),
          texto: (x.comentario ? x.comentario + ' · ' : '') + x.resPor, abrir: { pag: 'kpi' } });
    });
    if (yo.rol === 'operaciones' || yo.rol === 'gerente') equiposVista(yo).forEach(function (x) {
      if (x.estado === 'al dia' || x.estado === 'sin fecha') return;
      out.push({ id: 'mt-' + x.numero + '-' + x.proximo, tipo: x.porMedicion ? 'Aceite por cambiar' : x.estado === 'vencido' ? 'Mantenimiento vencido' : 'Mantenimiento por vencer',
        cuando: x.proximo + ' 00:00', titulo: x.tarea + ' · ' + x.equipo + (yo.rol === 'operaciones' ? ' (' + x.sucursalNombre + ')' : ''),
        texto: (x.estado === 'vencido' ? 'Tocaba el ' + dia(x.proximo).slice(0, 5) + ' (hace ' + (-x.restan) + (x.restan === -1 ? ' día)' : ' días)') : x.restan === 0 ? 'Toca hoy' : 'Toca en ' + x.restan + (x.restan === 1 ? ' día' : ' días')) +
          (x.porMedicion && x.medicion ? ' · según la medición: ' + x.medicion.motivo : ''),
        abrir: { pag: 'mant' } });
    });
    if (yo.rol === 'operaciones') {
      var ayerM = masDias(hoy, -1), sinMed = {}, aceF = null;
      filasEq().forEach(function (x) {
        if (!x.activo || !x.medir || (x.alta && x.alta > ayerM)) return;
        var hay = (aceF = aceF || filasAceite()).some(function (l) { return l.equipoNo === x.numero && l.fecha === ayerM; });
        if (!hay) (sinMed[x.sucursal] = sinMed[x.sucursal] || []).push(x.equipo);
      });
      Object.keys(sinMed).forEach(function (s) {
        out.push({ id: 'ace-falta-' + ayerM + '-' + s, tipo: 'Aceite sin medir', cuando: '', titulo: nombreUnidad(s) + ' no midió el aceite ayer',
          texto: sinMed[s].join(', '), abrir: { pag: 'mant' } });
      });
    }
    try { avisosMerma(yo, out, desde); } catch (e) {}
    try { avisosFirmas(yo, out, desde); } catch (e) {}
    try { avisosCalendario(yo, out, desde); } catch (e) {}
    try { avisosProductos(yo, out, desde); } catch (e) {}
    if (leeSug(yo)) filasSug().forEach(function (x) {
      if (!x.leidaPor) out.push({ id: 'sg-' + x.numero, tipo: 'Sugerencia nueva', cuando: x.en, titulo: x.anonima ? 'De alguien (anónima)' : 'De ' + x.de,
        texto: x.texto.slice(0, 120), abrir: { sug: true } });
    });
  } catch (e) {}
  // preguntas
  try {
    filasConsultas().forEach(function (x) {
      if (x.para === yo.nombre && x.estado === 'Abierta')
        out.push({ id: 'cq-' + x.numero, tipo: 'Pregunta de ' + x.de, cuando: x.en, titulo: x.tema, texto: x.pregunta, abrir: { cq: true } });
      if (x.de === yo.nombre && x.estado === 'Respondida' && x.resEn.slice(0, 10) >= desde)
        out.push({ id: 'cqr-' + x.numero, tipo: 'Respuesta de ' + x.para, cuando: x.resEn, titulo: x.tema, texto: x.respuesta, abrir: { cq: true } });
    });
  } catch (e) {}
  // presupuesto por aprobar / devuelto / aprobado
  try {
    var mesP = hoy.slice(0, 7), mesS = masDias(finDeMes(mesP), 1).slice(0, 7);
    [mesP, mesS].forEach(function (m) { var ep = estadoPres(m); if (!ep) return;
      if (ep.estado === PRES_EST.POR && yo.rol === 'finanzas')
        out.push({ id: 'pres-por-' + m + '-' + ep.en, tipo: 'Presupuesto por aprobar', cuando: ep.en, titulo: periodoTxt(m) + ' · lo mandó ' + ep.por,
          texto: 'Revíselo en Configuración → Presupuesto mensual.', abrir: { pres: m } });
      if (ep.estado !== PRES_EST.POR && ep.por === yo.nombre && ep.aprPor && ep.aprPor !== yo.nombre && ep.aprEn.slice(0, 10) >= desde)
        out.push({ id: 'pres-' + ep.estado + '-' + m + '-' + ep.aprEn, tipo: 'Presupuesto ' + ep.estado.toLowerCase(), cuando: ep.aprEn,
          titulo: periodoTxt(m) + ' · ' + ep.aprPor, texto: ep.comentario || (ep.estado === PRES_EST.APR ? 'Ya quedó aprobado.' : ''), abrir: { pres: m } });
    });
  } catch (e) {}
  // presupuesto
  if (yo.rol === 'finanzas' || esContador(yo)) {
    try { avisosPresupuestoDatos(out, presupuestoMes(cred, hoy.slice(0, 7))); } catch (e) {}
  }
}

/* ════════════ ARCHIVOS (cotizaciones en PDF o foto) ════════════
 * Se guardan en la carpeta «Rey Pizza · Cotizaciones» del Drive del dueño de la app
 * y se comparten con enlace (solo ver), para que se abran desde el teléfono. */
var CARPETA_COT = 'Rey Pizza · Cotizaciones';
var TIPOS_ARCHIVO = { 'application/pdf': 'pdf', 'image/jpeg': 'jpg', 'image/png': 'png' };
function guardaArchivo(a, prefijo, etiqueta) {
  a = a || {};
  var que = etiqueta || 'La cotización';
  var tipo = String(a.tipo || '').toLowerCase(), nombre = String(a.nombre || 'cotizacion').replace(/[\\/:*?"<>|]/g, ' ').slice(0, 80);
  if (!TIPOS_ARCHIVO[tipo] && /\.pdf$/i.test(nombre)) tipo = 'application/pdf';
  if (!TIPOS_ARCHIVO[tipo]) throw new Error(que + ' tiene que ser PDF o una foto (JPG o PNG).');
  var datos = String(a.datos || '').replace(/^data:[^,]*,/, '');
  if (!datos) throw new Error('Escoja el archivo: ' + que.toLowerCase() + '.');
  if (datos.length * 0.75 > 8 * 1024 * 1024) throw new Error('El archivo pesa más de 8 MB. Mande una versión más liviana.');
  var it = DriveApp.getFoldersByName(CARPETA_COT);
  var carpeta = it.hasNext() ? it.next() : DriveApp.createFolder(CARPETA_COT);
  var blob = Utilities.newBlob(Utilities.base64Decode(datos), tipo, (prefijo ? prefijo + ' · ' : '') + nombre);
  var f = carpeta.createFile(blob);
  try { f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) {}
  return { url: f.getUrl(), nombre: nombre };
}

/* ════════════ CORREOS ════════════
 * Cada quien guarda su correo en Configuración. Un disparador cada 5 minutos le
 * manda por correo los avisos nuevos de su campanita (lo que ya se le mandó queda
 * anotado en «Correos enviados» para no repetir). Las preguntas del jefe y el
 * resumen del presupuesto salen al momento. */
var H_CORREOS = ['Usuario', 'Aviso', 'Enviado en'];
function hojaCorreos(ss) {
  ss = ss || libro();
  var h = ss.getSheetByName('Correos enviados');
  if (!h) { h = hojaLimpia(ss, 'Correos enviados', H_CORREOS); h.getRange('C:C').setNumberFormat('dd/mm/yyyy hh:mm'); }
  return h;
}
function correoValido(c) { return /^[^\s@]+@[^\s@]+\.[a-z]{2,}$/i.test(String(c || '')); }
/* La dirección del botón «Abrir la app» de los correos. No se pregunta a Google: cuando el
 * correo sale del disparador automático, Google devuelve la dirección de prueba (/dev), que
 * solo abre para quien edita el proyecto («No se pudo abrir el archivo»). Cuando la app
 * instalable esté lista, aquí se pone https://app.caixsa.com */
var URL_APP = 'https://script.google.com/macros/s/AKfycbyjF_gjAOBS6H-86bONREc26kGuHmG9UIYQD8QoCN0bahyxWWPcuQ-CQY9hkQdxHmPovQ/exec';
function urlApp() { return URL_APP; }
function cuotaCorreo() { try { return MailApp.getRemainingDailyQuota(); } catch (e) { anotaErrorCorreo(e); return 0; } }
/** Lo último que pasó con los correos, para poder explicar por qué no llegan. */
function anotaErrorCorreo(e) {
  var m = String(e && e.message || e || '');
  try { CacheService.getScriptCache().put('correo_error', JSON.stringify({ m: m, en: fmtSello(new Date()) }), 21600); } catch (x) {}
}
function anotaOkCorreo() { try { CacheService.getScriptCache().put('correo_ok', fmtSello(new Date()), 21600); } catch (x) {} }
function explicaErrorCorreo(m) {
  m = String(m || '');
  if (!m) return '';
  if (/permis|autoriz|authoriz|scope|send_mail/i.test(m))
    return 'Google todavía no le dio permiso a la app para mandar correos. Hay que autorizarlo una vez desde el editor: función autorizarCorreos (ver los pasos que le mandé).';
  if (/cuota|quota|limit/i.test(m)) return 'Se acabó la cuota de correos de hoy de la cuenta de Google. Mañana se reanuda sola.';
  if (/invalid email|dirección|address/i.test(m)) return 'Google dice que la dirección de correo no es válida.';
  return 'Google no dejó mandar el correo: ' + m;
}
function htmlCorreo(titulo, cuerpo) {
  var u = urlApp();
  return '<div style="font-family:Arial,Helvetica,sans-serif;max-width:560px;margin:auto;color:#1b1d22">' +
    '<div style="background:#14284B;color:#fff;padding:14px 18px;border-radius:10px 10px 0 0;border-bottom:4px solid #F2B705"><b style="font-size:17px;letter-spacing:1px">REY PIZZA</b> ' +
    '<span style="color:#F2B705;font-size:11px;letter-spacing:2px;font-weight:bold">CAIX, S.A.</span>' +
    '<div style="font-size:13px;color:#F7F2E7;margin-top:2px">' + esHtml(titulo) + '</div></div>' +
    '<div style="background:#FFFFFF;border:1px solid #E3DBCB;border-top:0;padding:16px 18px;border-radius:0 0 10px 10px;color:#16223A">' + cuerpo +
    (u ? '<p style="margin-top:18px"><a href="' + u + '" style="background:#F2B705;color:#14284B;padding:10px 18px;border-radius:8px;text-decoration:none;font-weight:bold">Abrir la app</a></p>' : '') +
    '<p style="font-size:11px;color:#6c7079;margin-top:16px">Le llegó porque guardó este correo en Configuración de la app. Para dejar de recibirlos, bórrelo ahí.</p></div></div>';
}
function esHtml(t) { return String(t == null ? '' : t).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;'); }
function listaAvisosHtml(avisos) {
  return avisos.map(function (a) {
    return '<div style="padding:10px 0;border-top:1px solid #eef0f3"><div style="font-size:11px;color:#6c7079;text-transform:uppercase;letter-spacing:.05em">' +
      esHtml(a.tipo) + '</div><div style="font-weight:bold;margin:2px 0">' + esHtml(a.titulo) + '</div>' +
      (a.texto ? '<div style="font-size:13px;color:#3d4250">' + esHtml(a.texto) + '</div>' : '') + '</div>'; }).join('');
}
function enviaCorreo(para, asunto, html) {
  if (!correoValido(para)) return false;
  var q; try { q = MailApp.getRemainingDailyQuota(); } catch (e) { anotaErrorCorreo(e); return false; }
  if (q < 1) { anotaErrorCorreo('Se acabó la cuota (quota) de correos de hoy.'); return false; }
  try { MailApp.sendEmail({ to: para, subject: asunto, htmlBody: html, name: 'Rey Pizza · app' }); }
  catch (e) { anotaErrorCorreo(e); return false; }
  anotaOkCorreo();
  return true;
}
function ultimoErrorCorreo() {
  try { var c = CacheService.getScriptCache().get('correo_error'); return c ? JSON.parse(c) : null; } catch (e) { return null; }
}
function enviadosPor() {
  var s = {};
  leeTodo(hojaCorreos(), 3).forEach(function (r) { s[String(r[0]) + '|' + String(r[1])] = true; });
  return s;
}
function anotaEnviados(nombre, ids) {
  if (!ids.length) return;
  var h = hojaCorreos(), ahora = new Date(), f0 = h.getLastRow() + 1;
  h.getRange(f0, 1, ids.length, 3).setValues(ids.map(function (id) { return [nombre, id, ahora]; }));
}
/** Un aviso suelto, al momento (preguntas, respuestas). Queda anotado para que el
 *  disparador no lo repita. */
function correoAviso(nombre, aviso) {
  var u = null; usuariosCache().forEach(function (x) { if (x.nombre === nombre && x.activo) u = x; });
  if (!u || !correoValido(u.correo)) return false;
  var ok = false;
  try { ok = enviaCorreo(u.correo, 'Rey Pizza · ' + aviso.tipo + ': ' + aviso.titulo, htmlCorreo(aviso.tipo, listaAvisosHtml([aviso]))); } catch (e) {}
  if (ok) anotaEnviados(nombre, [aviso.id]);
  return ok;
}
/** Lo corre el disparador cada 5 minutos: a cada quien con correo, sus avisos nuevos. */
function enviarCorreosPendientes() {
  // Este trabajo tarda (calcula los avisos de cada persona y manda correos). Antes tomaba el candado de
  // todo el sistema y cualquiera que guardara algo en esos minutos se quedaba esperando o fallaba.
  // Ahora usa su propia marca: así nunca corren dos a la vez, y nunca estorba a quien guarda.
  var marca = null;
  try { marca = CacheService.getScriptCache(); if (marca.get('correos_en_curso')) return 'ocupado'; marca.put('correos_en_curso', '1', 280); } catch (e) { marca = null; }
  var enviados = 0;
  try {
    var ya = enviadosPor();
    usuariosCache().forEach(function (u) {
      if (!u.activo || !correoValido(u.correo)) return;
      var av = [];
      try { av = notificaciones({ usuario: u.nombre, pin: u.pin }).avisos; } catch (e) { return; }
      var nuevos = av.filter(function (a) { return a.id && !ya[u.nombre + '|' + a.id]; });
      if (!nuevos.length) return;
      var asunto = nuevos.length === 1 ? 'Rey Pizza · ' + nuevos[0].tipo + ': ' + nuevos[0].titulo : 'Rey Pizza · ' + nuevos.length + ' avisos nuevos';
      var ok = false;
      try { ok = enviaCorreo(u.correo, asunto, htmlCorreo(nuevos.length === 1 ? nuevos[0].tipo : 'Tiene ' + nuevos.length + ' avisos nuevos', listaAvisosHtml(nuevos))); } catch (e) {}
      if (ok) { anotaEnviados(u.nombre, nuevos.map(function (a) { return a.id; })); enviados++; }
    });
    limpiaCorreos();
  } finally { try { if (marca) marca.remove('correos_en_curso'); } catch (e) {} }
  return enviados;
}
/** Para que la hoja no crezca sin fin: se borra lo de más de 60 días. */
function limpiaCorreos() {
  var h = hojaCorreos(); if (h.getLastRow() < 4000) return;
  var lim = masDias(hoyISO(), -60), v = leeTodo(h, 3), quedan = v.filter(function (r) { return fmtSello(r[2]).slice(0, 10) >= lim; });
  h.getRange(2, 1, v.length, 3).setValues(v.map(function (r, i) { return quedan[i] || ['', '', '']; }));
}
function instalaDisparadorCorreos(lanzar) {
  try {
    var hay = ScriptApp.getProjectTriggers().some(function (t) { return t.getHandlerFunction() === 'enviarCorreosPendientes'; });
    if (!hay) ScriptApp.newTrigger('enviarCorreosPendientes').timeBased().everyMinutes(5).create();
    var hayD = ScriptApp.getProjectTriggers().some(function (t) { return t.getHandlerFunction() === 'reporteDiarioCorreo'; });
    if (!hayD) ScriptApp.newTrigger('reporteDiarioCorreo').timeBased().everyDays(1).atHour(21).nearMinute(15).create();
  } catch (e) { if (lanzar) throw e; }
}
function miCorreo(cred) {
  permiteAux();
  var yo = quien(cred);
  return { correo: yo.correo || '', cuota: cuotaCorreo() };
}
/** Guarda (o borra, si viene vacío) el correo de quien entra. Lo que ya tiene en la
 *  campanita se le manda en un solo correo de bienvenida, y desde ahí solo lo nuevo. */
function guardarMiCorreo(cred, correo) {
  permiteAux();
  var yo = quien(cred);
  correo = String(correo || '').trim().toLowerCase();
  if (correo && !correoValido(correo)) throw new Error('Ese correo no se ve bien. Ejemplo: nombre@gmail.com');
  var u = null; filasUsuarios().forEach(function (x) { if (x.nombre === yo.nombre) u = x; });
  if (!u) throw new Error('No se encontró su usuario.');
  var h = hojaUsuarios();
  if (h.getMaxColumns() < 9) encabezaAlFinal(h, 9, ['Correo']);
  h.getRange(u.fila, 9).setValue(correo);
  olvidaUsuarios();
  if (!correo) return { ok: true, correo: '', mensaje: 'Correo borrado. Ya no le llegan avisos por correo.' };
  var av = [];
  try { av = notificaciones(cred).avisos; } catch (e) {}
  var ok = false;
  try { ok = enviaCorreo(correo, 'Rey Pizza · Correo guardado', htmlCorreo('Desde hoy le llegan aquí los avisos de la app',
    '<p>Hola ' + esHtml(yo.nombre) + ': cada solicitud, aprobación o pregunta que le toque le llegará a este correo.</p>' +
    (av.length ? '<p><b>Lo que tiene pendiente ahora:</b></p>' + listaAvisosHtml(av.slice(0, 25)) : '<p>Por ahora no tiene nada pendiente.</p>'))); } catch (e) {}
  var ya = enviadosPor();
  anotaEnviados(yo.nombre, av.filter(function (a) { return a.id && !ya[yo.nombre + '|' + a.id]; }).map(function (a) { return a.id; }));
  var er = ok ? null : ultimoErrorCorreo();
  return { ok: true, correo: correo, enviado: ok, error: er ? explicaErrorCorreo(er.m) : '',
    mensaje: ok ? 'Correo guardado. Le mandamos un correo de prueba a ' + correo + ' (revise también Spam).' :
    'Correo guardado, pero el de prueba NO salió. ' + (er ? explicaErrorCorreo(er.m) : '') };
}

/* ════════════ PRESUPUESTO: APROBACIÓN ════════════
 * El contador lo arma y lo manda a aprobar; el director financiero lo aprueba (o lo
 * devuelve). Si lo arma el director financiero, queda aprobado al guardarlo. Al
 * aprobarse, al administrador le llega el resumen por correo. */
var H_PRESE = ['Mes', 'Estado', 'Enviado por', 'Enviado en', 'Aprobado por', 'Aprobado en', 'Comentario'];
var PRES_EST = { POR: 'Por aprobar', APR: 'Aprobado', DEV: 'Devuelto' };
function hojaPresEstado(ss) {
  ss = ss || libro();
  var h = ss.getSheetByName('Presupuesto aprobación');
  if (!h) { h = hojaLimpia(ss, 'Presupuesto aprobación', H_PRESE); h.getRange('A:A').setNumberFormat('@');
    h.getRange('D:D').setNumberFormat('dd/mm/yyyy hh:mm'); h.getRange('F:F').setNumberFormat('dd/mm/yyyy hh:mm'); }
  return h;
}
function estadoPres(mes) {
  var x = null;
  leeTodo(hojaPresEstado(), H_PRESE.length).forEach(function (r, i) {
    var m = String(r[0] || ''); if (/^\d{4}-\d{2}/.test(m) && m.slice(0, 7) === mes)
      x = { fila: i + 2, mes: mes, estado: String(r[1] || ''), por: String(r[2] || ''), en: fmtSello(r[3]),
        aprPor: String(r[4] || ''), aprEn: fmtSello(r[5]), comentario: String(r[6] || '') }; });
  return x;
}
function ponEstadoPres(mes, fila) {
  var h = hojaPresEstado(), x = estadoPres(mes);
  if (x) h.getRange(x.fila, 1, 1, H_PRESE.length).setValues([fila]);
  else { var f0 = h.getLastRow() + 1; h.getRange(f0, 1).setNumberFormat('@'); h.getRange(f0, 1, 1, H_PRESE.length).setValues([fila]); }
}
function aprobarPresupuesto(cred, mes, aprobar, comentario) {
  var yo = quien(cred);
  if (yo.rol !== 'finanzas') throw new Error('El presupuesto lo aprueba el director financiero.');
  var x = estadoPres(mes);
  if (!x || x.estado !== PRES_EST.POR) throw new Error('Ese presupuesto no está esperando aprobación.');
  comentario = String(comentario || '').trim().slice(0, 300);
  if (!aprobar && comentario.length < 4) throw new Error('Escriba qué hay que corregir.');
  ponEstadoPres(mes, [mes, aprobar ? PRES_EST.APR : PRES_EST.DEV, x.por, aFechaHora(x.en) || new Date(), yo.nombre, new Date(), comentario]);
  var r = presupuestoMes(cred, mes); r.ok = true;
  if (aprobar) { var env = correoPresupuesto(mes, yo.nombre);
    r.mensaje = 'Presupuesto de ' + periodoTxt(mes) + ' aprobado.' + (env ? ' Le mandamos el resumen a ' + env + '.' : ' (El administrador no tiene correo guardado: el resumen no salió por correo.)'); }
  else r.mensaje = 'Presupuesto devuelto a ' + x.por + ' para corregir.';
  return r;
}
/** El resumen del presupuesto aprobado, por correo a los administradores. */
function correoPresupuesto(mes, aprobo) {
  var admins = usuariosCache().filter(function (u) { return u.activo && u.rol === 'admin' && correoValido(u.correo); });
  if (!admins.length) return '';
  var p = presupuestoMes(null, mes, true), t = p.total;
  var fila = function (a, b, neg) { return '<tr><td style="padding:5px 8px;border-top:1px solid #eef0f3">' + a + '</td><td style="padding:5px 8px;border-top:1px solid #eef0f3;text-align:right' +
    (neg ? ';color:#b3261e' : '') + '">' + b + '</td></tr>'; };
  var h = '<p>' + esHtml(aprobo) + ' aprobó el presupuesto de <b>' + esHtml(periodoTxt(mes)) + '</b>.</p>' +
    '<table style="width:100%;border-collapse:collapse;font-size:14px">' + fila('Venta presupuestada', dinero(t.venta)) +
    fila('(−) Costo de ventas', dinero(t.costo)) + fila('<b>= Utilidad bruta</b>', '<b>' + dinero(t.utilidadBruta) + '</b>') +
    fila('(−) Gastos de operación', dinero(t.operacion)) + fila('<b>= Utilidad neta</b>', '<b>' + dinero(t.utilidadNeta) + '</b> (' +
      (t.venta ? Math.round(t.utilidadNeta / t.venta * 1000) / 10 : 0) + ' %)', t.utilidadNeta < t.metaNeta) +
    fila('Utilidad neta esperada', t.metaNetaPct + ' % · ' + dinero(t.metaNeta)) + '</table>';
  h += '<p style="margin-top:14px"><b>Por sucursal</b></p><table style="width:100%;border-collapse:collapse;font-size:13px"><tr style="color:#6c7079">' +
    '<td style="padding:4px 8px"></td><td style="padding:4px 8px;text-align:right">Venta</td><td style="padding:4px 8px;text-align:right">Gasto</td><td style="padding:4px 8px;text-align:right">Utilidad</td></tr>' +
    p.unidades.map(function (u) { return '<tr><td style="padding:5px 8px;border-top:1px solid #eef0f3">' + esHtml(u.nombre) + '</td><td style="padding:5px 8px;border-top:1px solid #eef0f3;text-align:right">' +
      dinero(u.venta) + '</td><td style="padding:5px 8px;border-top:1px solid #eef0f3;text-align:right">' + dinero(u.costo + u.operacion) +
      '</td><td style="padding:5px 8px;border-top:1px solid #eef0f3;text-align:right">' + dinero(u.utilidadNeta) + '</td></tr>'; }).join('') + '</table>';
  var tot = {};
  p.unidades.forEach(function (u) { u.rubros.forEach(function (r) { if (r.presupuesto) tot[r.rubro] = (tot[r.rubro] || 0) + r.presupuesto; }); });
  var ks = Object.keys(tot).sort(function (a, b) { return tot[b] - tot[a]; });
  if (ks.length) h += '<p style="margin-top:14px"><b>Rubros (las 4 unidades)</b></p><table style="width:100%;border-collapse:collapse;font-size:13px">' +
    ks.map(function (k) { return fila(esHtml(k), dinero(tot[k])); }).join('') + '</table>';
  var a = [];
  admins.forEach(function (u) { try { if (enviaCorreo(u.correo, 'Rey Pizza · Presupuesto de ' + periodoTxt(mes) + ' aprobado', htmlCorreo('Presupuesto para revisar', h))) a.push(u.correo); } catch (e) {} });
  return a.join(', ');
}

/* ════════════ PREGUNTAS DEL JEFE ════════════
 * El administrador (o un director a su equipo) pregunta por algo atrasado; le
 * aparece a la persona en su tablero y en la campanita (y por correo), y también a
 * su director. La respuesta le vuelve a quien preguntó. */
var H_CONS = ['No.', 'Preguntada en', 'De', 'Para', 'Tema', 'Pregunta', 'Estado', 'Respuesta', 'Respondida en'];
function hojaConsultas(ss) {
  ss = ss || libro();
  var h = ss.getSheetByName('Preguntas');
  if (!h) { h = hojaLimpia(ss, 'Preguntas', H_CONS); h.getRange('B:B').setNumberFormat('dd/mm/yyyy hh:mm');
    h.getRange('I:I').setNumberFormat('dd/mm/yyyy hh:mm'); h.setColumnWidth(6, 320); h.setColumnWidth(8, 320); }
  return h;
}
function filasConsultas() {
  return leeTodo(hojaConsultas(), H_CONS.length).map(function (r, i) {
    return { fila: i + 2, numero: String(r[0] || ''), en: fmtSello(r[1]), de: String(r[2] || ''), para: String(r[3] || ''),
      tema: String(r[4] || ''), pregunta: String(r[5] || ''), estado: String(r[6] || ''), respuesta: String(r[7] || ''),
      resEn: fmtSello(r[8]) };
  }).filter(function (x) { return x.numero; });
}
function usuarioPorNombre(n) { var u = null; usuariosCache().forEach(function (x) { if (x.nombre === n) u = x; }); return u; }
function esContadorU(u) { return u && u.rol === 'registro' && u.confirma; }
function puedePreguntarA(yo, u) {
  if (!u || !u.activo || u.nombre === yo.nombre || u.rol === 'dueno') return false;
  if (yo.esAdmin) return true;
  if (yo.rol === 'finanzas' && (u.rol === 'operaciones' || u.rol === 'cmo' || u.rol === 'admin')) return true;   // supervisa presupuestos, pagos y planillas: pregunta a los directores y avisa al administrador
  return seComunican(yo.nombre, u.nombre);
}
function listaConsultas(cred) {
  permiteCmo();
  var yo = quien(cred), mio = {};
  var ls = filasConsultas().reverse(), desde = masDias(hoyISO(), -30);
  var ve = function (x) {
    if (x.de === yo.nombre || x.para === yo.nombre || yo.esAdmin) return true;
    return equipoTotal(yo.nombre).indexOf(x.para) >= 0;   // el jefe ve las preguntas a su equipo
  };
  var vis = ls.filter(function (x) { return ve(x) && (x.estado === 'Abierta' || x.en.slice(0, 10) >= desde); });
  vis.forEach(function (x) { x.puedeResponder = x.para === yo.nombre && x.estado === 'Abierta'; x.mia = x.de === yo.nombre; });
  return { consultas: vis.slice(0, 40), abiertasParaMi: vis.filter(function (x) { return x.puedeResponder; }).length,
    aQuien: usuariosCache().filter(function (u) { return puedePreguntarA(yo, u); }).map(function (u) {
      return { nombre: u.nombre, puesto: puestoDe(u) }; }) };
}
function puestoDe(u) {
  if (!u) return '';
  if (u.rol === 'operaciones') return 'Director operativo';
  if (u.rol === 'finanzas') return 'Director financiero';
  if (esContadorU(u)) return 'Contador interno';
  if (u.rol === 'gerente') return 'Gerente ' + nombreUnidad(u.sucursal);
  if (u.rol === 'bodega') return 'Bodega';
  if (u.rol === 'cmo') return 'CMO · Chief Marketing Officer';
  if (u.rol === 'produccion') return 'Producción';
  if (u.rol === 'admin') return 'Administración';
  return 'Registra';
}
function preguntar(cred, d) {
  permiteCmo();
  var yo = quien(cred); d = d || {};
  var u = usuarioPorNombre(String(d.para || ''));
  if (!puedePreguntarA(yo, u)) throw new Error('A esa persona no le puede preguntar desde aquí.');
  var preg = String(d.pregunta || '').trim().slice(0, 500);
  if (preg.length < 5) throw new Error('Escriba la pregunta.');
  var tema = String(d.tema || '').trim().slice(0, 120) || 'Pregunta';
  var lock = LockService.getScriptLock(); lock.waitLock(15000); _LEE = {};
  var numero;
  try {
    var max = 0; filasConsultas().forEach(function (x) { var m = /^PR-(\d+)$/.exec(x.numero); if (m) max = Math.max(max, +m[1]); });
    numero = 'PR-' + ('0000' + (max + 1)).slice(-4);
    hojaConsultas().appendRow([numero, new Date(), yo.nombre, u.nombre, tema, preg, 'Abierta', '', '']);
  } finally { lock.releaseLock(); }
  var mando = correoAviso(u.nombre, { id: 'cq-' + numero, tipo: 'Pregunta de ' + yo.nombre, titulo: tema, texto: preg });
  var r = listaConsultas(cred); r.ok = true;
  r.mensaje = 'Pregunta enviada a ' + u.nombre + (mando ? ' (también por correo).' : '. La verá en su tablero y en la campanita.');
  return r;
}
function responderConsulta(cred, numero, respuesta) {
  permiteCmo();
  var yo = quien(cred), x = null;
  filasConsultas().forEach(function (y) { if (y.numero === String(numero)) x = y; });
  if (!x) throw new Error('No se encontró la pregunta ' + numero + '.');
  if (x.para !== yo.nombre) throw new Error('Esa pregunta es para ' + x.para + '.');
  if (x.estado !== 'Abierta') throw new Error('Esa pregunta ya se respondió.');
  respuesta = String(respuesta || '').trim().slice(0, 800);
  if (respuesta.length < 3) throw new Error('Escriba la respuesta.');
  hojaConsultas().getRange(x.fila, 7, 1, 3).setValues([['Respondida', respuesta, new Date()]]);
  correoAviso(x.de, { id: 'cqr-' + x.numero, tipo: 'Respuesta de ' + yo.nombre, titulo: x.tema, texto: respuesta });
  var r = listaConsultas(cred); r.ok = true; r.mensaje = 'Respuesta enviada a ' + x.de + '.';
  return r;
}

/* ════════════ ATRASOS ════════════
 * Lo que alguien tenía que haber hecho y no ha hecho. Cada atraso dice de quién es
 * y de qué equipo (director operativo o financiero), para que lo vean el
 * administrador, el director de ese equipo y la persona. */
function nombresDe(fn) { return usuariosCache().filter(function (u) { return u.activo && fn(u); }).map(function (u) { return u.nombre; }); }
function datosAtrasos() {
  var hoy = hoyISO(), ayer = masDias(hoy, -1), ahoraS = fmtSello(new Date()), out = [];
  var ope = nombresDe(function (u) { return u.rol === 'operaciones'; }), fin = nombresDe(function (u) { return u.rol === 'finanzas'; });
  var con = nombresDe(esContadorU), bod = nombresDe(function (u) { return u.rol === 'bodega'; });
  var ger = function (sid) { return nombresDe(function (u) { return u.rol === 'gerente' && u.sucursal === sid; }); };
  var pon = function (quien, equipo, id, que, detalle, horas, abrir) {
    out.push({ id: id, quien: quien.length ? quien.join(', ') : '(sin asignar)', equipo: equipo, que: que, detalle: detalle,
      horas: horas == null ? null : Math.round(horas), abrir: abrir || null }); };
  // cierres de caja atorados en un paso
  var pasos = { validar: [], recibir: [], banco: [] };
  filasPorConfirmar().forEach(function (x) {
    if (x.tipo !== 'ingreso') return;
    var p = pasoDe(x); if (!p) return;
    var h = horasEntre(p === 'validar' ? x.sello : p === 'recibir' ? x.validadoEn : x.recibidoEn, ahoraS);
    if (h != null) pasos[p].push({ x: x, h: h });
  });
  var lim = { validar: 3, recibir: 24, banco: 48 };
  var txt = { validar: 'por validar', recibir: 'con el efectivo sin recibir', banco: 'sin confirmar en el banco' };
  Object.keys(pasos).forEach(function (p) {
    var l = pasos[p].filter(function (y) { return y.h >= lim[p]; }).sort(function (a, b) { return b.h - a.h; });
    if (!l.length) return;
    var v = l[0].x;
    pon(p === 'validar' ? ope : con, p === 'validar' ? 'operaciones' : 'finanzas', 'cierre-' + p,
      (l.length === 1 ? 'Un cierre ' : l.length + ' cierres ') + txt[p],
      'El más viejo: ' + v.sucursal + ' ' + v.turno + ' del ' + dia(v.fecha).slice(0, 5) + ', ' + Math.round(l[0].h) + ' h esperando.', l[0].h, { pc: true });
  });
  // turnos que no se reportaron
  try {
    var ci = cierresCaja(3);
    SUC_KPI.forEach(function (sid) { var s = nombreUnidad(sid), faltan = [];
      ci.fechas.slice().reverse().forEach(function (f) { if (f >= hoy) return;
        ['AM', 'PM'].forEach(function (tu) { if (!ci.celdas[s + '|' + f + '|' + tu]) faltan.push({ f: f, tu: tu }); }); });
      if (!faltan.length) return;
      pon(ger(sid), 'operaciones', 'turno-' + sid + '-' + faltan[faltan.length - 1].f + faltan[faltan.length - 1].tu,
        s + ': ' + (faltan.length === 1 ? 'un corte sin reportar' : faltan.length + ' cortes sin reportar'),
        faltan.map(function (x) { return x.tu + ' ' + dia(x.f).slice(0, 5); }).join(', ') + '. El gerente los envía en Registrar.',
        horasEntre(faltan[0].f + ' 21:00', ahoraS));
    });
  } catch (e) {}
  // indicadores de ayer sin anotar (gerentes y contador)
  try {
    datosResultadosKpi(ayer, SUC_KPI.concat([AREA_CONTA])).sucursales.forEach(function (s) {
      var falta = s.esperados - s.hechos; if (!s.esperados || falta <= 0) return;
      var esC = s.id === AREA_CONTA;
      pon(esC ? con : ger(s.id), esC ? 'finanzas' : 'operaciones', 'kpi-' + s.id + '-' + ayer,
        (esC ? 'Contabilidad' : s.nombre) + ': ' + falta + ' de ' + s.esperados + (esC ? ' tareas' : ' indicadores') + ' sin anotar ayer',
        'Lo hecho: ' + Math.round(s.hechos / s.esperados * 100) + ' %.', 24, esC ? { pag: 'tareas' } : { pag: 'res' });
    });
  } catch (e) {}
  // impuestos
  try { datosImpuestos().periodos.forEach(function (x) { if (x.estado === 'Vencido')
    pon(fin, 'finanzas', 'imp-' + x.obl + '-' + x.periodo, x.nombre + ' (' + x.periodoTxt + ') vencido sin pago anotado', 'Venció el ' + dia(x.vence) + '.', horasEntre(x.vence + ' 23:59', ahoraS), { pag: 'imp' }); }); } catch (e) {}
  // requisiciones
  try { filasFondos().forEach(function (x) {
    if (x.estado === FONDO_EST.PED) { var h = horasEntre(x.en, ahoraS);
      if (h >= (x.urgencia === 'Urgente' ? 4 : 24)) pon(fin, 'finanzas', 'rq-apr-' + x.numero, x.numero + ' sin aprobar' + (x.urgencia === 'Urgente' ? ' (urgente)' : ''),
        x.tipo + ' · ' + x.unidad + ' · ' + dinero(x.monto), h, { pag: 'conta', ct: 'fondos' }); }
    if (x.estado === FONDO_EST.APR) { var h2 = horasEntre(x.aprEn, ahoraS);
      if (h2 >= 48) pon(con, 'finanzas', 'rq-pag-' + x.numero, x.numero + ' aprobada y sin pagar', x.tipo + ' · ' + x.unidad + ' · ' + dinero(x.monto), h2, { pag: 'conta', ct: 'fondos' }); }
    if (x.estado === FONDO_EST.PAG && !esEquipoBodega({ rol: rolDe(x.por) })) { var h3 = horasEntre(x.pagEn, ahoraS);
      if (h3 >= 24 * 7) pon(ope, 'operaciones', 'rq-fin-' + x.numero, x.numero + ' pagada hace ' + Math.round(h3 / 24) + ' días y sin marcar hecha', x.que.slice(0, 80), h3, { pag: 'fondos' }); }
  }); } catch (e) {}
  // presupuesto
  try {
    var mes = hoy.slice(0, 7), ep = estadoPres(mes);
    if (!ep && Number(hoy.slice(8)) >= 3) pon(con.concat(fin), 'finanzas', 'pres-falta-' + mes, 'El presupuesto de ' + periodoTxt(mes) + ' no se ha armado',
      'Lo arma el contador y lo aprueba el director financiero.', null, { pres: true });
    if (ep && ep.estado === PRES_EST.POR) { var hp = horasEntre(ep.en, ahoraS);
      if (hp >= 24) pon(fin, 'finanzas', 'pres-apr-' + mes, 'Presupuesto de ' + periodoTxt(mes) + ' sin aprobar', 'Lo mandó ' + ep.por + '.', hp, { pres: true }); }
    if (ep && ep.estado === PRES_EST.DEV) pon(con, 'finanzas', 'pres-dev-' + mes, 'Presupuesto de ' + periodoTxt(mes) + ' devuelto sin corregir', ep.comentario, horasEntre(ep.aprEn, ahoraS), { pres: true });
  } catch (e) {}
  // órdenes de compra cotizadas y sin pagar
  try { ordenesAgrupadas().forEach(function (o) {
    if (o.estado === 'Anulada' || o.pago !== 'Cotizada') return;
    var h = horasEntre(o.cotizadaEn, ahoraS);
    if (h >= 72) pon(con, 'finanzas', 'oc-pag-' + o.numero, o.numero + ' de ' + o.proveedor + ' cotizada y sin pagar', dinero(o.totalCotizado), h, { pag: 'conta', ct: 'pagos' }); }); } catch (e) {}
  // bodega: solicitudes sin aceptar y diferencias
  try { filasSol().forEach(function (s) { if (s.estado !== SOL.ENV) return; var h = horasEntre(s.enviadaEn, ahoraS);
    if (h >= 24) pon(bod, 'operaciones', 'sol-' + s.numero, 'Solicitud ' + s.numero + ' de ' + s.sucursal + ' sin aceptar', 'Entrega ' + dia(s.entrega).slice(0, 5) + '.', h, { tab: 'sol' }); }); } catch (e) {}
  try { var dif = lineasSol().filter(function (l) { return l.difEstado === DIF.PEND; });
    if (dif.length) pon(bod, 'operaciones', 'dif', dif.length + (dif.length === 1 ? ' diferencia' : ' diferencias') + ' de entrega por aclarar', 'Bodega tiene que decir qué pasó.', null, { tab: 'dif' }); } catch (e) {}
  // permisos que empiezan mañana o antes y nadie ha respondido
  try { filasPerm().forEach(function (x) { if (x.estado !== 'Pedido' || x.desde > masDias(hoy, 1)) return;
    var ju = usuarioPorNombre(jefeDe(x.por)); if (!ju || ju.rol !== 'operaciones') return;
    pon(ope, 'operaciones', 'perm-' + x.numero, 'Permiso de ' + x.por + ' sin responder', x.tipo + ' desde el ' + dia(x.desde).slice(0, 5) + '.', horasEntre(x.en, ahoraS), { perm: true }); }); } catch (e) {}
  // el presupuesto de la CMO sin mandar a contabilidad
  try {
    var actM = hoy.slice(0, 7);
    usuariosCache().forEach(function (u) {
      if (!u.activo || u.rol !== 'cmo') return;
      var un = unidadPorId(u.sucursal); if (!un) return;
      [actM, mesMas(actM, 1)].forEach(function (m) {
        var e = estadoPArea(m, un.nombre), lim = limitePArea(m);
        if ((!e || e.estado === PA_EST.BOR || e.estado === PA_EST.DEV) && hoy > lim)
          pon([u.nombre], 'cmo', 'pa-' + un.id + '-' + m, 'Presupuesto de ' + periodoTxt(m) + ' sin enviar',
            'Debía mandarlo a contabilidad a más tardar el ' + dia(lim).slice(0, 5) + (e && e.estado === PA_EST.DEV ? ' (lo devolvieron: ' + e.comentario + ')' : '') + '.',
            horasEntre(lim + ' 21:00', ahoraS), { cfg: 'miPres' });
      });
    });
  } catch (e) {}
  out.sort(function (a, b) { return (b.horas || 0) - (a.horas || 0); });
  return out;
}
/** Qué atrasos ve cada quien: el administrador todos; cada director los de su
 *  equipo; los demás, solo los suyos. */
function atrasosPara(yo, todos) {
  todos = todos || datosAtrasos();
  if (yo.esAdmin) return todos;
  if (yo.rol === 'operaciones' || yo.rol === 'finanzas') return todos.filter(function (a) { return a.equipo === yo.rol; });
  return todos.filter(function (a) { return a.quien.split(', ').indexOf(yo.nombre) >= 0; });
}
/** Lo de arriba de cada tablero: atrasos, preguntas y cierres pendientes. */
function miPanel(cred) {
  var yo = quien(cred), out = { atrasos: [], consultas: null, cierres: null };
  try { out.atrasos = atrasosPara(yo); } catch (e) {}
  try { out.consultas = listaConsultas(cred); } catch (e) {}
  try { if (yo.valida || yo.recibe || yo.banco) { var pc = listaPorConfirmar(yo); out.cierres = { pendientes: pc.pendientes.length }; } } catch (e) {}
  return out;
}

/* ── seguimiento por persona (administrador) ── */
function personasSeguimiento(cred, seg) {
  var todos = datosAtrasos(), hoy = hoyISO(), mes = hoy.slice(0, 7), cons = filasConsultas();
  var abiertas = function (n) { return cons.filter(function (x) { return x.para === n && x.estado === 'Abierta'; }).length; };
  var out = [];
  var fs = []; try { fs = filasFondos(); } catch (e) {}
  usuariosCache().forEach(function (u) {
    if (!u.activo) return;
    var p = { nombre: u.nombre, puesto: puestoDe(u), rol: u.rol, lineas: [], atrasos: [], preguntas: abiertas(u.nombre) };
    if (u.rol === 'operaciones') {
      p.atrasos = todos.filter(function (a) { return a.equipo === 'operaciones'; });
      var c = seg.cierres, k = seg.indicadores;
      p.lineas.push(['Cierres por validar', c.porPaso.validar, c.porPaso.validar ? 'no' : 'si']);
      k.sucursales.forEach(function (s) { p.lineas.push(['Indicadores ' + s.nombre + ' (7 días)', s.pct == null ? '—' : s.pct + ' %', s.pct == null ? '' : s.pct >= 90 ? 'si' : 'no']); });
      if (seg.inventario) p.lineas.push(['Solicitudes de bodega sin aceptar', seg.inventario.enviadas, seg.inventario.enviadas ? 'no' : 'si']);
      p.lineas.push(['Solicitudes pagadas sin terminar', fs.filter(function (x) { return x.estado === FONDO_EST.PAG; }).length, '']);
      p.ver = 'tab';
    } else if (u.rol === 'finanzas') {
      p.atrasos = todos.filter(function (a) { return a.equipo === 'finanzas'; });
      var ep = null; try { ep = estadoPres(mes); } catch (e) {}
      p.lineas.push(['Presupuesto de ' + periodoTxt(mes), ep ? ep.estado : 'Sin armar', ep && ep.estado === PRES_EST.APR ? 'si' : 'no']);
      p.lineas.push(['Solicitudes por aprobar', fs.filter(function (x) { return x.estado === FONDO_EST.PED; }).length, fs.some(function (x) { return x.estado === FONDO_EST.PED; }) ? 'no' : 'si']);
      if (seg.impuestos) p.lineas.push(['Impuestos vencidos', seg.impuestos.resumen.vencidos, seg.impuestos.resumen.vencidos ? 'no' : 'si']);
      if (seg.contabilidad && seg.contabilidad.tareas) p.lineas.push(['Tareas del contador (7 días)', seg.contabilidad.tareas.pct + ' %', seg.contabilidad.tareas.pct >= 90 ? 'si' : 'no']);
      p.ver = 'fin';
    } else if (esContadorU(u)) {
      p.atrasos = todos.filter(function (a) { return a.quien.split(', ').indexOf(u.nombre) >= 0; });
      var c2 = seg.cierres;
      p.lineas.push(['Efectivo por recibir', c2.porPaso.recibir, c2.porPaso.recibir ? 'no' : 'si']);
      p.lineas.push(['Por confirmar en banco', c2.porPaso.banco, c2.porPaso.banco ? 'no' : 'si']);
      if (seg.contabilidad) { p.lineas.push(['Órdenes de compra por pagar', seg.contabilidad.compras.porPagar, seg.contabilidad.compras.porPagar ? 'no' : 'si']);
        if (seg.contabilidad.tareas) p.lineas.push(['Sus tareas (7 días)', seg.contabilidad.tareas.pct + ' %', seg.contabilidad.tareas.pct >= 90 ? 'si' : 'no']); }
      p.lineas.push(['Solicitudes por pagar', fs.filter(function (x) { return x.estado === FONDO_EST.APR; }).length, '']);
      p.ver = 'fin';
    } else if (u.rol === 'cmo') {
      p.atrasos = todos.filter(function (a) { return a.quien.split(', ').indexOf(u.nombre) >= 0; });
      var unC = unidadPorId(u.sucursal), eC = unC ? estadoPArea(mes, unC.nombre) : null;
      p.lineas.push(['Presupuesto de ' + periodoTxt(mes), eC ? eC.estado : 'Sin armar', eC && eC.estado === PA_EST.APR ? 'si' : 'no']);
      var avC = null; try { avC = unC ? avanceArea(unC.id, mes) : null; } catch (e) {}
      if (avC && avC.aprobado) p.lineas.push(['Gastado del presupuesto', avC.pct == null ? dinero(avC.gastado) : avC.pct + ' %', avC.pct != null && avC.pct > 100 ? 'no' : 'si']);
      p.lineas.push(['Solicitudes de pago en trámite', fs.filter(function (x) { return x.por === u.nombre && [FONDO_EST.VER, FONDO_ADM, FONDO_EST.PED, FONDO_EST.APR].indexOf(x.estado) >= 0; }).length, '']);
      p.ver = 'fin';
    } else return;
    p.estado = p.atrasos.length ? 'atrasado' : 'al día';
    out.push(p);
  });
  var orden = { operaciones: 0, finanzas: 1, registro: 2, cmo: 3 };
  return out.sort(function (a, b) { return orden[a.rol] - orden[b.rol]; });
}

/* ════════════ TABLERO DE FINANZAS (director financiero y administrador) ════════════ */
function tableroFinanzas(cred) {
  var yo = quien(cred);
  if (!(yo.rol === 'finanzas' || yo.esAdmin || yo.rol === 'dueno')) throw new Error('Este tablero es del director financiero.');
  var hoy = hoyISO(), mes = hoy.slice(0, 7), dias = Number(finDeMes(mes).slice(8)), dHoy = Number(hoy.slice(8));
  var p = presupuestoMes(cred, mes), ep = estadoPres(mes);
  // venta del mes: real acumulada contra el ritmo del presupuesto
  var porDia = {}, porSuc = {};
  ventasRango(mes + '-01', hoy).forEach(function (v) { porDia[v.f] = (porDia[v.f] || 0) + v.m;
    (porSuc[v.u] = porSuc[v.u] || {})[v.f] = ((porSuc[v.u] || {})[v.f] || 0) + v.m; });
  var fechas = [], real = [], ritmo = [], acum = 0;
  for (var d = 1; d <= dias; d++) {
    var f = mes + '-' + ('0' + d).slice(-2); fechas.push(f);
    if (f <= hoy) { acum += porDia[f] || 0; real.push(r2(acum)); } else real.push(null);
    ritmo.push(r2(p.total.venta * d / dias));
  }
  // venta diaria por sucursal, 14 días
  var f14 = []; for (var k = 13; k >= 0; k--) f14.push(masDias(hoy, -k));
  var desde14 = f14[0], v14 = {};
  ventasRango(desde14, hoy).forEach(function (v) { (v14[v.u] = v14[v.u] || {})[v.f] = ((v14[v.u] || {})[v.f] || 0) + v.m; });
  var ventas14 = sucursalesDestino().map(function (u) { return { id: u.id, nombre: u.nombre,
    v: f14.map(function (f) { return f === hoy ? null : r2((v14[u.nombre] || {})[f] || 0); }) }; });
  // presupuesto por unidad: venta y gasto real contra lo presupuestado
  var unidades = p.unidades.map(function (u) {
    var gP = u.costo + u.operacion, gR = u.real.costo + u.real.operacion;
    var pasados = u.rubros.filter(function (r) { return r.presupuesto && r.pct >= 90; }).map(function (r) { return { rubro: r.rubro, pct: r.pct }; });
    return { id: u.id, nombre: u.nombre, vende: u.vende, venta: u.venta, ventaReal: u.real.venta, gasto: r2(gP), gastoReal: r2(gR), pasados: pasados }; });
  // contador: cumplimiento de tareas y horas para recibir el efectivo, 14 días
  var cumple = [];
  try { var tk = tableroKpiAreas([AREA_CONTA], 14); cumple = tk.cumplimiento[AREA_CONTA]; } catch (e) {}
  var horasRec = f14.map(function () { return []; });
  filasPorConfirmar().forEach(function (x) {
    if (x.tipo !== 'ingreso' || !x.recibidoEn || !x.validadoEn) return;
    var i = f14.indexOf(x.recibidoEn.slice(0, 10)); if (i < 0) return;
    var h = horasEntre(x.validadoEn, x.recibidoEn); if (h != null) horasRec[i].push(h); });
  var prom = function (l) { return l.length ? Math.round(l.reduce(function (a, b) { return a + b; }, 0) / l.length * 10) / 10 : null; };
  var cierres = { validar: 0, recibir: 0, banco: 0 };
  filasPorConfirmar().forEach(function (x) { if (x.tipo === 'ingreso') { var s = pasoDe(x); if (s) cierres[s]++; } });
  var fs = []; try { fs = filasFondos(); } catch (e) {}
  var imp = null; try { imp = datosImpuestos().resumen; } catch (e) {}
  var ind = null; try { ind = datosIndicadoresConta(); } catch (e) {}
  return { hoy: hoy, mes: mes, diaHoy: dHoy, diasMes: dias,
    presupuesto: { estado: ep, hay: p.hay, total: p.total }, unidades: unidades,
    venta: { fechas: fechas, real: real, ritmo: ritmo, proyeccion: dHoy ? r2(acum / dHoy * dias) : 0, acumulado: r2(acum) },
    ventas14: { fechas: f14, series: ventas14 },
    contador: { fechas: f14, cumplimiento: cumple, horasRecibir: horasRec.map(prom), ind: ind },
    cierres: cierres, impuestos: imp,
    solicitudes: { aprobar: fs.filter(function (x) { return x.estado === FONDO_EST.PED; }).length,
      pagar: fs.filter(function (x) { return x.estado === FONDO_EST.APR; }).length,
      mes: r2(fs.filter(function (x) { return x.pagEn.slice(0, 7) === mes; }).reduce(function (a, x) { return a + x.pagado; }, 0)) } };
}

/** Correo de prueba al que tiene guardado quien entra; dice exactamente qué pasó. */
function probarCorreo(cred) {
  var yo = quien(cred);
  if (!correoValido(yo.correo)) throw new Error('Primero guarde su correo.');
  var ok = enviaCorreo(yo.correo, 'Rey Pizza · Correo de prueba', htmlCorreo('Correo de prueba',
    '<p>Hola ' + esHtml(yo.nombre) + ': si lee esto, los avisos de la app le van a llegar a este correo.</p>'));
  var er = ok ? null : ultimoErrorCorreo();
  return { ok: ok, mensaje: ok ? 'Listo: salió un correo de prueba a ' + yo.correo + '. Si no lo ve en unos minutos, revise Spam y Promociones.'
    : 'No salió. ' + (er ? explicaErrorCorreo(er.m) : 'Motivo desconocido.') };
}
/** Para el administrador: si la app tiene permiso de mandar correos, si está el
 *  disparador, la cuota y quiénes tienen correo guardado. */
function estadoCorreos(cred) {
  exigeAdmin(cred);
  var out = { permiso: true, cuota: null, disparador: false, diario: false, conCorreo: [], sinCorreo: [], error: '', ultimoOk: '' };
  try { out.cuota = MailApp.getRemainingDailyQuota(); } catch (e) { out.permiso = false; out.error = explicaErrorCorreo(e.message); }
  try { ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'enviarCorreosPendientes') out.disparador = true;
    if (t.getHandlerFunction() === 'reporteDiarioCorreo') out.diario = true; }); }
  catch (e) { out.errorDisparador = 'La app no tiene permiso para crear el aviso automático (se arregla con autorizarCorreos en el editor).'; }
  usuariosCache().forEach(function (u) { if (!u.activo) return; (correoValido(u.correo) ? out.conCorreo : out.sinCorreo).push(u.nombre); });
  var er = ultimoErrorCorreo(); if (er && !out.error) out.error = explicaErrorCorreo(er.m) + ' (' + er.en + ')';
  try { out.ultimoOk = CacheService.getScriptCache().get('correo_ok') || ''; } catch (e) {}
  return out;
}
function activarCorreos(cred) {
  exigeAdmin(cred);
  instalaDisparadorCorreos();
  var r = estadoCorreos(cred); r.mensaje = r.disparador ? 'El envío automático de avisos está activo.' : 'No se pudo activar: ' + (r.errorDisparador || r.error);
  return r;
}
/** ▶ CORRA ESTA FUNCIÓN UNA VEZ DESDE EL EDITOR (Ejecutar). Google le pide los
 *  permisos de correo, Drive y disparadores; después deja activos los avisos cada
 *  5 minutos y el reporte diario, y manda un correo de prueba a cada quien que
 *  tenga su correo guardado. Lea el resultado en el «Registro de ejecución». */
function autorizarCorreos() {
  var l = ['AUTORIZAR CORREOS', ''];
  try { l.push('Cuota de correos que quedan hoy: ' + MailApp.getRemainingDailyQuota()); } catch (e) { l.push('✗ Correo: ' + e.message); }
  try { DriveApp.getRootFolder(); l.push('✓ Drive (cotizaciones): con permiso'); } catch (e) { l.push('✗ Drive: ' + e.message); }
  try { instalaDisparadorCorreos(true); l.push('✓ Avisos por correo cada 5 minutos y reporte diario (21:15): activos'); } catch (e) { l.push('✗ Disparador: ' + e.message); }
  var cuenta = ''; try { cuenta = Session.getEffectiveUser().getEmail(); } catch (e) {}
  if (cuenta) l.push('Los correos salen desde: ' + cuenta + ' (si se manda a esa misma cuenta, Gmail a veces lo deja solo en «Enviados»).');
  usuariosCache().forEach(function (u) {
    if (!u.activo || !correoValido(u.correo)) return;
    var ok = enviaCorreo(u.correo, 'Rey Pizza · Correo de prueba', htmlCorreo('Correo de prueba', '<p>Hola ' + esHtml(u.nombre) +
      ': los avisos de la app ya le llegan a este correo.</p>'));
    l.push((ok ? '✓ Prueba enviada a ' : '✗ No salió a ') + u.nombre + ' <' + u.correo + '>' + (ok ? '' : ': ' + explicaErrorCorreo((ultimoErrorCorreo() || {}).m)));
  });
  var sin = usuariosCache().filter(function (u) { return u.activo && !correoValido(u.correo); }).map(function (u) { return u.nombre; });
  if (sin.length) l.push('Sin correo guardado: ' + sin.join(', '));
  var t = l.join('\n'); Logger.log(t); return t;
}

/* ════════════ REPORTE DIARIO POR CORREO (dueños) ════════════
 * Cada noche (≈ 21:15) a los usuarios «Dueños» con correo guardado: la venta del día
 * por sucursal y turno, contra el mismo día de la semana pasada; el acumulado del
 * mes contra el presupuesto; lo que se pagó a proveedores y los gastos del día. */
function datosReporteDiario(fecha) {
  fecha = esFecha(fecha) ? fecha : hoyISO();
  var sem = masDias(fecha, -7), mes = fecha.slice(0, 7);
  var v = ventasRango(mes + '-01', fecha).concat(fecha.slice(0, 7) === sem.slice(0, 7) ? [] : ventasRango(sem, sem));
  var suc = sucursalesDestino().map(function (u) {
    var d = v.filter(function (x) { return x.u === u.nombre && x.f === fecha; }), s7 = v.filter(function (x) { return x.u === u.nombre && x.f === sem; });
    var t = function (l, k) { return r2(l.reduce(function (a, x) { return a + (x[k] || 0); }, 0)); };
    return { nombre: u.nombre, am: t(d.filter(function (x) { return x.t === 'AM'; }), 'm'), pm: t(d.filter(function (x) { return x.t === 'PM'; }), 'm'),
      total: t(d, 'm'), efectivo: t(d, 'e'), tarjeta: t(d, 'j'), comision: t(d, 'k'), semanaPasada: t(s7, 'm'),
      mes: t(v.filter(function (x) { return x.u === u.nombre && x.f.slice(0, 7) === mes && x.f <= fecha; }), 'm'), turnos: d.length };
  });
  var tot = { total: 0, semanaPasada: 0, mes: 0 };
  suc.forEach(function (s) { tot.total += s.total; tot.semanaPasada += s.semanaPasada; tot.mes += s.mes; });
  var p = null; try { p = presupuestoMes(null, mes, true); } catch (e) {}
  var dias = Number(finDeMes(mes).slice(8)), dHoy = Number(fecha.slice(8));
  var movs = []; try { movs = movimientos(3000).filter(function (m) { return !m.anulado && m.clase !== 'ingreso' && m.fecha === fecha; }); } catch (e) {}
  var pagosOC = []; try { ordenesAgrupadas().forEach(function (o) { if (o.pagadaEn && o.fechaPago === fecha) pagosOC.push({ numero: o.numero, proveedor: o.proveedor, monto: o.monto }); }); } catch (e) {}
  var pend = 0; try { filasPorConfirmar().forEach(function (x) { if (x.tipo === 'ingreso' && pasoDe(x)) pend++; }); } catch (e) {}
  return { fecha: fecha, semana: sem, sucursales: suc, total: r2(tot.total), semanaPasada: r2(tot.semanaPasada), mes: r2(tot.mes),
    presupuestoMes: p ? p.total.venta : 0, ritmo: p ? r2(p.total.venta * dHoy / dias) : 0,
    gastos: movs.map(function (m) { return { unidad: m.unidad, que: m.titulo, proveedor: m.proveedor || '', monto: m.monto }; }),
    pagosOC: pagosOC, cierresPendientes: pend };
}
function htmlReporteDiario(d) {
  var td = 'padding:6px 8px;border-top:1px solid #eef0f3;text-align:right', th = 'padding:6px 8px;text-align:right;color:#6c7079;font-weight:normal';
  var h = '<p style="font-size:15px"><b>Venta del ' + esHtml(dia(d.fecha)) + ': ' + dinero(d.total) + '</b>' +
    (d.semanaPasada ? ' · el mismo día de la semana pasada: ' + dinero(d.semanaPasada) + ' (' + (d.total >= d.semanaPasada ? '+' : '') +
      Math.round((d.total - d.semanaPasada) / d.semanaPasada * 100) + ' %)' : '') + '</p>';
  h += '<table style="width:100%;border-collapse:collapse;font-size:13px"><tr><th style="' + th + ';text-align:left">Sucursal</th><th style="' + th + '">AM</th><th style="' + th +
    '">PM</th><th style="' + th + '">Total</th><th style="' + th + '">Tarjeta</th><th style="' + th + '">Mes</th></tr>' +
    d.sucursales.map(function (s) { return '<tr><td style="' + td + ';text-align:left">' + esHtml(s.nombre) + (s.turnos < 2 ? ' <span style="color:#b3261e">(falta ' + (2 - s.turnos) + ' corte)</span>' : '') +
      '</td><td style="' + td + '">' + dinero(s.am) + '</td><td style="' + td + '">' + dinero(s.pm) + '</td><td style="' + td + '"><b>' + dinero(s.total) + '</b></td><td style="' + td + '">' +
      dinero(s.tarjeta) + '</td><td style="' + td + '">' + dinero(s.mes) + '</td></tr>'; }).join('') + '</table>';
  if (d.presupuestoMes) h += '<p>Acumulado del mes: <b>' + dinero(d.mes) + '</b> · el presupuesto pedía ' + dinero(d.ritmo) + ' a hoy (' +
    Math.round(d.mes / (d.ritmo || 1) * 100) + ' %) · presupuesto del mes ' + dinero(d.presupuestoMes) + '.</p>';
  h += '<p style="margin-top:14px"><b>Pagos a proveedores del día</b></p>' + (d.pagosOC.length ? '<ul>' + d.pagosOC.map(function (o) {
    return '<li>' + esHtml(o.numero) + ' · ' + esHtml(o.proveedor) + ': ' + dinero(o.monto) + '</li>'; }).join('') + '</ul>' : '<p style="color:#6c7079">Ninguno.</p>');
  h += '<p style="margin-top:14px"><b>Gastos registrados del día</b></p>' + (d.gastos.length ? '<ul>' + d.gastos.map(function (g) {
    return '<li>' + esHtml(g.unidad) + ' · ' + esHtml(g.que) + (g.proveedor ? ' (' + esHtml(g.proveedor) + ')' : '') + ': ' + dinero(g.monto) + '</li>'; }).join('') + '</ul>' : '<p style="color:#6c7079">Ninguno.</p>');
  if (d.cierresPendientes) h += '<p style="color:#b3261e">' + d.cierresPendientes + (d.cierresPendientes === 1 ? ' cierre de caja sigue' : ' cierres de caja siguen') + ' sin terminar de confirmar.</p>';
  return h;
}
function reporteDiarioCorreo() {
  var d = datosReporteDiario(hoyISO()), n = 0;
  usuariosCache().forEach(function (u) {
    if (!u.activo || u.rol !== 'dueno' || !correoValido(u.correo)) return;
    if (enviaCorreo(u.correo, 'Rey Pizza · Venta del ' + dia(d.fecha) + ': ' + dinero(d.total), htmlCorreo('Reporte del día', htmlReporteDiario(d)))) n++;
  });
  return n;
}
/** El reporte del día como lo ve el dueño en la app (y para probar el correo). */
function reporteDiario(cred, fecha) {
  var yo = quien(cred);
  if (!(yo.rol === 'dueno' || yo.esAdmin || yo.rol === 'finanzas')) throw new Error('Esto es para los dueños.');
  return datosReporteDiario(fecha);
}


/* ════════════ HORARIOS Y DESCANSO DE LOS GERENTES ════════════
 * El director operativo decide a qué hora quiere a cada gerente en su sucursal
 * (según la demanda de cada día) y qué día descansa: lo asigna él, o el gerente lo
 * pide y él lo aprueba. El descanso solo puede ser de lunes a jueves. */
var H_HOR = ['Sucursal', 'Día', 'Entrada', 'Salida', 'Actualizado por', 'Actualizado en'];
var H_DESC = ['No.', 'Pedido en', 'Gerente', 'Sucursal', 'Día', 'Estado', 'Resuelto por', 'Resuelto en', 'Comentario'];
var DIAS_NOMBRE = ['Domingo', 'Lunes', 'Martes', 'Miércoles', 'Jueves', 'Viernes', 'Sábado'];
var DIAS_DESCANSO = ['Lunes', 'Martes', 'Miércoles', 'Jueves'];
var DESC_EST = { PED: 'Pedido', APR: 'Aprobado', RECH: 'Rechazado', ASIG: 'Asignado' };
function hojaHor(ss) {
  ss = ss || libro();
  var h = ss.getSheetByName('Horarios gerentes');
  if (!h) { h = hojaLimpia(ss, 'Horarios gerentes', H_HOR); h.getRange('B:D').setNumberFormat('@'); h.getRange('F:F').setNumberFormat('dd/mm/yyyy hh:mm'); }
  return h;
}
function hojaDesc(ss) {
  ss = ss || libro();
  var h = ss.getSheetByName('Descansos gerentes');
  if (!h) { h = hojaLimpia(ss, 'Descansos gerentes', H_DESC); h.getRange('B:B').setNumberFormat('dd/mm/yyyy hh:mm'); h.getRange('H:H').setNumberFormat('dd/mm/yyyy hh:mm'); }
  return h;
}
function horaValida(t) { return /^([01]?\d|2[0-3]):[0-5]\d$/.test(String(t || '').trim()); }
function horaNorm(t) { var p = String(t).trim().split(':'); return ('0' + p[0]).slice(-2) + ':' + p[1]; }
function textoHora(v) { if (v instanceof Date) return Utilities.formatDate(v, ZONA, 'HH:mm'); return String(v == null ? '' : v).trim(); }
function filasHor() {
  return leeTodo(hojaHor(), H_HOR.length).map(function (r, i) {
    return { fila: i + 2, sucursal: String(r[0] || ''), dia: String(r[1] instanceof Date ? fmtDia(r[1]) : r[1] || '').trim(),
      entrada: textoHora(r[2]), salida: textoHora(r[3]), por: String(r[4] || ''), en: fmtSello(r[5]) };
  }).filter(function (x) { return x.sucursal && x.dia; });
}
function filasDesc() {
  return leeTodo(hojaDesc(), H_DESC.length).map(function (r, i) {
    return { fila: i + 2, numero: String(r[0] || ''), en: fmtSello(r[1]), gerente: String(r[2] || ''), sucursal: String(r[3] || ''),
      dia: String(r[4] || ''), estado: String(r[5] || ''), resPor: String(r[6] || ''), resEn: fmtSello(r[7]), comentario: String(r[8] || '') };
  }).filter(function (x) { return x.numero; });
}
/** El descanso que vale hoy de un gerente: el último aprobado o asignado. */
function descansoDe(nombre, todos) {
  var d = null;
  (todos || filasDesc()).forEach(function (x) { if (x.gerente === nombre && (x.estado === DESC_EST.APR || x.estado === DESC_EST.ASIG)) d = x; });
  return d;
}
function gerentesActivos() { return usuariosCache().filter(function (u) { return u.activo && u.rol === 'gerente' && u.sucursal; }); }
/** Horario de un gerente en una fecha: la excepción de ese día, o el de ese día de la semana. */
function horarioDe(sid, fecha, hors) {
  var nombreDia = DIAS_NOMBRE[diaSemana(fecha)], sem = null, exc = null;
  (hors || filasHor()).forEach(function (x) { if (x.sucursal !== sid) return; if (x.dia === fecha) exc = x; else if (x.dia === nombreDia) sem = x; });
  var x = exc || sem;
  return x ? { entrada: x.entrada, salida: x.salida, excepcion: !!exc } : null;
}
function semanaHorario(sid, hoy, hors, desc) {
  var out = [];
  for (var k = 0; k < 7; k++) {
    var f = masDias(hoy, k), nd = DIAS_NOMBRE[diaSemana(f)], hr = horarioDe(sid, f, hors);
    out.push({ fecha: f, dia: nd, entrada: hr ? hr.entrada : '', salida: hr ? hr.salida : '', excepcion: hr ? hr.excepcion : false,
      descansa: !!(desc && desc.dia === nd) });
  }
  return out;
}
/** Lo que ve cada quien: el gerente, su semana; el director operativo, todas; el administrador, todas (solo mira). */
function pantallaHorarios(cred) {
  var yo = quien(cred), hoy = hoyISO(), hors = filasHor(), todos = filasDesc();
  var esJefe = yo.rol === 'operaciones' || yo.esAdmin;
  if (!esJefe && yo.rol !== 'gerente') throw new Error('Los horarios de los gerentes los ve el director operativo.');
  var gs = gerentesActivos().filter(function (g) { return esJefe || g.nombre === yo.nombre; });
  var lista = gs.map(function (g) {
    var d = descansoDe(g.nombre, todos), pedido = null;
    todos.forEach(function (x) { if (x.gerente === g.nombre && x.estado === DESC_EST.PED) pedido = x; });
    var semana = semanaHorario(g.sucursal, hoy, hors, d);
    var base = {};
    DIAS_NOMBRE.forEach(function (nd) { var f = ''; for (var k = 0; k < 7; k++) if (DIAS_NOMBRE[diaSemana(masDias(hoy, k))] === nd) f = masDias(hoy, k);
      var x = null; hors.forEach(function (h) { if (h.sucursal === g.sucursal && h.dia === nd) x = h; }); base[nd] = x ? { entrada: x.entrada, salida: x.salida } : { entrada: '', salida: '' }; });
    return { gerente: g.nombre, sucursalId: g.sucursal, sucursal: nombreUnidad(g.sucursal), descanso: d ? d.dia : '', descansoPor: d ? d.resPor : '',
      pedido: pedido ? { numero: pedido.numero, dia: pedido.dia, en: pedido.en } : null, semana: semana, base: base };
  });
  var pedidos = esJefe ? todos.filter(function (x) { return x.estado === DESC_EST.PED; }) : [];
  return { hoy: hoy, dias: DIAS_NOMBRE.slice(1).concat(DIAS_NOMBRE[0]), diasDescanso: DIAS_DESCANSO, gerentes: lista, pedidos: pedidos,
    puedeEditar: yo.rol === 'operaciones' || yo.esAdmin, esGerente: yo.rol === 'gerente' };
}
/** El director operativo fija el horario semanal de la sucursal (o el de una fecha concreta). */
function guardarHorario(cred, sid, dias, fecha) {
  var yo = quien(cred);
  if (yo.rol !== 'operaciones' && !yo.esAdmin) throw new Error('El horario de los gerentes lo fija el director operativo.');
  if (!unidadPorId(sid) || !unidadPorId(sid).vende) throw new Error('Escoja la sucursal.');
  var filas = [];
  if (fecha) {
    if (!esFecha(fecha)) throw new Error('Escoja la fecha.');
    filas = [{ dia: fecha, entrada: (dias || {}).entrada, salida: (dias || {}).salida }];
  } else {
    DIAS_NOMBRE.forEach(function (nd) { var x = (dias || {})[nd]; if (x) filas.push({ dia: nd, entrada: x.entrada, salida: x.salida }); });
  }
  filas.forEach(function (x) {
    var vacia = !String(x.entrada || '').trim() && !String(x.salida || '').trim();
    x.vacia = vacia;
    if (vacia) return;
    if (!horaValida(x.entrada) || !horaValida(x.salida)) throw new Error('Escriba la entrada y la salida como hora, por ejemplo 08:00 y 17:00 (' + x.dia + ').');
    x.entrada = horaNorm(x.entrada); x.salida = horaNorm(x.salida);
    if (x.salida <= x.entrada) throw new Error('La salida tiene que ser después de la entrada (' + x.dia + ').');
  });
  var lock = LockService.getScriptLock(); lock.waitLock(15000); _LEE = {};
  try {
    var h = hojaHor(), actuales = filasHor();
    filas.forEach(function (x) {
      var ya = null; actuales.forEach(function (a) { if (a.sucursal === sid && a.dia === x.dia) ya = a; });
      if (x.vacia) { if (ya) h.getRange(ya.fila, 3, 1, 4).setValues([['', '', yo.nombre, new Date()]]); return; }
      var fila = [sid, x.dia, x.entrada, x.salida, yo.nombre, new Date()];
      if (ya) h.getRange(ya.fila, 1, 1, 6).setValues([fila]);
      else { h.appendRow(fila); h.getRange(h.getLastRow(), 2, 1, 3).setNumberFormat('@').setValues([[x.dia, x.entrada, x.salida]]); }
    });
  } finally { lock.releaseLock(); }
  var r = pantallaHorarios(cred); r.ok = true; r.mensaje = 'Horario de ' + nombreUnidad(sid) + ' guardado. El gerente lo ve en su inicio.';
  return r;
}
function numeroSiguiente(filas, prefijo) {
  var max = 0; filas.forEach(function (x) { var m = new RegExp('^' + prefijo + '-(\\d+)$').exec(x.numero); if (m) max = Math.max(max, +m[1]); });
  return prefijo + '-' + ('0000' + (max + 1)).slice(-4);
}
/** El gerente pide su día de descanso (lunes a jueves). */
function pedirDescanso(cred, diaPedido) {
  var yo = quien(cred);
  if (yo.rol !== 'gerente') throw new Error('Esto lo piden los gerentes.');
  if (DIAS_DESCANSO.indexOf(diaPedido) < 0) throw new Error('El descanso puede ser de lunes a jueves.');
  var lock = LockService.getScriptLock(); lock.waitLock(15000); _LEE = {}; var numero;
  try {
    var todos = filasDesc();
    todos.forEach(function (x) { if (x.gerente === yo.nombre && x.estado === DESC_EST.PED) throw new Error('Ya tiene un pedido de descanso esperando respuesta (' + x.numero + ').'); });
    numero = numeroSiguiente(todos, 'DE');
    hojaDesc().appendRow([numero, new Date(), yo.nombre, yo.sucursal, diaPedido, DESC_EST.PED, '', '', '']);
  } finally { lock.releaseLock(); }
  var r = pantallaHorarios(cred); r.ok = true; r.mensaje = numero + ' enviado. Lo decide el director operativo.';
  return r;
}
/** El director operativo asigna el descanso de un gerente, o responde a su pedido. */
function asignarDescanso(cred, gerente, diaAsignado) {
  var yo = quien(cred);
  if (yo.rol !== 'operaciones' && !yo.esAdmin) throw new Error('El descanso lo asigna el director operativo.');
  if (DIAS_DESCANSO.indexOf(diaAsignado) < 0) throw new Error('El descanso puede ser de lunes a jueves.');
  var g = null; gerentesActivos().forEach(function (u) { if (u.nombre === gerente) g = u; });
  if (!g) throw new Error('No se encontró a ese gerente.');
  var lock = LockService.getScriptLock(); lock.waitLock(15000); _LEE = {};
  try {
    var todos = filasDesc(), h = hojaDesc();
    todos.forEach(function (x) { if (x.gerente === gerente && x.estado === DESC_EST.PED) h.getRange(x.fila, 6, 1, 4).setValues([[DESC_EST.RECH, yo.nombre, new Date(), 'Se le asignó otro día: ' + diaAsignado]]); });
    h.appendRow([numeroSiguiente(todos, 'DE'), new Date(), gerente, g.sucursal, diaAsignado, DESC_EST.ASIG, yo.nombre, new Date(), '']);
  } finally { lock.releaseLock(); }
  var r = pantallaHorarios(cred); r.ok = true; r.mensaje = gerente + ' descansa los ' + diaAsignado.toLowerCase() + '. Lo verá en su inicio y en la campanita.';
  return r;
}
function resolverDescanso(cred, numero, aprobar, comentario) {
  var yo = quien(cred);
  if (yo.rol !== 'operaciones' && !yo.esAdmin) throw new Error('El descanso lo decide el director operativo.');
  var x = null; filasDesc().forEach(function (y) { if (y.numero === String(numero)) x = y; });
  if (!x) throw new Error('No se encontró ese pedido.');
  if (x.estado !== DESC_EST.PED) throw new Error('Ese pedido ya fue respondido.');
  comentario = String(comentario || '').trim().slice(0, 200);
  if (!aprobar && comentario.length < 4) throw new Error('Escriba por qué no se aprueba.');
  hojaDesc().getRange(x.fila, 6, 1, 4).setValues([[aprobar ? DESC_EST.APR : DESC_EST.RECH, yo.nombre, new Date(), comentario]]);
  var r = pantallaHorarios(cred); r.ok = true; r.mensaje = numero + (aprobar ? ' aprobado.' : ' no aprobado.') + ' ' + x.gerente + ' lo verá en su campanita.';
  return r;
}

/* ════════════ MANTENIMIENTO (freidoras, cambio de aceite, equipos) ════════════
 * El director operativo agrega los equipos de cada sucursal y cada cuántos días toca
 * su mantenimiento. El gerente marca cuándo se hizo. La app calcula el próximo
 * vencimiento y avisa. El administrador ve el resumen en su tablero. */
var H_EQ = ['No.', 'Sucursal', 'Equipo', 'Tarea', 'Cada (días)', 'Último hecho', 'Activo', 'Agregado por', 'Agregado en', 'Nota',
  'Tipo de periodo', 'Fecha programada', 'Medir aceite', 'Último registrado en'];
var H_MANT = ['No.', 'Equipo', 'Sucursal', 'Equipo (nombre)', 'Tarea', 'Fecha', 'Hecho por', 'Nota', 'Registrado en'];
var MANT_TAREAS = ['Cambio de aceite', 'Limpieza profunda', 'Mantenimiento preventivo', 'Cambio de filtro', 'Calibración', 'Otro'];
/* Periódico: cada tantos días. Programado: en una fecha que pone el director operativo
 * (una visita técnica, una revisión que no tiene un ritmo fijo). */
var MANT_TIPOS = ['Periódico', 'Programado'];
function hojaEq(ss) {
  ss = ss || libro();
  var h = ss.getSheetByName('Mantenimiento equipos');
  if (!h) { h = hojaLimpia(ss, 'Mantenimiento equipos', H_EQ); h.getRange('F:F').setNumberFormat('@'); h.getRange('L:L').setNumberFormat('@');
    h.getRange('I:I').setNumberFormat('dd/mm/yyyy hh:mm'); h.getRange('N:N').setNumberFormat('dd/mm/yyyy hh:mm'); h.setColumnWidth(3, 220); }
  else if (h.getLastColumn() < H_EQ.length) {
    h.getRange(1, 11, 1, H_EQ.length - 10).setValues([H_EQ.slice(10)])
      .setFontWeight('bold').setFontColor('#ffffff').setBackground('#14284B').setFontSize(10);
    h.getRange('L:L').setNumberFormat('@');
  }
  return h;
}
function hojaMant(ss) {
  ss = ss || libro();
  var h = ss.getSheetByName('Mantenimiento registros');
  if (!h) { h = hojaLimpia(ss, 'Mantenimiento registros', H_MANT); h.getRange('F:F').setNumberFormat('@'); h.getRange('I:I').setNumberFormat('dd/mm/yyyy hh:mm'); h.setColumnWidth(8, 260); }
  return h;
}
function filasEq() {
  return leeTodo(hojaEq(), H_EQ.length).map(function (r, i) {
    var tarea = String(r[3] || '');
    return { fila: i + 2, numero: String(r[0] || ''), sucursal: String(r[1] || ''), equipo: String(r[2] || ''), tarea: tarea,
      cada: Number(r[4]) || 0, ultimo: fmtDia(r[5]), activo: String(r[6]).trim().toLowerCase() !== 'no', por: String(r[7] || ''),
      alta: fmtSello(r[8]).slice(0, 10), nota: String(r[9] || ''),
      tipo: String(r[10] || '').trim() === 'Programado' ? 'Programado' : 'Periódico',
      programada: fmtDia(r[11]), medir: tarea === 'Cambio de aceite' && (String(r[12] == null ? '' : r[12]).trim() === '' || siNo(r[12])), ultimoEn: fmtSello(r[13]) };
  }).filter(function (x) { return x.numero; });
}

/* ── Medición diaria del aceite de las freidoras ──
 * El gerente mide cada día (con el medidor de TPM si lo tiene, o con la escala de color)
 * y marca las señales de desgaste. Con eso la app calcula a qué ritmo se gasta el aceite
 * y recomienda el día del cambio. Límite usual: 24 % de compuestos polares (TPM). */
var ACEITE = { tpmCambio: 24, tpmNuevo: 6, colorCambio: 5,
  colores: ['Claro, como nuevo', 'Dorado', 'Ámbar', 'Café', 'Oscuro'],
  senales: ['Hace espuma que no baja', 'Echa humo a la temperatura normal', 'Huele o sabe fuerte (rancio)',
    'Está espeso o pegajoso', 'Lo frito sale muy oscuro'] };
var H_ACE = ['No.', 'Fecha', 'Sucursal', 'Equipo No.', 'Equipo', 'TPM (%)', 'Color (1 a 5)', 'Señales', 'Nota', 'Medido por',
  'Registrado en', 'Recomienda cambiar', 'Motivo'];
function hojaAceite(ss) {
  ss = ss || libro();
  var h = ss.getSheetByName('Aceite mediciones');
  if (!h) { h = hojaLimpia(ss, 'Aceite mediciones', H_ACE); h.getRange('B:B').setNumberFormat('@'); h.getRange('L:L').setNumberFormat('@');
    h.getRange('K:K').setNumberFormat('dd/mm/yyyy hh:mm'); h.setColumnWidth(8, 260); h.setColumnWidth(13, 300); }
  return h;
}
function filasAceite() {
  return leeTodo(hojaAceite(), H_ACE.length).map(function (r, i) {
    return { fila: i + 2, numero: String(r[0] || ''), fecha: fmtDia(r[1]), sucursal: String(r[2] || ''), equipoNo: String(r[3] || ''),
      equipo: String(r[4] || ''), tpm: numONulo(r[5]), color: Number(r[6]) || 0,
      senales: String(r[7] || '').split(' · ').filter(Boolean), nota: String(r[8] || ''), por: String(r[9] || ''),
      en: fmtSello(r[10]), recomienda: fmtDia(r[11]), motivo: String(r[12] || '') };
  }).filter(function (x) { return x.numero; });
}
/** Las mediciones desde el último cambio de aceite, una por día (la última de ese día). */
function lecturasCiclo(x, todas) {
  var porDia = {};
  todas.forEach(function (l) {
    if (l.equipoNo !== x.numero) return;
    if (x.ultimo && (l.fecha < x.ultimo || (l.fecha === x.ultimo && x.ultimoEn && l.en && l.en < x.ultimoEn))) return;
    porDia[l.fecha] = l;
  });
  return Object.keys(porDia).sort().map(function (f) { return porDia[f]; });
}
/** Pendiente por mínimos cuadrados (cuánto sube por día). */
function pendienteAceite(pts) {
  if (pts.length < 2) return null;
  var n = pts.length, sx = 0, sy = 0;
  pts.forEach(function (p) { sx += p[0]; sy += p[1]; });
  var mx = sx / n, my = sy / n, num = 0, den = 0;
  pts.forEach(function (p) { num += (p[0] - mx) * (p[1] - my); den += (p[0] - mx) * (p[0] - mx); });
  return den ? num / den : null;
}
function recomiendaAceite(x, lecturas, hoy) {
  if (!lecturas.length) return null;
  var u = lecturas[lecturas.length - 1], base = x.ultimo || lecturas[0].fecha, ns = u.senales.length;
  var fin = function (dias, motivo) {
    var f = masDias(u.fecha, Math.max(0, dias)); if (f < hoy) f = hoy;
    var d = diasEntre(hoy, f);
    return { fecha: f, dias: d, motivo: motivo, nivel: d <= 0 ? 'hoy' : d <= 1 ? 'pronto' : 'ok', ultima: u };
  };
  if (u.tpm != null && u.tpm >= ACEITE.tpmCambio) return fin(0, 'Marca ' + num(u.tpm) + ' % de TPM: ya llegó al límite (' + ACEITE.tpmCambio + ' %). Cámbielo hoy.');
  if (u.color >= ACEITE.colorCambio) return fin(0, 'El aceite ya está oscuro (5 de 5). Cámbielo hoy.');
  if (ns >= 2) return fin(0, 'Tiene ' + ns + ' señales de aceite gastado. Cámbielo hoy.');
  var dias = null, motivo = '', ritmo;
  if (u.tpm != null) {
    var pts = lecturas.filter(function (l) { return l.tpm != null; }).map(function (l) { return [diasEntre(base, l.fecha), l.tpm]; });
    if (pts.length < 2 && x.ultimo && diasEntre(x.ultimo, u.fecha) > 0) pts.unshift([0, ACEITE.tpmNuevo]);
    ritmo = pendienteAceite(pts);
    motivo = 'Marca ' + num(u.tpm) + ' % de TPM';
    if (ritmo != null && ritmo > 0.05) { dias = Math.floor((ACEITE.tpmCambio - u.tpm) / ritmo); motivo += ' y sube cerca de ' + num(Math.round(ritmo * 10) / 10) + ' % por día.'; }
    else motivo += ritmo == null ? '. Con una medición más se calcula el ritmo.' : ' y casi no ha subido.';
  } else if (u.color) {
    var pc = lecturas.filter(function (l) { return l.color; }).map(function (l) { return [diasEntre(base, l.fecha), l.color]; });
    if (pc.length < 2 && x.ultimo && diasEntre(x.ultimo, u.fecha) > 0) pc.unshift([0, 1]);
    ritmo = pendienteAceite(pc);
    motivo = 'Color ' + u.color + ' de 5 (' + ACEITE.colores[u.color - 1].toLowerCase() + ')';
    if (ritmo != null && ritmo > 0.02) { dias = Math.floor((ACEITE.colorCambio - u.color) / ritmo); motivo += ' y se oscurece cerca de ' + num(Math.round(ritmo * 10) / 10) + ' por día.'; }
    else motivo += ritmo == null ? '. Con una medición más se calcula el ritmo.' : ' y casi no ha cambiado.';
  }
  if (ns === 1) { dias = dias == null ? 1 : Math.min(dias, 1); motivo += ' Señal de desgaste: ' + u.senales[0].toLowerCase() + '.'; }
  if (dias == null) return { fecha: '', dias: null, motivo: motivo + ' Siga midiendo cada día.', nivel: 'ok', ultima: u };
  return fin(dias, motivo);
}

/** Estado de un equipo hoy: cuándo toca y qué tan avanzado va el plazo. */
function estadoEq(x, hoy, lecturas) {
  var base = x.ultimo || x.alta || hoy, prox = '', umbral, total;
  if (x.tipo === 'Programado') {
    prox = x.programada || '';
    total = prox ? Math.max(1, diasEntre(base, prox)) : 0;
    umbral = 3;
  } else {
    prox = masDias(base, x.cada || 1); total = x.cada || 1;
    umbral = Math.max(1, Math.min(3, Math.round(x.cada * 0.15)));
  }
  var med = null, porMedicion = false;
  if (x.medir) {
    med = recomiendaAceite(x, lecturas || [], hoy);
    if (med && med.fecha && (!prox || med.fecha < prox)) { prox = med.fecha; porMedicion = true; umbral = 1; }
  }
  var medidoHoy = !!(med && med.ultima && med.ultima.fecha === hoy);
  if (!prox) return { proximo: '', restan: null, estado: 'sin fecha', avance: 0, sinRegistro: !x.ultimo,
    medicion: med, porMedicion: false, medidoHoy: medidoHoy, lecturas: (lecturas || []).slice(-10), toca: false, habilita: '' };
  var rest = diasEntre(hoy, prox);
  var estado = rest < 0 ? 'vencido' : rest === 0 ? 'hoy' : rest <= umbral ? 'pronto' : 'al dia';
  // se puede marcar cuando ya toca (o está por tocar), o si nunca se ha anotado; si no, queda bloqueado
  var toca = estado !== 'al dia' || (!x.ultimo && x.ultimo !== hoy), habilita = toca ? '' : masDias(prox, -umbral);
  var hecho = diasEntre(base, hoy), largo = porMedicion ? Math.max(1, diasEntre(base, prox)) : total;
  var avance = largo > 0 ? Math.max(0, Math.min(100, Math.round(hecho / largo * 100))) : 0;
  return { proximo: prox, restan: rest, estado: estado, avance: avance, sinRegistro: !x.ultimo,
    medicion: med, porMedicion: porMedicion, medidoHoy: medidoHoy, lecturas: (lecturas || []).slice(-10), toca: toca, habilita: habilita };
}
function equiposVista(yo) {
  var hoy = hoyISO(), eqs = filasEq().filter(function (x) { return x.activo && (yo.rol !== 'gerente' || x.sucursal === yo.sucursal); });
  var ace = eqs.some(function (x) { return x.medir; }) ? filasAceite() : [];
  return eqs.map(function (x) {
    var e = estadoEq(x, hoy, x.medir ? lecturasCiclo(x, ace) : null); e.sucursalNombre = nombreUnidad(x.sucursal);
    Object.keys(e).forEach(function (k) { x[k] = e[k]; });
    return x;
  }).sort(function (a, b) { return (a.restan == null ? 99999 : a.restan) - (b.restan == null ? 99999 : b.restan); });
}
function pantallaMant(cred) {
  var yo = quien(cred), esJefe = yo.rol === 'operaciones' || yo.esAdmin;
  if (!esJefe && yo.rol !== 'gerente') throw new Error('El mantenimiento lo ven los gerentes y el director operativo.');
  var hist = leeTodo(hojaMant(), H_MANT.length).map(function (r) {
    return { numero: String(r[0] || ''), equipoNo: String(r[1] || ''), sucursal: String(r[2] || ''), equipo: String(r[3] || ''), tarea: String(r[4] || ''),
      fecha: fmtDia(r[5]), por: String(r[6] || ''), nota: String(r[7] || ''), en: fmtSello(r[8]) };
  }).filter(function (x) { return x.numero && (yo.rol !== 'gerente' || x.sucursal === nombreUnidad(yo.sucursal)); }).reverse().slice(0, 40);
  return { hoy: hoyISO(), equipos: equiposVista(yo), historial: hist, tareas: MANT_TAREAS, tipos: MANT_TIPOS,
    aceite: { colores: ACEITE.colores, senales: ACEITE.senales, tpmCambio: ACEITE.tpmCambio },
    sucursales: UNIDADES.filter(function (u) { return u.vende; }).map(function (u) { return { id: u.id, nombre: u.nombre }; }),
    puedeAgregar: yo.rol === 'operaciones' || yo.esAdmin, puedeRegistrar: yo.rol === 'gerente' || yo.rol === 'operaciones' || yo.esAdmin,
    puedeMedir: yo.rol === 'gerente' || yo.rol === 'operaciones', esGerente: yo.rol === 'gerente' };
}
/** Lee tipo, cada y fecha programada del formulario del director. */
function periodoMant(d, hoy) {
  var tipo = MANT_TIPOS.indexOf(d.tipo) >= 0 ? d.tipo : 'Periódico', cada = 0, prog = '';
  if (tipo === 'Periódico') {
    cada = Math.round(Number(String(d.cada || '').replace(',', '.')));
    if (!(cada >= 1 && cada <= 730)) throw new Error('Escriba cada cuántos días toca (de 1 a 730).');
  } else {
    prog = esFecha(d.programada) ? d.programada : '';
    if (prog && prog < hoy) throw new Error('La fecha programada no puede ser en el pasado.');
  }
  return { tipo: tipo, cada: cada, programada: prog };
}
function agregarEquipo(cred, d) {
  var yo = quien(cred);
  if (yo.rol !== 'operaciones' && !yo.esAdmin) throw new Error('Los equipos los agrega el director operativo.');
  d = d || {};
  var equipo = String(d.equipo || '').trim().slice(0, 60);
  if (equipo.length < 3) throw new Error('Escriba el nombre del equipo (por ejemplo «Freidora 1»).');
  var tarea = MANT_TAREAS.indexOf(d.tarea) >= 0 ? d.tarea : '';
  if (!tarea) throw new Error('Escoja la tarea.');
  var hoy = hoyISO(), per = periodoMant(d, hoy);
  var medir = tarea === 'Cambio de aceite' && d.medir !== false;
  var sucs = (d.sucursales && d.sucursales.length ? d.sucursales : [d.sucursal]).filter(function (s) { return unidadPorId(s) && unidadPorId(s).vende; });
  if (!sucs.length) throw new Error('Escoja la sucursal.');
  var ultimo = esFecha(d.ultimo) ? d.ultimo : '';
  if (ultimo && ultimo > hoy) throw new Error('La última vez no puede ser una fecha futura.');
  var lock = LockService.getScriptLock(); lock.waitLock(15000); _LEE = {};
  try {
    var todos = filasEq(), h = hojaEq();
    sucs.forEach(function (s) {
      var numero = numeroSiguiente(todos, 'EQ');
      var fila = [numero, s, equipo, tarea, per.cada || '', ultimo, 'Sí', yo.nombre, new Date(), String(d.nota || '').slice(0, 120),
        per.tipo, per.programada, medir ? 'Sí' : 'No', ''];
      h.appendRow(fila);
      var f = h.getLastRow(); h.getRange(f, 6).setNumberFormat('@').setValue(ultimo); h.getRange(f, 12).setNumberFormat('@').setValue(per.programada);
      todos.push({ numero: numero });
    });
  } finally { lock.releaseLock(); }
  var r = pantallaMant(cred); r.ok = true; r.mensaje = equipo + ' agregado a ' + sucs.map(nombreUnidad).join(', ') + '.';
  return r;
}
function cambiarEquipo(cred, numero, d) {
  var yo = quien(cred);
  if (yo.rol !== 'operaciones' && !yo.esAdmin) throw new Error('Los equipos los cambia el director operativo.');
  var x = null; filasEq().forEach(function (y) { if (y.numero === String(numero)) x = y; });
  if (!x) throw new Error('No se encontró ese equipo.');
  d = d || {};
  var h = hojaEq();
  if (d.activo === false) {
    h.getRange(x.fila, 7).setValue('No');
    var r0 = pantallaMant(cred); r0.ok = true; r0.mensaje = x.equipo + ' ya no se sigue.'; return r0;
  }
  var per = periodoMant({ tipo: d.tipo || x.tipo, cada: d.cada != null ? d.cada : x.cada,
    programada: d.programada != null ? d.programada : x.programada }, hoyISO());
  h.getRange(x.fila, 5).setValue(per.cada || '');
  h.getRange(x.fila, 11).setValue(per.tipo);
  h.getRange(x.fila, 12).setNumberFormat('@').setValue(per.programada);
  if (x.tarea === 'Cambio de aceite' && d.medir != null) h.getRange(x.fila, 13).setValue(d.medir ? 'Sí' : 'No');
  var r = pantallaMant(cred); r.ok = true;
  r.mensaje = x.equipo + ': ' + (per.tipo === 'Periódico' ? 'cada ' + per.cada + ' días' : per.programada ? 'programado para el ' + dia(per.programada).slice(0, 5) : 'programado, sin fecha todavía') + '.';
  return r;
}
function registrarMant(cred, numero, d) {
  var yo = quien(cred);
  if (yo.rol !== 'gerente' && yo.rol !== 'operaciones' && !yo.esAdmin) throw new Error('El mantenimiento lo registra el gerente.');
  var x = null; filasEq().forEach(function (y) { if (y.numero === String(numero) && y.activo) x = y; });
  if (!x) throw new Error('No se encontró ese equipo.');
  if (yo.rol === 'gerente' && x.sucursal !== yo.sucursal) throw new Error('Ese equipo es de otra sucursal.');
  d = d || {};
  var hoy = hoyISO(), fecha = esFecha(d.fecha) ? d.fecha : hoy;
  if (fecha > hoy) throw new Error('La fecha no puede ser futura.');
  if (yo.rol === 'gerente' && diasEntre(fecha, hoy) > 7) throw new Error('Solo se puede anotar hasta 7 días atrás. Si fue antes, avise al director operativo.');
  if (yo.rol === 'gerente') {
    var ace = x.medir ? filasAceite() : [], e = estadoEq(x, hoy, x.medir ? lecturasCiclo(x, ace) : null);
    if (x.ultimo === hoy) throw new Error('Esto ya se anotó hoy.');
    if (!e.toca) throw new Error(e.estado === 'sin fecha' ? 'Todavía no tiene fecha: el director operativo la programa.'
      : 'Todavía no toca. Se habilita el ' + dia(e.habilita).slice(0, 5) + '.');
  }
  var nota = String(d.nota || '').trim().slice(0, 200);
  var lock = LockService.getScriptLock(); lock.waitLock(15000); _LEE = {};
  try {
    var hm = hojaMant(), n = numeroSiguiente(leeTodo(hm, 1).map(function (r) { return { numero: String(r[0]) }; }), 'MT');
    hm.appendRow([n, x.numero, nombreUnidad(x.sucursal), x.equipo, x.tarea, fecha, yo.nombre, nota, new Date()]);
    hm.getRange(hm.getLastRow(), 6).setNumberFormat('@').setValue(fecha);
    var he = hojaEq();
    if (!x.ultimo || fecha >= x.ultimo) { he.getRange(x.fila, 6).setNumberFormat('@').setValue(fecha); he.getRange(x.fila, 14).setValue(new Date()); }
    if (x.tipo === 'Programado' && (!x.programada || fecha >= masDias(x.programada, -15))) he.getRange(x.fila, 12).setNumberFormat('@').setValue('');
  } finally { lock.releaseLock(); }
  var r = pantallaMant(cred); r.ok = true;
  r.mensaje = x.tarea + ' de ' + x.equipo + ' anotado. ' + (x.tipo === 'Programado' ? 'La próxima fecha la pone el director operativo.'
    : 'Toca de nuevo en ' + x.cada + (x.cada === 1 ? ' día' : ' días') + (x.medir ? ' o antes, según la medición del aceite' : '') + '.');
  return r;
}
/** El gerente anota la medición del aceite del día y la app le dice cuándo cambiarlo. */
function medirAceite(cred, numero, d) {
  var yo = quien(cred);
  if (yo.rol !== 'gerente' && yo.rol !== 'operaciones') throw new Error('El aceite lo mide el gerente.');
  var x = null; filasEq().forEach(function (y) { if (y.numero === String(numero) && y.activo) x = y; });
  if (!x || !x.medir) throw new Error('A ese equipo no se le mide el aceite.');
  if (yo.rol === 'gerente' && x.sucursal !== yo.sucursal) throw new Error('Ese equipo es de otra sucursal.');
  d = d || {};
  var tpm = String(d.tpm == null ? '' : d.tpm).trim() === '' ? null : Number(String(d.tpm).replace(',', '.'));
  if (tpm != null && !(isFinite(tpm) && tpm >= 0 && tpm <= 45)) throw new Error('El TPM va de 0 a 45 %. Revise el número del medidor.');
  var color = Math.round(Number(d.color) || 0);
  if (color && !(color >= 1 && color <= 5)) throw new Error('Escoja el color del 1 al 5.');
  if (tpm == null && !color) throw new Error('Anote el TPM del medidor o escoja el color del aceite.');
  var sen = (d.senales || []).filter(function (s) { return ACEITE.senales.indexOf(s) >= 0; });
  var nota = String(d.nota || '').trim().slice(0, 200), hoy = hoyISO(), ahora = new Date();
  if (yo.rol === 'gerente' && filasAceite().some(function (l) { return l.equipoNo === x.numero && l.fecha === hoy; }))
    throw new Error('El aceite de ' + x.equipo + ' ya se midió hoy. La próxima medición es mañana.');
  var rec = null, lock = LockService.getScriptLock(); lock.waitLock(15000); _LEE = {};
  try {
    var todas = filasAceite();
    var nueva = { numero: '', equipoNo: x.numero, fecha: hoy, tpm: tpm == null ? null : Math.round(tpm * 10) / 10, color: color, senales: sen, en: fmtSello(ahora) };
    var ciclo = lecturasCiclo(x, todas).filter(function (l) { return l.fecha !== hoy; });
    ciclo.push(nueva);
    rec = recomiendaAceite(x, ciclo, hoy);
    var h = hojaAceite(), previa = null;
    todas.forEach(function (l) { if (l.equipoNo === x.numero && l.fecha === hoy) previa = l; });
    var fila = [previa ? previa.numero : numeroSiguiente(todas, 'AC'), hoy, nombreUnidad(x.sucursal), x.numero, x.equipo,
      nueva.tpm == null ? '' : nueva.tpm, color || '', sen.join(' · '), nota, yo.nombre, ahora, rec && rec.fecha || '', rec ? rec.motivo : ''];
    var f = previa ? previa.fila : h.getLastRow() + 1;
    h.getRange(f, 1, 1, H_ACE.length).setValues([fila]);
    h.getRange(f, 2).setNumberFormat('@').setValue(hoy); h.getRange(f, 12).setNumberFormat('@').setValue(rec && rec.fecha || '');
  } finally { lock.releaseLock(); }
  var r = pantallaMant(cred); r.ok = true; r.recomendacion = rec;
  r.mensaje = 'Medición de ' + x.equipo + ' guardada. ' + (!rec || !rec.fecha ? (rec ? rec.motivo : '')
    : rec.nivel === 'hoy' ? 'Recomendación: cambie el aceite hoy. ' + rec.motivo
    : 'Recomendación: cambie el aceite el ' + DIAS_NOMBRE[aDate(rec.fecha).getDay()].toLowerCase() + ' ' + dia(rec.fecha).slice(0, 5) +
      ' (en ' + rec.dias + (rec.dias === 1 ? ' día' : ' días') + '). ' + rec.motivo);
  return r;
}
/** Para el tablero del administrador. */
function resumenMant() {
  var eqs = equiposVista({ rol: 'admin' }), porSuc = {};
  UNIDADES.forEach(function (u) { if (u.vende) porSuc[u.id] = { nombre: u.nombre, total: 0, vencidos: 0, pronto: 0 }; });
  eqs.forEach(function (x) { var s = porSuc[x.sucursal]; if (!s) return; s.total++; if (x.estado === 'vencido') s.vencidos++; if (x.estado === 'pronto' || x.estado === 'hoy') s.pronto++; });
  var hoy = hoyISO();
  return { total: eqs.length, vencidos: eqs.filter(function (x) { return x.estado === 'vencido'; }).length,
    pronto: eqs.filter(function (x) { return x.estado === 'pronto' || x.estado === 'hoy'; }).length,
    sinFecha: eqs.filter(function (x) { return x.estado === 'sin fecha'; }).length,
    sinMedir: eqs.filter(function (x) { return x.medir && !x.medidoHoy; }).length,
    sucursales: Object.keys(porSuc).map(function (k) { return porSuc[k]; }),
    urgentes: eqs.filter(function (x) { return x.estado !== 'al dia' && x.estado !== 'sin fecha'; }).slice(0, 8).map(function (x) {
      return { numero: x.numero, sucursal: x.sucursalNombre, equipo: x.equipo, tarea: x.tarea, proximo: x.proximo, restan: x.restan, estado: x.estado,
        porMedicion: x.porMedicion }; }) };
}

/* ════════════ CAJA DE SUGERENCIAS ════════════
 * Cualquiera con usuario deja una sugerencia (con su nombre o sin él). Solo el
 * director operativo y el administrador las leen. */
var H_SUG = ['No.', 'Enviada en', 'De', 'Rol', 'Sucursal', 'Sugerencia', 'Anónima', 'Leída por', 'Leída en'];
function hojaSug(ss) {
  ss = ss || libro();
  var h = ss.getSheetByName('Sugerencias');
  if (!h) { h = hojaLimpia(ss, 'Sugerencias', H_SUG); h.getRange('B:B').setNumberFormat('dd/mm/yyyy hh:mm'); h.getRange('I:I').setNumberFormat('dd/mm/yyyy hh:mm'); h.setColumnWidth(6, 420); }
  return h;
}
function filasSug() {
  return leeTodo(hojaSug(), H_SUG.length).map(function (r, i) {
    return { fila: i + 2, numero: String(r[0] || ''), en: fmtSello(r[1]), de: String(r[2] || ''), rol: String(r[3] || ''), sucursal: String(r[4] || ''),
      texto: String(r[5] || ''), anonima: siNo(r[6]), leidaPor: String(r[7] || ''), leidaEn: fmtSello(r[8]) };
  }).filter(function (x) { return x.numero; });
}
function leeSug(yo) { return yo.rol === 'operaciones' || yo.esAdmin; }
function pantallaSugerencias(cred) {
  permiteCmo();
  var yo = quien(cred), todas = filasSug();
  var out = { puedeEnviar: yo.rol !== 'dueno', puedeLeer: leeSug(yo), mias: [] };
  if (out.puedeLeer) {
    out.lista = todas.slice().reverse().slice(0, 100).map(function (x) {
      return { numero: x.numero, en: x.en, de: x.anonima ? 'Anónima' : x.de, rol: x.anonima ? '' : x.rol, sucursal: x.anonima ? '' : x.sucursal, texto: x.texto, leida: !!x.leidaPor }; });
    out.sinLeer = todas.filter(function (x) { return !x.leidaPor; }).length;
  }
  out.mias = todas.filter(function (x) { return x.de === yo.nombre && !x.anonima; }).reverse().slice(0, 5).map(function (x) { return { numero: x.numero, en: x.en, texto: x.texto }; });
  return out;
}
function enviarSugerencia(cred, texto, anonima) {
  permiteCmo();
  var yo = quien(cred);
  if (yo.rol === 'dueno') throw new Error('Su usuario no envía sugerencias desde la app.');
  texto = String(texto || '').trim().slice(0, 600);
  if (texto.length < 8) throw new Error('Escriba la sugerencia (al menos una frase).');
  var lock = LockService.getScriptLock(); lock.waitLock(15000); _LEE = {}; var numero;
  try {
    numero = numeroSiguiente(filasSug(), 'SG');
    hojaSug().appendRow([numero, new Date(), anonima ? '' : yo.nombre, anonima ? '' : yo.rol, anonima ? '' : (yo.sucursal ? nombreUnidad(yo.sucursal) : ''), texto, anonima ? 'Sí' : 'No', '', '']);
  } finally { lock.releaseLock(); }
  var r = pantallaSugerencias(cred); r.ok = true; r.mensaje = 'Gracias. Su sugerencia (' + numero + ') le llegó al director operativo y al administrador.';
  return r;
}
function leerSugerencias(cred) {
  var yo = quien(cred);
  if (!leeSug(yo)) throw new Error('Las sugerencias las leen el director operativo y el administrador.');
  var h = hojaSug(), ahora = new Date();
  filasSug().forEach(function (x) { if (!x.leidaPor) h.getRange(x.fila, 8, 1, 2).setValues([[yo.nombre, ahora]]); });
  var r = pantallaSugerencias(cred); r.sinLeer = 0; return r;
}

/* ════════════ PRODUCTOS DESDE LA CONFIGURACIÓN ════════════
 * Alvaro y Samuel agregan productos nuevos (para una sucursal, varias o todas).
 * Ellos y los gerentes deciden cómo se cuenta cada producto: por unidad (cajas,
 * bolsas…) o por peso (libras, onzas…). El conteo lo muestra ya configurado. */
function puedeProductos(yo) { return yo.esAdmin || yo.rol === 'operaciones' || yo.rol === 'bodega'; }
function unidadesConteo() { return UNIDADES_TRASLADO.filter(function (u) { return PESOS.indexOf(u) < 0; }); }
function productosConfig(cred) {
  var yo = quien(cred);
  if (!puedeProductos(yo) && yo.rol !== 'gerente') throw new Error('Su usuario no cambia los productos.');
  var grupos = {}, ver = puedeProductos(yo);
  var lista = filasBodega().filter(function (b) {
    return ver ? true : (b.activo && b.suc[yo.sucursal]);
  }).map(function (b) {
    if (b.activo) grupos[b.grupo] = true;
    var o = { codigo: b.codigo, nombre: b.nombre, grupo: b.grupo, area: b.area, unidad: b.unidad, medida: b.medida,
      suelto: b.suelto, porUnidad: b.porUnidad, menor: b.menor, porSuelto: b.porSuelto, pesoCada: b.pesoCada, pesoEn: b.pesoEn,
      presentacion: b.presentacion, descripcion: b.descripcion, activo: b.activo, suc: b.suc };
    if (ver) { o.proveedor = b.proveedor; o.proveedor2 = b.proveedor2; o.precio = b.precioPres; o.presCompra = b.presCompra; o.rinde = b.rinde;
      o.precioEn = b.precioEn; o.precioPor = b.precioPor; o.precioAnt = b.precioAnt; }
    if (!b.activo && b.desactPor) { o.desactPor = b.desactPor; o.desactEn = b.desactEn ? fmtSello(b.desactEn) : ''; }
    return o;
  });
  return { productos: lista, unidades: unidadesConteo(), pesos: PESOS, grupos: Object.keys(grupos).sort(), areas: AREAS_BODEGA,
    categorias: CATEGORIAS, puedeAgregar: puedeProductos(yo), esGerente: yo.rol === 'gerente',
    proveedores: ver ? filasProveedores().filter(function (p) { return p.activo !== false; }).map(function (p) { return p.nombre; }).sort() : [],
    presentacionesCompra: PRES_COMPRA,
    sucursales: UNIDADES.filter(function (u) { return u.vende; }).map(function (u) { return { id: u.id, nombre: u.nombre }; }) };
}
/** El nombre de un producto, ordenado: mayúsculas, sin espacios de más. */
function nombreProducto(v) {
  var n = String(v == null ? '' : v).replace(/\s+/g, ' ').trim().toUpperCase().slice(0, 80);
  if (n.length < 3) throw new Error('Escriba el nombre del producto (mínimo 3 letras).');
  return n;
}
/** Editar un producto (nombre, categoría, área, descripción y cómo se cuenta) en un solo paso.
 *  El código no cambia nunca: los conteos y traslados viejos siguen ligados a él. */
function editarProducto(cred, codigo, d) {
  var yo = quien(cred);
  if (!puedeProductos(yo)) throw new Error('Los productos los edita el director operativo o la bodega.');
  d = d || {};
  var b = null; filasBodega().forEach(function (x) { if (x.codigo === String(codigo)) b = x; });
  if (!b) throw new Error('No se encontró ese producto.');
  var nombre = nombreProducto(d.nombre);
  var area = AREAS_BODEGA.indexOf(d.area) >= 0 ? d.area : b.area;
  var grupo = String(d.grupo || '').trim() ? nombreCategoria(d.grupo) : b.grupo;
  filasBodega().forEach(function (x) { if (x.grupo.toLowerCase() === grupo.toLowerCase()) grupo = x.grupo; });
  var desc = String(d.descripcion == null ? '' : d.descripcion).replace(/\s+/g, ' ').trim().slice(0, 200);
  var m = medidaValida(d), extra = null;
  var lock = LockService.getScriptLock(); lock.waitLock(15000); _LEE = {};
  try {
    filasBodega().forEach(function (x) {
      if (x.codigo !== b.codigo && x.nombre.toUpperCase() === nombre) throw new Error('Ya existe «' + x.nombre + '» en el catálogo (' + x.codigo + ').');
    });
    var h = hojaBodega(), ahora = new Date();
    h.getRange(b.fila, 2).setValue(nombre);
    h.getRange(b.fila, 3).setValue(area);
    h.getRange(b.fila, 4).setValue(grupo);
    h.getRange(b.fila, 9, 1, 3).setValues([[m.unidad, m.suelto, m.porUnidad]]);
    h.getRange(b.fila, 17, 1, 2).setValues([[ahora, yo.nombre]]);
    h.getRange(b.fila, 25).setValue(m.suelto ? (b.mandaSuelto ? 'Sí' : 'No') : 'No');
    h.getRange(b.fila, 27).setValue(m.medida);
    h.getRange(b.fila, 28, 1, 4).setValues([[m.menor, m.porSuelto, m.pesoCada, m.pesoEn]]);
    h.getRange(b.fila, 32).setValue(desc);
    extra = editarProductoExtra(yo, b, d, h, ahora);
    _BOD = null;
  } finally { lock.releaseLock(); }
  if (extra && extra.cambioPrecio) { try { avisaCambioPrecioProducto(yo, b, extra.cambioPrecio); } catch (e) {} }
  var r = productosConfig(cred); r.ok = true;
  r.mensaje = nombre + ' quedó guardado.' + (extra && extra.mensaje ? ' ' + extra.mensaje : '') + (nombre !== b.nombre.toUpperCase() ? ' Antes se llamaba «' + b.nombre + '»; lo ya registrado conserva el nombre de ese día.' : '');
  return r;
}
/** Activar o desactivar un producto en todo el sistema (nunca se elimina). Al desactivar, le llega el aviso al director operativo. */
function activarProducto(cred, codigo, activo) {
  var yo = quien(cred);
  if (!puedeProductos(yo)) throw new Error('Los productos los activa o desactiva el director operativo o la bodega.');
  var b = null; filasBodega().forEach(function (x) { if (x.codigo === String(codigo)) b = x; });
  if (!b) throw new Error('No se encontró ese producto.');
  activo = activo === true || activo === 'true' || activo === 'Sí';
  var h = hojaBodega(), ahora = new Date();
  if (activo === b.activo) { var r0 = productosConfig(cred); r0.ok = true; r0.mensaje = b.nombre + (activo ? ' ya estaba activo.' : ' ya estaba desactivado.'); return r0; }
  h.getRange(b.fila, 16).setValue(activo ? 'Sí' : 'No');
  h.getRange(b.fila, 17, 1, 2).setValues([[ahora, yo.nombre]]);
  h.getRange(b.fila, 33, 1, 2).setValues([activo ? ['', ''] : [yo.nombre, ahora]]);
  _BOD = null;
  var r = productosConfig(cred); r.ok = true;
  r.mensaje = b.nombre + (activo ? ' se activó.' : ' se desactivó' + (yo.rol === 'operaciones' ? '.' : '. Le llegó el aviso al director operativo.'));
  return r;
}
/** Aviso al director operativo: productos que otra persona desactivó en los últimos días. */
function avisosProductos(yo, out, desde) {
  if (yo.rol !== 'operaciones') return;
  filasBodega().forEach(function (b) {
    if (b.activo || !b.desactEn || !b.desactPor || b.desactPor === yo.nombre) return;
    if (fmtDia(b.desactEn) < desde) return;
    out.push({ id: 'prd-off-' + b.codigo + '-' + fmtDia(b.desactEn) + '-' + Utilities.formatDate(b.desactEn, ZONA, 'HHmm'),
      tipo: 'Producto desactivado', cuando: fmtSello(b.desactEn), titulo: b.nombre,
      texto: 'Lo desactivó ' + b.desactPor + ' para todas las sucursales. Ya no sale en conteos ni en pedidos.', abrir: { prod: true } });
  });
}
function codigoProducto(nombre) {
  var base = sinAcentoS(nombre).toUpperCase().replace(/[^A-Z0-9 ]/g, '').split(/\s+/).filter(function (w) { return w.length > 0; });
  var c = base.map(function (w) { return w.slice(0, 3); }).join('').slice(0, 6) || 'PROD';
  var usados = {}; filasBodega().forEach(function (b) { usados[b.codigo.toUpperCase()] = true; });
  var cand = c, n = 1;
  while (usados[cand]) { n++; cand = c.slice(0, 6) + n; }
  return cand.slice(0, 12);
}
function medidaValida(d) {
  d = d || {};
  var medida = d.medida === 'peso' ? 'peso' : 'unidad';
  var unidad = singular(String(d.unidad || '').slice(0, 20));
  if (medida === 'peso') {
    if (PESOS.indexOf(unidad) < 0) throw new Error('Escoja en qué se pesa: ' + PESOS.join(', ') + '.');
    return { medida: 'peso', unidad: unidad, suelto: '', porUnidad: '', menor: '', porSuelto: '', pesoCada: '', pesoEn: '' };
  }
  if (!unidad || PESOS.indexOf(unidad) >= 0) unidad = unidad && PESOS.indexOf(unidad) < 0 ? unidad : 'unidad';
  var cifraN = function (v) { return Math.round((Number(String(v == null ? '' : v).replace(',', '.')) || 0) * 1000) / 1000; };
  var suelto = singular(String(d.suelto || '').slice(0, 20)), porUnidad = cifraN(d.porUnidad);
  if (suelto) {
    if (suelto === unidad) throw new Error('Lo que trae adentro tiene que ser distinto de «' + unidad + '».');
    if (!(porUnidad >= 2)) throw new Error('Diga cuántas ' + plur(2, suelto) + ' trae cada ' + unidad + '.');
  } else porUnidad = '';
  var menor = suelto ? singular(String(d.menor || '').slice(0, 20)) : '', porSuelto = cifraN(d.porSuelto);
  if (menor) {
    if (menor === suelto || menor === unidad) throw new Error('La unidad más pequeña tiene que ser distinta de «' + suelto + '» y de «' + unidad + '».');
    if (!(porSuelto >= 2)) throw new Error('Diga cuántas ' + plur(2, menor) + ' trae cada ' + suelto + '.');
  } else porSuelto = '';
  var pesoCada = cifraN(d.pesoCada), pesoEn = String(d.pesoEn || '').trim().toLowerCase();
  if (d.pesa) {
    if (!(pesoCada > 0 && pesoCada <= 5000)) throw new Error('Diga cuánto pesa cada ' + (menor || suelto || unidad) + ' cerrado.');
    if (PESOS.indexOf(pesoEn) < 0) throw new Error('Escoja en qué se pesa lo abierto: ' + PESOS.join(', ') + '.');
  } else { pesoCada = ''; pesoEn = ''; }
  return { medida: 'unidad', unidad: unidad, suelto: suelto, porUnidad: porUnidad, menor: menor, porSuelto: porSuelto, pesoCada: pesoCada, pesoEn: pesoEn };
}
function agregarProducto(cred, p) {
  var yo = quien(cred);
  if (!puedeProductos(yo)) throw new Error('Los productos nuevos los agrega el director operativo o la bodega.');
  p = p || {};
  var nombre = String(p.nombre || '').trim().toUpperCase().slice(0, 80);
  if (nombre.length < 3) throw new Error('Escriba el nombre del producto.');
  var suc = p.todas ? { centro: true, almendras: true, parque: true } : (p.suc || {});
  if (!suc.centro && !suc.almendras && !suc.parque) throw new Error('Marque para qué sucursal es (o «Todas»).');
  var m = medidaValida(p);
  var area = AREAS_BODEGA.indexOf(p.area) >= 0 ? p.area : 'Cocina';
  var grupo = String(p.grupo || '').trim() ? nombreCategoria(p.grupo) : 'Otros';
  filasBodega().forEach(function (x) { if (x.grupo.toLowerCase() === grupo.toLowerCase()) grupo = x.grupo; });
  var categoria = CATEGORIAS.indexOf(p.categoria) >= 0 ? p.categoria : 'Insumos de cocina';
  var lock = LockService.getScriptLock(); lock.waitLock(15000); _LEE = {}; var codigo;
  try {
    filasBodega().forEach(function (b) { if (b.nombre.toUpperCase() === nombre) throw new Error('Ya existe «' + b.nombre + '» en el catálogo (' + b.codigo + ').'); });
    codigo = codigoProducto(nombre);
    var almacen = almacenDe(grupo), fila = [codigo, nombre, area, grupo, suc.centro ? 'Sí' : 'No', suc.almendras ? 'Sí' : 'No', suc.parque ? 'Sí' : 'No',
      categoria, m.unidad, m.suelto, m.porUnidad, '', '', 1, '', 'Sí', new Date(), yo.nombre, almacen, estanteDe(almacen, categoria), '', '', '', '',
      m.suelto ? 'Sí' : 'No', '', m.medida, m.menor, m.porSuelto, m.pesoCada, m.pesoEn,
      String(p.descripcion || '').replace(/\s+/g, ' ').trim().slice(0, 200)];
    hojaBodega().appendRow(fila); _BOD = null;
  } finally { lock.releaseLock(); }
  var r = productosConfig(cred); r.ok = true;
  r.mensaje = nombre + ' se agregó (' + codigo + ') para ' + (p.todas ? 'todas las sucursales' : ['centro', 'almendras', 'parque'].filter(function (k) { return suc[k]; }).map(nombreUnidad).join(', ')) + ', contado por ' + m.unidad + '.';
  return r;
}
/* ── Categorías (el «grupo» de cada producto): se usan para ordenar el conteo y la lista.
 * Las renombran el director operativo, la bodega o el administrador; renombrar a una
 * categoría que ya existe junta las dos. */
function nombreCategoria(v) {
  var s = String(v || '').replace(/\s+/g, ' ').trim().slice(0, 40);
  if (s.length < 2) throw new Error('Escriba el nombre de la categoría.');
  return s.charAt(0).toUpperCase() + s.slice(1);
}
function renombrarCategoria(cred, de, a) {
  var yo = quien(cred);
  if (!puedeProductos(yo)) throw new Error('Las categorías las cambia el director operativo o la bodega.');
  de = String(de || '').trim(); a = nombreCategoria(a);
  var existentes = {}; filasBodega().forEach(function (b) { existentes[b.grupo.toLowerCase()] = b.grupo; });
  if (!existentes[de.toLowerCase()]) throw new Error('No se encontró la categoría «' + de + '».');
  if (existentes[a.toLowerCase()] && a.toLowerCase() !== de.toLowerCase()) a = existentes[a.toLowerCase()];
  var n = 0, lock = LockService.getScriptLock(); lock.waitLock(15000); _LEE = {};
  try {
    var h = hojaBodega(), ahora = new Date();
    filasBodega().forEach(function (b) {
      if (b.grupo.toLowerCase() !== de.toLowerCase()) return;
      h.getRange(b.fila, 4).setValue(a); h.getRange(b.fila, 17, 1, 2).setValues([[ahora, yo.nombre]]); n++;
    });
    _BOD = null;
  } finally { lock.releaseLock(); }
  var r = productosConfig(cred); r.ok = true;
  r.mensaje = (existentes[a.toLowerCase()] && a.toLowerCase() !== de.toLowerCase() ? '«' + de + '» se juntó con «' + a + '»' : '«' + de + '» ahora se llama «' + a + '»') +
    ': ' + n + (n === 1 ? ' producto.' : ' productos.');
  return r;
}
/** Cómo se cuenta un producto: por unidad o por peso. Vale desde el próximo conteo. */
function guardarMedida(cred, codigo, d) {
  var yo = quien(cred);
  if (!puedeProductos(yo) && yo.rol !== 'gerente') throw new Error('Su usuario no cambia los productos.');
  var b = null; filasBodega().forEach(function (x) { if (x.codigo === String(codigo)) b = x; });
  if (!b) throw new Error('No se encontró ese producto.');
  if (yo.rol === 'gerente' && !b.suc[yo.sucursal]) throw new Error('Ese producto no lo maneja su sucursal.');
  d = d || {};
  var m = medidaValida(d);
  var h = hojaBodega();
  h.getRange(b.fila, 9, 1, 3).setValues([[m.unidad, m.suelto, m.porUnidad]]);
  h.getRange(b.fila, 28, 1, 4).setValues([[m.menor, m.porSuelto, m.pesoCada, m.pesoEn]]);
  h.getRange(b.fila, 17, 1, 2).setValues([[new Date(), yo.nombre]]);
  h.getRange(b.fila, 25).setValue(m.suelto ? (b.mandaSuelto ? 'Sí' : 'No') : 'No');
  h.getRange(b.fila, 27).setValue(m.medida);
  _BOD = null;
  var r = productosConfig(cred); r.ok = true;
  var nb = null; filasBodega().forEach(function (x) { if (x.codigo === b.codigo) nb = x; });
  r.mensaje = b.nombre + ': ' + (nb ? nb.presentacion : '') + '. Aparece así en el próximo conteo.';
  return r;
}

/* ════════════ PRODUCTO DAÑADO O VENCIDO ════════════
 * El gerente anota lo que se dañó o se venció en su sucursal (sale de su inventario).
 * Al director operativo le llega el aviso: lo evalúa y decide si fue una pérdida
 * justificada o quién responde (el gerente u otra persona). Si alguien responde, el
 * gerente escribe la respuesta: qué pasó y qué se hará para que no se repita. */
var H_MER = ['No.', 'Registrado en', 'Fecha', 'Sucursal', 'Código', 'Producto', 'Cantidad', 'Unidad', 'Motivo', 'Qué pasó',
  'A cargo', 'Costo unitario', 'Costo', 'Registró', 'Foto', 'Estado', 'Evaluó', 'Evaluado en', 'Decisión', 'Responsable',
  'Comentario', 'Monto a cubrir', 'Respuesta', 'Respondió', 'Respondido en'];
var MER_MOT = ['Dañado', 'Vencido', 'Mal estado', 'Otro'];
var MER_EST = { EVAL: 'Por evaluar', RESP: 'Por responder', CERR: 'Cerrado', ANUL: 'Anulado' };
var MER_DEC = { JUST: 'Pérdida justificada', GER: 'Responde el gerente', OTRO: 'Responde otra persona', ERR: 'Registro equivocado' };
function hojaMerma(ss) {
  ss = ss || libro();
  var h = ss.getSheetByName('Producto dañado');
  if (!h) { h = hojaLimpia(ss, 'Producto dañado', H_MER); h.getRange('B:B').setNumberFormat('dd/mm/yyyy hh:mm'); h.getRange('C:C').setNumberFormat('@');
    h.getRange('R:R').setNumberFormat('dd/mm/yyyy hh:mm'); h.getRange('Y:Y').setNumberFormat('dd/mm/yyyy hh:mm');
    h.setColumnWidth(6, 240); h.setColumnWidth(10, 300); h.setColumnWidth(21, 260); h.setColumnWidth(23, 300); }
  return h;
}
function filasMerma() {
  return leeTodo(hojaMerma(), H_MER.length).map(function (r, i) {
    return { fila: i + 2, numero: String(r[0] || ''), en: fmtSello(r[1]), fecha: fmtDia(r[2]), sucursal: String(r[3] || ''),
      codigo: String(r[4] || ''), producto: String(r[5] || ''), cantidad: Number(r[6]) || 0, unidad: String(r[7] || ''),
      motivo: String(r[8] || ''), que: String(r[9] || ''), aCargo: String(r[10] || ''), costoU: Number(r[11]) || 0, costo: Number(r[12]) || 0,
      por: String(r[13] || ''), foto: String(r[14] || ''), estado: String(r[15] || '') || MER_EST.EVAL, evalPor: String(r[16] || ''),
      evalEn: fmtSello(r[17]), decision: String(r[18] || ''), responsable: String(r[19] || ''), comentario: String(r[20] || ''),
      monto: numONulo(r[21]), respuesta: String(r[22] || ''), respPor: String(r[23] || ''), respEn: fmtSello(r[24]) };
  }).filter(function (x) { return x.numero; });
}
function veMerma(yo) { return yo.rol === 'gerente' || yo.rol === 'operaciones' || yo.esAdmin; }
function pantallaMerma(cred) {
  var yo = quien(cred);
  if (!veMerma(yo)) throw new Error('Esto lo ven los gerentes y el director operativo.');
  var ger = yo.rol === 'gerente', suc = ger ? nombreUnidad(yo.sucursal) : '', hoy = hoyISO(), mes = hoy.slice(0, 7);
  var lista = filasMerma().filter(function (x) { return !ger || x.sucursal === suc; }).reverse();
  var delMes = lista.filter(function (x) { return x.fecha.slice(0, 7) === mes && x.estado !== MER_EST.ANUL; });
  var porSuc = {};
  delMes.forEach(function (x) { var s = porSuc[x.sucursal] = porSuc[x.sucursal] || { nombre: x.sucursal, n: 0, costo: 0 }; s.n++; s.costo = r2(s.costo + x.costo); });
  var out = { hoy: hoy, motivos: MER_MOT, decisiones: [MER_DEC.JUST, MER_DEC.GER, MER_DEC.OTRO, MER_DEC.ERR],
    puedeRegistrar: ger, puedeEvaluar: yo.rol === 'operaciones' || yo.esAdmin, esGerente: ger, sucursal: suc,
    lista: lista.slice(0, 120).map(function (x) {
      var o = {}; Object.keys(x).forEach(function (k) { if (k !== 'fila') o[k] = x[k]; });
      o.puedeResponder = ger && x.estado === MER_EST.RESP;
      return o; }),
    mes: { n: delMes.length, costo: r2(delMes.reduce(function (a, x) { return a + x.costo; }, 0)),
      porSucursal: Object.keys(porSuc).map(function (k) { return porSuc[k]; }),
      porMotivo: MER_MOT.map(function (m) { var xs = delMes.filter(function (x) { return x.motivo === m; });
        return { motivo: m, n: xs.length, costo: r2(xs.reduce(function (a, x) { return a + x.costo; }, 0)) }; }).filter(function (m) { return m.n; }) },
    pendientes: { evaluar: lista.filter(function (x) { return x.estado === MER_EST.EVAL; }).length,
      responder: lista.filter(function (x) { return x.estado === MER_EST.RESP; }).length } };
  if (ger) out.productos = productosInv().filter(function (b) { return maneja(b, yo.sucursal); }).map(function (b) {
    return { codigo: b.codigo, nombre: b.nombre, unidad: b.unidad, medida: b.medida, grupo: b.grupo }; });
  return out;
}
function registrarMerma(cred, d) {
  var yo = quien(cred);
  if (yo.rol !== 'gerente') throw new Error('El producto dañado o vencido lo anota el gerente de la sucursal.');
  d = d || {};
  var sid = yo.sucursal, u = unidadPorId(sid);
  var b = productosInv().filter(function (p) { return p.codigo === String(d.codigo || '') && maneja(p, sid); })[0];
  if (!b) throw new Error('Escoja el producto.');
  var q = Number(String(d.cantidad == null ? '' : d.cantidad).replace(',', '.'));
  if (!(isFinite(q) && q > 0)) throw new Error('Escriba cuánto se dañó o se venció.');
  if (q > 100000) throw new Error('La cantidad parece equivocada.');
  if (b.medida !== 'peso' && Math.abs(q - Math.round(q * 100) / 100) > 1e-9) throw new Error('Use hasta dos decimales.');
  q = r3(q);
  var motivo = MER_MOT.indexOf(d.motivo) >= 0 ? d.motivo : '';
  if (!motivo) throw new Error('Escoja si está dañado, vencido o en mal estado.');
  var que = String(d.que || '').trim().slice(0, 400);
  if (que.length < 10) throw new Error('Cuente qué pasó (al menos una frase).');
  var hoy = hoyISO(), fecha = esFecha(d.fecha) ? d.fecha : hoy;
  if (fecha > hoy) throw new Error('La fecha no puede ser futura.');
  if (diasEntre(fecha, hoy) > 7) throw new Error('Solo se puede anotar hasta 7 días atrás. Si fue antes, avise al director operativo.');
  var aCargo = String(d.aCargo || '').trim().slice(0, 60);
  var cu = costoBodega(b, indiceCompras()).costo, costo = r2(cu * q);
  var foto = d.foto && d.foto.datos ? guardaArchivo(d.foto, 'Producto dañado ' + fecha + ' · ' + b.nombre + ' foto', 'La foto').url : '';   // fuera del candado
  var numero, lock = LockService.getScriptLock(); lock.waitLock(15000); _LEE = {};
  try {
    numero = numeroSiguiente(filasMerma(), 'PD');
    var h = hojaMerma(), ahora = new Date();
    var fila = [numero, ahora, fecha, u.nombre, b.codigo, b.nombre, q, b.unidad, motivo, que, aCargo, cu, costo, yo.nombre, foto,
      MER_EST.EVAL, '', '', '', '', '', '', '', '', ''];
    var f = h.getLastRow() + 1;
    h.getRange(f, 1, 1, H_MER.length).setValues([fila]);
    h.getRange(f, 3).setNumberFormat('@').setValue(fecha);
    escribeKardex([[ahora, u.nombre, b.codigo, b.nombre, MOV.MERMA, -q, numero, yo.nombre, motivo + ': ' + que.slice(0, 120)]]);
  } finally { lock.releaseLock(); }
  var r = pantallaMerma(cred); r.ok = true;
  r.mensaje = numero + ' anotado: ' + num(q) + ' ' + plur(q, b.unidad) + ' de ' + b.nombre + (costo ? ' (' + dinero(costo) + ')' : '') +
    '. Salió de su inventario y le llegó al director operativo para que lo evalúe.';
  return r;
}
function buscaMerma(numero) {
  var x = filasMerma().filter(function (y) { return y.numero === String(numero); })[0];
  if (!x) throw new Error('No se encontró ' + numero + '.');
  return x;
}
/** El director operativo evalúa: pérdida justificada, responde el gerente, responde otra persona o registro equivocado. */
function evaluarMerma(cred, numero, d) {
  var yo = quien(cred);
  if (yo.rol !== 'operaciones' && !yo.esAdmin) throw new Error('Esto lo evalúa el director operativo.');
  d = d || {};
  var x, est, resp = '', lock = LockService.getScriptLock(); lock.waitLock(15000); _LEE = {};
  try {
    x = buscaMerma(numero);
    if (x.estado !== MER_EST.EVAL) throw new Error(numero + ' ya se evaluó.');
    var dec = '';
    Object.keys(MER_DEC).forEach(function (k) { if (MER_DEC[k] === d.decision) dec = d.decision; });
    if (!dec) throw new Error('Escoja qué decide.');
    var com = String(d.comentario || '').trim().slice(0, 300);
    if (com.length < 4) throw new Error('Escriba un comentario: por qué lo decide así.');
    var monto = '';
    if (dec === MER_DEC.GER || dec === MER_DEC.OTRO) {
      resp = dec === MER_DEC.GER ? x.por : String(d.responsable || '').trim().slice(0, 60);
      if (resp.length < 3) throw new Error('Escriba quién responde.');
      var m = String(d.monto == null ? '' : d.monto).trim();
      monto = m === '' ? x.costo : Number(m.replace(',', '.'));
      if (!(isFinite(monto) && monto >= 0)) throw new Error('El monto a cubrir no es válido.');
      monto = r2(monto);
    }
    est = dec === MER_DEC.JUST ? MER_EST.CERR : dec === MER_DEC.ERR ? MER_EST.ANUL : MER_EST.RESP;
    hojaMerma().getRange(x.fila, 16, 1, 7).setValues([[est, yo.nombre, new Date(), dec, resp, com, monto]]);
    if (est === MER_EST.ANUL)
      escribeKardex([[new Date(), x.sucursal, x.codigo, x.producto, MOV.MERMA, x.cantidad, numero, yo.nombre, 'Anulado: ' + com.slice(0, 120)]]);
  } finally { lock.releaseLock(); }
  var r = pantallaMerma(cred); r.ok = true;
  r.mensaje = numero + ': ' + (est === MER_EST.CERR ? 'pérdida justificada, quedó cerrado.' : est === MER_EST.ANUL ? 'anulado; el producto regresó al inventario de ' + x.sucursal + '.'
    : 'responde ' + resp + '. ' + x.por + ' ya lo sabe y debe responder.');
  return r;
}
/** El gerente responde por lo que se dañó: qué pasó y qué hará para que no se repita. */
function responderMerma(cred, numero, texto) {
  var yo = quien(cred);
  if (yo.rol !== 'gerente') throw new Error('Esto lo responde el gerente de la sucursal.');
  var t = String(texto || '').trim().slice(0, 500);
  if (t.length < 10) throw new Error('Escriba su respuesta: qué pasó y qué hará para que no se repita.');
  var lock = LockService.getScriptLock(); lock.waitLock(15000); _LEE = {};
  try {
    var x = buscaMerma(numero);
    if (x.sucursal !== nombreUnidad(yo.sucursal)) throw new Error('Eso es de otra sucursal.');
    if (x.estado !== MER_EST.RESP) throw new Error(numero + ' no está esperando respuesta.');
    hojaMerma().getRange(x.fila, 16).setValue(MER_EST.CERR);
    hojaMerma().getRange(x.fila, 23, 1, 3).setValues([[t, yo.nombre, new Date()]]);
  } finally { lock.releaseLock(); }
  var r = pantallaMerma(cred); r.ok = true; r.mensaje = 'Respuesta enviada. ' + numero + ' quedó cerrado; el director operativo la verá.';
  return r;
}
/** Avisos de la campanita. */
function avisosMerma(yo, out, desde) {
  if (!veMerma(yo)) return;
  var suc = yo.rol === 'gerente' ? nombreUnidad(yo.sucursal) : '';
  filasMerma().forEach(function (x) {
    var t = num(x.cantidad) + ' ' + plur(x.cantidad, x.unidad) + ' de ' + x.producto;
    if (yo.rol === 'operaciones') {
      if (x.estado === MER_EST.EVAL)
        out.push({ id: 'pd-' + x.numero, tipo: 'Producto ' + x.motivo.toLowerCase() + ' por evaluar', cuando: x.en,
          titulo: x.sucursal + ' · ' + t + (x.costo ? ' · ' + dinero(x.costo) : ''), texto: x.que + ' · anotó ' + x.por, abrir: { tab: 'merma' } });
      if (x.respPor && x.respEn.slice(0, 10) >= desde)
        out.push({ id: 'pd-r-' + x.numero, tipo: 'Respuesta sobre producto dañado', cuando: x.respEn, titulo: x.numero + ' · ' + x.sucursal + ' · ' + t,
          texto: x.respuesta + ' · ' + x.respPor, abrir: { tab: 'merma' } });
    }
    if (yo.rol === 'gerente' && x.sucursal === suc) {
      if (x.estado === MER_EST.RESP)
        out.push({ id: 'pd-e-' + x.numero, tipo: 'Producto dañado: debe responder', cuando: x.evalEn, titulo: x.numero + ' · ' + t,
          texto: 'Responde: ' + x.responsable + (x.monto ? ' · a cubrir ' + dinero(x.monto) : '') + ' · ' + x.comentario, abrir: { tab: 'merma' } });
      else if ((x.estado === MER_EST.CERR || x.estado === MER_EST.ANUL) && x.evalEn && !x.respPor && x.evalEn.slice(0, 10) >= desde)
        out.push({ id: 'pd-e-' + x.numero, tipo: 'Producto dañado evaluado', cuando: x.evalEn, titulo: x.numero + ' · ' + x.decision.toLowerCase(),
          texto: t + ' · ' + x.comentario + ' · ' + x.evalPor, abrir: { tab: 'merma' } });
    }
  });
}
/** Para el tablero del administrador. */
function resumenMerma() {
  var mes = hoyISO().slice(0, 7), xs = filasMerma();
  var m = xs.filter(function (x) { return x.fecha.slice(0, 7) === mes && x.estado !== MER_EST.ANUL; });
  return { n: m.length, costo: r2(m.reduce(function (a, x) { return a + x.costo; }, 0)),
    evaluar: xs.filter(function (x) { return x.estado === MER_EST.EVAL; }).length,
    responder: xs.filter(function (x) { return x.estado === MER_EST.RESP; }).length,
    cubrir: r2(m.filter(function (x) { return x.monto; }).reduce(function (a, x) { return a + x.monto; }, 0)) };
}

/* ════════════ FIRMA DE LOS TRASLADOS ════════════
 * Cada traslado de bodega (TR-…) lo firma en la app el gerente de cada sucursal que
 * recibió: revisa lo que llegó, firma con el dedo y puede dejar una nota si algo no
 * llegó bien. Si la bodega corrige el traslado después, hay que firmarlo otra vez.
 * Imprimir la hoja es opcional: ya sale con las firmas. */
var H_FIRMAS = ['Traslado', 'Sucursal', 'Firmó', 'Firmado en', 'Firma', 'Nota', 'Huella', 'Fecha del traslado'];
var DIAS_FIRMA = 21;
function hojaFirmas(ss) {
  ss = ss || libro();
  var h = ss.getSheetByName('Firmas de traslados');
  if (!h) { h = hojaLimpia(ss, 'Firmas de traslados', H_FIRMAS); h.getRange('D:D').setNumberFormat('dd/mm/yyyy hh:mm'); h.getRange('H:H').setNumberFormat('@');
    h.setColumnWidth(5, 120); h.setColumnWidth(6, 260); }
  return h;
}
function firmasGuardadas() {
  var out = {};
  leeTodo(hojaFirmas(), H_FIRMAS.length).forEach(function (r, i) {
    var n = String(r[0] || ''), s = String(r[1] || ''); if (!n || !s) return;
    out[n + '|' + s] = { fila: i + 2, por: String(r[2] || ''), en: fmtSello(r[3]), firma: String(r[4] || ''), nota: String(r[5] || ''), huella: String(r[6] || '') };
  });
  return out;
}
/** Resumen corto de lo que se mandó: si cambia, la firma ya no vale. */
function huellaLineas(lineas) {
  var txt = lineas.map(function (l) { return String(l.producto || l.nombre).toLowerCase() + '|' + String(l.medida || l.unidad) + '|' + r3(l.cantidad); }).sort().join(';');
  var h = 5381; for (var i = 0; i < txt.length; i++) h = ((h * 33) ^ txt.charCodeAt(i)) >>> 0;
  return lineas.length + '-' + h.toString(36);
}
/** Los traslados recientes de una sucursal (o de todas), con sus productos. */
function trasladosRecientes(sucNombres, dias) {
  var ss = libro(), g = ss.getSheetByName('Gastos'), porNum = {}, ids = {};
  if (!g || g.getLastRow() < 2) return [];
  var limite = masDias(hoyISO(), -(dias || DIAS_FIRMA));
  var desde = Math.max(2, g.getLastRow() - 6000), ancho = Math.min(H_GASTOS.length, g.getLastColumn());
  g.getRange(desde, 1, g.getLastRow() - desde + 1, ancho).getValues().forEach(function (r) {
    var id = String(r[0] || ''), n = String(r[15] || '').trim(), s = String(r[4] || '');
    if (!/^T\d/.test(id) || /-B$/.test(id) || !/^TR-/.test(n) || sucNombres.indexOf(s) < 0) return;
    if (String(r[10]).toLowerCase() === 'sí') return;
    var f = fmtDia(r[6]); if (f < limite) return;
    var k = n + '|' + s;
    if (!porNum[k]) porNum[k] = { numero: n, sucursal: s, fecha: f, sello: fmtSello(r[1]), quien: String(r[2] || ''), lineas: [] };
    ids[id] = k;
  });
  if (Object.keys(ids).length) leeTodo(ss.getSheetByName('Gastos detalle'), H_DETALLE.length).forEach(function (r) {
    var k = ids[String(r[0])]; if (!k) return;
    if (String(r[13]).toLowerCase() !== 'no') return;
    porNum[k].lineas.push({ producto: String(r[8] || ''), cantidad: Number(r[9]) || 0, medida: String(r[10] || ''), monto: Number(r[12]) || 0 });
  });
  return Object.keys(porNum).map(function (k) { return porNum[k]; }).filter(function (t) { return t.lineas.length; })
    .sort(function (a, b) { return a.sello < b.sello ? 1 : -1; });
}
function estadoFirma(t, fs) {
  var f = fs[t.numero + '|' + t.sucursal], hu = huellaLineas(t.lineas);
  if (!f) return { estado: 'Por firmar' };
  return { estado: f.huella === hu ? 'Firmado' : 'Cambió: firmar otra vez', por: f.por, en: f.en, nota: f.nota, firma: f.firma };
}
function veFirmas(yo) { return yo.rol === 'gerente' || yo.rol === 'operaciones' || yo.rol === 'bodega' || yo.esAdmin; }
function pantallaFirmas(cred) {
  var yo = quien(cred);
  if (!veFirmas(yo)) throw new Error('Esto lo ven los gerentes, la bodega y el director operativo.');
  var sucs = yo.rol === 'gerente' ? [nombreUnidad(yo.sucursal)] : sucursalesDestino().map(function (u) { return u.nombre; });
  var fs = firmasGuardadas(), verPrecio = veCostosTraslado(yo);
  var lista = trasladosRecientes(sucs, DIAS_FIRMA).map(function (t) {
    var e = estadoFirma(t, fs);
    t.lineas.sort(function (a, b) { return a.producto.localeCompare(b.producto, 'es'); });
    var o = { numero: t.numero, sucursal: t.sucursal, fecha: t.fecha, sello: t.sello, quien: t.quien, lineas: t.lineas,
      estado: e.estado, por: e.por || '', en: e.en || '', nota: e.nota || '', firma: e.firma || '' };
    if (verPrecio) o.total = r2(t.lineas.reduce(function (a, l) { return a + (l.monto || 0); }, 0));
    else o.lineas = t.lineas.map(function (l) { return { producto: l.producto, cantidad: l.cantidad, medida: l.medida }; });
    return o;
  });
  return { hoy: hoyISO(), verPrecios: verPrecio, dias: DIAS_FIRMA, puedeFirmar: yo.rol === 'gerente', sucursal: yo.rol === 'gerente' ? sucs[0] : '',
    lista: lista, porFirmar: lista.filter(function (x) { return x.estado !== 'Firmado'; }).length };
}
function firmarTraslado(cred, numero, firma, nota) {
  var yo = quien(cred);
  if (yo.rol !== 'gerente') throw new Error('El traslado lo firma el gerente de la sucursal que lo recibe.');
  var suc = nombreUnidad(yo.sucursal);
  firma = String(firma || '');
  if (!/^data:image\/(png|jpeg);base64,/.test(firma)) throw new Error('Firme en el recuadro con el dedo.');
  if (firma.length < 900) throw new Error('La firma quedó vacía. Firme en el recuadro.');
  if (firma.length > 48000) throw new Error('La firma quedó muy pesada. Toque «Borrar» y firme otra vez.');
  nota = String(nota || '').trim().slice(0, 300);
  var t = trasladosRecientes([suc], 60).filter(function (x) { return x.numero === String(numero); })[0];
  if (!t) throw new Error('No se encontró el traslado ' + numero + ' para ' + suc + '.');
  var lock = LockService.getScriptLock(); lock.waitLock(15000); _LEE = {};
  try {
    var fs = firmasGuardadas(), f = fs[t.numero + '|' + suc], hu = huellaLineas(t.lineas), h = hojaFirmas();
    if (f && f.huella === hu) throw new Error(numero + ' ya está firmado por ' + f.por + '.');
    var fila = [t.numero, suc, yo.nombre, new Date(), firma, nota, hu, t.fecha];
    var n = f ? f.fila : h.getLastRow() + 1;
    h.getRange(n, 1, 1, H_FIRMAS.length).setValues([fila]);
    h.getRange(n, 8).setNumberFormat('@').setValue(t.fecha);
  } finally { lock.releaseLock(); }
  var r = pantallaFirmas(cred); r.ok = true;
  r.mensaje = numero + ' firmado: recibió ' + suc + '.' + (nota ? ' La bodega verá su nota.' : '');
  return r;
}
/** Avisos: el gerente, lo que le falta firmar; la bodega, las notas; el director, lo que lleva más de un día sin firma. */
function avisosFirmas(yo, out, desde) {
  if (!veFirmas(yo) || yo.esAdmin) return;
  var sucs = yo.rol === 'gerente' ? [nombreUnidad(yo.sucursal)] : sucursalesDestino().map(function (u) { return u.nombre; });
  var fs = firmasGuardadas(), ayer = masDias(hoyISO(), -1);
  trasladosRecientes(sucs, DIAS_FIRMA).forEach(function (t) {
    var e = estadoFirma(t, fs), txt = t.lineas.length + (t.lineas.length === 1 ? ' producto' : ' productos') + ' · mandó ' + t.quien;
    if (yo.rol === 'gerente' && e.estado !== 'Firmado')
      out.push({ id: 'tf-' + t.numero + '-' + t.sucursal + '-' + huellaLineas(t.lineas), tipo: e.estado === 'Por firmar' ? 'Traslado por firmar' : 'Traslado corregido: firmar otra vez',
        cuando: t.sello, titulo: t.numero + ' · ' + dia(t.fecha).slice(0, 5), texto: txt, abrir: { tab: 'firmas' } });
    if (yo.rol === 'bodega' && e.estado === 'Firmado' && e.nota && e.en.slice(0, 10) >= desde)
      out.push({ id: 'tfn-' + t.numero + '-' + t.sucursal, tipo: 'Nota al firmar un traslado', cuando: e.en,
        titulo: t.numero + ' · ' + t.sucursal, texto: e.nota + ' · ' + e.por, abrir: { pag: 'tras' } });
    if (yo.rol === 'operaciones' && e.estado !== 'Firmado' && t.fecha <= ayer)
      out.push({ id: 'tfo-' + t.numero + '-' + t.sucursal, tipo: 'Traslado sin firmar', cuando: t.sello,
        titulo: t.sucursal + ' · ' + t.numero + ' del ' + dia(t.fecha).slice(0, 5), texto: 'El gerente todavía no firma de recibido.', abrir: { tab: 'firmas' } });
  });
}

/* ════════════ PEDIDOS DE LAS SUCURSALES → TRASLADO ════════════
 * Samuel ya no escribe el traslado a mano: los pedidos que mandan los gerentes llenan
 * solos las columnas de cada sucursal. El sistema muestra cuánto hay en bodega y, si no
 * alcanza para todos, reparte parejo entre las sucursales (sin dar a nadie más de lo que
 * pidió). Las cantidades solo se cambian con «Editar» en cada producto. Al «Aceptar y
 * preparar», los pedidos quedan aceptados y preparados y a cada gerente le llega el aviso. */
function esPesoMedida(m) { return PESOS.indexOf(singular(m)) >= 0; }
/** Reparte `disponible` entre los pedidos: parejo, sin pasar de lo pedido. */
function repartoParejo(pedidos, disponible, entero) {
  var ids = Object.keys(pedidos).filter(function (k) { return pedidos[k] > 0; }), out = {};
  Object.keys(pedidos).forEach(function (k) { out[k] = 0; });
  var queda = Math.max(0, disponible);
  var paso = entero ? 1 : 0.01;
  // se da de a poco a los que todavía no completan, hasta que se acabe (reparto parejo)
  var falta = function (k) { return r3(pedidos[k] - out[k]); };
  var guardia = 0;
  while (queda >= paso - 1e-9 && ids.some(function (k) { return falta(k) > 1e-9; }) && guardia < 100000) {
    var abiertos = ids.filter(function (k) { return falta(k) > 1e-9; });
    var cuota = entero ? Math.max(1, Math.floor(queda / abiertos.length)) : Math.max(paso, Math.floor(queda / abiertos.length * 100) / 100);
    abiertos.sort(function (a, b) { return out[a] - out[b] || falta(b) - falta(a); });
    for (var i = 0; i < abiertos.length && queda >= paso - 1e-9; i++) {
      var k = abiertos[i], d = Math.min(cuota, falta(k), queda);
      if (entero) d = Math.floor(d + 1e-9);
      if (d <= 0) continue;
      out[k] = r3(out[k] + d); queda = r3(queda - d);
    }
    guardia++;
  }
  return out;
}
function pantallaPedidosBodega(cred) {
  var yo = quien(cred);
  if (yo.rol !== 'bodega' && !yo.esAdmin) throw new Error('Esto es de la bodega.');
  var sucs = sucursalesDestino(), porNombre = {};
  sucs.forEach(function (s) { porNombre[s.nombre] = s.id; });
  var abiertas = filasSol().filter(function (s) { return s.estado === SOL.ENV && porNombre[s.sucursal]; });
  var catT = {}; productosTraslado().forEach(function (p) { catT[p.key] = p; });
  var cat = {}; productosInv().forEach(function (b) { cat[b.codigo] = b; });
  var estB = estadoInventario('bodega'), conInv = !!estB.ultimo;
  // lo ya aceptado y todavía sin recibir sigue en bodega pero ya tiene dueño
  var ya = {};
  var comp = {}; filasSol().forEach(function (s) { if (s.estado === SOL.ACE || s.estado === SOL.CAM) comp[s.numero] = true; });
  var nums = {}; abiertas.forEach(function (s) { nums[s.numero] = s; });
  var lineas = {}, orden = [];
  lineasSol().forEach(function (l) {
    if (comp[l.numero]) { var f0 = catT[l.clave] ? catT[l.clave].factor : 1; ya[l.codigo] = (ya[l.codigo] || 0) + (l.enviado != null ? l.enviado : l.pedido) * f0; return; }
    var s = nums[l.numero]; if (!s || !(l.pedido > 0)) return;
    var sid = porNombre[s.sucursal];
    if (!lineas[l.clave]) {
      var p = catT[l.clave] || {}, b = cat[l.codigo] || {};
      lineas[l.clave] = { clave: l.clave, codigo: l.codigo, nombre: l.nombre, medida: l.medida, estante: b.estante || '', grupo: b.grupo || '',
        factor: p.factor || 1, entero: !esPesoMedida(l.medida), pedido: {}, propuesta: {} };
      orden.push(l.clave);
    }
    lineas[l.clave].pedido[sid] = r3((lineas[l.clave].pedido[sid] || 0) + l.pedido);
  });
  // disponible en bodega, por producto (en su unidad principal); primero las cajas, luego lo suelto
  var queda = {};
  orden.sort(function (a, b) { var x = lineas[a], y = lineas[b];
    return (x.estante || '').localeCompare(y.estante || '', 'es') || x.nombre.localeCompare(y.nombre, 'es') || (y.factor - x.factor); });
  var out = orden.map(function (k) {
    var l = lineas[k], total = 0;
    Object.keys(l.pedido).forEach(function (s) { total = r3(total + l.pedido[s]); });
    var disp = null;
    if (conInv) {
      if (queda[l.codigo] == null) queda[l.codigo] = Math.max(0, r3((estB.existencia[l.codigo] || 0) - (ya[l.codigo] || 0)));
      disp = r3(queda[l.codigo] / l.factor);
      if (l.entero) disp = Math.floor(disp + 1e-9);
    }
    var alcanza = disp == null || disp >= total - 1e-9;
    l.propuesta = alcanza ? JSON.parse(JSON.stringify(l.pedido)) : repartoParejo(l.pedido, disp, l.entero);
    var usa = 0; Object.keys(l.propuesta).forEach(function (s) { usa = r3(usa + l.propuesta[s]); });
    if (disp != null) queda[l.codigo] = Math.max(0, r3(queda[l.codigo] - usa * l.factor));
    var b = cat[l.codigo] || {};
    return { clave: l.clave, codigo: l.codigo, nombre: l.nombre, medida: l.medida, estante: l.estante, grupo: l.grupo, entero: l.entero,
      pedido: l.pedido, propuesta: l.propuesta, total: total, disponible: disp,
      disponibleTxt: disp == null ? '' : num(disp) + ' ' + plur(disp, l.medida), alcanza: alcanza };
  });
  return { hoy: hoyISO(), conInventario: conInv, puede: yo.rol === 'bodega',
    sucursales: sucs.map(function (s) { return { id: s.id, nombre: s.nombre }; }),
    solicitudes: abiertas.map(function (s) { return { numero: s.numero, sucursal: s.sucursal, sid: porNombre[s.sucursal], enviadaPor: s.enviadaPor,
      enviadaEn: s.enviadaEn, entrega: s.entrega, nota: s.nota }; }),
    lineas: out, faltan: out.filter(function (x) { return !x.alcanza; }).length };
}
/** envios = { clave: { centro: 3, almendras: 2 } }. Acepta y deja preparados todos los pedidos abiertos. */
function aceptarYPreparar(cred, envios, opciones) {
  var yo = exigeSoloBodega(cred);
  envios = envios || {}; opciones = opciones || {};
  var sucs = sucursalesDestino(), porNombre = {};
  sucs.forEach(function (s) { porNombre[s.nombre] = s.id; });
  var catT = {}; productosTraslado().forEach(function (p) { catT[p.key] = p; });
  var hechas = [], lock = LockService.getScriptLock(); lock.waitLock(25000); _LEE = {};
  try {
    var abiertas = filasSol().filter(function (s) { return s.estado === SOL.ENV && porNombre[s.sucursal]; });
    if (!abiertas.length) throw new Error('No hay pedidos nuevos de las sucursales.');
    var h = hojaSolD(), ahora = new Date(), total = 0, porSol = {};
    abiertas.forEach(function (s) { porSol[s.numero] = { s: s, n: 0, lineas: lineasSol(s.numero) }; });
    var lee = function (v, nombre) {
      var n = Number(String(v == null ? '' : v).replace(',', '.'));
      if (String(v == null ? '' : v).trim() === '') return null;
      if (!isFinite(n) || n < 0 || n > 10000) throw new Error('La cantidad de «' + nombre + '» no es válida.');
      return r3(n);
    };
    Object.keys(porSol).forEach(function (num0) {
      var x = porSol[num0], sid = porNombre[x.s.sucursal], vistas = {};
      x.lineas.forEach(function (l) {
        vistas[l.clave] = true;
        var e = envios[l.clave] || {}, n = lee(e[sid], l.nombre);
        if (n == null) n = l.pedido;
        h.getRange(l.fila, 11).setValue(n); x.n += n; total += n;
        var falta = r3(l.pedido - n);
        if (falta > 0.0005 && !(opciones.no && opciones.no[l.clave])) h.getRange(l.fila, 17, 1, 2).setValues([[falta, String((opciones.llega || {})[l.clave] || '').trim().slice(0, 40)]]);
        else h.getRange(l.fila, 17, 1, 2).setValues([['', '']]);
      });
      // lo que la bodega manda sin que la sucursal lo haya pedido
      var nuevas = [];
      Object.keys(envios).forEach(function (k) {
        if (vistas[k]) return;
        var p = catT[k]; if (!p) return;
        var n = lee((envios[k] || {})[sid], p.nombre);
        if (!(n > 0)) return;
        nuevas.push([x.s.numero, x.s.sucursal, p.codigo, p.nombre, p.unidad, k, '', 0, 'lo agregó la bodega', 0, n, '', '', '', '', '', '', '', '', '', '']);
        x.n += n; total += n;
      });
      if (nuevas.length) h.getRange(h.getLastRow() + 1, 1, nuevas.length, H_SOLD.length).setValues(nuevas);
    });
    Object.keys(porSol).forEach(function (num0) {
      var x = porSol[num0];
      if (!(x.n > 0)) return;          // a esa sucursal no se le pudo mandar nada: queda pendiente (y sigue «Enviada» hasta que haya)
      hojaSol().getRange(x.s.fila, 7, 1, 3).setValues([[SOL.ACE, ahora, yo.nombre]]);
      hojaSol().getRange(x.s.fila, 17).setValue(ahora);
      hechas.push(x.s.numero + ' ' + x.s.sucursal);
    });
    if (!hechas.length) throw new Error('No va a mandar nada. Revise las cantidades.');
  } finally { lock.releaseLock(); }
  var r = pantallaPedidosBodega(cred); r.ok = true;
  r.mensaje = 'Preparado: ' + hechas.join(' · ') + '. Los gerentes ya recibieron el aviso.';
  return r;
}

/* ════════════ CALENDARIOS DE TRABAJO (quincenales) ════════════
 * Cada gerente arma el calendario de su sucursal: por persona, por día, mañana (AM) y tarde (PM).
 * Una quincena se divide en dos semanas (bloques). De ahí salen las horas de cada persona, el
 * semáforo de horas, el costo de mano de obra de cada día y la planilla proyectada de la quincena.
 * Solo el administrador y el director operativo ven el dinero; el gerente ve horas.
 * Cuando se paga la planilla, el administrador la confirma: se anota como gasto de «Nómina y
 * bonificaciones» de la sucursal y el calendario de esa quincena queda cerrado. */
var H_CEQ = ['ID', 'Sucursal', 'Nombre', 'Área', 'Pago por hora', 'Bono quincenal', 'Alta', 'Baja', 'Creó',
  'Rol de quien creó', 'Creado en', 'Actualizado por', 'Actualizado en'];
var H_CTU = ['Sucursal', 'Fecha', 'ID', 'Mañana', 'Tarde', 'Actualizado por', 'Actualizado en'];
var H_CCF = ['Sucursal', 'Clave', 'Valor', 'Actualizado por', 'Actualizado en'];
var H_CBO = ['Sucursal', 'Inicio', 'ID', 'Bono', 'Actualizado por', 'Actualizado en'];
var H_CHR = ['Sucursal', 'Inicio', 'ID', 'Horas reales', 'Actualizado por', 'Actualizado en'];
var H_CBR = ['Usuario', 'Sucursal', 'Inicio', 'Cambios', 'Actualizado en'];
var H_CPL = ['Sucursal', 'Inicio', 'Fin', 'Horas', 'Proyectado', 'Pagado', 'Fecha de pago', 'Confirmado por',
  'Confirmado en', 'ID del gasto', 'Estado', 'Nota', 'Detalle'];
var CAL_AREAS = ['Cocina', 'Barra'];
var CAL_MESES = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto', 'septiembre',
  'octubre', 'noviembre', 'diciembre'];
var CAL_D3 = ['Dom', 'Lun', 'Mar', 'Mié', 'Jue', 'Vie', 'Sáb'];
/* Como están hoy en la hoja de cálculo de la quincena. Se cambian en la pestaña Configuración. */
var CAL_DEF = { am_i: '08:00', am_f: '13:00', pm_i: '13:00', pm_f: '20:30', ap_i: '16:00', ap_f: '20:30', ap2_i: '15:00', ap2_f: '20:30',
  base: 13, bonoGeneral: 0, bajoDe: 33, bajoA: 38, optDe: 38.5, optA: 42.5 };

function calRed(n) { return Math.round((Number(n) || 0) * 100) / 100; }
function calTxt(v) { return v instanceof Date ? fmtDia(v) : String(v == null ? '' : v).trim(); }
function calHoja(nombre, enc, textoCols) {
  var existe = libro().getSheetByName(nombre), h = hojaCat(nombre, enc);
  if (!existe) (textoCols || []).forEach(function (c) { h.getRange(c + ':' + c).setNumberFormat('@'); });
  return h;
}

/* ── lectura de las hojas ── */
function calLeeEq() {
  return leeTodo(calHoja('Calendario equipo', H_CEQ, ['G', 'H']), H_CEQ.length).map(function (r, i) {
    return { fila: i + 2, id: calTxt(r[0]), sid: calTxt(r[1]), nombre: calTxt(r[2]), area: calTxt(r[3]),
      pago: calTxt(r[4]) === '' ? null : Number(r[4]), bono: Number(r[5]) || 0, alta: calTxt(r[6]), baja: calTxt(r[7]),
      por: calTxt(r[8]), rolPor: calTxt(r[9]), creado: fmtSello(r[10]) };
  }).filter(function (x) { return x.id && x.sid; });
}
function calLeeTu() {
  return leeTodo(calHoja('Calendario turnos', H_CTU, ['B']), H_CTU.length).map(function (r, i) {
    return { fila: i + 2, sid: calTxt(r[0]), f: calTxt(r[1]), id: calTxt(r[2]), am: calTxt(r[3]).toLowerCase(),
      pm: calTxt(r[4]).toLowerCase() };
  }).filter(function (x) { return x.sid && x.f && x.id; });
}
function calLeeCf() {
  return leeTodo(calHoja('Calendario config', H_CCF, ['C']), H_CCF.length).map(function (r, i) {
    return { fila: i + 2, sid: calTxt(r[0]), k: calTxt(r[1]), v: calTxt(r[2]) };
  }).filter(function (x) { return x.sid && x.k; });
}
function calLeeBo() {
  return leeTodo(calHoja('Calendario bonos', H_CBO, ['B']), H_CBO.length).map(function (r, i) {
    return { fila: i + 2, sid: calTxt(r[0]), ini: calTxt(r[1]), id: calTxt(r[2]), bono: calTxt(r[3]) === '' ? null : Number(r[3]) || 0 };
  }).filter(function (x) { return x.sid && x.ini && x.id && x.bono != null; });
}
function calLeeHr() {
  return leeTodo(calHoja('Calendario horas reales', H_CHR, ['B']), H_CHR.length).map(function (r, i) {
    return { fila: i + 2, sid: calTxt(r[0]), ini: calTxt(r[1]), id: calTxt(r[2]), horas: calTxt(r[3]) === '' ? null : Number(r[3]) };
  }).filter(function (x) { return x.sid && x.ini && x.id; });
}
function calLeePl() {
  return leeTodo(calHoja('Calendario planilla', H_CPL, ['B', 'C', 'G']), H_CPL.length).map(function (r, i) {
    return { fila: i + 2, sid: calTxt(r[0]), ini: calTxt(r[1]), fin: calTxt(r[2]), horas: Number(r[3]) || 0,
      proyectado: Number(r[4]) || 0, pagado: Number(r[5]) || 0, fechaPago: calTxt(r[6]), por: calTxt(r[7]),
      en: fmtSello(r[8]), gasto: calTxt(r[9]), estado: calTxt(r[10]), nota: calTxt(r[11]), detalle: calTxt(r[12]) };
  }).filter(function (x) { return x.sid && x.ini; });
}
function calDatos() { return { eq: calLeeEq(), tu: calLeeTu(), cf: calLeeCf(), bo: calLeeBo(), hr: calLeeHr(), pl: calLeePl() }; }

/* ── configuración de una sucursal ── */
function calCfg(sid, filas) {
  var c = { ger: {} }, k;
  for (k in CAL_DEF) c[k] = CAL_DEF[k];
  (filas || []).forEach(function (x) {
    if (x.sid === '*') {                                       // lo que vale para todas las sucursales
      if (x.k === 'bono_general') { var bg = Number(x.v); if (x.v !== '' && isFinite(bg) && bg >= 0) c.bonoGeneral = bg; }
      return;
    }
    if (x.sid !== sid) return;
    if (/^(am|pm|ap|ap2)_[if]$/.test(x.k)) { if (horaValida(x.v)) c[x.k] = horaNorm(x.v); return; }
    if (x.k === 'base' || /^(bajo|opt)(De|A)$/.test(x.k)) { var n = Number(x.v); if (x.v !== '' && isFinite(n)) c[x.k] = n; return; }
    var m = /^ger_(salario|bono):(.+)$/.exec(x.k);
    if (m) { c.ger[m[2]] = c.ger[m[2]] || { salario: 0, bono: 0 }; c.ger[m[2]][m[1]] = Number(x.v) || 0; }
  });
  return c;
}
function calMin(t) { var p = String(t).split(':'); return (+p[0]) * 60 + (+p[1] || 0); }
function calHoras(c) {
  var h = function (i, f) { var d = calMin(f) - calMin(i); if (d <= 0) d += 1440; return calRed(d / 60); };
  return { am: h(c.am_i, c.am_f), pm: h(c.pm_i, c.pm_f), ap: h(c.ap_i, c.ap_f), ap2: h(c.ap2_i, c.ap2_f) };
}

/* ── la quincena: del 1 al 15 y del 16 al último día del mes; dos semanas ── */
function calQuincena(fecha) {
  fecha = esFecha(fecha) ? fecha : hoyISO();
  var y = +fecha.slice(0, 4), m = +fecha.slice(5, 7), d = +fecha.slice(8, 10), mm = ('0' + m).slice(-2);
  var ini, fin;
  if (d <= 15) { ini = y + '-' + mm + '-01'; fin = y + '-' + mm + '-15'; }
  else { ini = y + '-' + mm + '-16'; fin = y + '-' + mm + '-' + ('0' + new Date(y, m, 0).getDate()).slice(-2); }
  var dias = [], f = ini;
  while (f <= fin) { dias.push(f); f = masDias(f, 1); }
  var n1 = Math.floor(dias.length / 2);
  return { ini: ini, fin: fin, dias: dias, n: dias.length, n1: n1, n2: dias.length - n1,
    nombre: 'Quincena del ' + (+ini.slice(8)) + ' al ' + (+fin.slice(8)) + ' de ' + CAL_MESES[m - 1] + ' de ' + y,
    corto: ('0' + (+ini.slice(8))).slice(-2) + '–' + ('0' + (+fin.slice(8))).slice(-2) + ' ' + CAL_MESES[m - 1] + ' ' + y,
    ant: masDias(ini, -1), sig: masDias(fin, 1) };
}
/** El color de las horas, como en la hoja de cálculo: bajo (naranja), óptimo (verde) y alto (rojo).
 *  Los rangos son por semana; el de la quincena es el doble (k = 2). */
function calClase(h, cfg, k) {
  if (h > cfg.optA * k + 1e-9) return 'alto';
  if (h >= cfg.optDe * k - 1e-9) return 'optimo';
  if (h >= cfg.bajoDe * k - 1e-9) return 'bajo';
  return '';
}

/* ── quién puede qué ── */
function calAcceso(yo) {
  var todas = ['centro', 'almendras', 'parque'];
  if (yo.rol === 'gerente') return { sucs: [yo.sucursal], turnos: true, equipo: true, pagos: false, dinero: false, confirma: false };
  if (yo.esAdmin) return { sucs: todas, turnos: true, equipo: true, pagos: true, dinero: true, confirma: true };
  if (yo.rol === 'operaciones') return { sucs: todas, turnos: true, equipo: true, pagos: true, dinero: true, confirma: false };
  if (yo.rol === 'finanzas') return { sucs: todas, turnos: true, equipo: true, pagos: true, dinero: true, confirma: false, soloVer: true };   // ve todo (turnos, equipo, pagos, planillas) sin cambiar nada
  if (yo.rol === 'dueno') return { sucs: todas, turnos: false, equipo: false, pagos: false, dinero: true, confirma: false };
  throw new Error('Los calendarios los ven los gerentes y la administración.');
}
function calSucursal(ac, sid) {
  if (!sid || ac.sucs.indexOf(sid) < 0) sid = ac.sucs[0];
  return sid;
}

/* ── el cálculo de una quincena de una sucursal ── */
function calModelo(D, sid, fecha) {
  var q = calQuincena(fecha), cfg = calCfg(sid, D.cf), hr = calHoras(cfg);
  var mapa = {}, ov = {}, rh = {};
  (D.hr || []).forEach(function (x) { if (x.sid === sid && x.ini === q.ini && x.horas != null) rh[x.id] = x.horas; });
  D.tu.forEach(function (r) { if (r.sid === sid && r.f >= q.ini && r.f <= q.fin) mapa[r.id + '|' + r.f] = r; });
  D.bo.forEach(function (b) { if (b.sid === sid && b.ini === q.ini) ov[b.id] = b.bono; });
  var gente = D.eq.filter(function (p) {
    return p.sid === sid && (!p.baja || p.baja >= q.ini) && (!p.alta || p.alta <= q.fin);
  }).sort(function (a, b) {
    return CAL_AREAS.indexOf(a.area) - CAL_AREAS.indexOf(b.area) || (a.id < b.id ? -1 : a.id > b.id ? 1 : 0);
  });
  var suma = function (dias, a, b) { var t = 0; for (var k = a; k < b; k++) t += dias[k].h; return calRed(t); };
  var sinDesc = function (dias, a, b) { for (var k = a; k < b; k++) if (dias[k].h === 0) return false; return true; };
  var personas = gente.map(function (p) {
    var dias = q.dias.map(function (f) {
      var r = mapa[p.id + '|' + f], am = r && r.am === 'x' ? 'x' : '', pm = r && (r.pm === 'x' || r.pm === 'xa' || r.pm === 'xb') ? r.pm : '';
      return { f: f, am: am, pm: pm, h: calRed((am ? hr.am : 0) + (pm === 'x' ? hr.pm : pm === 'xa' ? hr.ap : pm === 'xb' ? hr.ap2 : 0)) };
    });
    var hb1 = suma(dias, 0, q.n1), hb2 = suma(dias, q.n1, q.n), hq = calRed(hb1 + hb2);
    var tarifa = p.pago != null ? p.pago : cfg.base;
    // El bono es parejo para todos (el general); solo cambia si alguien tiene uno propio o se ajustó en esta quincena.
    var propio = p.bono > 0, bonoBase = propio ? p.bono : cfg.bonoGeneral;
    var ajustado = ov.hasOwnProperty(p.id) && ov[p.id] !== bonoBase, bono = ov.hasOwnProperty(p.id) ? ov[p.id] : bonoBase;
    // Las horas del pago son las reales cuando ya se anotaron al final de la quincena; si no, las del calendario.
    var hreal = rh.hasOwnProperty(p.id) ? rh[p.id] : null, hp = hreal != null ? hreal : hq;
    var bonoProy = hq > 0 ? bono : 0;
    if (!(hp > 0)) bono = 0;
    return { id: p.id, nombre: p.nombre, area: p.area, dias: dias, hb1: hb1, hb2: hb2, hq: hq,
      c1: calClase(hb1, cfg, 1), c2: calClase(hb2, cfg, 1), cq: calClase(hq, cfg, 2),
      sd1: sinDesc(dias, 0, q.n1), sd2: sinDesc(dias, q.n1, q.n),
      tarifa: tarifa, tarifaPropia: p.pago != null, bono: bono, bonoBase: bonoBase, bonoAjustado: ajustado, bonoPropio: propio,
      bonoOrigen: ajustado ? 'ajustado' : propio ? 'propio' : 'general',
      hreal: hreal, hp: hp, pagoProy: calRed(hq * tarifa + bonoProy),
      pagoHoras: calRed(hp * tarifa), pago: calRed(hp * tarifa + bono) };
  });
  var dias = q.dias.map(function (f, i) {
    var t = { f: f, dow: diaSemana(f), am: 0, pm: 0, gente: 0, horas: 0, costo: 0 };
    personas.forEach(function (p) {
      var d = p.dias[i]; if (d.am) t.am++; if (d.pm) t.pm++;
      if (d.h > 0) t.gente++; t.horas += d.h; t.costo += d.h * p.tarifa;
    });
    t.horas = calRed(t.horas); t.costo = calRed(t.costo); return t;
  });
  var areas = CAL_AREAS.map(function (a) {
    var fs = personas.filter(function (p) { return p.area === a; }), t = { area: a, personas: fs.length, hb1: 0, hb2: 0, hq: 0, hp: 0, bonos: 0, pago: 0, proy: 0 };
    fs.forEach(function (p) { t.hb1 += p.hb1; t.hb2 += p.hb2; t.hq += p.hq; t.hp += p.hp; t.bonos += p.bono; t.pago += p.pago; t.proy += p.pagoProy; });
    t.hb1 = calRed(t.hb1); t.hb2 = calRed(t.hb2); t.hq = calRed(t.hq); t.hp = calRed(t.hp); t.bonos = calRed(t.bonos); t.pago = calRed(t.pago); t.proy = calRed(t.proy);
    return t;
  });
  var gerencia = usuariosCache().filter(function (u) { return u.activo && u.rol === 'gerente' && u.sucursal === sid; }).map(function (u) {
    var g = cfg.ger[u.nombre] || { salario: 0, bono: 0 };
    return { nombre: u.nombre, puesto: 'Gerente de sucursal ' + nombreUnidad(sid), salario: g.salario, bono: g.bono, pago: calRed(g.salario + g.bono) };
  });
  var tot = { horas: 0, horasCal: 0, personal: 0, bonos: 0, gerencia: 0, total: 0, proy: 0, sinReales: 0, conReales: 0 };
  areas.forEach(function (a) { tot.horas += a.hp; tot.horasCal += a.hq; tot.personal += a.pago; tot.bonos += a.bonos; tot.proy += a.proy; });
  personas.forEach(function (p) { if (p.hq > 0 || p.hreal != null) { if (p.hreal != null) tot.conReales++; else tot.sinReales++; } });
  gerencia.forEach(function (g) { tot.gerencia += g.pago; });
  tot.horas = calRed(tot.horas); tot.horasCal = calRed(tot.horasCal); tot.personal = calRed(tot.personal); tot.bonos = calRed(tot.bonos);
  tot.gerencia = calRed(tot.gerencia); tot.total = calRed(tot.personal + tot.gerencia); tot.proy = calRed(tot.proy + tot.gerencia);
  return { q: q, cfg: cfg, hr: hr, personas: personas, dias: dias, areas: areas, gerencia: gerencia, tot: tot };
}

function calPlanillaDe(D, sid, ini) {
  var out = null;
  D.pl.forEach(function (x) { if (x.sid === sid && x.ini === ini && x.estado === 'Pagada') out = x; });
  return out;
}

/** La venta de cada día de la quincena y la venta típica de cada día de la semana
 *  (promedio de las últimas 8 semanas): sirve para poner más gente en los días fuertes. */
function calVentas(sid, q) {
  var hoy = hoyISO(), base = masDias(hoy, -56), desde = q.ini < base ? q.ini : base, nom = nombreUnidad(sid), porDia = {};
  try {
    ventasRango(desde, hoy).forEach(function (v) { if (v.u === nom) porDia[v.f] = (porDia[v.f] || 0) + v.m; });
  } catch (e) {}
  var s = [0, 0, 0, 0, 0, 0, 0], c = [0, 0, 0, 0, 0, 0, 0];
  Object.keys(porDia).forEach(function (f) {
    if (f >= base && f < hoy && porDia[f] > 0) { var d = diaSemana(f); s[d] += porDia[f]; c[d]++; }
  });
  return { porDia: porDia, tipica: s.map(function (t, d) { return c[d] ? Math.round(t / c[d]) : 0; }) };
}

/** Todo lo que pinta la ventana de Calendarios. */
function calPantalla(yo, ac, sid, fecha, D) {
  sid = calSucursal(ac, sid); D = D || calDatos();
  var m = calModelo(D, sid, fecha), q = m.q, pl = calPlanillaDe(D, sid, q.ini), ve = calVentas(sid, q), hoy = hoyISO();
  var out = {
    hoy: hoy, sid: sid, sucursal: nombreUnidad(sid),
    sucursales: ac.sucs.map(function (s) { return { id: s, nombre: nombreUnidad(s) }; }),
    q: { ini: q.ini, fin: q.fin, n: q.n, n1: q.n1, n2: q.n2, nombre: q.nombre, corto: q.corto, ant: q.ant, sig: q.sig, dias: q.dias },
    acceso: { turnos: ac.turnos, equipo: ac.equipo, pagos: ac.pagos, dinero: ac.dinero, confirma: ac.confirma, soloVer: !!ac.soloVer },
    horas: m.hr, turnos: { am: [m.cfg.am_i, m.cfg.am_f], pm: [m.cfg.pm_i, m.cfg.pm_f], ap: [m.cfg.ap_i, m.cfg.ap_f], ap2: [m.cfg.ap2_i, m.cfg.ap2_f] },
    rangos: { bajoDe: m.cfg.bajoDe, bajoA: m.cfg.bajoA, optDe: m.cfg.optDe, optA: m.cfg.optA },
    cerrada: !!pl, venta: {}, tipica: ve.tipica,
    personas: m.personas.map(function (p) {
      var o = { id: p.id, nombre: p.nombre, area: p.area, dias: p.dias.map(function (d) { return { f: d.f, am: d.am, pm: d.pm }; }),
        hb1: p.hb1, hb2: p.hb2, hq: p.hq, c1: p.c1, c2: p.c2, cq: p.cq, sd1: p.sd1, sd2: p.sd2 };
      if (ac.dinero) { o.tarifa = p.tarifa; o.bono = p.bono; o.hreal = p.hreal; }
      return o;
    }),
    equipo: D.eq.filter(function (p) { return p.sid === sid && !p.baja; }).map(function (p) {
      var o = { id: p.id, nombre: p.nombre, area: p.area };
      if (ac.pagos) { o.pago = p.pago; o.bono = p.bono; o.bonoEf = p.bono > 0 ? p.bono : m.cfg.bonoGeneral; }
      return o;
    })
  };
  q.dias.forEach(function (f) { if (f <= hoy && ve.porDia[f] > 0) out.venta[f] = Math.round(ve.porDia[f]); });
  if (ac.dinero) {
    out.costoDias = m.dias.map(function (d) { return d.costo; });
    out.pagos = (pl && pl.detalle && calJson(pl.detalle, null)) || calPagos(m);
    out.sinReales = m.tot.sinReales; out.conReales = m.tot.conReales; out.finalizada = hoy >= q.fin;
    out.planilla = pl ? { pagado: pl.pagado, proyectado: pl.proyectado, fechaPago: pl.fechaPago, por: pl.por, en: pl.en, nota: pl.nota } : null;
    out.base = m.cfg.base; out.bonoGeneral = m.cfg.bonoGeneral;
    if (ac.pagos) {
      out.config = { bonoGeneral: m.cfg.bonoGeneral, am: [m.cfg.am_i, m.cfg.am_f], pm: [m.cfg.pm_i, m.cfg.pm_f], ap: [m.cfg.ap_i, m.cfg.ap_f], ap2: [m.cfg.ap2_i, m.cfg.ap2_f], base: m.cfg.base,
        bajoDe: m.cfg.bajoDe, bajoA: m.cfg.bajoA, optDe: m.cfg.optDe, optA: m.cfg.optA,
        gerentes: m.gerencia.map(function (g) { return { nombre: g.nombre, salario: g.salario, bono: g.bono }; }) };
    }
  }
  if (ac.turnos && !pl) out.borrador = calBorradorDe(yo, sid, q.ini);
  out.avisos = calAvisosModelo(m);
  return out;
}
function calJson(t, def) { try { return JSON.parse(t); } catch (e) { return def; } }
function calPagos(m) {
  return { areas: m.areas.map(function (a) {
      return { area: a.area, personas: a.personas, hb1: a.hb1, hb2: a.hb2, hq: a.hq, hp: a.hp, bonos: a.bonos, pago: a.pago,
        filas: m.personas.filter(function (p) { return p.area === a.area; }).map(function (p) {
          return { id: p.id, nombre: p.nombre, hb1: p.hb1, hb2: p.hb2, hq: p.hq, hreal: p.hreal, hp: p.hp, tarifa: p.tarifa, bono: p.bono, bonoBase: p.bonoBase,
            ajustado: p.bonoAjustado, origen: p.bonoOrigen, pago: p.pago };
        }) };
    }),
    gerencia: m.gerencia, tot: m.tot };
}
/** Avisos del calendario: quien no descansa ni un día, y quien se pasa de horas. */
function calAvisosModelo(m) {
  var out = [];
  m.personas.forEach(function (p) {
    if (p.hq <= 0) return;
    if (p.sd1) out.push({ id: p.id, tipo: 'descanso', texto: p.nombre + ' no tiene día de descanso en la semana 1.' });
    if (p.sd2) out.push({ id: p.id, tipo: 'descanso', texto: p.nombre + ' no tiene día de descanso en la semana 2.' });
    if (p.c1 === 'alto') out.push({ id: p.id, tipo: 'horas', texto: p.nombre + ' pasa del rango óptimo en la semana 1 (' + p.hb1 + ' h).' });
    if (p.c2 === 'alto') out.push({ id: p.id, tipo: 'horas', texto: p.nombre + ' pasa del rango óptimo en la semana 2 (' + p.hb2 + ' h).' });
  });
  return out;
}

/* ── pantalla ── */
function pantallaCalendarios(cred, sid, fecha) {
  var yo = quien(cred), ac = calAcceso(yo);
  return calPantalla(yo, ac, sid, fecha);
}
function calRespuesta(yo, ac, sid, fecha, mensaje) {
  var r = calPantalla(yo, ac, sid, fecha); r.ok = true; r.mensaje = mensaje; return r;
}

/* ── turnos: quién trabaja cada día ── */
function guardarTurnos(cred, sid, fecha, cambios) {
  var yo = quien(cred), ac = calAcceso(yo); soloVerCal(ac);
  if (!ac.turnos) throw new Error('Su usuario solo consulta los calendarios.');
  if (ac.sucs.indexOf(sid) < 0) throw new Error('Ese calendario no es de su sucursal.');
  if (!Array.isArray(cambios) || !cambios.length) throw new Error('No hay cambios que guardar.');
  if (cambios.length > 400) throw new Error('Son demasiados cambios de una vez. Guarde por partes.');
  var lock = LockService.getScriptLock(); lock.waitLock(15000); _LEE = {};
  try {
    var D = calDatos(), q = calQuincena(fecha);
    if (calPlanillaDe(D, sid, q.ini)) throw new Error('La planilla de esta quincena ya está pagada: el calendario quedó cerrado.');
    var ids = {};
    D.eq.forEach(function (p) { if (p.sid === sid && (!p.baja || p.baja >= q.ini)) ids[p.id] = true; });
    var idx = {}; D.tu.forEach(function (r) { if (r.sid === sid) idx[r.id + '|' + r.f] = r; });
    var h = calHoja('Calendario turnos', H_CTU, ['B']), ahora = new Date(), nuevos = [], n = 0, ult = {};
    cambios.forEach(function (c) {
      c = c || {};
      var id = String(c.id || ''), f = String(c.f || '');
      if (!ids[id]) throw new Error('Una de las personas ya no está en el equipo. Vuelva a abrir el calendario.');
      if (q.dias.indexOf(f) < 0) throw new Error('Una fecha no es de esta quincena.');
      var am = (c.am === true || c.am === 'x') ? 'x' : '';
      var pm = c.pm === 'x' || c.pm === 'xa' || c.pm === 'xb' ? c.pm : '';
      ult[id + '|' + f] = { id: id, f: f, am: am, pm: pm };
    });
    Object.keys(ult).forEach(function (k) {
      var u = ult[k], r = idx[k];
      if (r) {
        if (r.am === u.am && r.pm === u.pm) return;
        h.getRange(r.fila, 4, 1, 4).setValues([[u.am, u.pm, yo.nombre, ahora]]); n++;
      } else {
        if (!u.am && !u.pm) return;
        nuevos.push([sid, u.f, u.id, u.am, u.pm, yo.nombre, ahora]); n++;
      }
    });
    if (nuevos.length) {
      var d0 = h.getLastRow() + 1;
      h.getRange(d0, 2, nuevos.length, 1).setNumberFormat('@');
      h.getRange(d0, 1, nuevos.length, H_CTU.length).setValues(nuevos);
    }
    calBorradorPon(yo, sid, q.ini, []);
    return calRespuesta(yo, ac, sid, fecha, n ? 'Calendario guardado (' + n + (n === 1 ? ' cambio' : ' cambios') + ').' : 'No había nada nuevo que guardar.');
  } finally { lock.releaseLock(); }
}

/* ── el equipo de la sucursal: registrar, editar y eliminar ── */
function calNumero(v, campo, max) {
  if (v === '' || v == null) return null;
  var n = Number(String(v).replace(',', '.'));
  if (!isFinite(n) || n < 0 || n > max) throw new Error(campo + ' no es válido.');
  return calRed(n);
}
function guardarPersonaCal(cred, sid, p, fecha) {
  var yo = quien(cred), ac = calAcceso(yo); soloVerCal(ac);
  if (!ac.equipo) throw new Error('Su usuario solo consulta los calendarios.');
  if (ac.sucs.indexOf(sid) < 0) throw new Error('Esa no es su sucursal.');
  p = p || {};
  var nombre = String(p.nombre || '').replace(/\s+/g, ' ').trim();
  if (nombre.length < 2) throw new Error('Escriba el nombre de la persona.');
  if (nombre.length > 40) throw new Error('El nombre es muy largo.');
  var area = String(p.area || '');
  if (CAL_AREAS.indexOf(area) < 0) throw new Error('Escoja si trabaja en Cocina o en Barra.');
  var lock = LockService.getScriptLock(); lock.waitLock(15000); _LEE = {};
  try {
    var eq = calLeeEq(), h = calHoja('Calendario equipo', H_CEQ, ['G', 'H']), ahora = new Date(), id = String(p.id || ''), previo = null;
    eq.forEach(function (x) { if (x.id === id) previo = x; });
    if (id && (!previo || previo.sid !== sid || previo.baja)) throw new Error('Esa persona ya no está en el equipo.');
    eq.forEach(function (x) {
      if (x.sid === sid && !x.baja && x.id !== id && x.nombre.toLowerCase() === nombre.toLowerCase())
        throw new Error('Ya hay alguien que se llama «' + nombre + '» en esta sucursal. Agregue el apellido.');
    });
    var pago = null, bono = null;
    if (ac.pagos) { pago = calNumero(p.pago, 'El pago por hora', 500); bono = calNumero(p.bono, 'El bono', 20000); }
    if (previo) {
      h.getRange(previo.fila, 3, 1, 2).setValues([[nombre, area]]);
      if (ac.pagos) h.getRange(previo.fila, 5, 1, 2).setValues([[pago == null ? '' : pago, bono == null ? 0 : bono]]);
      h.getRange(previo.fila, 12, 1, 2).setValues([[yo.nombre, ahora]]);
    } else {
      var mx = 0; eq.forEach(function (x) { var n = parseInt(x.id.slice(1), 10) || 0; if (n > mx) mx = n; });
      id = 'P' + ('000' + (mx + 1)).slice(-4);
      var fila = h.getLastRow() + 1;
      h.getRange(fila, 7, 1, 2).setNumberFormat('@');
      h.getRange(fila, 1, 1, H_CEQ.length).setValues([[id, sid, nombre, area, pago == null ? '' : pago, bono == null ? 0 : bono,
        hoyISO(), '', yo.nombre, yo.rol, ahora, yo.nombre, ahora]]);
    }
    return calRespuesta(yo, ac, sid, fecha, previo ? nombre + ' quedó actualizado.' : nombre + ' ya está en el equipo. Póngale sus turnos en el calendario.');
  } finally { lock.releaseLock(); }
}
/** Se elimina del equipo. Lo ya trabajado queda en las quincenas anteriores; los días que faltan se limpian. */
function eliminarPersonaCal(cred, sid, id, fecha) {
  var yo = quien(cred), ac = calAcceso(yo); soloVerCal(ac);
  if (!ac.equipo) throw new Error('Su usuario solo consulta los calendarios.');
  if (ac.sucs.indexOf(sid) < 0) throw new Error('Esa no es su sucursal.');
  var lock = LockService.getScriptLock(); lock.waitLock(15000); _LEE = {};
  try {
    var eq = calLeeEq(), tu = calLeeTu(), p = null, hoy = hoyISO();
    eq.forEach(function (x) { if (x.id === String(id) && x.sid === sid && !x.baja) p = x; });
    if (!p) throw new Error('Esa persona ya no está en el equipo.');
    var propios = tu.filter(function (r) { return r.id === p.id && (r.am || r.pm); });
    var h = calHoja('Calendario equipo', H_CEQ, ['G', 'H']);
    if (!propios.length) h.deleteRow(p.fila);
    else {
      h.getRange(p.fila, 8).setValue(hoy);
      h.getRange(p.fila, 12, 1, 2).setValues([[yo.nombre, new Date()]]);
      var ht = calHoja('Calendario turnos', H_CTU, ['B']);
      propios.forEach(function (r) { if (r.f > hoy) ht.getRange(r.fila, 4, 1, 2).setValues([['', '']]); });
    }
    return calRespuesta(yo, ac, sid, fecha, p.nombre + ' salió del equipo.' + (propios.length ? ' Lo que ya trabajó se conserva.' : ''));
  } finally { lock.releaseLock(); }
}

/* ── configuración: turnos, pago base, semáforo y pago de los gerentes ── */
function guardarConfigCal(cred, sid, cfg, fecha) {
  var yo = quien(cred), ac = calAcceso(yo); soloVerCal(ac);
  if (!ac.pagos) throw new Error('La configuración de pagos la cambia el administrador o el director operativo.');
  if (ac.sucs.indexOf(sid) < 0) throw new Error('Escoja la sucursal.');
  cfg = cfg || {};
  var lista = [], t = cfg.turnos || {};
  [['am', 'La mañana'], ['pm', 'La tarde'], ['ap', 'El apoyo de tarde'], ['ap2', 'El apoyo 2']].forEach(function (a) {
    var par = t[a[0]] || [];
    if (a[0] === 'ap2' && !t.ap2) return;
    if (!horaValida(par[0]) || !horaValida(par[1])) throw new Error(a[1] + ': escriba hora de inicio y de fin.');
    lista.push([a[0] + '_i', horaNorm(par[0])], [a[0] + '_f', horaNorm(par[1])]);
  });
  var base = calNumero(cfg.base, 'El pago por hora base', 500);
  if (base == null) throw new Error('Escriba el pago por hora base.');
  lista.push(['base', base]);
  var s = cfg.semaforo || {}, v = {};
  ['bajoDe', 'bajoA', 'optDe', 'optA'].forEach(function (k) {
    v[k] = calNumero(s[k], 'El rango de horas', 200);
    if (v[k] == null) throw new Error('Complete los rangos de horas.');
  });
  if (!(v.bajoDe <= v.bajoA && v.bajoA <= v.optDe && v.optDe <= v.optA)) throw new Error('Los rangos de horas deben ir de menor a mayor: bajo, óptimo.');
  lista.push(['bajoDe', v.bajoDe], ['bajoA', v.bajoA], ['optDe', v.optDe], ['optA', v.optA]);
  var nombresG = usuariosCache().filter(function (u) { return u.activo && u.rol === 'gerente' && u.sucursal === sid; }).map(function (u) { return u.nombre; });
  (cfg.gerentes || []).forEach(function (g) {
    if (!g || nombresG.indexOf(g.nombre) < 0) return;
    var sal = calNumero(g.salario, 'El salario', 100000), bo = calNumero(g.bono, 'El bono', 100000);
    lista.push(['ger_salario:' + g.nombre, sal == null ? 0 : sal], ['ger_bono:' + g.nombre, bo == null ? 0 : bo]);
  });
  var lock = LockService.getScriptLock(); lock.waitLock(15000); _LEE = {};
  try {
    var filas = calLeeCf(), h = calHoja('Calendario config', H_CCF, ['C']), ahora = new Date(), nuevas = [];
    lista.forEach(function (par) {
      var previa = null; filas.forEach(function (x) { if (x.sid === sid && x.k === par[0]) previa = x; });
      if (previa) h.getRange(previa.fila, 3, 1, 3).setValues([[String(par[1]), yo.nombre, ahora]]);
      else nuevas.push([sid, par[0], String(par[1]), yo.nombre, ahora]);
    });
    if (nuevas.length) {
      var d0 = h.getLastRow() + 1;
      h.getRange(d0, 3, nuevas.length, 1).setNumberFormat('@');
      h.getRange(d0, 1, nuevas.length, H_CCF.length).setValues(nuevas);
    }
    return calRespuesta(yo, ac, sid, fecha, 'Configuración de ' + nombreUnidad(sid) + ' guardada. Las quincenas ya pagadas no cambian.');
  } finally { lock.releaseLock(); }
}
/** Horas de los dos apoyos de tarde: el gerente dice a qué hora entran (turnos escalonados). Solo cambia las quincenas sin pagar. */
function guardarHorasApoyo(cred, sid, fecha, ap, ap2) {
  var yo = quien(cred), ac = calAcceso(yo); soloVerCal(ac);
  if (!ac.turnos) throw new Error('Los horarios los cambia el gerente o la administración.');
  if (ac.sucs.indexOf(sid) < 0) throw new Error('Escoja la sucursal.');
  var lista = [];
  [['ap', 'Apoyo 1', ap], ['ap2', 'Apoyo 2', ap2]].forEach(function (a) {
    var par = a[2] || [];
    if (!horaValida(par[0]) || !horaValida(par[1])) throw new Error(a[1] + ': escriba hora de entrada y de salida.');
    lista.push([a[0] + '_i', horaNorm(par[0])], [a[0] + '_f', horaNorm(par[1])]);
  });
  var lock = LockService.getScriptLock(); lock.waitLock(15000); _LEE = {};
  try {
    var filas = calLeeCf(), h = calHoja('Calendario config', H_CCF, ['C']), ahora = new Date(), nuevas = [];
    lista.forEach(function (par) {
      var previa = null; filas.forEach(function (x) { if (x.sid === sid && x.k === par[0]) previa = x; });
      if (previa) h.getRange(previa.fila, 3, 1, 3).setValues([[String(par[1]), yo.nombre, ahora]]);
      else nuevas.push([sid, par[0], String(par[1]), yo.nombre, ahora]);
    });
    if (nuevas.length) {
      var d0 = h.getLastRow() + 1;
      h.getRange(d0, 3, nuevas.length, 1).setNumberFormat('@');
      h.getRange(d0, 1, nuevas.length, H_CCF.length).setValues(nuevas);
    }
    return calRespuesta(yo, ac, sid, fecha, 'Horas de apoyo de ' + nombreUnidad(sid) + ' guardadas.');
  } finally { lock.releaseLock(); }
}
/** Bono de esta quincena por persona (cambia el general solo para esta quincena). */
function guardarBonosCal(cred, sid, fecha, bonos) {
  var yo = quien(cred), ac = calAcceso(yo); soloVerCal(ac);
  if (!ac.pagos) throw new Error('Los bonos los cambia el administrador o el director operativo.');
  if (ac.sucs.indexOf(sid) < 0) throw new Error('Escoja la sucursal.');
  var lock = LockService.getScriptLock(); lock.waitLock(15000); _LEE = {};
  try {
    var D = calDatos(), q = calQuincena(fecha);
    if (calPlanillaDe(D, sid, q.ini)) throw new Error('La planilla de esta quincena ya está pagada.');
    var ids = {}, bg = calCfg(sid, D.cf).bonoGeneral; D.eq.forEach(function (p) { if (p.sid === sid) { ids[p.id] = p; p.bonoEf = p.bono > 0 ? p.bono : bg; } });
    var h = calHoja('Calendario bonos', H_CBO, ['B']), ahora = new Date(), n = 0, nuevas = [];
    Object.keys(bonos || {}).forEach(function (id) {
      if (!ids[id]) return;
      var b = calNumero(bonos[id], 'El bono', 20000);
      var previo = null; D.bo.forEach(function (x) { if (x.sid === sid && x.ini === q.ini && x.id === id) previo = x; });
      if (b == null || b === ids[id].bonoEf) {                  // igual al que le corresponde: se quita el ajuste
        if (previo) { h.getRange(previo.fila, 4, 1, 3).setValues([['', yo.nombre, ahora]]); n++; }
        return;
      }
      if (previo) { if (previo.bono !== b) { h.getRange(previo.fila, 4, 1, 3).setValues([[b, yo.nombre, ahora]]); n++; } }
      else { nuevas.push([sid, q.ini, id, b, yo.nombre, ahora]); n++; }
    });
    if (nuevas.length) {
      var d0 = h.getLastRow() + 1;
      h.getRange(d0, 2, nuevas.length, 1).setNumberFormat('@');
      h.getRange(d0, 1, nuevas.length, H_CBO.length).setValues(nuevas);
    }
    return calRespuesta(yo, ac, sid, fecha, n ? 'Bonos de la quincena guardados.' : 'No había cambios en los bonos.');
  } finally { lock.releaseLock(); }
}

/* ── el bono general: el mismo para todos, se cambia en Configuración ── */
function calGuardaCf(sid, k, v, quien, ahora) {
  var filas = calLeeCf(), h = calHoja('Calendario config', H_CCF, ['C']), previa = null;
  filas.forEach(function (x) { if (x.sid === sid && x.k === k) previa = x; });
  if (previa) h.getRange(previa.fila, 3, 1, 3).setValues([[String(v), quien, ahora]]);
  else {
    var d0 = h.getLastRow() + 1;
    h.getRange(d0, 3, 1, 1).setNumberFormat('@');
    h.getRange(d0, 1, 1, H_CCF.length).setValues([[sid, k, String(v), quien, ahora]]);
  }
}
/** Quién tiene un bono propio y cuál es el general: lo que pinta Configuración → Bono general. */
function leerBonoGeneral(cred) {
  var yo = quien(cred), ac = calAcceso(yo);
  if (!ac.dinero) throw new Error('El bono lo ven el administrador, el director operativo y finanzas.');
  var D = calDatos(), g = calCfg('centro', D.cf).bonoGeneral;
  var propios = D.eq.filter(function (p) { return !p.baja && p.bono > 0; }).map(function (p) {
    return { id: p.id, sucursal: nombreUnidad(p.sid), nombre: p.nombre, bono: p.bono };
  });
  var gente = D.eq.filter(function (p) { return !p.baja; }).length;
  return { ok: true, bonoGeneral: g, propios: propios, personas: gente, puedeCambiar: ac.pagos };
}
/** El bono de todos de una vez. Con «reemplazar» también se borran los bonos propios y todos quedan parejos. */
function guardarBonoGeneral(cred, valor, reemplazar, sid, fecha) {
  var yo = quien(cred), ac = calAcceso(yo); soloVerCal(ac);
  if (!ac.pagos) throw new Error('El bono lo cambia el administrador o el director operativo.');
  var b = calNumero(valor, 'El bono', 20000);
  if (b == null) throw new Error('Escriba el bono. Si no hay bono, escriba 0.');
  var lock = LockService.getScriptLock(); lock.waitLock(15000); _LEE = {};
  try {
    var ahora = new Date(), limpios = 0;
    calGuardaCf('*', 'bono_general', b, yo.nombre, ahora);
    if (reemplazar) {
      var h = calHoja('Calendario equipo', H_CEQ, ['G', 'H']);
      calLeeEq().forEach(function (p) {
        if (p.bono > 0) { h.getRange(p.fila, 6).setValue(0); h.getRange(p.fila, 12, 1, 2).setValues([[yo.nombre, ahora]]); limpios++; }
      });
    }
    var texto = 'Bono general de ' + dinero(b) + ' por quincena guardado para todo el personal.' +
      (limpios ? ' ' + limpios + (limpios === 1 ? ' bono propio se quitó' : ' bonos propios se quitaron') + ' y todos quedan parejos.' : '') +
      ' Las quincenas ya pagadas no cambian.';
    if (sid && ac.sucs.indexOf(sid) >= 0) return calRespuesta(yo, ac, sid, fecha, texto);
    return { ok: true, mensaje: texto, bonoGeneral: b, quitados: limpios };
  } finally { lock.releaseLock(); }
}

/* ── horas reales: al final de la quincena se anotan las que de verdad trabajó cada persona y sobre ellas se paga ── */
function guardarHorasReales(cred, sid, fecha, horas) {
  var yo = quien(cred), ac = calAcceso(yo); soloVerCal(ac);
  if (!ac.pagos) throw new Error('Las horas trabajadas las anota el administrador o el director operativo.');
  if (ac.sucs.indexOf(sid) < 0) throw new Error('Escoja la sucursal.');
  var lock = LockService.getScriptLock(); lock.waitLock(15000); _LEE = {};
  try {
    var D = calDatos(), q = calQuincena(fecha);
    if (calPlanillaDe(D, sid, q.ini)) throw new Error('La planilla de esta quincena ya está pagada.');
    var ids = {}; D.eq.forEach(function (p) { if (p.sid === sid) ids[p.id] = p; });
    var h = calHoja('Calendario horas reales', H_CHR, ['B']), ahora = new Date(), n = 0, nuevas = [];
    Object.keys(horas || {}).forEach(function (id) {
      if (!ids[id]) return;
      var v = horas[id], b = (v === '' || v == null) ? null : calNumero(v, 'Las horas de ' + ids[id].nombre, 250);
      var previo = null; D.hr.forEach(function (x) { if (x.sid === sid && x.ini === q.ini && x.id === id) previo = x; });
      if (previo) { if (previo.horas !== b) { h.getRange(previo.fila, 4, 1, 3).setValues([[b == null ? '' : b, yo.nombre, ahora]]); n++; } }
      else if (b != null) { nuevas.push([sid, q.ini, id, b, yo.nombre, ahora]); n++; }
    });
    if (nuevas.length) {
      var d0 = h.getLastRow() + 1;
      h.getRange(d0, 2, nuevas.length, 1).setNumberFormat('@');
      h.getRange(d0, 1, nuevas.length, H_CHR.length).setValues(nuevas);
    }
    return calRespuesta(yo, ac, sid, fecha, n ? 'Horas trabajadas guardadas. La planilla se calculó con ellas.' : 'No había cambios en las horas.');
  } finally { lock.releaseLock(); }
}

/* ── avances del gerente: el calendario a medias, guardado sin publicarlo (solo lo ve quien lo guardó) ── */
function calBorradorDe(yo, sid, ini) {
  var out = null;
  try {
    leeTodo(calHoja('Calendario borradores', H_CBR, ['C']), H_CBR.length).forEach(function (r) {
      if (calTxt(r[0]) === yo.nombre && calTxt(r[1]) === sid && calTxt(r[2]) === ini) {
        var c = calJson(calTxt(r[3]), null);
        if (Array.isArray(c) && c.length) out = { cambios: c, ms: r[4] instanceof Date ? r[4].getTime() : 0 };
      }
    });
  } catch (e) {}
  return out;
}
function calBorradorPon(yo, sid, ini, lista) {
  var h = calHoja('Calendario borradores', H_CBR, ['C']), filas = leeTodo(h, H_CBR.length), fila = 0;
  filas.forEach(function (r, i) { if (calTxt(r[0]) === yo.nombre && calTxt(r[1]) === sid && calTxt(r[2]) === ini) fila = i + 2; });
  if (!lista.length) { if (fila) h.deleteRow(fila); return; }
  var txt = JSON.stringify(lista);
  if (fila) h.getRange(fila, 4, 1, 2).setValues([[txt, new Date()]]);
  else {
    var d0 = h.getLastRow() + 1;
    h.getRange(d0, 3, 1, 1).setNumberFormat('@');
    h.getRange(d0, 1, 1, H_CBR.length).setValues([[yo.nombre, sid, ini, txt, new Date()]]);
  }
}
function guardarBorradorCal(cred, sid, fecha, cambios) {
  var yo = quien(cred), ac = calAcceso(yo); soloVerCal(ac);
  if (!ac.turnos) throw new Error('Su usuario solo consulta los calendarios.');
  if (ac.sucs.indexOf(sid) < 0) throw new Error('Ese calendario no es de su sucursal.');
  if (!Array.isArray(cambios)) cambios = [];
  if (cambios.length > 400) throw new Error('Son demasiados cambios para guardar como avance. Publique el calendario.');
  var q = calQuincena(fecha), lista = [];
  cambios.forEach(function (c) {
    c = c || {};
    var f = String(c.f || ''), id = String(c.id || '');
    if (!id || q.dias.indexOf(f) < 0) return;
    lista.push({ id: id, f: f, am: (c.am === true || c.am === 'x') ? 'x' : '', pm: c.pm === 'x' || c.pm === 'xa' || c.pm === 'xb' ? c.pm : '' });
  });
  var lock = LockService.getScriptLock(); lock.waitLock(15000); _LEE = {};
  try {
    calBorradorPon(yo, sid, q.ini, lista);
    return { ok: true, mensaje: lista.length ? 'Avances guardados (' + lista.length + (lista.length === 1 ? ' cambio' : ' cambios') +
      '). Todavía no se publican: siga cuando quiera y toque «Guardar calendario» al terminar.' : 'No había avances que guardar.', guardados: lista.length };
  } finally { lock.releaseLock(); }
}

/* ── la planilla: se confirma cuando se paga ── */
function confirmarPlanilla(cred, sid, fecha, pagado, fechaPago, nota) {
  var yo = quien(cred), ac = calAcceso(yo); soloVerCal(ac);
  if (!ac.confirma) throw new Error('La planilla la confirma el administrador.');
  if (ac.sucs.indexOf(sid) < 0) throw new Error('Escoja la sucursal.');
  var lock = LockService.getScriptLock(); lock.waitLock(15000); _LEE = {};
  try {
    var D = calDatos(), m = calModelo(D, sid, fecha), q = m.q, hoy = hoyISO();
    if (calPlanillaDe(D, sid, q.ini)) throw new Error('Esta quincena ya está confirmada.');
    if (hoy < q.fin) throw new Error('La quincena termina el ' + dia(q.fin) + '. Confirme la planilla cuando se pague.');
    if (!(m.tot.total > 0)) throw new Error('No hay horas ni salarios en esta quincena. Arme primero el calendario.');
    var monto = pagado === '' || pagado == null ? m.tot.total : calNumero(pagado, 'El monto pagado', 500000);
    if (!(monto > 0)) throw new Error('Escriba lo que se pagó.');
    var fp = fechaPago || hoy;
    if (!esFecha(fp) || fp > hoy) throw new Error('La fecha de pago no es válida.');
    nota = String(nota || '').trim().slice(0, 120);
    if (Math.abs(monto - m.tot.total) >= 1 && nota.length < 4)
      throw new Error('Lo pagado (' + dinero(monto) + ') no es lo calculado (' + dinero(m.tot.total) + '). Escriba por qué en la nota.');
    var v = preparaGasto({ unidad: sid, fechaPago: fp, categoria: 'Nómina y bonificaciones', monto: monto,
      proveedor: 'Planilla ' + nombreUnidad(sid), factura: 'PL-' + sid + '-' + q.ini,
      nota: ('Planilla ' + q.corto + ' · ' + m.tot.horas + ' h' + (nota ? ' · ' + nota : '')).slice(0, 140) });
    /* La planilla se paga al terminar la quincena (la del 1 al 15 el día 16; la del 16 al fin de mes el día 1 del mes siguiente),
     * pero el gasto es del mes que se trabajó: se anota en ese mes aunque el pago caiga en el siguiente. */
    v.mes = q.ini.slice(0, 7);
    var g = escribeGasto(v, yo.nombre, new Date(), null);
    var h = calHoja('Calendario planilla', H_CPL, ['B', 'C', 'G']), fila = h.getLastRow() + 1;
    h.getRange(fila, 2, 1, 2).setNumberFormat('@'); h.getRange(fila, 7).setNumberFormat('@');
    h.getRange(fila, 1, 1, H_CPL.length).setValues([[sid, q.ini, q.fin, m.tot.horas, m.tot.proy, monto, fp, yo.nombre,
      new Date(), g.nuevo.id, 'Pagada', nota, JSON.stringify(calPagos(m))]]);
    return calRespuesta(yo, ac, sid, fecha, 'Planilla de ' + nombreUnidad(sid) + ' confirmada: ' + dinero(monto) + '. Quedó como gasto de Nómina y bonificaciones.');
  } finally { lock.releaseLock(); }
}
/** Las quincenas recientes de una sucursal y cuáles se pueden pagar: las ya pagadas y las que no han terminado quedan bloqueadas. */
function planillasPorPagar(cred, sid) {
  var yo = quien(cred), ac = calAcceso(yo);
  if (!ac.confirma && !ac.soloVer) throw new Error('La planilla la confirma el administrador.');
  if (ac.sucs.indexOf(sid) < 0) throw new Error('Escoja la sucursal.');
  var D = calDatos(), hoy = hoyISO(), f = calQuincena(hoy).ini, lista = [];
  for (var i = 0; i < 9; i++) {
    var m = calModelo(D, sid, f), q = m.q, pl = calPlanillaDe(D, sid, q.ini), estado, motivo = '';
    if (pl) { estado = 'Pagada'; motivo = 'ya se pagó' + (pl.fechaPago ? ' el ' + dia(pl.fechaPago) : ''); }
    else if (hoy < q.fin) { estado = 'Falta'; motivo = 'termina el ' + dia(q.fin); }
    else if (!(m.tot.total > 0)) { estado = 'Sin datos'; motivo = 'sin horas ni salarios'; }
    else estado = 'Por pagar';
    lista.push({ ini: q.ini, fin: q.fin, nombre: q.nombre, corto: q.corto, mesGasto: q.ini.slice(0, 7), pagarEl: masDias(q.fin, 1),
      estado: estado, bloqueada: estado !== 'Por pagar', motivo: motivo, total: m.tot.total, horas: m.tot.horas,
      sinReales: m.tot.sinReales, conReales: m.tot.conReales });
    f = masDias(q.ini, -1);
  }
  return { ok: true, hoy: hoy, sid: sid, lista: lista };
}

/** Una sola vez: las planillas que ya estaban anotadas en el mes en que se pagaron pasan al mes que se trabajó. */
function corrigeMesPlanillas() {
  try {
    var pr = typeof PropertiesService !== 'undefined' ? PropertiesService.getScriptProperties() : null;
    if (pr && pr.getProperty('mesPlanillas1') === 'ok') return;
    var h = libro().getSheetByName('Gastos');
    if (h && h.getLastRow() > 1) {
      var n = h.getLastRow() - 1, fac = h.getRange(2, 16, n, 1).getValues(), mes = h.getRange(2, 8, n, 1).getValues(), cambio = false;
      for (var i = 0; i < n; i++) {
        var mt = /^PL-[a-z0-9]+-(\d{4}-\d{2})-\d{2}$/i.exec(String(fac[i][0] || ''));
        if (mt && String(mes[i][0]).slice(0, 7) !== mt[1]) { h.getRange(i + 2, 8).setNumberFormat('@').setValue(mt[1]); cambio = true; }
      }
      if (cambio) _LEE = {};
    }
    if (pr) pr.setProperty('mesPlanillas1', 'ok');
  } catch (e) { /* si algo falla, no se estorba a la persona: se intenta otra vez la próxima */ }
}

function reabrirPlanilla(cred, sid, fecha, motivo) {
  var yo = quien(cred), ac = calAcceso(yo); soloVerCal(ac);
  if (!ac.confirma) throw new Error('La planilla la reabre el administrador.');
  if (ac.sucs.indexOf(sid) < 0) throw new Error('Escoja la sucursal.');
  motivo = String(motivo || '').trim();
  if (motivo.length < 4) throw new Error('Escriba el motivo.');
  var lock = LockService.getScriptLock(); lock.waitLock(15000); _LEE = {};
  try {
    var D = calDatos(), q = calQuincena(fecha), pl = calPlanillaDe(D, sid, q.ini);
    if (!pl) throw new Error('Esta quincena no estaba confirmada.');
    if (pl.gasto) anular({ cred: cred, clase: 'gasto', id: pl.gasto, motivo: 'Planilla reabierta: ' + motivo });
    var h = calHoja('Calendario planilla', H_CPL, ['B', 'C', 'G']);
    h.getRange(pl.fila, 11).setValue('Reabierta'); h.getRange(pl.fila, 12).setValue((pl.nota ? pl.nota + ' · ' : '') + 'reabierta: ' + motivo);
    return calRespuesta(yo, ac, sid, fecha, 'La planilla se reabrió y el gasto se anuló. Ya puede corregir y volver a confirmar.');
  } finally { lock.releaseLock(); }
}

/* ── el PDF para los compañeros: horarios y horas, sin dinero ── */
/** Por qué un calendario todavía no está terminado (lista vacía = terminado). */
function calIncompleto(m, pendiente) {
  var r = [];
  if (pendiente) r.push('hay cambios sin guardar');
  var sin = m.personas.filter(function (p) { return !(p.hq > 0); }).length;
  if (sin) r.push(sin + (sin === 1 ? ' persona sin turnos' : ' personas sin turnos'));
  var sd = m.personas.filter(function (p) { return p.hq > 0 && (p.sd1 || p.sd2); }).length;
  if (sd) r.push(sd + (sd === 1 ? ' persona sin su día de descanso' : ' personas sin su día de descanso'));
  var vac = m.dias.filter(function (d) { return d.gente === 0; }).length;
  if (vac) r.push(vac + (vac === 1 ? ' día sin nadie' : ' días sin nadie'));
  return r;
}

function calHtmlPdf(m, sucNombre, motivos, soloSemana) {
  var q = m.q, h = [], colores = { am: '#FFE6A0', pm: '#14284B', ap: '#8FA6CC', ap2: '#C9D6EC' };
  /* Encabezados con colores alternos (por día y por AM/PM); en el cuerpo solo alternan las filas, para seguir el turno de cada persona. */
  var fondo = ['#F7F2E7', '#DDE5F3'];
  var hh = function (t) { var p = String(t).split(':'), n = +p[0]; return (n % 12 || 12) + ':' + p[1] + (n < 12 ? ' a.m.' : ' p.m.'); };
  h.push('<html><head><meta charset="utf-8"><style>');
  h.push('body{font-family:Arial,Helvetica,sans-serif;font-size:10px;color:#16223A;margin:22px}');
  h.push('table.band{width:100%;background:#14284B;color:#fff;border-bottom:4px solid #F2B705;margin:0 0 10px;border-collapse:collapse}');
  h.push('table.band td{border:0;padding:10px 12px;vertical-align:middle}.band img{width:40px}');
  h.push('.band td.r{text-align:right;font-size:12px;color:#F7F2E7;font-weight:bold}');
  h.push('h1{font-size:19px;color:#fff;margin:0;letter-spacing:1px}.band .co{color:#F2B705;font-size:9px;letter-spacing:2px;font-weight:bold}');
  h.push('h2{font-size:13px;color:#14284B;margin:6px 0 6px}');
  h.push('table.g{width:100%;border-collapse:collapse;table-layout:fixed}');
  h.push('table.g th,table.g td{border:1px solid #CFC6B2;text-align:center;padding:4px 0;font-size:8.5px}');
  h.push('table.g th.n{background:#14284B;color:#F2B705;width:22%}table.g th.hs{background:#14284B;color:#F2B705;width:11%}');
  h.push('table.g td.nom{text-align:left;padding-left:5px;font-weight:bold;font-size:9px;color:#14284B}table.g td.hs{font-weight:bold;color:#14284B}');
  h.push('table.g tr.ar td{background:#14284B;color:#F2B705;text-align:left;padding:4px 6px;font-weight:bold;letter-spacing:1px}');
  h.push('table.g tr.tot td{background:#F2B705;font-weight:bold;color:#14284B}');
  h.push('td.am{background:' + colores.am + ';color:#14284B;font-weight:bold}td.pm{background:' + colores.pm + ';color:#fff;font-weight:bold}td.ap{background:' + colores.ap + ';color:#14284B;font-weight:bold}td.ap2{background:' + colores.ap2 + ';color:#14284B;font-weight:bold}');
  h.push('.ley{margin-top:10px;font-size:9px;color:#44506A;line-height:1.7}.ley b{display:inline-block;padding:1px 7px;border-radius:3px;margin-right:3px}');
  h.push('.pie{margin-top:12px;border-top:1px solid #E3DBCB;padding-top:6px;font-size:8.5px;color:#6F7686}');
  h.push('table.inc{width:100%;border-collapse:collapse;margin:0 0 8px}table.inc td{background:#C62828;color:#fff;text-align:center;padding:7px 8px;font-weight:bold;font-size:12px;letter-spacing:1px;border:0}table.inc td small{display:block;font-weight:normal;font-size:8.5px;letter-spacing:0;margin-top:2px}');
  h.push('</style></head><body>');
  var bloques = [[0, q.n1, 'Semana 1'], [q.n1, q.n, 'Semana 2']];
  if (soloSemana === 1 || soloSemana === 2) bloques = [bloques[soloSemana - 1]];
  bloques.forEach(function (b, bi) {
    var a = b[0], z = b[1], dias = q.dias.slice(a, z), ultima = bi === bloques.length - 1;
    h.push('<div' + (!ultima ? ' style="page-break-after:always"' : '') + '>');
    h.push('<table class="band"><tr><td style="width:52px"><img src="' + LOGO_PNG + '"></td><td><h1>REY PIZZA</h1><div class="co">CAIX, S.A.</div></td>' +
      '<td class="r">Calendario de trabajo<br>' + htm(sucNombre) + '</td></tr></table>');
    if (soloSemana) h.push('<div style="text-align:right;font-size:9px;font-weight:bold;letter-spacing:1px;color:#1F7A4D;margin:-4px 0 6px">CALENDARIO OFICIAL</div>');
    if (motivos && motivos.length) h.push('<table class="inc"><tr><td>INFORMACIÓN INCOMPLETA<small>Muestra: el calendario aún no está terminado (' + htm(motivos.join(', ')) + '). No es el definitivo.</small></td></tr></table>');
    h.push('<h2>' + b[2] + ' · ' + dia(dias[0]).slice(0, 5) + ' al ' + dia(dias[dias.length - 1]) + ' &nbsp;<span style="font-weight:normal;color:#6F7686;font-size:10px">(' + htm(q.nombre) + ')</span></h2>');
    var enc = function (i, tx) {                            // encabezados de día: dorado y azul claro, uno sí y uno no
      var bg = i % 2 ? '#C9D6EC' : '#F2B705';
      return '<th colspan="2" style="background:' + bg + ';color:#14284B;font-size:8.5px">' + tx + '</th>';
    };
    h.push('<table class="g"><tr><th class="n" rowspan="3">&nbsp;</th>');
    dias.forEach(function (f, i) { h.push(enc(i, dia(f).slice(0, 5))); });
    if (ultima) h.push('<th class="hs" rowspan="3">Total de la quincena</th>');
    h.push('</tr><tr>');
    dias.forEach(function (f, i) { h.push(enc(i, CAL_D3[diaSemana(f)])); });
    h.push('</tr><tr>');
    dias.forEach(function () {
      h.push('<th style="background:#FFF1C2;color:#14284B;font-size:7.5px">AM</th><th style="background:#DCE3F0;color:#14284B;font-size:7.5px">PM</th>');
    });
    h.push('</tr>');
    var fila = 0, colspan = dias.length * 2 + 1 + (ultima ? 1 : 0);
    CAL_AREAS.forEach(function (ar) {
      var fs = m.personas.filter(function (p) { return p.area === ar; });
      if (!fs.length) return;
      h.push('<tr class="ar"><td colspan="' + colspan + '">' + (ar === 'Barra' ? 'BARRA · SERVICIO' : 'COCINA') + '</td></tr>');
      fs.forEach(function (p) {
        var r = fila++ % 2;
        h.push('<tr><td class="nom" style="background:' + fondo[r] + '">' + htm(p.nombre) + '</td>');
        p.dias.slice(a, z).forEach(function (d) {
          var st = ' style="background:' + fondo[r] + '"';
          h.push('<td' + (d.am ? ' class="am"' : st) + '>' + (d.am ? 'x' : '&nbsp;') + '</td>' +
            '<td' + (d.pm ? ' class="' + (d.pm === 'xa' ? 'ap' : d.pm === 'xb' ? 'ap2' : 'pm') + '"' : st) + '>' + (d.pm || '&nbsp;') + '</td>');
        });
        if (ultima) h.push('<td class="hs" style="background:' + fondo[r] + '">' + calRed(p.hq) + ' h</td>');
        h.push('</tr>');
      });
    });
    h.push('</table>');
    h.push('<div class="ley"><b style="background:' + colores.am + ';color:#14284B">x</b> Mañana ' + hh(m.cfg.am_i) + ' a ' + hh(m.cfg.am_f) +
      ' &nbsp; <b style="background:' + colores.pm + ';color:#fff">x</b> Tarde ' + hh(m.cfg.pm_i) + ' a ' + hh(m.cfg.pm_f) +
      ' &nbsp; <b style="background:' + colores.ap2 + ';color:#14284B">xb</b> Apoyo 2 ' + hh(m.cfg.ap2_i) + ' a ' + hh(m.cfg.ap2_f) +
      ' &nbsp; <b style="background:' + colores.ap + ';color:#14284B">xa</b> Apoyo 1 ' + hh(m.cfg.ap_i) + ' a ' + hh(m.cfg.ap_f) +
      '<br>Casilla vacía = día de descanso.</div>');
    h.push('<div class="pie">Generado el ' + Utilities.formatDate(new Date(), ZONA, 'dd/MM/yyyy HH:mm') + ' desde la app REY PIZZA · CAIX, S.A. Cualquier cambio lo avisa su gerente.</div>');
    h.push('</div>');
  });
  h.push('</body></html>');
  return h.join('');
}
function pdfCalendario(cred, sid, fecha, pendiente, opciones) {
  var yo = quien(cred), ac = calAcceso(yo);
  if (ac.sucs.indexOf(sid) < 0) throw new Error('Escoja la sucursal.');
  var m = calModelo(calDatos(), sid, fecha);
  if (!m.personas.length) throw new Error('Este calendario no tiene equipo. Agregue a las personas primero.');
  opciones = opciones || {};
  var sem = opciones.tipo === 'oficial' ? (Number(opciones.semana) === 2 ? 2 : 1) : 0;
  var base = 'Calendario ' + nombreUnidad(sid) + ' ' + m.q.corto.replace('–', '-');
  if (sem) {
    /* El oficial va por semana, y solo sale si ESA semana está completa (sin cambios pendientes, todos con turnos y su descanso, ningún día vacío). */
    var faltan = calIncompletoSemana(m, sem, pendiente);
    if (faltan.length) throw new Error('La semana ' + sem + ' todavía no está lista para el calendario oficial: ' + faltan.join(', ') + '. Puede descargar la muestra.');
    var nombreO = base + ' Semana ' + sem + ' (oficial).pdf';
    var blobO = Utilities.newBlob(calHtmlPdf(m, nombreUnidad(sid), [], sem), 'text/html', 'c.html').getAs('application/pdf').setName(nombreO);
    return { ok: true, nombre: nombreO, tipo: 'application/pdf', base64: Utilities.base64Encode(blobO.getBytes()), incompleto: false, motivos: [], oficial: true, semana: sem };
  }
  /* La muestra es la quincena completa; si falta algo, lleva el letrero INFORMACIÓN INCOMPLETA. */
  var motivos = calIncompleto(m, pendiente);
  var nombre = base + (motivos.length ? ' (muestra, INCOMPLETO)' : ' (muestra)') + '.pdf';
  var blob = Utilities.newBlob(calHtmlPdf(m, nombreUnidad(sid), motivos), 'text/html', 'c.html').getAs('application/pdf').setName(nombre);
  return { ok: true, nombre: nombre, tipo: 'application/pdf', base64: Utilities.base64Encode(blob.getBytes()), incompleto: motivos.length > 0, motivos: motivos };
}
/** Lo que le falta a una semana para el oficial. */
function calIncompletoSemana(m, sem, pendiente) {
  var r = [], a = sem === 2 ? m.q.n1 : 0, z = sem === 2 ? m.q.n : m.q.n1;
  if (pendiente) r.push('hay cambios sin guardar');
  var sin = m.personas.filter(function (p) { return !((sem === 2 ? p.hb2 : p.hb1) > 0); }).length;
  if (sin) r.push(sin + (sin === 1 ? ' persona sin turnos' : ' personas sin turnos'));
  var sd = m.personas.filter(function (p) { return (sem === 2 ? p.hb2 : p.hb1) > 0 && (sem === 2 ? p.sd2 : p.sd1); }).length;
  if (sd) r.push(sd + (sd === 1 ? ' persona sin su día de descanso' : ' personas sin su día de descanso'));
  var vac = m.dias.slice(a, z).filter(function (d) { return d.gente === 0; }).length;
  if (vac) r.push(vac + (vac === 1 ? ' día sin nadie' : ' días sin nadie'));
  return r;
}

/* ── campanita ── */
function avisosCalendario(yo, out, desde) {
  var esJefe = yo.esAdmin || yo.rol === 'operaciones';
  if (!esJefe) return;
  var D = calDatos(), hoy = hoyISO();
  D.eq.forEach(function (p) {
    if (p.baja || p.rolPor !== 'gerente' || p.creado.slice(0, 10) < desde || p.pago != null || p.bono) return;
    out.push({ id: 'cal-p-' + p.id, tipo: 'Personal nuevo', cuando: p.creado, titulo: nombreUnidad(p.sid) + ' · ' + p.nombre + ' (' + p.area + ')',
      texto: 'Lo agregó ' + p.por + '. Póngale su pago por hora si no es el de la base.', abrir: { pag: 'cal', ct: 'equipo' } });
  });
  if (!yo.esAdmin) return;
  var vistas = {};
  [masDias(hoy, -1), masDias(hoy, -16)].forEach(function (f) {
    var q = calQuincena(f);
    if (q.fin >= hoy || vistas[q.ini]) return;
    vistas[q.ini] = true;
    ['centro', 'almendras', 'parque'].forEach(function (sid) {
      if (calPlanillaDe(D, sid, q.ini)) return;
      var m = calModelo(D, sid, q.ini);
      if (!(m.tot.total > 0)) return;
      out.push({ id: 'cal-pl-' + sid + '-' + q.ini, tipo: 'Planilla por confirmar', cuando: q.fin + ' 23:00',
        titulo: nombreUnidad(sid) + ' · ' + q.corto, texto: (m.tot.sinReales ? 'Faltan las horas trabajadas de ' + m.tot.sinReales + (m.tot.sinReales === 1 ? ' persona' : ' personas') + '. ' : 'Con las horas trabajadas. ') +
          dinero(m.tot.total) + ' · ' + m.tot.horas + ' h. Confírmela cuando la pague.',
        abrir: { pag: 'cal', ct: 'pagos', sid: sid, fecha: q.ini } });
    });
  });
}

/* ════════════ ZONAS DE INVENTARIO (PRODUCTOS / ÁREAS) ════════════
 * Cada gerente arma las zonas de SU sucursal con el nombre que quiera («Pizzas», «Barra», «Refri»…) y marca
 * qué productos van en cada una. En «Contar y pedir» esas zonas funcionan como filtro: el gerente cuenta
 * una zona a la vez y termina más rápido. Un producto puede estar en más de una zona.
 * Hoja: «Zonas inventario» (una fila por zona; «Productos» guarda los códigos separados por coma). */
var H_ZONAS = ['Sucursal', 'Id', 'Zona', 'Orden', 'Productos', 'Actualizado por', 'Actualizado en'];
var MAX_ZONAS = 20;
function hojaZonas(ss) {
  ss = ss || libro();
  var h = ss.getSheetByName('Zonas inventario');
  if (!h) {
    h = hojaLimpia(ss, 'Zonas inventario', H_ZONAS);
    h.getRange('G:G').setNumberFormat('dd/mm/yyyy hh:mm');
    h.setColumnWidth(3, 200); h.setColumnWidth(5, 420);
  }
  return h;
}
function filasZonas() {
  return leeTodo(hojaZonas(), H_ZONAS.length).map(function (r) {
    return { suc: String(r[0] || '').trim(), id: String(r[1] || '').trim(), nombre: String(r[2] || '').trim(),
      orden: Number(r[3]) || 0, por: String(r[5] || ''), en: r[6],
      codigos: String(r[4] || '').split(',').map(function (c) { return c.trim(); }).filter(Boolean) };
  }).filter(function (z) { return z.suc && z.id && z.nombre; });
}
function zonasDe(unidadId) {
  return filasZonas().filter(function (z) { return z.suc === unidadId; })
    .sort(function (a, b) { return a.orden - b.orden; });
}
/** Gerente: su sucursal. Administrador y director operativo: la que pidan (solo sucursales que venden). */
function unidadZonas(yo, pedida) {
  if (yo.rol === 'gerente') return yo.sucursal;
  if (veInv(yo)) { var u = unidadPorId(pedida); return u && u.vende ? u.id : SUC_KPI[0]; }
  throw new Error('Las zonas de inventario las arma el gerente de cada sucursal.');
}
function pantallaZonas(cred, unidadId) {
  var yo = quien(cred), uid = unidadZonas(yo, unidadId), un = unidadPorId(uid);
  var prods = productosInv().filter(function (b) { return maneja(b, uid); }).map(function (b) {
    return { codigo: b.codigo, nombre: b.nombre, area: b.area, grupo: b.grupo, estante: b.estante };
  });
  return { unidad: { id: uid, nombre: un ? un.nombre : uid },
    sucursales: veInv(yo) ? SUC_KPI.map(function (s) { return { id: s, nombre: nombreUnidad(s) }; }) : [],
    zonas: zonasDe(uid).map(function (z) { return { id: z.id, nombre: z.nombre, codigos: z.codigos }; }),
    productos: prods, max: MAX_ZONAS };
}
/** Guarda de una vez TODAS las zonas de la sucursal: zonas = [{id, nombre, codigos:[…]}]. */
function guardarZonas(cred, unidadId, zonas) {
  var yo = quien(cred), uid = unidadZonas(yo, unidadId);
  zonas = Array.isArray(zonas) ? zonas : [];
  var validos = {};
  productosInv().forEach(function (b) { if (maneja(b, uid)) validos[b.codigo] = true; });
  var vistos = {}, limpias = [], marca = Utilities.formatDate(new Date(), ZONA, 'yyMMddHHmmss');
  zonas.forEach(function (z, i) {
    var nombre = String((z && z.nombre) || '').replace(/\s+/g, ' ').trim().slice(0, 30);
    if (!nombre) return;
    var llave = sinAcentoS(nombre);
    if (vistos[llave]) throw new Error('Hay dos zonas con el nombre «' + nombre + '». Use nombres distintos.');
    vistos[llave] = true;
    var id = /^Z[0-9]{8,20}$/.test(String(z.id || '')) ? String(z.id) : 'Z' + marca + (i < 10 ? '0' : '') + i;
    var cods = [];
    (z.codigos || []).forEach(function (c) { c = String(c).trim(); if (validos[c] && cods.indexOf(c) < 0) cods.push(c); });
    limpias.push({ id: id, nombre: nombre, codigos: cods });
  });
  if (limpias.length > MAX_ZONAS) throw new Error('Son demasiadas zonas (máximo ' + MAX_ZONAS + ').');
  var lock = LockService.getScriptLock();
  lock.waitLock(15000); _LEE = {};
  try {
    var h = hojaZonas(), ahora = new Date();
    var otras = filasZonas().filter(function (z) { return z.suc !== uid; });
    var filas = otras.map(function (z) { return [z.suc, z.id, z.nombre, z.orden, z.codigos.join(','), z.por, z.en]; });
    limpias.forEach(function (z, i) { filas.push([uid, z.id, z.nombre, i + 1, z.codigos.join(','), yo.nombre, ahora]); });
    var filasViejas = Math.max(0, h.getLastRow() - 1);
    if (filasViejas) h.getRange(2, 1, filasViejas, H_ZONAS.length).clearContent();
    if (filas.length) h.getRange(2, 1, filas.length, H_ZONAS.length).setValues(filas);
  } finally { lock.releaseLock(); }
  _LEE = {};
  var r = pantallaZonas(cred, uid); r.ok = true;
  r.mensaje = limpias.length ? 'Guardado: ' + limpias.length + (limpias.length === 1 ? ' zona.' : ' zonas.') : 'Se quitaron todas las zonas.';
  return r;
}

/* ════════════ INDICADORES DE TECNOLOGÍA ════════════
 * Los decide y los anota el administrador (es el responsable del área). Funcionan igual que los de
 * marketing, pero no pasan por «borrador»: quedan activos al guardarlos. */
var KPI_CATS_TEC = ['Sistemas y app', 'Equipos y red', 'Datos y respaldos', 'Seguridad', 'Soporte'];
var KPIS_TEC_RECOMENDADOS = [
  { c: 'Sistemas y app', n: 'La app abre y responde bien en las 3 sucursales', t: 'Sí / No', u: '', min: null, max: null, f: 'Diario',
    i: 'Abra la app desde un teléfono, entre con su PIN y revise que cargue la pantalla de inicio.', p: 'Si la app falla, las sucursales no pueden registrar ventas ni inventario.' },
  { c: 'Sistemas y app', n: 'Tiempo que tarda en abrir la app', t: 'Número', u: 'segundos', min: null, max: 5, f: 'Semanal',
    i: 'Cronometre desde que toca el ícono hasta ver la pantalla de inicio, con datos móviles.', p: 'Si tarda mucho, el personal deja de usarla.' },
  { c: 'Sistemas y app', n: 'Errores de la app reportados por el personal en la semana', t: 'Número', u: 'errores', min: null, max: 3, f: 'Semanal',
    i: 'Cuente lo que reportaron los gerentes (mensajes de error, pantallas que no cargan).', p: 'Muestra qué tanto estorba el sistema en el trabajo diario.' },
  { c: 'Datos y respaldos', n: 'Copia de seguridad de la hoja de cálculo hecha', t: 'Sí / No', u: '', min: null, max: null, f: 'Semanal',
    i: 'Haga una copia del archivo «Caja Rey Pizza» en Drive (Archivo → Hacer una copia) y póngale la fecha.', p: 'Toda la operación vive en esa hoja: sin copia, un error borra todo.' },
  { c: 'Datos y respaldos', n: 'Correos automáticos enviados sin errores', t: 'Sí / No', u: '', min: null, max: null, f: 'Semanal',
    i: 'Revise la hoja «Correos enviados»: ninguna fila debe decir que falló.', p: 'Los avisos y reportes llegan por correo; si fallan, nadie se entera.' },
  { c: 'Equipos y red', n: 'Internet funcionando en las 3 sucursales', t: 'Sí / No', u: '', min: null, max: null, f: 'Diario',
    i: 'Confirme con cada gerente que la conexión no se cayó durante el servicio.', p: 'Sin internet no se registra la caja.' },
  { c: 'Equipos y red', n: 'Teléfonos, tabletas e impresoras de las sucursales funcionando', t: 'Sí / No', u: '', min: null, max: null, f: 'Semanal',
    i: 'Pregunte a cada gerente si hay algún equipo dañado o sin batería.', p: 'Un equipo dañado detiene el registro del día.' },
  { c: 'Seguridad', n: 'Nadie usa el PIN inicial (1234)', t: 'Sí / No', u: '', min: null, max: null, f: 'Mensual',
    i: 'Revise la hoja «Usuarios» y pida cambiar los PIN que sigan en el valor inicial.', p: 'Un PIN conocido deja entrar a cualquiera con el nombre de otro.' },
  { c: 'Seguridad', n: 'Usuarios del personal que ya no trabaja, dados de baja', t: 'Sí / No', u: '', min: null, max: null, f: 'Mensual',
    i: 'Compare la lista de usuarios con el personal actual y dé de baja a quien ya no esté.', p: 'Una persona que se fue no debe seguir viendo datos de la empresa.' },
  { c: 'Soporte', n: 'Solicitudes de soporte atendidas en menos de 24 horas', t: 'Número', u: '%', min: 90, max: null, f: 'Semanal',
    i: 'De las solicitudes que llegaron en la semana, el porcentaje que se resolvió en un día.', p: 'La rapidez en resolver ayuda a que el equipo confíe en el sistema.' },
  { c: 'Soporte', n: 'Solicitudes de soporte pendientes', t: 'Número', u: 'solicitudes', min: null, max: 3, f: 'Semanal',
    i: 'Cuente lo que sigue sin resolver al cerrar la semana.', p: 'Si se acumulan, el sistema pierde confianza.' }
];
function exigeAdminTec(cred) { var yo = exigeAdmin(cred); _KPI_TEC = true; return yo; }
function listaKpisTec(cred) { exigeAdminTec(cred); return listaKpis(cred); }
function guardarKpiTec(cred, k) { exigeAdminTec(cred); return guardarKpi(cred, k); }
function borrarKpiTec(cred, id) { exigeAdminTec(cred); return borrarKpi(cred, id); }
function enviarKpisTec(cred, ids) { exigeAdminTec(cred); return enviarKpis(cred, ids); }
function agregarRecomendadosTec(cred, claves) { exigeAdminTec(cred); return agregarRecomendados(cred, claves); }
function resultadosKpiTec(cred, desde, hasta) { exigeAdminTec(cred); return resultadosKpiRango(cred, desde, hasta, ''); }
function pantallaKpiTec(cred) { exigeAdminTec(cred); return pantallaKpi(cred); }
function registrarKpisTec(cred, valores) { exigeAdminTec(cred); return registrarKpis(cred, valores); }

/* ════════════ RESUMEN POR ÁREAS (administrador) ════════════
 * La pantalla de inicio del administrador: Operaciones, Finanzas, Marketing y Tecnología, cada una con
 * su avance (qué porcentaje de lo que tocaba anotar sí se anotó), los últimos 14 días y algunas señales. */
function nivelPct(p) { return p == null ? 'sin' : p >= 90 ? 'ok' : p >= 70 ? 'alerta' : 'mal'; }
function resumenAreas(cred, fresco) {
  exigeAdmin(cred);
  var cc = null; try { cc = CacheService.getScriptCache(); } catch (e) {}
  if (cc && !fresco) { try { var t = cc.get('resumen_areas_v2'); if (t) return JSON.parse(t); } catch (e) {} }
  var r = resumenAreasCalc();
  if (cc) { try { var s = JSON.stringify(r); if (s.length < 90000) cc.put('resumen_areas_v2', s, 300); } catch (e) {} }
  return r;
}
/** Solo lo que sale de los indicadores: rápido. Las «señales» (inventario, impuestos, presupuesto…) van aparte, en resumenAreasSenales. */
function resumenAreasCalc() {
  var t0 = Date.now(), hoy = hoyISO(), desde7 = masDias(hoy, -6);
  var tb = tableroKpiAreas(SUC_KPI.concat([AREA_CONTA, AREA_MKT, AREA_TEC, AREA_COO]), 14);
  var regs = registrosKpi(5000, true), activos = filasKpi().filter(function (k) { return k.estado === KPI_EST.ENV; });
  var prom = function (l) {
    l = l.filter(function (x) { return x != null; });
    return l.length ? Math.round(l.reduce(function (a, b) { return a + b; }, 0) / l.length) : null;
  };
  function armar(id, nombre, ids, detalle) {
    var serie = tb.fechas.map(function (f, i) {
      return prom(ids.map(function (a) { return tb.cumplimiento[a][i]; }));
    });
    var pct = prom(serie.slice(-8)), antes = prom(serie.slice(0, -8));
    var nombres = ids.map(nombreArea);
    var fuera = regs.filter(function (x) { return x.enRango === 'No' && x.fecha >= desde7 && nombres.indexOf(x.sucursal) >= 0; }).length;
    var total = activos.filter(function (k) { return kpiEnAreas(k, ids); }).length;
    return { id: id, nombre: nombre, pct: pct, antes: antes, nivel: nivelPct(pct), serie: serie,
      indicadores: total, fuera: fuera, extras: [], detalle: detalle || [] };
  }
  var detOps = SUC_KPI.map(function (s) {
    return { id: s, nombre: nombreUnidad(s), pct: prom(tb.cumplimiento[s].slice(-8)) };
  });
  return { hoy: hoy, fechas: tb.fechas, ms: { base: Date.now() - t0, total: Date.now() - t0 }, areas: [
    armar('operaciones', 'Operaciones', SUC_KPI, detOps),
    armar('finanzas', 'Finanzas', [AREA_CONTA]),
    armar('marketing', 'Marketing', [AREA_MKT]),
    armar('tecnologia', 'Tecnología', [AREA_TEC]),
    armar('direccion', 'Dirección operativa', [AREA_COO])
  ] };
}

/** Las señales de cada área (lo lento: lee inventario, impuestos, mantenimiento, presupuesto). Cambian poco: se guardan 10 minutos.
 *  Devuelve también cuánto tardó cada una, para saber cuál optimizar. */
function resumenAreasSenales(cred, fresco) {
  exigeAdmin(cred);
  var cc = null; try { cc = CacheService.getScriptCache(); } catch (e) {}
  if (cc && !fresco) { try { var t = cc.get('resumen_senales_v1'); if (t) return JSON.parse(t); } catch (e) {} }
  var r = resumenSenalesCalc();
  if (cc) { try { var s = JSON.stringify(r); if (s.length < 90000) cc.put('resumen_senales_v1', s, 600); } catch (e) {} }
  return r;
}
function resumenSenalesCalc() {
  var t0 = Date.now(), hoy = hoyISO(), tiempos = {};
  function seguro(nombre, f) {
    var a = Date.now();
    try { return f(); } catch (e) { return null; } finally { tiempos[nombre] = Date.now() - a; }
  }
  var extra = function (k, v, nivel) { return { k: k, v: String(v), nivel: nivel || '' }; };
  var ops = [], fin = [], mkt = [];

  var inv = seguro('inventario', function () {
    var aj = ajustesInv(), al = 0;
    SUC_KPI.forEach(function (u) {
      var ult = estadoInventario(u).ultimo, cal = calendarioConteo(u, ult ? ult.fecha : '', aj, hoy);
      if (!cal.atrasado) al++;
    });
    return al;
  });
  if (inv != null) ops.push(extra('Inventarios al día', inv + ' de ' + SUC_KPI.length, inv === SUC_KPI.length ? 'ok' : inv === 0 ? 'mal' : 'alerta'));
  var mant = seguro('mantenimiento', function () { return resumenMant(); });
  if (mant) ops.push(extra('Mantenimientos vencidos', mant.vencidos, mant.vencidos ? 'mal' : 'ok'));

  var imp = seguro('impuestos', function () { return datosImpuestos().periodos; });
  if (imp) {
    var venc = imp.filter(function (x) { return x.estado === 'Vencido'; }).length;
    var pv = imp.filter(function (x) { return x.estado === 'Por vencer' && x.dias <= 7; }).length;
    fin.push(extra('Impuestos vencidos', venc, venc ? 'mal' : 'ok'));
    if (pv) fin.push(extra('Vencen en 7 días', pv, 'alerta'));
  }
  var porVal = seguro('porValidar', function () { return filasPorConfirmar().filter(function (x) { return x.estado === EST_PEND; }).length; });
  if (porVal != null) fin.push(extra('Registros por validar', porVal, porVal > 10 ? 'alerta' : porVal ? '' : 'ok'));

  var av = seguro('presupuesto', function () {
    var ids = {}, mejor = null;
    usuariosCache().forEach(function (u) { if (u.activo && u.rol === 'cmo' && u.sucursal) ids[u.sucursal] = true; });
    Object.keys(ids).forEach(function (id) { var a = avanceArea(id, hoy.slice(0, 7)); if (a && a.pct != null && (!mejor || a.presupuesto > mejor.presupuesto)) mejor = a; });
    return mejor;
  });
  if (av) mkt.push(extra('Presupuesto gastado', av.pct + ' %', av.pct > 100 ? 'mal' : av.pct > 85 ? 'alerta' : 'ok'));

  tiempos.total = Date.now() - t0;
  return { hoy: hoy, areas: { operaciones: ops, finanzas: fin, marketing: mkt, tecnologia: [] }, ms: tiempos };
}

/* ════════════ PUERTA INTERNA PARA EL SERVIDOR PROPIO ════════════
 * El servidor nuevo (Render) necesita saber quién es quién para reconocer a cada persona sin preguntarle a Google.
 * Se lo cuenta esta puerta, que SOLO abre con la clave secreta «SERVER_TOKEN» guardada en las propiedades del script
 * (Configuración del proyecto → Propiedades del script). La clave nunca está escrita en el código.
 * Aquí viajan los PIN: por eso es una puerta aparte, con clave larga, y el servidor nuevo los guarda ya transformados
 * (nunca en texto legible). */
var INTERNAS = ['exportarUsuarios', 'exportarIndicadores', 'exportarResumenLibro', 'exportarHojaTipada', 'versionCodigo'];
function igualesSeguro(a, b) {
  a = String(a); b = String(b);
  if (a.length !== b.length) return false;
  var d = 0;
  for (var i = 0; i < a.length; i++) d |= a.charCodeAt(i) ^ b.charCodeAt(i);
  return d === 0;
}
function interno(fn, q) {
  var token = '';
  try { token = PropertiesService.getScriptProperties().getProperty('SERVER_TOKEN') || ''; } catch (e) {}
  // sin clave puesta (o demasiado corta) la puerta queda cerrada para todos
  if (token.length < 24 || !igualesSeguro(q && q.token || '', token)) throw new Error('No autorizado.');
  if (INTERNAS.indexOf(fn) < 0) throw new Error('No existe.');
  return globalThis[fn](q);
}
/** La lista completa de usuarios. Si el servidor ya tiene la última versión (misma «huella»), solo contesta «sin cambios». */
function exportarUsuarios(q) {
  extiendeLugares();                      // los lugares creados en la app (oficinas) valen como «sucursal» de un usuario
  var us = filasUsuarios().map(function (u) {
    return { nombre: u.nombre, pin: u.pin, rol: u.rol, activo: u.activo, sucursal: u.sucursal, confirma: u.confirma, correo: u.correo };
  });
  var huella = Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, JSON.stringify(us)));
  if (q && q.huella && String(q.huella) === huella) return { sinCambios: true, huella: huella };
  return { huella: huella, usuarios: us };
}

/** Los indicadores para el servidor propio: la lista (solo si cambió desde la última huella) y los registros desde la fila que pida,
 *  por páginas de hasta 3000. Así el servidor mantiene una copia al día sin releer todo cada vez. */
function exportarIndicadores(q) {
  q = q || {};
  var lista = filasKpi().map(function (k) { delete k.fila; return k; });
  var huella = Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, JSON.stringify(lista)));
  var h = hojaKpiReg(), ultima = h.getLastRow();
  var ancho = Math.min(H_KPIREG.length, h.getMaxColumns());
  var pagina = Math.max(1, Math.min(3000, Number(q.max) || 3000));
  var desde = Math.max(2, Number(q.desdeFila) || 2), tope = Math.min(ultima, desde + pagina - 1);
  var filas = [];
  if (desde <= ultima) {
    filas = h.getRange(desde, 1, tope - desde + 1, ancho).getValues().map(function (r, i) {
      return { fila: desde + i, en: fmtSello(r[0]), fecha: fmtDia(r[1]), turno: String(r[2] || ''),
        sucursal: String(r[3] || ''), id: String(r[4] || ''), nombre: String(r[5] || ''),
        valor: r[6] === '' || r[6] == null ? '' : r[6], unidad: String(r[7] || ''),
        min: numONulo(r[8]), max: numONulo(r[9]), enRango: String(r[10] || ''),
        por: String(r[11] || ''), nota: String(r[12] || ''), per: String(r[13] || '') };
    });
  }
  var out = { huella: huella, ultimaFila: ultima, desde: desde, hasta: tope, hayMas: tope < ultima, registros: filas };
  if (String(q.huella || '') !== huella) out.indicadores = lista;
  return out;
}

/* ════════════ COPIA COMPLETA PARA EL SERVIDOR PROPIO ════════════
 * El servidor propio guarda una copia de TODAS las pestañas y corre este mismo código sobre ella para contestar las lecturas.
 * Cada cambio que pasa por aquí avisa qué pestañas tocó («tocadas»), y el servidor solo vuelve a traer esas. */
var VERSION_CODIGO = '2026-10-14-a';       // se cambia a mano cada vez que se cambia este archivo; el servidor compara que coincida con la suya
function versionCodigo() { return VERSION_CODIGO; }
function huellaTexto(t) { return Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, t)).slice(0, 22); }
/** Un resumen barato de cada pestaña: cuántas filas y columnas tiene y una huella de sus últimas 25 filas. */
function exportarResumenLibro() {
  var ss = libro();
  return { hojas: ss.getSheets().map(function (h) {
    var f = h.getLastRow(), c = h.getLastColumn(), cola = '';
    if (f > 0 && c > 0) { var d = Math.max(1, f - 24); cola = huellaTexto(JSON.stringify(h.getRange(d, 1, f - d + 1, c).getValues())); }
    return { nombre: h.getName(), filas: f, cols: c, huella: cola };
  }), disparadores: disparadoresActivos(), propiedades: propiedadesVisibles() };
}
/** El contenido de una pestaña, con las fechas marcadas para no perder su tipo; por páginas de hasta 2000 filas. */
function exportarHojaTipada(q) {
  q = q || {};
  var h = libro().getSheetByName(String(q.nombre || ''));
  if (!h) return { existe: false };
  var f = h.getLastRow(), c = h.getLastColumn();
  var pagina = Math.max(1, Math.min(2000, Number(q.max) || 2000)), desde = Math.max(1, Number(q.desde) || 1), hasta = Math.min(f, desde + pagina - 1);
  var celdas = [];
  if (f > 0 && c > 0 && desde <= f) celdas = h.getRange(desde, 1, hasta - desde + 1, c).getValues().map(function (fila) {
    return fila.map(function (v) { return v instanceof Date ? { $d: v.getTime() } : v; });
  });
  var textoCols = [];
  if (c > 0) h.getRange(Math.min(2, Math.max(1, f)), 1, 1, c).getNumberFormats()[0].forEach(function (x, i) { if (x === '@') textoCols.push(i + 1); });
  return { existe: true, nombre: h.getName(), filas: f, cols: c, maxCols: h.getMaxColumns(), desde: desde, hasta: hasta, hayMas: hasta < f, celdas: celdas, textoCols: textoCols };
}
/** Las tareas programadas que ya tiene el proyecto (el servidor las necesita para saber que ya están puestas). */
function disparadoresActivos() { try { return ScriptApp.getProjectTriggers().map(function (t) { return t.getHandlerFunction(); }); } catch (e) { return []; } }
/** Las propiedades del script (marcas como «ya se corrigió tal cosa»), sin la clave secreta. El servidor las necesita para correr este código igual que aquí. */
function propiedadesVisibles() {
  try { var p = PropertiesService.getScriptProperties().getProperties(); delete p.SERVER_TOKEN; return p; } catch (e) { return {}; }
}

/* ════════════ INDICADORES DEL DIRECTOR OPERATIVO ════════════
 * Los decide el administrador (como los de la CMO: quedan en borrador hasta que los envía) y los anota el director operativo en
 * «Mis KPIs». El director sigue decidiendo y enviando los de sus gerentes por su lado. */
var KPI_CATS_COO = ['Bodega central', 'Producción', 'Sucursales', 'Equipo'];
var KPIS_COO_RECOMENDADOS = [
  { c: 'Sucursales', n: 'Sucursales con todos sus indicadores anotados en la semana', t: 'Número', u: '%', min: 90, max: null, f: 'Semanal',
    i: 'De los indicadores que tocaban en la semana, el porcentaje que los gerentes sí anotaron (lo ve en Resultados).', p: 'Si no anotan, no se sabe cómo va la operación.' },
  { c: 'Sucursales', n: 'Cierres de caja validados el mismo día', t: 'Número', u: '%', min: 95, max: null, f: 'Semanal',
    i: 'De los cierres de la semana, el porcentaje que validó el mismo día en que se reportó.', p: 'Un cierre sin validar esconde diferencias de efectivo.' },
  { c: 'Sucursales', n: 'Mantenimientos vencidos en las sucursales', t: 'Número', u: 'equipos', min: null, max: 0, f: 'Semanal',
    i: 'Cuántos equipos tienen el mantenimiento vencido al cerrar la semana (lo ve en Mantenimiento).', p: 'Un equipo sin mantenimiento termina parando el servicio.' },
  { c: 'Sucursales', n: 'Faltantes de producto en sucursales (veces en la semana)', t: 'Número', u: 'veces', min: null, max: 3, f: 'Semanal',
    i: 'Cuántas veces una sucursal se quedó sin un producto importante, según lo que reportaron los gerentes.', p: 'Los faltantes son ventas que se pierden.' },
  { c: 'Bodega central', n: 'Solicitudes de sucursales atendidas a tiempo por bodega', t: 'Número', u: '%', min: 95, max: null, f: 'Semanal',
    i: 'De las solicitudes de la semana, el porcentaje que bodega entregó en el plazo acordado.', p: 'Si bodega se atrasa, las sucursales piden emergencias.' },
  { c: 'Bodega central', n: 'Diferencias de bodega sin aclarar', t: 'Número', u: 'diferencias', min: null, max: 0, f: 'Semanal',
    i: 'Cuántas diferencias de inventario de bodega siguen sin explicación al cerrar la semana (lo ve en Diferencias).', p: 'Lo que no se aclara se repite.' },
  { c: 'Bodega central', n: 'Compras recibidas completas', t: 'Número', u: '%', min: 95, max: null, f: 'Semanal',
    i: 'De las órdenes recibidas en la semana, el porcentaje que llegó completo y en buen estado.', p: 'Una compra incompleta afecta a todas las sucursales.' },
  { c: 'Producción', n: 'Lotes con rendimiento dentro de lo esperado', t: 'Número', u: '%', min: 90, max: null, f: 'Semanal',
    i: 'De los lotes de la semana, el porcentaje cuyo rendimiento quedó dentro de la receta (lo ve en Producción).', p: 'Un mal rendimiento es dinero que se pierde en el insumo.' },
  { c: 'Producción', n: 'Merma de producción en la semana', t: 'Número', u: '%', min: null, max: 5, f: 'Semanal',
    i: 'Lo que se perdió de los insumos producidos, como porcentaje de lo producido.', p: 'Ayuda a ver si la producción cuida los insumos.' },
  { c: 'Equipo', n: 'Reunión semanal con los gerentes realizada', t: 'Sí / No', u: '', min: null, max: null, f: 'Semanal',
    i: 'Marque «Sí» si se reunió con los tres gerentes esta semana.', p: 'La reunión es donde se corrigen los problemas antes de que crezcan.' },
  { c: 'Equipo', n: 'Indicadores enviados a los gerentes al día', t: 'Sí / No', u: '', min: null, max: null, f: 'Semanal',
    i: 'Marque «Sí» si los indicadores nuevos o cambiados ya se enviaron a los gerentes.', p: 'Un indicador sin enviar es una orden que nadie recibió.' }
];
function exigeAdminCoo(cred) { var yo = exigeAdmin(cred); _KPI_COO = true; return yo; }
function listaKpisCoo(cred) { exigeAdminCoo(cred); return listaKpis(cred); }
function guardarKpiCoo(cred, k) { exigeAdminCoo(cred); return guardarKpi(cred, k); }
function borrarKpiCoo(cred, id) { exigeAdminCoo(cred); return borrarKpi(cred, id); }
function enviarKpisCoo(cred, ids) { exigeAdminCoo(cred); return enviarKpis(cred, ids); }
function agregarRecomendadosCoo(cred, claves) { exigeAdminCoo(cred); return agregarRecomendados(cred, claves); }
function resultadosKpiCoo(cred, desde, hasta) { exigeAdminCoo(cred); return resultadosKpiRango(cred, desde, hasta, ''); }
/** Lo que anota el director operativo: sus propios indicadores. */
function exigeCoo(cred) {
  var yo = quien(cred);
  if (yo.rol !== 'operaciones') throw new Error('Esto es para el director operativo.');
  _KPI_COO = true;
  return yo;
}
function pantallaKpiCoo(cred) { exigeCoo(cred); return pantallaKpi(cred); }
function registrarKpisCoo(cred, valores) { exigeCoo(cred); return registrarKpis(cred, valores); }

/* ════════════ RECETAS POR PORCENTAJE (MEZCLAS) Y ETIQUETAS ════════════
 * Una receta de varios ingredientes (masa de pizza: harina, levadura, sal, pimienta). La base (la harina) es el 100 %; cada
 * ingrediente es un % de su peso. Paulino dice cuántas unidades de la base usa (15 sacos) y la app calcula cuánto sale de
 * bodega de cada ingrediente, cuántas porciones deben salir (8 bolsas por saco = 120) y arma las etiquetas de cada porción. */
var H_MEZ = ['ID', 'Nombre', 'Base (código)', 'Base', 'Peso de cada unidad de la base', 'Medida del peso', 'Producto que sale (código)',
  'Producto que sale', 'Presentaciones por unidad de base', 'Vida útil (días)', 'Activa', 'Creada por', 'Creada en', 'Nota', 'Ingredientes (JSON)', 'Porciones por presentación'];
var H_LOTEQ = ['No. de lote', 'Receta', 'Nombre', 'Fecha de producción', 'Vida útil (días)', 'Vence', 'Peso por porción (g)', 'Ingredientes (JSON)', 'Sobrante de la base (g)'];
var GRAMOS = { libra: 453.59237, onza: 28.349523125, kilo: 1000, gramo: 1 };
var ETQ_DEF = { hoja: 'carta', ancho: 6.67, alto: 2.54, margenSup: 1.27, margenIzq: 0.48, sepH: 0.32, sepV: 0, bordes: false, barras: true };
var HOJAS_ETQ = { carta: { ancho: 21.59, alto: 27.94 }, oficio: { ancho: 21.59, alto: 33.02 } };

function hojasMezcla(ss) {
  ss = ss || libro();
  var m = ss.getSheetByName('Recetas por porcentaje');
  if (!m) { m = hojaLimpia(ss, 'Recetas por porcentaje', H_MEZ); m.getRange('M:M').setNumberFormat('dd/mm/yyyy hh:mm'); m.setColumnWidth(2, 240); m.setColumnWidth(15, 420); }
  else if (m.getLastColumn() < 16) encabezaAlFinal(m, 16, ['Porciones por presentación']);
  var e = ss.getSheetByName('Producción lotes · etiquetas');
  if (!e) { e = hojaLimpia(ss, 'Producción lotes · etiquetas', H_LOTEQ); e.getRange('D:D').setNumberFormat('@'); e.getRange('F:F').setNumberFormat('@'); e.setColumnWidth(8, 420); }
  else if (e.getLastColumn() < 9) encabezaAlFinal(e, 9, ['Sobrante de la base (g)']);
  return { rec: m, eq: e };
}
function jsonSeguro(t) { try { var v = JSON.parse(String(t || '')); return Array.isArray(v) ? v : []; } catch (e) { return []; } }
function filasMezclas() {
  return leeTodo(hojasMezcla().rec, H_MEZ.length).map(function (r, i) {
    return { fila: i + 2, id: String(r[0] || ''), nombre: String(r[1] || ''), base: String(r[2] || ''), baseNombre: String(r[3] || ''),
      basePeso: Number(r[4]) || 0, baseMedida: String(r[5] || '').trim().toLowerCase() || 'libra', salida: String(r[6] || ''), salidaNombre: String(r[7] || ''),
      porciones: Number(r[8]) || 0, vida: Number(r[9]) || 0, activa: String(r[10]).trim().toLowerCase() !== 'no',
      creadaPor: String(r[11] || ''), creadaEn: fmtSello(r[12]), nota: String(r[13] || ''), ingredientes: jsonSeguro(r[14]), porPres: Number(r[15]) > 0 ? Number(r[15]) : 1 };
  }).filter(function (x) { return x.id; });
}
function filasLotesEtq() {
  var o = {};
  leeTodo(hojasMezcla().eq, H_LOTEQ.length).forEach(function (r, i) {
    var n = String(r[0] || ''); if (!n) return;
    o[n] = { fila: i + 2, numero: n, receta: String(r[1] || ''), nombre: String(r[2] || ''), fecha: fmtDia(r[3]), vida: Number(r[4]) || 0,
      vence: fmtDia(r[5]), pesoPorcion: Number(r[6]) || 0, ingredientes: jsonSeguro(r[7]), sobrante: Number(r[8]) || 0 };
  });
  return o;
}
/** Cuánto pesa una unidad de bodega de un producto (si se sabe): {peso, medida}. */
function pesoUnidadDe(b) {
  if (!b) return null;
  if (b.medida === 'peso' && GRAMOS[b.unidad]) return { peso: 1, medida: b.unidad };
  if (b.pesoCada > 0 && GRAMOS[b.pesoEn]) {          // «cada una cerrada pesa»: se refiere a la unidad más pequeña que se cuenta
    var k = (b.menor && b.porSuelto >= 2 && b.suelto && b.porUnidad >= 2) ? b.porUnidad * b.porSuelto : (b.suelto && b.porUnidad >= 2) ? b.porUnidad : 1;
    return { peso: r3(b.pesoCada * k), medida: b.pesoEn };
  }
  return null;
}
function numPos(v) { var n = Number(String(v == null ? '' : v).replace(',', '.')); return isFinite(n) ? n : 0; }
/** Lo que pide una receta para `unidades` de la base. */
function calculoMezcla(m, unidades, bod) {
  var gBase = unidades * m.basePeso * GRAMOS[m.baseMedida];
  var ings = m.ingredientes.map(function (x) {
    var g = gBase * x.pct / 100, gU = x.peso * GRAMOS[x.medida], b = bod[x.codigo] || {};
    return { codigo: x.codigo, nombre: b.nombre || x.nombre, pct: x.pct, g: Math.round(g * 100) / 100, unidades: gU > 0 ? Math.round(g / gU * 10000) / 10000 : 0,
      unidad: b.unidad || x.unidad || 'unidad' };
  });
  var esperadas = Math.round(unidades * m.porciones * 1000) / 1000;
  var gTot = gBase + ings.reduce(function (a, x) { return a + x.g; }, 0);
  return { gBase: Math.round(gBase * 10) / 10, ingredientes: ings, esperadas: esperadas, pesoPorcion: esperadas > 0 ? Math.round(gTot / esperadas * 10) / 10 : 0 };
}
function mezclasPantalla(ex, idx, bod) {
  return filasMezclas().filter(function (m) { return m.activa; }).map(function (m) {
    var b = bod[m.base] || {}, s = bod[m.salida] || {}, uno = calculoMezcla(m, 1, bod);
    var costo = (b.codigo ? costoBodega(b, idx).costo : 0) + uno.ingredientes.reduce(function (a, x) { var y = bod[x.codigo]; return a + (y ? costoBodega(y, idx).costo * x.unidades : 0); }, 0);
    return { id: m.id, nombre: m.nombre, base: m.base, baseNombre: b.nombre || m.baseNombre, baseUnidad: b.unidad || 'unidad', basePeso: m.basePeso, baseMedida: m.baseMedida,
      salida: m.salida, salidaNombre: s.nombre || m.salidaNombre, salidaUnidad: s.unidad || 'bolsa', porciones: m.porciones, porPres: m.porPres, vida: m.vida, nota: m.nota,
      hayBase: r3(ex[m.base] || 0), hayBaseTxt: b.codigo ? textoCantidad(ex[m.base] || 0, b) : '—',
      ingredientes: m.ingredientes.map(function (x) { var y = bod[x.codigo] || {};
        return { codigo: x.codigo, nombre: y.nombre || x.nombre, pct: x.pct, peso: x.peso, medida: x.medida, unidad: y.unidad || x.unidad || 'unidad',
          hay: r3(ex[x.codigo] || 0), hayTxt: y.codigo ? textoCantidad(ex[x.codigo] || 0, y) : '—' }; }),
      porUnidad: uno, costoPorUnidad: Math.round(costo * 100) / 100,
      costoPorcion: uno.esperadas > 0 ? Math.round(costo / uno.esperadas * 10000) / 10000 : 0 };
  });
}
/** Receta por % nueva o cambiada. r = {id, nombre, base, basePeso, baseMedida, salida, salidaNombre, porciones, vida, nota, ingredientes:[{codigo, pct, peso, medida}]} */
function guardarMezcla(cred, r) {
  var yo = exigeProduccion(cred);
  r = r || {};
  var bod = {}; filasBodega().forEach(function (b) { bod[b.codigo] = b; });
  var nombre = String(r.nombre || '').trim().replace(/\s+/g, ' ').slice(0, 70);
  if (nombre.length < 3) throw new Error('Escriba el nombre de la receta (por ejemplo «Mezcla para masa de pizza»).');
  var base = bod[String(r.base || '')];
  if (!base) throw new Error('Escoja el ingrediente base (el 100 %, por ejemplo la harina).');
  var basePeso = numPos(r.basePeso), baseMedida = String(r.baseMedida || '').toLowerCase();
  if (!(basePeso > 0) || !GRAMOS[baseMedida]) throw new Error('Diga cuánto pesa cada ' + base.unidad + ' de ' + base.nombre + ' (por ejemplo 50 libras).');
  var porciones = numPos(r.porciones);          // presentaciones (bolsas con etiqueta) que salen de cada unidad de la base
  if (!(porciones > 0 && porciones <= 10000)) throw new Error('Diga cuánto lleva cada porción de la base (por ejemplo 500 g).');
  var porPres = Math.round(numPos(r.porPres)) || 1;
  if (!(porPres >= 1 && porPres <= 1000)) throw new Error('Las porciones por presentación van de 1 a 1000.');
  var vida = Math.round(numPos(r.vida));
  if (!(vida >= 1 && vida <= 730)) throw new Error('Diga cuántos días dura el producto (entre 1 y 730).');
  var lista = Array.isArray(r.ingredientes) ? r.ingredientes : [];
  if (!lista.length) throw new Error('Agregue al menos un ingrediente además de la base.');
  if (lista.length > 20) throw new Error('Una receta puede tener hasta 20 ingredientes.');
  var vistos = {}; vistos[base.codigo] = true;
  var ings = lista.map(function (x) {
    var b = bod[String(x.codigo || '')];
    if (!b) throw new Error('Hay un ingrediente sin escoger.');
    if (vistos[b.codigo]) throw new Error(b.codigo === base.codigo ? b.nombre + ' ya es la base de la receta.' : b.nombre + ' está dos veces.');
    vistos[b.codigo] = true;
    var pct = Math.round(numPos(x.pct) * 100000) / 100000;
    if (!(pct > 0 && pct <= 500)) throw new Error('Revise el porcentaje de ' + b.nombre + '.');
    var peso = numPos(x.peso), medida = String(x.medida || '').toLowerCase();
    if (!(peso > 0) || !GRAMOS[medida]) throw new Error('Diga cuánto pesa cada ' + b.unidad + ' de ' + b.nombre + ' (lo que sale de bodega se descuenta en ' + plur(2, b.unidad) + ').');
    return { codigo: b.codigo, nombre: b.nombre, unidad: b.unidad, pct: pct, peso: peso, medida: medida };
  });
  var codigo = String(r.salida || '').trim().toUpperCase(), nombreS = String(r.salidaNombre || '').trim().replace(/\s+/g, ' ').slice(0, 70);
  if (vistos[codigo]) throw new Error('Lo que sale no puede ser uno de los ingredientes.');
  var lock = LockService.getScriptLock(); lock.waitLock(15000); _LEE = {};
  var id;
  try {
    if (!bod[codigo]) {
      if (nombreS.length < 3) throw new Error('Escriba el nombre de lo que sale (por ejemplo «Bolsa de mezcla para masa»).');
      codigo = codigoNuevo(nombreS, bod);
      hojaBodega().appendRow([codigo, nombreS.toUpperCase(), 'Cocina', 'Producción', 'Sí', 'Sí', 'Sí', 'Insumos de cocina',
        'bolsa', '', '', '', '', 1, '', 'Sí', new Date(), yo.nombre, almacenDe('Producción'), '', '', '', '', '', 'No', '']);
      _BOD = null; filasBodega().forEach(function (b) { bod[b.codigo] = b; });
    }
    var h = hojasMezcla().rec, x = null;
    filasMezclas().forEach(function (y) { if (y.id === String(r.id || '')) x = y; });
    id = x ? x.id : 'RM' + Utilities.formatDate(new Date(), ZONA, 'yyyyMMddHHmmss');
    var fila = [id, nombre, base.codigo, base.nombre, basePeso, baseMedida, codigo, bod[codigo] ? bod[codigo].nombre : nombreS.toUpperCase(), porciones, vida,
      r.activa === false ? 'No' : 'Sí', x ? x.creadaPor : yo.nombre, x ? x.creadaEn : new Date(), String(r.nota || '').slice(0, 200), JSON.stringify(ings), porPres];
    if (x) h.getRange(x.fila, 1, 1, H_MEZ.length).setValues([fila]); else h.appendRow(fila);
  } finally { lock.releaseLock(); }
  var d = pantallaProduccion(cred); d.ok = true; d.mensaje = 'Receta «' + nombre + '» guardada.'; d.mezcla = id;
  return d;
}
function gramosTxt(g) { return g >= 1000 ? num(Math.round(g / 10) / 100) + ' kg' : num(Math.round(g * 10) / 10) + ' g'; }
/** Un lote de una receta por %: d = {receta, unidades (los sacos), vida, nota, sinExistencia}.
 *  Salen solo las presentaciones COMPLETAS. Los sacos salen enteros de bodega y la harina que sobra queda anotada (en gramos) en el lote.
 *  Los demás ingredientes salen de bodega solo por lo que llevan las presentaciones completas; lo de la parte incompleta también se anota. */
function registrarLoteMezcla(cred, d) {
  var yo = exigeProduccion(cred);
  d = d || {};
  var m = null; filasMezclas().forEach(function (y) { if (y.id === String(d.receta)) m = y; });
  if (!m || !m.activa) throw new Error('Escoja la receta.');
  var bod = {}; filasBodega().forEach(function (b) { bod[b.codigo] = b; });
  var base = bod[m.base], sal = bod[m.salida];
  if (!base || !sal) throw new Error('La base o lo que sale ya no están en el catálogo de bodega.');
  var sacos = Math.round(numPos(d.unidades) * 100000) / 100000;
  if (!(sacos > 0)) throw new Error('Escriba cuántas ' + plur(2, base.unidad) + ' de ' + base.nombre + ' va a usar.');
  var exactas = sacos * m.porciones, completas = Math.floor(exactas + 0.001);          // 11.9998 (por redondeo de los sacos) cuenta como 12
  if (completas < 1) throw new Error('Con ' + num(sacos) + ' ' + plur(sacos, base.unidad) + ' no alcanza para una ' + sal.nombre + ' completa (salen ' + num(Math.round(exactas * 100) / 100) + '). Use más ' + plur(2, base.unidad) + '.');
  var usado = completas / m.porciones;                                     // la parte de los sacos que entra en las presentaciones completas
  var c = calculoMezcla(m, usado, bod), todo = calculoMezcla(m, sacos, bod);
  var sobranteG = Math.max(0, Math.round((todo.gBase - c.gBase) * 10) / 10);
  var vida = String(d.vida == null ? '' : d.vida).trim() === '' ? m.vida : Math.round(numPos(d.vida));
  if (!(vida >= 1 && vida <= 730)) throw new Error('La vida útil va de 1 a 730 días.');
  var nota = String(d.nota || '').trim().slice(0, 200);
  var ex = estadoInventario('bodega').existencia, falta = [];
  [{ codigo: base.codigo, unidades: sacos }].concat(c.ingredientes).forEach(function (x) {
    var b = bod[x.codigo]; if (b && (ex[x.codigo] || 0) < x.unidades - 0.0005) falta.push(b.nombre + ': hay ' + textoCantidad(ex[x.codigo] || 0, b) + ' y se usan ' + textoCantidad(x.unidades, b));
  });
  if (falta.length && !d.sinExistencia) throw new Error('No alcanza en bodega. ' + falta.join(' · ') + '. Revise las cantidades o pida que cuenten la bodega.');
  var idx = indiceCompras(), costo = costoBodega(base, idx).costo * sacos;           // los sacos salen enteros: su costo completo va a este lote
  c.ingredientes.forEach(function (x) { var b = bod[x.codigo]; if (b) costo += costoBodega(b, idx).costo * x.unidades; });
  costo = Math.round(costo * 100) / 100;
  var costoPor = Math.round(costo / completas * 10000) / 10000;
  var pesoPres = Math.round((c.gBase + c.ingredientes.reduce(function (a, x) { return a + x.g; }, 0)) / completas * 10) / 10;
  var hoy = hoyISO(), vence = masDias(hoy, vida), lock = LockService.getScriptLock(); lock.waitLock(15000); _LEE = {};
  var numero;
  try {
    var max = 0;
    leeTodo(hojasProduccion().lote, 1).forEach(function (r) { var mm = /^PR-(\d+)$/.exec(String(r[0])); if (mm) max = Math.max(max, +mm[1]); });
    numero = 'PR-' + ('0000' + (max + 1)).slice(-4);
    var ahora = new Date(), hl = hojasProduccion().lote;
    hl.appendRow([numero, ahora, hoy, m.id, base.codigo, base.nombre, sacos, base.unidad, sacos, sal.codigo, sal.nombre, completas, completas, 100, costo, costoPor, yo.nombre, nota]);
    hl.getRange(hl.getLastRow(), 3).setNumberFormat('@').setValue(hoy);
    var he = hojasMezcla().eq;
    he.appendRow([numero, m.id, m.nombre, hoy, vida, vence, pesoPres, JSON.stringify([{ codigo: base.codigo, nombre: base.nombre, g: c.gBase, sobra: sobranteG }].concat(c.ingredientes.map(function (x, i) {
      return { codigo: x.codigo, nombre: x.nombre, g: x.g, sobra: Math.round((todo.ingredientes[i].g - x.g) * 10) / 10 }; }))), sobranteG]);
    he.getRange(he.getLastRow(), 4).setNumberFormat('@').setValue(hoy); he.getRange(he.getLastRow(), 6).setNumberFormat('@').setValue(vence);
    escribeKardex([[ahora, nombreBodega(), base.codigo, base.nombre, MOV.SAL_PROD, -sacos, numero, yo.nombre, m.nombre + (sobranteG > 0.5 ? ' · sobran ' + gramosTxt(sobranteG) : '')]]
      .concat(c.ingredientes.map(function (x) { return [ahora, nombreBodega(), x.codigo, x.nombre, MOV.SAL_PROD, -x.unidades, numero, yo.nombre, m.nombre + ' · ' + num(x.g) + ' g']; }))
      .concat([[ahora, nombreBodega(), sal.codigo, sal.nombre, MOV.ENT_PROD, completas, numero, yo.nombre, num(completas * m.porPres) + ' porciones']]));
    if (costoPor > 0) { hojaBodega().getRange(sal.fila, 15).setValue(costoPor); _BOD = null; }
  } finally { lock.releaseLock(); }
  var r = pantallaProduccion(cred); r.ok = true; r.lote = numero; r.sobrante = sobranteG;
  r.mensaje = numero + ': ' + num(completas) + ' ' + plur(completas, sal.unidad) + ' de ' + sal.nombre + (m.porPres > 1 ? ' (' + num(completas * m.porPres) + ' porciones)' : '') +
    (sobranteG > 0.5 ? '. Sobran ' + gramosTxt(sobranteG) + ' de ' + base.nombre + ' (anotados en el lote)' : '') + '. Vence el ' + dia(vence) + '.';
  return r;
}
/** Lo que va en las etiquetas de un lote (de receta por % o de porcionar). Con `dias`, se cambia su vida útil. */
function etiquetasLote(cred, numero, dias) {
  var yo = quien(cred);
  if (!veProduccion(yo)) throw new Error('Su usuario no ve producción.');
  var l = null; filasLotes().forEach(function (x) { if (x.numero === String(numero)) l = x; });
  if (!l) throw new Error('No se encontró el lote ' + numero + '.');
  var bod = {}; filasBodega().forEach(function (b) { bod[b.codigo] = b; });
  var eq = filasLotesEtq()[l.numero], sal = bod[l.porcion] || {};
  var vida = eq ? eq.vida : 0, fecha = eq ? eq.fecha : l.fecha, ings = eq ? eq.ingredientes : [{ codigo: l.insumo, nombre: l.insumoNombre, g: GRAMOS[l.medida] ? l.usada * GRAMOS[l.medida] : 0 }];
  var peso = eq ? eq.pesoPorcion : (GRAMOS[l.medida] && l.obtenidas > 0 ? Math.round(l.usada * GRAMOS[l.medida] / l.obtenidas * 10) / 10 : 0);
  if (!vida) { var mz = null; filasMezclas().forEach(function (y) { if (y.id === l.receta) mz = y; }); vida = mz ? mz.vida : 0; }
  var nueva = String(dias == null ? '' : dias).trim() === '' ? 0 : Math.round(numPos(dias));
  if (nueva && !(nueva >= 1 && nueva <= 730)) throw new Error('La vida útil va de 1 a 730 días.');
  if (nueva && nueva !== vida) {
    if (yo.rol !== 'produccion' && !yo.esAdmin) throw new Error('La vida útil la cambia producción.');
    vida = nueva;
    var he = hojasMezcla().eq, vence = masDias(fecha, vida);
    if (eq) { he.getRange(eq.fila, 5).setValue(vida); he.getRange(eq.fila, 6).setNumberFormat('@').setValue(vence); }
    else { he.appendRow([l.numero, l.receta, l.porcionNombre, fecha, vida, vence, peso, JSON.stringify(ings)]);
      he.getRange(he.getLastRow(), 4).setNumberFormat('@').setValue(fecha); he.getRange(he.getLastRow(), 6).setNumberFormat('@').setValue(vence); }
  }
  // La receta es confidencial: a la etiqueta (y a quien la imprime) solo le llega el nombre de la porción, el lote y las fechas.
  return { numero: l.numero, producto: sal.nombre || l.porcionNombre, sku: l.porcion, fecha: fecha, vida: vida, vence: vida ? masDias(fecha, vida) : '',
    cantidad: Math.round(l.obtenidas), elaboro: l.por, empresa: 'REY PIZZA · CAIX, S.A.', config: configEtiquetas(), hojas: HOJAS_ETQ, puedeCambiar: yo.rol === 'produccion' || !!yo.esAdmin };
}
function configEtiquetas() {
  var c = {}; try { c = JSON.parse(ajusteCrudo('etiquetas_config') || '{}') || {}; } catch (e) { c = {}; }
  var o = {}; Object.keys(ETQ_DEF).forEach(function (k) { o[k] = c[k] != null ? c[k] : ETQ_DEF[k]; });
  return o;
}
/** Se guarda el tamaño de las etiquetas para la próxima vez (en cm). */
function guardarConfigEtiquetas(cred, c) {
  var yo = quien(cred);
  if (yo.rol !== 'produccion' && !yo.esAdmin) throw new Error('El tamaño de las etiquetas lo configura producción.');
  c = c || {};
  var hoja = HOJAS_ETQ[c.hoja] ? c.hoja : 'carta', H = HOJAS_ETQ[hoja];
  var n = function (k, min, max) { var v = Math.round(numPos(c[k]) * 100) / 100; if (!(v >= min && v <= max)) throw new Error('Revise la medida «' + k + '» (entre ' + min + ' y ' + max + ' cm).'); return v; };
  var o = { hoja: hoja, ancho: n('ancho', 1, H.ancho), alto: n('alto', 1, H.alto), margenSup: n('margenSup', 0, 10), margenIzq: n('margenIzq', 0, 10),
    sepH: n('sepH', 0, 5), sepV: n('sepV', 0, 5), bordes: !!c.bordes, barras: !!c.barras };
  if (o.margenIzq + o.ancho > H.ancho + 0.001 || o.margenSup + o.alto > H.alto + 0.001) throw new Error('Con esas medidas no cabe ni una etiqueta en la hoja.');
  ponAjusteCrudo('etiquetas_config', JSON.stringify(o), 'Tamaño de las etiquetas de producción (cm)');
  return { ok: true, config: o, mensaje: 'Tamaño de etiquetas guardado.' };
}

/* ════════════ COMPRAS: devolver y corregir una orden ════════════
 * Contabilidad devuelve la orden a bodega (con el motivo) cuando la cotización no cuadra con lo pedido; Samuel corrige las cantidades y la
 * reenvía: vuelve a «Por cotizar» y se cotiza de nuevo, con la cotización en línea. */
function devolverOrden(cred, numero, motivo) {
  var yo = quien(cred);
  if (!(esContador(yo) || yo.esAdmin || yo.rol === 'finanzas')) throw new Error('Las órdenes las devuelve contabilidad.');
  motivo = String(motivo || '').trim();
  if (motivo.length < 5) throw new Error('Escriba por qué se devuelve (por ejemplo: la cotización no cuadra con la cantidad pedida).');
  var o = null; ordenesAgrupadas().forEach(function (x) { if (x.numero === String(numero)) o = x; });
  if (!o) throw new Error('No se encontró la orden ' + numero + '.');
  if (o.estado === 'Anulada') throw new Error('Esa orden está anulada.');
  if (o.estado !== 'Pedida') throw new Error('Esa orden ya se recibió en bodega.');
  if (o.pago === 'Pagada') throw new Error('Esa orden ya se pagó.');
  if (o.devueltaEn) throw new Error('Esa orden ya está devuelta a bodega.');
  var lock = LockService.getScriptLock(); lock.waitLock(15000); _LEE = {};
  try {
    var hp = hojaOCP(), pagos = filasOCP();
    if (pagos[o.numero]) hp.getRange(pagos[o.numero].fila, 18, 1, 3).setValues([[new Date(), yo.nombre, motivo.slice(0, 300)]]);
    else hp.appendRow([o.numero, o.proveedor, '', 0, '', '', '', '', '', '', '', '', '', '', '', '', 0, new Date(), yo.nombre, motivo.slice(0, 300), '']);
  } finally { lock.releaseLock(); }
  var r = ordenesPago(cred); r.ok = true; r.mensaje = o.numero + ' devuelta a bodega. Samuel la corrige y la reenvía.';
  return r;
}
/** Samuel corrige las cantidades de una orden devuelta y la reenvía. cantidades = { código: cantidad } */
function corregirOrden(cred, numero, cantidades) {
  var yo = quien(cred);
  if (!(yo.rol === 'bodega' || yo.esAdmin)) throw new Error('La orden la corrige la bodega.');
  var ls = filasOC().filter(function (o) { return o.numero === String(numero); });
  if (!ls.length) throw new Error('No se encontró la orden ' + numero + '.');
  if (ls[0].estado !== 'Pedida') throw new Error('Solo se corrige una orden que todavía no llega.');
  var pg = filasOCP()[String(numero)];
  if (!pg || !pg.devueltaEn) throw new Error('Esa orden no está devuelta: no hace falta corregirla.');
  if (pg.pagadaEn) throw new Error('Esa orden ya se pagó.');
  cantidades = cantidades || {};
  var lock = LockService.getScriptLock(); lock.waitLock(15000); _LEE = {};
  var cambios = 0;
  try {
    var h = hojaOC();
    ls.forEach(function (o) {
      var t = cantidades[o.codigo], v = t == null || String(t).trim() === '' ? o.pedido : Number(String(t).replace(',', '.'));
      if (!(v > 0) || !isFinite(v)) throw new Error('La cantidad de «' + (o.compra || o.nombre) + '» debe ser mayor que cero.');
      v = r3(v);
      if (v !== o.pedido) { h.getRange(o.fila, 8).setValue(v); cambios++; }
      h.getRange(o.fila, 16, 1, 2).setValues([['', '']]);              // la cotización anterior ya no vale
    });
    var hp = hojaOCP();
    hp.getRange(pg.fila, 3, 1, 4).setValues([['', '', '', '']]);      // no. de cotización, total, cotizada en, cotizada por
    hp.getRange(pg.fila, 16).setValue('');                            // el archivo de la cotización anterior
    hp.getRange(pg.fila, 18, 1, 3).setValues([['', '', '']]);         // ya no está devuelta
  } finally { lock.releaseLock(); }
  var r = listaOrdenes(cred); r.ok = true;
  r.mensaje = String(numero) + ' corregida' + (cambios ? ' (' + cambios + (cambios === 1 ? ' cantidad' : ' cantidades') + ')' : '') + ' y reenviada: ahora se cotiza de nuevo.';
  return r;
}

/* ════════════ PRECIOS: un solo correo con los cambios, y reiniciar todos ════════════ */
function avisaCambioPrecios(yo, hechos, titulo) {
  if (!hechos || !hechos.length) return [];
  var para = usuariosCache().filter(function (u) { return u.activo && correoValido(u.correo) && u.nombre !== yo.nombre &&
    (u.rol === 'operaciones' || u.rol === 'finanzas' || (u.rol === 'registro' && u.confirma)); });
  if (!para.length) return [];
  var td = 'padding:6px 8px;border-top:1px solid #eef0f3';
  var filas = hechos.slice(0, 300).map(function (h) { return '<tr><td style="' + td + '">' + esHtml(h.nombre) + '<div style="font-size:11px;color:#6c7079">por ' + esHtml(h.pres) + '</div></td>' +
    '<td style="' + td + ';text-align:right;color:#6c7079">' + (h.antes > 0 ? dinero(h.antes) : '—') + '</td><td style="' + td + ';text-align:right;font-weight:bold">' + (h.ahora > 0 ? dinero(h.ahora) : 'sin precio') + '</td></tr>'; }).join('');
  var html = htmlCorreo(titulo, '<p>' + esHtml(yo.nombre) + ' ' + (hechos.length === 1 ? 'cambió 1 precio' : 'cambió ' + hechos.length + ' precios') + ' el ' +
    Utilities.formatDate(new Date(), ZONA, 'dd/MM/yyyy HH:mm') + '.</p><table style="width:100%;border-collapse:collapse;font-size:13px"><tr><th align="left" style="padding:6px 8px">Producto</th>' +
    '<th align="right" style="padding:6px 8px">Antes</th><th align="right" style="padding:6px 8px">Ahora</th></tr>' + filas + '</table>' + (hechos.length > 300 ? '<p>…y ' + (hechos.length - 300) + ' más.</p>' : ''));
  var ok = [];
  para.forEach(function (u) { try { if (enviaCorreo(u.correo, 'Rey Pizza · ' + titulo + ' (' + hechos.length + ')', html)) ok.push(u.nombre); } catch (e) {} });
  return ok;
}
/** Borra todos los precios de bodega para registrarlos de nuevo. Solo el administrador, escribiendo REINICIAR. */
function reiniciarPrecios(cred, confirmacion) {
  var yo = exigeAdmin(cred);
  if (String(confirmacion || '').trim().toUpperCase() !== 'REINICIAR') throw new Error('Para borrar todos los precios, escriba REINICIAR.');
  var lock = LockService.getScriptLock(); lock.waitLock(15000); _LEE = {};
  var hechos = [];
  try {
    var h = hojaBodega(), ahora = new Date();
    filasBodega().forEach(function (b) {
      if (!(b.precioPres > 0)) return;
      hechos.push({ nombre: b.nombre, pres: b.presCompra || b.unidad, antes: b.precioPres, ahora: 0 });
      h.getRange(b.fila, 35, 1, 3).setValues([['', ahora, yo.nombre + ' (reinicio)']]);
    });
    _BOD = null;
  } finally { lock.releaseLock(); }
  var avisados = []; try { avisados = avisaCambioPrecios(yo, hechos, 'Se reiniciaron los precios de bodega'); } catch (e) {}
  var r = listaPrecios(cred); r.ok = true;
  r.mensaje = hechos.length + (hechos.length === 1 ? ' precio borrado' : ' precios borrados') + '. Ya se pueden registrar de nuevo.' + (avisados.length ? ' Se avisó por correo a ' + avisados.join(', ') + '.' : '');
  return r;
}

/* ════════════ EQUIPOS: inventario, responsables, revisión de cada mes y pedidos ════════════
 * El director operativo registra los equipos de cada lugar (sucursales y bodega/producción) y quién responde por ellos.
 * Cada mes los revisa con el responsable: está, falta o está dañado. El responsable confirma la revisión; lo que falta se reporta.
 * Paulino y los gerentes piden equipo; el director decide: asignar uno que ya existe, comprarlo (pasa a pagos) o rechazarlo. */
var H_EQA = ['No.', 'Lugar', 'Equipo', 'Detalle (marca, serie)', 'Cantidad', 'Responsable', 'A cargo desde', 'Estado', 'Activo', 'Agregado por', 'Agregado en', 'Nota'];
var H_EQI = ['No.', 'Lugar', 'Fecha', 'Hecho por', 'Responsable', 'Equipos', 'Están', 'Faltan', 'Dañados', 'Confirmado por', 'Confirmado en', 'Nota', 'Detalle (JSON)'];
var H_EQS = ['No.', 'Pedido en', 'Pedido por', 'Lugar', 'Qué necesita', 'Para qué', 'Estado', 'Resuelto por', 'Resuelto en', 'Cómo', 'Equipo asignado', 'Solicitud de pago', 'Nota', 'Recibido en'];
var EQA_ESTADOS = ['Bueno', 'Dañado', 'En reparación'];
var EQS_EST = { PED: 'Pedido', ASIG: 'Asignado, por recibir', COMPRA: 'En compra', REC: 'Recibido', RECH: 'Rechazado' };
var DIAS_REV_EQ = 30;
function hojasEquipos(ss) {
  ss = ss || libro();
  var a = ss.getSheetByName('Equipos');
  if (!a) { a = hojaLimpia(ss, 'Equipos', H_EQA); a.getRange('G:G').setNumberFormat('@'); a.getRange('K:K').setNumberFormat('dd/mm/yyyy hh:mm'); a.setColumnWidth(3, 220); }
  var i = ss.getSheetByName('Equipos revisiones');
  if (!i) { i = hojaLimpia(ss, 'Equipos revisiones', H_EQI); i.getRange('C:C').setNumberFormat('@'); i.getRange('K:K').setNumberFormat('dd/mm/yyyy hh:mm'); i.setColumnWidth(13, 400); }
  var s = ss.getSheetByName('Equipos pedidos');
  if (!s) { s = hojaLimpia(ss, 'Equipos pedidos', H_EQS); s.getRange('B:B').setNumberFormat('dd/mm/yyyy hh:mm'); s.getRange('I:I').setNumberFormat('dd/mm/yyyy hh:mm'); s.getRange('N:N').setNumberFormat('dd/mm/yyyy hh:mm'); }
  return { a: a, i: i, s: s };
}
function filasEqA() {
  return leeTodo(hojasEquipos().a, H_EQA.length).map(function (r, k) {
    return { fila: k + 2, numero: String(r[0] || ''), lugar: String(r[1] || ''), equipo: String(r[2] || ''), detalle: String(r[3] || ''), cantidad: Number(r[4]) || 1,
      responsable: String(r[5] || ''), desde: fmtDia(r[6]), estado: String(r[7] || 'Bueno'), activo: String(r[8]).trim().toLowerCase() !== 'no',
      agregadoPor: String(r[9] || ''), agregadoEn: fmtSello(r[10]), nota: String(r[11] || '') };
  }).filter(function (x) { return x.numero; });
}
function filasEqI() {
  return leeTodo(hojasEquipos().i, H_EQI.length).map(function (r, k) {
    var det = []; try { det = JSON.parse(String(r[12] || '[]')) || []; } catch (e) {}
    return { fila: k + 2, numero: String(r[0] || ''), lugar: String(r[1] || ''), fecha: fmtDia(r[2]), por: String(r[3] || ''), responsable: String(r[4] || ''),
      total: Number(r[5]) || 0, estan: Number(r[6]) || 0, faltan: Number(r[7]) || 0, danados: Number(r[8]) || 0,
      confirmadoPor: String(r[9] || ''), confirmadoEn: fmtSello(r[10]), nota: String(r[11] || ''), detalle: det };
  }).filter(function (x) { return x.numero; });
}
function filasEqS() {
  return leeTodo(hojasEquipos().s, H_EQS.length).map(function (r, k) {
    return { fila: k + 2, numero: String(r[0] || ''), en: fmtSello(r[1]), por: String(r[2] || ''), lugar: String(r[3] || ''), que: String(r[4] || ''), para: String(r[5] || ''),
      estado: String(r[6] || ''), resueltoPor: String(r[7] || ''), resueltoEn: fmtSello(r[8]), como: String(r[9] || ''), equipo: String(r[10] || ''),
      solicitud: String(r[11] || ''), nota: String(r[12] || ''), recibidoEn: fmtSello(r[13]) };
  }).filter(function (x) { return x.numero; });
}
function jefeEquipos(yo) { return !!yo && (yo.esAdmin || yo.rol === 'operaciones'); }
/** El lugar de cada persona: su sucursal; bodega y producción, la Bodega Central. */
function lugarDe(yo) { return yo.rol === 'gerente' ? yo.sucursal : (yo.rol === 'bodega' || yo.rol === 'produccion') ? 'bodega' : ''; }
function veEquipo(yo, x) { return jefeEquipos(yo) || yo.rol === 'finanzas' || yo.rol === 'dueno' || x.responsable === yo.nombre || (lugarDe(yo) && x.lugar === lugarDe(yo)); }
function lugaresEq() { return UNIDADES.filter(function (u) { return u.vende || u.id === 'bodega'; }).map(function (u) { return { id: u.id, nombre: u.nombre }; }); }
/** El responsable que se propone para un lugar: el gerente de la sucursal (o nadie, en bodega). */
function responsablePorDefecto(lugar) {
  var g = usuariosCache().filter(function (u) { return u.activo && u.rol === 'gerente' && u.sucursal === lugar; })[0];
  return g ? g.nombre : '';
}
function pantallaEquipos(cred) {
  var yo = quien(cred);
  if (!jefeEquipos(yo) && ['gerente', 'bodega', 'produccion', 'finanzas', 'dueno'].indexOf(yo.rol) < 0) throw new Error('Su usuario no ve los equipos.');
  var hoy = hoyISO(), eqs = filasEqA().filter(function (x) { return x.activo && veEquipo(yo, x); });
  var revs = filasEqI(), ult = {};
  revs.forEach(function (r) { if (!ult[r.lugar] || r.fecha >= ult[r.lugar].fecha) ult[r.lugar] = r; });
  var lugares = lugaresEq().filter(function (l) { return jefeEquipos(yo) || yo.rol === 'finanzas' || yo.rol === 'dueno' || l.id === lugarDe(yo) || eqs.some(function (x) { return x.lugar === l.id; }); })
    .map(function (l) { var u = ult[l.id], n = eqs.filter(function (x) { return x.lugar === l.id; }).length;
      return { id: l.id, nombre: l.nombre, equipos: n, ultima: u ? u.fecha : '', faltan: u ? u.faltan : 0,
        toca: n > 0 && (!u || diasEntre(u.fecha, hoy) >= DIAS_REV_EQ), responsable: responsablePorDefecto(l.id) }; });
  var faltantes = [];
  Object.keys(ult).forEach(function (k) { var r = ult[k]; r.detalle.forEach(function (d) { if (d.falta > 0) faltantes.push({ lugar: nombreUnidad(r.lugar), equipo: d.equipo, falta: d.falta,
    responsable: r.responsable, fecha: r.fecha, revision: r.numero, nota: d.nota || '' }); }); });
  var peds = filasEqS().filter(function (p) { return jefeEquipos(yo) || p.por === yo.nombre; }).reverse().slice(0, 40);
  var misRev = revs.filter(function (r) { return !r.confirmadoEn && r.responsable === yo.nombre; });
  return { hoy: hoy, yo: { nombre: yo.nombre, rol: yo.rol }, jefe: jefeEquipos(yo), lugarPropio: lugarDe(yo), estados: EQA_ESTADOS,
    lugares: lugares, equipos: eqs.map(function (x) { x.lugarNombre = nombreUnidad(x.lugar); return x; }),
    revisiones: revs.filter(function (r) { return jefeEquipos(yo) || yo.rol === 'finanzas' || yo.rol === 'dueno' || r.responsable === yo.nombre || r.lugar === lugarDe(yo); }).reverse().slice(0, 30)
      .map(function (r) { r.lugarNombre = nombreUnidad(r.lugar); return r; }),
    porConfirmar: misRev.map(function (r) { r.lugarNombre = nombreUnidad(r.lugar); return r; }),
    faltantes: jefeEquipos(yo) || yo.rol === 'finanzas' || yo.rol === 'dueno' ? faltantes : faltantes.filter(function (f) { return f.responsable === yo.nombre; }),
    pedidos: peds.map(function (p) { p.lugarNombre = nombreUnidad(p.lugar); return p; }),
    puedePedir: !!lugarDe(yo) || yo.rol === 'operaciones',
    personas: jefeEquipos(yo) ? usuariosCache().filter(function (u) { return u.activo && ['gerente', 'bodega', 'produccion', 'operaciones'].indexOf(u.rol) >= 0; }).map(function (u) { return u.nombre; }).sort() : [] };
}
/** Agregar o cambiar un equipo (el director operativo o el administrador). */
function guardarEquipoA(cred, d) {
  var yo = quien(cred);
  if (!jefeEquipos(yo)) throw new Error('Los equipos los registra el director operativo.');
  d = d || {};
  var equipo = String(d.equipo || '').trim().replace(/\s+/g, ' ').slice(0, 70);
  if (equipo.length < 3) throw new Error('Escriba el equipo (por ejemplo «Estufa industrial 4 hornillas»).');
  var lugar = String(d.lugar || '');
  if (!lugaresEq().some(function (l) { return l.id === lugar; })) throw new Error('Escoja dónde está el equipo.');
  var cant = Math.round(Number(String(d.cantidad == null ? 1 : d.cantidad).replace(',', '.')));
  if (!(cant >= 1 && cant <= 999)) throw new Error('La cantidad va de 1 a 999.');
  var resp = String(d.responsable || '').trim() || responsablePorDefecto(lugar);
  if (resp && !usuariosCache().some(function (u) { return u.activo && u.nombre === resp; })) throw new Error('El responsable «' + resp + '» no es un usuario activo.');
  var estado = EQA_ESTADOS.indexOf(d.estado) >= 0 ? d.estado : 'Bueno';
  var lock = LockService.getScriptLock(); lock.waitLock(15000); _LEE = {};
  var numero, x = null;
  try {
    var h = hojasEquipos().a, todos = filasEqA();
    todos.forEach(function (y) { if (y.numero === String(d.numero || '')) x = y; });
    if (x) {
      var desde = x.responsable !== resp || x.lugar !== lugar ? hoyISO() : x.desde;
      h.getRange(x.fila, 2, 1, 7).setValues([[lugar, equipo, String(d.detalle || '').slice(0, 120), cant, resp, desde, estado]]);
      h.getRange(x.fila, 7).setNumberFormat('@').setValue(desde);
      h.getRange(x.fila, 12).setValue(String(d.nota || '').slice(0, 200));
      numero = x.numero;
    } else {
      numero = numeroSiguiente(todos, 'EQP');
      h.appendRow([numero, lugar, equipo, String(d.detalle || '').slice(0, 120), cant, resp, hoyISO(), estado, 'Sí', yo.nombre, new Date(), String(d.nota || '').slice(0, 200)]);
      h.getRange(h.getLastRow(), 7).setNumberFormat('@').setValue(hoyISO());
    }
  } finally { lock.releaseLock(); }
  var r = pantallaEquipos(cred); r.ok = true; r.numero = numero;
  r.mensaje = equipo + (x ? ' actualizado.' : ' registrado en ' + nombreUnidad(lugar) + (resp ? ', a cargo de ' + resp : '') + '.');
  return r;
}
/** Dar de baja un equipo (ya no existe, se vendió, se botó): sale del inventario. */
function bajaEquipo(cred, numero, motivo) {
  var yo = quien(cred);
  if (!jefeEquipos(yo)) throw new Error('Los equipos los da de baja el director operativo.');
  motivo = String(motivo || '').trim();
  if (motivo.length < 4) throw new Error('Escriba por qué se da de baja.');
  var x = null; filasEqA().forEach(function (y) { if (y.numero === String(numero)) x = y; });
  if (!x) throw new Error('No se encontró ese equipo.');
  var h = hojasEquipos().a; h.getRange(x.fila, 9).setValue('No'); h.getRange(x.fila, 12).setValue(('Baja: ' + motivo + ' (' + yo.nombre + ', ' + hoyISO() + ')').slice(0, 200));
  var r = pantallaEquipos(cred); r.ok = true; r.mensaje = x.equipo + ' dado de baja.'; return r;
}
/** La revisión de cada mes: d = { lugar, items: [{ numero, encontrados, estado, nota }], nota } */
function revisarEquipos(cred, d) {
  var yo = quien(cred);
  if (!jefeEquipos(yo)) throw new Error('La revisión de equipos la hace el director operativo con el responsable.');
  d = d || {};
  var lugar = String(d.lugar || ''), eqs = filasEqA().filter(function (x) { return x.activo && x.lugar === lugar; });
  if (!eqs.length) throw new Error('Ese lugar no tiene equipos registrados.');
  var porNum = {}; (d.items || []).forEach(function (i) { porNum[String(i.numero)] = i; });
  var falta = eqs.filter(function (x) { return !porNum[x.numero]; });
  if (falta.length) throw new Error('Falta revisar: ' + falta.slice(0, 3).map(function (x) { return x.equipo; }).join(', ') + (falta.length > 3 ? ' y ' + (falta.length - 3) + ' más' : '') + '.');
  var det = [], estan = 0, faltan = 0, danados = 0;
  eqs.forEach(function (x) {
    var i = porNum[x.numero], enc = Math.round(Number(String(i.encontrados == null ? '' : i.encontrados).replace(',', '.')));
    if (!(enc >= 0 && enc <= 999)) throw new Error('Diga cuántos hay de «' + x.equipo + '».');
    var est = EQA_ESTADOS.indexOf(i.estado) >= 0 ? i.estado : x.estado, fl = Math.max(0, x.cantidad - enc), nota = String(i.nota || '').trim().slice(0, 150);
    if (fl > 0 && nota.length < 3) throw new Error('Escriba qué pasó con «' + x.equipo + '» (faltan ' + fl + ').');
    if (est !== 'Bueno' && enc > 0) danados++;
    if (fl > 0) faltan += fl; else estan++;
    det.push({ numero: x.numero, equipo: x.equipo, esperado: x.cantidad, encontrados: enc, falta: fl, estado: est, nota: nota });
  });
  var resp = (eqs.map(function (x) { return x.responsable; }).filter(Boolean)[0]) || responsablePorDefecto(lugar);
  var lock = LockService.getScriptLock(); lock.waitLock(15000); _LEE = {};
  var numero;
  try {
    var hs = hojasEquipos(), todos = filasEqI();
    numero = numeroSiguiente(todos, 'REQ');
    hs.i.appendRow([numero, lugar, hoyISO(), yo.nombre, resp, eqs.length, estan, faltan, danados, '', '', String(d.nota || '').slice(0, 200), JSON.stringify(det)]);
    hs.i.getRange(hs.i.getLastRow(), 3).setNumberFormat('@').setValue(hoyISO());
    var porFila = {}; eqs.forEach(function (x) { porFila[x.numero] = x; });
    det.forEach(function (z) { var x = porFila[z.numero]; if (x && z.estado !== x.estado) hs.a.getRange(x.fila, 8).setValue(z.estado); });
  } finally { lock.releaseLock(); }
  var r = pantallaEquipos(cred); r.ok = true; r.numero = numero;
  r.mensaje = 'Revisión de ' + nombreUnidad(lugar) + ': ' + (faltan ? faltan + (faltan === 1 ? ' equipo falta' : ' equipos faltan') + ' (se reporta)' : 'están todos') +
    (danados ? ' · ' + danados + (danados === 1 ? ' dañado' : ' dañados') : '') + '.' + (resp ? ' ' + resp + ' la confirma en su app.' : '');
  return r;
}
/** El responsable confirma la revisión (como firmar). */
function confirmarRevisionEq(cred, numero) {
  var yo = quien(cred);
  var x = null; filasEqI().forEach(function (y) { if (y.numero === String(numero)) x = y; });
  if (!x) throw new Error('No se encontró esa revisión.');
  if (x.responsable !== yo.nombre) throw new Error('La confirma ' + (x.responsable || 'el responsable del lugar') + '.');
  if (x.confirmadoEn) throw new Error('Ya la confirmó.');
  hojasEquipos().i.getRange(x.fila, 10, 1, 2).setValues([[yo.nombre, new Date()]]);
  var r = pantallaEquipos(cred); r.ok = true; r.mensaje = 'Revisión ' + x.numero + ' confirmada.'; return r;
}
/** Pedir equipo: d = { que, para } (para su lugar). El director operativo decide. */
function pedirEquipo(cred, d) {
  var yo = quien(cred), lugar = lugarDe(yo) || (yo.rol === 'operaciones' ? String((d || {}).lugar || '') : '');
  if (!lugar) throw new Error('Los equipos los piden los gerentes, la bodega y producción.');
  d = d || {};
  var que = String(d.que || '').trim().slice(0, 120), para = String(d.para || '').trim().slice(0, 200);
  if (que.length < 4) throw new Error('Escriba qué equipo necesita (por ejemplo «Tanque de gas de 100 lb»).');
  if (para.length < 4) throw new Error('Escriba para qué lo necesita.');
  var lock = LockService.getScriptLock(); lock.waitLock(15000); _LEE = {};
  var numero;
  try { var h = hojasEquipos().s; numero = numeroSiguiente(filasEqS(), 'PEQ'); h.appendRow([numero, new Date(), yo.nombre, lugar, que, para, EQS_EST.PED, '', '', '', '', '', '', '']); }
  finally { lock.releaseLock(); }
  var r = pantallaEquipos(cred); r.ok = true; r.mensaje = numero + ' enviado al director operativo: ' + que + '.'; return r;
}
/** El director decide: d = { como: 'asignar' | 'comprar' | 'rechazar', equipo (no. del que se asigna), monto, nota } */
function resolverPedidoEq(cred, numero, d) {
  var yo = quien(cred);
  if (!jefeEquipos(yo)) throw new Error('Los pedidos de equipo los decide el director operativo.');
  d = d || {};
  var p = null; filasEqS().forEach(function (y) { if (y.numero === String(numero)) p = y; });
  if (!p) throw new Error('No se encontró ese pedido.');
  if (p.estado !== EQS_EST.PED) throw new Error('Ese pedido ya se resolvió (' + p.estado.toLowerCase() + ').');
  var hs = hojasEquipos(), nota = String(d.nota || '').trim().slice(0, 200), msg;
  if (d.como === 'rechazar') {
    if (nota.length < 4) throw new Error('Escriba por qué no se da.');
    hs.s.getRange(p.fila, 7, 1, 4).setValues([[EQS_EST.RECH, yo.nombre, new Date(), 'Rechazado']]); hs.s.getRange(p.fila, 13).setValue(nota);
    msg = numero + ' rechazado. ' + p.por + ' lo verá en su campanita.';
  } else if (d.como === 'asignar') {
    var x = null; filasEqA().forEach(function (y) { if (y.numero === String(d.equipo || '') && y.activo) x = y; });
    if (!x) throw new Error('Escoja el equipo que se le asigna.');
    hs.a.getRange(x.fila, 2).setValue(p.lugar); hs.a.getRange(x.fila, 6).setValue(p.por); hs.a.getRange(x.fila, 7).setNumberFormat('@').setValue(hoyISO());
    hs.s.getRange(p.fila, 7, 1, 5).setValues([[EQS_EST.ASIG, yo.nombre, new Date(), 'Asignar uno que ya existe', x.numero]]); if (nota) hs.s.getRange(p.fila, 13).setValue(nota);
    msg = numero + ': se le asigna ' + x.equipo + ' (' + x.numero + '). Queda a cargo de ' + p.por + ' cuando confirme que lo recibió.';
  } else if (d.como === 'comprar') {
    var so = pedirFondos(cred, { tipo: 'Equipo', unidad: nombreUnidad(p.lugar), que: 'Para ' + p.por + ': ' + p.que + (p.para ? ' (' + p.para + ')' : ''), monto: d.monto, urgencia: d.urgencia || 'Normal', archivo: d.archivo });
    var sn = (/SO-\d+/.exec(so.mensaje || '') || [''])[0];
    hs.s.getRange(p.fila, 7, 1, 6).setValues([[EQS_EST.COMPRA, yo.nombre, new Date(), 'Comprar', '', sn]]); if (nota) hs.s.getRange(p.fila, 13).setValue(nota);
    msg = numero + ': se compra. Va a contabilidad como solicitud ' + sn + '. Cuando llegue, regístrelo a cargo de ' + p.por + '.';
  } else throw new Error('Escoja qué hacer con el pedido.');
  var r = pantallaEquipos(cred); r.ok = true; r.mensaje = msg; return r;
}
/** Un equipo comprado ya llegó: se registra a cargo de quien lo pidió. */
function entregarCompraEq(cred, numero, d) {
  var yo = quien(cred);
  if (!jefeEquipos(yo)) throw new Error('Lo registra el director operativo.');
  var p = null; filasEqS().forEach(function (y) { if (y.numero === String(numero)) p = y; });
  if (!p || p.estado !== EQS_EST.COMPRA) throw new Error('Ese pedido no está en compra.');
  d = d || {};
  var r0 = guardarEquipoA(cred, { lugar: p.lugar, equipo: d.equipo || p.que, detalle: d.detalle || '', cantidad: d.cantidad || 1, responsable: p.por, estado: 'Bueno', nota: 'Pedido ' + p.numero });
  hojasEquipos().s.getRange(p.fila, 7).setValue(EQS_EST.ASIG); hojasEquipos().s.getRange(p.fila, 11).setValue(r0.numero);
  var r = pantallaEquipos(cred); r.ok = true; r.mensaje = (d.equipo || p.que) + ' registrado. ' + p.por + ' confirma que lo recibió.'; return r;
}
/** Quien pidió confirma que lo recibió: desde hoy queda a su cargo. */
function recibirEquipo(cred, numero) {
  var yo = quien(cred);
  var p = null; filasEqS().forEach(function (y) { if (y.numero === String(numero)) p = y; });
  if (!p) throw new Error('No se encontró ese pedido.');
  if (p.por !== yo.nombre) throw new Error('Lo confirma ' + p.por + '.');
  if (p.estado !== EQS_EST.ASIG) throw new Error('Ese pedido no tiene un equipo por recibir.');
  var hs = hojasEquipos(); hs.s.getRange(p.fila, 7).setValue(EQS_EST.REC); hs.s.getRange(p.fila, 14).setValue(new Date());
  var x = null; filasEqA().forEach(function (y) { if (y.numero === p.equipo) x = y; });
  if (x) hs.a.getRange(x.fila, 7).setNumberFormat('@').setValue(hoyISO());
  var r = pantallaEquipos(cred); r.ok = true; r.mensaje = 'Recibido: ' + (x ? x.equipo : p.que) + ' queda a su cargo desde hoy.'; return r;
}
/** Los avisos de equipos (para la campanita). */
function avisosEquipos(yo) {
  var out = [], hoy = hoyISO();
  try {
    if (jefeEquipos(yo)) {
      filasEqS().forEach(function (p) { if (p.estado === EQS_EST.PED) out.push({ id: 'peq-' + p.numero, tipo: 'Pedido de equipo', cuando: p.en, titulo: p.por + ' pide: ' + p.que, texto: p.para, abrir: { pag: 'equipos' } }); });
      var eqs = filasEqA().filter(function (x) { return x.activo; }), ult = {};
      filasEqI().forEach(function (r) { if (!ult[r.lugar] || r.fecha >= ult[r.lugar].fecha) ult[r.lugar] = r; });
      lugaresEq().forEach(function (l) { if (!eqs.some(function (x) { return x.lugar === l.id; })) return; var u = ult[l.id];
        if (!u || diasEntre(u.fecha, hoy) >= DIAS_REV_EQ) out.push({ id: 'revq-' + l.id + '-' + hoy.slice(0, 7), tipo: 'Revisión de equipos', cuando: '', titulo: 'Toca revisar los equipos de ' + l.nombre,
          texto: u ? 'La última fue el ' + dia(u.fecha) + '.' : 'Todavía no se han revisado.', abrir: { pag: 'equipos' } }); });
      Object.keys(ult).forEach(function (k) { var r = ult[k]; if (r.faltan > 0) out.push({ id: 'faleq-' + r.numero, tipo: 'Faltan equipos', cuando: r.fecha, titulo: (r.faltan === 1 ? 'Falta 1 equipo' : 'Faltan ' + r.faltan + ' equipos') + ' en ' + nombreUnidad(r.lugar),
        texto: r.detalle.filter(function (d) { return d.falta > 0; }).map(function (d) { return d.equipo + ' (' + d.falta + ')'; }).join(', '), abrir: { pag: 'equipos' } }); });
    }
    filasEqI().forEach(function (r) { if (!r.confirmadoEn && r.responsable === yo.nombre) out.push({ id: 'confeq-' + r.numero, tipo: 'Revisión de equipos', cuando: r.fecha,
      titulo: 'Confirme la revisión de equipos de ' + nombreUnidad(r.lugar), texto: r.faltan ? 'Faltan ' + r.faltan + '.' : 'Están todos.', abrir: { pag: 'equipos' } }); });
    filasEqS().forEach(function (p) { if (p.por !== yo.nombre) return;
      if (p.estado === EQS_EST.ASIG) out.push({ id: 'asigeq-' + p.numero, tipo: 'Equipo', cuando: p.resueltoEn, titulo: 'Le asignaron: ' + p.que, texto: 'Confirme cuando lo reciba.', abrir: { pag: 'equipos' } });
      if (p.estado === EQS_EST.RECH) out.push({ id: 'recheq-' + p.numero, tipo: 'Equipo', cuando: p.resueltoEn, titulo: 'No se aprobó: ' + p.que, texto: p.nota, abrir: { pag: 'equipos' } });
      if (p.estado === EQS_EST.COMPRA) out.push({ id: 'compeq-' + p.numero, tipo: 'Equipo', cuando: p.resueltoEn, titulo: 'Se va a comprar: ' + p.que, texto: p.solicitud ? 'Solicitud ' + p.solicitud : '', abrir: { pag: 'equipos' } });
    });
  } catch (e) {}
  return out;
}

/* ════════════ SOLICITUDES DE PAGO: contabilidad puede devolverlas ════════════
 * Si la factura o la cotización no tiene la misma cantidad que el dinero que se pide, contabilidad la devuelve a quien la pidió (con el motivo).
 * Esa persona corrige el monto (o cambia el archivo) y la reenvía: vuelve a empezar su camino de aprobación, porque el monto cambió. */
var FONDO_DEV = 'Devuelta';
function devolverFondo(cred, numero, motivo) {
  var yo = quien(cred);
  if (!(esContador(yo) || yo.esAdmin || yo.rol === 'finanzas')) throw new Error('Las solicitudes las devuelve contabilidad.');
  var x = buscaFondo(numero);
  if (x.estado !== FONDO_EST.APR && x.estado !== FONDO_EST.PED) throw new Error('Esa solicitud ya está ' + x.estado.toLowerCase() + '.');
  motivo = String(motivo || '').trim().slice(0, 300);
  if (motivo.length < 5) throw new Error('Escriba por qué se devuelve (por ejemplo: la factura no tiene la misma cantidad que el dinero que se solicita).');
  var h = hojaFondos();
  h.getRange(x.fila, 10).setValue(FONDO_DEV);
  h.getRange(x.fila, 28, 1, 4).setValues([[yo.nombre, new Date(), motivo, x.monto]]);
  var r = listaFondos(cred); r.ok = true;
  r.mensaje = numero + ' devuelta a ' + x.por + '. Corrige el monto o el archivo y la reenvía.';
  return r;
}
/** Quien la pidió corrige: d = { monto, que, archivo }. Vuelve al inicio de su camino de aprobación. */
function corregirFondo(cred, numero, d) {
  permiteCmo();
  var yo = quien(cred);
  var x = buscaFondo(numero);
  if (x.por !== yo.nombre) throw new Error('La corrige ' + x.por + ', quien la pidió.');
  if (x.estado !== FONDO_DEV) throw new Error('Esa solicitud no está devuelta.');
  d = d || {};
  var monto = d.monto == null || String(d.monto).trim() === '' ? x.monto : r2(String(d.monto).replace(/[Q,\s]/g, ''));
  if (!(monto > 0)) throw new Error('Escriba cuánto dinero se necesita.');
  var que = d.que == null || String(d.que).trim() === '' ? x.que : String(d.que).trim().slice(0, 300);
  if (que.length < 8) throw new Error('Describa qué se necesita y por qué.');
  var arch = d.archivo && d.archivo.datos ? guardaArchivo(d.archivo, 'Solicitud de ' + yo.nombre + (rolDe(yo.nombre) === 'cmo' ? ' · factura' : ' · cotización'), rolDe(yo.nombre) === 'cmo' ? 'La factura' : '') : null;
  if (monto === x.monto && !arch && que === x.que) throw new Error('Cambie el monto o suba el archivo correcto; si todo está bien, contabilidad la puede pagar como está.');
  var rol = rolDe(yo.nombre), vuelveA = estadoInicialFondo(rol, x.reqDinero);          // vuelve a empezar su camino de aprobación
  var lock = LockService.getScriptLock(); lock.waitLock(15000); _LEE = {};
  try {
    var h = hojaFondos();
    h.getRange(x.fila, 6).setValue(que); h.getRange(x.fila, 8).setValue(monto);
    if (arch) h.getRange(x.fila, 23, 1, 2).setValues([[arch.url, arch.nombre]]);
    h.getRange(x.fila, 10, 1, 4).setValues([[vuelveA, '', '', '']]);          // vuelve a empezar: nadie la ha aprobado con este monto
    h.getRange(x.fila, 25, 1, 3).setValues([['', '', '']]);
    h.getRange(x.fila, 33, 1, 3).setValues([['', '', '']]); h.getRange(x.fila, 36).setValue('');
  } finally { lock.releaseLock(); }
  var r = listaFondos(cred); r.ok = true;
  r.mensaje = numero + ' corregida' + (monto !== x.monto ? ' (' + dinero(x.monto) + ' → ' + dinero(monto) + ')' : '') + ' y reenviada ' + (vuelveA === FONDO_EST.VER ? 'al director operativo.' : vuelveA === FONDO_ADM ? 'al administrador.' : 'a finanzas para aprobar.');
  return r;
}

/* ════════════ PENDIENTE DE ENTREGA ════════════
 * Si Samuel manda menos de lo que pidió una sucursal (porque no hay en bodega), lo que falta queda PENDIENTE DE ENTREGA en el pedido, no se pierde.
 * Se respalda con una orden de compra abierta: sin orden, el pendiente sale marcado «Sin orden de compra» (hay que pedirlo al proveedor).
 * Cuando llega a bodega (la orden se recibe), pasa a «Listo para despachar» y Samuel lo manda en un pedido nuevo derivado del original.
 * El gerente ve en su pedido qué le falta y en qué va: orden de compra hecha, por entrar a bodega, o ya listo. */
var PEND_EST = { SIN: 'Sin orden de compra', OC: 'En orden de compra', LISTO: 'Listo para despachar', ENT: 'Entregado', CAN: 'No se entrega' };
function pendDeLinea(l) { return l.pendCanPor || !(l.pend > 0) ? 0 : Math.max(0, r3(l.pend - l.pendEntregado)); }
/** Lo que hace falta saber de bodega para juzgar cada pendiente: órdenes abiertas, existencia libre (sin lo ya comprometido) y factores. */
function contextoPend() {
  var ocs = {}; filasOC().forEach(function (o) { if (o.estado === 'Pedida' && o.codigo && !ocs[o.codigo]) ocs[o.codigo] = o.numero; });
  var catT = {}; productosTraslado().forEach(function (p) { catT[p.key] = p; });
  var ex = {}; try { ex = estadoInventario('bodega').existencia || {}; } catch (e) {}
  var comp = {}, abiertas = {};
  filasSol().forEach(function (s) { if (s.estado === SOL.ACE || s.estado === SOL.CAM) abiertas[s.numero] = true; });
  lineasSol().forEach(function (l) { if (!abiertas[l.numero]) return; var f = catT[l.clave] ? catT[l.clave].factor : 1; comp[l.codigo] = (comp[l.codigo] || 0) + (l.enviado != null ? l.enviado : l.pedido) * f; });
  return { ocs: ocs, catT: catT, ex: ex, comp: comp };
}
/** Cuánto hay libre en bodega de esta línea, en la unidad en que se pidió (cajas, unidades sueltas…). */
function libreDeLinea(l, ctx) {
  var f = ctx.catT[l.clave] ? ctx.catT[l.clave].factor : 1;
  return f > 0 ? Math.max(0, r3(((ctx.ex[l.codigo] || 0) - (ctx.comp[l.codigo] || 0)) / f)) : 0;
}
function estadoPend(l, ctx) {
  if (!(l.pend > 0)) return '';
  if (l.pendCanPor) return PEND_EST.CAN;
  var deb = pendDeLinea(l);
  if (deb <= 0.0005) return PEND_EST.ENT;
  if (libreDeLinea(l, ctx) >= deb - 1e-9) return PEND_EST.LISTO;
  return ctx.ocs[l.codigo] ? PEND_EST.OC : PEND_EST.SIN;
}
function textoPend(est, l, ctx) {
  var ll = l.pendLlega ? ' · llega ' + l.pendLlega : '';
  if (est === PEND_EST.LISTO) return 'Ya hay en bodega: sale en la próxima entrega';
  if (est === PEND_EST.OC) return 'Orden de compra ' + ctx.ocs[l.codigo] + ' hecha: por entrar a bodega central' + ll;
  if (est === PEND_EST.SIN) return 'Todavía no hay orden de compra: bodega lo está viendo' + ll;
  if (est === PEND_EST.CAN) return 'No se entrega' + (l.pendMotivo ? ': ' + l.pendMotivo : '');
  return est === PEND_EST.ENT ? 'Entregado' : '';
}
/** La lista de lo pendiente: la bodega y el director ven todo; el gerente, lo de su sucursal. */
function listaPendientes(cred) {
  var yo = quien(cred), esBod = yo.rol === 'bodega';
  if (!esBod && yo.rol !== 'gerente' && !veInv(yo)) throw new Error('Su usuario no ve los pendientes de entrega.');
  var mias = yo.rol === 'gerente' ? nombreUnidad(yo.sucursal) : '';
  var sols = {}; filasSol().forEach(function (s) { sols[s.numero] = s; });
  var ls = lineasSol().filter(function (l) { return l.pend > 0 && pendDeLinea(l) > 0.0005 && sols[l.numero] && sols[l.numero].estado !== SOL.ANU && (!mias || sols[l.numero].sucursal === mias); });
  var out = [], ctx = ls.length ? contextoPend() : null, cat = {};
  if (ls.length) productosInv().forEach(function (b) { cat[b.codigo] = b; });
  var orden = {}; orden[PEND_EST.LISTO] = 0; orden[PEND_EST.SIN] = 1; orden[PEND_EST.OC] = 2;
  ls.forEach(function (l) {
    var s = sols[l.numero], est = estadoPend(l, ctx), deb = pendDeLinea(l), b = cat[l.codigo] || {};
    out.push({ numero: l.numero, sucursal: s.sucursal, enviadaEn: s.enviadaEn, clave: l.clave, codigo: l.codigo, nombre: l.nombre, medida: l.medida, pedido: l.pedido, enviado: l.enviado,
      pendiente: deb, pendTotal: l.pend, entregado: l.pendEntregado, estado: est, orden: est === PEND_EST.OC ? ctx.ocs[l.codigo] : '', llega: l.pendLlega,
      libre: libreDeLinea(l, ctx), enBodegaTxt: b.codigo ? textoCantidad(ctx.ex[l.codigo] || 0, b) : '', texto: textoPend(est, l, ctx),
      puedeDespachar: esBod && est === PEND_EST.LISTO, puedeGestionar: esBod });
  });
  out.sort(function (a, b) { return (orden[a.estado] - orden[b.estado]) || a.sucursal.localeCompare(b.sucursal, 'es') || a.nombre.localeCompare(b.nombre, 'es'); });
  var cuenta = function (e) { return out.filter(function (x) { return x.estado === e; }).length; };
  return { pendientes: out, resumen: { listos: cuenta(PEND_EST.LISTO), sinOrden: cuenta(PEND_EST.SIN), enOrden: cuenta(PEND_EST.OC) }, puedeGestionar: esBod, esGerente: yo.rol === 'gerente' };
}
function lineaPend(numero, clave) {
  var x = null; lineasSol(String(numero)).forEach(function (l) { if (l.clave === String(clave)) x = l; });
  if (!x || !(x.pend > 0)) throw new Error('Ese producto no está pendiente en ' + numero + '.');
  return x;
}
/** Samuel avisa cuándo llega: «el jueves», «mañana». Es lo que ve el gerente. */
function llegaPendiente(cred, numero, clave, texto) {
  var yo = exigeSoloBodega(cred), l = lineaPend(numero, clave);
  texto = String(texto || '').trim().slice(0, 40);
  hojaSolD().getRange(l.fila, 18).setValue(texto);
  var r = listaPendientes(cred); r.ok = true; r.mensaje = texto ? l.nombre + ': el gerente verá que llega ' + texto + '.' : 'Se quitó la fecha de llegada.'; return r;
}
/** No se va a entregar (se descontinuó, ya no hace falta): sale de pendientes, con el motivo. */
function cancelarPendiente(cred, numero, clave, motivo) {
  var yo = exigeSoloBodega(cred), l = lineaPend(numero, clave);
  motivo = String(motivo || '').trim().slice(0, 150);
  if (motivo.length < 4) throw new Error('Escriba por qué no se va a entregar.');
  if (pendDeLinea(l) <= 0.0005) throw new Error('Eso ya no está pendiente.');
  hojaSolD().getRange(l.fila, 20, 1, 2).setValues([[yo.nombre, motivo]]);
  var r = listaPendientes(cred); r.ok = true; r.mensaje = l.nombre + ' ya no se entrega a ' + numero + '. El gerente verá el motivo.'; return r;
}
/** Samuel despacha lo pendiente que ya hay en bodega. items = [{ numero, clave, cantidad? }]. Por cada sucursal se arma un pedido nuevo, ya aceptado, que sigue el camino de siempre (camión, recibir, diferencias). */
function despacharPendientes(cred, items) {
  var yo = exigeSoloBodega(cred);
  items = items || [];
  if (!items.length) throw new Error('Escoja qué se va a despachar.');
  var sols = {}; filasSol().forEach(function (s) { sols[s.numero] = s; });
  var ctx = contextoPend(), queda = {}, porSuc = {}, usadas = {};
  items.forEach(function (it) {
    var l = lineaPend(it.numero, it.clave), s = sols[l.numero];
    if (!s || s.estado === SOL.ANU) throw new Error('El pedido ' + l.numero + ' está anulado.');
    var key = l.numero + '|' + l.clave; if (usadas[key]) return; usadas[key] = true;
    var deb = pendDeLinea(l); if (deb <= 0.0005) throw new Error(l.nombre + ' ya no está pendiente.');
    var cant = it.cantidad == null || String(it.cantidad).trim() === '' ? deb : r3(Number(String(it.cantidad).replace(',', '.')));
    if (!(cant > 0) || !isFinite(cant)) throw new Error('La cantidad de «' + l.nombre + '» no es válida.');
    if (cant > deb + 0.0005) throw new Error('De «' + l.nombre + '» solo falta entregar ' + num(deb) + ' ' + plur(deb, l.medida) + '.');
    var f = ctx.catT[l.clave] ? ctx.catT[l.clave].factor : 1;
    if (queda[l.codigo] == null) queda[l.codigo] = Math.max(0, (ctx.ex[l.codigo] || 0) - (ctx.comp[l.codigo] || 0));
    if (cant * f > queda[l.codigo] + 1e-9) throw new Error('De «' + l.nombre + '» no hay suficiente en bodega: libre ' + num(libreDeLinea(l, ctx)) + ' ' + plur(libreDeLinea(l, ctx), l.medida) + '.');
    queda[l.codigo] = r3(queda[l.codigo] - cant * f);
    (porSuc[s.sucursal] = porSuc[s.sucursal] || { s: s, lineas: [], origenes: {} }).lineas.push({ l: l, cant: cant });
    porSuc[s.sucursal].origenes[l.numero] = true;
  });
  var hechas = [], lock = LockService.getScriptLock(); lock.waitLock(25000); _LEE = {};
  try {
    var hs = hojaSol(), hd = hojaSolD(), ahora = new Date(), hoy = hoyISO();
    Object.keys(porSuc).forEach(function (suc) {
      var g = porSuc[suc], numero = siguienteSol(), orig = Object.keys(g.origenes).join(', ');
      hs.appendRow([numero, suc, ahora, g.s.enviadaPor, '', hoy, SOL.ACE, ahora, yo.nombre, '', '', '', 'Pendiente de ' + orig, '', '', ahora, ahora, '', '', '']);
      hs.getRange(hs.getLastRow(), 6).setNumberFormat('@').setValue(hoy);
      var filas = g.lineas.map(function (x) { return [numero, suc, x.l.codigo, x.l.nombre, x.l.medida, x.l.clave, '', 0, 'Pendiente de ' + x.l.numero, x.cant, x.cant, '', '', '', '', '', '', '', '', '', '']; });
      hd.getRange(hd.getLastRow() + 1, 1, filas.length, H_SOLD.length).setValues(filas);
      g.lineas.forEach(function (x) { hd.getRange(x.l.fila, 19).setValue(r3(x.l.pendEntregado + x.cant)); });      // lo que ya se entregó del pendiente
      hechas.push(numero + ' para ' + suc + ' (' + g.lineas.length + (g.lineas.length === 1 ? ' producto' : ' productos') + ')');
    });
  } finally { lock.releaseLock(); }
  var r = listaPendientes(cred); r.ok = true;
  r.mensaje = 'Listo: ' + hechas.join(' · ') + '. Ya está aceptado; imprima la hoja y márquelo «Salió el camión» cuando salga.';
  return r;
}
/** Avisos de pendientes: la bodega (ya hay en bodega / falta ordenar), el gerente (ya hay en bodega) y el director (sin orden de compra). */
function avisosPendientes(yo) {
  var out = [];
  var ls = lineasSol().filter(function (l) { return l.pend > 0 && pendDeLinea(l) > 0.0005; });
  if (!ls.length) return out;
  var sols = {}; filasSol().forEach(function (s) { sols[s.numero] = s; });
  var ctx = contextoPend(), sinOrden = 0;
  ls.forEach(function (l) {
    var s = sols[l.numero]; if (!s || s.estado === SOL.ANU) return;
    var est = estadoPend(l, ctx), deb = pendDeLinea(l), que = num(deb) + ' ' + plur(deb, l.medida);
    if (yo.rol === 'bodega' && est === PEND_EST.LISTO)
      out.push({ id: 'pendl-' + l.numero + '-' + l.clave, tipo: 'Pendiente listo', cuando: '', titulo: 'Ya hay en bodega: ' + l.nombre, texto: 'Falta entregar ' + que + ' a ' + s.sucursal + ' (' + l.numero + ').', abrir: { tab: 'pend' } });
    if (yo.rol === 'bodega' && est === PEND_EST.SIN)
      out.push({ id: 'pendso-' + l.numero + '-' + l.clave, tipo: 'Falta ordenar', cuando: '', titulo: 'Sin orden de compra: ' + l.nombre, texto: s.sucursal + ' espera ' + que + ' (' + l.numero + ').', abrir: { tab: 'compras' } });
    if (est === PEND_EST.SIN) sinOrden++;
    if (yo.rol === 'gerente' && s.sucursal === nombreUnidad(yo.sucursal) && est === PEND_EST.LISTO)
      out.push({ id: 'pendg-' + l.numero + '-' + l.clave, tipo: 'Pendiente de entrega', cuando: '', titulo: 'Ya hay en bodega: ' + l.nombre, texto: 'Le faltaba ' + que + ' de ' + l.numero + '. Sale en la próxima entrega.', abrir: { tab: 'pend' } });
  });
  if ((yo.rol === 'operaciones' || yo.esAdmin) && sinOrden)
    out.push({ id: 'pendsin-' + hoyISO(), tipo: 'Pendientes de entrega', cuando: '', titulo: sinOrden + (sinOrden === 1 ? ' producto pendiente sin orden de compra' : ' productos pendientes sin orden de compra'),
      texto: 'Las sucursales los esperan y bodega todavía no los pidió al proveedor.', abrir: { tab: 'pend' } });
  return out;
}

/* ════════════ LO QUE HAY QUE HACER HOY ════════════
 * Cada aviso que pide una acción (aprobar, pagar, aceptar, contar…) se marca como TAREA, con su verbo y su prioridad (1 dinero · 2 operación · 3 rutina).
 * El Inicio de cada persona los muestra primero; la campanita sigue recordando que hay algo pendiente. Lo demás (ya se pagó, ya llegó…) es solo información. */
var TAREA_DE = [
  ['em-rech-', null], ['recheq-', null],
  ['sd-apr-', 'Aprobar', 1], ['sd-ver-', 'Verificar', 1], ['sd-aut-', 'Autorizar', 1], ['sd-ok-', 'Proceder', 2], ['cierref-', 'Revisar', 2], ['cierre-', 'Cerrar mes', 2], ['sd-pag-', 'Pagar', 1], ['sd-dev-', 'Corregir', 1], ['val-', 'Validar', 1], ['rec-', 'Recibir efectivo', 1], ['banco-', 'Confirmar', 1],
  ['rech-', 'Corregir', 1], ['ocdev-', 'Corregir', 1], ['oc-cot-', 'Cotizar', 1], ['oc-pagar-', 'Pagar', 1], ['pres-por-', 'Aprobar', 1], ['rep-', 'Pagar', 1], ['dif-caja-', 'Revisar', 1],
  ['peq-', 'Decidir', 2], ['confeq-', 'Confirmar', 2], ['revq-', 'Revisar', 2], ['asigeq-', 'Confirmar', 2], ['pendl-', 'Despachar', 2], ['pendso-', 'Ordenar', 2], ['pendsin-', 'Revisar', 2],
  ['nueva-', 'Aceptar', 2], ['cargar-', 'Despachar', 2], ['cam-', 'Recibir', 2], ['dif-', 'Aclarar', 2], ['de-', 'Decidir', 2], ['perm-r-', 'Decidir', 2], ['cq-', 'Responder', 2],
  ['em-env-', 'Recibir', 2], ['em-', 'Enviar', 2],
  ['contar-', 'Contar', 3], ['kpifalta-', 'Anotar', 3], ['ace-falta-', 'Medir', 3], ['sg-', 'Leer', 3]
];
function tareaDeAviso(a) {
  var id = String(a.id || '');
  for (var i = 0; i < TAREA_DE.length; i++) if (id.indexOf(TAREA_DE[i][0]) === 0) return TAREA_DE[i][1] ? { accion: TAREA_DE[i][1], prio: TAREA_DE[i][2] } : null;
  return null;
}

/* ════════════ SOLICITUDES DE PAGO: el camino de aprobación según el rango de quien pide ════════════
 *  · Gerentes, Samuel y Paulino (los de Álvaro):  Álvaro verifica → el administrador autoriza → Daniel aprueba el pago → contabilidad entrega el dinero.
 *  · Los directores (Álvaro, Amalia):             el administrador autoriza → Daniel aprueba el pago → contabilidad.
 *  · Daniel (director financiero):                el administrador autoriza → contabilidad.
 *  · Si NO requiere dinero (un traslado, algo con lo que ya contamos): solo pasa por el administrador (y por Álvaro, si es de los suyos);
 *    al autorizarse, la solicitud vuelve a quien la pidió para que proceda. */
var FONDO_ADM = 'Por autorizar', FONDO_AUT = 'Autorizada';
var ETAPA_ESTADO = { ver: FONDO_EST.VER, adm: FONDO_ADM, fin: FONDO_EST.PED, con: FONDO_EST.APR };
function rutaFondo(rol, dinero) {
  var r = [];
  if (rol === 'gerente' || rol === 'bodega' || rol === 'produccion') r.push('ver');
  if (rol !== 'admin') r.push('adm');
  if (dinero) { if (rol !== 'finanzas') r.push('fin'); r.push('con'); }
  return r;
}
function estadoInicialFondo(rol, dinero) { var r = rutaFondo(rol, dinero); return r.length ? ETAPA_ESTADO[r[0]] : FONDO_AUT; }
function estadoSiguienteFondo(x, etapa) { var r = rutaFondo(rolDe(x.por), x.reqDinero), n = r[r.indexOf(etapa) + 1]; return n ? ETAPA_ESTADO[n] : FONDO_AUT; }
function estadoTxtFondo(x) {
  var m = {}; m[FONDO_EST.VER] = 'Por verificar'; m[FONDO_ADM] = 'Por autorizar'; m[FONDO_EST.PED] = 'Por aprobar el pago'; m[FONDO_EST.APR] = 'Por pagar';
  m[FONDO_EST.PAG] = 'Pagada'; m[FONDO_EST.FIN] = 'Terminada'; m[FONDO_AUT] = 'Autorizada'; m[FONDO_EST.RECH] = 'Rechazada'; m[FONDO_DEV] = 'Devuelta';
  return m[x.estado] || x.estado;
}
function etapaTxtFondo(x) {
  var m = {}; m[FONDO_EST.VER] = 'Espera al director operativo'; m[FONDO_ADM] = 'Espera al administrador'; m[FONDO_EST.PED] = 'Espera al director financiero (aprobar el pago)';
  m[FONDO_EST.APR] = 'Espera a contabilidad (entregar el dinero)'; m[FONDO_AUT] = 'Autorizada: ya puede proceder'; m[FONDO_DEV] = 'Devuelta: hay que corregirla';
  return m[x.estado] || '';
}
/** Los pasos de una solicitud, según su camino: quién hizo cada uno y cuándo. */
function pasosFondo(x) {
  var ruta = rutaFondo(rolDe(x.por), x.reqDinero), idx = function (e) { var i = ruta.indexOf(e); return i < 0 ? ruta.length : i; };
  var ei = ruta.length;
  if (x.estado === FONDO_EST.VER) ei = idx('ver'); else if (x.estado === FONDO_ADM) ei = idx('adm'); else if (x.estado === FONDO_EST.PED) ei = idx('fin'); else if (x.estado === FONDO_EST.APR) ei = idx('con');
  else if (x.estado === FONDO_DEV) ei = -1;
  else if (x.estado === FONDO_EST.RECH) { var er = x.etapaRech || ({ operaciones: 'ver', admin: 'adm', finanzas: 'fin' }[rolDe(x.aprPor)] || 'fin'); ei = idx(er); }
  var datos = { ver: ['Verificada', x.verPor, x.verEn], adm: ['Autorizada', x.autPor, x.autEn], fin: ['Aprobada', x.aprPor, x.aprEn], con: ['Pagada', x.pagPor, x.pagEn] };
  var pasos = [{ n: 'Pedida', por: x.por, en: x.en }];
  ruta.forEach(function (e, k) { var d = datos[e], hecho = k < ei; pasos.push({ n: d[0], por: hecho ? d[1] : '', en: hecho ? (d[2] || x.en) : '' }); });
  if (x.reqDinero) { if (rolDe(x.por) !== 'cmo' && !esEquipoBodega({ rol: rolDe(x.por) })) pasos.push({ n: 'Terminada', por: '', en: x.finEn }); }
  else pasos.push({ n: 'Hecha', por: x.por, en: x.finEn });
  return pasos;
}
function listaFondos(cred) {
  permiteCmo();
  var yo = quien(cred);
  if (!veFondos(yo)) throw new Error('Su usuario no ve las solicitudes.');
  var mes = hoyISO().slice(0, 7), ls = filasFondos().reverse(), soloSuyas = yo.rol === 'gerente' || yo.cmo || esEquipoBodega(yo);
  if (soloSuyas) ls = ls.filter(function (x) { return x.por === yo.nombre; });          // el gerente, la CMO y el equipo de bodega ven las suyas
  var p = null;
  if (!soloSuyas) { try { p = presupuestoMes({ usuario: yo.nombre, pin: cred.pin }, mes); } catch (e) {} }
  var primer = estadoInicialFondo(yo.esAdmin ? 'admin' : rolDe(yo.nombre), true);
  var r = { solicitudes: ls.slice(0, 60).map(function (x) {
      x.rolSol = rolDe(x.por);
      x.deBodega = esEquipoBodega({ rol: x.rolSol });
      x.deGerente = x.rolSol === 'gerente' || (!x.deBodega && (!!x.verPor || x.estado === FONDO_EST.VER));
      x.deCmo = x.rolSol === 'cmo';
      x.estadoTxt = estadoTxtFondo(x); x.etapa = etapaTxtFondo(x); x.pasos = pasosFondo(x);
      x.puedeVerificar = (yo.rol === 'operaciones' || yo.esAdmin) && x.estado === FONDO_EST.VER;
      x.puedeAutorizar = !!yo.esAdmin && x.estado === FONDO_ADM;
      x.puedeAprobar = (yo.rol === 'finanzas' || yo.esAdmin) && x.estado === FONDO_EST.PED;
      x.puedePagar = (esContador(yo) || yo.esAdmin) && x.estado === FONDO_EST.APR;
      x.puedeDevolver = (esContador(yo) || yo.esAdmin) && x.estado === FONDO_EST.APR;
      x.puedeCorregir = x.por === yo.nombre && x.estado === FONDO_DEV;
      x.devuelta = x.devEn && x.estado === FONDO_DEV ? { por: x.devPor, en: x.devEn, motivo: x.devMotivo, montoAntes: x.montoAntes } : null;
      x.corregida = !!(x.devEn && x.estado !== FONDO_DEV && x.montoAntes && x.montoAntes !== x.monto);
      x.puedeTerminar = (yo.rol === 'operaciones' || yo.esAdmin) && x.estado === FONDO_EST.PAG && !x.deCmo && !x.deBodega;
      x.puedeHecha = (x.por === yo.nombre || !!yo.esAdmin) && x.estado === FONDO_AUT;
      x.puedeAdjuntar = (yo.rol === 'operaciones' || yo.esAdmin || esContador(yo) || (x.por === yo.nombre && (x.estado === FONDO_EST.VER || x.estado === FONDO_ADM || (yo.cmo && x.estado === FONDO_EST.PED)))) &&
        x.estado !== FONDO_EST.FIN && x.estado !== FONDO_EST.RECH && x.estado !== FONDO_AUT;
      if (p) { var u = p.unidades.filter(function (y) { return y.nombre === x.unidad; })[0];
        var rb = u && u.rubros.filter(function (q) { return q.rubro === rubroDeFondo(x.tipo); })[0];
        if (rb) x.presupuesto = { rubro: rb.rubro, presupuesto: rb.presupuesto, gastado: rb.gastado }; }
      return x; }),
    tipos: yo.cmo ? Object.keys(FONDO_TIPOS_CMO) : esEquipoBodega(yo) ? Object.keys(FONDO_TIPOS_BODEGA) : Object.keys(FONDO_TIPOS),
    unidades: (yo.rol === 'gerente' || yo.cmo) ? [nombreUnidad(yo.sucursal)] : esEquipoBodega(yo) ? [nombreUnidad('bodega')] : UNIDADES.map(function (u) { return u.nombre; }),
    formas: FORMAS_PAGO, puedePedir: yo.rol === 'operaciones' || yo.esAdmin || yo.rol === 'finanzas' || yo.rol === 'gerente' || yo.cmo || esEquipoBodega(yo),
    esGerente: yo.rol === 'gerente', esCmo: !!yo.cmo, esDirector: yo.rol === 'operaciones' || yo.rol === 'finanzas', esBodega: esEquipoBodega(yo),
    pendientes: { devueltas: ls.filter(function (x) { return x.estado === FONDO_DEV; }).length, verificar: ls.filter(function (x) { return x.estado === FONDO_EST.VER; }).length,
      autorizar: ls.filter(function (x) { return x.estado === FONDO_ADM; }).length, aprobar: ls.filter(function (x) { return x.estado === FONDO_EST.PED; }).length,
      pagar: ls.filter(function (x) { return x.estado === FONDO_EST.APR; }).length } };
  r.botonPedir = primer === FONDO_EST.VER ? 'Enviar al director operativo' : primer === FONDO_ADM ? 'Enviar al administrador' : 'Enviar a finanzas para aprobar';
  r.ayudaPedir = primer === FONDO_EST.VER ? 'Su solicitud la verifica el director operativo, la autoriza el administrador, la aprueba el director financiero y contabilidad entrega el dinero.'
    : yo.rol === 'finanzas' ? 'Lo que usted pide lo autoriza el administrador y después contabilidad lo paga.'
    : 'Lo que usted pide lo autoriza el administrador, lo aprueba el director financiero y después contabilidad lo paga.';
  r.ayudaSinDinero = 'Si es un traslado o algo con lo que ya contamos, no necesita dinero: solo pasa' + (primer === FONDO_EST.VER ? ' por el director operativo y' : '') + ' por el administrador, y al autorizarse le regresa para que proceda.';
  return r;
}
function pedirFondos(cred, d) {
  permiteCmo();
  var yo = quien(cred), rol = rolDe(yo.nombre) || yo.rol;
  if (yo.rol !== 'operaciones' && !yo.esAdmin && yo.rol !== 'finanzas' && yo.rol !== 'gerente' && !yo.cmo && !esEquipoBodega(yo)) throw new Error('Las solicitudes las hacen los gerentes, la bodega, producción, los directores y el administrador.');
  d = d || {};
  var conDinero = !(d.dinero === false || String(d.dinero).toLowerCase() === 'no');
  if (!(yo.cmo ? FONDO_TIPOS_CMO[d.tipo] : esEquipoBodega(yo) ? FONDO_TIPOS_BODEGA[d.tipo] : FONDO_TIPOS[d.tipo])) throw new Error('Escoja qué tipo de solicitud es.');
  if (yo.rol === 'gerente' || yo.cmo) d.unidad = nombreUnidad(yo.sucursal);
  if (esEquipoBodega(yo)) d.unidad = nombreUnidad('bodega');
  if (yo.cmo && conDinero && d.tipo !== 'Sueldo' && !(d.archivo && d.archivo.datos)) throw new Error('Adjunte la factura (PDF o foto): contabilidad la necesita para pagar.');
  var u = null; UNIDADES.forEach(function (x) { if (x.nombre === d.unidad) u = x; });
  if (!u) throw new Error('Escoja la sucursal.');
  var que = String(d.que || '').trim().slice(0, 300);
  if (que.length < 8) throw new Error('Describa qué se necesita y por qué.');
  var monto = conDinero ? r2(String(d.monto || '').replace(/[Q,\s]/g, '')) : 0;
  if (conDinero && !(monto > 0)) throw new Error('Escriba cuánto dinero se necesita (aproximado). Si no hace falta dinero, marque «No necesita dinero».');
  var urg = ['Normal', 'Urgente'].indexOf(d.urgencia) >= 0 ? d.urgencia : 'Normal';
  var arch = d.archivo && d.archivo.datos ? guardaArchivo(d.archivo, 'Solicitud de ' + yo.nombre + (yo.cmo ? ' · factura' : ' · cotización'), yo.cmo ? 'La factura' : '') : { url: '', nombre: '' };
  var estado0 = estadoInicialFondo(yo.esAdmin ? 'admin' : rol, conDinero);
  var lock = LockService.getScriptLock();
  lock.waitLock(15000); _LEE = {};
  var numero;
  try {
    var max = 0;
    filasFondos().forEach(function (x) { var m = /^(?:SD|RQ|SO)-(\d+)$/.exec(x.numero); if (m) max = Math.max(max, +m[1]); });
    numero = 'SO-' + ('0000' + (max + 1)).slice(-4);
    hojaFondos().appendRow([numero, new Date(), yo.nombre, d.tipo, u.nombre, que, String(d.proveedor || '').slice(0, 60), monto, urg, estado0,
      '', '', '', '', '', '', '', '', '', '', '', '', arch.url, arch.nombre, '', '', '', '', '', '', '', conDinero ? 'Sí' : 'No', '', '', '', '']);
  } finally { lock.releaseLock(); }
  var r = listaFondos(cred); r.ok = true;
  var a = estado0 === FONDO_EST.VER ? ' enviada al director operativo para que la verifique' : estado0 === FONDO_ADM ? ' enviada al administrador para que la autorice'
    : estado0 === FONDO_EST.PED ? ' enviada al director financiero para que apruebe el pago' : ' registrada y autorizada';
  r.mensaje = numero + a + (conDinero ? ': ' + dinero(monto) : ' (no necesita dinero)') + '.';
  return r;
}
function verificarFondos(cred, numero, confirmar, comentario) {
  var yo = quien(cred);
  if (yo.rol !== 'operaciones' && !yo.esAdmin) throw new Error('Lo de los gerentes, la bodega y producción lo verifica el director operativo.');
  var x = buscaFondo(numero);
  if (x.estado !== FONDO_EST.VER) throw new Error('Esa solicitud ya no está por verificar (' + x.estado.toLowerCase() + ').');
  comentario = String(comentario || '').trim().slice(0, 200);
  if (!confirmar && comentario.length < 4) throw new Error('Escriba por qué se rechaza.');
  var h = hojaFondos();
  h.getRange(x.fila, 25, 1, 3).setValues([[yo.nombre, new Date(), comentario]]);
  if (confirmar) h.getRange(x.fila, 10).setValue(estadoSiguienteFondo(x, 'ver'));
  else { h.getRange(x.fila, 10, 1, 4).setValues([[FONDO_EST.RECH, yo.nombre, new Date(), comentario]]); h.getRange(x.fila, 36).setValue('ver'); }
  var r = listaFondos(cred); r.ok = true;
  r.mensaje = numero + (confirmar ? ' verificada. Ahora la autoriza el administrador.' : ' rechazada. ' + x.por + ' lo verá en su campanita.');
  return r;
}
/** El administrador autoriza (o rechaza). Sin dinero, aquí termina el camino y la solicitud regresa a quien la pidió para que proceda. */
function autorizarFondos(cred, numero, autorizar, comentario) {
  var yo = quien(cred);
  if (!yo.esAdmin) throw new Error('Las solicitudes las autoriza el administrador.');
  var x = buscaFondo(numero);
  if (x.estado !== FONDO_ADM) throw new Error('Esa solicitud no está por autorizar (' + estadoTxtFondo(x).toLowerCase() + ').');
  comentario = String(comentario || '').trim().slice(0, 200);
  if (!autorizar && comentario.length < 4) throw new Error('Escriba por qué se rechaza.');
  var h = hojaFondos();
  h.getRange(x.fila, 33, 1, 3).setValues([[yo.nombre, new Date(), comentario]]);
  var sig = autorizar ? estadoSiguienteFondo(x, 'adm') : FONDO_EST.RECH;
  if (autorizar) h.getRange(x.fila, 10).setValue(sig);
  else { h.getRange(x.fila, 10, 1, 4).setValues([[FONDO_EST.RECH, yo.nombre, new Date(), comentario]]); h.getRange(x.fila, 36).setValue('adm'); }
  var r = listaFondos(cred); r.ok = true;
  r.mensaje = numero + (!autorizar ? ' rechazada. ' + x.por + ' lo verá en su campanita.' : sig === FONDO_AUT ? ' autorizada. ' + x.por + ' ya puede proceder.'
    : sig === FONDO_EST.PED ? ' autorizada. Ahora la aprueba el director financiero.' : ' autorizada. Contabilidad entrega el dinero.');
  return r;
}
function aprobarFondos(cred, numero, aprobar, comentario) {
  var yo = quien(cred);
  if (yo.rol !== 'finanzas' && !yo.esAdmin) throw new Error('El pago lo aprueba el director financiero.');
  var x = buscaFondo(numero);
  if (x.estado !== FONDO_EST.PED) throw new Error(x.estado === FONDO_ADM || x.estado === FONDO_EST.VER ? 'Todavía falta que la autoricen (' + estadoTxtFondo(x).toLowerCase() + ').' : 'Esa solicitud ya está ' + estadoTxtFondo(x).toLowerCase() + '.');
  comentario = String(comentario || '').trim().slice(0, 200);
  if (!aprobar && comentario.length < 4) throw new Error('Escriba por qué se rechaza.');
  hojaFondos().getRange(x.fila, 10, 1, 4).setValues([[aprobar ? FONDO_EST.APR : FONDO_EST.RECH, yo.nombre, new Date(), comentario]]);
  if (!aprobar) hojaFondos().getRange(x.fila, 36).setValue('fin');
  var r = listaFondos(cred); r.ok = true; r.mensaje = numero + (aprobar ? ' aprobada. El contador entrega el dinero.' : ' rechazada. ' + x.por + ' lo verá en su campanita.');
  return r;
}
/** Quien pidió (sin dinero) marca que ya lo hizo. */
function hechaFondos(cred, numero, nota) {
  var yo = quien(cred), x = buscaFondo(numero);
  if (x.por !== yo.nombre && !yo.esAdmin) throw new Error('La marca ' + x.por + ', quien la pidió.');
  if (x.estado !== FONDO_AUT) throw new Error('Esa solicitud no está autorizada, o ya se marcó.');
  hojaFondos().getRange(x.fila, 10).setValue(FONDO_EST.FIN);
  hojaFondos().getRange(x.fila, 21, 1, 2).setValues([[new Date(), String(nota || '').slice(0, 200)]]);
  var r = listaFondos(cred); r.ok = true; r.mensaje = numero + ' marcada como hecha.';
  return r;
}

/* ════════════ DANIEL: VE TODO, NO CAMBIA NADA ════════════ */
function soloVerCal(ac) {
  if (ac && ac.soloVer) throw new Error('Usted puede ver los calendarios y las planillas, pero no cambiarlos. Si ve algo raro, pregúntele al director o avise al administrador.');
}

/* ════════════ PRESUPUESTOS: REINICIAR UN MES ════════════
 * Borra todo lo del mes (presupuesto por sucursal y rubro, su aprobación, y los presupuestos de área) para volver a empezar.
 * El presupuesto es un indicador, no el pago real: los pagos salen de las planillas y de las solicitudes. Solo el administrador, escribiendo REINICIAR. */
function reiniciarPresupuestos(cred, mes, confirmacion) {
  var yo = exigeAdmin(cred);
  mes = /^\d{4}-\d{2}$/.test(String(mes || '')) ? String(mes) : hoyISO().slice(0, 7);
  if (String(confirmacion || '').trim().toUpperCase() !== 'REINICIAR') throw new Error('Para borrar los presupuestos de ' + periodoTxt(mes) + ', escriba REINICIAR.');
  var mesDe = function (v) { return v instanceof Date ? Utilities.formatDate(v, ZONA, 'yyyy-MM') : String(v == null ? '' : v).slice(0, 7); };
  var borra = function (h, ancho) {
    if (!h || h.getLastRow() < 2) return 0;
    var vals = h.getRange(2, 1, h.getLastRow() - 1, Math.min(ancho, h.getLastColumn())).getValues(), k = 0;
    for (var i = vals.length - 1; i >= 0; i--) if (mesDe(vals[i][0]) === mes) { h.deleteRow(i + 2); k++; }
    return k;
  };
  var lock = LockService.getScriptLock(); lock.waitLock(20000); _LEE = {};
  var n = 0;
  try { n += borra(hojaPres(), H_PRES.length); n += borra(hojaPresEstado(), H_PRESE.length); n += borra(hojaPArea(), H_PAREA.length); n += borra(hojaPAreaEst(), H_PAREA_EST.length); }
  finally { lock.releaseLock(); }
  olvidaLectura();
  var r = presupuestoMes(cred, mes); r.ok = true;
  r.mensaje = 'Presupuestos de ' + periodoTxt(mes) + ' reiniciados (' + n + (n === 1 ? ' renglón borrado' : ' renglones borrados') + '). Se puede volver a armar cuando haya datos reales.';
  return r;
}

/* ════════════ HISTORIAL DE INVENTARIO (auditoría) ════════════
 * Por sucursal y por mes, producto por producto: cuánto contaron antes de pedir, cuánto pidieron, cuánto les mandaron, cuánto recibieron,
 * con cuánto quedaron (lo contado + lo recibido) y cuánto gastaron hasta el siguiente conteo. Para revisar que no pidan de más.
 * Lo ven el administrador, el director operativo y el director financiero. */
function veHistorialInv(yo) { return !!yo && (yo.esAdmin || yo.rol === 'operaciones' || yo.rol === 'finanzas'); }
function historialInventario(cred, unidadId, mes) {
  var yo = quien(cred);
  if (!veHistorialInv(yo)) throw new Error('El historial de inventario lo ven el administrador y los directores.');
  return historialCalc(unidadId, mes);
}
function historialCalc(unidadId, mes) {
  var suc = UNIDADES.filter(function (u) { return u.vende; }), u = suc.filter(function (x) { return x.id === unidadId; })[0] || suc[0];
  mes = /^\d{4}-\d{2}$/.test(String(mes || '')) ? String(mes) : hoyISO().slice(0, 7);
  var sols = filasSol().filter(function (x) { return x.sucursal === u.nombre && x.estado !== SOL.ANU; }), porNum = {};
  sols.forEach(function (x) { porNum[x.numero] = x; });
  // los ciclos de cada producto: todos los pedidos (para saber el siguiente conteo aunque caiga en otro mes)
  var ciclos = {};
  lineasSol().forEach(function (l) {
    var x = porNum[l.numero]; if (!x || !(l.pedido > 0 || l.enviado > 0)) return;
    (ciclos[l.clave] = ciclos[l.clave] || { codigo: l.codigo, nombre: l.nombre, medida: l.medida, filas: [] }).filas.push({
      numero: x.numero, fecha: x.enviadaEn, entrega: diaAbasto(x), contado: l.existencia, pedido: l.pedido, enviado: l.enviado, recibido: l.recibido,
      recibidoEn: x.recibidaEn, estado: x.estado, pendiente: pendDeLinea(l) });
  });
  var productos = [];
  Object.keys(ciclos).forEach(function (k) {
    var c = ciclos[k]; c.filas.sort(function (a, b) { return a.fecha < b.fecha ? -1 : 1; });
    c.filas.forEach(function (f, i) {
      f.quedo = f.recibido != null ? r3(f.contado + f.recibido) : null;          // con lo que quedó la sucursal al recibir
      var sig = c.filas[i + 1];
      f.gasto = f.quedo != null && sig ? r3(f.quedo - sig.contado) : null;      // lo que se usó hasta el siguiente conteo
    });
    var delMes = c.filas.filter(function (f) { return f.entrega.slice(0, 7) === mes; });
    if (!delMes.length) return;
    var suma = function (k2) { return r3(delMes.reduce(function (a, f) { return a + (Number(f[k2]) || 0); }, 0)); };
    productos.push({ clave: k, codigo: c.codigo, nombre: c.nombre, medida: c.medida, ciclos: delMes,
      totales: { pedido: suma('pedido'), enviado: suma('enviado'), recibido: suma('recibido'), gasto: suma('gasto') },
      alerta: delMes.some(function (f) { return f.contado > 0 && f.pedido > 0 && f.contado >= f.pedido * 2; }) });   // pidió teniendo el doble de lo que pidió
  });
  productos.sort(function (a, b) { return a.nombre.localeCompare(b.nombre, 'es'); });
  var meses = []; for (var i = 0; i < 6; i++) { var d = new Date(hoyISO().slice(0, 7) + '-15T12:00:00'); d.setMonth(d.getMonth() - i); meses.push(Utilities.formatDate(d, ZONA, 'yyyy-MM')); }
  return { unidad: { id: u.id, nombre: u.nombre }, sucursales: suc.map(function (x) { return { id: x.id, nombre: x.nombre }; }), mes: mes, mesTxt: periodoTxt(mes), meses: meses,
    productos: productos, pedidos: sols.filter(function (x) { return diaAbasto(x).slice(0, 7) === mes; }).length, calendario: calendarioAbasto(mes),
    cierres: suc.map(function (x) { var c = filasCierres().filter(function (y) { return y.mes === mes && y.sucursal === x.nombre; })[0]; return { id: x.id, nombre: x.nombre, cerrado: !!c, por: c ? c.por : '', en: c ? c.en : '' }; }) };
}

/* ════════════ MENÚS POR ROL ════════════
 * El administrador decide qué menús y secciones ve cada rol (no cada persona): lo que se apaga desaparece del menú de todos los de ese rol.
 * El contador es su propio perfil (dentro del rol «registro», los que confirman ingresos). Por defecto el contador solo ve lo de pagar. */
var H_MENUS = ['Perfil', 'Apagados (JSON)', 'Cambiado por', 'Cambiado en'];
var PERFILES_MENU = [['gerente', 'Gerente'], ['bodega', 'Bodega'], ['produccion', 'Producción'], ['operaciones', 'Director operativo'], ['finanzas', 'Director financiero'],
  ['contador', 'Contador'], ['registro', 'Registro (caja)'], ['cmo', 'Marketing (CMO)'], ['auxiliar', 'Auxiliar'], ['dueno', 'Dueños']];
var MENUS_DEFECTO = { contador: ['inicio', 'registrar', 'reportes', 'conta:resumen'] };     // el contador: solo pagos (Daniel)
function hojaMenus(ss) {
  ss = ss || libro();
  var h = ss.getSheetByName('Menús por rol');
  if (!h) { h = hojaLimpia(ss, 'Menús por rol', H_MENUS); h.getRange('D:D').setNumberFormat('dd/mm/yyyy hh:mm'); h.setColumnWidth(2, 420); }
  return h;
}
function perfilDe(yo) { return yo.esAdmin ? 'admin' : (yo.rol === 'registro' && yo.recibe) ? 'contador' : yo.rol; }
function menusApagados() {
  var out = JSON.parse(JSON.stringify(MENUS_DEFECTO)), h = libro().getSheetByName('Menús por rol');
  if (h && h.getLastRow() > 1) leeTodo(h, H_MENUS.length).forEach(function (r) { var k = String(r[0] || ''); if (!k) return; try { out[k] = JSON.parse(String(r[1] || '[]')) || []; } catch (e) {} });
  return out;
}
/** Lo que tiene apagado mi perfil (la app lo esconde del menú). */
function misMenus(cred) {
  permiteAux(); permiteCmo();
  var yo = quien(cred), p = perfilDe(yo);
  return { perfil: p, apagados: p === 'admin' ? [] : (menusApagados()[p] || []) };
}
function menusRol(cred) {
  exigeAdmin(cred);
  var ap = menusApagados(), us = usuariosCache().filter(function (u) { return u.activo; });
  return { perfiles: PERFILES_MENU.map(function (x) {
      var n = us.filter(function (u) { return x[0] === 'contador' ? (u.rol === 'registro' && u.confirma) : x[0] === 'registro' ? (u.rol === 'registro' && !u.confirma) : u.rol === x[0]; }).length;
      return { id: x[0], nombre: x[1], usuarios: n, apagados: ap[x[0]] || [] }; }) };
}
function guardarMenusRol(cred, perfil, apagados) {
  var yo = exigeAdmin(cred);
  if (!PERFILES_MENU.some(function (x) { return x[0] === perfil; })) throw new Error('Ese rol no existe.');
  var lista = (apagados || []).map(function (x) { return String(x).trim(); })
    .filter(function (x, i, a) { return /^[a-z]+(:[a-z0-9]+)?$/.test(x) && x !== 'config' && a.indexOf(x) === i; }).slice(0, 80);     // Configuración nunca se apaga
  var lock = LockService.getScriptLock(); lock.waitLock(15000); _LEE = {};
  try {
    var h = hojaMenus(), fila = 0;
    if (h.getLastRow() > 1) leeTodo(h, H_MENUS.length).forEach(function (r, i) { if (String(r[0]) === perfil) fila = i + 2; });
    var val = [[perfil, JSON.stringify(lista), yo.nombre, new Date()]];
    if (fila) h.getRange(fila, 1, 1, 4).setValues(val); else h.appendRow(val[0]);
  } finally { lock.releaseLock(); }
  var r = menusRol(cred); r.ok = true;
  var nom = PERFILES_MENU.filter(function (x) { return x[0] === perfil; })[0][1];
  r.mensaje = 'Menús de «' + nom + '» guardados' + (lista.length ? ' (' + lista.length + ' apagados)' : '') + '. Lo verán la próxima vez que abran la app.';
  return r;
}

/* ════════════ ABASTECIMIENTO DEL MES: calendario con semáforo, detalle de cada pedido, PDF por sucursal y cierre de mes ════════════
 *  A (verde)    Abastecido sin inconvenientes.
 *  P (amarillo) Abastecido, pero quedan pendientes o hubo un inconveniente (diferencias, faltante).
 *  F (rojo)     Falta algún paso para que se abastezca (sin aceptar, sin despachar, en camino sin recibir).
 *  El cierre de mes: al terminar el mes cada sucursal cuenta lo que queda. Eso es lo que «debería haber» contra lo contado, y es el punto de partida del mes siguiente. */
var H_CIERRE = ['Mes', 'Sucursal', 'Código', 'Producto', 'Medida', 'Debía haber', 'Contado', 'Diferencia', 'Cerrado por', 'Cerrado en', 'Nota'];
function hojaCierres(ss) {
  ss = ss || libro();
  var h = ss.getSheetByName('Cierres de mes');
  if (!h) { h = hojaLimpia(ss, 'Cierres de mes', H_CIERRE); h.getRange('A:A').setNumberFormat('@'); h.getRange('C:C').setNumberFormat('@'); h.getRange('J:J').setNumberFormat('dd/mm/yyyy hh:mm'); h.setColumnWidth(4, 240); }
  return h;
}
function mesDeCelda(v) { return v instanceof Date ? Utilities.formatDate(v, ZONA, 'yyyy-MM') : String(v == null ? '' : v).slice(0, 7); }
function filasCierres() {
  var h = libro().getSheetByName('Cierres de mes'); if (!h || h.getLastRow() < 2) return [];
  return leeTodo(h, H_CIERRE.length).map(function (r, i) {
    return { fila: i + 2, mes: mesDeCelda(r[0]), sucursal: String(r[1] || ''), codigo: String(r[2] || ''), nombre: String(r[3] || ''), medida: String(r[4] || ''),
      debia: r[5] === '' || r[5] == null ? null : Number(r[5]), contado: Number(r[6]) || 0, dif: Number(r[7]) || 0, por: String(r[8] || ''), en: fmtSello(r[9]), nota: String(r[10] || '') };
  }).filter(function (x) { return x.mes && x.codigo; });
}
function finDeMes(mes) { var p = mes.split('-'); var d = new Date(+p[0], +p[1], 0); return Utilities.formatDate(d, ZONA, 'yyyy-MM-dd'); }
function mesAnterior(mes) { var p = mes.split('-'), d = new Date(+p[0], +p[1] - 2, 15); return Utilities.formatDate(d, ZONA, 'yyyy-MM'); }
function diaAbasto(x) { return String(x.entrega || x.enviadaEn || '').slice(0, 10); }
var ABASTO_ORDEN = { A: 0, P: 1, F: 2 };
/** La letra de un pedido: A, P o F, y por qué. */
function abastoDe(x, lineas) {
  if (x.estado !== SOL.REC && x.estado !== SOL.DIF) {
    var q = x.estado === SOL.ENV ? 'bodega todavía no lo acepta' : x.estado === SOL.ACE ? 'bodega lo está preparando o falta que salga el camión' : 'va en camino, falta que lo reciban';
    return { e: 'F', txt: 'Falta un paso: ' + q + '.' };
  }
  var pend = lineas.filter(function (l) { return pendDeLinea(l) > 0.0005; }).length, dif = lineas.filter(function (l) { return Math.abs(l.diferencia || 0) > 0.0005 || (l.pendCanPor && l.pend > 0); }).length;
  if (x.estado === SOL.DIF || pend || dif) {
    var m = []; if (pend) m.push(pend + (pend === 1 ? ' producto pendiente de entrega' : ' productos pendientes de entrega')); if (dif) m.push(dif + (dif === 1 ? ' producto con diferencia' : ' productos con diferencias'));
    return { e: 'P', txt: 'Abastecido, pero ' + (m.length ? m.join(' y ') : 'hubo un inconveniente') + '.' };
  }
  return { e: 'A', txt: 'Abastecido sin inconvenientes.' };
}
/** El calendario del mes: filas = sucursales, columnas = los días con abastecimiento. */
function calendarioAbasto(mes) {
  var sucs = UNIDADES.filter(function (u) { return u.vende; }), sols = filasSol().filter(function (x) { return x.estado !== SOL.ANU && diaAbasto(x).slice(0, 7) === mes; });
  var lins = {}; lineasSol().forEach(function (l) { (lins[l.numero] = lins[l.numero] || []).push(l); });
  var dias = {}, filas = sucs.map(function (u) {
    var celdas = {}, tot = { A: 0, P: 0, F: 0 };
    sols.filter(function (x) { return x.sucursal === u.nombre; }).forEach(function (x) {
      var f = diaAbasto(x), a = abastoDe(x, lins[x.numero] || []); dias[f] = true;
      var c = celdas[f] || (celdas[f] = { e: 'A', n: 0, numeros: [] }); c.n++; c.numeros.push(x.numero);
      if (ABASTO_ORDEN[a.e] > ABASTO_ORDEN[c.e]) c.e = a.e; tot[a.e]++;
    });
    return { id: u.id, nombre: u.nombre, celdas: celdas, total: tot };
  });
  return { dias: Object.keys(dias).sort(), filas: filas };
}
function detalleAbastecimiento(cred, numero) {
  var yo = quien(cred);
  if (!veHistorialInv(yo)) throw new Error('El historial de abastecimiento lo ven el administrador y los directores.');
  var x = null; filasSol().forEach(function (y) { if (y.numero === String(numero)) x = y; });
  if (!x) throw new Error('No se encontró el pedido ' + numero + '.');
  var ls = lineasSol(x.numero).filter(function (l) { return l.pedido > 0 || l.enviado > 0 || l.recibido > 0; }), a = abastoDe(x, ls);
  return { numero: x.numero, sucursal: x.sucursal, estado: x.estado, letra: a.e, txt: a.txt, entrega: x.entrega, enviadaEn: x.enviadaEn, enviadaPor: x.enviadaPor, aceptadaEn: x.aceptadaEn, aceptadaPor: x.aceptadaPor,
    despachadaEn: x.despachadaEn, despachadaPor: x.despachadaPor, recibidaEn: x.recibidaEn, recibidaPor: x.recibidaPor, traslado: x.traslado, nota: x.nota,
    lineas: ls.map(function (l) { return { nombre: l.nombre, medida: l.medida, contado: l.existencia, pedido: l.pedido, enviado: l.enviado, recibido: l.recibido, diferencia: l.diferencia, difEstado: l.difEstado,
      pendiente: pendDeLinea(l), cancelado: !!l.pendCanPor, motivo: l.pendMotivo }; }) };
}
/** Lo que se necesita para el PDF: el calendario, los productos con sus movimientos, y el cierre. */
function datosHistorialPdf(cred, unidadId, mes) {
  var h = historialInventario(cred, unidadId, mes), u = h.unidad;
  var cierres = filasCierres().filter(function (c) { return c.mes === h.mes && c.sucursal === u.nombre; }), porCod = {};
  cierres.forEach(function (c) { porCod[c.codigo] = c; });
  var fila = h.calendario.filas.filter(function (f) { return f.id === u.id; })[0];
  var lins = {}; lineasSol().forEach(function (l) { (lins[l.numero] = lins[l.numero] || []).push(l); });
  var pedidos = filasSol().filter(function (x) { return x.sucursal === u.nombre && x.estado !== SOL.ANU && diaAbasto(x).slice(0, 7) === h.mes; })
    .sort(function (a, b) { return diaAbasto(a) < diaAbasto(b) ? -1 : 1; }).map(function (x) { var ls = lins[x.numero] || [], a = abastoDe(x, ls);
      return { numero: x.numero, dia: diaAbasto(x), letra: a.e, txt: a.txt, productos: ls.filter(function (l) { return l.pedido > 0 || l.enviado > 0; }).length, recibidaPor: x.recibidaPor, estado: x.estado }; });
  var productos = h.productos.map(function (p) {
    var ult = null, c = porCod[p.codigo]; for (var i = p.ciclos.length - 1; i >= 0; i--) if (p.ciclos[i].quedo != null) { ult = p.ciclos[i]; break; }
    var debia = ult ? ult.quedo : null;
    return { nombre: p.nombre, medida: p.medida, ciclos: p.ciclos, totales: p.totales, debia: debia, cierre: c ? c.contado : null, uso: debia != null && c ? r3(debia - c.contado) : null };
  });
  var sinMov = cierres.filter(function (c) { return !h.productos.some(function (p) { return p.codigo === c.codigo; }); })
    .map(function (c) { return { nombre: c.nombre, medida: c.medida, ciclos: [], totales: { pedido: 0, enviado: 0, recibido: 0, gasto: 0 }, debia: c.debia, cierre: c.contado, uso: c.debia != null ? r3(c.debia - c.contado) : null }; });
  return { sucursal: u.nombre, mes: h.mes, mesTxt: h.mesTxt, pedidos: pedidos, productos: productos.concat(sinMov).sort(function (a, b) { return a.nombre.localeCompare(b.nombre, 'es'); }),
    cierre: cierres.length ? { por: cierres[0].por, en: cierres[0].en, nota: cierres[0].nota, productos: cierres.length } : null, resumen: fila ? fila.total : { A: 0, P: 0, F: 0 } };
}
function htmlHistorialPdf(d) {
  var h = [], n = function (v) { return v == null ? '—' : String(Math.round(v * 100) / 100).replace('.', ','); };
  var col = { A: '#1F7A4D', P: '#B8860B', F: '#B3261E' }, nom = { A: 'Abastecido', P: 'Con pendientes o inconveniente', F: 'Falta un paso' };
  h.push('<html><head><meta charset="utf-8"><style>');
  h.push('body{font-family:Arial,Helvetica,sans-serif;font-size:10px;color:#16223A;margin:22px}');
  h.push('table.band{width:100%;background:#14284B;color:#fff;border-bottom:4px solid #F2B705;margin:0 0 10px;border-collapse:collapse}');
  h.push('table.band td{border:0;padding:10px 12px;vertical-align:middle}.band img{width:40px}.band td.r{text-align:right;font-size:12px;color:#F7F2E7;font-weight:bold}');
  h.push('h1{font-size:19px;color:#fff;margin:0;letter-spacing:1px}.band .co{color:#F2B705;font-size:9px;letter-spacing:2px;font-weight:bold}');
  h.push('h2{font-size:13px;color:#14284B;margin:14px 0 6px}');
  h.push('table.t{width:100%;border-collapse:collapse;margin:0 0 10px}table.t th,table.t td{border:1px solid #CFC6B2;padding:4px 6px;font-size:9.5px;text-align:right}');
  h.push('table.t th{background:#14284B;color:#F2B705}table.t td.l,table.t th.l{text-align:left}table.t tr.p td{background:#F7F2E7;font-weight:bold}');
  h.push('table.t td.fin{background:#EEF2FA;font-weight:bold}.note{color:#6F7686;font-size:9px}');
  h.push('</style></head><body>');
  h.push('<table class="band"><tr><td style="width:52px"><img src="' + LOGO_PNG + '"></td><td><h1>REY PIZZA</h1><div class="co">CAIX, S.A.</div></td><td class="r">Historial de abastecimiento<br>' + htm(d.sucursal) + ' · ' + htm(d.mesTxt) + '</td></tr></table>');
  h.push('<h2>Abastecimientos del mes</h2>');
  if (!d.pedidos.length) h.push('<div class="note">No hubo pedidos de esta sucursal en este mes.</div>');
  else {
    h.push('<table class="t"><tr><th class="l">Fecha</th><th class="l">Pedido</th><th class="l">Estado</th><th>Productos</th><th class="l">Lo recibió</th></tr>');
    d.pedidos.forEach(function (p) { h.push('<tr><td class="l">' + htm(dia(p.dia)) + '</td><td class="l">' + htm(p.numero) + '</td><td class="l" style="color:' + col[p.letra] + ';font-weight:bold">' + p.letra + ' · ' + htm(p.txt) + '</td><td>' + p.productos + '</td><td class="l">' + htm(p.recibidaPor || '—') + '</td></tr>'); });
    h.push('</table><div class="note">A = ' + nom.A + ' · P = ' + nom.P + ' · F = ' + nom.F + '. Resumen: ' + d.resumen.A + ' sin inconvenientes, ' + d.resumen.P + ' con pendientes, ' + d.resumen.F + ' sin terminar.</div>');
  }
  h.push('<h2>Producto por producto</h2><div class="note">«Total» = lo que contaron + lo que recibieron. «Debería haber» es el total del último pedido del mes. «Cierre» es lo que de verdad contaron al cerrar el mes; la diferencia es lo que se usó desde el último pedido.</div>');
  d.productos.forEach(function (p) {
    h.push('<table class="t" style="page-break-inside:avoid;margin-top:8px"><tr><th class="l" colspan="7">' + htm(p.nombre) + ' <span style="font-weight:normal;color:#C9D6EC">(' + htm(p.medida) + ')</span></th></tr>');
    if (p.ciclos.length) {
      h.push('<tr class="p"><td class="l">Fecha</td><td class="l">Pedido</td><td>Contó</td><td>Pidió</td><td>Se envió</td><td>Se recibió</td><td>Total</td></tr>');
      p.ciclos.forEach(function (c) { h.push('<tr><td class="l">' + htm(dia(c.entrega || c.fecha)) + '</td><td class="l">' + htm(c.numero) + '</td><td>' + n(c.contado) + '</td><td>' + n(c.pedido) + '</td><td>' + n(c.enviado) + '</td><td>' + n(c.recibido) + '</td><td class="fin">' + n(c.quedo) + '</td></tr>'); });
    }
    h.push('<tr class="p"><td class="l" colspan="3">Debería haber: <b>' + n(p.debia) + '</b></td><td colspan="2">Cierre contado: <b>' + (p.cierre == null ? 'pendiente' : n(p.cierre)) + '</b></td><td colspan="2">Se usó: <b>' + (p.uso == null ? '—' : n(p.uso)) + '</b></td></tr></table>');
  });
  if (!d.productos.length) h.push('<div class="note">No hay productos con movimientos.</div>');
  h.push('<div class="note" style="margin-top:14px">' + (d.cierre ? 'Cierre de mes registrado por ' + htm(d.cierre.por) + ' el ' + htm(String(d.cierre.en).slice(0, 10)) + (d.cierre.nota ? ' · ' + htm(d.cierre.nota) : '') + '.' : 'El cierre de mes de esta sucursal todavía no se ha hecho.') +
    ' Generado el ' + Utilities.formatDate(new Date(), ZONA, 'dd/MM/yyyy HH:mm') + ' desde la app REY PIZZA · CAIX, S.A.</div></body></html>');
  return h.join('');
}
function pdfHistorialInventario(cred, unidadId, mes) {
  var d = datosHistorialPdf(cred, unidadId, mes);
  var nombre = 'Historial de abastecimiento ' + d.sucursal + ' ' + d.mes + '.pdf';
  var blob = Utilities.newBlob(htmlHistorialPdf(d), 'text/html', 'h.html').getAs('application/pdf').setName(nombre);
  return { ok: true, nombre: nombre, tipo: 'application/pdf', base64: Utilities.base64Encode(blob.getBytes()) };
}

/* ── cierre de mes ── */
function totalConteoProducto(b, x) {
  var lee = function (v, que) {
    var s = String(v == null ? '' : v).trim(); if (s === '') return 0;
    var n = Number(s.replace(',', '.')); if (!isFinite(n) || n < 0 || n > 100000) throw new Error('La cantidad de «' + b.nombre + '» (' + que + ') no es válida.');
    return r3(n);
  };
  var e = lee(x.e, b.unidad), s = tieneSuelto(b) ? lee(x.s, b.suelto) : 0, m = b.menor ? lee(x.m, b.menor) : 0, p = b.pesoCada ? lee(x.p, b.pesoEn) : 0;
  return r3(e + (tieneSuelto(b) ? s / b.porUnidad : 0) + (b.menor ? m / (b.porUnidad * b.porSuelto) : 0) + (b.pesoCada ? p / (b.pesoCada * nivelPeso(b).n) : 0));
}
function llenoConteoProducto(b, x) { var h = function (v) { return String(v == null ? '' : v).trim() !== ''; }; return h(x.e) || h(x.s) || (b.menor && h(x.m)) || (b.pesoCada && h(x.p)); }
function puedeCerrarMes(yo, unidadId) { return yo.esAdmin || yo.rol === 'operaciones' || (yo.rol === 'gerente' && yo.sucursal === unidadId); }
function cierreAbierto(mes, hoy) { return mes < hoy.slice(0, 7) || (mes === hoy.slice(0, 7) && hoy >= masDias(finDeMes(mes), -2)); }      // desde 2 días antes de que acabe el mes
function pantallaCierreMes(cred, unidadId, mes) {
  var yo = quien(cred), hoy = hoyISO();
  if (!(yo.esAdmin || yo.rol === 'operaciones' || yo.rol === 'finanzas' || yo.rol === 'gerente')) throw new Error('Su usuario no ve el cierre de mes.');
  var suc = UNIDADES.filter(function (u) { return u.vende; });
  if (yo.rol === 'gerente') unidadId = yo.sucursal;
  var u = suc.filter(function (x) { return x.id === unidadId; })[0] || suc[0];
  mes = /^\d{4}-\d{2}$/.test(String(mes || '')) ? String(mes) : (hoy.slice(8) >= '25' || hoy.slice(8) <= '05' ? (hoy.slice(8) <= '05' ? mesAnterior(hoy.slice(0, 7)) : hoy.slice(0, 7)) : hoy.slice(0, 7));
  var guardado = {}; filasCierres().forEach(function (c) { if (c.mes === mes && c.sucursal === u.nombre) guardado[c.codigo] = c; });
  var hist = historialCalc(u.id, mes), debia = {};
  hist.productos.forEach(function (p) {                       // lo que debería haber: el total del último pedido ya recibido del mes (en la medida base del producto)
    for (var i = p.ciclos.length - 1; i >= 0; i--) if (p.ciclos[i].quedo != null) { if (debia[p.codigo] == null || p.medida === (productosInv().filter(function (b) { return b.codigo === p.codigo; })[0] || {}).unidad) debia[p.codigo] = p.ciclos[i].quedo; break; } });
  var exi = {}; try { exi = estadoInventario(u.id).existencia || {}; } catch (e) {}
  var prods = productosInv().filter(function (b) { return maneja(b, u.id); }).map(function (b) {
    var g = guardado[b.codigo], d0 = debia[b.codigo] != null ? debia[b.codigo] : (exi[b.codigo] != null ? r3(exi[b.codigo]) : null);
    return { codigo: b.codigo, nombre: b.nombre, grupo: b.grupo, estante: b.estante, unidad: b.unidad, suelto: tieneSuelto(b) ? b.suelto : '', porUnidad: tieneSuelto(b) ? b.porUnidad : 0,
      medida: b.medida, menor: b.menor, porSuelto: b.porSuelto, pesoCada: b.pesoCada, pesoEn: b.pesoEn, pesa: b.pesoCada ? nivelPeso(b).nombre : '', presentacion: b.presentacion,
      debia: d0, contado: g ? g.contado : null }; });
  var meses = []; for (var i = 0; i < 4; i++) { var d = new Date(hoy.slice(0, 7) + '-15T12:00:00'); d.setMonth(d.getMonth() - i); meses.push(Utilities.formatDate(d, ZONA, 'yyyy-MM')); }
  var uno = Object.keys(guardado).map(function (k) { return guardado[k]; })[0];
  return { unidad: { id: u.id, nombre: u.nombre }, sucursales: suc.map(function (x) { return { id: x.id, nombre: x.nombre }; }), mes: mes, mesTxt: periodoTxt(mes), meses: meses,
    abierto: cierreAbierto(mes, hoy), puede: puedeCerrarMes(yo, u.id), fin: finDeMes(mes), productos: prods,
    cerrado: uno ? { por: uno.por, en: uno.en, nota: uno.nota } : null };
}
function guardarCierreMes(cred, d) {
  var yo = quien(cred), hoy = hoyISO(); d = d || {};
  var unidadId = yo.rol === 'gerente' ? yo.sucursal : String(d.unidad || '');
  if (!puedeCerrarMes(yo, unidadId)) throw new Error('El cierre de mes lo hace el gerente de la sucursal (o el director operativo).');
  var u = UNIDADES.filter(function (x) { return x.vende && x.id === unidadId; })[0]; if (!u) throw new Error('Escoja la sucursal.');
  var mes = String(d.mes || ''); if (!/^\d{4}-\d{2}$/.test(mes)) throw new Error('Escoja el mes.');
  if (!cierreAbierto(mes, hoy)) throw new Error('El cierre de ' + periodoTxt(mes) + ' se hace en sus últimos 2 días o después de que termine el mes.');
  var pant = pantallaCierreMes(cred, unidadId, mes), cant = d.cantidades || {}, filas = [], faltan = [];
  var prods = productosInv().filter(function (b) { return maneja(b, u.id); }), porCod = {}; pant.productos.forEach(function (p) { porCod[p.codigo] = p; });
  var ahora = new Date();
  prods.forEach(function (b) {
    var x = cant[b.codigo] || {};
    if (!llenoConteoProducto(b, x)) { faltan.push(b.nombre); return; }
    var total = totalConteoProducto(b, x), debia = porCod[b.codigo] ? porCod[b.codigo].debia : null;
    filas.push([mes, u.nombre, b.codigo, b.nombre, b.unidad, debia == null ? '' : debia, total, debia == null ? '' : r3(total - debia), yo.nombre, ahora, String(d.nota || '').slice(0, 200)]);
  });
  if (faltan.length) throw new Error('Faltan cantidades: ' + faltan.slice(0, 3).join(', ') + (faltan.length > 3 ? ' y ' + (faltan.length - 3) + ' más' : '') + '. Escriba 0 si ya no queda nada.');
  if (!filas.length) throw new Error('Esta sucursal no tiene productos para contar.');
  var lock = LockService.getScriptLock(); lock.waitLock(20000); _LEE = {};
  try {
    var h = hojaCierres();
    for (var r = h.getLastRow(); r >= 2; r--) if (mesDeCelda(h.getRange(r, 1).getValue()) === mes && String(h.getRange(r, 2).getValue()) === u.nombre) h.deleteRow(r);        // un cierre nuevo reemplaza al anterior
    var f0 = h.getLastRow() + 1;
    h.getRange(f0, 1, filas.length, 1).setNumberFormat('@'); h.getRange(f0, 3, filas.length, 1).setNumberFormat('@');
    h.getRange(f0, 1, filas.length, H_CIERRE.length).setValues(filas);
  } finally { lock.releaseLock(); }
  var usados = filas.filter(function (f) { return f[7] !== '' && f[7] < -0.0005; }).length;
  var r2 = pantallaCierreMes(cred, unidadId, mes); r2.ok = true;
  r2.mensaje = 'Cierre de ' + periodoTxt(mes) + ' de ' + u.nombre + ' guardado: ' + filas.length + ' productos. Es el punto de partida del próximo mes.';
  return r2;
}
/** Para el primer «contar y pedir» del mes: lo que se contó al cerrar el mes anterior. */
function cierreRefDe(unidadId, ult, hoy) {
  if (unidadId === 'bodega') return null;
  var u = unidadPorId(unidadId), mesAct = hoy.slice(0, 7), ant = mesAnterior(mesAct);
  if (ult && ult.fecha >= mesAct + '-01') return null;                         // ya hubo un conteo este mes: la referencia solo es para el primero
  var por = {}, hay = false;
  filasCierres().forEach(function (c) { if (c.mes === ant && c.sucursal === u.nombre) { por[c.codigo] = c.contado; hay = true; } });
  return hay ? { mes: ant, mesTxt: periodoTxt(ant), por: por } : null;
}
/** El recordatorio: al gerente, desde 2 días antes de que acabe el mes hasta el día 10; al director, si pasan los primeros días y falta. */
function avisosCierre(yo) {
  var out = [], hoy = hoyISO(), mesAct = hoy.slice(0, 7), dd = +hoy.slice(8);
  var m = hoy >= masDias(finDeMes(mesAct), -2) ? mesAct : (dd <= 10 ? mesAnterior(mesAct) : '');
  if (!m) return out;
  var hechos = {}; filasCierres().forEach(function (c) { if (c.mes === m) hechos[c.sucursal] = true; });
  var suc = UNIDADES.filter(function (u) { return u.vende; });
  if (yo.rol === 'gerente') { var u = suc.filter(function (x) { return x.id === yo.sucursal; })[0];
    if (u && !hechos[u.nombre]) out.push({ id: 'cierre-' + u.id + '-' + m, tipo: 'Cierre de mes', cuando: '', titulo: 'Cierre de mes de ' + periodoTxt(m) + ': cuente lo que queda',
      texto: 'Anote cuánto hay de cada producto. Es el punto de partida del próximo mes.', abrir: { tab: 'cierre' } }); }
  if ((yo.esAdmin || yo.rol === 'operaciones') && m < mesAct && dd >= 4) {
    var f = suc.filter(function (x) { return !hechos[x.nombre]; }).map(function (x) { return x.nombre; });
    if (f.length) out.push({ id: 'cierref-' + m, tipo: 'Cierre de mes', cuando: '', titulo: 'Falta el cierre de ' + periodoTxt(m) + ': ' + f.join(', '), texto: 'Cada gerente tiene que contar lo que queda al terminar el mes.', abrir: { tab: 'hist' } });
  }
  return out;
}

/* ════════════ PRODUCTO COMPLETO: sucursales, proveedores y precio dentro del producto ════════════
 * Lo editan Samuel (bodega), Álvaro y el administrador. El precio se pone una vez (y se confirma); después lo actualizan solas las compras
 * (cada cotización). Si alguien lo vuelve a cambiar a mano, se les avisa a Álvaro y al administrador. */
function editarProductoExtra(yo, b, d, h, ahora) {
  var out = { mensajes: [] };
  if (d.sucursales && typeof d.sucursales === 'object') {
    var s = d.sucursales, val = [['centro', 5], ['almendras', 6], ['parque', 7]].map(function (x) { return s[x[0]] == null ? (b.suc[x[0]] ? 'Sí' : 'No') : (s[x[0]] === true || s[x[0]] === 'true' ? 'Sí' : 'No'); });
    var antes = ['centro', 'almendras', 'parque'].map(function (k) { return b.suc[k] ? 'Sí' : 'No'; });
    if (val.join() !== antes.join()) { h.getRange(b.fila, 5, 1, 3).setValues([val]); out.mensajes.push('Sucursales actualizadas.'); }
  }
  var provs = null, provOk = function (n) {
    if (!n) return '';
    provs = provs || filasProveedores().map(function (p) { return p.nombre; });
    var hit = provs.filter(function (p) { return p.toLowerCase() === String(n).trim().toLowerCase(); })[0];
    if (!hit) throw new Error('El proveedor «' + n + '» no está registrado. Agréguelo primero en Proveedores.');
    return hit;
  };
  if (d.proveedor != null) { var p1 = provOk(String(d.proveedor).trim()); if (p1 !== b.proveedor) h.getRange(b.fila, 13).setValue(p1); }
  if (d.proveedor2 != null) { var p2 = provOk(String(d.proveedor2).trim());
    if (p2 && p2 === (d.proveedor != null ? provOk(String(d.proveedor).trim()) : b.proveedor)) throw new Error('El proveedor secundario tiene que ser distinto del principal.');
    if (p2 !== b.proveedor2) h.getRange(b.fila, 38).setValue(p2); }
  if (d.presCompra != null) { var pc = String(d.presCompra).trim().slice(0, 40); if (pc !== b.presCompra) h.getRange(b.fila, 26).setValue(pc); }
  if (d.rinde != null && String(d.rinde).trim() !== '') { var rd = Number(String(d.rinde).replace(',', '.'));
    if (!(rd > 0 && rd <= 100000)) throw new Error('Diga cuántas ' + plur(2, b.unidad) + ' trae lo que se compra.'); if (rd !== b.rinde) h.getRange(b.fila, 14).setValue(r3(rd)); }
  if (d.precio != null && String(d.precio).trim() !== '') {
    var pr = r2(String(d.precio).replace(/[Q,\s]/g, ''));
    if (!(pr > 0)) throw new Error('Escriba el precio de lo que se compra (mayor que cero).');
    if (pr !== b.precioPres) {
      if (b.precioPres > 0 && d.confirmaCambio !== true) throw new Error('Este producto ya tiene precio (' + dinero(b.precioPres) + '). Confirme el cambio: se les avisará a Álvaro y al administrador.');
      h.getRange(b.fila, 35, 1, 3).setValues([[pr, ahora, yo.nombre]]); h.getRange(b.fila, 39).setValue(b.precioPres > 0 ? b.precioPres : '');
      if (b.precioPres > 0) out.cambioPrecio = { antes: b.precioPres, ahora: pr };
      out.mensajes.push(b.precioPres > 0 ? 'Precio cambiado de ' + dinero(b.precioPres) + ' a ' + dinero(pr) + ': se les avisó a Álvaro y al administrador.' : 'Precio guardado: ' + dinero(pr) + '. De aquí en adelante lo actualizan las compras.');
    }
  }
  out.mensaje = out.mensajes.join(' ');
  return out;
}
/** Un cambio de precio a mano: correo a Álvaro y al administrador (el aviso de la campanita lo arma avisosPrecioProducto). */
function avisaCambioPrecioProducto(yo, b, c) {
  var para = usuariosCache().filter(function (u) { return u.activo && correoValido(u.correo) && u.nombre !== yo.nombre && (u.rol === 'operaciones' || u.rol === 'admin'); });
  if (!para.length) return [];
  var html = htmlCorreo('Cambio de precio: ' + b.nombre, '<p>' + esHtml(yo.nombre) + ' cambió el precio de <b>' + esHtml(b.nombre) + '</b>' + (b.presCompra ? ' (por ' + esHtml(b.presCompra) + ')' : '') +
    ': de <b>' + dinero(c.antes) + '</b> a <b>' + dinero(c.ahora) + '</b>, el ' + Utilities.formatDate(new Date(), ZONA, 'dd/MM/yyyy HH:mm') + '.</p>');
  var ok = []; para.forEach(function (u) { try { if (enviaCorreo(u.correo, 'Rey Pizza · Cambio de precio: ' + b.nombre, html)) ok.push(u.nombre); } catch (e) {} });
  return ok;
}
function avisosPrecioProducto(yo, out, desde) {
  if (!(yo.rol === 'operaciones' || yo.esAdmin)) return;
  filasBodega().forEach(function (b) {
    if (!(b.precioAnt > 0) || !b.precioEnRaw || !b.precioPor || /^Cotizaci/.test(b.precioPor) || b.precioPor === yo.nombre) return;
    if (fmtDia(b.precioEnRaw) < desde) return;
    out.push({ id: 'prc-' + b.codigo + '-' + Utilities.formatDate(b.precioEnRaw, ZONA, 'yyyyMMddHHmm'), tipo: 'Cambio de precio', cuando: fmtSello(b.precioEnRaw), titulo: b.nombre + ': ' + dinero(b.precioAnt) + ' → ' + dinero(b.precioPres),
      texto: 'Lo cambió ' + b.precioPor + (b.presCompra ? ' · por ' + b.presCompra : '') + '.', abrir: { prod: true } });
  });
}

/* ════════════ COMPRAS AL POR MAYOR (BULK) ════════════
 * Cada producto dice en qué presentación lo compra la bodega (paquete, caja, fardo…) y cuántas unidades trae. Con eso:
 * 3 paquetes × Q150 = Q450 → 1,500 unidades → Q0.30 cada unidad, que es el costo con que sale a las sucursales. */
var PRES_COMPRA = ['paquete', 'caja', 'fardo', 'bolsa', 'saco', 'bulto', 'cubeta', 'galón', 'docena', 'ciento', 'millar', 'rollo', 'display', 'unidad'];
function unidadDeCodigo(c) {
  var m = _MEMO.uniCod; if (!m) { m = _MEMO.uniCod = {}; filasBodega().forEach(function (b) { m[b.codigo] = b; }); }
  return m[c] || null;
}
function compraTxtDe(b) {
  if (!b || !b.presCompra || !(b.rinde > 1)) return '';
  var pres = String(b.presCompra).replace(/\s+de\s+[\d.,]+.*$/i, '');
  return 'Bodega lo compra por ' + pres + ' de ' + num(b.rinde) + ' ' + plur(b.rinde, b.unidad);
}

/* ════════════ PROVEEDORES DENTRO DE PRODUCTOS ════════════
 * Samuel, Álvaro y el administrador ven, agregan y editan los proveedores desde Productos (sin pasar por la caja).
 * Cada proveedor muestra qué productos le compramos (como principal o como secundario). */
function proveedoresProductos(cred) {
  var yo = quien(cred);
  if (!puedeProductos(yo)) throw new Error('Los proveedores los ven la bodega, el director operativo y el administrador.');
  var prin = {}, sec = {};
  filasBodega().forEach(function (b) {
    if (!b.activo) return;
    if (b.proveedor) (prin[b.proveedor.toLowerCase()] = prin[b.proveedor.toLowerCase()] || []).push(b.nombre);
    if (b.proveedor2) (sec[b.proveedor2.toLowerCase()] = sec[b.proveedor2.toLowerCase()] || []).push(b.nombre);
  });
  return { proveedores: filasProveedores().filter(function (p) { return p.tipo !== 'Repartidor' && p.tipo !== 'Sucursal'; }).map(function (p) {
      var k = p.nombre.toLowerCase();
      return { nombre: p.nombre, nit: p.nit, tel: p.tel, activo: p.activo, principal: (prin[k] || []).sort(), secundario: (sec[k] || []).sort() };
    }).sort(function (a, b) { return a.nombre.localeCompare(b.nombre, 'es'); }),
    puedeEditar: puedeProductos(yo) };
}
/** Cuánto se ha usado un proveedor (para saber si se puede borrar sin perder historia). */
function usoProveedor(nombre) {
  var k = String(nombre).trim().toLowerCase(), u = { gastos: 0, detalle: 0, ordenes: 0, pagos: 0, productos: 0, bodega: 0 };
  var cuenta = function (h, ancho, col) {
    if (!h) return 0;
    return leeTodo(h, ancho).filter(function (r) { return String(r[col] || '').trim().toLowerCase() === k; }).length;
  };
  var ss = libro();
  u.gastos = cuenta(ss.getSheetByName('Gastos'), H_GASTOS.length, 14);
  u.detalle = cuenta(ss.getSheetByName('Gastos detalle'), H_DETALLE.length, 4);
  u.ordenes = cuenta(ss.getSheetByName('Órdenes de compra'), H_OC.length, 3);
  u.pagos = cuenta(ss.getSheetByName('Órdenes de compra pagos'), H_OCP.length, 1);
  u.productos = filasProductos().filter(function (p) { return p.proveedor.toLowerCase() === k; }).length;
  u.bodega = filasBodega().filter(function (b) { return (b.proveedor || '').toLowerCase() === k || (b.proveedor2 || '').toLowerCase() === k; }).length;
  u.total = u.gastos + u.detalle + u.ordenes + u.pagos + u.productos + u.bodega;
  return u;
}
/** Eliminar un proveedor. Solo se borra si nunca se usó; si ya tiene compras, órdenes o productos
 *  no se borra (se perdería la historia): se deshabilita y deja de salir en las listas. */
function eliminarProveedor(cred, nombre) {
  var yo = quien(cred);
  if (!(yo.esAdmin || yo.rol === 'operaciones')) throw new Error('Solo el administrador y el director operativo eliminan proveedores.');
  nombre = String(nombre || '').replace(/\s+/g, ' ').trim();
  var lock = LockService.getScriptLock(); lock.waitLock(15000); _LEE = {}; _BOD = null;
  var msg, archivado = false;
  try {
    var h = hojaCat('Proveedores', H_PROVEEDORES), x = null;
    filasProveedores().forEach(function (y) { if (y.nombre.toLowerCase() === nombre.toLowerCase()) x = y; });
    if (!x) throw new Error('No se encontró ese proveedor.');
    var u = usoProveedor(x.nombre);
    if (u.total > 0) {
      var partes = [];
      if (u.gastos) partes.push(u.gastos + (u.gastos === 1 ? ' gasto' : ' gastos'));
      if (u.ordenes) partes.push(u.ordenes + (u.ordenes === 1 ? ' orden de compra' : ' órdenes de compra'));
      if (u.productos + u.bodega) partes.push((u.productos + u.bodega) + ' productos');
      if (!partes.length) partes.push('movimientos');
      h.getRange(x.fila, 4).setValue('No'); archivado = true;
      msg = '«' + x.nombre + '» tiene historial (' + partes.join(', ') + '), así que no se borra para no perderlo. Quedó deshabilitado: ya no sale en las listas.';
    } else {
      h.deleteRow(x.fila);
      msg = '«' + x.nombre + '» eliminado.';
    }
  } finally { lock.releaseLock(); }
  var r = proveedoresProductos(cred); r.ok = true; r.mensaje = msg; r.nombre = nombre; r.archivado = archivado;
  return r;
}

/** Agregar un proveedor nuevo, o cambiar el NIT, el teléfono o si está activo (el nombre de uno que ya existe no se cambia aquí). */
function guardarProveedorProd(cred, p) {
  var yo = quien(cred);
  if (!puedeProductos(yo)) throw new Error('Los proveedores los agregan la bodega, el director operativo y el administrador.');
  p = p || {};
  var original = String(p.original || '').trim(), nombre = original || String(p.nombre || '').replace(/\s+/g, ' ').trim();
  if (nombre.length < 2) throw new Error('Escriba el nombre del proveedor.');
  if (nombre.length > 60) throw new Error('El nombre del proveedor es muy largo.');
  var lock = LockService.getScriptLock(); lock.waitLock(15000); _LEE = {};
  var msg;
  try {
    var h = hojaCat('Proveedores', H_PROVEEDORES), x = null;
    filasProveedores().forEach(function (y) { if (y.nombre.toLowerCase() === nombre.toLowerCase()) x = y; });
    var nit = String(p.nit == null ? (x ? x.nit : '') : p.nit).trim().slice(0, 20), tel = String(p.tel == null ? (x ? x.tel : '') : p.tel).trim().slice(0, 20);
    var activo = p.activo === false ? 'No' : 'Sí';
    if (x) { if (!original && x.activo) throw new Error('Ya existe «' + x.nombre + '». Escójalo de la lista.');
      h.getRange(x.fila, 2, 1, 3).setValues([[nit, tel, activo]]); msg = x.nombre + (x.activo ? ' actualizado.' : ' vuelve a estar en la lista.'); nombre = x.nombre; }
    else { if (original) throw new Error('No se encontró ese proveedor.');
      var fila = [nombre, nit, tel, activo, new Date(), yo.nombre, '', '']; h.appendRow(fila); msg = nombre + ' agregado a los proveedores.'; }
  } finally { lock.releaseLock(); }
  var r = proveedoresProductos(cred); r.ok = true; r.mensaje = msg; r.nombre = nombre;
  return r;
}
