/* Guardado rápido (Fase 4).
 *
 * Para ciertos registros (por ahora gastos e ingresos) el código ORIGINAL de la app corre aquí, sobre la copia de las hojas:
 *   1. se ejecuta y se anota cada celda que cambia («antes» y «ahora»);
 *   2. se guarda esa lista en la base de datos (cola_escrituras) → desde ese momento el registro está a salvo y la persona recibe su respuesta;
 *   3. por detrás, la lista se reenvía a la hoja de Google (aplicarCambios), que antes de escribir comprueba que ninguna celda haya cambiado.
 * Si algo no es seguro (borrar filas, mandar correos, errores, la copia no está al día…) no se usa el camino rápido: se deshace todo y lo hace Google, como siempre.
 * Mientras haya algo pendiente de reenviar: la copia no se pisa con lo de Google, y lo que va a Google espera (hasta BARRERA_MS) a que se reenvíe.
 *
 * La llave de cada registro (rid) decide por dónde va, para que NUNCA se guarde dos veces:
 *   - si ya se guardó por aquí (solicitudes o cola_escrituras) → se contesta lo mismo;
 *   - si alguna vez se mandó a Google (rid_google) → se repite en Google, que la reconoce;
 *   - si después de dejarlo a salvo algo falla, se contesta que sí se guardó (nunca se manda también a Google). */
import { unaVez } from './unavez.js';
import { cod } from './gas_libro.js';

const ESPERAS_DEF = [1500, 4000, 10000, 30000, 60000];       // reintentos del reenvío
class Alternativa extends Error {}                           // «mejor que lo haga Google»
export class Transitorio extends Error {}                    // «no se sabe si quedó: la app debe repetir con la misma llave»
const espera = (ms) => new Promise((r) => setTimeout(r, ms));

export function crearEscrituras({ cfg, pool, copia, lecturas, origen, emite = () => {}, reloj = () => Date.now() }) {
  const permitidas = new Set(cfg.escrituraFn || []), ESPERAS = cfg.esperasReenvio || ESPERAS_DEF, seguidas = new Map();     // seguidas: cuántas veces seguidas una función no fue apta
  const cola = [];                                             // lo pendiente, en orden: { id, rid, fn, ops, intentos }
  const recientes = new Map();                                 // rid → { v } de lo guardado por aquí en las últimas horas (por si la base de datos no contesta)
  const recuerda = (rid, v) => { if (!rid) return; recientes.set(rid, { v }); if (recientes.size > 2000) recientes.delete(recientes.keys().next().value); };
  const noAptas = new Map();                                   // función → hasta cuándo no se intenta el camino rápido
  const est = { rapidas: 0, alternativas: 0, aplicadas: 0, ajustes: 0, conflictos: 0, reintentos: 0, ultimoError: '', ultimaAplicadaMs: 0, motivos: {} };
  let reserva = false, iniciando = false, drenando = false, temporizador = null, vacia = [], persiste = null, verifica = null, cadena = Promise.resolve(), apagada = false;
  const activaModo = () => cfg.escrituraModo === 'auto' && !!pool && permitidas.size > 0;
  const pendientes = () => cola.length;
  const retiene = () => reserva || cola.length > 0;
  const ocupado = () => reserva || iniciando || cola.length > 0;
  copia.ponRetiene(retiene);

  const motivoAlt = (m) => { est.alternativas++; est.motivos[m] = (est.motivos[m] || 0) + 1; return new Alternativa(m); };
  function avisaVacia() { if (!ocupado()) { const w = vacia; vacia = []; w.forEach((f) => f()); } }
  /** Espera a que no quede nada pendiente (hasta `ms`). Lo usa lo que va a Google, para no leer ni escribir sobre algo viejo. */
  function esperaVacia(ms) {
    if (!ocupado()) return Promise.resolve(true);
    return new Promise((res) => { const t = setTimeout(() => { vacia = vacia.filter((f) => f !== fin); res(false); }, ms); const fin = () => { clearTimeout(t); res(true); }; vacia.push(fin); });
  }
  function programaPersistencia() {
    if (persiste) return;
    persiste = setTimeout(() => { persiste = null; copia.guardaSucias().catch(() => {}); }, 2000); persiste.unref?.();
  }
  function programaVerificacion() {
    clearTimeout(verifica); verifica = setTimeout(() => { if (!retiene() && !apagada) copia.sincronizar().catch(() => {}); }, 300); verifica.unref?.();
  }

  /** ¿Se puede intentar el camino rápido para esta función ahora mismo? */
  async function puede(fn) {
    if (!activaModo() || !permitidas.has(fn) || iniciando) return false;
    const n = noAptas.get(fn); if (n && n > reloj()) return false;
    if (cola.length >= cfg.maxCola) return false;
    return lecturas.activa();
  }

  /** Lo que ya se sabe de esta llave: { v } (ya se guardó aquí), 'google' (se mandó a Google) o null (nueva). */
  async function buscaPrevio(rid) {
    if (!rid) return null;
    {
      const s = (await pool.query(`SELECT estado, resultado FROM solicitudes WHERE rid=$1`, [rid])).rows[0];
      if (s && s.estado === 'ok') return { v: s.resultado };
      if (s && s.estado === 'en_curso') {                    // quedó a medias de un reinicio: se libera (aquí todo va de a uno, así que nadie más la está guardando)
        await pool.query(`DELETE FROM solicitudes WHERE rid=$1 AND estado='en_curso'`, [rid]);
      }
    }
    const q = (await pool.query(`SELECT id, rid, fn, ops, resultado, estado, intentos FROM cola_escrituras WHERE rid=$1`, [rid])).rows[0];
    if (q) { adopta(q); return { v: q.resultado }; }
    const g = (await pool.query(`SELECT 1 FROM rid_google WHERE rid=$1`, [rid])).rows[0];
    if (g) return 'google';
    return null;
  }
  /** Una fila de la cola que no está en memoria (la base de datos la guardó pero aquí no se supo): se pone en la copia y se reenvía. */
  function adopta(q) {
    if (q.estado !== 'pendiente' || cola.some((x) => x.id === q.id)) return;
    reserva = true;
    try { const malas = copia.libro.aplicaOps(q.ops); if (malas.length) copia.marcaTocadas(malas); }
    finally { reserva = false; }
    cola.push({ id: q.id, rid: q.rid, fn: q.fn, ops: q.ops, intentos: q.intentos || 0 }); recuerda(q.rid, q.resultado);
    programaPersistencia(); arranca(0);
  }
  async function marcaGoogle(rid, fn) {
    if (!rid) return;
    await pool.query(`INSERT INTO rid_google (rid, fn) VALUES ($1,$2) ON CONFLICT (rid) DO NOTHING`, [rid, fn]);
  }

  /** Corre fn(args) por el camino rápido. Devuelve { v } si se hizo (o ya estaba hecho), o null si debe hacerlo Google.
   *  Lanza Transitorio cuando no se puede saber si quedó (la app repite con la misma llave).
   *  Todo —revisar la llave, guardar y, si toca, anotarla como de Google— pasa de a uno: dos intentos con la misma llave nunca se cruzan. */
  async function ejecuta(fn, args, rid, sid) {
    if (!activaModo() || !permitidas.has(fn)) return null;
    const turno = cadena; let suelta; cadena = new Promise((r) => { suelta = r; });
    await turno.catch(() => {});
    try {
      const mem = rid && recientes.get(rid);
      if (mem) return { v: mem.v };                            // se guardó por aquí hace poco (aunque la base de datos no conteste)
      let previo;
      try { previo = await buscaPrevio(rid); }
      catch (e) {
        if (e instanceof Transitorio) throw e;
        return null;                                           // la base de datos no contesta y esta llave no se guardó por aquí: que lo haga Google (reconoce la llave)
      }
      if (previo === 'google') return null;
      if (previo) return previo;
      const r = await rapido(fn, args, rid, sid);
      if (r === null) { try { await marcaGoogle(rid, fn); } catch (e) {} }        // desde ya, esta llave es de Google
      return r;
    } finally { suelta(); }
  }

  async function rapido(fn, args, rid, sid) {
    if (!(await puede(fn))) return null;
    let hecho = null;
    try {
      await copia.esperaSync().catch(() => {});             // si la copia se estaba poniendo al día desde Google, se espera a que termine
      if (!copia.fresca()) { await copia.alDia().catch(() => false); if (!copia.fresca()) throw motivoAlt('copia no al día'); }
      reserva = true;                                        // desde aquí la copia no se pisa con lo de Google
      try {
        const v = await unaVez(pool, rid, fn, async () => { const x = await guarda(fn, args, rid, sid); hecho = { v: x }; return x; });
        return { v };
      } finally { reserva = false; avisaVacia(); }
    } catch (e) {
      if (hecho) return hecho;                               // ya quedó a salvo y en la cola: se contesta que sí (nunca se manda también a Google)
      if (e instanceof Alternativa) return null;
      throw e;
    }
  }

  async function guarda(fn, args, rid, sid) {
    const x = lecturas.correrEscritura(fn, args);
    if (!x.ok) {
      if (x.noApta) {                                              // crear una pestaña que faltaba es cosa de una vez; si pasa seguido, esa función ya no se intenta por 30 minutos
        const n = (seguidas.get(fn) || 0) + 1; seguidas.set(fn, n); if (n >= 3) { noAptas.set(fn, reloj() + 30 * 60 * 1000); seguidas.set(fn, 0); }
        throw motivoAlt('no apta: ' + x.motivo);
      }
      throw motivoAlt('error: ' + String(x.motivo).slice(0, 80));      // un error de la app: que lo conteste Google, que es quien manda
    }
    seguidas.set(fn, 0);
    const v = x.r.v === undefined ? null : x.r.v;
    if (!x.ops.length) { est.rapidas++; return v; }
    if (x.ops.length > 2000) { copia.libro.deshaz(x.undo); throw motivoAlt('demasiados cambios'); }
    const ops = x.ops.map((o) => (o.t === 'v' ? { t: 'v', h: o.h, r: o.r, c: o.c, a: cod(o.a), v: cod(o.v) } : o));
    let id;
    try {
      id = (await pool.query('INSERT INTO cola_escrituras (rid, fn, ops, resultado) VALUES ($1,$2,$3,$4) RETURNING id', [rid || null, fn, JSON.stringify(ops), JSON.stringify(v)])).rows[0].id;
    } catch (e) {
      // ¿quedó o no quedó? Si se puede saber, se sigue; si no, se deshace aquí y la app repite con la misma llave (entonces se sabrá)
      let ya = null, supe = false;
      if (rid) { try { ya = (await pool.query('SELECT id FROM cola_escrituras WHERE rid=$1', [rid])).rows[0] || null; supe = true; } catch (e2) { supe = false; } }
      if (ya) id = ya.id;
      else {
        copia.libro.deshaz(x.undo);
        if (supe || !rid) throw motivoAlt('la base de datos no contestó');      // seguro que no quedó: que lo haga Google
        throw new Transitorio('El servidor tuvo un problema. Intente de nuevo.');
      }
    }
    cola.push({ id, rid, fn, ops, intentos: 0 }); recuerda(rid, v);
    est.rapidas++; programaPersistencia();
    try { emite({ tipo: 'cambio', fn, sid: sid || '', rapida: true }); } catch (e) {}
    arranca(0);
    return v;
  }

  function arranca(ms) { if (apagada || temporizador || drenando) return; temporizador = setTimeout(() => { temporizador = null; drena().catch(() => {}); }, ms); temporizador.unref?.(); }

  async function marca(item, estado, detalle) {
    for (let i = 0; i < 4; i++) {
      try { await pool.query(`UPDATE cola_escrituras SET estado=$2, detalle=$3, intentos=$4, aplicada_en=CASE WHEN $2='pendiente' THEN NULL ELSE now() END WHERE id=$1`, [item.id, estado, detalle ? String(detalle).slice(0, 500) : null, item.intentos]); return true; }
      catch (e) { await espera(300 * (i + 1)); }
    }
    return false;            // si la base no contesta, la fila queda «pendiente»: al reiniciar se reenvía y Google la reconoce (no se duplica)
  }
  /** Reenvía a Google, de a una y en orden. */
  async function drena() {
    if (drenando || apagada) return; drenando = true;
    try {
      while (cola.length && !apagada) {
        const item = cola[0];
        const r = await origen.llamar({ interno: true, fn: 'aplicarCambios', token: cfg.serverToken, id: 'q' + item.id, ops: item.ops });
        const v = r && r.ok ? r.v : null;
        if (v && v.ok) {
          cola.shift(); est.aplicadas++; est.ultimaAplicadaMs = Date.now(); est.ultimoError = ''; item.intentos++;
          const notas = [v.repetido ? 'ya estaba' : '', v.desplazadas ? 'filas corridas: ' + JSON.stringify(v.desplazadas) : '', v.yaEstaban ? 'ya estaban en: ' + v.yaEstaban.join(', ') : '',
            v.descartadas ? 'celdas que mandó la hoja: ' + v.descartadas.join(' | ') : ''].filter(Boolean).join(' · ');
          if (v.desplazadas || v.descartadas || v.yaEstaban) {        // la hoja quedó un poco distinta a la copia: esas pestañas se vuelven a traer
            est.ajustes++; console.error(JSON.stringify({ nivel: 'aviso', msg: 'Guardado rápido: la hoja de Google tenía cambios en medio', id: item.id, fn: item.fn, notas }));
            copia.marcaTocadas([...new Set([...Object.keys(v.desplazadas || {}), ...(v.yaEstaban || []), ...(v.descartadas ? item.ops.map((o) => o.h) : [])])]);
          }
          await marca(item, 'aplicada', notas);
          try { emite({ tipo: 'cambio', fn: item.fn, reenviada: true }); } catch (e) {}
          continue;
        }
        if (v && v.conflicto) {          // no se pudo escribir sin pisar algo: Google manda. Se conserva lo que se iba a escribir para poder recuperarlo a mano.
          cola.shift(); est.conflictos++; est.ultimoError = 'conflicto: ' + v.conflicto[0]; item.intentos++; await marca(item, 'conflicto', v.conflicto.join(' | '));
          console.error(JSON.stringify({ nivel: 'error', msg: 'Guardado rápido: conflicto con la hoja de Google', id: item.id, fn: item.fn, detalle: v.conflicto }));
          copia.marcaTocadas([...new Set(item.ops.map((o) => o.h))]);        // la copia vuelve a traer esas pestañas
          continue;
        }
        // Google no contestó o falló: se reintenta con espera creciente (el registro sigue a salvo en la base de datos)
        item.intentos++; est.reintentos++; est.ultimoError = String((r && r.error) || 'sin respuesta').slice(0, 200);
        if (item.intentos === 1 || item.intentos % 5 === 0) await marca(item, 'pendiente', est.ultimoError);
        drenando = false; arranca(ESPERAS[Math.min(item.intentos - 1, ESPERAS.length - 1)]); return;
      }
    } finally { if (drenando) drenando = false; }
    if (!cola.length) { avisaVacia(); programaVerificacion(); copia.guardaSucias().catch(() => {}); }
  }

  /** Al arrancar: lo que quedó pendiente de la vez anterior se vuelve a poner en la copia (si no está) y se reenvía.
   *  Mientras tanto nada va por el camino rápido y lo que va a Google espera (así el orden no se mezcla). */
  async function recupera() {
    if (!activaModo()) return 0;
    try { await pool.query(`DELETE FROM rid_google WHERE en < now() - interval '3 days'`); } catch (e) {}
    const n = (await pool.query(`SELECT count(*)::int AS n FROM cola_escrituras WHERE estado='pendiente'`)).rows[0].n;
    if (!n) return 0;
    iniciando = true;
    try {
      for (let i = 0; i < 240 && !copia.cargada(); i++) await espera(500);      // la copia se carga primero (de la base de datos)
      const r = await pool.query(`SELECT id, rid, fn, ops, intentos FROM cola_escrituras WHERE estado='pendiente' ORDER BY id`);
      if (copia.cargada()) {
        reserva = true; const malas = new Set();
        try { for (const x of r.rows) copia.libro.aplicaOps(x.ops).forEach((h) => malas.add(h)); } finally { reserva = false; }
        if (malas.size) copia.marcaTocadas([...malas]);                         // no cuadró: esas pestañas se vuelven a traer de Google cuando se vacíe la cola
      }
      for (const x of r.rows) if (!cola.some((y) => y.id === x.id)) cola.push({ id: x.id, rid: x.rid, fn: x.fn, ops: x.ops, intentos: x.intentos });
    } finally { iniciando = false; }
    if (cola.length) arranca(500); else avisaVacia();
    return cola.length;
  }
  async function resumen() {
    if (!pool) return { pendientes: 0 };
    const r = await pool.query(`SELECT estado, count(*)::int AS n, max(creada_en) AS ultima FROM cola_escrituras WHERE creada_en > now() - interval '7 days' GROUP BY estado`);
    return { porEstado: r.rows };
  }
  async function conflictos() { const r = await pool.query(`SELECT id, fn, creada_en, detalle, ops FROM cola_escrituras WHERE estado='conflicto' ORDER BY id DESC LIMIT 50`); return r.rows; }
  const estado = () => ({ modo: cfg.escrituraModo, funciones: [...permitidas], pendientes: cola.length, ...est, noAptas: [...noAptas].filter(([, h]) => h > reloj()).map(([f]) => f) });
  return { ejecuta, puede, esperaVacia, pendientes, retiene, ocupado, recupera, resumen, conflictos, estado, drena, parar() { apagada = true; clearTimeout(temporizador); clearTimeout(persiste); clearTimeout(verifica); } };
}
