/* Corre el código ORIGINAL de la app (Codigo.gs, sin cambios) sobre la copia de las hojas para contestar las LECTURAS sin esperar a Google.
 *
 *  apagado → todo pasa a Google, como antes (es el modo con el que arranca).
 *  sombra  → Google sigue contestando; en segundo plano se corre lo mismo aquí y se compara. No cambia nada para nadie:
 *            solo junta evidencia de que el servidor da las mismas respuestas.
 *  nativa  → las lecturas las contesta el servidor. Lo que intente cambiar algo (guardar, borrar…) se detecta solo y se manda a Google.
 *            Una parte de las respuestas se sigue comparando con Google; si una función da una respuesta distinta, se deja de usar aquí un rato. */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { compilaCodigo, crearEntorno, ejecuta } from './gas_ejecutor.js';
import { crearPropiedades } from './gas_servicios.js';

const aqui = path.dirname(fileURLToPath(import.meta.url));
const IGNORA = new Set(['ms', 'tocadas', 'ahora', 'hora', 'generado', 'generadoEn']);   // lo que cambia solo con el paso del tiempo

/** La primera diferencia entre dos respuestas (o null si son iguales). */
export function diferencia(a, b, ruta = '$') {
  if (a === b) return null;
  if (typeof a === 'number' && typeof b === 'number') return Math.abs(a - b) < 1e-9 ? null : `${ruta}: ${a} ≠ ${b}`;
  if (a === null || b === null || typeof a !== 'object' || typeof b !== 'object') return `${ruta}: ${corto(a)} ≠ ${corto(b)}`;
  if (Array.isArray(a) !== Array.isArray(b)) return `${ruta}: lista ≠ objeto`;
  if (Array.isArray(a)) {
    if (a.length !== b.length) return `${ruta}: ${a.length} elementos ≠ ${b.length}`;
    for (let i = 0; i < a.length; i++) { const d = diferencia(a[i], b[i], `${ruta}[${i}]`); if (d) return d; }
    return null;
  }
  const claves = new Set([...Object.keys(a), ...Object.keys(b)]);
  for (const k of claves) { if (IGNORA.has(k)) continue; const d = diferencia(a[k], b[k], `${ruta}.${k}`); if (d) return d; }
  return null;
}
const corto = (v) => { const t = JSON.stringify(v); return t === undefined ? 'undefined' : t.length > 60 ? t.slice(0, 57) + '…' : t; };

export function crearLecturas({ cfg, pool, copia, origen }) {
  let script = null, versionLocal = '', versionOk = null, versionRevisadaMs = 0, motivo = null;
  const escriben = new Set(), bloqueadas = new Map(), ventanas = new Map();      // ventanas: los últimos resultados de comparar cada función
  /* Modo auto: una pantalla se «gradúa» (se contesta aquí) cuando sus últimas comparaciones con Google salieron todas iguales.
   * Una sola diferencia la regresa a Google, que la sigue comparando hasta que vuelva a graduarse. */
  const graduadas = new Set(); let graduadasMs = 0;
  async function refrescaGraduadas(forzar) {
    if (!pool || (!forzar && Date.now() - graduadasMs < 120000)) return;
    graduadasMs = Date.now();
    try {
      const r = await pool.query(`SELECT fn, count(*)::int AS n, bool_and(igual) AS todas FROM (
          SELECT fn, igual, row_number() OVER (PARTITION BY fn ORDER BY en DESC, id DESC) AS rn FROM comparaciones WHERE en > now() - interval '7 days') x
        WHERE rn <= $1 GROUP BY fn`, [cfg.minGraduar]);
      graduadas.clear();
      r.rows.forEach((x) => { if (x.n >= cfg.minGraduar && x.todas) graduadas.add(x.fn); });
    } catch (e) { /* si no se puede consultar, se queda como estaba */ }
  }
  async function graduada(fn) {
    if (cfg.lecturasModo !== 'auto') return false;
    await refrescaGraduadas(false);
    const b = bloqueadas.get(fn);
    return graduadas.has(fn) && !(b && b > Date.now()) && !escriben.has(fn);
  }

  function carga() {
    if (script || cfg.lecturasModo === 'apagado') return;
    try {
      const ruta = cfg.codigoGs || [path.join(aqui, 'Codigo.gs'), path.join(aqui, '..', 'Codigo.gs')].find((r) => fs.existsSync(r));
      if (!ruta) { motivo = 'No se encontró Codigo.gs junto al servidor.'; return; }
      const texto = fs.readFileSync(ruta, 'utf8'); script = compilaCodigo(texto);
      versionLocal = (/var VERSION_CODIGO = '([^']+)'/.exec(texto) || [])[1] || '';
    } catch (e) { motivo = 'No se pudo cargar Codigo.gs: ' + e.message; script = null; }
  }
  /** El Codigo.gs de aquí y el de Apps Script tienen que ser el mismo; si no, no se contesta nada desde aquí. */
  async function versionCoincide() {
    if (versionOk !== null && Date.now() - versionRevisadaMs < 300000) return versionOk;
    versionRevisadaMs = Date.now();
    try {
      const r = await origen.llamar({ interno: true, fn: 'versionCodigo', token: cfg.serverToken });
      versionOk = !!(r.ok && versionLocal && r.v === versionLocal);
      motivo = versionOk ? null : `El Codigo.gs del servidor (${versionLocal || '?'}) no es el de Apps Script (${r.ok ? r.v : 'sin respuesta'}).`;
    } catch (e) { versionOk = false; motivo = 'No se pudo comprobar la versión del código.'; }
    return versionOk;
  }
  async function activa() { carga(); return cfg.lecturasModo !== 'apagado' && !!script && copia.listo() && await versionCoincide(); }

  /** Corre la consulta aquí, sin permitir que cambie nada. {ok:true, r} o {ok:false, motivo}. */
  async function nativa(fn, args) {
    if (!(await activa())) return { ok: false, motivo: motivo || 'apagada' };
    if (escriben.has(fn)) return { ok: false, motivo: 'escribe' };
    const b = bloqueadas.get(fn); if (b && b > Date.now()) return { ok: false, motivo: 'bloqueada por diferencias' };
    if (!(await copia.alDia())) return { ok: false, motivo: 'copia no al día' };
    const libro = copia.libro; const t0 = Date.now();
    libro.soloLectura = true; libro.huboEscritura = false; libro.escrituraDatos = false;
    let r;
    try { r = ejecuta(script, { libro, cache: copia.cache, propiedades: crearPropiedades(copia.propiedades()), correos: [], registro: [], disparadores: copia.disparadores() }, { fn, args, rid: '' }); }
    finally { libro.soloLectura = false; }
    if (libro.huboEscritura) {
      const datos = libro.escrituraDatos; libro.huboEscritura = false; libro.escrituraDatos = false;
      if (datos) { escriben.add(fn); return { ok: false, motivo: 'escribe' }; }        // guarda datos: siempre por Google
      return { ok: false, motivo: 'repara' };      // solo iba a crear una pestaña o unos títulos que faltan: que lo haga Google esta vez; la copia lo trae después
    }
    if (r && r._interno) return { ok: false, motivo: 'error del motor: ' + r.error };
    delete r.tocadas; r.ms = Date.now() - t0;
    return { ok: true, r };
  }

  async function anota(fn, igual, msGoogle, msPropio, detalle) {
    // Una diferencia suelta puede ser solo el caché de Google (que a veces guarda unos minutos una respuesta). Si en las últimas 20 comparaciones
    // hubo 3 o más distintas, esa función se deja de contestar aquí durante 30 minutos.
    if (!igual) graduadas.delete(fn);               // modo auto: una diferencia y esa pantalla vuelve a Google
    const v = ventanas.get(fn) || []; v.push(igual); if (v.length > 20) v.shift(); ventanas.set(fn, v);
    if (v.filter((x) => !x).length >= 3) { bloqueadas.set(fn, Date.now() + 30 * 60 * 1000); ventanas.set(fn, []); }
    if (!pool) return;
    try { await pool.query('INSERT INTO comparaciones (fn, igual, ms_google, ms_propio, detalle) VALUES ($1,$2,$3,$4,$5)', [fn, igual, msGoogle ?? null, msPropio ?? null, detalle ? String(detalle).slice(0, 300) : null]); } catch (e) {}
  }
  /** Modo sombra: Google ya contestó; se corre lo mismo aquí y se compara. */
  async function sombra(fn, args, rGoogle, msGoogle) {
    try {
      if (!(await activa()) || escriben.has(fn) || !copia.estado().cargada || copia.estado().sucio) return;
      const n = await nativa(fn, args);
      if (!n.ok) { if (n.motivo !== 'escribe' && n.motivo !== 'repara') await anota(fn, false, msGoogle, null, 'no se pudo correr aquí: ' + n.motivo); return; }
      const d = diferencia(rGoogle, n.r);
      await anota(fn, !d, msGoogle, n.r.ms, d);
    } catch (e) { /* la comparación nunca debe afectar a nadie */ }
  }
  /** Modo nativa: a veces se le pregunta también a Google (después de contestar) para seguir comprobando. */
  async function verificaDespues(fn, args, rPropia) {
    const muestra = cfg.lecturasModo === 'auto' ? Math.max(cfg.muestraVerificacion, 0.1) : cfg.muestraVerificacion;      // en auto se sigue comparando 1 de cada 10
    if (muestra <= 0 || Math.random() >= muestra) return;
    try {
      const t0 = Date.now(); const g = await origen.llamar({ fn, args, rid: '' }); const ms = Date.now() - t0;
      if (g._caido) return; delete g._ms; delete g._caido;
      const tocadas = g.tocadas; delete g.tocadas; if (tocadas && tocadas.length) copia.marcaTocadas(tocadas);
      const d = diferencia(g, rPropia); await anota(fn, !d, ms, rPropia.ms, d);
    } catch (e) {}
  }
  async function resumen() {
    if (!pool) return [];
    const r = await pool.query(`SELECT fn, count(*)::int AS veces, count(*) FILTER (WHERE igual)::int AS iguales,
        round(avg(ms_google))::int AS google_ms, round(avg(ms_propio))::int AS propio_ms,
        max(en) FILTER (WHERE NOT igual) AS ultima_diferencia,
        (array_agg(detalle ORDER BY en DESC) FILTER (WHERE NOT igual))[1] AS ejemplo
      FROM comparaciones WHERE en > now() - interval '7 days' GROUP BY fn ORDER BY (count(*) FILTER (WHERE NOT igual)) DESC, count(*) DESC`);
    return r.rows;
  }
  const estado = () => ({ modo: cfg.lecturasModo, codigoCargado: !!script, versionLocal, versionCoincide: versionOk, motivo, escriben: [...escriben].sort(), bloqueadas: [...bloqueadas].filter(([, h]) => h > Date.now()).map(([f]) => f),
    graduadas: [...graduadas].sort(), minGraduar: cfg.minGraduar });
  return { nativa, sombra, verificaDespues, resumen, estado, activa, carga, anota, graduada, refrescaGraduadas };
}
