/* Una copia completa de TODAS las pestañas de la hoja de cálculo, en memoria y (cifrada) en la base de datos.
 * La hoja de Google sigue mandando. Esta copia se pone al día sola:
 *   - cada cambio que pasa por este servidor trae la lista de pestañas que tocó («tocadas»), y solo esas se vuelven a traer;
 *   - cada pocos minutos se compara un resumen de todas las pestañas (filas, columnas, huella de las últimas filas) y se corrigen las que difieren;
 *   - si algo falla o la copia está vieja, no se usa (contesta Google). */
import { Libro, Hoja } from './gas_libro.js';
import { cifra, descifra } from './gas_cifrado.js';
import { crearCache } from './gas_servicios.js';
import crypto from 'node:crypto';

const huella = (t) => crypto.createHash('sha256').update(t).digest('base64').slice(0, 22);
const dec = (v) => (v && typeof v === 'object' && '$d' in v ? new Date(v.$d) : v);

/** Lo mismo que calcula Apps Script (exportarResumenLibro) pero sobre la copia local. */
export function resumenLocal(h) {
  const f = h.getLastRow(), c = h.getLastColumn(); let cola = '';
  if (f > 0 && c > 0) { const d = Math.max(1, f - 24); cola = huella(JSON.stringify(h.getRange(d, 1, f - d + 1, c).getValues())); }
  return { nombre: h.getName(), filas: f, cols: c, huella: cola };
}

export function crearCopia({ cfg, pool, origen }) {
  const libro = new Libro(); const cache = crearCache();
  const est = { disparadores: [], propiedades: {}, cargada: false, ultimaVerificacionMs: 0, ultimaOkMs: 0, error: null, sucio: false, pendientes: new Set(), hojas: 0, filas: 0, desdeBaseDatos: false };
  let enCurso = null, temporizador = null, actividadMs = 0, completaMs = 0;
  const listo = () => !!(pool && cfg.appsScript && cfg.serverToken && cfg.pinPepper && cfg.lecturasModo !== 'apagado');

  const llamar = async (fn, extra = {}) => { const r = await origen.llamar({ interno: true, fn, token: cfg.serverToken, ...extra }); if (!r.ok) throw new Error(String(r.error || 'Apps Script no contestó')); return r.v; };

  async function traeHoja(nombre) {
    const filas = []; let desde = 1, v;
    do {
      v = await llamar('exportarHojaTipada', { nombre, desde, max: cfg.paginaHoja });
      if (!v.existe) return null;
      v.celdas.forEach((f) => filas.push(f.map(dec)));
      desde = v.hasta + 1;
    } while (v.hayMas);
    const h = new Hoja(libro, nombre);
    h.filas = filas; h.maxCols = Math.max(26, v.maxCols || 0, v.cols || 0); h.maxFilas = Math.max(1000, v.filas || 0);
    h.textoCols = (v.textoCols || []).map((c) => ({ c1: c, c2: c, r1: 1, r2: 0 })); h.sucia = false;
    return h;
  }
  async function guardaHoja(h) {
    if (!pool) return;
    const datos = cifra({ filas: h.filas.map((f) => f.map((x) => (x instanceof Date ? { $d: x.getTime() } : x))), textoCols: h.textoCols, maxFilas: h.maxFilas, maxCols: h.maxCols }, cfg.pinPepper);
    await pool.query(`INSERT INTO copia_hojas (nombre, datos, actualizado) VALUES ($1,$2,now()) ON CONFLICT (nombre) DO UPDATE SET datos=$2, actualizado=now()`, [h.nombre, datos]);
  }
  async function cargaDeBaseDatos() {
    const r = await pool.query('SELECT nombre, datos FROM copia_hojas');
    for (const x of r.rows) {
      const d = descifra(x.datos, cfg.pinPepper); const h = new Hoja(libro, x.nombre);
      h.filas = d.filas.map((f) => f.map(dec)); h.textoCols = d.textoCols || []; h.maxFilas = d.maxFilas; h.maxCols = d.maxCols; h.sucia = false; libro.hojas.push(h);
    }
    return r.rows.length;
  }
  function cuenta() { est.hojas = libro.hojas.length; est.filas = libro.hojas.reduce((a, h) => a + h.getLastRow(), 0); }

  async function sincronizarUna({ completa } = {}) {
    if (!listo()) return { ok: false, motivo: 'sin configurar' };
    try {
      if (!est.cargada) { try { est.desdeBaseDatos = (await cargaDeBaseDatos()) > 0; } catch (e) { libro.hojas.length = 0; est.desdeBaseDatos = false; } }
      const pend = [...est.pendientes]; est.pendientes.clear();
      const resumen = await llamar('exportarResumenLibro'); const res = resumen.hojas; est.disparadores = Array.isArray(resumen.disparadores) ? resumen.disparadores : est.disparadores; if (resumen.propiedades && typeof resumen.propiedades === 'object') est.propiedades = resumen.propiedades;
      const enGoogle = new Map(res.map((x) => [x.nombre, x]));
      const aTraer = new Set(pend.filter((n) => n !== '*' && enGoogle.has(n)));
      const forzarTodo = completa || !est.cargada && !est.desdeBaseDatos;
      for (const x of res) {
        const h = libro.getSheetByName(x.nombre);
        if (forzarTodo || !h) { aTraer.add(x.nombre); continue; }
        const l = resumenLocal(h);
        if (l.filas !== x.filas || l.cols !== x.cols || l.huella !== x.huella) aTraer.add(x.nombre);
      }
      for (const h of [...libro.hojas]) if (!enGoogle.has(h.nombre)) { libro.quitaHoja(h.nombre); await pool.query('DELETE FROM copia_hojas WHERE nombre=$1', [h.nombre]); }
      for (const n of aTraer) { const h = await traeHoja(n); if (h) { libro.reemplazaHoja(n, h); await guardaHoja(h); } }
      cache.limpia(['fallos_']);                   // lo que el código guardó en su caché ya no vale
      est.cargada = true; est.error = null; est.sucio = est.pendientes.size > 0; est.ultimaOkMs = est.ultimaVerificacionMs = Date.now(); if (completa) completaMs = Date.now(); cuenta();
      return { ok: true, traidas: [...aTraer] };
    } catch (e) { est.error = String(e.message || e).slice(0, 300); est.sucio = true; return { ok: false, motivo: est.error }; }
  }
  function sincronizar(opc) { if (!enCurso) enCurso = sincronizarUna(opc).finally(() => { enCurso = null; }); return enCurso; }

  /** Una ejecución cambió estas pestañas: se vuelven a traer (la siguiente lectura espera a eso). */
  function marcaTocadas(lista) { if (!lista || !lista.length || !listo()) return; lista.forEach((n) => est.pendientes.add(n)); est.sucio = true; sincronizar().catch(() => {}); }
  function actividad() { actividadMs = Date.now(); }
  /** ¿Se puede contestar desde la copia con confianza? Si hace falta, primero se actualiza. */
  async function alDia() {
    if (!listo()) return false;
    if (!est.cargada || est.sucio || Date.now() - est.ultimaVerificacionMs > cfg.frescoCopiaMs) {
      const p = sincronizar(); await Promise.race([p, new Promise((r) => setTimeout(r, est.cargada ? 15000 : cfg.esperaPrimeraCopiaMs))]);
      if (est.sucio && enCurso) await Promise.race([enCurso, new Promise((r) => setTimeout(r, 15000))]);
    }
    return est.cargada && !est.sucio && Date.now() - est.ultimaVerificacionMs <= cfg.maxViejaCopiaMs;
  }
  function iniciar() {
    if (!listo()) return;
    sincronizar({ completa: false }).catch(() => {});
    temporizador = setInterval(() => {
      const cada = (Date.now() - actividadMs < 10 * 60 * 1000 ? cfg.segundosCopia : 600) * 1000;
      const completa = !completaMs || Date.now() - completaMs > cfg.horasCompletaCopia * 3600 * 1000;
      if (completa || Date.now() - est.ultimaVerificacionMs >= cada) sincronizar({ completa: completa && est.cargada }).catch(() => {});
    }, 15000);
    temporizador.unref?.();
  }
  const estado = () => ({ lista: listo(), cargada: est.cargada, hojas: est.hojas, filas: est.filas, error: est.error, sucio: est.sucio, desdeBaseDatos: est.desdeBaseDatos,
    ultimaOk: est.ultimaOkMs ? new Date(est.ultimaOkMs).toISOString() : null });
  return { libro, cache, disparadores: () => est.disparadores, propiedades: () => est.propiedades, listo, sincronizar, marcaTocadas, actividad, alDia, iniciar, estado, parar() { clearInterval(temporizador); } };
}
