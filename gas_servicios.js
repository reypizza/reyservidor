/* Las piezas de Google que el código de la app usa, hechas con lo que hay en el servidor. */
import crypto from 'node:crypto';
import zlib from 'node:zlib';

const ZONA_DEF = 'America/Guatemala';
const MESES = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre'];
const DIAS = ['domingo', 'lunes', 'martes', 'miércoles', 'jueves', 'viernes', 'sábado'];

const DIAS_EN = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
const FORMATEADORES = new Map();
function formateador(tz) {                       // crear un Intl.DateTimeFormat es caro: se crea uno por zona y se reutiliza
  let f = FORMATEADORES.get(tz);
  if (!f) { f = new Intl.DateTimeFormat('en-GB', { timeZone: tz, year: 'numeric', month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit', second: '2-digit', hourCycle: 'h23', weekday: 'long' }); FORMATEADORES.set(tz, f); }
  return f;
}
const dos = (n) => (n < 10 ? '0' + n : '' + n);
const DESDE_2007 = Date.UTC(2007, 0, 1);
function partes(d, tz) {
  const zona = tz || ZONA_DEF, ms = d.getTime();
  if (zona === 'America/Guatemala' && ms >= DESDE_2007) {       // Guatemala está en UTC−6 todo el año desde 2007: sin Intl, solo aritmética
    const t = new Date(ms - 6 * 3600000);
    return { year: '' + t.getUTCFullYear(), month: dos(t.getUTCMonth() + 1), day: dos(t.getUTCDate()), hour: dos(t.getUTCHours()), minute: dos(t.getUTCMinutes()), second: dos(t.getUTCSeconds()), weekday: DIAS_EN[t.getUTCDay()] };
  }
  const o = {}; for (const p of formateador(zona).formatToParts(d)) o[p.type] = p.value; return o;
}
const TOKENS = /'([^']*)'|yyyy|yy|MMMM|MMM|MM|M|dd|d|HH|H|hh|h|mm|ss|SSS|a|EEEE|EEE/g;
const CACHE_FMT = new Map();
/** Los patrones de Java que usa la app: yyyy yy MM M dd d HH H hh h mm ss SSS a EEE EEEE MMM MMMM  (y texto entre comillas simples) */
export function formatDate(d, tz, patron) {
  if (!(d instanceof Date) || isNaN(d)) throw new Error('Argumentos no válidos: formatDate');
  const llave = d.getTime() + '|' + (tz || '') + '|' + patron;
  const hit = CACHE_FMT.get(llave); if (hit !== undefined) return hit;
  const p = partes(d, tz); const ms = String(((d.getTime() % 1000) + 1000) % 1000).padStart(3, '0'); const h12 = (Number(p.hour) % 12) || 12;
  const dowIdx = DIAS_EN.indexOf(p.weekday);
  const r = String(patron).replace(TOKENS, (t, lit) => {
    if (lit !== undefined) return lit === '' ? "'" : lit;
    switch (t) {
      case 'yyyy': return p.year; case 'yy': return p.year.slice(-2); case 'MM': return p.month; case 'M': return String(Number(p.month));
      case 'MMMM': return MESES[Number(p.month) - 1]; case 'MMM': return MESES[Number(p.month) - 1].slice(0, 3);
      case 'dd': return p.day; case 'd': return String(Number(p.day)); case 'HH': return p.hour; case 'H': return String(Number(p.hour));
      case 'hh': return String(h12).padStart(2, '0'); case 'h': return String(h12); case 'mm': return p.minute; case 'ss': return p.second;
      case 'SSS': return ms; case 'a': return Number(p.hour) < 12 ? 'AM' : 'PM'; case 'EEEE': return DIAS[dowIdx]; case 'EEE': return DIAS[dowIdx].slice(0, 3);
    } return t;
  });
  if (CACHE_FMT.size > 60000) CACHE_FMT.clear();
  CACHE_FMT.set(llave, r);
  return r;
}

class Blob {
  constructor(bytes, tipo, nombre) { this._b = Buffer.from(bytes); this._t = tipo || 'application/octet-stream'; this._n = nombre || ''; }
  getBytes() { return Array.from(this._b).map((x) => (x > 127 ? x - 256 : x)); }
  getDataAsString() { return this._b.toString('utf8'); } getContentType() { return this._t; } getName() { return this._n; } setName(n) { this._n = n; return this; }
  setContentType(t) { this._t = t; return this; }
  getAs(tipo) { throw new Error('Convertir a ' + tipo + ' todavía no está disponible en el servidor propio.'); }
  copyBlob() { return new Blob(this._b, this._t, this._n); }
}
const aBuf = (x) => (Buffer.isBuffer(x) ? x : Array.isArray(x) ? Buffer.from(x.map((n) => (n < 0 ? n + 256 : n))) : Buffer.from(String(x)));

export function crearUtilities() {
  return {
    formatDate,
    sleep(ms) { Atomics.wait(new Int32Array(new SharedArrayBuffer(4)), 0, 0, Math.min(Number(ms) || 0, 2000)); },
    computeDigest(alg, v) { const h = crypto.createHash(String(alg === 'SHA_256' || alg === 'SHA_256' ? 'sha256' : alg === 'MD5' ? 'md5' : alg === 'SHA_1' ? 'sha1' : 'sha256')); h.update(typeof v === 'string' ? Buffer.from(v, 'utf8') : aBuf(v)); return Array.from(h.digest()).map((x) => (x > 127 ? x - 256 : x)); },
    DigestAlgorithm: { SHA_256: 'SHA_256', SHA_1: 'SHA_1', MD5: 'MD5' },
    base64Encode(v) { return (typeof v === 'string' ? Buffer.from(v, 'utf8') : aBuf(v)).toString('base64'); },
    base64Decode(s) { return Array.from(Buffer.from(String(s), 'base64')).map((x) => (x > 127 ? x - 256 : x)); },
    base64EncodeWebSafe(v) { return (typeof v === 'string' ? Buffer.from(v, 'utf8') : aBuf(v)).toString('base64url'); },
    newBlob(datos, tipo, nombre) { return new Blob(typeof datos === 'string' ? Buffer.from(datos, 'utf8') : aBuf(datos), tipo, nombre); },
    gzip(blob, nombre) { return new Blob(zlib.gzipSync(blob._b), 'application/x-gzip', nombre || blob._n + '.gz'); },
    ungzip(blob) { return new Blob(zlib.gunzipSync(blob._b), 'application/octet-stream', blob._n); },
    getUuid() { return crypto.randomUUID(); },
    parseDate() { throw new Error('Utilities.parseDate no se usa en esta app.'); },
  };
}

/** El «caché del script»: guarda textos unos segundos o minutos. Vive mientras viva el servidor. */
export function crearCache() {
  const m = new Map();
  const vivo = (k) => { const x = m.get(k); if (!x) return null; if (x.hasta && x.hasta < Date.now()) { m.delete(k); return null; } return x; };
  const c = {
    get(k) { const x = vivo(String(k)); return x ? x.v : null; },
    put(k, v, seg) { m.set(String(k), { v: String(v), hasta: Date.now() + (Number(seg) || 600) * 1000 }); },
    remove(k) { m.delete(String(k)); },
    getAll(ks) { const o = {}; (ks || []).forEach((k) => { const x = vivo(String(k)); if (x) o[k] = x.v; }); return o; },
    /** Borra todo menos lo que empiece con alguno de estos prefijos (los contadores de intentos fallidos se conservan). */
    limpia(conservar = []) { for (const k of [...m.keys()]) if (!conservar.some((p) => k.startsWith(p))) m.delete(k); },
    putAll(o, seg) { Object.keys(o || {}).forEach((k) => c.put(k, o[k], seg)); }, removeAll(ks) { (ks || []).forEach((k) => m.delete(String(k))); },
  };
  return c;
}
export function crearPropiedades(inicial = {}) {
  const m = new Map(Object.entries(inicial));
  return { getProperty: (k) => (m.has(k) ? m.get(k) : null), setProperty(k, v) { m.set(k, String(v)); return this; }, deleteProperty(k) { m.delete(k); return this; },
    getProperties: () => Object.fromEntries(m), setProperties(o) { Object.entries(o).forEach(([k, v]) => m.set(k, String(v))); return this; }, getKeys: () => [...m.keys()] };
}
