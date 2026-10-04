/* Toda la configuración sale de variables de entorno: nada de claves dentro del código. */
const num = (v, d) => { const n = Number(v); return Number.isFinite(n) && n > 0 ? n : d; };
const numOCero = (v, d) => { if (v === undefined || v === '') return d; const n = Number(v); return Number.isFinite(n) && n >= 0 ? n : d; };   // aquí 0 sí vale: «nunca»
const lista = (v) => String(v || '').split(',').map((s) => s.trim()).filter(Boolean);

export function leeConfig(env = process.env) {
  return {
    puerto: num(env.PORT, 3000),
    baseDatos: env.DATABASE_URL || '',
    baseDatosSsl: env.DB_SSL === '1',
    appsScript: env.APPS_SCRIPT_URL || '',          // la dirección /exec del Apps Script actual
    origenes: lista(env.ORIGENES || 'https://app.caixsa.com'),
    claveAdmin: env.ADMIN_CLAVE || '',               // para ver /velocidad
    confiaProxy: env.TRUST_PROXY === '1',            // detrás de Render/Netlify: la IP real viene en x-forwarded-for
    minutosCalentar: numOCero(env.WARM_MINUTES, 4),       // cada cuánto «despierta» a Apps Script (0 = nunca)
    msEsperaOrigen: num(env.ORIGEN_TIMEOUT_MS, 55000),
    maxCuerpo: num(env.MAX_CUERPO_BYTES, 8 * 1024 * 1024),
    limitePorMinuto: num(env.LIMITE_POR_MINUTO, 400),
    limiteEntrarPorMinuto: num(env.LIMITE_ENTRAR_POR_MINUTO, 40),
    diasBitacora: num(env.DIAS_BITACORA, 30),
    serverToken: env.SERVER_TOKEN || '',             // clave secreta compartida con Apps Script (puerta interna)
    pinPepper: env.PIN_PEPPER || '',                 // clave secreta para transformar los PIN (solo vive aquí, nunca en la base)
    minutosSync: numOCero(env.SYNC_MINUTES, 2),           // cada cuánto aprende de la hoja Usuarios
    maxIntentos: num(env.MAX_INTENTOS, 8),
  };
}
