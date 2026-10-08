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
    segundosKpi: numOCero(env.SYNC_KPI_SEGUNDOS, 60),    // cada cuánto revisa la hoja de indicadores (con actividad)
    paginaKpi: num(env.PAGINA_KPI, 3000),                 // filas por página al traer registros
    frescoKpiMs: num(env.FRESCO_KPI_MS, 45000),           // hasta qué antigüedad se contesta sin volver a revisar la hoja
    maxViejaKpiMs: num(env.MAX_VIEJA_KPI_MS, 180000),     // más viejo que esto: contesta Apps Script
    esperaPrimeraKpiMs: num(env.ESPERA_PRIMERA_KPI_MS, 6000),
    horasCompletaKpi: num(env.HORAS_COMPLETA_KPI, 6),
    // Lecturas completas con el código de la app corriendo aquí (Fase 3)
    lecturasModo: ['sombra', 'nativa', 'auto'].includes(env.LECTURAS_MODO) ? env.LECTURAS_MODO : 'apagado',   // apagado = todo como antes; auto = cada pantalla pasa sola cuando demuestra que da lo mismo que Google
    minGraduar: Math.max(3, num(env.MIN_GRADUAR, 12)),    // modo auto: cuántas comparaciones seguidas iguales necesita una pantalla para contestarse aquí
    codigoGs: env.CODIGO_GS || '',                       // dónde está el Codigo.gs (por defecto, junto al programa)
    paginaHoja: num(env.PAGINA_HOJA, 2000),
    frescoCopiaMs: num(env.FRESCO_COPIA_MS, 120000),     // cada cuánto se verifica la copia (2 min con la app en uso)
    maxViejaCopiaMs: num(env.MAX_VIEJA_COPIA_MS, 360000),
    esperaPrimeraCopiaMs: num(env.ESPERA_PRIMERA_COPIA_MS, 8000),
    segundosCopia: numOCero(env.SYNC_COPIA_SEGUNDOS, 90),
    horasCompletaCopia: num(env.HORAS_COMPLETA_COPIA, 3),
    muestraVerificacion: Math.min(1, Math.max(0, Number(env.MUESTRA_VERIFICACION ?? 0.05))),   // en modo nativa: qué parte de las respuestas se comparan después con Google
  };
}
