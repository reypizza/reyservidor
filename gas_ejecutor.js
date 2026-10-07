/* Corre el código ORIGINAL de la app (el mismo Codigo.gs de Apps Script), sin cambios, dentro del servidor.
 * Cada consulta corre en un entorno nuevo (igual que en Google): se vuelven a definir las funciones y variables, y se llama a doPost. */
import vm from 'node:vm';
import { Libro } from './gas_libro.js';
import { crearUtilities, crearCache, crearPropiedades } from './gas_servicios.js';

const encadenado = () => new Proxy(function () {}, { get: (_, k) => (k === 'then' ? undefined : encadenado()), apply: () => encadenado() });

const propsLectura = (p, efecto) => ({ ...p, setProperty(k, v) { efecto(); return p.setProperty(k, v); }, setProperties(o) { efecto(); return p.setProperties(o); }, deleteProperty(k) { efecto(); return p.deleteProperty(k); } });
export function compilaCodigo(texto) { return new vm.Script(texto, { filename: 'Codigo.gs' }); }

export function crearEntorno({ libro, cache = crearCache(), propiedades = crearPropiedades(), correos = [], registro = [], disparadores = [] } = {}) {
  return { libro: libro || new Libro(), cache, propiedades, correos, registro, disparadores };
}

/** Ejecuta {fn, args, rid} y devuelve lo mismo que devolvía doPost: { ok, v, ms } o { ok:false, error } */
export function ejecuta(script, entorno, q) {
  const { libro, cache, propiedades, correos, registro } = entorno; const disparadores = entorno.disparadores || [];
  const lock = { waitLock() {}, tryLock: () => true, releaseLock() {}, hasLock: () => true };
  const efecto = () => { if (libro.soloLectura) libro._intento(true); };      // mandar un correo, programar algo o guardar una propiedad también cuentan como «cambiar algo»
  const sandbox = {
    console, Date, Proxy, Atomics, SharedArrayBuffer,                       // Date del servidor: así «instanceof Date» vale para las fechas de las hojas
    SpreadsheetApp: { openById: () => libro, getActiveSpreadsheet: () => libro, getActive: () => libro, openByUrl: () => libro,
      newDataValidation: () => encadenado(), newConditionalFormatRule: () => encadenado(), newFilterCriteria: () => encadenado(),         // solo adornos de la hoja (listas desplegables, colores): no cambian datos
      WrapStrategy: new Proxy({}, { get: (_, k) => String(k) }), BorderStyle: new Proxy({}, { get: (_, k) => String(k) }), DataValidationCriteria: new Proxy({}, { get: (_, k) => String(k) }) },
    Utilities: crearUtilities(),
    CacheService: { getScriptCache: () => cache, getUserCache: () => cache, getDocumentCache: () => cache },
    PropertiesService: { getScriptProperties: () => propsLectura(propiedades, efecto), getUserProperties: () => propsLectura(propiedades, efecto), getDocumentProperties: () => propsLectura(propiedades, efecto) },
    LockService: { getScriptLock: () => lock, getUserLock: () => lock, getDocumentLock: () => lock },
    MailApp: { sendEmail(a, b, c, d) { efecto(); correos.push(typeof a === 'object' ? a : { to: a, subject: b, body: c, ...(d || {}) }); }, getRemainingDailyQuota: () => 1500 },
    GmailApp: { sendEmail(to, subject, body, o) { efecto(); correos.push({ to, subject, body, ...(o || {}) }); }, getRemainingDailyQuota: () => 1500 },
    ScriptApp: { getProjectTriggers: () => disparadores.map((f) => ({ getHandlerFunction: () => f, getUniqueId: () => f, getEventType: () => 'CLOCK' })),
      newTrigger: (f) => { efecto(); const b = new Proxy({}, { get: (_, k) => (k === 'create' ? () => { if (!libro.soloLectura) disparadores.push(f); return {}; } : () => b) }); return b; },
      deleteTrigger(t) { efecto(); const i = disparadores.indexOf(t && t.getHandlerFunction ? t.getHandlerFunction() : t); if (i >= 0 && !libro.soloLectura) disparadores.splice(i, 1); }, getService: () => ({ getUrl: () => '' }) },
    HtmlService: { createHtmlOutputFromFile: () => encadenado(), createHtmlOutput: () => encadenado(), createTemplateFromFile: () => encadenado() },
    ContentService: { MimeType: { JSON: 'json', TEXT: 'text' }, createTextOutput: (t) => ({ _t: t, setMimeType() { return this; } }) },
    Session: { getScriptTimeZone: () => 'America/Guatemala', getActiveUser: () => ({ getEmail: () => '' }), getEffectiveUser: () => ({ getEmail: () => '' }) },
    DriveApp: new Proxy({}, { get: (_, k) => { if (k === 'Access' || k === 'Permission') return new Proxy({}, { get: (_, x) => String(x) }); return () => { throw new Error('Google Drive todavía no está disponible en el servidor propio.'); }; } }),
    UrlFetchApp: { fetch() { throw new Error('UrlFetchApp no está disponible en el servidor propio.'); } },
    Logger: { log: (...a) => { registro.push(a.join(' ')); }, getLog: () => registro.join('\n') },
  };
  const ctx = vm.createContext(sandbox);
  script.runInContext(ctx);
  if (q && q.__directa) { try { return { ok: true, v: vm.runInContext(q.__directa, ctx).apply(null, q.args || []) }; } catch (e) { return { ok: false, error: String(e && e.message ? e.message : e) }; } }     // solo para pruebas: una función suelta
  let salida;
  try { salida = vm.runInContext('doPost', ctx)({ postData: { contents: JSON.stringify(q) } }); }
  catch (e) { return { ok: false, error: String(e && e.message ? e.message : e), _interno: true }; }
  return JSON.parse(salida._t);
}
