/* El código de la app (Codigo.gs) corre aquí igual que en Google: con la hora de Guatemala para las fechas. */
process.env.TZ = process.env.TZ || 'America/Guatemala';
import { leeConfig } from './config.js';
import { crearPool, migrarConReintentos } from './db.js';
import { crearServidor } from './server.js';
import { iniciarCalentador } from './calentar.js';

const cfg = leeConfig();
const pool = crearPool(cfg);
const log = (o) => console.log(JSON.stringify(o));

if (cfg.appsScript && (!cfg.serverToken || !cfg.pinPepper)) log({ nivel: 'aviso', msg: 'Faltan SERVER_TOKEN o PIN_PEPPER: el servidor todavía no aprende de la hoja Usuarios (todo sigue pasando a Apps Script).' });
if (!cfg.appsScript) log({ nivel: 'aviso', msg: 'Falta APPS_SCRIPT_URL: las consultas que no estén migradas fallarán.' });

const { server, bitacora, origen, usuarios, indicadores, copia, lecturas } = crearServidor({ cfg, pool });
server.listen(cfg.puerto, () => log({ nivel: 'info', msg: 'Servidor escuchando', puerto: cfg.puerto }));

// Las migraciones corren en segundo plano: si la base tarda en estar lista, las consultas que pasan a Apps Script igual funcionan.
if (pool) {
  migrarConReintentos(pool)
    .then((h) => { log({ nivel: 'info', msg: 'Base de datos lista', migraciones: h }); usuarios.iniciar(); indicadores.iniciar(); lecturas.carga(); copia.iniciar(); return bitacora.limpiar(); })
    .catch((e) => log({ nivel: 'error', msg: 'La base de datos no quedó lista', detalle: String(e.message || e) }));
  setInterval(() => bitacora.limpiar(), 6 * 3600 * 1000).unref();
} else log({ nivel: 'aviso', msg: 'Sin DATABASE_URL: funciona solo como puerta hacia Apps Script, sin bitácora.' });

const calentador = iniciarCalentador({ cfg, origen, bitacora });

function apagar(senal) {
  log({ nivel: 'info', msg: 'Apagando', senal });
  calentador.parar(); usuarios.parar(); indicadores.parar(); copia.parar();
  server.close(async () => { await bitacora.vaciar(); if (pool) await pool.end().catch(() => {}); process.exit(0); });
  setTimeout(() => process.exit(0), 8000).unref();
}
process.on('SIGTERM', () => apagar('SIGTERM'));
process.on('SIGINT', () => apagar('SIGINT'));
