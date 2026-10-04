import pg from 'pg';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const aqui = path.dirname(fileURLToPath(import.meta.url));
/* Las migraciones (.sql) pueden estar junto al código (carpeta plana, como se sube a GitHub desde una tablet)
 * o en db/migraciones (carpeta ordenada, como se desarrolla). Se usa la primera que tenga archivos .sql. */
function buscaCarpeta() {
  for (const c of [aqui, path.join(aqui, '..', 'db', 'migraciones')]) {
    try { if (fs.readdirSync(c).some((f) => f.endsWith('.sql'))) return c; } catch (e) { /* no existe */ }
  }
  return aqui;
}
export const CARPETA_MIGRACIONES = buscaCarpeta();

export function crearPool(cfg) {
  if (!cfg.baseDatos) return null;
  const pool = new pg.Pool({
    connectionString: cfg.baseDatos,
    ssl: cfg.baseDatosSsl ? { rejectUnauthorized: false } : undefined,
    max: 10,
    connectionTimeoutMillis: 5000,
    idleTimeoutMillis: 30000,
  });
  pool.on('error', (e) => console.error(JSON.stringify({ nivel: 'error', donde: 'pool', msg: String(e.message || e) })));
  return pool;
}

/** Aplica, en orden, las migraciones que todavía no se han aplicado. Segura si dos copias arrancan a la vez. */
export async function migrar(pool, carpeta = CARPETA_MIGRACIONES) {
  const cli = await pool.connect();
  const hechas = [];
  try {
    await cli.query('SELECT pg_advisory_lock(7203001)');
    await cli.query('CREATE TABLE IF NOT EXISTS migraciones (nombre text PRIMARY KEY, aplicada_en timestamptz NOT NULL DEFAULT now())');
    const ya = new Set((await cli.query('SELECT nombre FROM migraciones')).rows.map((r) => r.nombre));
    const archivos = fs.readdirSync(carpeta).filter((f) => f.endsWith('.sql')).sort();
    for (const f of archivos) {
      if (ya.has(f)) continue;
      const sql = fs.readFileSync(path.join(carpeta, f), 'utf8');
      try {
        await cli.query('BEGIN');
        await cli.query(sql);
        await cli.query('INSERT INTO migraciones (nombre) VALUES ($1)', [f]);
        await cli.query('COMMIT');
        hechas.push(f);
      } catch (e) {
        await cli.query('ROLLBACK').catch(() => {});
        throw new Error(`La migración ${f} falló: ${e.message}`);
      }
    }
  } finally {
    await cli.query('SELECT pg_advisory_unlock(7203001)').catch(() => {});
    cli.release();
  }
  return hechas;
}

/** Espera a que la base esté lista (al arrancar, Render a veces levanta el servidor antes que la base). */
export async function migrarConReintentos(pool, intentos = 8, esperaMs = 2500) {
  let ultimo;
  for (let i = 1; i <= intentos; i++) {
    try { return await migrar(pool); }
    catch (e) { ultimo = e; await new Promise((r) => setTimeout(r, esperaMs)); }
  }
  throw ultimo;
}
