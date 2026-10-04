import http from 'node:http';
import crypto from 'node:crypto';
import { crearPuerta } from './puerta.js';
import { crearOrigen } from './origen.js';
import { crearBitacora } from './bitacora.js';
import { crearLimite } from './limite.js';

const VERSION = '0.1.0';

export function crearServidor({ cfg, pool, fetchFn }) {
  const origen = crearOrigen(cfg, fetchFn);
  const bitacora = crearBitacora(pool, cfg);
  const limite = crearLimite();
  const puerta = crearPuerta({ cfg, pool, origen, bitacora, limite });
  const desde = Date.now();

  const origenPermitido = (o) => !!o && (cfg.origenes.includes('*') || cfg.origenes.includes(o));
  function cabeceras(req, extra = {}) {
    const h = { 'Cache-Control': 'no-store', 'X-Content-Type-Options': 'nosniff', ...extra };
    const o = req.headers.origin;
    if (origenPermitido(o)) { h['Access-Control-Allow-Origin'] = o; h['Vary'] = 'Origin'; }
    return h;
  }
  const ipDe = (req) => {
    if (cfg.confiaProxy) { const x = String(req.headers['x-forwarded-for'] || '').split(',')[0].trim(); if (x) return x; }
    return req.socket.remoteAddress || '?';
  };
  const json = (req, res, estado, obj) => {
    res.writeHead(estado, cabeceras(req, { 'Content-Type': 'application/json; charset=utf-8' }));
    res.end(JSON.stringify(obj));
  };
  function leeCuerpo(req) {
    return new Promise((resolve, reject) => {
      let n = 0; const partes = [];
      req.on('data', (c) => { n += c.length; if (n > cfg.maxCuerpo) { reject(Object.assign(new Error('grande'), { grande: true })); req.destroy(); } else partes.push(c); });
      req.on('end', () => resolve(Buffer.concat(partes).toString('utf8')));
      req.on('error', reject);
    });
  }
  const igual = (a, b) => { const x = Buffer.from(String(a)), y = Buffer.from(String(b)); return x.length === y.length && crypto.timingSafeEqual(x, y); };

  const server = http.createServer(async (req, res) => {
    try {
      const url = new URL(req.url, 'http://x');
      if (req.method === 'OPTIONS') {
        res.writeHead(204, cabeceras(req, { 'Access-Control-Allow-Methods': 'POST, GET, OPTIONS', 'Access-Control-Allow-Headers': 'Content-Type', 'Access-Control-Max-Age': '86400' }));
        return res.end();
      }
      if (req.method === 'GET' && (url.pathname === '/salud' || url.pathname === '/')) {
        let db = null;
        if (pool) { try { await pool.query('SELECT 1'); db = true; } catch (e) { db = false; } }
        return json(req, res, 200, { ok: true, servidor: 'Rey Pizza', version: VERSION, baseDatos: db, appsScript: !!cfg.appsScript, activoMin: Math.round((Date.now() - desde) / 60000) });
      }
      if (req.method === 'GET' && url.pathname === '/velocidad') {
        if (!cfg.claveAdmin || !igual(url.searchParams.get('clave') || '', cfg.claveAdmin)) return json(req, res, 401, { ok: false, error: 'Falta la clave.' });
        await bitacora.vaciar();
        return json(req, res, 200, { ok: true, horas: Number(url.searchParams.get('horas')) || 24, consultas: await bitacora.resumen(Number(url.searchParams.get('horas')) || 24) });
      }
      if (req.method === 'POST' && url.pathname === '/') {
        let q;
        try { q = JSON.parse((await leeCuerpo(req)) || '{}'); }
        catch (e) { return json(req, res, 200, { ok: false, error: e.grande ? 'Lo enviado es demasiado grande.' : 'La app envió algo que no se entiende.' }); }
        const out = await puerta(q, ipDe(req));
        return json(req, res, 200, out);          // siempre 200: la app lee ok / error (igual que con Apps Script)
      }
      json(req, res, 404, { ok: false, error: 'No existe.' });
    } catch (e) {
      console.error(JSON.stringify({ nivel: 'error', msg: String(e && e.message ? e.message : e) }));
      try { json(req, res, 200, { ok: false, error: 'El servidor tuvo un problema. Intente de nuevo.' }); } catch (e2) {}
    }
  });
  server.keepAliveTimeout = 65000;     // más que el de los balanceadores, para no cortar conexiones
  server.headersTimeout = 66000;
  return { server, bitacora, origen };
}
