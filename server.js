import http from 'node:http';
import crypto from 'node:crypto';
import { crearPuerta } from './puerta.js';
import { crearOrigen } from './origen.js';
import { crearBitacora } from './bitacora.js';
import { crearLimite } from './limite.js';
import { crearUsuarios } from './usuarios.js';
import { crearIndicadores } from './indicadores.js';
import { crearCopia } from './copia.js';
import { crearLecturas } from './lecturas.js';

const VERSION = '0.1.0';

export function crearServidor({ cfg, pool, fetchFn, reloj }) {
  const origen = crearOrigen(cfg, fetchFn);
  const bitacora = crearBitacora(pool, cfg);
  const limite = crearLimite();
  const usuarios = crearUsuarios({ cfg, pool, origen });
  const indicadores = crearIndicadores({ cfg, pool, origen });
  const copia = crearCopia({ cfg, pool, origen });
  const lecturas = crearLecturas({ cfg, pool, copia, origen });
  const puerta = crearPuerta({ cfg, pool, origen, bitacora, limite, usuarios, indicadores, reloj, copia, lecturas });
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
      if (req.method === 'GET' && url.pathname === '/estado') {
        if (!cfg.claveAdmin || !igual(url.searchParams.get('clave') || '', cfg.claveAdmin)) return json(req, res, 401, { ok: false, error: 'Falta la clave.' });
        let est = null, roles = [], ind = null;
        try { est = await usuarios.estado(); roles = est ? await usuarios.porRol() : []; ind = await indicadores.estado(); } catch (e) { est = { error: 'La base de datos no contestó.' }; }
        return json(req, res, 200, { ok: true, usuarios: est && !est.error ? { ...est, porRol: roles, sincronizacionLista: usuarios.listo(),
          ultimaOk: est.ultimaOk ? new Date(est.ultimaOk).toISOString() : null, ultimoIntento: est.ultimoIntento ? new Date(est.ultimoIntento).toISOString() : null } : est,
          indicadores: ind ? { ...ind, sincronizacionLista: indicadores.listo() } : null,
          lecturas: { ...lecturas.estado(), copia: copia.estado() } });
      }
      if (req.method === 'GET' && url.pathname === '/comparaciones') {
        if (!cfg.claveAdmin || !igual(url.searchParams.get('clave') || '', cfg.claveAdmin)) return json(req, res, 401, { ok: false, error: 'Falta la clave.' });
        let porFuncion = []; try { porFuncion = await lecturas.resumen(); } catch (e) {}
        const tot = porFuncion.reduce((a, x) => ({ veces: a.veces + x.veces, iguales: a.iguales + x.iguales }), { veces: 0, iguales: 0 });
        return json(req, res, 200, { ok: true, estado: lecturas.estado(), total: tot, funciones: porFuncion });
      }
      if (req.method === 'POST' && url.pathname === '/resincronizar') {
        if (!cfg.claveAdmin || !igual(url.searchParams.get('clave') || '', cfg.claveAdmin)) return json(req, res, 401, { ok: false, error: 'Falta la clave.' });
        return json(req, res, 200, { ok: true, indicadores: await indicadores.sincronizar({ completa: true }), usuarios: await usuarios.sincronizar({ forzar: true }), copia: copia.listo() ? await copia.sincronizar({ completa: true }) : null });
      }
      if (req.method === 'POST' && url.pathname === '/') {
        let q;
        try { q = JSON.parse((await leeCuerpo(req)) || '{}'); }
        catch (e) { return json(req, res, 200, { ok: false, error: e.grande ? 'Lo enviado es demasiado grande.' : 'La app envió algo que no se entiende.' }); }
        indicadores.actividad(); copia.actividad();
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
  return { server, bitacora, origen, usuarios, indicadores, copia, lecturas };
}
