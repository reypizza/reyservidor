/* Aviso en tiempo real: cuando algo cambia (se guardó un registro, se aprobó algo…), todos los teléfonos conectados se enteran en
 * milisegundos y se ponen al día. Es un canal de una sola vía (Server-Sent Events). NO lleva datos: solo «algo cambió» y quién lo hizo
 * (un código al azar de ese teléfono, para que el que guardó no se avise a sí mismo). Lo que cada quien ve se vuelve a pedir con su PIN. */
export function crearEventos({ maxConexiones = 600, cadaMs = 20000 } = {}) {
  const conexiones = new Set(); let n = 0, enviados = 0;
  const latido = setInterval(() => { for (const c of conexiones) { try { c.res.write(': latido\n\n'); } catch (e) { conexiones.delete(c); } } }, cadaMs);
  latido.unref?.();
  return {
    /** Registra una conexión abierta. `cabeceras` ya trae el CORS. */
    abre(req, res, cabeceras, sid) {
      if (conexiones.size >= maxConexiones) { res.writeHead(503, cabeceras); return res.end(); }
      res.writeHead(200, { ...cabeceras, 'Content-Type': 'text/event-stream; charset=utf-8', 'Cache-Control': 'no-cache, no-transform', 'Connection': 'keep-alive', 'X-Accel-Buffering': 'no' });
      res.write('retry: 2000\n\n: listo\n\n');
      const c = { res, sid: String(sid || '').slice(0, 40) }; conexiones.add(c);
      req.on('close', () => conexiones.delete(c));
    },
    /** Avisa a todos que algo cambió. */
    emite(ev) {
      if (!conexiones.size) return;
      const dato = JSON.stringify({ n: ++n, t: Date.now(), fn: String((ev && ev.fn) || ''), sid: String((ev && ev.sid) || '') });
      for (const c of conexiones) { try { c.res.write('id: ' + n + '\nevent: cambio\ndata: ' + dato + '\n\n'); enviados++; } catch (e) { conexiones.delete(c); } }
    },
    estado: () => ({ conectados: conexiones.size, avisos: n, enviados }),
    parar() { clearInterval(latido); for (const c of conexiones) { try { c.res.end(); } catch (e) {} } conexiones.clear(); },
  };
}
