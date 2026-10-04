/* Aprende de la hoja «Usuarios» (a través de la puerta interna de Apps Script) y guarda una copia en la base de datos.
 * Por ahora la hoja es la que manda: si alguien cambia un PIN o agrega una persona, aquí se entera en pocos minutos
 * (o de inmediato si alguien falla al entrar: ver pedirSync). */
import { hashPin, verificaPin } from './clave.js';

export function crearUsuarios({ cfg, pool, origen }) {
  let enCurso = null;
  let ultimoPedido = 0;
  let temporizador = null;
  const listo = () => !!(pool && cfg.appsScript && cfg.serverToken && cfg.pinPepper);

  async function nota(campos) {
    const { huella = null, ok = false, error = null } = campos;
    await pool.query(
      `INSERT INTO sincronizacion (clave, huella, ultima_ok, ultimo_intento, error)
       VALUES ('usuarios', $1, CASE WHEN $2 THEN now() END, now(), $3)
       ON CONFLICT (clave) DO UPDATE SET
         huella = COALESCE($1, sincronizacion.huella),
         ultima_ok = CASE WHEN $2 THEN now() ELSE sincronizacion.ultima_ok END,
         ultimo_intento = now(), error = $3`, [huella, ok, error]);
  }

  async function sincronizarUna({ forzar }) {
    if (!listo()) return { ok: false, motivo: 'sin configurar' };
    try {
      const previa = (await pool.query(`SELECT huella FROM sincronizacion WHERE clave='usuarios'`)).rows[0];
      const r = await origen.llamar({ interno: true, fn: 'exportarUsuarios', token: cfg.serverToken, huella: forzar ? '' : (previa && previa.huella) || '' });
      if (!r.ok) { await nota({ error: String(r.error || 'Apps Script no contestó').slice(0, 300) }); return { ok: false, motivo: r.error }; }
      const v = r.v || {};
      if (v.sinCambios) { await nota({ huella: v.huella, ok: true }); return { ok: true, cambios: false }; }
      const lista = Array.isArray(v.usuarios) ? v.usuarios : [];
      const hay = Number((await pool.query('SELECT count(*)::int AS n FROM usuarios')).rows[0].n);
      // una lista vacía casi seguro es un fallo al leer la hoja: no se borra lo que ya se sabía
      if (!lista.length && hay) { await nota({ error: 'La hoja Usuarios llegó vacía; se conserva lo anterior.' }); return { ok: false, motivo: 'vacía' }; }

      const cli = await pool.connect();
      try {
        // se prepara todo (el scrypt tarda) antes de abrir la transacción; el último con el mismo nombre gana, como en Apps Script
        const porNombre = new Map();
        for (const u of lista) { const nombre = String(u.nombre || '').trim(); if (nombre) porNombre.set(nombre.toLowerCase(), u); }
        const previos = new Map((await cli.query('SELECT nombre_lc, pin_hash FROM usuarios')).rows.map((x) => [x.nombre_lc, x.pin_hash]));
        const filas = [];
        for (const [lc, u] of porNombre) {
          const pin = String(u.pin == null ? '' : u.pin).trim();
          let h = previos.get(lc);
          if (!(h && await verificaPin(pin, h, cfg.pinPepper))) h = await hashPin(pin, cfg.pinPepper);     // solo se vuelve a transformar si el PIN cambió
          filas.push([String(u.nombre).trim(), lc, h, pin.length, String(u.rol || 'registro'), u.activo !== false,
            String(u.sucursal || ''), !!u.confirma, String(u.correo || '')]);
        }
        await cli.query('BEGIN');
        for (const f of filas)
          await cli.query(
            `INSERT INTO usuarios (nombre, nombre_lc, pin_hash, pin_largo, rol, activo, sucursal, confirma, correo)
             VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9)
             ON CONFLICT (nombre_lc) DO UPDATE SET nombre=$1, pin_hash=$3, pin_largo=$4, rol=$5, activo=$6, sucursal=$7, confirma=$8, correo=$9, actualizado_en=now()`, f);
        await cli.query(`UPDATE usuarios SET activo=false, actualizado_en=now() WHERE activo AND nombre_lc <> ALL($1::text[])`, [filas.map((f) => f[1])]);
        await cli.query('COMMIT');
      } catch (e) { await cli.query('ROLLBACK').catch(() => {}); throw e; } finally { cli.release(); }
      await nota({ huella: v.huella, ok: true });
      return { ok: true, cambios: true, usuarios: lista.length };
    } catch (e) {
      try { await nota({ error: String(e.message || e).slice(0, 300) }); } catch (e2) {}
      return { ok: false, motivo: String(e.message || e) };
    }
  }

  /** Una sola sincronización a la vez: si ya hay una corriendo, todos esperan a esa. */
  function sincronizar({ forzar = false } = {}) {
    if (!enCurso) enCurso = sincronizarUna({ forzar }).finally(() => { enCurso = null; });
    return enCurso;
  }
  /** Cuando alguien falla al entrar, puede ser que su PIN o su usuario sean nuevos en la hoja: se revisa ya (máximo una vez cada 15 s). */
  function pedirSync() {
    const ahora = Date.now();
    if (ahora - ultimoPedido < 15000) return enCurso || Promise.resolve({ ok: true, cambios: false, omitido: true });
    ultimoPedido = ahora;
    return sincronizar({ forzar: true });
  }
  async function estado() {
    if (!pool) return null;
    const r = (await pool.query(`SELECT
        (SELECT count(*)::int FROM usuarios) AS total,
        (SELECT count(*)::int FROM usuarios WHERE activo) AS activos,
        (SELECT ultima_ok FROM sincronizacion WHERE clave='usuarios') AS ultima_ok,
        (SELECT ultimo_intento FROM sincronizacion WHERE clave='usuarios') AS ultimo_intento,
        (SELECT error FROM sincronizacion WHERE clave='usuarios') AS error`)).rows[0];
    return { total: r.total, activos: r.activos, ultimaOk: r.ultima_ok ? new Date(r.ultima_ok).getTime() : null,
      ultimoIntento: r.ultimo_intento ? new Date(r.ultimo_intento).getTime() : null, error: r.error || null };
  }
  async function porRol() {
    return (await pool.query(`SELECT rol, count(*)::int AS n FROM usuarios WHERE activo GROUP BY rol ORDER BY rol`)).rows;
  }
  function iniciar() {
    if (!listo() || !cfg.minutosSync) return;
    sincronizar().catch(() => {});
    temporizador = setInterval(() => { sincronizar().catch(() => {}); }, cfg.minutosSync * 60 * 1000);
    temporizador.unref?.();
  }
  return { sincronizar, pedirSync, estado, porRol, iniciar, listo, parar() { clearInterval(temporizador); } };
}
