/* Las consultas que ya contesta ESTE servidor (con su base de datos). Todo lo que no esté aquí se pasa a Apps Script.
 * Cada función recibe (ctx, ...args) y devuelve el valor; si algo está mal, lanza un Error con un mensaje en español.
 * ctx = { pool, rid, ip, cfg, usuarios }
 * Una función puede devolver PASAR: «no estoy seguro, que conteste Apps Script».
 * Al migrar un módulo, sus funciones se agregan aquí, y desaparecen de Apps Script. */
import { hoyISO, ahoraTxt } from './fecha.js';

export const PASAR = Symbol('pasar a Apps Script');
const FRESCO_MS = 30 * 60 * 1000;      // si hace más de 30 min que no se aprende de la hoja, mejor que conteste Google

export const nativas = {
  /** Solo contesta «aquí estoy»: sirve para medir la velocidad real hasta el servidor. */
  ping() { return { ok: true }; },

  /** Lo primero que pide la app al abrir (antes de pedir el PIN). Igual que la de Apps Script, pero sin esperar a Google. */
  async listaEntrada(ctx) {
    const est = ctx.usuarios && await ctx.usuarios.estado();
    if (!est || !est.ultimaOk || Date.now() - est.ultimaOk > FRESCO_MS) return PASAR;
    const ahora = new Date();
    return { usuarios: [], hoy: hoyISO(ahora), ahora: ahoraTxt(ahora), hay: est.activos,
      problema: est.activos ? '' : 'No hay ninguna persona activa.' };
  },
};

/* Las que cambian datos y por eso se protegen con «rid» (no se ejecutan dos veces). Se llena al migrar. */
export const escriben = new Set([]);

/* Lecturas que, si algo falla aquí (la base de datos, por ejemplo), se vuelven a pedir a Apps Script en vez de dar error.
 * Solo vale mientras los datos de esa consulta sigan también en las hojas. */
export const conRespaldoEnOrigen = new Set(['listaEntrada']);
