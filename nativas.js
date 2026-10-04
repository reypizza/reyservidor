/* Las consultas que ya contesta ESTE servidor (con su base de datos). Todo lo que no esté aquí se pasa a Apps Script.
 * Cada función recibe (ctx, ...args) y devuelve el valor; si algo está mal, lanza un Error con un mensaje en español.
 * ctx = { pool, rid, ip }
 * Al migrar un módulo, sus funciones se agregan aquí, y desaparecen de Apps Script. */
export const nativas = {
  /** Solo contesta «aquí estoy»: sirve para medir la velocidad real hasta el servidor. */
  ping() { return { ok: true }; },
};

/* Las que cambian datos y por eso se protegen con «rid» (no se ejecutan dos veces). Se llena al migrar. */
export const escriben = new Set([]);
