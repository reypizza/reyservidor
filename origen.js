/* Habla con el Apps Script de siempre. Mientras un módulo no esté migrado, ahí se contesta. */
export function crearOrigen(cfg, fetchFn = fetch) {
  async function llamar(cuerpo) {
    if (!cfg.appsScript) return { ok: false, error: 'El servidor todavía no sabe dónde está el Apps Script (falta APPS_SCRIPT_URL).' };
    const t0 = Date.now();
    let r, texto;
    try {
      // Apps Script contesta con un redirect (302): fetch lo sigue solo, y por eso el teléfono ya no tiene que hacerlo
      r = await fetchFn(cfg.appsScript, {
        method: 'POST',
        redirect: 'follow',
        headers: { 'Content-Type': 'text/plain;charset=utf-8' },
        body: JSON.stringify(cuerpo),
        signal: AbortSignal.timeout(cfg.msEsperaOrigen),
      });
      texto = await r.text();
    } catch (e) {
      const tarde = e && (e.name === 'TimeoutError' || e.name === 'AbortError');
      return { ok: false, error: tarde ? 'Google tardó demasiado en contestar. Espere un momento y vuelva a intentar.' : 'No se pudo hablar con Google. Revise la conexión e intente de nuevo.', _ms: Date.now() - t0, _caido: true };
    }
    let j;
    try { j = JSON.parse(texto); } catch (e) {
      // cuando Google se pasa de cuota o falla, contesta una página HTML en vez de JSON
      return { ok: false, error: 'Google no contestó como se esperaba (HTTP ' + r.status + '). Intente de nuevo en un momento.', _ms: Date.now() - t0, _caido: true };
    }
    if (!j || typeof j !== 'object') return { ok: false, error: 'Respuesta inesperada de Google.', _ms: Date.now() - t0, _caido: true };
    j._ms = Date.now() - t0;
    return j;
  }
  return { llamar };
}
