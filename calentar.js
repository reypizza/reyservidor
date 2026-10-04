/* Apps Script se «duerme» si nadie lo usa un rato, y la primera consulta tarda mucho más. Este temporizador
 * le toca la puerta cada pocos minutos para que siempre esté despierto. No lee ni guarda nada. */
export function iniciarCalentador({ cfg, origen, bitacora }) {
  if (!cfg.appsScript || !cfg.minutosCalentar) return { parar() {} };
  const tocar = async () => {
    const t0 = Date.now();
    const r = await origen.llamar({ fn: 'ping', args: [], rid: '' });
    bitacora.anota({ fn: '(calentar Apps Script)', nativa: false, ok: !!r.ok, msTotal: Date.now() - t0, msOrigen: r._ms ?? null });
  };
  const id = setInterval(() => { tocar().catch(() => {}); }, cfg.minutosCalentar * 60 * 1000);
  id.unref?.();
  tocar().catch(() => {});
  return { parar() { clearInterval(id); }, tocar };
}
