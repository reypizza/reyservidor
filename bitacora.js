/* Guarda cuánto tarda cada consulta, por detrás (nunca hace esperar a nadie) y sin guardar datos de la consulta. */
export function crearBitacora(pool, { diasBitacora = 30 } = {}) {
  let cola = [];
  let temporizador = null;
  const MAX = 500;

  async function vaciar() {
    if (!pool || !cola.length) { cola = []; return; }
    const lote = cola; cola = [];
    try {
      const vals = [], params = [];
      lote.forEach((x, i) => { const b = i * 5; vals.push(`($${b + 1},$${b + 2},$${b + 3},$${b + 4},$${b + 5})`); params.push(x.fn, x.nativa, x.ok, x.msTotal, x.msOrigen); });
      await pool.query(`INSERT INTO llamadas (fn, nativa, ok, ms_total, ms_origen) VALUES ${vals.join(',')}`, params);
    } catch (e) { /* si la base falla, se pierde la medición: nunca debe afectar a las consultas */ }
  }
  return {
    anota(x) {
      if (!pool) return;
      if (cola.length >= MAX) cola.shift();
      cola.push(x);
      if (!temporizador) { temporizador = setTimeout(() => { temporizador = null; vaciar(); }, 2000); temporizador.unref?.(); }
    },
    vaciar,
    async limpiar() {
      if (!pool) return;
      try {
        await pool.query(`DELETE FROM llamadas WHERE en < now() - ($1 || ' days')::interval`, [String(diasBitacora)]);
        await pool.query(`DELETE FROM solicitudes WHERE creada_en < now() - interval '2 days'`);
      } catch (e) {}
    },
    async resumen(horas = 24) {
      if (!pool) return [];
      const r = await pool.query(
        `SELECT fn, nativa, count(*)::int AS veces,
                round(avg(ms_total))::int AS prom_ms, round(percentile_cont(0.95) WITHIN GROUP (ORDER BY ms_total))::int AS p95_ms,
                max(ms_total)::int AS max_ms, round(avg(ms_origen))::int AS origen_prom_ms,
                count(*) FILTER (WHERE NOT ok)::int AS fallos
           FROM llamadas WHERE en > now() - ($1 || ' hours')::interval
          GROUP BY fn, nativa ORDER BY avg(ms_total) DESC`, [String(horas)]);
      return r.rows;
    },
  };
}
