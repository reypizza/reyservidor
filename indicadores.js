/* Una copia al día de las hojas «Indicadores» e «Indicadores registros» en la base de datos.
 *
 * La HOJA sigue mandando (todo lo que cambia datos sigue pasando por Apps Script). Esta copia solo sirve para contestar
 * LECTURAS rápido. Para que nunca se vea algo viejo:
 *   - tras cada cambio que pasa por este servidor se actualiza de inmediato (marcaSucio) y la siguiente lectura espera a eso;
 *   - con actividad, se revisa la hoja cada minuto (cada 5 minutos si nadie usa la app);
 *   - si por algo la copia tiene más de 3 minutos sin actualizarse, las lecturas las contesta Apps Script, como antes;
 *   - cada 6 horas se vuelve a traer todo, por si alguien editó filas viejas a mano. */
export function crearIndicadores({ cfg, pool, origen }) {
  let enCurso = null, marcaVer = 0, cargado = false;
  const est = { ultimaOkMs: 0, huella: '', cursor: 2, ultimaFila: 0, completaMs: 0, ultimoIntentoMs: 0, verOk: 0, consistente: true };
  let actividadMs = 0, temporizador = null, primeraVez = true;
  const listo = () => !!(pool && cfg.appsScript && cfg.serverToken);

  async function carga() {
    if (cargado || !pool) return;
    const r = (await pool.query(`SELECT huella, ultima_ok, cursor_fila, ultima_fila, completa_en FROM sincronizacion WHERE clave='indicadores'`)).rows[0];
    if (r) Object.assign(est, { huella: r.huella || '', ultimaOkMs: r.ultima_ok ? new Date(r.ultima_ok).getTime() : 0, cursor: r.cursor_fila, ultimaFila: r.ultima_fila, completaMs: r.completa_en ? new Date(r.completa_en).getTime() : 0 });
    cargado = true;
  }
  async function guarda(extra = {}) {
    const { ok = false, error = null, completa = false } = extra;
    await pool.query(
      `INSERT INTO sincronizacion (clave, huella, ultima_ok, ultimo_intento, error, cursor_fila, ultima_fila, completa_en)
       VALUES ('indicadores', $1, CASE WHEN $2 THEN now() END, now(), $3, $4, $5, CASE WHEN $6 THEN now() END)
       ON CONFLICT (clave) DO UPDATE SET huella=$1, ultima_ok = CASE WHEN $2 THEN now() ELSE sincronizacion.ultima_ok END, ultimo_intento=now(), error=$3,
         cursor_fila=$4, ultima_fila=$5, completa_en = CASE WHEN $6 THEN now() ELSE sincronizacion.completa_en END`,
      [est.huella || null, ok, error, est.cursor, est.ultimaFila, completa]);
  }

  async function una({ completa }) {
    if (!listo()) return { ok: false, motivo: 'sin configurar' };
    await carga();
    const v0 = marcaVer;
    est.ultimoIntentoMs = Date.now();
    try {
      const debeCompleta = completa || !est.completaMs || Date.now() - est.completaMs > cfg.horasCompletaKpi * 3600 * 1000;
      let cursor = debeCompleta ? 2 : est.cursor, huella = debeCompleta ? '' : est.huella, ultimaFila = est.ultimaFila, paginas = 0, nuevos = 0, huboCambioLista = false;
      let reinicio = debeCompleta;
      if (debeCompleta) est.consistente = false;        // se van a reescribir filas que ya estaban: mientras tanto nadie lee la copia
      while (true) {
        const r = await origen.llamar({ interno: true, fn: 'exportarIndicadores', token: cfg.serverToken, huella, desdeFila: cursor, max: cfg.paginaKpi });
        if (!r.ok) throw new Error(String(r.error || 'Apps Script no contestó'));
        const v = r.v || {};
        // la hoja tiene menos filas de las que ya se habían traído: alguien borró filas. Se empieza de cero.
        if (!reinicio && v.ultimaFila < cursor - 1) { reinicio = true; est.consistente = false; cursor = 2; huella = ''; continue; }
        const cli = await pool.connect();
        try {
          await cli.query('BEGIN');
          reinicio = false;      // (traer todo desde la fila 2 reescribe cada fila; lo que sobra se borra más abajo)
          if (Array.isArray(v.indicadores)) {
            await cli.query('DELETE FROM indicadores');
            const ls = v.indicadores;
            for (let i = 0; i < ls.length; i += 200) {
              const vals = [], par = [];
              ls.slice(i, i + 200).forEach((k, j) => { const b = j * 3; vals.push(`($${b + 1},$${b + 2},$${b + 3}::jsonb)`); par.push(i + j, String(k.id || ''), JSON.stringify(k)); });
              await cli.query(`INSERT INTO indicadores (orden, id, doc) VALUES ${vals.join(',')}`, par);
            }
            huboCambioLista = true;
          }
          const rs = Array.isArray(v.registros) ? v.registros : [];
          for (let i = 0; i < rs.length; i += 500) {
            const vals = [], par = [];
            rs.slice(i, i + 500).forEach((x, j) => { const b = j * 3; vals.push(`($${b + 1},$${b + 2},$${b + 3}::jsonb)`); par.push(x.fila, String(x.id || ''), JSON.stringify(x)); });
            await cli.query(`INSERT INTO indicadores_registros (fila, id, doc) VALUES ${vals.join(',')} ON CONFLICT (fila) DO UPDATE SET id=EXCLUDED.id, doc=EXCLUDED.doc`, par);
          }
          // si la hoja ya no llega hasta donde había llegado antes, se borra lo que sobra
          await cli.query('DELETE FROM indicadores_registros WHERE fila > $1', [v.ultimaFila]);
          await cli.query('COMMIT');
          nuevos += rs.length;
        } catch (e) { await cli.query('ROLLBACK').catch(() => {}); throw e; } finally { cli.release(); }
        huella = v.huella; ultimaFila = v.ultimaFila; cursor = Math.max(cursor, v.hasta + 1); paginas++;
        if (!v.hayMas) break;
        if (paginas > 60) throw new Error('La hoja de registros es demasiado grande para traerla de una vez.');
      }
      Object.assign(est, { huella, cursor: Math.max(cursor, ultimaFila + 1), ultimaFila });
      est.ultimaOkMs = Date.now(); est.consistente = true; if (marcaVer === v0) est.verOk = marcaVer;
      if (debeCompleta) est.completaMs = Date.now();
      await guarda({ ok: true, completa: debeCompleta });
      return { ok: true, nuevos, completa: debeCompleta, listaCambio: huboCambioLista };
    } catch (e) {
      try { await guarda({ error: String(e.message || e).slice(0, 300) }); } catch (e2) {}
      return { ok: false, motivo: String(e.message || e) };
    }
  }

  /** Una sola sincronización a la vez: quien llegue mientras otra corre, espera a esa. */
  function sincronizar({ completa = false } = {}) {
    if (!enCurso) enCurso = una({ completa }).finally(() => { enCurso = null; });
    return enCurso;
  }
  /** Se avisa cuando algo pasó por aquí que cambia las hojas: la siguiente lectura va a esperar a que la copia lo alcance. */
  function marcaSucio() { marcaVer++; if (listo()) sincronizar().catch(() => {}); }
  function actividad() { actividadMs = Date.now(); }

  /** ¿La copia está lista para contestar con confianza? Si hace falta, primero se actualiza. false = que conteste Apps Script. */
  async function alDia() {
    if (!listo()) return false;
    await carga();
    const nuncaTuvo = !est.ultimaOkMs;
    const vieja = Date.now() - est.ultimaOkMs > cfg.frescoKpiMs;
    if (nuncaTuvo || vieja || est.verOk !== marcaVer) {
      const p = sincronizar();
      // la primera carga puede tardar (hoja grande): no se hace esperar a nadie más de unos segundos
      await Promise.race([p, new Promise((r) => setTimeout(r, nuncaTuvo ? cfg.esperaPrimeraKpiMs : 20000))]);
      // si se marcó sucio mientras corría esa, otra vuelta (una sola)
      if (est.verOk !== marcaVer && est.ultimaOkMs && Date.now() - est.ultimaOkMs < cfg.frescoKpiMs) await Promise.race([sincronizar(), new Promise((r) => setTimeout(r, 20000))]);
    }
    if (!est.consistente && enCurso) await Promise.race([enCurso, new Promise((r) => setTimeout(r, 20000))]);
    return !!est.ultimaOkMs && est.consistente && Date.now() - est.ultimaOkMs <= cfg.maxViejaKpiMs && est.verOk === marcaVer;
  }

  /** Los datos que usa el tablero: la lista y las últimas n filas (las que Apps Script llamaba «registrosKpi(n)»). */
  async function leer(n = 8000) {
    const desde = Math.max(2, est.ultimaFila - n + 1);
    const [a, b] = await Promise.all([
      pool.query('SELECT doc FROM indicadores ORDER BY orden'),
      pool.query(`SELECT doc FROM indicadores_registros WHERE fila >= $1 AND id <> '' ORDER BY fila`, [desde]),
    ]);
    return { indicadores: a.rows.map((r) => r.doc), registros: b.rows.map((r) => r.doc), ultimaFila: est.ultimaFila };
  }
  async function estado() {
    if (!pool) return null;
    await carga();
    const r = (await pool.query(`SELECT (SELECT count(*)::int FROM indicadores) AS indicadores, (SELECT count(*)::int FROM indicadores_registros) AS filas,
      (SELECT error FROM sincronizacion WHERE clave='indicadores') AS error`)).rows[0];
    return { indicadores: r.indicadores, filas: r.filas, ultimaFila: est.ultimaFila, error: r.error || null, ultimaOk: est.ultimaOkMs ? new Date(est.ultimaOkMs).toISOString() : null,
      completaEn: est.completaMs ? new Date(est.completaMs).toISOString() : null };
  }
  function iniciar() {
    if (!listo() || !cfg.segundosKpi) return;
    sincronizar().catch(() => {});
    temporizador = setInterval(() => {
      const hayActividad = Date.now() - actividadMs < 10 * 60 * 1000;
      const cada = (hayActividad ? cfg.segundosKpi : 300) * 1000;
      if (Date.now() - est.ultimoIntentoMs >= cada) sincronizar().catch(() => {});
    }, 15000);
    temporizador.unref?.();
  }
  return { sincronizar, marcaSucio, actividad, alDia, leer, estado, iniciar, listo, parar() { clearInterval(temporizador); } };
}
