/* Para lo que GUARDA datos. Si el teléfono repite la misma consulta (misma «rid») porque se cayó la señal,
 * el servidor no la ejecuta dos veces: devuelve lo que ya contestó. Igual que hacía Apps Script con su caché. */
export async function unaVez(pool, rid, fn, ejecutar) {
  if (!pool || !rid) return ejecutar();
  const ins = await pool.query(
    `INSERT INTO solicitudes (rid, fn, estado) VALUES ($1,$2,'en_curso') ON CONFLICT (rid) DO NOTHING RETURNING rid`, [rid, fn]);
  if (!ins.rowCount) {
    // ya existe: o terminó (se devuelve lo guardado) o sigue corriendo (se espera un poco)
    for (let i = 0; i < 30; i++) {
      const r = (await pool.query('SELECT estado, resultado FROM solicitudes WHERE rid=$1', [rid])).rows[0];
      if (!r) break;
      if (r.estado === 'ok') return r.resultado;
      if (r.estado === 'error') { await pool.query('DELETE FROM solicitudes WHERE rid=$1 AND estado=\'error\'', [rid]); break; }
      await new Promise((res) => setTimeout(res, 500));
    }
    const r2 = (await pool.query('SELECT estado, resultado FROM solicitudes WHERE rid=$1', [rid])).rows[0];
    if (r2 && r2.estado === 'ok') return r2.resultado;
    if (r2 && r2.estado === 'en_curso') throw new Error('El servidor sigue trabajando en eso. Espere un momento y revise antes de guardar de nuevo.');
    return unaVez(pool, rid, fn, ejecutar);
  }
  try {
    const v = await ejecutar();
    await pool.query(`UPDATE solicitudes SET estado='ok', resultado=$2, terminada_en=now() WHERE rid=$1`, [rid, JSON.stringify(v === undefined ? null : v)]);
    return v;
  } catch (e) {
    await pool.query(`DELETE FROM solicitudes WHERE rid=$1`, [rid]).catch(() => {});   // si falló, el reintento sí debe volver a correr
    throw e;
  }
}
