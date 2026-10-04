/* Reconoce a la persona igual que «quien» de Apps Script: mismos mensajes, mismas reglas, mismo resultado. */
import { verificaPin } from './clave.js';

const CIERRE_MIN = 10;     // cuánto dura el cierre tras demasiados intentos

/**
 * cred = { usuario, pin }.  opciones:
 *   cuenta: true solo cuando la persona aprieta «Entrar» (así el arranque automático nunca suma intentos fallidos)
 *   aux / cmo: esta función sí la pueden usar el acceso limitado / la CMO (igual que permiteAux / permiteCmo)
 */
export async function autoriza({ pool, cfg, usuarios }, cred, { cuenta = false, aux = false, cmo = false } = {}) {
  cred = cred || {};
  const nombre = String(cred.usuario || '').trim();
  const pin = String(cred.pin || '').trim();
  if (!nombre || !pin) throw new Error('Escoja su nombre y escriba su PIN.');
  const lc = nombre.toLowerCase();

  if (cuenta) {
    const f = (await pool.query('SELECT n FROM intentos WHERE nombre_lc=$1 AND hasta > now()', [lc])).rows[0];
    if (f && f.n >= cfg.maxIntentos) throw new Error(`Demasiados intentos con ese nombre. Espere ${CIERRE_MIN} minutos y vuelva a probar.`);
  }

  const busca = async () => {
    const u = (await pool.query('SELECT * FROM usuarios WHERE nombre_lc=$1', [lc])).rows[0];
    return u && u.activo && await verificaPin(pin, u.pin_hash, cfg.pinPepper) ? u : null;
  };
  let hallado = await busca();
  if (!hallado) {
    // ¿será que el usuario o su PIN son nuevos en la hoja y todavía no se aprendieron? Se revisa una vez antes de rechazar
    const est = await usuarios.estado();
    if (!est || !est.total) { usuarios.pedirSync().catch(() => {}); throw new Error('El servidor está aprendiendo quién es quién. Intente de nuevo en un momento.'); }
    await usuarios.pedirSync();
    hallado = await busca();
  }
  if (!hallado) {
    if (cuenta) await pool.query(
      `INSERT INTO intentos (nombre_lc, n, hasta) VALUES ($1, 1, now() + interval '${CIERRE_MIN} minutes')
       ON CONFLICT (nombre_lc) DO UPDATE SET n = CASE WHEN intentos.hasta > now() THEN intentos.n + 1 ELSE 1 END, hasta = now() + interval '${CIERRE_MIN} minutes'`, [lc]);
    throw new Error('Nombre o PIN incorrectos.');
  }
  if (cuenta) await pool.query('DELETE FROM intentos WHERE nombre_lc=$1', [lc]);

  const rol = hallado.rol;
  if (rol === 'auxiliar' && !aux) throw new Error('Su usuario tiene acceso limitado: solo registra los gastos de su lugar.');
  if (rol === 'cmo' && !aux && !cmo) throw new Error('Su usuario no entra a esta parte. Tiene su presupuesto, gastos, solicitudes e indicadores.');
  const conLugar = rol === 'gerente' || rol === 'auxiliar' || rol === 'cmo';
  if (conLugar && !hallado.sucursal) throw new Error('A su usuario le falta la sucursal o el lugar. Pida al administrador que lo ponga.');
  const conf = !!hallado.confirma;
  return {
    nombre: hallado.nombre, rol, esAdmin: rol === 'admin',
    codigoCorto: hallado.pin_largo < 6,
    sucursal: conLugar ? hallado.sucursal : '',
    limitado: rol === 'auxiliar' || rol === 'cmo',
    cmo: rol === 'cmo',
    valida: rol === 'admin' || rol === 'operaciones',
    recibe: rol === 'admin' || (conf && rol === 'registro'),
    banco: rol === 'admin' || (conf && rol === 'registro'),
    correo: hallado.correo || '',
    confirma: ['admin', 'operaciones', 'finanzas'].includes(rol) || (conf && rol === 'registro'),
    soloMira: rol === 'operaciones' || rol === 'finanzas' || rol === 'dueno',
    soloReportes: rol === 'dueno',
  };
}
