/* El PIN nunca se guarda legible. Se transforma así: HMAC con la clave secreta del servidor (PIN_PEPPER) y luego
 * scrypt con una sal propia de cada persona. Quien se robe solo la base de datos no puede probar PIN: le falta la clave secreta. */
import crypto from 'node:crypto';

const scrypt = (clave, sal) => new Promise((res, rej) =>
  crypto.scrypt(clave, sal, 32, { N: 16384, r: 8, p: 1, maxmem: 64 * 1024 * 1024 }, (e, k) => (e ? rej(e) : res(k))));

/** Igual que «mismoPin» de Apps Script: sin espacios, y 012345 vale lo mismo que 12345 (la hoja a veces guarda el PIN como número). */
export function normalizaPin(p) {
  p = String(p == null ? '' : p).trim();
  if (/^\d+$/.test(p)) p = p.replace(/^0+(?=\d)/, '');
  return p;
}
const mezcla = (pin, pepper) => crypto.createHmac('sha256', pepper).update(normalizaPin(pin)).digest();

export async function hashPin(pin, pepper) {
  if (!pepper) throw new Error('Falta PIN_PEPPER.');
  if (!normalizaPin(pin)) return '';
  const sal = crypto.randomBytes(16);
  const h = await scrypt(mezcla(pin, pepper), sal);
  return `s1$${sal.toString('hex')}$${h.toString('hex')}`;
}

export async function verificaPin(pin, guardado, pepper) {
  if (!guardado || !pepper || !normalizaPin(pin)) return false;
  const [v, sal, h] = String(guardado).split('$');
  if (v !== 's1' || !sal || !h) return false;
  const calc = await scrypt(mezcla(pin, pepper), Buffer.from(sal, 'hex'));
  const esperado = Buffer.from(h, 'hex');
  return calc.length === esperado.length && crypto.timingSafeEqual(calc, esperado);
}
