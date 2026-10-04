/* Cifra y comprime la copia antes de guardarla en la base de datos (AES-256-GCM). La clave sale de PIN_PEPPER, que nunca está en la base. */
import crypto from 'node:crypto';
import zlib from 'node:zlib';
const clave = (secreto) => crypto.createHash('sha256').update('copia-v1|' + secreto).digest();
export function cifra(obj, secreto) {
  if (!secreto) throw new Error('Falta PIN_PEPPER para cifrar la copia.');
  const iv = crypto.randomBytes(12), c = crypto.createCipheriv('aes-256-gcm', clave(secreto), iv);
  const d = Buffer.concat([c.update(zlib.gzipSync(Buffer.from(JSON.stringify(obj), 'utf8'))), c.final()]);
  return Buffer.concat([iv, c.getAuthTag(), d]);
}
export function descifra(buf, secreto) {
  const iv = buf.subarray(0, 12), tag = buf.subarray(12, 28), d = buf.subarray(28);
  const c = crypto.createDecipheriv('aes-256-gcm', clave(secreto), iv); c.setAuthTag(tag);
  return JSON.parse(zlib.gunzipSync(Buffer.concat([c.update(d), c.final()])).toString('utf8'));
}
