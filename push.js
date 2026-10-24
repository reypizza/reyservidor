/* Avisos al celular del personal (Tripulación Real): cuando la gerente confirma la semana, Supabase deja el aviso en una cola
 * y este servidor lo manda (Web Push, sin librerías: firma VAPID ES256 y cifrado aes128gcm, RFC 8291 y 8292).
 * La llave de los avisos la crea este servidor la primera vez y la guarda en Supabase (tabla sin acceso público). Nunca va en el código. */
import crypto from 'node:crypto';

const b64u = (b) => Buffer.from(b).toString('base64').replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
const deB64u = (s) => Buffer.from(String(s).replace(/-/g, '+').replace(/_/g, '/'), 'base64');
const hmac = (llave, dato) => crypto.createHmac('sha256', llave).update(dato).digest();
// solo a los servicios de avisos de los navegadores (lo mismo que revisa Supabase al guardar)
const SERVICIOS = /^https:\/\/(fcm\.googleapis\.com|android\.googleapis\.com|updates\.push\.services\.mozilla\.com|[a-z0-9.-]+\.notify\.windows\.com|web\.push\.apple\.com)\//;

/** Una llave nueva: la pública (65 bytes, sin comprimir) y la privada (32 bytes), en base64url. */
export function llaveNueva() {
  const { publicKey, privateKey } = crypto.generateKeyPairSync('ec', { namedCurve: 'prime256v1' });
  const j = publicKey.export({ format: 'jwk' }), d = privateKey.export({ format: 'jwk' }).d;
  return { publica: b64u(Buffer.concat([Buffer.from([4]), deB64u(j.x), deB64u(j.y)])), privada: d };
}
function llavePrivada(publica, privada) {
  const p = deB64u(publica);
  return crypto.createPrivateKey({ key: { kty: 'EC', crv: 'P-256', x: b64u(p.subarray(1, 33)), y: b64u(p.subarray(33, 65)), d: privada }, format: 'jwk' });
}
/** El pase VAPID para un servicio de avisos (vale 12 horas). */
export function paseVapid(endpoint, llave, contacto = 'mailto:avisos@reypizza.gt') {
  const aud = new URL(endpoint).origin;
  const cab = b64u(JSON.stringify({ typ: 'JWT', alg: 'ES256' })), cuerpo = b64u(JSON.stringify({ aud, exp: Math.floor(Date.now() / 1000) + 12 * 3600, sub: contacto }));
  const firma = crypto.sign('sha256', Buffer.from(cab + '.' + cuerpo), { key: llavePrivada(llave.publica, llave.privada), dsaEncoding: 'ieee-p1363' });
  return 'vapid t=' + cab + '.' + cuerpo + '.' + b64u(firma) + ', k=' + llave.publica;
}
/** Cifra el aviso para un celular (aes128gcm). */
export function cifra(texto, p256dh, auth, sal = crypto.randomBytes(16), efimera = null) {
  const ua = deB64u(p256dh), sec = deB64u(auth);
  if (ua.length !== 65 || sec.length < 16) throw new Error('suscripción no válida');
  if (!efimera) { efimera = crypto.createECDH('prime256v1'); efimera.generateKeys(); }
  const as = efimera.getPublicKey(), compartido = efimera.computeSecret(ua);
  const prkK = hmac(sec, compartido), ikm = hmac(prkK, Buffer.concat([Buffer.from('WebPush: info\0'), ua, as, Buffer.from([1])]));
  const prk = hmac(sal, ikm);
  const cek = hmac(prk, Buffer.from('Content-Encoding: aes128gcm\0\x01')).subarray(0, 16), nonce = hmac(prk, Buffer.from('Content-Encoding: nonce\0\x01')).subarray(0, 12);
  const c = crypto.createCipheriv('aes-128-gcm', cek, nonce);
  const cuerpo = Buffer.concat([c.update(Buffer.concat([Buffer.from(texto), Buffer.from([2])])), c.final(), c.getAuthTag()]);
  const rs = Buffer.alloc(4); rs.writeUInt32BE(4096);
  return Buffer.concat([sal, rs, Buffer.from([as.length]), as, cuerpo]);
}

/** El repartidor de avisos: cada 30 segundos pide a Supabase lo que hay que mandar. */
export function iniciarAvisos({ cfg, copia, origen, log = () => {}, fetchFn = fetch, cada = 30000 }) {
  let cred = null, credMs = 0, temporizador = null, ocupado = false;
  async function credencial() {
    if (cred && Date.now() - credMs < 3600e3) return cred;
    const p = (copia && copia.propiedades && copia.propiedades()) || {};
    let url = process.env.SUPABASE_URL || p.SUPABASE_URL || '', key = process.env.SUPABASE_SERVICE_KEY || p.SUPABASE_SERVICE_KEY || '';
    if ((!url || !key) && origen && cfg.serverToken) {
      try { const r = await origen.llamar({ interno: true, fn: 'supaCredencial', token: cfg.serverToken }); if (r.ok && r.v) { url = r.v.url; key = r.v.key; } } catch (e) {}
    }
    url = String(url).replace(/\/+$/, '');
    if (!/^https:\/\/[a-z0-9-]+\.supabase\.co$/i.test(url) || String(key).length < 20) return null;
    cred = { url, key }; credMs = Date.now(); return cred;
  }
  async function rpc(c, fn, cuerpo) {
    const r = await fetchFn(c.url + '/rest/v1/rpc/' + fn, { method: 'POST', headers: { apikey: c.key, Authorization: 'Bearer ' + c.key, 'Content-Type': 'application/json' }, body: JSON.stringify(cuerpo) });
    const t = await r.text(); if (!r.ok) throw new Error('Supabase ' + r.status + ': ' + t.slice(0, 200));
    return t ? JSON.parse(t) : null;
  }
  async function vuelta() {
    if (ocupado) return; ocupado = true;
    try {
      const c = await credencial(); if (!c) return;
      let t = await rpc(c, 'per_push_trabajo', { p_max: 50 });
      let llave = t && t.llave;
      if (!llave) {                                        // la primera vez: la llave nace aquí y se guarda en Supabase (no se reemplaza nunca)
        const n = llaveNueva(); const pub = await rpc(c, 'per_push_llave_poner', { p_publica: n.publica, p_privada: n.privada });
        llave = pub === n.publica ? n : null;
        if (!llave) { t = await rpc(c, 'per_push_trabajo', { p_max: 0 }); llave = t && t.llave; }
        log({ nivel: 'info', msg: 'Avisos al celular: llave lista' });
      }
      const envios = (t && t.envios) || []; if (!envios.length || !llave) return;
      const hecho = { ok: [], muertos: [], error: {} };
      for (const e of envios) {
        const texto = JSON.stringify({ titulo: e.titulo, texto: e.texto, url: e.url || '/' });
        let alguno = false;
        for (const s of e.subs || []) {
          if (!SERVICIOS.test(s.endpoint) || /[\s\x00-\x1f]/.test(s.endpoint)) { hecho.muertos.push(s.endpoint); continue; }
          try {
            const r = await fetchFn(s.endpoint, { method: 'POST', redirect: 'error', body: cifra(texto, s.p256dh, s.auth),
              headers: { Authorization: paseVapid(s.endpoint, llave), TTL: '86400', 'Content-Encoding': 'aes128gcm', 'Content-Type': 'application/octet-stream', Urgency: 'normal' } });
            if (r.status === 404 || r.status === 410) hecho.muertos.push(s.endpoint);
            else if (r.status >= 200 && r.status < 300) alguno = true;
            else hecho.error[e.id] = 'El servicio de avisos contestó ' + r.status;
          } catch (x) { hecho.error[e.id] = String(x.message || x).slice(0, 150); }
        }
        if (alguno || !(e.subs || []).length) hecho.ok.push(e.id);       // sin celular registrado: queda como hecho (lo ve al abrir su app)
      }
      await rpc(c, 'per_push_hecho', { p: hecho });
      log({ nivel: 'info', msg: 'Avisos al celular', mandados: hecho.ok.length, quitados: hecho.muertos.length });
    } catch (x) { log({ nivel: 'error', msg: 'Avisos al celular', detalle: String(x.message || x).slice(0, 200) }); }
    finally { ocupado = false; }
  }
  temporizador = setInterval(vuelta, cada); temporizador.unref && temporizador.unref();
  const t0 = setTimeout(vuelta, 5000); if (t0.unref) t0.unref();
  return { vuelta, parar() { clearInterval(temporizador); } };
}
