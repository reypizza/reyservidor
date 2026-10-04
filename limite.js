/* Límite sencillo por IP: no deja que alguien pruebe PINs a lo loco ni sature el servidor. En memoria (una sola copia del servidor). */
export function crearLimite({ ventanaMs = 60000 } = {}) {
  const cubos = new Map();
  let ultimaLimpieza = Date.now();
  return {
    /** true = puede pasar; false = ya se pasó del máximo en esta ventana */
    pasa(clave, maximo, ahora = Date.now()) {
      if (ahora - ultimaLimpieza > ventanaMs * 2) {
        for (const [k, v] of cubos) if (ahora - v.ini > ventanaMs) cubos.delete(k);
        ultimaLimpieza = ahora;
      }
      let c = cubos.get(clave);
      if (!c || ahora - c.ini > ventanaMs) { c = { ini: ahora, n: 0 }; cubos.set(clave, c); }
      c.n++;
      return c.n <= maximo;
    },
  };
}
