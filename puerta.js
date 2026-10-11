/* La «puerta»: recibe una consulta {fn, args, rid} y la contesta aquí (si ya está migrada) o se la pasa a Apps Script.
 * Devuelve siempre lo mismo que devolvía Apps Script: { ok:true, v, ms } o { ok:false, error }. Por eso la app
 * (el frontend) no necesita cambios: solo apunta a esta dirección en lugar de la de Google. */
import { nativas, escriben, conRespaldoEnOrigen, invalidaIndicadores, PASAR } from './nativas.js';
import { ErrorApp } from './errores.js';
import { unaVez } from './unavez.js';
import { Transitorio } from './escrituras.js';

const NOMBRE_VALIDO = /^[A-Za-z][A-Za-z0-9_]{0,63}$/;

export function crearPuerta({ cfg, pool, origen, bitacora, limite, usuarios, indicadores, reloj, copia, lecturas, escrituras, emite }) {
  return async function puerta(q, ip) {
    const t0 = Date.now();
    const fn = String((q && q.fn) || '');
    const args = Array.isArray(q && q.args) ? q.args : [];
    const rid = String((q && q.rid) || '').replace(/[^A-Za-z0-9_-]/g, '').slice(0, 40);
    const fin = (out, nativa, msOrigen) => {
      bitacora.anota({ fn: NOMBRE_VALIDO.test(fn) ? fn : '(inválida)', nativa, ok: !!out.ok, msTotal: Date.now() - t0, msOrigen: msOrigen ?? null });
      return out;
    };

    if (!NOMBRE_VALIDO.test(fn)) return fin({ ok: false, error: 'La app pidió algo que no existe.' }, true);

    // límites por IP: más estricto para entrar (así nadie prueba PINs sin parar)
    const maximo = fn === 'entrar' ? cfg.limiteEntrarPorMinuto : cfg.limitePorMinuto;
    if (!limite.pasa(ip + '|' + (fn === 'entrar' ? 'e' : 'g'), maximo))
      return fin({ ok: false, error: 'Demasiados intentos seguidos. Espere un minuto y vuelva a intentar.' }, true);

    // 1) ¿ya está migrada? Se contesta aquí, con la base de datos propia
    if (Object.prototype.hasOwnProperty.call(nativas, fn)) {
      try {
        const ctx = { pool, rid, ip, cfg, usuarios, indicadores, reloj };
        const correr = () => nativas[fn](ctx, ...args);
        const v = escriben.has(fn) ? await unaVez(pool, rid, fn, correr) : await correr();
        if (v !== PASAR) return fin({ ok: true, v: v === undefined ? null : v, ms: Date.now() - t0 }, true);
        // PASAR: esta vez no es seguro contestar aquí; sigue el camino de Apps Script
      } catch (e) {
        if (e instanceof ErrorApp || !conRespaldoEnOrigen.has(fn)) return fin({ ok: false, error: String(e && e.message ? e.message : e) }, true);
        // es una lectura que también sabe contestar Apps Script: si aquí falló algo, que conteste Google
      }
    }

    // 1a) guardado rápido: gastos e ingresos se guardan aquí primero y se reenvían a Google por detrás (solo si se encendió; por defecto está apagado)
    if (escrituras) {
      try {
        const e = await escrituras.ejecuta(fn, args, rid, String((q && q.sid) || '').replace(/[^A-Za-z0-9_-]/g, '').slice(0, 40));
        if (e) return fin({ ok: true, v: e.v === undefined ? null : e.v, ms: Date.now() - t0 }, true);
      } catch (e) {
        // nunca se pasa a Google desde aquí: podría quedar guardado dos veces. La app repite con la misma llave y entonces se sabrá si quedó.
        const m = String(e && e.message ? e.message : e);
        if (!(e instanceof Transitorio)) console.error(JSON.stringify({ nivel: 'error', msg: 'Guardado rápido: falla inesperada', fn, detalle: m.slice(0, 200) }));
        return fin({ ok: false, error: e instanceof Transitorio ? m : 'El servidor tuvo un problema. Intente de nuevo.' }, true);
      }
    }

    // 1b) lecturas completas: el código de la app corre aquí sobre la copia de las hojas (solo si se encendió; por defecto está apagado)
    if (lecturas && (cfg.lecturasModo === 'nativa' || (cfg.lecturasModo === 'auto' && await lecturas.graduada(fn)))) {
      const n = await lecturas.nativa(fn, args);
      if (n.ok) { lecturas.verificaDespues(fn, args, n.r); return fin(n.r, true); }
    }

    // 2) todavía no: se la pasamos a Apps Script, igual que hacía el teléfono, pero desde un servidor con mejor conexión
    if (escrituras && escrituras.ocupado()) await escrituras.esperaVacia(cfg.msBarrera);        // algo guardado aquí todavía no llegó a Google: se espera, para no leer ni escribir sobre algo viejo
    const r = await origen.llamar({ fn, args, rid });
    const msOrigen = r._ms; delete r._ms; delete r._caido;
    if (copia && r.tocadas) copia.marcaTocadas(r.tocadas);        // qué pestañas cambió esta ejecución: la copia las vuelve a traer
    const tocadas = r.tocadas || []; delete r.tocadas;
    if (emite && tocadas.length) emite({ tipo: 'cambio', fn, sid: String((q && q.sid) || '').replace(/[^A-Za-z0-9_-]/g, '').slice(0, 40) });      // algo cambió en las hojas: los demás teléfonos se ponen al día
    if (indicadores && invalidaIndicadores.has(fn)) indicadores.marcaSucio();      // algo pudo cambiar en las hojas de indicadores: la copia se pone al día
    if (lecturas && (cfg.lecturasModo === 'sombra' || cfg.lecturasModo === 'auto') && !tocadas.length && typeof r.ok === 'boolean') lecturas.sombra(fn, args, JSON.parse(JSON.stringify(r)), msOrigen);
    return fin(r, false, msOrigen);
  };
}
