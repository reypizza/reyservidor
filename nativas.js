/* Las consultas que ya contesta ESTE servidor (con su base de datos). Todo lo que no esté aquí se pasa a Apps Script.
 * Cada función recibe (ctx, ...args) y devuelve el valor; si algo está mal, lanza un Error con un mensaje en español.
 * ctx = { pool, rid, ip, cfg, usuarios, indicadores, reloj }
 * Una función puede devolver PASAR: «no estoy seguro, que conteste Apps Script».
 * Al migrar un módulo, sus funciones se agregan aquí, y desaparecen de Apps Script. */
import { hoyISO, ahoraTxt, horaTxt } from './fecha.js';
import { exigeAdmin, exigeOperaciones } from './auth.js';
import { tableroKpiAreas, resumenAreasCalc, SUC_KPI, AREA_CONTA } from './kpi.js';

export const PASAR = Symbol('pasar a Apps Script');
const FRESCO_MS = 30 * 60 * 1000;      // si hace más de 30 min que no se aprende de la hoja, mejor que conteste Google

export const nativas = {
  /** Solo contesta «aquí estoy»: sirve para medir la velocidad real hasta el servidor. */
  ping() { return { ok: true }; },

  /** Lo primero que pide la app al abrir (antes de pedir el PIN). Igual que la de Apps Script, pero sin esperar a Google. */
  async listaEntrada(ctx) {
    const est = ctx.usuarios && await ctx.usuarios.estado();
    if (!est || !est.ultimaOk || Date.now() - est.ultimaOk > FRESCO_MS) return PASAR;
    const ahora = new Date();
    return { usuarios: [], hoy: hoyISO(ahora), ahora: ahoraTxt(ahora), hay: est.activos,
      problema: est.activos ? '' : 'No hay ninguna persona activa.' };
  },

  /** El tablero de indicadores (pestaña «Tablero»): cumplimiento, checklists, tendencias y lo que más falla. */
  async tableroKpi(ctx, cred, dias) {
    const yo = await exigeOperaciones(ctx, cred);
    const areas = yo.esAdmin || yo.rol === 'operaciones' ? SUC_KPI.slice() : [AREA_CONTA];     // finanzas ve contabilidad
    if (!(await ctx.indicadores.alDia())) return PASAR;
    const { indicadores, registros } = await ctx.indicadores.leer(8000);
    const ahora = ctx.reloj ? ctx.reloj() : new Date();
    return tableroKpiAreas({ areas, dias, hoy: hoyISO(ahora), hora: horaTxt(ahora), indicadores, registros });
  },

  /** Las tarjetas de Inicio del administrador (anillos y líneas). Las señales (inventario, impuestos…) siguen en resumenAreasSenales. */
  async resumenAreas(ctx, cred) {
    await exigeAdmin(ctx, cred);
    if (!(await ctx.indicadores.alDia())) return PASAR;
    const t0 = Date.now();
    const { indicadores, registros, ultimaFila } = await ctx.indicadores.leer(8000);
    const ahora = ctx.reloj ? ctx.reloj() : new Date();
    const r = resumenAreasCalc({ hoy: hoyISO(ahora), hora: horaTxt(ahora), indicadores, registros, ultimaFila });
    return { ...r, ms: { base: Date.now() - t0, total: Date.now() - t0 } };
  },
};

/* Lo que cambia las hojas de indicadores: cuando una de estas pasa por aquí, la copia se actualiza de inmediato. */
export const invalidaIndicadores = new Set([
  'guardarKpi', 'borrarKpi', 'enviarKpis', 'agregarRecomendados',
  'guardarKpiCmo', 'borrarKpiCmo', 'enviarKpisCmo', 'agregarRecomendadosCmo',
  'guardarKpiTec', 'borrarKpiTec', 'enviarKpisTec', 'agregarRecomendadosTec',
  'registrarKpis', 'registrarKpisTec',
]);

/* Las que cambian datos y por eso se protegen con «rid» (no se ejecutan dos veces). Se llena al migrar. */
export const escriben = new Set([]);

/* Lecturas que, si algo falla aquí (la base de datos, por ejemplo), se vuelven a pedir a Apps Script en vez de dar error.
 * Solo vale mientras los datos de esa consulta sigan también en las hojas. */
export const conRespaldoEnOrigen = new Set(['listaEntrada', 'tableroKpi', 'resumenAreas']);
