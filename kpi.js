/* La parte de cuentas de los indicadores: el cumplimiento de cada área, los checklists, las tendencias y lo que más falla.
 * Es el MISMO cálculo que hacía Apps Script (tableroKpiAreas y resumenAreasCalc), pero sobre la copia en la base de datos.
 * Las pruebas comparan los dos con los mismos datos y exigen resultados idénticos. */

export const SUC_KPI = ['centro', 'almendras', 'parque'];
export const AREA_CONTA = 'contabilidad', AREA_MKT = 'marketing', AREA_TEC = 'tecnologia';
const NOMBRES = { centro: 'Centro', almendras: 'Almendras', parque: 'Parque', contabilidad: 'Contabilidad', marketing: 'Marketing', tecnologia: 'Tecnología' };
export const nombreArea = (sid) => (Object.prototype.hasOwnProperty.call(NOMBRES, sid) ? NOMBRES[sid] : sid);
const INICIO_PM = { centro: '13:00', almendras: '13:30', parque: '13:30', contabilidad: '13:00', marketing: '13:00', tecnologia: '13:00' };
const ENVIADO = 'Enviado';

/* Fechas «yyyy-MM-dd» sin zona horaria: Guatemala no cambia de hora, así que basta con calendario puro. */
const aDia = (iso) => { const [y, m, d] = String(iso).split('-').map(Number); return new Date(Date.UTC(y, m - 1, d, 12)); };
export const masDias = (iso, n) => { const d = aDia(iso); d.setUTCDate(d.getUTCDate() + n); return d.toISOString().slice(0, 10); };
const diaSemana = (iso) => aDia(iso).getUTCDay();
export const lunesDe = (iso) => { const d = diaSemana(iso); return masDias(iso, d === 0 ? -6 : 1 - d); };

export function periodosKpi(k, sid, fecha, hora) {
  if (k.frec === 'Cada turno') {
    const ps = [{ turno: 'AM', per: fecha + '|AM' }];
    if (!hora || hora >= INICIO_PM[sid]) ps.push({ turno: 'PM', per: fecha + '|PM' });
    return ps;
  }
  if (k.frec === 'Semanal') return [{ turno: '', per: 'S' + lunesDe(fecha) }];
  if (k.frec === 'Mensual') return [{ turno: '', per: 'M' + fecha.slice(0, 7) }];
  if (k.frec === 'Apertura') return [{ turno: 'Apertura', per: fecha + '|A' }];
  if (k.frec === 'Cierre') return [{ turno: 'Cierre', per: fecha + '|C' }];
  return [{ turno: '', per: fecha }];
}
export const kpiEnAreas = (k, areas) => areas.some((a) => k.suc[a]);

/** indicadores = la lista completa en el orden de la hoja; registros = lo que Apps Script llama «las últimas 8000 filas» (ya sin las filas sin id). */
export function tableroKpiAreas({ areas, dias, hoy, hora, indicadores, registros }) {
  dias = Math.min(31, Math.max(7, Number(dias) || 14));
  const fechas = [];
  for (let k = dias - 1; k >= 0; k--) fechas.push(masDias(hoy, -k));
  const kpisDe = (sid, alDia) => indicadores.filter((k) => k.estado === ENVIADO && k.suc[sid] && (!alDia || String(k.desde).slice(0, 10) <= alDia));
  const idx = {};
  registros.forEach((x) => { idx[x.sucursal + '|' + x.id + '|' + x.per] = x; });
  const cumplimiento = {}, checklists = {}, numericos = {}, fallas = {};
  areas.forEach((sid) => {
    const suc = nombreArea(sid);
    cumplimiento[sid] = []; checklists[sid] = { Apertura: [], Cierre: [] };
    fechas.forEach((f) => {
      let esp = 0, hechos = 0;
      const ch = { Apertura: { t: 0, si: 0, hecho: 0 }, Cierre: { t: 0, si: 0, hecho: 0 } };
      kpisDe(sid, f).forEach((k) => {
        if (k.frec === 'Semanal' || k.frec === 'Mensual') return;
        periodosKpi(k, sid, f, f === hoy ? hora : null).forEach((p) => {
          const x = idx[suc + '|' + k.id + '|' + p.per];
          esp++; if (x) hechos++;
          if (k.frec === 'Apertura' || k.frec === 'Cierre') { const c = ch[k.frec]; c.t++; if (x) { c.hecho++; if (x.enRango === 'Sí') c.si++; } }
        });
      });
      // hoy todavía no termina: no se grafica para no parecer una caída
      cumplimiento[sid].push(esp && f !== hoy ? Math.round(hechos / esp * 100) : null);
      ['Apertura', 'Cierre'].forEach((m) => {
        const c = ch[m];
        checklists[sid][m].push(!c.t ? null : { t: c.t, hecho: c.hecho, si: c.si, e: c.hecho === 0 ? (f === hoy ? 'hoy' : 'falta') : c.si === c.t ? 'ok' : 'parcial' });
      });
    });
  });
  const ks = indicadores.filter((k) => k.estado === ENVIADO && kpiEnAreas(k, areas));
  const desde = fechas[0], h30 = masDias(hoy, -29);
  ks.forEach((k) => {
    if (k.tipo !== 'Número') return;
    const serie = {};
    areas.forEach((sid) => {
      if (!k.suc[sid]) return;
      const suc = nombreArea(sid), porDia = {};
      registros.forEach((x) => {
        if (x.id !== k.id || x.sucursal !== suc || x.fecha < desde) return;
        const v = Number(x.valor); if (!Number.isFinite(v)) return;
        (porDia[x.fecha] = porDia[x.fecha] || []).push(v);
      });
      serie[sid] = fechas.map((f) => { const l = porDia[f]; if (!l) return null; return Math.round(l.reduce((a, b) => a + b, 0) / l.length * 100) / 100; });
    });
    numericos[k.id] = { id: k.id, nombre: k.nombre, unidad: k.unidad, min: k.min, max: k.max, frec: k.frec, serie };
  });
  const nombresArea = areas.map(nombreArea);
  registros.forEach((x) => {
    if (x.enRango !== 'No' || x.fecha < h30 || nombresArea.indexOf(x.sucursal) < 0) return;
    const f = fallas[x.id] || (fallas[x.id] = { nombre: x.nombre, n: 0, por: {} });
    f.n++; f.por[x.sucursal] = (f.por[x.sucursal] || 0) + 1;
  });
  return { fechas, areas: areas.map((a) => ({ id: a, nombre: nombreArea(a) })), cumplimiento, checklists,
    numericos: Object.keys(numericos).map((k) => numericos[k]),
    fallas: Object.keys(fallas).map((k) => fallas[k]).sort((a, b) => b.n - a.n).slice(0, 6) };
}

const nivelPct = (p) => (p == null ? 'sin' : p >= 90 ? 'ok' : p >= 70 ? 'alerta' : 'mal');

/** Lo que muestran las tarjetas de Inicio del administrador (sin las «señales», que salen de inventario, impuestos, etc.). */
export function resumenAreasCalc({ hoy, hora, indicadores, registros, ultimaFila }) {
  const desde7 = masDias(hoy, -6);
  const tb = tableroKpiAreas({ areas: SUC_KPI.concat([AREA_CONTA, AREA_MKT, AREA_TEC]), dias: 14, hoy, hora, indicadores, registros });
  const primera5000 = Math.max(2, ultimaFila - 5000 + 1);
  const regs = registros.filter((x) => x.fila >= primera5000);
  const activos = indicadores.filter((k) => k.estado === ENVIADO);
  const prom = (l) => { l = l.filter((x) => x != null); return l.length ? Math.round(l.reduce((a, b) => a + b, 0) / l.length) : null; };
  const armar = (id, nombre, ids, detalle) => {
    const serie = tb.fechas.map((f, i) => prom(ids.map((a) => tb.cumplimiento[a][i])));
    const pct = prom(serie.slice(-8)), antes = prom(serie.slice(0, -8));
    const nombres = ids.map(nombreArea);
    const fuera = regs.filter((x) => x.enRango === 'No' && x.fecha >= desde7 && nombres.indexOf(x.sucursal) >= 0).length;
    const total = activos.filter((k) => kpiEnAreas(k, ids)).length;
    return { id, nombre, pct, antes, nivel: nivelPct(pct), serie, indicadores: total, fuera, extras: [], detalle: detalle || [] };
  };
  const detOps = SUC_KPI.map((s) => ({ id: s, nombre: nombreArea(s), pct: prom(tb.cumplimiento[s].slice(-8)) }));
  return { hoy, fechas: tb.fechas, areas: [
    armar('operaciones', 'Operaciones', SUC_KPI, detOps),
    armar('finanzas', 'Finanzas', [AREA_CONTA]),
    armar('marketing', 'Marketing', [AREA_MKT]),
    armar('tecnologia', 'Tecnología', [AREA_TEC]),
  ] };
}
