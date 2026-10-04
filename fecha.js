/* La fecha y la hora de Guatemala (la misma zona que usa Apps Script), sin depender de la zona del servidor. */
const ZONA = 'America/Guatemala';
export function partes(d = new Date()) {
  const f = new Intl.DateTimeFormat('en-GB', { timeZone: ZONA, year: 'numeric', month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit', hourCycle: 'h23' });
  const o = {}; for (const p of f.formatToParts(d)) o[p.type] = p.value;
  return o;
}
export const hoyISO = (d) => { const p = partes(d); return `${p.year}-${p.month}-${p.day}`; };
/** «dd/MM/yyyy, HH:mm», el formato que usa la app para el reloj */
export const ahoraTxt = (d) => { const p = partes(d); return `${p.day}/${p.month}/${p.year}, ${p.hour}:${p.minute}`; };
