/* Una «hoja de cálculo» en memoria que se comporta como la de Google para lo que usa el código de la app.
 * Los valores son: '' | texto | número | true/false | Date (igual que devuelve getValues en Apps Script). */

const colANum = (s) => { let n = 0; for (const ch of s.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64); return n; };
export const numACol = (n) => { let s = ''; while (n > 0) { const r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = Math.floor((n - 1) / 26); } return s; };

/** «A1», «A1:C5», «B:B», «A:C», «2:4» → {r1,c1,r2,c2}  (r2 / c2 = 0 significa «hasta el final») */
export function parseA1(a1) {
  const t = String(a1).replace(/\$/g, '').trim();
  const celda = (x) => { const m = /^([A-Za-z]*)(\d*)$/.exec(x); if (!m || (!m[1] && !m[2])) throw new Error('Rango no válido: ' + a1); return { c: m[1] ? colANum(m[1]) : 0, r: m[2] ? Number(m[2]) : 0 }; };
  const [i, f] = t.split(':'); const a = celda(i), b = f ? celda(f) : a;
  return { r1: a.r || 1, c1: a.c || 1, r2: b.r || 0, c2: b.c || 0 };
}

const NOOPS = new Set(['setFontWeight', 'setFontColor', 'setFontSize', 'setFontFamily', 'setFontStyle', 'setBackground', 'setBackgrounds', 'setHorizontalAlignment', 'setVerticalAlignment',
  'setWrap', 'setWrapStrategy', 'setBorder', 'setNote', 'setDataValidation', 'setDataValidations', 'merge', 'mergeAcross', 'breakApart', 'setTextRotation', 'setRichTextValue', 'protect',
  'setConditionalFormatRules', 'activate', 'setFontLine', 'setBackgroundRGB', 'setShowHyperlink', 'setTextStyle', 'setFontColors', 'setFontWeights']);

function interpreta(v) {          // lo que Google hace con un texto escrito en una celda que no es de «texto sin formato»
  if (typeof v !== 'string') return v;
  const t = v.trim();
  if (t === '') return v;
  if (/^[-+]?\d+(\.\d+)?$/.test(t) && !/^[-+]?0\d/.test(t)) return Number(t);
  if (/^(true|false)$/i.test(t)) return t.toLowerCase() === 'true';
  return v;
}

export class Rango {
  constructor(hoja, r, c, nr, nc) { this.hoja = hoja; this.r = r; this.c = c; this.nr = nr; this.nc = nc; }
  getRow() { return this.r; } getColumn() { return this.c; } getNumRows() { return this.nr; } getNumColumns() { return this.nc; } getSheet() { return this.hoja; }
  getValues() { const o = []; for (let i = 0; i < this.nr; i++) { const fila = []; for (let j = 0; j < this.nc; j++) fila.push(this.hoja._lee(this.r + i, this.c + j)); o.push(fila); } return o; }
  getDisplayValues() { return this.getValues().map((f) => f.map((v) => (v instanceof Date ? v.toISOString() : String(v)))); }
  getValue() { return this.hoja._lee(this.r, this.c); }
  getNumberFormats() { const o = []; for (let i = 0; i < this.nr; i++) { const f = []; for (let j = 0; j < this.nc; j++) f.push(this.hoja._esTexto(this.r + i, this.c + j) ? '@' : 'General'); o.push(f); } return o; }
  getFormula() { return this.hoja.formulas.get(this.r + ',' + this.c) || ''; }
  setValues(vals) {
    if (!Array.isArray(vals) || (this.nr && vals.length !== this.nr) || (vals[0] && vals[0].length !== this.nc))
      throw new Error(`Los datos (${vals && vals.length} por ${vals && vals[0] && vals[0].length}) no coinciden con el rango (${this.nr} por ${this.nc}).`);
    vals.forEach((fila, i) => fila.forEach((v, j) => this.hoja._escribe(this.r + i, this.c + j, v)));
    return this;
  }
  setValue(v) { for (let i = 0; i < this.nr; i++) for (let j = 0; j < this.nc; j++) this.hoja._escribe(this.r + i, this.c + j, v); return this; }
  setFormula(f) { this.hoja.formulas.set(this.r + ',' + this.c, String(f)); this.hoja._escribe(this.r, this.c, ''); return this; }
  setFormulas(fs) { fs.forEach((fila, i) => fila.forEach((f, j) => { if (f) { this.hoja.formulas.set((this.r + i) + ',' + (this.c + j), String(f)); this.hoja._escribe(this.r + i, this.c + j, ''); } })); return this; }
  clearContent() { for (let i = 0; i < this.nr; i++) for (let j = 0; j < this.nc; j++) this.hoja._borra(this.r + i, this.c + j); return this; }
  clear() { this.clearContent(); this.setNumberFormat('General'); return this; }
  setNumberFormat(f) { this.hoja._formato(this.r, this.c, this.nr, this.nc, f); return this; }
  setNumberFormats(fs) { fs.forEach((fila, i) => fila.forEach((f, j) => this.hoja._formato(this.r + i, this.c + j, 1, 1, f))); return this; }
  sort(espec) {
    const claves = (Array.isArray(espec) ? espec : [espec]).map((e) => (typeof e === 'number' ? { column: e, ascending: true } : { column: e.column, ascending: e.ascending !== false }));
    const filas = this.getValues();
    const cmp = (a, b) => { const x = a instanceof Date ? a.getTime() : a, y = b instanceof Date ? b.getTime() : b;
      if (x === y) return 0; if (x === '') return 1; if (y === '') return -1;
      if (typeof x === 'number' && typeof y === 'number') return x - y; return String(x).toLowerCase() < String(y).toLowerCase() ? -1 : 1; };
    const ordenadas = filas.map((f, i) => [f, i]).sort((p, q) => { for (const k of claves) { const col = k.column - this.c; const d = cmp(p[0][col], q[0][col]); if (d) return k.ascending ? d : -d; } return p[1] - q[1]; }).map((p) => p[0]);
    ordenadas.forEach((f, i) => f.forEach((v, j) => this.hoja._ponCruda(this.r + i, this.c + j, v)));
    return this;
  }
  getA1Notation() { return numACol(this.c) + this.r + ':' + numACol(this.c + this.nc - 1) + (this.r + this.nr - 1); }
}
// los métodos de formato (colores, negritas, bordes…) no cambian datos: se aceptan y no hacen nada
for (const m of NOOPS) Rango.prototype[m] = function () { return this; };

export class Hoja {
  constructor(libro, nombre) { this.libro = libro; this.nombre = nombre; this.filas = []; this.maxFilas = 1000; this.maxCols = 26; this.textoCols = []; this.formulas = new Map(); this.sucia = true; }
  _lee(r, c) { const f = this.filas[r - 1]; const v = f ? f[c - 1] : undefined; return v === undefined || v === null ? '' : v; }
  _igual(a, b) { return a instanceof Date && b instanceof Date ? a.getTime() === b.getTime() : a === b; }
  _ponCruda(r, c, v) {
    if (this.libro.soloLectura) { if (!this._igual(this._lee(r, c), v)) this.libro._intento(r > 1 && this.libro.hojas.includes(this)); return; }     // en modo lectura nada cambia (escribir en una pestaña que no existe o en su fila de títulos es «reparar», no «guardar datos»)
    while (this.filas.length < r) this.filas.push([]); const f = this.filas[r - 1]; while (f.length < c - 1) f.push(''); f[c - 1] = v; if (r > this.maxFilas) this.maxFilas = r; if (c > this.maxCols) this.maxCols = c; this.sucia = true; }
  _esTexto(r, c) { return this.textoCols.some((x) => c >= x.c1 && (!x.c2 || c <= x.c2) && r >= x.r1 && (!x.r2 || r <= x.r2)); }
  _escribe(r, c, v) {
    if (v instanceof Date) v = new Date(v.getTime());
    else if (v === null || v === undefined) v = '';
    else if (this._esTexto(r, c)) v = typeof v === 'number' ? String(v) : v;       // en una celda de «texto sin formato» nada se interpreta
    else v = interpreta(v);
    this._ponCruda(r, c, v);
  }
  _borra(r, c) { if (this.libro.soloLectura) { if (this._lee(r, c) !== '') this.libro._intento(r > 1 && this.libro.hojas.includes(this)); return; } const f = this.filas[r - 1]; if (f && c - 1 < f.length) { f[c - 1] = ''; this.sucia = true; } }
  _formato(r, c, nr, nc, f) {
    if (this.libro.soloLectura) return;
    const x = { r1: r, c1: c, r2: nr ? r + nr - 1 : 0, c2: nc ? c + nc - 1 : 0 };
    if (f === '@') this.textoCols.push(x);
    else this.textoCols = this.textoCols.filter((t) => !(t.c1 === x.c1 && t.c2 === x.c2 && t.r1 === x.r1 && t.r2 === x.r2));
  }
  getName() { return this.nombre; } setName(n) { this.nombre = String(n); this.libro._renombra(this); return this; }
  getSheetId() { return this.libro.hojas.indexOf(this) + 1; } getParent() { return this.libro; }
  getLastRow() { for (let i = this.filas.length - 1; i >= 0; i--) if (this.filas[i].some((v) => v !== '' && v !== null && v !== undefined)) return i + 1; return 0; }
  getLastColumn() { let m = 0; for (const f of this.filas) for (let j = f.length - 1; j >= m; j--) if (f[j] !== '' && f[j] !== undefined && f[j] !== null) { m = j + 1; break; } return m; }
  getMaxColumns() { return this.maxCols; } getMaxRows() { return this.maxFilas; }
  getRange(a, b, c, d) {
    if (typeof a === 'string') { const p = parseA1(a); const r2 = p.r2 || this.maxFilas, c2 = p.c2 || this.maxCols; return new Rango(this, p.r1, p.c1, r2 - p.r1 + 1, c2 - p.c1 + 1); }
    return new Rango(this, a, b, c || 1, d || 1);
  }
  getDataRange() { return new Rango(this, 1, 1, Math.max(1, this.getLastRow()), Math.max(1, this.getLastColumn())); }
  appendRow(fila) { const r = this.getLastRow() + 1; fila.forEach((v, j) => this._escribe(r, j + 1, v)); return this; }
  _bloqueada() { if (this.libro.soloLectura) { this.libro._intento(this.libro.hojas.includes(this)); return true; } return false; }
  deleteRow(n) { if (this._bloqueada()) return this; if (n <= this.filas.length) this.filas.splice(n - 1, 1); this.sucia = true; return this; }
  deleteRows(n, k) { if (this._bloqueada()) return this; this.filas.splice(n - 1, k); this.sucia = true; return this; }
  insertColumnAfter(n) { if (this._bloqueada()) return this; this.maxCols++; this.filas.forEach((f) => { if (f.length > n) f.splice(n, 0, ''); }); this.sucia = true; return this; }
  insertColumnBefore(n) { return this.insertColumnAfter(n - 1); }
  insertRowAfter(n) { if (this._bloqueada()) return this; this.filas.splice(n, 0, []); this.sucia = true; return this; }
  insertRowsAfter(n, k) { for (let i = 0; i < k; i++) this.insertRowAfter(n); return this; }
  setColumnWidth() { return this; } setRowHeight() { return this; } setFrozenRows() { return this; } setFrozenColumns() { return this; } setTabColor() { return this; }
  hideSheet() { return this; } showSheet() { return this; } setHiddenGridlines() { return this; } autoResizeColumn() { return this; } setConditionalFormatRules() { return this; }
  getProtections() { return []; } protect() { return { setWarningOnly() { return this; }, addEditor() { return this; }, removeEditors() { return this; }, setDescription() { return this; } }; }
  clear() { if (this._bloqueada()) return this; this.filas = []; this.sucia = true; return this; } clearContents() { return this.clear(); }
  activate() { return this; }
}

export class Libro {
  constructor(id = 'libro-local') { this.id = id; this.hojas = []; this.nombre = 'Caja Rey Pizza'; this.soloLectura = false; this.huboEscritura = false; this.escrituraDatos = false; }
  /** Un intento de cambiar algo en modo lectura. «Reparar» (crear una pestaña que falta o su fila de títulos) no es lo mismo que escribir datos. */
  _intento(datos) { this.huboEscritura = true; if (datos) this.escrituraDatos = true; }
  /** Cambia una pestaña entera por otra (así se pone al día la copia). */
  reemplazaHoja(nombre, hoja) { const i = this.hojas.findIndex((h) => h.nombre === nombre); hoja.libro = this; if (i >= 0) this.hojas[i] = hoja; else this.hojas.push(hoja); }
  quitaHoja(nombre) { const i = this.hojas.findIndex((h) => h.nombre === nombre); if (i >= 0) this.hojas.splice(i, 1); }
  _renombra() {}
  getId() { return this.id; } getName() { return this.nombre; } getUrl() { return 'https://docs.google.com/spreadsheets/d/' + this.id + '/edit'; }
  getSheetByName(n) { return this.hojas.find((h) => h.nombre === n) || null; }
  getSheets() { return this.hojas.slice(); }
  insertSheet(n) { if (this.soloLectura) { this._intento(false); return new Hoja(this, String(n === undefined ? 'temporal' : n)); } const nombre = n === undefined ? 'Hoja ' + (this.hojas.length + 1) : String(n); if (this.getSheetByName(nombre)) throw new Error(`Ya existe una hoja llamada «${nombre}».`); const h = new Hoja(this, nombre); this.hojas.push(h); return h; }
  deleteSheet(h) { if (this.soloLectura) { this._intento(true); return; } const i = this.hojas.indexOf(h); if (i >= 0) this.hojas.splice(i, 1); }
  setActiveSheet(h) { return h; } moveActiveSheet() {} getActiveSheet() { return this.hojas[0]; } setSpreadsheetTimeZone() {} setName(n) { this.nombre = n; return this; } flush() {} getSpreadsheetTimeZone() { return 'America/Guatemala'; }
  duplicateActiveSheet() { throw new Error('No disponible'); }
  /** Para guardar y recuperar todo el libro (las fechas se marcan para no perder su tipo). */
  exporta() {
    const cod = (v) => (v instanceof Date ? { $d: v.getTime() } : v);
    return { id: this.id, hojas: this.hojas.map((h) => ({ nombre: h.nombre, filas: h.filas.map((f) => f.map(cod)), textoCols: h.textoCols, maxFilas: h.maxFilas, maxCols: h.maxCols, formulas: [...h.formulas] })) };
  }
  static importa(d) {
    const l = new Libro(d.id);
    for (const x of d.hojas) { const h = new Hoja(l, x.nombre); h.filas = x.filas.map((f) => f.map((v) => (v && typeof v === 'object' && '$d' in v ? new Date(v.$d) : v))); h.textoCols = x.textoCols || []; h.maxFilas = x.maxFilas; h.maxCols = x.maxCols; h.formulas = new Map(x.formulas || []); h.sucia = false; l.hojas.push(h); }
    return l;
  }
}
