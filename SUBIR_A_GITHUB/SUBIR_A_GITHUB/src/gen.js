// Generador del "REPORTE DIARIO - MANTENIMIENTO ELECTRICO EN CAMPAMENTOS" (.xlsx)
// Sin dependencias salvo JSZip. Funciona en navegador y en Node.
(function (root) {
  'use strict';

  // ---------- utilidades ----------
  const COLS = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';
  const colIdx = (c) => COLS.indexOf(c);           // 'A' -> 0
  const colLetter = (i) => COLS[i];
  const esc = (s) => String(s == null ? '' : s)
    .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '')
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

  const MESES = ['ENERO', 'FEBRERO', 'MARZO', 'ABRIL', 'MAYO', 'JUNIO', 'JULIO', 'AGOSTO', 'SETIEMBRE', 'OCTUBRE', 'NOVIEMBRE', 'DICIEMBRE'];

  // fecha 'YYYY-MM-DD' -> número de serie de Excel
  function dateSerial(iso) {
    const [y, m, d] = iso.split('-').map(Number);
    return Math.round((Date.UTC(y, m - 1, d) - Date.UTC(1899, 11, 30)) / 86400000);
  }
  function addDays(iso, n) {
    const [y, m, d] = iso.split('-').map(Number);
    const t = new Date(Date.UTC(y, m - 1, d + n));
    return t.toISOString().slice(0, 10);
  }
  const hm = (min) => `${Math.floor(min / 60)}:${String(min % 60).padStart(2, '0')}`;

  // ---------- reparto de horas ----------
  // items: [{dur: horas opcional}]  -> [{ini, fin}] en minutos
  function distribuir(items, ini, fin, paso) {
    paso = paso || 15;
    const n = items.length;
    if (!n) return [];
    const total = fin - ini;
    const fijos = items.map((a) => {
      const h = parseFloat(String(a.dur == null ? '' : a.dur).replace(',', '.'));
      return isFinite(h) && h > 0 ? Math.round(h * 60) : null;
    });
    const usado = fijos.reduce((s, v) => s + (v || 0), 0);
    const libres = fijos.filter((v) => v == null).length;
    const resto = Math.max(0, total - usado);
    const base = libres ? Math.max(paso, Math.round(resto / libres / paso) * paso) : 0;
    const out = [];
    let t = ini;
    for (let i = 0; i < n; i++) {
      let d = fijos[i] != null ? fijos[i] : base;
      let f = (i === n - 1) ? fin : Math.min(fin, t + d);
      if (f < t) f = t;
      out.push({ ini: t, fin: f });
      t = f;
    }
    return out;
  }

  // ---------- estilos ----------
  const AZUL = 'FFB4C6E7';
  function Styles() {
    this.fonts = ['<font><sz val="11"/><name val="Calibri"/><family val="2"/></font>'];
    this.fills = ['<fill><patternFill patternType="none"/></fill>', '<fill><patternFill patternType="gray125"/></fill>'];
    this.borders = ['<border><left/><right/><top/><bottom/><diagonal/></border>'];
    this.numFmts = [];
    this.xfs = ['<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'];
    this.map = new Map();
    this.idx = { fonts: new Map(), fills: new Map(), borders: new Map(), fmts: new Map() };
  }
  Styles.prototype._get = function (kind, xml, arr) {
    const m = this.idx[kind];
    if (m.has(xml)) return m.get(xml);
    arr.push(xml); m.set(xml, arr.length - 1); return arr.length - 1;
  };
  Styles.prototype.id = function (s) {
    s = s || {};
    const key = JSON.stringify(s);
    if (this.map.has(key)) return this.map.get(key);
    const f = s.font || {};
    const fontXml = `<font>${f.b ? '<b/>' : ''}${f.i ? '<i/>' : ''}<sz val="${f.sz || 10}"/>${f.color ? `<color rgb="${f.color}"/>` : '<color theme="1"/>'}<name val="${f.name || 'Arial'}"/><family val="2"/></font>`;
    const fontId = this._get('fonts', fontXml, this.fonts);
    let fillId = 0;
    if (s.fill) fillId = this._get('fills', `<fill><patternFill patternType="solid"><fgColor rgb="${s.fill}"/><bgColor indexed="64"/></patternFill></fill>`, this.fills);
    let borderId = 0;
    if (s.border) {
      const b = s.border;
      const side = (n, v) => v ? `<${n} style="${v}"><color indexed="64"/></${n}>` : `<${n}/>`;
      borderId = this._get('borders', `<border>${side('left', b.l)}${side('right', b.r)}${side('top', b.t)}${side('bottom', b.b)}<diagonal/></border>`, this.borders);
    }
    let numFmtId = 0;
    if (s.fmt != null) {
      const builtin = { '0%': 9, '0.00': 2, 'h:mm': 20, '0': 1, '@': 49 };
      if (builtin[s.fmt] != null) numFmtId = builtin[s.fmt];
      else {
        const m = this.idx.fmts;
        if (!m.has(s.fmt)) { const id = 164 + this.numFmts.length; this.numFmts.push(`<numFmt numFmtId="${id}" formatCode="${esc(s.fmt)}"/>`); m.set(s.fmt, id); }
        numFmtId = m.get(s.fmt);
      }
    }
    const a = s.align || {};
    const ind = (a.h === 'left' || a.h === 'right' || a.h === 'justify') ? ' indent="1"' : '';
    const alignXml = (a.h || a.v || a.wrap) ? `<alignment${a.h ? ` horizontal="${a.h}"` : ''}${a.v ? ` vertical="${a.v}"` : ''}${a.wrap ? ' wrapText="1"' : ''}${ind}/>` : '';
    const xf = `<xf numFmtId="${numFmtId}" fontId="${fontId}" fillId="${fillId}" borderId="${borderId}" xfId="0"${numFmtId ? ' applyNumberFormat="1"' : ''} applyFont="1"${fillId ? ' applyFill="1"' : ''}${borderId ? ' applyBorder="1"' : ''}${alignXml ? ' applyAlignment="1">' + alignXml + '</xf>' : '/>'}`;
    this.xfs.push(xf);
    const id = this.xfs.length - 1;
    this.map.set(key, id);
    return id;
  };
  Styles.prototype.xml = function () {
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
      '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
      (this.numFmts.length ? `<numFmts count="${this.numFmts.length}">${this.numFmts.join('')}</numFmts>` : '') +
      `<fonts count="${this.fonts.length}">${this.fonts.join('')}</fonts>` +
      `<fills count="${this.fills.length}">${this.fills.join('')}</fills>` +
      `<borders count="${this.borders.length}">${this.borders.join('')}</borders>` +
      '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
      `<cellXfs count="${this.xfs.length}">${this.xfs.join('')}</cellXfs>` +
      '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>' +
      '</styleSheet>';
  };

  // ---------- hoja ----------
  // Columnas del formato original (D..P). A-C angostas como margen.
  const WIDTHS = { A: 1.5, B: 1.5, C: 1.5, D: 13, E: 23.78, F: 9.33, G: 7.44, H: 8.89, I: 12, J: 21.55, K: 12.11, L: 16.22, M: 14.44, N: 10.66, O: 13, P: 30.66 };
  const colPx = (c) => Math.round((WIDTHS[c] || 8.43) * 7 + 5);
  const PX = 9525; // EMU por pixel

  function Sheet(name) {
    this.name = name;
    this.cells = {};   // ref -> {v, f, t, s}
    this.merges = [];
    this.heights = {};
    this.images = [];  // {data, ext, col, colOff, row, rowOff, cx, cy}
    this.blocks = [];  // [ini, fin] filas que no deben partirse entre páginas
    this.breaks = [];
    this.r = 1;
  }
  Sheet.prototype.set = function (ref, v, s, f) {
    this.cells[ref] = { v: v, s: Object.assign({}, s || {}), f: f };
  };
  Sheet.prototype.style = function (ref, s) {
    const c = this.cells[ref] || (this.cells[ref] = { v: null, s: {} });
    c.s = Object.assign({}, c.s, s);
  };
  // rango 'D5:P5' con valor y estilo; aplica estilo a todas las celdas
  Sheet.prototype.box = function (c1, r1, c2, r2, v, s, f) {
    const a = colIdx(c1), b = colIdx(c2);
    for (let r = r1; r <= r2; r++) for (let c = a; c <= b; c++) {
      const ref = colLetter(c) + r;
      if (r === r1 && c === a) this.set(ref, v, s, f); else this.set(ref, null, s);
    }
    if (a !== b || r1 !== r2) this.merges.push(`${c1}${r1}:${c2}${r2}`);
  };
  Sheet.prototype.h = function (r, pt) { this.heights[r] = pt; };

  // posición en px (x desde inicio de columna D) -> col/offset
  function locateX(xPx) {
    let c = colIdx('D');
    let x = xPx;
    while (c < colIdx('P') && x >= colPx(colLetter(c))) { x -= colPx(colLetter(c)); c++; }
    return { col: c, off: Math.max(0, Math.round(x)) };
  }
  const AREA_PX = (() => { let s = 0; for (let c = colIdx('D'); c <= colIdx('P'); c++) s += colPx(colLetter(c)); return s; })();

  Sheet.prototype.image = function (img, xPx, row, yPx, wPx, hPx) {
    const loc = locateX(xPx);
    this.images.push({ data: img.data, ext: img.ext, col: loc.col, colOff: loc.off * PX, row: row - 1, rowOff: Math.round(yPx) * PX, cx: Math.round(wPx) * PX, cy: Math.round(hPx) * PX });
  };

  // Saltos de página: A4 vertical ajustado a 1 página de ancho
  Sheet.prototype.calcBreaks = function () {
    const anchoPt = AREA_PX * 0.75;
    const escala = Math.min(1, (8.27 - 0.6) * 72 / anchoPt);
    const cap = ((11.69 - 0.8) * 72) / escala - 40; // margen de seguridad
    const alto = (r) => (this.heights[r] != null ? this.heights[r] : 14.4);
    let usado = 0, fila = 2;
    const avanzar = (hasta) => { for (; fila <= hasta; fila++) { const h = alto(fila); if (usado + h > cap) { this.breaks.push(fila); usado = 0; } usado += h; } };
    this.blocks.forEach(([a, b]) => {
      avanzar(a - 1);
      let hb = 0; for (let k = a; k <= b; k++) hb += alto(k);
      if (usado + hb > cap && hb <= cap && usado > 0) { this.breaks.push(a); usado = 0; }
      avanzar(b);
    });
    avanzar(this.r - 1);
  };

  Sheet.prototype.xml = function (styles) {
    const rows = {};
    Object.keys(this.cells).forEach((ref) => {
      const m = ref.match(/^([A-Z]+)(\d+)$/);
      (rows[+m[2]] = rows[+m[2]] || []).push([colIdx(m[1]), ref, this.cells[ref]]);
    });
    const allRows = new Set(Object.keys(rows).map(Number).concat(Object.keys(this.heights).map(Number)));
    const sorted = Array.from(allRows).sort((a, b) => a - b);
    let sd = '';
    sorted.forEach((r) => {
      const ht = this.heights[r];
      sd += `<row r="${r}"${ht ? ` ht="${ht}" customHeight="1"` : ''}>`;
      (rows[r] || []).sort((a, b) => a[0] - b[0]).forEach(([, ref, c]) => {
        const s = styles.id(c.s);
        if (c.f) sd += `<c r="${ref}" s="${s}"><f>${esc(c.f)}</f>${c.v != null ? `<v>${c.v}</v>` : ''}</c>`;
        else if (c.v == null || c.v === '') sd += `<c r="${ref}" s="${s}"/>`;
        else if (typeof c.v === 'number') sd += `<c r="${ref}" s="${s}"><v>${c.v}</v></c>`;
        else if (typeof c.v === 'object' && c.v.rich) {
          const f = (c.s && c.s.font) || {};
          const runs = c.v.rich.map((run) => `<r><rPr>${run.b || f.b ? '<b/>' : ''}<sz val="${f.sz || 10}"/><color theme="1"/><rFont val="${f.name || 'Arial'}"/><family val="2"/></rPr><t xml:space="preserve">${esc(run.t)}</t></r>`).join('');
          sd += `<c r="${ref}" s="${s}" t="inlineStr"><is>${runs}</is></c>`;
        }
        else sd += `<c r="${ref}" s="${s}" t="inlineStr"><is><t xml:space="preserve">${esc(c.v)}</t></is></c>`;
      });
      sd += '</row>';
    });
    let cols = '<cols>';
    Object.keys(WIDTHS).forEach((c) => { const i = colIdx(c) + 1; cols += `<col min="${i}" max="${i}" width="${WIDTHS[c]}" customWidth="1"/>`; });
    cols += '</cols>';
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
      '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
      '<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>' +
      `<dimension ref="A1:P${this.r}"/>` +
      '<sheetViews><sheetView showGridLines="0" zoomScale="70" zoomScaleNormal="70" workbookViewId="0"/></sheetViews>' +
      '<sheetFormatPr defaultRowHeight="14.4"/>' + cols +
      `<sheetData>${sd}</sheetData>` +
      (this.merges.length ? `<mergeCells count="${this.merges.length}">${this.merges.map((m) => `<mergeCell ref="${m}"/>`).join('')}</mergeCells>` : '') +
      '<printOptions horizontalCentered="1"/>' +
      '<pageMargins left="0.3" right="0.3" top="0.4" bottom="0.4" header="0.2" footer="0.2"/>' +
      '<pageSetup paperSize="9" orientation="portrait" fitToWidth="1" fitToHeight="0"/>' +
      (this.breaks.length ? `<rowBreaks count="${this.breaks.length}" manualBreakCount="${this.breaks.length}">${this.breaks.map((b) => `<brk id="${b - 1}" max="16383" man="1"/>`).join('')}</rowBreaks>` : '') +
      (this.images.length ? '<drawing r:id="rId1"/>' : '') +
      '</worksheet>';
  };

  // ---------- construcción del informe ----------
  const T = { thin: 'thin', med: 'medium' };
  const ALL = { l: 'thin', r: 'thin', t: 'thin', b: 'thin' };
  const S = {
    title: { font: { b: true, sz: 18 }, fill: AZUL, border: { l: 'medium', r: 'medium', t: 'medium', b: 'medium' }, align: { h: 'center', v: 'center', wrap: true } },
    head: { font: { b: true, sz: 11 }, fill: AZUL, border: ALL, align: { h: 'center', v: 'center' } },
    headR: { font: { b: true, sz: 11 }, fill: AZUL, border: ALL, align: { h: 'right', v: 'center' } },
    headDate: { font: { b: true, sz: 10 }, fill: AZUL, border: ALL, align: { h: 'center', v: 'center' }, fmt: '[$-280A]d" de "mmmm" de "yyyy;@' },
    colh: { font: { b: true, sz: 10, color: 'FF2F5597' }, border: ALL, align: { h: 'center', v: 'center', wrap: true } },
    label: { font: { b: true, sz: 11 }, border: ALL, align: { h: 'left', v: 'center' } },
    labelR: { font: { b: true, sz: 11 }, border: ALL, align: { h: 'right', v: 'center' } },
    val: { font: { sz: 11 }, border: ALL, align: { h: 'center', v: 'center', wrap: true } },
    valB: { font: { b: true, sz: 11 }, border: ALL, align: { h: 'center', v: 'center', wrap: true } },
    item: { font: { b: true, sz: 9 }, border: ALL, align: { h: 'center', v: 'center' } },
    cell: { font: { b: true, sz: 9 }, border: ALL, align: { h: 'center', v: 'center', wrap: true } },
    cellN: { font: { sz: 9 }, border: ALL, align: { h: 'center', v: 'center', wrap: true } },
    time: { font: { b: true, sz: 9 }, border: ALL, align: { h: 'center', v: 'center' }, fmt: 'h:mm' },
    pct: { font: { b: true, sz: 9 }, border: ALL, align: { h: 'center', v: 'center' }, fmt: '0%' },
    num: { font: { sz: 9 }, border: ALL, align: { h: 'center', v: 'center' }, fmt: '0.00' },
    ptitle18: { font: { b: true, sz: 18 }, border: ALL, align: { h: 'center', v: 'center', wrap: true } },
    ptitle16: { font: { b: true, sz: 16 }, border: ALL, align: { h: 'center', v: 'center', wrap: true } },
    ptext: { font: { sz: 13 }, border: ALL, align: { h: 'justify', v: 'top', wrap: true } },
    pbox: { border: { l: 'thin', r: 'thin' } },
    firmaH: { font: { b: true, sz: 9 }, fill: AZUL, border: { l: 'medium', r: 'medium', t: 'medium', b: 'medium' }, align: { h: 'center', v: 'center' } },
    firmaB: { font: { sz: 9 }, border: { l: 'medium', r: 'medium', t: 'medium', b: 'medium' }, align: { h: 'center', v: 'bottom' } },
  };

  // Convierte **negrita** en tramos de texto enriquecido
  function richText(txt) {
    const partes = String(txt).split(/(\*\*[^*]+\*\*)/g).filter((x) => x !== '');
    if (!partes.some((x) => /^\*\*[^*]+\*\*$/.test(x))) return txt;
    return { rich: partes.map((x) => /^\*\*[^*]+\*\*$/.test(x) ? { t: x.slice(2, -2), b: true } : { t: x }) };
  }
  // Líneas que ocupa un texto con letra de 'sz' pt en el ancho del panel (D:P)
  function lineasPanel(texto, sz) {
    const anchoPt = AREA_PX * 0.75 - 24;               // ancho útil menos márgenes
    const porLinea = Math.floor(anchoPt / (sz * 0.53)); // ancho medio de carácter en Arial (conservador)
    return String(texto || '').replace(/\*\*/g, '').split(/\n/).reduce((s, p) => s + Math.max(1, Math.ceil(p.length / porLinea)), 0);
  }

  function estimaAlto(texto, ptPorLinea, charsPorLinea, min) {
    const lineas = String(texto || '').split(/\n/).reduce((s, p) => s + Math.max(1, Math.ceil(p.length / charsPorLinea)), 0);
    return Math.max(min, Math.round(lineas * ptPorLinea + 8));
  }

  function construirHoja(sh, d, fecha, assets) {
    let r = 2;
    const fechaSerial = dateSerial(fecha);
    const mes = (d.mes && String(d.mes).trim()) || MESES[+fecha.slice(5, 7) - 1];

    // Encabezado
    for (let i = 0; i < 4; i++) sh.h(r + i, 18);
    sh.box('D', r, 'E', r + 3, null, { border: { l: 'medium', r: 'medium', t: 'medium', b: 'medium' } });
    sh.box('F', r, 'O', r + 3, 'REPORTE DIARIO\nMANTENIMIENTO ELECTRICO EN CAMPAMENTOS', S.title);
    sh.box('P', r, 'P', r + 2, 'Página', Object.assign({}, S.title, { font: { b: true, sz: 10 }, border: { l: 'medium', r: 'medium', t: 'medium' } }));
    sh.box('P', r + 3, 'P', r + 3, 1, Object.assign({}, S.title, { font: { b: true, sz: 10 }, border: { l: 'medium', r: 'medium', b: 'medium' } }));
    if (assets.logo) {
      const w = colPx('D') + colPx('E'), h = 4 * 18 * 4 / 3;
      const k = Math.min((w - 12) / assets.logo.w, (h - 10) / assets.logo.h);
      const iw = assets.logo.w * k, ih = assets.logo.h * k;
      sh.image(assets.logo, (w - iw) / 2, r, (h - ih) / 2, iw, ih);
    }
    r += 4; sh.h(r, 6); r++;

    // Datos generales
    const top = r;
    sh.box('D', r, 'P', r, 'DATOS GENERALES', S.head); r++;
    sh.box('D', r, 'E', r, 'LUGAR DE TRABAJO:', S.label);
    sh.box('F', r, 'J', r, d.lugar || '', S.val);
    sh.box('K', r, 'L', r, 'AREA DE TRABAJO:', S.labelR);
    sh.box('M', r, 'P', r, d.area || '', S.val); r++;
    sh.box('D', r, 'E', r, 'ADMINISTRADOR CONTRATISTA:', S.label);
    sh.box('F', r, 'J', r, d.administrador || '', S.val);
    sh.box('K', r, 'L', r, 'FRENTE DE TRABAJO:', S.labelR);
    sh.box('M', r, 'P', r, d.frente || '', S.val); r++;
    const tecs = (d.tecnicos || []).filter((t) => t && String(t.nombre || '').trim());
    const nT = Math.max(1, tecs.length);
    sh.box('D', r, 'E', r + nT - 1, 'TECNICO(S) EJECUTOR(ES):', Object.assign({}, S.label, { align: { h: 'left', v: 'center', wrap: true } }));
    for (let i = 0; i < nT; i++) sh.box('F', r + i, 'J', r + i, tecs[i] ? String(tecs[i].nombre).toUpperCase() : '', S.valB);
    sh.box('K', r, 'L', r, 'MES:', S.labelR);
    sh.box('M', r, 'P', r, mes, S.valB);
    if (nT > 1) sh.box('K', r + 1, 'P', r + nT - 1, null, S.val);
    r += nT;
    sh.h(r, 6); r++;

    // Actividades ejecutadas
    sh.box('D', r, 'L', r, 'ACTIVIDADES EJECUTADAS', S.head);
    sh.box('M', r, 'M', r, 'FECHA:', S.headR);
    sh.box('N', r, 'P', r, fechaSerial, S.headDate); r++;
    sh.h(r, 26);
    sh.box('D', r, 'E', r, 'ITEM', S.colh);
    sh.box('F', r, 'K', r, 'PARTIDAS EJECUTADAS', S.colh);
    sh.box('L', r, 'L', r, 'HORA DE INICIO', S.colh);
    sh.box('M', r, 'M', r, 'HORA FIN', S.colh);
    sh.box('N', r, 'N', r, '% AVANCE', S.colh);
    sh.box('O', r, 'P', r, 'OBSERVACIONES', S.colh); r++;
    const filas = calcularActividades(d);
    filas.forEach((a, i) => {
      const alto = Math.max(estimaAlto(a.desc, 12, 60, 17), estimaAlto(a.obs, 12, 44, 17));
      sh.h(r, alto);
      sh.box('D', r, 'E', r, 'ACTIVIDAD ' + (i + 1), S.item);
      sh.box('F', r, 'K', r, String(a.desc || '').toUpperCase(), S.cell);
      sh.box('L', r, 'L', r, a.ini / 1440, S.time);
      sh.box('M', r, 'M', r, a.fin / 1440, S.time);
      sh.box('N', r, 'N', r, a.avance == null ? 1 : a.avance, S.pct);
      sh.box('O', r, 'P', r, String(a.obs || '').toUpperCase(), S.cell);
      r++;
    });
    sh.h(r, 6); r++;

    // Panel fotográfico
    const r0Panel = r;
    sh.box('D', r, 'P', r, 'PANEL FOTOGRAFICO', S.head); r++;
    (d.panel || []).forEach((p, i) => {
      const b0 = r;
      const tit = `ACTIVIDAD ${i + 1}: ${String(p.titulo || '').toUpperCase()}`;
      const tsz = tit.length <= 60 ? 18 : 16;             // título corto = 18, largo = 16
      sh.h(r, Math.min(409, lineasPanel(tit, tsz * 1.12) * tsz * 1.35 + 14));
      sh.box('D', r, 'P', r, tit, tsz === 18 ? S.ptitle18 : S.ptitle16); r++;
      const fotos = (p.fotos || []).filter(Boolean);
      const n = fotos.length;
      if (n) {
        const porFila = Math.max(1, Math.min(n, +p.porFila || (n <= 4 ? n : (n <= 6 ? 3 : 4))));
        const altoPt = +p.altoFila || 240;
        const altoPx = altoPt * 4 / 3;
        const slotW = AREA_PX / porFila;
        for (let k = 0; k < n; k += porFila) {
          sh.h(r, altoPt);
          sh.box('D', r, 'P', r, null, S.pbox);
          const fila = fotos.slice(k, k + porFila);
          const ancho = slotW * fila.length, x0 = (AREA_PX - ancho) / 2; // centra filas incompletas
          fila.forEach((f, j) => {
            const m = 6;
            const sc = Math.min((slotW - 2 * m) / f.w, (altoPx - 2 * m) / f.h);
            const iw = f.w * sc, ih = f.h * sc;
            sh.image(f, x0 + j * slotW + (slotW - iw) / 2, r, (altoPx - ih) / 2, iw, ih);
          });
          r++;
        }
      }
      // Redacción: alto según el largo del texto. Excel limita cada fila a 409 pt,
      // así que si hace falta más el recuadro se extiende sobre varias filas combinadas.
      const txt = String(p.texto || '').trim();
      const szT = S.ptext.font.sz;
      const altoTxt = Math.max(70, lineasPanel(txt, szT) * szT * 1.3 + 24);
      const nFil = Math.ceil(altoTxt / 400);
      for (let k = 0; k < nFil; k++) sh.h(r + k, Math.ceil(altoTxt / nFil));
      sh.box('D', r, 'P', r + nFil - 1, txt ? richText(txt) : '', S.ptext); r += nFil;
      sh.blocks.push([b0 === r0Panel + 1 ? r0Panel : b0, r - 1]);
    });
    sh.h(r, 6); r++;

    // Materiales
    const bMat = r;
    sh.box('D', r, 'P', r, 'MATERIALES Y EQUIPOS UTILIZADOS', S.head); r++;
    sh.box('D', r, 'D', r, 'ITEM', S.head);
    sh.box('E', r, 'N', r, 'DESCRIPCION', S.head);
    sh.box('O', r, 'O', r, 'UNIDAD', S.head);
    sh.box('P', r, 'P', r, 'CANTIDAD', S.head); r++;
    const mats = [{ desc: 'HERRAMIENTAS DE PODER', unidad: 'UNIDAD', cant: '-' }, { desc: 'HERRAMIENTAS MANUALES', unidad: 'UNIDAD', cant: '-' }]
      .concat((d.materiales || []).filter((m) => m && String(m.desc || '').trim()));
    mats.forEach((m, i) => {
      sh.h(r, 17);
      sh.box('D', r, 'D', r, i + 1, S.cellN);
      sh.box('E', r, 'N', r, String(m.desc).toUpperCase(), S.cellN);
      sh.box('O', r, 'O', r, String(m.unidad || 'UNIDAD').toUpperCase(), S.cellN);
      const q = parseFloat(String(m.cant).replace(',', '.'));
      sh.box('P', r, 'P', r, (m.cant === '-' || m.cant === '' || m.cant == null || !isFinite(q)) ? (m.cant || '-') : q, S.cellN);
      r++;
    });
    sh.h(r, 6); r++;

    sh.blocks.push([bMat, r - 1]);
    // Horas hombre
    const bHH = r;
    sh.box('D', r, 'P', r, 'HORAS HOMBRE UTILIZADAS', S.head); r++;
    sh.h(r, 28);
    sh.box('D', r, 'D', r, 'N°', S.head);
    sh.box('E', r, 'H', r, 'CARGO', S.head);
    sh.box('I', r, 'I', r, 'CANT.', S.head);
    sh.box('J', r, 'J', r, 'UND', S.head);
    sh.box('K', r, 'K', r, 'CANT. PARCIAL', Object.assign({}, S.head, { align: { h: 'center', v: 'center', wrap: true } }));
    sh.box('L', r, 'L', r, 'CANTIDAD TOTAL', Object.assign({}, S.head, { align: { h: 'center', v: 'center', wrap: true } }));
    sh.box('M', r, 'P', r, 'OBSERVACIONES', S.head); r++;
    const hh = +d.hhPorTecnico || 10;
    const r0 = r;
    (tecs.length ? tecs : [{ nombre: '', cargo: 'TÉCNICO ELECTRICISTA' }]).forEach((t, i) => {
      sh.h(r, 16);
      sh.box('D', r, 'D', r, String(i + 1).padStart(2, '0'), S.cellN);
      sh.box('E', r, 'H', r, String(t.cargo || 'TÉCNICO ELECTRICISTA').toUpperCase(), S.cellN);
      sh.box('I', r, 'I', r, 1, S.num);
      sh.box('J', r, 'J', r, 'HH', S.cellN);
      sh.box('K', r, 'K', r, hh, S.num);
      sh.box('L', r, 'L', r, hh, S.num, `K${r}*I${r}`);
      sh.box('M', r, 'P', r, '', S.cellN);
      r++;
    });
    const nTt = r - r0;
    sh.h(r, 16);
    sh.box('D', r, 'D', r, '', S.cellN);
    sh.box('E', r, 'H', r, 'TOTAL', S.cell);
    sh.box('I', r, 'I', r, nTt, Object.assign({}, S.num, { font: { b: true, sz: 9 } }), `SUM(I${r0}:I${r - 1})`);
    sh.box('J', r, 'J', r, 'HH', S.cellN);
    sh.box('K', r, 'K', r, '', S.cellN);
    sh.box('L', r, 'L', r, nTt * hh, Object.assign({}, S.num, { font: { b: true, sz: 9 } }), `SUM(L${r0}:L${r - 1})`);
    sh.box('M', r, 'P', r, '', S.cellN); r++;

    // Checklist
    sh.box('D', r, 'P', r, 'CHECKLIST', S.head); r++;
    sh.h(r, 17);
    sh.box('D', r, 'I', r, 'HERRAMIENTAS MANUALES Y DE PODER', Object.assign({}, S.cellN, { align: { h: 'left', v: 'center' } }));
    sh.box('J', r, 'J', r, 'X', S.cell);
    sh.box('K', r, 'P', r, d.checklist || 'TODO OPERATIVO SIN OBSERVACIONES DE LAS HERRAMIENTAS', Object.assign({}, S.cellN, { align: { h: 'left', v: 'center' } })); r++;
    sh.h(r, 6); r++;

    sh.blocks.push([bHH, r - 1]);
    // Actividades programadas
    const bProg = r;
    sh.box('D', r, 'L', r, 'ACTIVIDADES PROGRAMADAS', S.head);
    sh.box('M', r, 'M', r, 'FECHA:', S.headR);
    sh.box('N', r, 'P', r, dateSerial(addDays(fecha, 1)), S.headDate); r++;
    sh.box('D', r, 'E', r, 'ITEM', S.colh);
    sh.box('F', r, 'L', r, 'PARTIDAS PROGRAMADAS', S.colh);
    sh.box('M', r, 'P', r, 'OBSERVACIONES', S.colh); r++;
    const prog = (d.programadas || []).filter((p) => p && String(p.desc || '').trim());
    const nP = Math.max(3, prog.length);
    for (let i = 0; i < nP; i++) {
      const p = prog[i] || {};
      sh.h(r, Math.max(estimaAlto(p.desc, 12, 70, 16), estimaAlto(p.obs, 12, 60, 16)));
      sh.box('D', r, 'E', r, 'ACTIVIDAD ' + (i + 1), S.item);
      sh.box('F', r, 'L', r, String(p.desc || '').toUpperCase(), S.cell);
      sh.box('M', r, 'P', r, String(p.obs || '').toUpperCase(), S.cell);
      r++;
    }
    const bottom = r - 1;
    sh.h(r, 10); r++;

    // Firmas
    const bFir = r;
    sh.box('D', r, 'J', r, d.firmaIzq || 'MINERA YANAQUIHUA', S.firmaH);
    sh.box('K', r, 'P', r, d.firmaDer || 'CONTRATISTA', S.firmaH); r++;
    for (let i = 0; i < 4; i++) sh.h(r + i, i === 3 ? 30 : 22);
    sh.box('D', r, 'J', r + 3, 'Sello y firma Supervisor', S.firmaB);
    sh.box('K', r, 'P', r + 3, 'Sello y firma Administrador', S.firmaB);
    if (assets.firma) {
      const w = ['K', 'L', 'M', 'N', 'O', 'P'].reduce((s, c) => s + colPx(c), 0);
      const h = (22 * 3 + 16) * 4 / 3;
      const k = Math.min((w * 0.5) / assets.firma.w, (h - 8) / assets.firma.h);
      const iw = assets.firma.w * k, ih = assets.firma.h * k;
      const xK = ['D', 'E', 'F', 'G', 'H', 'I', 'J'].reduce((s, c) => s + colPx(c), 0);
      sh.image(assets.firma, xK + (w - iw) / 2, r, 4, iw, ih);
    }
    r += 4;
    sh.blocks.push([bProg, r - 1]);

    // Marco exterior grueso
    for (let k = top; k <= bottom; k++) {
      if (sh.cells['D' + k]) sh.style('D' + k, { border: Object.assign({}, (sh.cells['D' + k].s.border || {}), { l: 'medium' }) });
      if (sh.cells['P' + k]) sh.style('P' + k, { border: Object.assign({}, (sh.cells['P' + k].s.border || {}), { r: 'medium' }) });
    }
    sh.r = r;
    sh.calcBreaks();
  }

  const FIJAS_MANANA = [
    { desc: 'REUNIÓN DE SEGURIDAD', ini: 390, fin: 410 },
    { desc: 'TRASLADO DE PERSONAL A PLANTA', ini: 410, fin: 430 },
    { desc: 'LLENADO DE HERRAMIENTAS DE GESTIÓN', ini: 430, fin: 450 },
    { desc: 'PREPARACIÓN DE MATERIALES', ini: 450, fin: 480 },
    { desc: 'PREPARACIÓN DE HERRAMIENTAS', ini: 480, fin: 510 },
  ];
  const HORAS = { inicioManana: 510, inicioDescanso: 720, finDescanso: 825, finJornada: 1110 };

  function calcularActividades(d) {
    const fijas = (d.fijas || FIJAS_MANANA).map((f) => Object.assign({ avance: 1, obs: '' }, f));
    const man = (d.manana || []).filter((a) => a && String(a.desc || '').trim());
    const tar = (d.tarde || []).filter((a) => a && String(a.desc || '').trim());
    const finFijas = fijas.length ? fijas[fijas.length - 1].fin : HORAS.inicioManana;
    const hm1 = distribuir(man, finFijas, HORAS.inicioDescanso);
    const ht = distribuir(tar, HORAS.finDescanso, HORAS.finJornada);
    const pct = (v) => { const n = parseFloat(String(v == null ? '' : v).replace(',', '.')); return isFinite(n) ? n / 100 : 1; };
    return fijas
      .concat(man.map((a, i) => ({ desc: a.desc, obs: a.obs, avance: pct(a.avance), ini: hm1[i].ini, fin: hm1[i].fin })))
      .concat([{ desc: 'DESCANSO', obs: '', avance: 1, ini: HORAS.inicioDescanso, fin: HORAS.finDescanso }])
      .concat(tar.map((a, i) => ({ desc: a.desc, obs: a.obs, avance: pct(a.avance), ini: ht[i].ini, fin: ht[i].fin })));
  }

  function nombreHoja(n, usados) {
    let s = String(n || 'INFORME').replace(/[\[\]\:\*\?\/\\]/g, ' ').trim().slice(0, 31) || 'INFORME';
    let k = 2, base = s;
    while (usados.has(s.toUpperCase())) s = (base.slice(0, 28) + ' ' + k++);
    usados.add(s.toUpperCase());
    return s;
  }

  // data = {fecha:'YYYY-MM-DD', hojas:[...]} ; assets = {logo:{data,ext,w,h}, firma:{...}}
  async function generar(data, assets, JSZip) {
    const styles = new Styles();
    const usados = new Set();
    const sheets = data.hojas.map((h, i) => {
      const sh = new Sheet(nombreHoja(h.nombreHoja || (i ? 'INFORME ' + (i + 1) : 'INFORME'), usados));
      construirHoja(sh, h, data.fecha, { logo: assets.logo, firma: h.incluirFirma === false ? null : (h.firmaImg || assets.firma) });
      return sh;
    });
    const zip = new JSZip();
    let ct = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>' +
      '<Default Extension="png" ContentType="image/png"/><Default Extension="jpeg" ContentType="image/jpeg"/>' +
      '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
      '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' +
      '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>' +
      '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>';
    let wbSheets = '', wbRels = '', defNames = '';
    let imgN = 0;
    sheets.forEach((sh, i) => {
      const n = i + 1;
      ct += `<Override PartName="/xl/worksheets/sheet${n}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`;
      wbSheets += `<sheet name="${esc(sh.name)}" sheetId="${n}" r:id="rId${n}"/>`;
      wbRels += `<Relationship Id="rId${n}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${n}.xml"/>`;
      defNames += `<definedName name="_xlnm.Print_Area" localSheetId="${i}">'${esc(sh.name.replace(/'/g, "''"))}'!$D$2:$P$${sh.r - 1}</definedName>`;
      zip.file(`xl/worksheets/sheet${n}.xml`, sh.xml(styles));
      if (sh.images.length) {
        ct += `<Override PartName="/xl/drawings/drawing${n}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>`;
        zip.file(`xl/worksheets/_rels/sheet${n}.xml.rels`, '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
          `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing${n}.xml"/></Relationships>`);
        let dr = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">';
        let rels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">';
        sh.images.forEach((im, k) => {
          imgN++;
          const file = `image${imgN}.${im.ext}`;
          zip.file(`xl/media/${file}`, im.data);
          rels += `<Relationship Id="rId${k + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/${file}"/>`;
          dr += `<xdr:oneCellAnchor><xdr:from><xdr:col>${im.col}</xdr:col><xdr:colOff>${im.colOff}</xdr:colOff><xdr:row>${im.row}</xdr:row><xdr:rowOff>${im.rowOff}</xdr:rowOff></xdr:from>` +
            `<xdr:ext cx="${im.cx}" cy="${im.cy}"/><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="${k + 2}" name="Imagen ${k + 1}"/><xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr></xdr:nvPicPr>` +
            `<xdr:blipFill><a:blip r:embed="rId${k + 1}"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>` +
            `<xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${im.cx}" cy="${im.cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic><xdr:clientData/></xdr:oneCellAnchor>`;
        });
        dr += '</xdr:wsDr>'; rels += '</Relationships>';
        zip.file(`xl/drawings/drawing${n}.xml`, dr);
        zip.file(`xl/drawings/_rels/drawing${n}.xml.rels`, rels);
      }
    });
    const nS = sheets.length;
    wbRels += `<Relationship Id="rId${nS + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>`;
    ct += '</Types>';
    zip.file('[Content_Types].xml', ct);
    zip.file('_rels/.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' +
      '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>' +
      '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/></Relationships>');
    zip.file('docProps/core.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">' +
      `<dc:title>Informe diario de mantenimiento eléctrico ${data.fecha}</dc:title><dc:creator>Generador de informes</dc:creator></cp:coreProperties>`);
    zip.file('docProps/app.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>Microsoft Excel</Application></Properties>');
    zip.file('xl/_rels/workbook.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${wbRels}</Relationships>`);
    zip.file('xl/workbook.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
      `<bookViews><workbookView activeTab="0"/></bookViews><sheets>${wbSheets}</sheets><definedNames>${defNames}</definedNames><calcPr calcId="191029" fullCalcOnLoad="1"/></workbook>`);
    zip.file('xl/styles.xml', styles.xml());
    return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE', compressionOptions: { level: 6 } });
  }

  function nombreArchivo(fecha) {
    const [y, m, d] = fecha.split('-');
    return `INFORME_DIARIO_DE_MANTENIMIENTO_ELECTRICO_${d}-${m}-${y}.xlsx`;
  }

  const api = { generar, calcularActividades, distribuir, nombreArchivo, addDays, hm, MESES, FIJAS_MANANA, HORAS };
  if (typeof module !== 'undefined' && module.exports) module.exports = api; else root.InformeGen = api;
})(typeof self !== 'undefined' ? self : this);
