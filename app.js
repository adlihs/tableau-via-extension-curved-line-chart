/* global tableau, d3 */
// Curved Line Viz Extension — full build
// Features: encodings + summary fallback, filters/params/selection subscriptions,
// auto render on changes/resize, numeric parsing, curve/tension, BG & line color,
// text color, axis lines toggle, area gradient, point value labels, shadows, white-dot style.

const state = {
  options: {
    curve: 'cardinal',
    tension: 0.5,
    bgColor: '#0b1020',
    lineColor: '#22aaff', // if set, overrides per-series palette
    textColor: '#e8eefb', // axis/ticks/labels
    axisLines: true,      // show axis domain/tick lines
    areaFill: false,      // draw gradient area under line
    showPointValues: false // show Y labels at points
  }
};

document.addEventListener('DOMContentLoaded', init);

async function init() {
  try {
    await tableau.extensions.initializeAsync();

    // Bind UI
    const curveSel = document.getElementById('curveSelect');
    const tensionIn = document.getElementById('tensionInput');
    const bgInput = document.getElementById('bgColor');
    const lineInput = document.getElementById('lineColor');
    const textInput = document.getElementById('textColor');
    const areaToggle = document.getElementById('areaToggle');
    const axisToggle = document.getElementById('axisLinesToggle');
    const pointValsToggle = document.getElementById('pointValuesToggle');

    if (curveSel) curveSel.addEventListener('change', (e) => {
      state.options.curve = e.target.value;
      document.getElementById('tensionWrap').style.display =
        (state.options.curve === 'cardinal' || state.options.curve === 'catmullRom') ? 'inline-flex' : 'none';
      renderFromEncodings();
    });
    if (tensionIn) tensionIn.addEventListener('input', (e) => {
      state.options.tension = parseFloat(e.target.value);
      const out = document.getElementById('tensionOut');
      if (out) out.textContent = state.options.tension.toFixed(2);
      renderFromEncodings();
    });

    if (bgInput) {
      const hBg = (e) => { state.options.bgColor = e.target.value || '#0b1020'; renderFromEncodings(); };
      bgInput.addEventListener('input', hBg); bgInput.addEventListener('change', hBg);
    }
    if (lineInput) {
      const hLine = (e) => { state.options.lineColor = e.target.value || ''; renderFromEncodings(); };
      lineInput.addEventListener('input', hLine); lineInput.addEventListener('change', hLine);
    }
    if (textInput) {
      const hText = (e) => { state.options.textColor = e.target.value || '#e8eefb'; document.documentElement.style.setProperty('--text', state.options.textColor); renderFromEncodings(); };
      textInput.addEventListener('input', hText); textInput.addEventListener('change', hText);
    }
    if (areaToggle) {
      areaToggle.checked = !!state.options.areaFill;
      areaToggle.addEventListener('change', (e) => { state.options.areaFill = !!e.target.checked; renderFromEncodings(); });
    }
    if (axisToggle) {
      axisToggle.checked = state.options.axisLines !== false;
      axisToggle.addEventListener('change', (e) => { state.options.axisLines = !!e.target.checked; renderFromEncodings(); });
    }
    if (pointValsToggle) {
      pointValsToggle.checked = !!state.options.showPointValues;
      pointValsToggle.addEventListener('change', (e) => { state.options.showPointValues = !!e.target.checked; renderFromEncodings(); });
    }

    subscribeToWorksheetEvents();
    await renderFromEncodings();
    setAutoRender();
    setStatus('✅ Listo (usa los tiles X/Y/Series en Marks)');
  } catch (err) {
    console.error(err);
    setStatus('Error de inicialización: ' + (err.message || err), true);
  }
}

function setStatus(msg, isError = false) {
  const el = document.getElementById('status');
  if (el) {
    el.textContent = msg;
    el.style.color = isError ? '#ffb3b3' : 'var(--muted)';
  }
}

// ---- Auto-render utilities ----
function debounce(fn, wait = 250) { let t; return (...a) => { clearTimeout(t); t = setTimeout(() => fn(...a), wait); }; }

let __lastSig = null;
async function computeSignatureLight() {
  try {
    const wc = tableau.extensions.worksheetContent;
    if (wc && typeof wc.getDataAsync === 'function') {
      const tbl = await wc.getDataAsync();
      const cols = (tbl.columns||[]).map(c => [
        c.encodingId || c.encoding || c.role || c.fieldRole || '',
        c.displayName || c.fieldName || '',
        c.dataType || ''
      ].join(':')).join('|');
      return 'enc:' + cols + '::rows=' + (tbl.data ? tbl.data.length : 0);
    }
  } catch(_) {}
  try {
    const wc = tableau.extensions.worksheetContent;
    if (wc && wc.worksheet && wc.worksheet.getSummaryDataAsync) {
      const s = await wc.worksheet.getSummaryDataAsync({ maxRows: 0, ignoreSelection: false });
      const cols = (s.columns||[]).map(c => (c.fieldName||'') + ':' + (c.dataType||'')).join('|');
      return 'sum:' + cols;
    }
  } catch(_) {}
  return null;
}

function setAutoRender() {
  const debouncedRender = debounce(() => renderFromEncodings(), 150);

  window.addEventListener('resize', debouncedRender);
  try {
    const ro = new ResizeObserver(debouncedRender);
    const root = document.getElementById('chartRoot');
    if (root) ro.observe(root);
    window.__curvedLineRO = ro;
  } catch(_) {}

  if (window.__curvedLineAutoTimer) clearInterval(window.__curvedLineAutoTimer);
  window.__curvedLineAutoTimer = setInterval(async () => {
    try {
      const sig = await computeSignatureLight();
      if (sig && sig !== __lastSig) {
        __lastSig = sig;
        renderFromEncodings();
      }
    } catch(_) {}
  }, 1500);
}

function subscribeToWorksheetEvents() {
  try {
    const wc = tableau.extensions.worksheetContent;
    if (!wc || !wc.worksheet) return;
    if (window.__curvedLineEventsBound) return;
    window.__curvedLineEventsBound = true;
    const ws = wc.worksheet;
    if (tableau && tableau.TableauEventType) {
      try { ws.addEventListener(tableau.TableauEventType.FilterChanged, () => renderFromEncodings()); } catch (_) {}
      try { ws.addEventListener(tableau.TableauEventType.ParameterChanged, () => renderFromEncodings()); } catch (_) {}
      try { ws.addEventListener(tableau.TableauEventType.MarkSelectionChanged, () => renderFromEncodings()); } catch (_) {}
    }
  } catch (e) {
    console.warn('No se pudieron suscribir eventos de la hoja:', e);
  }
}

// -------- Data helpers --------
function cellRaw(cell) {
  if (cell == null) return null;
  if (typeof cell === 'object') {
    if ('value' in cell && cell.value != null) return cell.value;
    if ('nativeValue' in cell && cell.nativeValue != null) return cell.nativeValue;
    if ('formattedValue' in cell && cell.formattedValue != null) return cell.formattedValue;
  }
  return cell;
}

function cellFormatted(cell) {
  if (cell && typeof cell === 'object' && 'formattedValue' in cell && cell.formattedValue != null) return cell.formattedValue;
  if (cell && typeof cell === 'object' && 'value' in cell && cell.value != null) return String(cell.value);
  return cell != null ? String(cell) : '';
}

function toNumber(v) {
  if (v == null) return NaN;
  if (typeof v === 'number') return v;
  if (typeof v === 'boolean') return v ? 1 : 0;
  let s = String(v).trim();
  if (s === '') return NaN;
  if (s.includes(',') && !s.includes('.')) s = s.replace(',', '.'); // comma decimal
  s = s.replace(/[\s\u00A0,](?=\d{3}\b)/g, ''); // thousands
  s = s.replace(/[^0-9eE\.\-\+]/g, '');
  const n = Number(s);
  return Number.isFinite(n) ? n : NaN;
}

// Encoded data or throw
async function getEncodedData() {
  const wc = tableau.extensions.worksheetContent;
  if (wc && typeof wc.getDataAsync === 'function') {
    return wc.getDataAsync(); // { columns, data }
  }
  const ctx = tableau.extensions.context;
  if (ctx && typeof ctx.getDataAsync === 'function') return ctx.getDataAsync();
  throw new Error('Esta build no expone getDataAsync para encodings.');
}

// Summary fallback
async function getSummaryData() {
  const wc = tableau.extensions.worksheetContent;
  if (!wc || !wc.worksheet) throw new Error('Worksheet host no disponible.');
  return wc.worksheet.getSummaryDataAsync({ ignoreSelection: false });
}

async function renderFromEncodings() {
  try {
    setStatus('Cargando…');

    // Apply BG to root early
    const root = document.getElementById('chartRoot');
    if (root && state.options && state.options.bgColor) root.style.background = state.options.bgColor;

    let columns = [], data = [];

    try {
      const bound = await getEncodedData();
      columns = bound.columns || [];
      data = bound.data || [];
    } catch (e) { console.warn('getEncodedData() no disponible, usando SummaryData:', e); }

    if (!data || data.length === 0) {
      const s = await getSummaryData();
      columns = s.columns || [];
      data = s.data || [];
    }

    if (!columns.length || !data.length) {
      throw new Error('No hay filas/columnas disponibles para renderizar.');
    }

    // Map encodings
    let iX, iY, iS;
    columns.forEach((c, i) => {
      const enc = (c.encodingId || c.encoding || c.role || c.fieldRole || '').toString().toLowerCase();
      const name = (c.displayName || c.fieldName || '').toString().toLowerCase();
      if (enc === 'x' || name === 'x') iX = i;
      else if (enc === 'y' || name === 'y') iY = i;
      else if (enc === 'series' || name === 'series') iS = i;
    });

    // Heurística si no hay metadatos de encoding
    if (iX == null || iY == null) {
      const isNumType = (t) => ['float','double','int','integer','number','numeric'].includes((t||'').toLowerCase());
      let yIdx = -1;
      for (let ci = 0; ci < columns.length; ci++) {
        const v = cellRaw((data[0] || [])[ci]);
        const n = toNumber(v);
        if (Number.isFinite(n) || isNumType(columns[ci].dataType)) { yIdx = ci; break; }
      }
      iY = (iY != null) ? iY : (yIdx >= 0 ? yIdx : 0);
      iX = (iX != null) ? iX : (iY === 0 ? 1 : 0);
      iS = (iS != null) ? iS : ([0,1,2].find(k => k !== iX && k !== iY && k < columns.length));
    }

    if (iX == null || iY == null || iX === iY) {
      throw new Error('Asigna campos válidos a X e Y en Marks (o verifica que haya al menos una dimensión y una medida en la vista).');
    }

    // Build rows
    const hasSeries = iS != null && iS >= 0;
    let rows = data.map(r => {
      const xCell = r[iX];
      const yCell = r[iY];
      const sCell = hasSeries ? r[iS] : null;
      const xv = cellRaw(xCell);
      const yv = toNumber(cellRaw(yCell));
      const sv = hasSeries ? cellRaw(sCell) : 'Series';
      const yLabel = cellFormatted(yCell);
      return { x: xv, y: yv, yLabel, s: sv };
    }).filter(d => d.x != null && Number.isFinite(d.y));

    if (!rows.length) throw new Error('No hay datos válidos (X nulo o Y no numérico tras parseo).');

    // Type of X
    const asDate = rows.every(d => !isNaN(Date.parse(d.x)));
    const asNum  = !asDate && rows.every(d => !isNaN(toNumber(d.x)));

    // Collapse noisy series
    if (hasSeries) {
      const sameField = (columns[iS]?.fieldName || columns[iS]?.displayName) === (columns[iX]?.fieldName || columns[iX]?.displayName);
      if (sameField) rows = rows.map(d => ({ ...d, s: 'Series' }));
      else {
        const uniq = new Set(rows.map(d => String(d.s))).size;
        if (uniq > rows.length * 0.5) rows = rows.map(d => ({ ...d, s: 'Series' }));
      }
    }

    // Sort by X
    rows.sort((a, b) => {
      let av=a.x, bv=b.x;
      if (asDate) { av=new Date(a.x).getTime(); bv=new Date(b.x).getTime(); }
      else if (asNum) { av=toNumber(a.x); bv=toNumber(b.x); }
      return av - bv;
    });

    drawChart({ rows, dimType: asDate ? 'date' : (asNum ? 'number' : 'string') });
    setStatus(`✅ Renderizado: ${rows.length} puntos`);
  } catch (e) {
    console.error(e);
    setStatus('Error: ' + (e.message || e), true);
  }
}

function curveFactory() {
  const { curve, tension } = state.options;
  switch (curve) {
    case 'basis':      return d3.curveBasis;
    case 'monotoneX':  return d3.curveMonotoneX;
    case 'natural':    return d3.curveNatural;
    case 'catmullRom': return d3.curveCatmullRom.tension(tension);
    case 'cardinal':
    default:           return d3.curveCardinal.tension(tension);
  }
}

function drawChart({ rows, dimType }) {
  const root = document.getElementById('chartRoot');
  root.innerHTML = '';

  // Apply background color
  if (state.options && state.options.bgColor) root.style.background = state.options.bgColor;

  const margin = { top: 16, right: 16, bottom: 40, left: 56 };
  let width  = root.clientWidth;
  let height = root.clientHeight;
  if (!width || width < 50) width = 900;
  if (!height || height < 50) height = 480;

  const svg = d3.select(root).append('svg').attr('width', width).attr('height', height);
  const defs = svg.append('defs');
  const g = svg.append('g').attr('transform', `translate(${margin.left},${margin.top})`);
  const innerW = width  - margin.left - margin.right;
  const innerH = height - margin.top  - margin.bottom;

  // Group by series + collapse duplicate X by mean
  const grouped = d3.group(rows, d => d.s);
  const seriesData = [];
  for (const [key, pts] of grouped.entries()) {
    const cleaned = prepSeriesPoints(pts, dimType);
    if (cleaned.length) seriesData.push({ key, pts: cleaned });
  }

  // Scales
  let xDomain;
  if (dimType === 'date')      xDomain = d3.extent(seriesData.flatMap(s => s.pts.map(d => new Date(d.x))));
  else if (dimType === 'number') xDomain = d3.extent(seriesData.flatMap(s => s.pts.map(d => +d.x)));
  else                          xDomain = Array.from(new Set(seriesData.flatMap(s => s.pts.map(d => d.x))));

  let x;
  if (dimType === 'string') x = d3.scalePoint().domain(xDomain).range([0, innerW]).padding(0.5);
  else if (dimType === 'date') x = d3.scaleTime().domain(xDomain).range([0, innerW]);
  else x = d3.scaleLinear().domain(xDomain).range([0, innerW]);

  let yExtent = d3.extent(seriesData.flatMap(s => s.pts.map(d => d.y)));
  if (!yExtent || yExtent[0] == null || yExtent[1] == null) yExtent = [0,1];
  let yDomain = yExtent;
  if (yDomain[0] === yDomain[1]) {
    const pad = Math.abs(yDomain[0]) || 1;
    yDomain = [yDomain[0] - pad * 0.5, yDomain[1] + pad * 0.5];
  }
  const y = d3.scaleLinear().domain(yDomain).nice().range([innerH, 0]);

  // Axes
  const xAxisG = g.append('g').attr('class','x-axis').attr('transform', `translate(0,${innerH})`).call(d3.axisBottom(x));
  const yAxisG = g.append('g').attr('class','y-axis').call(d3.axisLeft(y));
  const tcol = (state.options && state.options.textColor) ? state.options.textColor : '#e8eefb';
  const axisOpacity = (state.options && state.options.axisLines) ? 0.5 : 0;
  xAxisG.selectAll('text').attr('fill', tcol);
  yAxisG.selectAll('text').attr('fill', tcol);
  xAxisG.selectAll('path,line').attr('stroke', tcol).attr('opacity', axisOpacity);
  yAxisG.selectAll('path,line').attr('stroke', tcol).attr('opacity', axisOpacity);

  const palette = d3.scaleOrdinal(d3.schemeTableau10).domain(seriesData.map(s => s.key));
  const useCustomColor = !!(state.options && state.options.lineColor);

  const line = d3.line()
    .defined(d => Number.isFinite(d.y) && d.x != null)
    .curve(curveFactory())
    .x(d => dimType === 'date' ? x(new Date(d.x)) : (dimType === 'number' ? x(+d.x) : x(d.x)))
    .y(d => y(d.y));

  for (const { key, pts } of seriesData) {
    const strokeCol = useCustomColor ? state.options.lineColor : palette(key);

    // Optional area under line with gradient
    if (state.options.areaFill) {
      const gradId = `grad-${cssSafe(key)}`;
      let grad = defs.select(`#${gradId}`);
      if (grad.empty()) {
        grad = defs.append('linearGradient').attr('id', gradId).attr('x1','0').attr('y1','0').attr('x2','0').attr('y2','1');
        grad.append('stop').attr('offset','0%');
        grad.append('stop').attr('offset','100%');
      }
      grad.select('stop[offset="0%"]').attr('stop-color', strokeCol).attr('stop-opacity', 0.38);
      grad.select('stop[offset="100%"]').attr('stop-color', strokeCol).attr('stop-opacity', 0);

      const area = d3.area()
        .defined(d => Number.isFinite(d.y) && d.x != null)
        .curve(curveFactory())
        .x(d => dimType === 'date' ? x(new Date(d.x)) : (dimType === 'number' ? x(+d.x) : x(d.x)))
        .y1(d => y(d.y))
        .y0(innerH);

      g.append('path')
        .datum(pts)
        .attr('class', `series-area area-${cssSafe(key)}`)
        .attr('d', area)
        .attr('fill', `url(#${gradId})`)
        .attr('opacity', 1);
    }

    // Ensure line shadow filter (per series) exists and matches current color
    const filtId = `shadow-${cssSafe(key)}`;
    let filt = defs.select(`#${filtId}`);
    if (filt.empty()) {
      filt = defs.append('filter')
        .attr('id', filtId)
        .attr('x','-50%').attr('y','-50%').attr('width','200%').attr('height','200%');
      filt.append('feDropShadow')
        .attr('dx', 0).attr('dy', 2).attr('stdDeviation', 3)
        .attr('flood-color', strokeCol)
        .attr('flood-opacity', 0.45);
    } else {
      filt.select('feDropShadow')
        .attr('flood-color', strokeCol)
        .attr('flood-opacity', 0.45);
    }

    // Path (line)
    g.append('path')
      .datum(pts)
      .attr('class', `series-line line-${cssSafe(key)}`)
      .attr('d', line)
      .attr('fill', 'none')
      .attr('stroke', strokeCol)
      .attr('stroke-width', 4)
      .attr('stroke-linejoin', 'round')
      .attr('stroke-linecap', 'round')
      .attr('filter', `url(#${filtId})`);

    // Dots
    const dots = g.selectAll(`.dot-${cssSafe(key)}`)
      .data(pts)
      .enter()
      .append('circle')
      .attr('class', `dot dot-${cssSafe(key)}`)
      .attr('cx', d => dimType === 'date' ? x(new Date(d.x)) : (dimType === 'number' ? x(+d.x) : x(d.x)))
      .attr('cy', d => y(d.y))
      .attr('r', 3.5)
      .attr('fill', '#ffffff')
      .attr('stroke', strokeCol)
      .attr('stroke-width', 2);

    dots.append('title').text(d => `${key}\n${d.x}: ${d.y}`);

    // Optional point value labels
    if (state.options.showPointValues) {
      const lblCol = (state.options && state.options.textColor) ? state.options.textColor : '#e8eefb';
      g.selectAll(`.dot-label-${cssSafe(key)}`)
        .data(pts)
        .enter()
        .append('text')
        .attr('class', `dot-label dot-label-${cssSafe(key)}`)
        .attr('x', d => dimType === 'date' ? x(new Date(d.x)) : (dimType === 'number' ? x(+d.x) : x(d.x)))
        .attr('y', d => y(d.y) - 8)
        .attr('text-anchor', 'middle')
        .attr('font-size', 10)
        .attr('fill', lblCol)
        .attr('pointer-events', 'none')
        .text(d => (d.yLabel != null && d.yLabel !== '') ? d.yLabel : (Number.isFinite(d.y) ? d3.format('~g')(d.y) : ''));
    }
  }
}

function prepSeriesPoints(pts, dimType) {
  const keyFn = (v) => dimType === 'date' ? new Date(v).getTime()
                 : (dimType === 'number' ? +v : String(v));
  const bucket = new Map();
  for (const p of pts) {
    if (p.x == null || !Number.isFinite(+p.y)) continue;
    const k = keyFn(p.x);
    const b = bucket.get(k) || { sum: 0, count: 0, xRep: p.x };
    b.sum += +p.y; b.count += 1; bucket.set(k, b);
  }
  const collapsed = Array.from(bucket.values()).map(b => ({ x: b.xRep, y: b.sum / b.count }));
  collapsed.sort((a, b) => keyFn(a.x) - keyFn(b.x));
  return collapsed;
}

function cssSafe(s) { return String(s).replace(/[^a-z0-9_-]/gi, '-'); }
