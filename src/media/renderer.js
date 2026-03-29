/* ===========================================================
   CONSTANTS
   =========================================================== */
// Cell aspect ratio: height = CELL_RATIO * width
const CELL_RATIO     = 2.0;
const BLINK_MS       = 1000; // ms per blink cycle
const MAX_UNDO_STEPS = 50;   // Configurable: maximum number of undo/redo steps

// Grid themes
const GRID_THEMES = {
  dark:  { bg: '#000000', fg: '#ffffff' },
  light: { bg: '#ffffff', fg: '#000000' },
};

// Mutable palette — updated by theme toggle
let gridTheme = 'dark';
let GRID_BG   = GRID_THEMES.dark.bg;
let GRID_FG   = GRID_THEMES.dark.fg;

// null means "inherit GRID_FG" so default + neutral follow the theme
const FIELD_COLORS = {
  default:   null,
  pink:      '#ff69b4',
  red:       '#ff4444',
  green:     '#44ff44',
  turquoise: '#44ffee',
  blue:      '#4488ff',
  neutral:   null,
  yellow:    '#ffff44',
};

const FILL_TYPES = { empty: 'empty', X: 'X', '*': '*', '#': '#', ID: 'ID' };

function applyGridTheme() {
  const t = GRID_THEMES[gridTheme];
  GRID_BG = t.bg;
  GRID_FG = t.fg;
  const grid    = document.getElementById('bmsGrid');
  const wrapper = document.getElementById('gridWrapper');
  if (grid) {
    grid.style.background = GRID_BG;
    grid.style.color      = GRID_FG;
  }
  if (wrapper) {
    // Border uses grid foreground at reduced opacity
    wrapper.style.setProperty('--grid-border', GRID_FG + '66');
  }
  // Expose as CSS variable so cp-swatch-auto can reference it
  document.documentElement.style.setProperty('--grid-fg', GRID_FG);
}

/* ===========================================================
   STATE
   =========================================================== */
let bmsSource = window.__BMS_SOURCE__;

// acquireVsCodeApi can only be called once per page lifetime
let _vscodeApi = null;
function getVsCodeApi() {
  if (!_vscodeApi && typeof acquireVsCodeApi === 'function') {
    _vscodeApi = acquireVsCodeApi();
  }
  return _vscodeApi;
}

function getInitialConfig() {
  const configNode = document.getElementById('bmsConfigData');
  if (!configNode) return {};
  try {
    return JSON.parse(configNode.textContent || '{}');
  } catch (_) {
    return {};
  }
}

function getCurrentConfig() {
  return {
    fill: fillMode,
    sync: !!document.getElementById('autoSyncChk')?.checked,
    autoResize: !!document.getElementById('autoResizeChk')?.checked,
    theme: gridTheme,
  };
}

function persistConfig(overrides = {}) {
  const vsc = getVsCodeApi();
  if (!vsc) return;
  vsc.postMessage({ command: 'saveConfig', ...getCurrentConfig(), ...overrides });
}

let mapDef    = null;   // { rows, cols }
let fields    = [];     // parsed + augmented field objects
let cells     = [];     // flat array of cell DOM elements (row-major)

let selectedIds = new Set();  // field indices that are selected
let selectedStopperId = null; // index of the UNPROT field whose stopper is selected

let fillMode      = 'empty';
let autoResize    = false;    // auto-grow ASKIP label fields while typing
let clipboard     = null;     // copy-paste buffer: [{...field, relRow, relCol}]
let lastMouseGrid = null;     // {row,col} when mouse is over grid, else null
let _fieldIdCounter = 0;      // counter for generating unique paste IDs

// groups: array of Sets of field indices. Rebuilt after every fields mutation.
// A group id is stored as f.groupId (string) on each field.
let _nextGroupId = 1;
function _newGroupId() { return 'g' + (_nextGroupId++); }

/* ===========================================================
   ARRAY HELPERS
   =========================================================== */
// Generate a BMS-safe 7-char label for an array field:
// first 5 alphanumeric chars of arrayId (uppercase) + 2-digit 1-based index
function arrLabel(arrayId, index) {
  const base = (arrayId.replace(/[^A-Za-z0-9]/g, '').toUpperCase() || 'ARR').slice(0, 5);
  return base + String(index + 1).padStart(2, '0');
}

// Emit the BMS comment line preceding each array field:
// * {arrayId padded to col 35}{arrayId}
// col 1='*', col 2=' ', col 3=first id start, col 35=second id start (1-indexed)
function arrayCommentLine(arrayId) {
  return '* ' + arrayId.padEnd(32, ' ').slice(0, 32) + arrayId;
}

function sanitizeFieldId(value) {
  return String(value)
    .replace(/[^A-Za-z0-9-]/g, '')
    .slice(0, 25);
}

// drag/resize state
let dragState = null;
// { type: 'move'|'resize-left'|'resize-right', fieldIdx, startX, startY, origRow, origCol, origLen }

// live-sync flags (module-level so renderBms can read them)
let _syncFromFile       = false;  // true while re-rendering from an incoming file update
let _gridSizeInitialized = false; // true after first renderBms — inputs become authoritative

/* ===========================================================
   MODULE: PARSER
   =========================================================== */
const Parser = (() => {
  function parseMap(source) {
    const size = source.match(/SIZE=\((\d+),(\d+)\)/);
    const line = source.match(/LINE=(\d+)/);
    const col  = source.match(/COLUMN=(\d+)/);
    return {
      rows:          size ? +size[1] : 23,
      cols:          size ? +size[2] : 80,
      line:          line ? +line[1] : 1,
      column:        col  ? +col[1]  : 1,
      sizeFromSource: !!size,
    };
  }

  function parseFields(source) {
    // Two-pass: find every DFHMDF start position, then slice source into blocks.
    const labelRe     = /^(\w{1,7})\s+DFHMDF\b/gm;
    const unlabeledRe = /^\s{2,}DFHMDF\b/gm;
    const starts  = [];
    let m;
    while ((m = labelRe.exec(source)) !== null) {
      starts.push({ id: m[1], index: m.index });
    }
    while ((m = unlabeledRe.exec(source)) !== null) {
      starts.push({ id: '', index: m.index });
    }
    starts.sort((a, b) => a.index - b.index);

    // Detect array comment lines: "* {id}{spaces}{id}" where id is at col 3 and col 35
    // Build a map: charIndex of a DFHMDF label line → arrayId from the preceding comment
    const sourceLines = source.split('\n');
    const lineCharStart = [];
    let _cp = 0;
    for (const ln of sourceLines) { lineCharStart.push(_cp); _cp += ln.length + 1; }
    const arrayCommentForLine = new Map(); // nextLineCharStart → arrayId
    for (let li = 0; li < sourceLines.length; li++) {
      const ln = sourceLines[li];
      if (ln.length >= 3 && ln[0] === '*' && ln[1] === ' ') {
        const idM = ln.slice(2).match(/^(\S+)/);
        if (idM) {
          const id = idM[1];
          // Second occurrence must start at position 34 (0-indexed) = column 35
          if (ln.length >= 34 + id.length &&
              ln.slice(34, 34 + id.length) === id &&
              (ln.length === 34 + id.length || ln[34 + id.length] === ' ')) {
            if (li + 1 < sourceLines.length) {
              arrayCommentForLine.set(lineCharStart[li + 1], id);
            }
          }
        }
      }
    }
    starts.forEach(s => { s.arrayId = arrayCommentForLine.get(s.index) || null; });

    const result = [];
    for (let i = 0; i < starts.length; i++) {
      const blockStart = starts[i].index;
      const blockEnd   = i + 1 < starts.length ? starts[i + 1].index : source.length;
      const block      = source.slice(blockStart, blockEnd);
      const id         = starts[i].id;

      const pos    = block.match(/POS=\((\d+),(\d+)\)/);
      const len    = block.match(/LENGTH=(\d+)/i);
      const init   = block.match(/INITIAL='([^']*)'/i);

      // Parse ATTRB= — can be ATTRB=X or ATTRB=(X,Y,...)
      const attrbRaw = block.match(/ATTRB=\(?([^)\n,]+(?:,[^)\n]+)*)\)?/i);
      const attrbList = attrbRaw
        ? attrbRaw[1].split(',').map(s => s.trim().toUpperCase())
        : [];
      const prot   = attrbList.includes('PROT') || attrbList.includes('ASKIP') || /\bPROT\b/.test(block);
      const unprot = attrbList.includes('UNPROT') || /\bUNPROT\b/.test(block);
      const askip  = attrbList.includes('ASKIP');
      const bright = attrbList.includes('BRT')  ? 'brt'
                   : attrbList.includes('DRK')  ? 'drk'
                   : attrbList.includes('NORM') ? 'norm'
                   : '';   // no brightness keyword → '—' in panel
      const numeric = attrbList.includes('NUM');
      const ic      = attrbList.includes('IC');
      const fset    = attrbList.includes('FSET');

      // Parse COLOR and HILIGHT
      const colorM  = block.match(/COLOR=(\w+)/i);
      const hiliteM = block.match(/HILIGHT=(\w+)/i);
      const colorMap = { BLUE:'blue', GREEN:'green', RED:'red', YELLOW:'yellow',
                         TURQUOISE:'turquoise', PINK:'pink', WHITE:'neutral', DEFAULT:'default' };
      const parsedColor  = colorM  ? (colorMap[colorM[1].toUpperCase()]  || 'default') : 'default';
      const parsedHilit  = hiliteM ? hiliteM[1].toLowerCase() : 'off';

      // Parse OUTLINE= — can be OUTLINE=BOX, OUTLINE=OVER, OUTLINE=(OVER,UNDER), etc.
      const outlineM   = block.match(/OUTLINE=\(?([^)\n]+)\)?/i);
      const outlineRaw = outlineM
        ? outlineM[1].split(',').map(s => s.trim().toUpperCase())
        : [];
      const OUTLINE_VALID = ['OVER','UNDER','LEFT','RIGHT'];
      const parsedOutline = outlineRaw.includes('BOX')
        ? ['OVER','UNDER','LEFT','RIGHT']
        : outlineRaw.filter(s => OUTLINE_VALID.includes(s));

      if (!pos || !len) continue;

      result.push({
        id,
        row:         +pos[1] - 1,
        col:         +pos[2] - 1,
        length:      +len[1],
        initialText: init ? init[1] : '',
        prot,
        unprot,
        askip,
        brightness:  bright,
        numeric,
        ic,
        fset,
        color:     parsedColor,
        highlight: parsedHilit,
        outline:   parsedOutline,
        _srcArrayId: starts[i].arrayId, // temporary — resolved in post-processing
      });
    }

    // Post-process: detect ASKIP LENGTH=0 stopper fields.
    const stopperEntries = result.filter(f => f.askip && f.length === 0);
    const regularFields  = result.filter(f => !(f.askip && f.length === 0));
    stopperEntries.forEach(sf => {
      const parent = regularFields.find(uf =>
        uf.unprot &&
        uf.row === sf.row &&
        uf.col + uf.length === sf.col
      );
      if (parent) parent.stopper = true;
    });
    regularFields.forEach(f => {
      if (f.unprot && f.stopper === undefined) f.stopper = false;
    });

    // Post-process: reconstruct array groups from array comment markers.
    const arrayGroups = new Map(); // arrayId → [field indices in regularFields]
    regularFields.forEach((f, i) => {
      if (f._srcArrayId) {
        if (!arrayGroups.has(f._srcArrayId)) arrayGroups.set(f._srcArrayId, []);
        arrayGroups.get(f._srcArrayId).push(i);
      }
      delete f._srcArrayId;
    });
    arrayGroups.forEach((indices, arrayId) => {
      const gid = _newGroupId();
      // Infer direction: if first two members share a row → horizontal, else vertical
      let dir = 'h';
      if (indices.length >= 2) {
        const a = regularFields[indices[0]], b = regularFields[indices[1]];
        if (a.col === b.col) dir = 'v';
        else if (a.row !== b.row) dir = 'hv';
      }
      // Infer 2D structure: arrayCols, arrayRows, colStep, rowStep
      const f0 = regularFields[indices[0]];
      let arrayCols, arrayRows, colStep, rowStep;
      if (dir === 'h') {
        arrayCols = indices.length;
        arrayRows = 1;
        colStep   = indices.length >= 2 ? regularFields[indices[1]].col - f0.col : f0.length + 1;
        rowStep   = 1;
      } else if (dir === 'v') {
        arrayCols = 1;
        arrayRows = indices.length;
        colStep   = f0.length + 1;
        rowStep   = indices.length >= 2 ? regularFields[indices[1]].row - f0.row : 1;
      } else {
        // 2D: find where the row first changes to determine arrayCols
        arrayCols = indices.length;
        for (let k = 1; k < indices.length; k++) {
          if (regularFields[indices[k]].row !== f0.row) { arrayCols = k; break; }
        }
        arrayRows = Math.ceil(indices.length / arrayCols);
        colStep   = arrayCols >= 2             ? regularFields[indices[1]].col - f0.col       : f0.length + 1;
        rowStep   = arrayCols < indices.length ? regularFields[indices[arrayCols]].row - f0.row : 1;
      }
      indices.forEach((fi, pos) => {
        const f = regularFields[fi];
        f.isArray    = true;
        f.arrayId    = arrayId;
        f.arrayDir   = dir;
        f.arrayIndex = pos;
        f.groupId    = gid;
        f.arrayCols  = arrayCols;
        f.arrayRows  = arrayRows;
        f.colStep    = colStep;
        f.rowStep    = rowStep;
      });
    });

    return regularFields;
  }

  return { parseMap, parseFields };
})();

/* ===========================================================
   MODULE: FILL
   Computes display text for a field given the current fill mode.
   =========================================================== */
const Fill = (() => {
  function getText(field, mode) {
    // INITIAL text always takes precedence regardless of mode
    if (field.initialText) return formatText(field.initialText, field.length);

    // Only apply fill to PROT/UNPROT fields when they have no INITIAL
    if (!field.prot && !field.unprot) return ' '.repeat(field.length);

    switch (mode) {
      case 'empty': return ' '.repeat(field.length);
      case 'X':
      case '*':
      case '#':   return mode.repeat(field.length);
      case 'ID':  return formatText(field.id, field.length, '.');
      default:    return ' '.repeat(field.length);
    }
  }

  function formatText(str, len, padChar = ' ') {
    if (padChar !== ' ') {
      str = `<${str}:${len}>`;
    }
    if (str.length > len) return str.slice(0, len);
    return str + padChar.repeat(len - str.length);
  }

  return { getText };
})();

/* ===========================================================
   MODULE: GRID BUILDER
   =========================================================== */
const GridBuilder = (() => {
  function build(def) {
    const grid = document.getElementById('bmsGrid');
    grid.innerHTML = '';
    grid.style.background = GRID_BG;
    grid.style.color      = GRID_FG;

    const newCells = [];
    const total = def.rows * def.cols;
    const fragment = document.createDocumentFragment();
    for (let i = 0; i < total; i++) {
      const el = document.createElement('div');
      el.className = 'grid-cell';
      el.textContent = ' ';
      fragment.appendChild(el);
      newCells.push(el);
    }
    grid.appendChild(fragment);
    return newCells;
  }

  return { build };
})();

/* ===========================================================
   GRID SIZING
   Computes px cell size from wrapper width and applies it.
   Called after initial render and on every container resize.
   =========================================================== */
function fitGrid(def) {
  if (!def) return;
  const wrapper = document.getElementById('gridWrapper');
  const grid    = document.getElementById('bmsGrid');
  if (!wrapper || !grid) return;

  // Available width minus the 2×2px border
  const avail  = wrapper.clientWidth - 4;
  const cellW  = avail / def.cols;
  const cellH  = cellW * CELL_RATIO;
  // Font-size so one character (0.6em wide) fits inside cellW
  const fontSize = cellW / 0.6;

  grid.style.gridTemplateColumns = `repeat(${def.cols}, ${cellW}px)`;
  grid.style.gridTemplateRows    = `repeat(${def.rows}, ${cellH}px)`;
  grid.style.fontSize            = `${fontSize}px`;

  SelectionOverlay.resize(def);
  SelectionOverlay.draw(def, selectedIds, fields);
}

/* ===========================================================
   MODULE: FIELD RENDERER
   Applies field data to cell elements and handles color/highlight.
   =========================================================== */
const FieldRenderer = (() => {
  // Active blink intervals: fieldIdx -> intervalId
  const blinkTimers = {};

  function clearAllBlink() {
    Object.values(blinkTimers).forEach(clearInterval);
    Object.keys(blinkTimers).forEach(k => delete blinkTimers[k]);
  }

  function applyAll(mapDef, fields, cells, fillMode) {
    clearAllBlink();

    // Reset all cells first
    cells.forEach(c => {
      c.textContent = ' ';
      c.className = 'grid-cell';
      c.style.cssText = '';
    });

    fields.forEach((f, idx) => {
      const text = Fill.getText(f, fillMode);
      applyField(mapDef, cells, f, idx, text);
    });
    renderStoppers(mapDef, fields, cells);
  }

  function applyField(mapDef, cells, f, idx, text) {
    const fgColor = FIELD_COLORS[f.color] ?? GRID_FG;
    const isReverse   = f.highlight === 'reverse';
    const isBlink     = f.highlight === 'blink';
    const isUnderline = f.highlight === 'underline';

    // For reverse/blink: swap fg and fgColor as background.
    // For normal: fg is the field color; leave backgroundColor empty so the
    // CSS class `.is-field` subtle tint shows through.
    const fg = isReverse ? GRID_BG : fgColor;
    const bg = (isReverse || isBlink) ? fgColor : '';

    // Outline borders — always use the field color (not fg which may be inverted)
    const outline   = f.outline || [];
    const hasOver   = outline.includes('OVER');
    const hasUnder  = outline.includes('UNDER');
    const hasLeft   = outline.includes('LEFT');
    const hasRight  = outline.includes('RIGHT');
    const bw        = outline.length > 0 ? '1px solid ' + fgColor : '';

    for (let i = 0; i < f.length; i++) {
      const cellIdx = f.row * mapDef.cols + f.col + i;
      const cell = cells[cellIdx];
      if (!cell) continue;

      cell.textContent = text[i] ?? ' ';
      cell.className = 'grid-cell is-field';
      cell.dataset.fieldIdx = idx;

      cell.style.color           = fg;
      cell.style.backgroundColor = bg;
      cell.style.textDecoration  = isUnderline ? 'underline' : '';

      // Apply per-cell outline borders
      cell.style.borderTop    = hasOver  ? bw : '';
      cell.style.borderBottom = hasUnder ? bw : '';
      cell.style.borderLeft   = (hasLeft  && i === 0)              ? bw : '';
      cell.style.borderRight  = (hasRight && i === f.length - 1)   ? bw : '';

      if (isBlink) {
        startBlink(idx, f, mapDef, cells, fgColor);
      }
    }
  }

  function startBlink(idx, f, mapDef, cells, fgColor) {
    if (blinkTimers[idx]) return; // already running
    let phase = false;

    blinkTimers[idx] = setInterval(() => {
      phase = !phase;
      const fg = phase ? GRID_BG : fgColor;
      const bg = phase ? fgColor : '';   // '' lets CSS tint show in off-phase
      for (let i = 0; i < f.length; i++) {
        const cellIdx = f.row * mapDef.cols + f.col + i;
        const cell = cells[cellIdx];
        if (!cell) continue;
        cell.style.color           = fg;
        cell.style.backgroundColor = bg;
      }
    }, BLINK_MS);
  }

  function refreshField(mapDef, cells, f, idx, fillMode) {
    // Clear blink timer for this field if any
    if (blinkTimers[idx]) {
      clearInterval(blinkTimers[idx]);
      delete blinkTimers[idx];
    }
    const text = Fill.getText(f, fillMode);
    applyField(mapDef, cells, f, idx, text);
  }

  return { applyAll, refreshField, clearAllBlink };
})();

/* ===========================================================
   MODULE: SELECTION OVERLAY
   Drawn with XOR composite so the indicator is always visible
   regardless of field color, highlight, or grid theme.
   =========================================================== */
const SelectionOverlay = (() => {
  let canvas, ctx, cellWpx, cellHpx;
  // Lasso drag state
  let lassoRect   = null; // { x1,y1,x2,y2 } in canvas px
  // Palette drop preview
  let dropPreview = null; // [{ row, col, length }] or legacy single candidate
  // Collision state — turns selection indicator red
  let isColliding = false;

  function init() {
    canvas = document.getElementById('selectionOverlay');
    ctx    = canvas.getContext('2d');
  }

  function measure() {
    const firstCell = document.querySelector('.grid-cell');
    if (!firstCell) return;
    const r = firstCell.getBoundingClientRect();
    cellWpx = r.width;
    cellHpx = r.height;
  }

  function resize(def) {
    const grid = document.getElementById('bmsGrid');
    canvas.width  = grid.offsetWidth;
    canvas.height = grid.offsetHeight;
    canvas.style.width  = grid.offsetWidth  + 'px';
    canvas.style.height = grid.offsetHeight + 'px';
    measure();
  }

  function draw(def, selectedIds, fields) {
    if (!ctx) return;
    ctx.clearRect(0, 0, canvas.width, canvas.height);

    // Selection indicator: VSCode button background colour (reads current theme via CSS var)
    const vscBlue = getComputedStyle(document.documentElement)
      .getPropertyValue('--vscode-button-background').trim() || '#0078d4';

    // Check if selection is a single array group
    const selArr = [...selectedIds].map(i => fields[i]).filter(Boolean);
    const isSameArray = selArr.length > 1 &&
      selArr.every(f => f.isArray) &&
      new Set(selArr.map(f => f.arrayId)).size === 1;

    selectedIds.forEach(idx => {
      const f = fields[idx];
      if (!f) return;
      const x = f.col    * cellWpx;
      const y = f.row    * cellHpx;
      const w = f.length * cellWpx;
      const h = cellHpx;

      const selColor = isColliding ? '#ff4040' : vscBlue;
      ctx.strokeStyle = selColor;
      ctx.lineWidth   = 2;
      ctx.setLineDash([4, 3]);
      ctx.strokeRect(x + 1, y + 1, w - 2, h - 2);

      // Resize handles — solid filled squares, only for single selection
      if (selectedIds.size === 1) {
        ctx.fillStyle = selColor;
        ctx.setLineDash([]);
        ctx.fillRect(x,         y + Math.floor((h - 7) / 2), 6, 7);
        ctx.fillRect(x + w - 6, y + Math.floor((h - 7) / 2), 6, 7);
      }
    });

    // Array bounding box + resize handles
    if (isSameArray) {
      const first     = selArr[0];
      const arrayCols = first.arrayCols ?? 1;
      const arrayRows = first.arrayRows ?? 1;
      const colStep   = first.colStep   ?? (first.length + 1);
      const rowStep   = first.rowStep   ?? 1;
      // Find index-0 element for the anchor
      const anchor    = selArr.reduce((a, b) => (a.arrayIndex < b.arrayIndex ? a : b));
      const bx = anchor.col * cellWpx;
      const by = anchor.row * cellHpx;
      const bw = (anchor.length + (arrayCols - 1) * colStep) * cellWpx;
      const bh = (1          + (arrayRows - 1) * rowStep)  * cellHpx;
      const boxColor = isColliding ? '#ff4040' : vscBlue;
      ctx.save();
      ctx.strokeStyle = boxColor;
      ctx.lineWidth   = 2;
      ctx.setLineDash([6, 4]);
      ctx.strokeRect(bx - 3, by - 3, bw + 6, bh + 6);
      ctx.setLineDash([]);
      // Right-edge handle (add/remove columns)
      ctx.fillStyle = boxColor;
      const hx = bx + bw + 3;
      const hy = by + Math.floor((bh - 7) / 2);
      ctx.fillRect(hx, hy, 7, 7);
      // Bottom-edge handle (add/remove rows)
      const vhx = bx + Math.floor((bw - 7) / 2);
      const vhy = by + bh + 3;
      ctx.fillRect(vhx, vhy, 7, 7);
      ctx.restore();
    }

    // Stopper selection highlight
    if (selectedStopperId !== null && fields[selectedStopperId]) {
      const f = fields[selectedStopperId];
      const stopperCol = f.col + f.length;
      if (stopperCol < def.cols) {
        const x = stopperCol * cellWpx;
        const y = f.row      * cellHpx;
        ctx.save();
        ctx.strokeStyle = '#ff9944';
        ctx.lineWidth   = 2;
        ctx.setLineDash([3, 2]);
        ctx.strokeRect(x + 1, y + 1, cellWpx - 2, cellHpx - 2);
        ctx.restore();
      }
    }

    // Lasso rect (drawn normally, not XOR)
    if (lassoRect) {
      const { x1, y1, x2, y2 } = lassoRect;
      ctx.save();
      ctx.strokeStyle = 'rgba(255,255,255,0.8)';
      ctx.lineWidth   = 1;
      ctx.setLineDash([4, 3]);
      ctx.fillStyle   = 'rgba(100,180,255,0.15)';
      const rx = Math.min(x1, x2), ry = Math.min(y1, y2);
      const rw = Math.abs(x2 - x1),  rh = Math.abs(y2 - y1);
      ctx.fillRect(rx, ry, rw, rh);
      ctx.strokeRect(rx, ry, rw, rh);
      ctx.restore();
    }

    // Drop-target preview when dragging from the palette
    if (dropPreview) {
      const previews = Array.isArray(dropPreview) ? dropPreview : [dropPreview];
      ctx.save();
      ctx.fillStyle   = isColliding ? 'rgba(220,50,50,0.25)'  : 'rgba(0,127,212,0.25)';
      ctx.strokeStyle = isColliding ? '#dd3333'               : '#007fd4';
      ctx.lineWidth   = 1;
      ctx.setLineDash([4, 3]);
      previews.forEach(({ row, col, length }) => {
        ctx.fillRect(col * cellWpx, row * cellHpx, length * cellWpx, cellHpx);
        ctx.strokeRect(col * cellWpx, row * cellHpx, length * cellWpx, cellHpx);
      });
      ctx.restore();
    }
  }

  function setLasso(rect) {
    lassoRect = rect;
  }

  function setDropPreview(preview) {
    dropPreview = preview;
  }

  function setColliding(v) { isColliding = v; }

  return { init, resize, draw, setLasso, setDropPreview, setColliding, getCellSize: () => ({ cellWpx, cellHpx }) };
})();

/* ===========================================================
   MODULE: INLINE EDITOR
   Allows editing ASKIP field text directly on the grid.
   =========================================================== */
const InlineEditor = (() => {
  let activeIdx = null;

  function open(idx, fields, mapDef, cells) {
    const f = fields[idx];
    // Only ASKIP (label) fields with length > 1 are editable inline
    if (!f.askip || f.length <= 1) return;

    activeIdx = idx;
    const preEditState = JSON.parse(JSON.stringify(fields)); // capture state before editing
    const grid = document.getElementById('bmsGrid');
    const r    = grid.getBoundingClientRect();
    const cellW = r.width  / mapDef.cols;
    const cellH = r.height / mapDef.rows;

    const input = document.createElement('input');
    input.type  = 'text';
    input.value = f.initialText || '';   // never pre-fill with fill-mode text
    input.style.cssText = [
      `position:absolute`,
      `left:${f.col * cellW}px`,
      `top:${f.row * cellH}px`,
      `width:${f.length * cellW}px`,
      `height:${cellH}px`,
      `font-family:'Courier New',monospace`,
      `font-size:${grid.style.fontSize}`,
      `background:${GRID_BG}`,
      `color:${FIELD_COLORS[f.color] ?? GRID_FG}`,
      `border:none`,
      `outline:2px solid #fff`,
      `padding:0`,
      `box-sizing:border-box`,
      `z-index:20`,
    ].join(';');

    const wrapper = document.getElementById('gridWrapper');
    wrapper.appendChild(input);
    input.focus();
    input.select();

    let committed = false;
    function commit(save) {
      if (committed) return;
      committed = true;
      if (save) {
        History.pushRaw(preEditState);
        f.initialText = input.value;
      }
      input.remove();
      activeIdx = null;
      FieldRenderer.refreshField(mapDef, cells, f, idx, fillMode);
      SelectionOverlay.draw(mapDef, selectedIds, fields);
      updatePanel(fields);
    }

    input.addEventListener('input', () => {
      if (!autoResize) return;
      const txt = input.value;
      if (txt.length <= f.length) return;
      const newLen = txt.length;
      if (newLen <= mapDef.cols - f.col &&
          !fieldCollides({ row: f.row, col: f.col, length: newLen }, fields, new Set([idx]))) {
        f.length = newLen;
        input.style.width = f.length * cellW + 'px';
        FieldRenderer.applyAll(mapDef, fields, cells, fillMode);
        SelectionOverlay.resize(mapDef);
        SelectionOverlay.draw(mapDef, selectedIds, fields);
      } else if (txt.length > f.length) {
        input.value = txt.slice(0, f.length);
      }
    });

    input.addEventListener('keydown', ev => {
      if (ev.key === 'Enter')  { ev.preventDefault(); commit(true);  }
      if (ev.key === 'Escape') { ev.preventDefault(); commit(false); }
    });
    input.addEventListener('blur', () => commit(true));
  }

  function isActive() { return activeIdx !== null; }

  return { open, isActive };
})();

/* ===========================================================
   MODULE: CURSOR
   Updates the grid cursor based on hover / drag state.
   =========================================================== */
const Cursor = (() => {
  function set(type) {
    const grid = document.getElementById('bmsGrid');
    if (!grid) return;
    const map = {
      default:        'default',
      field:          'pointer',
      move:           'grab',
      grabbing:       'grabbing',
      'resize-left':  'ew-resize',
      'resize-right': 'ew-resize',
      'resize-bottom':'ns-resize',
      'ns-resize':    'ns-resize',
      lasso:          'crosshair',
      ctrlhover:      'alias',
    };
    grid.style.cursor = map[type] ?? 'default';
  }
  return { set };
})();

/* ===========================================================
   COLLISION DETECTION
   Returns true if candidate ({row, col, length}) overlaps any
   field not in excludeSet (a Set of field indices).
   =========================================================== */
function fieldCollides(candidate, allFields, excludeSet) {
  for (let i = 0; i < allFields.length; i++) {
    if (excludeSet && excludeSet.has(i)) continue;
    const f = allFields[i];
    if (f.row !== candidate.row) continue;
    const aEnd = candidate.col + candidate.length - 1;
    const bEnd = f.col + f.length - 1;
    if (candidate.col <= bEnd && aEnd >= f.col) return true;
  }
  return false;
}

function anyCollision(candidates, allFields, excludeSet) {
  return candidates.some(c => fieldCollides(c, allFields, excludeSet));
}

/* ===========================================================
   MODULE: INTERACTION
   Handles click, multi-select, drag, resize, lasso, inline-edit,
   ctrl+click-to-reveal, and cursor updates.
   =========================================================== */
const Interaction = (() => {
  let isDragging = false;
  let dragInfo   = null;

  // Lasso state (drag on empty space to rubber-band select)
  let isLasso   = false;
  let lassoInfo = null; // { startX, startY, startGridX, startGridY }

  // How many px the mouse must move before we commit to a drag vs a click
  const DRAG_THRESHOLD = 4;

  // Stored handler references so we can remove them on re-init
  let _hMousedown  = null;
  let _hMousemove  = null;
  let _hMouseleave = null;
  let _hDblclick   = null;
  let _hWinMove    = null;
  let _hWinUp      = null;

  function init(mapDef, fields, cells) {
    const grid = document.getElementById('bmsGrid');

    // Remove any previously registered handlers before adding new ones.
    // This prevents stale closures (with old fields/cells refs) from
    // accumulating after every renderBms() call.
    if (_hMousedown)  grid.removeEventListener('mousedown',  _hMousedown);
    if (_hMousemove)  grid.removeEventListener('mousemove',  _hMousemove);
    if (_hMouseleave) grid.removeEventListener('mouseleave', _hMouseleave);
    if (_hDblclick)   grid.removeEventListener('dblclick',   _hDblclick);
    if (_hWinMove)    window.removeEventListener('mousemove', _hWinMove);
    if (_hWinUp)      window.removeEventListener('mouseup',   _hWinUp);

    _hMousedown  = e => onGridMouseDown(e, mapDef, fields, cells);
    _hMousemove  = e => onGridHover(e, mapDef, fields);
    _hMouseleave = () => { lastMouseGrid = null; Cursor.set('default'); };
    _hDblclick   = e => onGridDblClick(e, mapDef, fields, cells);
    _hWinMove    = e => onMouseMove(e, mapDef, fields, cells);
    _hWinUp      = e => onMouseUp(e, mapDef, fields, cells);

    grid.addEventListener('mousedown',  _hMousedown);
    grid.addEventListener('mousemove',  _hMousemove);
    grid.addEventListener('mouseleave', _hMouseleave);
    grid.addEventListener('dblclick',   _hDblclick);
    window.addEventListener('mousemove', _hWinMove);
    window.addEventListener('mouseup',   _hWinUp);
  }

  /* ---- helpers ---- */
  function getFieldAtCell(fields, row, col) {
    for (let i = fields.length - 1; i >= 0; i--) {
      const f = fields[i];
      if (f.row === row && col >= f.col && col < f.col + f.length) return i;
    }
    return -1;
  }

  function getCellCoords(e, mapDef) {
    const grid  = document.getElementById('bmsGrid');
    const r     = grid.getBoundingClientRect();
    const x     = e.clientX - r.left;
    const y     = e.clientY - r.top;
    const cellW = r.width  / mapDef.cols;
    const cellH = r.height / mapDef.rows;
    return {
      col:   Math.floor(x / cellW),
      row:   Math.floor(y / cellH),
      px: x, py: y, cellW, cellH,
    };
  }

  function getArrayBounds(arrayId, fields, mapDef) {
    const grid    = document.getElementById('bmsGrid');
    const r       = grid.getBoundingClientRect();
    const cellW   = r.width  / mapDef.cols;
    const cellH   = r.height / mapDef.rows;
    const members = fields.filter(f => f.isArray && f.arrayId === arrayId);
    if (!members.length) return null;
    const anchor  = members.reduce((a, b) => (a.arrayIndex < b.arrayIndex ? a : b));
    const arrayCols = anchor.arrayCols ?? 1;
    const arrayRows = anchor.arrayRows ?? 1;
    const colStep   = anchor.colStep   ?? (anchor.length + 1);
    const rowStep   = anchor.rowStep   ?? 1;
    const bx = anchor.col * cellW + r.left;
    const by = anchor.row * cellH + r.top;
    const bw = (anchor.length + (arrayCols - 1) * colStep) * cellW;
    const bh = (1            + (arrayRows - 1)  * rowStep) * cellH;
    return { bx, by, bw, bh, anchor, arrayCols, arrayRows, colStep, rowStep };
  }

  function getResizeHandle(e, fieldIdx, fields, mapDef) {
    const f     = fields[fieldIdx];
    const grid  = document.getElementById('bmsGrid');
    const r     = grid.getBoundingClientRect();
    const cellW = r.width / mapDef.cols;
    const fx    = f.col * cellW + r.left;
    const fw    = f.length * cellW;
    const mx    = e.clientX;
    const my    = e.clientY;

    if (f.isArray) {
      // Handle requires all selected fields to be from this array
      const selArr = [...selectedIds].map(i => fields[i]).filter(Boolean);
      const allSameArray = selArr.length > 0 && selArr.every(s => s.isArray && s.arrayId === f.arrayId);
      if (!allSameArray) return null;
      const bounds = getArrayBounds(f.arrayId, fields, mapDef);
      if (!bounds) return null;
      const { bx, by, bw, bh } = bounds;
      const isVOnly = bounds.arrayCols === 1 && bounds.arrayRows > 1;
      if (isVOnly) {
        // V-only: right edge changes length
        if (Math.abs(mx - (bx + bw)) < 8) return 'resize-right';
        return null;
      }
      // Right edge → change cols
      if (Math.abs(mx - (bx + bw + 3)) < 9) return 'resize-right';
      // Bottom edge → change rows
      if (Math.abs(my - (by + bh + 3)) < 9) return 'resize-bottom';
      return null;
    }

    if (selectedIds.size !== 1) return null;
    if (Math.abs(mx - fx)        < 8) return 'resize-left';
    if (Math.abs(mx - (fx + fw)) < 8) return 'resize-right';
    return null;
  }

  /* ---- hover cursor update ---- */
  function onGridHover(e, mapDef, fields) {
    const { col, row } = getCellCoords(e, mapDef);
    lastMouseGrid = { row, col };
    if (isDragging || isLasso) return;
    const idx = getFieldAtCell(fields, row, col);
    if (idx < 0) {
      // Check if hovering the array bounding-box edge handles (they extend outside field cells)
      const selArr = [...selectedIds].map(i => fields[i]).filter(Boolean);
      const isSameArr = selArr.length > 1 && selArr.every(f => f.isArray) &&
                        new Set(selArr.map(f => f.arrayId)).size === 1;
      if (isSameArr) {
        const dummy = selArr[0];
        const handle = getResizeHandle(e, fields.indexOf(dummy), fields, mapDef);
        if (handle) { Cursor.set(handle === 'resize-bottom' ? 'ns-resize' : handle); return; }
      }
      Cursor.set('default'); return;
    }
    if (e.ctrlKey || e.metaKey) { Cursor.set('ctrlhover'); return; }
    const handle = getResizeHandle(e, idx, fields, mapDef);
    if (handle) { Cursor.set(handle === 'resize-bottom' ? 'ns-resize' : handle); return; }
    Cursor.set('field');
  }

  /* ---- double-click: inline edit (ASKIP/label fields only) ---- */
  function onGridDblClick(e, mapDef, fields, cells) {
    const { col, row } = getCellCoords(e, mapDef);
    const idx = getFieldAtCell(fields, row, col);
    if (idx < 0) return;
    const f = fields[idx];
    // Only open inline editor for ASKIP (label) fields with length > 1
    if (f.askip && f.length > 1) {
      InlineEditor.open(idx, fields, mapDef, cells);
    }
  }

  /* ---- mousedown ---- */
  function onGridMouseDown(e, mapDef, fields, cells) {
    if (e.button !== 0) return;
    if (InlineEditor.isActive()) return;

    // Blur any focused toolbar/panel control (e.g. fill dropdown) so that
    // keyboard shortcuts like Delete work immediately after clicking the grid.
    const ae = document.activeElement;
    if (ae && ae !== document.body && !ae.closest('#bmsGrid')) ae.blur();

    const { col, row, px, py } = getCellCoords(e, mapDef);
    const idx = getFieldAtCell(fields, row, col);

    // Clicking a stopper marker selects it; Delete then disables it
    const _stopperCell = cells[row * mapDef.cols + col];
    if (_stopperCell && _stopperCell.dataset.stopperParent !== undefined) {
      const parentIdx = parseInt(_stopperCell.dataset.stopperParent, 10);
      if (!isNaN(parentIdx) && fields[parentIdx]) {
        selectedIds.clear();
        selectedStopperId = parentIdx;
        updatePanelAndOverlay(mapDef, fields, cells);
        e.preventDefault();
        return;
      }
    }

    // Clicking anything else clears the stopper selection
    if (selectedStopperId !== null) {
      selectedStopperId = null;
    }

    // Ctrl/Cmd + click: reveal in editor
    if ((e.ctrlKey || e.metaKey) && idx >= 0) {
      e.preventDefault();
      const f = fields[idx];
      const vscode = getVsCodeApi();
      if (vscode) vscode.postMessage({ command: 'revealField', fieldId: f.id });
      return;
    }

    // Clicked empty space — check array bounding-box handles first, then lasso
    if (idx < 0) {
      // Check if clicking on array bounding-box resize handles (extend outside field cells)
      let bbHandle = null, bbIdx = -1;
      const selArr3 = [...selectedIds].map(i => fields[i]).filter(Boolean);
      const isSA3 = selArr3.length > 1 && selArr3.every(f => f.isArray) &&
                    new Set(selArr3.map(f => f.arrayId)).size === 1;
      if (isSA3) {
        bbIdx    = fields.indexOf(selArr3[0]);
        bbHandle = getResizeHandle(e, bbIdx, fields, mapDef);
      }
      if (bbHandle && bbIdx >= 0) {
        // Fall through to handle logic with the bounding-box handle
        const hf2 = fields[bbIdx];
        // reset idx to bbIdx so the handle block below fires
        const isArrayResize2 = true;
        selectedIds.clear();
        fields.forEach((fi, i) => { if (fi.isArray && fi.arrayId === hf2.arrayId) selectedIds.add(i); });
        isDragging = true;
        dragInfo = {
          type: bbHandle, fieldIdx: bbIdx,
          startX: e.clientX, startY: e.clientY,
          origRow: hf2.row, origCol: hf2.col, origLen: hf2.length,
          origState: JSON.parse(JSON.stringify(fields)),
          isArrayResize: isArrayResize2,
          arrayId:       hf2.arrayId,
          origArrayCols: hf2.arrayCols ?? 1,
          origArrayRows: hf2.arrayRows ?? 1,
          origColStep:   hf2.colStep   ?? (hf2.length + 1),
          origRowStep:   hf2.rowStep   ?? 1,
        };
        Cursor.set(bbHandle === 'resize-bottom' ? 'ns-resize' : bbHandle);
        e.preventDefault();
        return;
      }

      if (!e.shiftKey) {
        selectedIds.clear();
        updatePanelAndOverlay(mapDef, fields, cells);
      }
      // Begin lasso
      isLasso   = true;
      lassoInfo = { startX: e.clientX, startY: e.clientY, startPx: px, startPy: py };
      Cursor.set('lasso');
      e.preventDefault();
      return;
    }

    const handle = getResizeHandle(e, idx, fields, mapDef);

    if (handle) {
      const hf = fields[idx];
      const isArrayResize = hf.isArray && (handle === 'resize-right' || handle === 'resize-bottom');
      if (isArrayResize) {
        // Ensure all array members are selected before starting the drag
        selectedIds.clear();
        fields.forEach((fi, i) => { if (fi.isArray && fi.arrayId === hf.arrayId) selectedIds.add(i); });
      }
      isDragging = true;
      dragInfo = {
        type: handle, fieldIdx: idx,
        startX: e.clientX, startY: e.clientY,
        origRow: hf.row, origCol: hf.col, origLen: hf.length,
        origState: JSON.parse(JSON.stringify(fields)),
        isArrayResize,
        arrayId:       isArrayResize ? hf.arrayId              : null,
        origArrayCols: isArrayResize ? (hf.arrayCols ?? 1)     : null,
        origArrayRows: isArrayResize ? (hf.arrayRows ?? 1)     : null,
        origColStep:   isArrayResize ? (hf.colStep ?? (hf.length + 1)) : null,
        origRowStep:   isArrayResize ? (hf.rowStep ?? 1)       : null,
      };
      Cursor.set(handle === 'resize-bottom' ? 'ns-resize' : handle);
      e.preventDefault();
      return;
    }

    // Toggle selection on Shift/Ctrl, else single-select
    if (e.shiftKey || e.ctrlKey || e.metaKey) {
      if (selectedIds.has(idx)) selectedIds.delete(idx);
      else selectedIds.add(idx);
    } else {
      if (!selectedIds.has(idx)) {
        selectedIds.clear();
        selectedIds.add(idx);
      }
    }
    expandSelectionToGroups();

    updatePanelAndOverlay(mapDef, fields, cells);

    // Begin drag-move
    isDragging = true;
    dragInfo = {
      type: 'move', fieldIdx: idx,
      startX: e.clientX, startY: e.clientY,
      origRow: fields[idx].row, origCol: fields[idx].col, origLen: fields[idx].length,
      movedEnough: false,
      snapshot: [...selectedIds].map(i => ({ idx: i, row: fields[i].row, col: fields[i].col })),
      origState: JSON.parse(JSON.stringify(fields)),
    };
    Cursor.set('move');
    e.preventDefault();
  }

  /* ---- mousemove ---- */
  function onMouseMove(e, mapDef, fields, cells) {
    // Lasso mode
    if (isLasso && lassoInfo) {
      const grid  = document.getElementById('bmsGrid');
      const r     = grid.getBoundingClientRect();
      const x1    = lassoInfo.startPx;
      const y1    = lassoInfo.startPy;
      const x2    = e.clientX - r.left;
      const y2    = e.clientY - r.top;
      SelectionOverlay.setLasso({ x1, y1, x2, y2 });
      SelectionOverlay.resize(mapDef);
      SelectionOverlay.draw(mapDef, selectedIds, fields);
      return;
    }

    if (!isDragging || !dragInfo) return;

    const distX = Math.abs(e.clientX - dragInfo.startX);
    const distY = Math.abs(e.clientY - dragInfo.startY);
    if (!dragInfo.movedEnough && distX < DRAG_THRESHOLD && distY < DRAG_THRESHOLD) return;
    dragInfo.movedEnough = true;

    Cursor.set(dragInfo.type === 'move' ? 'grabbing' : dragInfo.type);

    const grid  = document.getElementById('bmsGrid');
    const r     = grid.getBoundingClientRect();
    const cellW = r.width  / mapDef.cols;
    const cellH = r.height / mapDef.rows;
    const dx    = Math.round((e.clientX - dragInfo.startX) / cellW);
    const dy    = Math.round((e.clientY - dragInfo.startY) / cellH);

    if (dragInfo.type === 'move') {
      // Clamp dx/dy so the entire group stays in bounds as a unit
      let clampedDx = dx, clampedDy = dy;
      dragInfo.snapshot.forEach(snap => {
        const f = fields[snap.idx];
        const maxDx = mapDef.cols - f.length - snap.col;
        const minDx = -snap.col;
        const maxDy = mapDef.rows - 1 - snap.row;
        const minDy = -snap.row;
        if (clampedDx > maxDx) clampedDx = maxDx;
        if (clampedDx < minDx) clampedDx = minDx;
        if (clampedDy > maxDy) clampedDy = maxDy;
        if (clampedDy < minDy) clampedDy = minDy;
      });
      dragInfo.snapshot.forEach(snap => {
        const f = fields[snap.idx];
        f.row = snap.row + clampedDy;
        f.col = snap.col + clampedDx;
      });
      dragInfo.colliding = anyCollision(
        dragInfo.snapshot.map(snap => fields[snap.idx]), fields, selectedIds
      );
    } else if (dragInfo.type === 'resize-bottom' && dragInfo.isArrayResize) {
      const arrayCols = dragInfo.origArrayCols;
      const arrayRows = dragInfo.origArrayRows;
      const rowStep   = dragInfo.origRowStep;
      const newRows   = Math.max(1, arrayRows + Math.round(dy / rowStep));
      const ok = rebuildArray(dragInfo.arrayId, arrayCols, newRows, dragInfo.origColStep, rowStep, dragInfo.origLen, true);
      dragInfo.colliding = !ok;
    } else if (dragInfo.type === 'resize-right') {
      if (dragInfo.isArrayResize) {
        const arrayCols = dragInfo.origArrayCols;
        const arrayRows = dragInfo.origArrayRows;
        const colStep   = dragInfo.origColStep;
        const isVOnly   = arrayCols === 1 && arrayRows > 1;
        let ok;
        if (isVOnly) {
          // V-only array: resize-right changes the length of all elements
          const newLen = Math.min(Math.max(1, dragInfo.origLen + dx), mapDef.cols - dragInfo.origCol);
          ok = rebuildArray(dragInfo.arrayId, arrayCols, arrayRows, colStep, dragInfo.origRowStep, newLen, true);
        } else {
          // H/2D array: resize-right adds/removes columns
          const newCols = Math.max(1, arrayCols + Math.round(dx / colStep));
          ok = rebuildArray(dragInfo.arrayId, newCols, arrayRows, colStep, dragInfo.origRowStep, dragInfo.origLen, true);
        }
        dragInfo.colliding = !ok;
      } else {
        const f = fields[dragInfo.fieldIdx];
        f.length = Math.min(Math.max(1, dragInfo.origLen + dx), mapDef.cols - f.col);
        dragInfo.colliding = fieldCollides(f, fields, new Set([dragInfo.fieldIdx]));
      }
    } else if (dragInfo.type === 'resize-left') {
      const f    = fields[dragInfo.fieldIdx];
      const newCol = Math.max(0, Math.min(dragInfo.origCol + dx, dragInfo.origCol + dragInfo.origLen - 1));
      const delta  = newCol - dragInfo.origCol;
      f.col    = newCol;
      f.length = Math.max(1, dragInfo.origLen - delta);
      dragInfo.colliding = fieldCollides(f, fields, new Set([dragInfo.fieldIdx]));
    }

    SelectionOverlay.setColliding(dragInfo.colliding || false);
    setDragError(dragInfo.colliding || false);
    FieldRenderer.applyAll(mapDef, fields, cells, fillMode);
    SelectionOverlay.resize(mapDef);
    SelectionOverlay.draw(mapDef, selectedIds, fields);
    updatePanel(fields);
  }

  /* ---- mouseup ---- */
  function onMouseUp(e, mapDef, fields, cells) {
    // Commit lasso selection
    if (isLasso && lassoInfo) {
      const grid  = document.getElementById('bmsGrid');
      const r     = grid.getBoundingClientRect();
      const cellW = r.width  / mapDef.cols;
      const cellH = r.height / mapDef.rows;

      const gx1 = lassoInfo.startPx;
      const gy1 = lassoInfo.startPy;
      const gx2 = e.clientX - r.left;
      const gy2 = e.clientY - r.top;

      const minCol = Math.max(0, Math.floor(Math.min(gx1, gx2) / cellW));
      const maxCol = Math.min(mapDef.cols - 1, Math.floor(Math.max(gx1, gx2) / cellW));
      const minRow = Math.max(0, Math.floor(Math.min(gy1, gy2) / cellH));
      const maxRow = Math.min(mapDef.rows - 1, Math.floor(Math.max(gy1, gy2) / cellH));

      if (!e.shiftKey && !e.ctrlKey && !e.metaKey) selectedIds.clear();

      fields.forEach((f, idx) => {
        // A field is inside the lasso if any of its cols overlap the lasso columns
        const fColEnd = f.col + f.length - 1;
        if (f.row >= minRow && f.row <= maxRow &&
            fColEnd >= minCol && f.col <= maxCol) {
          selectedIds.add(idx);
        }
      });

      isLasso   = false;
      lassoInfo = null;
      SelectionOverlay.setLasso(null);
      expandSelectionToGroups();
      updatePanelAndOverlay(mapDef, fields, cells);
      Cursor.set('default');
      return;
    }

    if (!isDragging) return;
    const wasColliding = dragInfo && dragInfo.colliding;
    const info = dragInfo;
    isDragging = false;
    dragInfo   = null;
    SelectionOverlay.setColliding(false);
    setDragError(false);

    if (wasColliding && info) {
      // Revert to the positions that were recorded at drag start
      if (info.type === 'move') {
        info.snapshot.forEach(snap => {
          fields[snap.idx].row = snap.row;
          fields[snap.idx].col = snap.col;
        });
      } else if (info.isArrayResize && info.origState) {
        // Array resize: restore full state (handles added/removed elements)
        fields.length = 0;
        info.origState.forEach(f => fields.push(f));
        selectedIds.clear();
        fields.forEach((f, i) => { if (f.isArray && f.arrayId === info.arrayId) selectedIds.add(i); });
        // After restoring, re-select to reset indices
        expandSelectionToGroups();
      } else {
        fields[info.fieldIdx].col    = info.origCol;
        fields[info.fieldIdx].length = info.origLen;
      }
      FieldRenderer.applyAll(mapDef, fields, cells, fillMode);
    }

    // Push undo entry if the drag committed a real change
    if (!wasColliding && info && info.movedEnough && info.origState) {
      History.pushRaw(info.origState);
    }
    Cursor.set('default');
    updatePanelAndOverlay(mapDef, fields, cells);
  }

  return { init, getCellCoords, getFieldAtCell };
})();

/* ===========================================================
   STOPPER RENDERING
   Renders a virtual '|' marker one cell to the right of UNPROT
   fields that have stopper=true, provided that cell is free and
   not beyond the grid edge.  stopperCol (0-indexed) = f.col + f.length + 1
   which maps to BMS POS column (1-indexed) = f.col + f.length + 2.
   =========================================================== */
function renderStoppers(mapDef, fields, cells) {
  if (!mapDef || !cells) return;
  // Clear any previously painted stopper cells
  cells.forEach(c => {
    if (c.dataset.stopperParent !== undefined) {
      delete c.dataset.stopperParent;
      c.classList.remove('is-stopper');
      if (!c.classList.contains('is-field')) {
        c.textContent = ' ';
        c.style.color = '';
      }
    }
  });
  fields.forEach((f, idx) => {
    if (!f.unprot || !f.stopper) return;
    const stopperCol = f.col + f.length;
    if (stopperCol >= mapDef.cols) return; // at/beyond grid edge
    // Blocked if the stopper cell is occupied by another field, OR if a field
    // starts immediately to the right (its attribute byte provides the stop).
    const blocked = fields.some(f2 =>
      f2.row === f.row && (
        (f2.col <= stopperCol && stopperCol < f2.col + f2.length) ||
        f2.col === stopperCol + 1
      )
    );
    if (blocked) return;
    const cellIdx = f.row * mapDef.cols + stopperCol;
    const cell = cells[cellIdx];
    if (!cell || cell.classList.contains('is-field')) return;
    cell.textContent = '|';
    cell.classList.add('is-stopper');
    cell.dataset.stopperParent = String(idx);
  });
}

/* ===========================================================
   MODULE: PANEL
   Controls the bottom field details panel.
   =========================================================== */
function updatePanel(fields) {
  const panel    = document.getElementById('fieldPanel');
  const idInput  = document.getElementById('panelId');
  const rowInput = document.getElementById('panelRow');
  const colInput = document.getElementById('panelCol');
  const lenInput = document.getElementById('panelLen');
  const hlSel    = document.getElementById('panelHighlight');
  const textGroup = document.getElementById('panelTextGroup');
  const textInput = document.getElementById('panelText');
  const attrGroup = document.getElementById('panelAttrGroup');
  const panelType    = document.getElementById('panelType');
  const panelBright  = document.getElementById('panelBright');
  const panelNum     = document.getElementById('panelNum');
  const panelNumLbl  = document.getElementById('panelNumLabel');
  const panelIc      = document.getElementById('panelIc');
  const panelIcLbl   = document.getElementById('panelIcLabel');
  const panelFset    = document.getElementById('panelFset');
  const panelFsetLbl = document.getElementById('panelFsetLabel');

  const outlineGroup = document.getElementById('panelOutlineGroup');
  const panelOlOver  = document.getElementById('panelOlOver');
  const panelOlUnder = document.getElementById('panelOlUnder');
  const panelOlLeft  = document.getElementById('panelOlLeft');
  const panelOlRight = document.getElementById('panelOlRight');

  if (selectedIds.size === 0) {
    panel.style.display = 'none';
    return;
  }

  panel.style.display = 'flex';
  const multi    = selectedIds.size > 1;
  const firstIdx = [...selectedIds][0];
  const first    = fields[firstIdx];
  const selArr   = [...selectedIds].map(i => fields[i]);
  // True when every selected field belongs to the same array group
  const isSameArray = selArr.every(f => f.isArray) &&
                      new Set(selArr.map(f => f.arrayId)).size === 1;

  // ID row: for arrays show arrayId (editable), else normal behaviour
  idInput.disabled  = multi && !isSameArray;
  rowInput.disabled = multi && !isSameArray;
  colInput.disabled = multi && !isSameArray;
  lenInput.disabled = multi && !isSameArray;

  if (isSameArray) {
    idInput.value = first.arrayId;
    // Show position of the first element (arrayIndex === 0) as the reference point
    const firstEl  = selArr.find(f => f.arrayIndex === 0) || first;
    rowInput.value = firstEl.row + 1;
    colInput.value = firstEl.col + 1;
    lenInput.value = firstEl.length;
  } else {
    idInput.value  = multi ? '(multiple)' : first.id;
    rowInput.value = multi ? '' : first.row + 1;
    colInput.value = multi ? '' : first.col + 1;
    lenInput.value = multi ? '' : first.length;
  }

  // Text field: only for single-field ASKIP (label) selection
  const showText = !multi && first.askip;
  textGroup.style.display = showText ? 'flex' : 'none';
  if (showText) {
    textInput.disabled = false;
    textInput.value    = first.initialText || '';
  }

  // Type dropdown
  const getType = f => f.askip ? 'askip' : f.prot ? 'prot' : 'unprot';
  const types   = new Set(selArr.map(getType));
  panelType.value    = types.size === 1 ? getType(first) : '';
  panelType.disabled = false;

  const colors  = new Set(selArr.map(f => f.color));
  const hls     = new Set(selArr.map(f => f.highlight));
  const brights = new Set(selArr.map(f => f.brightness || ''));

  setColorPickerValue(colors.size === 1 ? first.color : 'default');
  hlSel.value = hls.size === 1 ? first.highlight : 'off';

  // Brightness dropdown — works for single and multi selection
  panelBright.value    = brights.size === 1 ? (first.brightness || '') : '';
  panelBright.disabled = false;

  // Attribute checkboxes — NUM, IC, FSET only valid for UNPROT (variable) fields
  const allUnprot = selArr.every(f => f.unprot);
  if (!multi) {
    const isUnprot = !!first.unprot;
    panelNum.checked  = !!first.numeric;
    panelIc.checked   = !!first.ic;
    panelFset.checked = !!first.fset;
    panelNum.indeterminate = panelIc.indeterminate = panelFset.indeterminate = false;
    panelNumLbl.style.opacity  = isUnprot ? '1' : '0.4';
    panelIcLbl.style.opacity   = isUnprot ? '1' : '0.4';
    panelFsetLbl.style.opacity = isUnprot ? '1' : '0.4';
    panelNum.disabled  = !isUnprot;
    panelIc.disabled   = !isUnprot;
    panelFset.disabled = !isUnprot;
  } else {
    // Multi: enable checkboxes only if all selected fields are UNPROT
    panelNumLbl.style.opacity  = allUnprot ? '1' : '0.4';
    panelIcLbl.style.opacity   = allUnprot ? '1' : '0.4';
    panelFsetLbl.style.opacity = allUnprot ? '1' : '0.4';
    panelNum.disabled  = !allUnprot;
    panelIc.disabled   = !allUnprot;
    panelFset.disabled = !allUnprot;
    if (allUnprot) {
      const nums  = new Set(selArr.map(f => !!f.numeric));
      const ics   = new Set(selArr.map(f => !!f.ic));
      const fsets = new Set(selArr.map(f => !!f.fset));
      panelNum.checked  = nums.size  === 1 && nums.has(true);
      panelIc.checked   = ics.size   === 1 && ics.has(true);
      panelFset.checked = fsets.size === 1 && fsets.has(true);
      panelNum.indeterminate  = nums.size  > 1;
      panelIc.indeterminate   = ics.size   > 1;
      panelFset.indeterminate = fsets.size > 1;
    } else {
      panelNum.checked = panelIc.checked = panelFset.checked = false;
      panelNum.indeterminate = panelIc.indeterminate = panelFset.indeterminate = false;
    }
  }

  // Outline checkboxes — available for any field type
  if (outlineGroup) {
    const SIDES = [
      { el: panelOlOver,  key: 'OVER'  },
      { el: panelOlUnder, key: 'UNDER' },
      { el: panelOlLeft,  key: 'LEFT'  },
      { el: panelOlRight, key: 'RIGHT' },
    ];
    SIDES.forEach(({ el, key }) => {
      if (!el) return;
      if (!multi) {
        el.checked = (first.outline || []).includes(key);
        el.indeterminate = false;
      } else {
        const vals = new Set(selArr.map(f => (f.outline || []).includes(key)));
        el.checked       = vals.size === 1 && vals.has(true);
        el.indeterminate = vals.size === 2;
      }
      el.disabled = false;
    });
  }

  // Stopper checkbox — only visible when all selected fields are UNPROT
  const panelStopperRow = document.getElementById('panelStopperRow');
  const panelStopper    = document.getElementById('panelStopper');
  if (panelStopperRow && panelStopper) {
    const allUnprot = selArr.every(f => f.unprot);
    panelStopperRow.style.display = allUnprot ? '' : 'none';
    if (allUnprot) {
      if (!multi) {
        panelStopper.checked       = !!first.stopper;
        panelStopper.indeterminate = false;
      } else {
        const vals = new Set(selArr.map(f => !!f.stopper));
        panelStopper.checked       = vals.size === 1 && vals.has(true);
        panelStopper.indeterminate = vals.size === 2;
      }
      panelStopper.disabled = false;
    }
  }

  // Array config row — visible when all selected fields are from the same array
  const panelArrayRow = document.getElementById('panelArrayRow');
  if (panelArrayRow) {
    if (isSameArray) {
      const arrayCols = first.arrayCols ?? (first.arrayDir === 'v' ? 1 : selArr.length);
      const arrayRows = first.arrayRows ?? (first.arrayDir === 'v' ? selArr.length : 1);
      const colStep   = first.colStep   ?? (first.length + 1);
      const rowStep   = first.rowStep   ?? 1;
      const panelArrayCols    = document.getElementById('panelArrayCols');
      const panelArrayRows    = document.getElementById('panelArrayRows');
      const panelArrayColStep = document.getElementById('panelArrayColStep');
      const panelArrayRowStep = document.getElementById('panelArrayRowStep');
      if (panelArrayCols)    panelArrayCols.value    = arrayCols;
      if (panelArrayRows)    panelArrayRows.value    = arrayRows;
      if (panelArrayColStep) panelArrayColStep.value = colStep - first.length; // display gap, not step
      if (panelArrayRowStep) panelArrayRowStep.value = rowStep - 1; // display gap, not step
      panelArrayRow.style.display = 'flex';
    } else {
      panelArrayRow.style.display = 'none';
    }
  }
}

// Color options data — must match the HTML picker options
const COLOR_OPTIONS = [
  { value: 'default',   hex: '',        label: 'Default'   },
  { value: 'pink',      hex: '#ff69b4', label: 'Pink'      },
  { value: 'red',       hex: '#ff4444', label: 'Red'       },
  { value: 'green',     hex: '#44ff44', label: 'Green'     },
  { value: 'turquoise', hex: '#44ffee', label: 'Turquoise' },
  { value: 'blue',      hex: '#4488ff', label: 'Blue'      },
  { value: 'neutral',   hex: '',        label: 'Neutral'   },
  { value: 'yellow',    hex: '#ffff44', label: 'Yellow'    },
];

function setColorPickerValue(val) {
  const opt = COLOR_OPTIONS.find(o => o.value === val) || COLOR_OPTIONS[0];
  const swatchEl = document.getElementById('colorPickerSwatch');
  const labelEl  = document.getElementById('colorPickerLabel');
  if (swatchEl) {
    if (opt.hex) {
      swatchEl.className = 'cp-swatch';
      swatchEl.style.background = opt.hex;
    } else {
      swatchEl.className = 'cp-swatch cp-swatch-auto';
      swatchEl.style.background = '';
    }
  }
  if (labelEl) labelEl.textContent = opt.label;
  const menu = document.getElementById('colorPickerMenu');
  if (menu) {
    menu.querySelectorAll('.cp-option').forEach(el => {
      el.classList.toggle('cp-selected', el.dataset.value === val);
    });
  }
}

function getColorPickerValue() {
  const sel = document.querySelector('#colorPickerMenu .cp-option.cp-selected');
  return sel ? sel.dataset.value : 'default';
}

function updatePanelAndOverlay(mapDef, fields, cells) {
  updatePanel(fields);
  SelectionOverlay.resize(mapDef);
  SelectionOverlay.draw(mapDef, selectedIds, fields);
}

/* ===========================================================
   MODULE: HISTORY  (Undo / Redo)
   Uses JSON round-trip to deep-clone the entire fields array so
   any future field properties are automatically captured.
   =========================================================== */
const History = (() => {
  let undoStack = [];
  let redoStack = [];

  function snapshot() {
    return JSON.parse(JSON.stringify(fields));
  }

  // Call BEFORE mutating fields (captures current state, clears redo)
  function push() {
    undoStack.push(snapshot());
    if (undoStack.length > MAX_UNDO_STEPS) undoStack.shift();
    redoStack = [];
  }

  // Push an externally-captured snapshot (used after drag completes)
  function pushRaw(snap) {
    undoStack.push(snap);
    if (undoStack.length > MAX_UNDO_STEPS) undoStack.shift();
    redoStack = [];
  }

  function restore(snap) {
    fields.length = 0;
    snap.forEach(f => fields.push(f));
    selectedIds.clear();
    selectedStopperId = null;
    if (!mapDef) return;
    FieldRenderer.applyAll(mapDef, fields, cells, fillMode);
    SelectionOverlay.resize(mapDef);
    SelectionOverlay.draw(mapDef, selectedIds, fields);
    updatePanel(fields);
  }

  function undo() {
    if (!undoStack.length) return;
    redoStack.push(snapshot());
    restore(undoStack.pop());
  }

  function redo() {
    if (!redoStack.length) return;
    undoStack.push(snapshot());
    if (undoStack.length > MAX_UNDO_STEPS) undoStack.shift();
    restore(redoStack.pop());
  }

  function clear() { undoStack = []; redoStack = []; }

  return { push, pushRaw, undo, redo, clear };
})();

/* ===========================================================
   GRID MESSAGE BAR
   Displays paste errors and terminal-size warnings below the grid.
   =========================================================== */
let _gridSizeWarn  = false;
let _pasteError    = '';
let _pasteErrTimer = null;
let _dragError     = false;

function setGridSizeWarn(v) {
  _gridSizeWarn = v;
  _refreshGridMsg();
}

function setPasteError(msg) {
  _pasteError = msg;
  _refreshGridMsg();
  clearTimeout(_pasteErrTimer);
  if (msg) {
    _pasteErrTimer = setTimeout(() => { _pasteError = ''; _refreshGridMsg(); }, 4000);
  }
}

function setDragError(v) {
  if (_dragError === v) return;
  _dragError = v;
  _refreshGridMsg();
}

function _refreshGridMsg() {
  const el = document.getElementById('gridMsg');
  if (!el) return;
  if (_pasteError) {
    el.className   = 'error';
    el.textContent = _pasteError;
    el.style.display = 'block';
  } else if (_dragError) {
    el.className   = 'error';
    el.textContent = '\u26d4 Cannot place here — overlaps another field';
    el.style.display = 'block';
  } else if (_gridSizeWarn) {
    el.className   = 'warning';
    el.textContent = '\u26a0 Some terminals may not support sizes larger than 80\u00d724';
    el.style.display = 'block';
  } else {
    el.style.display = 'none';
    el.textContent   = '';
  }
}

/* ===========================================================
   MODULE: GROUPS
   Fields carry a `groupId` string property. Fields sharing the
   same groupId are treated as a group: clicking one selects all.
   =========================================================== */
function groupSelected() {
  if (selectedIds.size < 2 || !mapDef) return;
  History.push();
  // Collect all existing groupIds from the selection so we can
  // absorb members of other groups into this new group.
  const memberIds = new Set(selectedIds);
  const involvedGroupIds = new Set(
    [...selectedIds].map(i => fields[i].groupId).filter(Boolean)
  );
  // Pull in any other fields that share one of those group ids
  fields.forEach((f, i) => {
    if (f.groupId && involvedGroupIds.has(f.groupId)) memberIds.add(i);
  });
  const gid = _newGroupId();
  memberIds.forEach(i => { fields[i].groupId = gid; });
  // Expand selection to all group members
  memberIds.forEach(i => selectedIds.add(i));
  updatePanelAndOverlay(mapDef, fields, cells);
}

function ungroupSelected() {
  if (selectedIds.size === 0 || !mapDef) return;
  // Array groups can never be ungrouped
  if ([...selectedIds].some(i => fields[i] && fields[i].isArray)) {
    setPasteError('Array groups cannot be ungrouped. Delete the fields to remove the array.');
    return;
  }
  History.push();
  // Collect all groupIds touched by the selection
  const involvedGroupIds = new Set(
    [...selectedIds].map(i => fields[i].groupId).filter(Boolean)
  );
  // Remove groupId from every member of those groups
  fields.forEach(f => {
    if (f.groupId && involvedGroupIds.has(f.groupId)) delete f.groupId;
  });
  updatePanelAndOverlay(mapDef, fields, cells);
}

/* ===========================================================
   REBUILD ARRAY
   Replaces array members with a new set given updated structure.
   Returns true on success, false if any position is out-of-bounds
   or would collide with a non-member field.
   =========================================================== */
function rebuildArray(arrayId, newArrayCols, newArrayRows, newColStep, newRowStep, newLength, skipHistory = false) {
  const memberIndices = fields
    .map((f, i) => i)
    .filter(i => fields[i].isArray && fields[i].arrayId === arrayId)
    .sort((a, b) => fields[a].arrayIndex - fields[b].arrayIndex);
  if (memberIndices.length === 0 || !mapDef) return false;

  const firstF   = fields[memberIndices[0]];
  const firstRow = firstF.row;
  const firstCol = firstF.col;
  const gid      = firstF.groupId;
  const newCount = newArrayCols * newArrayRows;

  // Compute target positions
  const newPositions = [];
  for (let ri = 0; ri < newArrayRows; ri++) {
    for (let ci = 0; ci < newArrayCols; ci++) {
      newPositions.push({
        row:    firstRow + ri * newRowStep,
        col:    firstCol + ci * newColStep,
        length: newLength,
      });
    }
  }

  // Validate: in bounds and no collision with non-member fields
  const memberIdxSet = new Set(memberIndices);
  const valid = newPositions.every(p =>
    p.row >= 0 && p.row < mapDef.rows &&
    p.col >= 0 && p.col + p.length <= mapDef.cols &&
    !fieldCollides(p, fields, memberIdxSet)
  );
  if (!valid) return false;

  if (!skipHistory) History.push();

  let newDir = 'h';
  if (newArrayCols === 1 && newArrayRows > 1) newDir = 'v';
  else if (newArrayCols > 1 && newArrayRows > 1) newDir = 'hv';

  const template = { ...firstF };

  const applyProps = (f, idx) => Object.assign(f, template, {
    id:         arrLabel(arrayId, idx),
    row:        newPositions[idx].row,
    col:        newPositions[idx].col,
    length:     newLength,
    arrayIndex: idx,
    arrayCols:  newArrayCols,
    arrayRows:  newArrayRows,
    colStep:    newColStep,
    rowStep:    newRowStep,
    arrayDir:   newDir,
    groupId:    gid,
  });

  if (newCount <= memberIndices.length) {
    memberIndices.slice(0, newCount).forEach((fi, idx) => applyProps(fields[fi], idx));
    // Remove excess in descending order so earlier splices don't shift later indices
    memberIndices.slice(newCount).sort((a, b) => b - a).forEach(i => {
      fields.splice(i, 1);
      if (selectedStopperId !== null) {
        if (selectedStopperId === i)  selectedStopperId = null;
        else if (selectedStopperId > i) selectedStopperId--;
      }
    });
  } else {
    memberIndices.forEach((fi, idx) => applyProps(fields[fi], idx));
    for (let idx = memberIndices.length; idx < newCount; idx++) {
      const nf = {};
      applyProps(nf, idx);
      fields.push(nf);
    }
  }

  // Rebuild selection to match all current array members
  selectedIds.clear();
  fields.forEach((f, i) => { if (f.isArray && f.arrayId === arrayId) selectedIds.add(i); });
  return true;
}

// When a field belonging to a group is clicked, expand selectedIds to
// include all group members. Call after any single-click selection change.
function expandSelectionToGroups() {
  let changed = true;
  while (changed) {
    changed = false;
    const groupIds = new Set(
      [...selectedIds].map(i => fields[i] && fields[i].groupId).filter(Boolean)
    );
    fields.forEach((f, i) => {
      if (f.groupId && groupIds.has(f.groupId) && !selectedIds.has(i)) {
        selectedIds.add(i);
        changed = true;
      }
    });
  }
}

/* ===========================================================
   DELETE SELECTED FIELDS
   =========================================================== */
function deleteSelected() {
  if (selectedIds.size === 0 || !mapDef) return;
  History.push();
  // Sort descending so splice doesn't shift indices of remaining items
  const toRemove = [...selectedIds].sort((a, b) => b - a);
  toRemove.forEach(idx => fields.splice(idx, 1));
  selectedIds.clear();
  FieldRenderer.applyAll(mapDef, fields, cells, fillMode);
  SelectionOverlay.resize(mapDef);
  SelectionOverlay.draw(mapDef, selectedIds, fields);
  updatePanel(fields);
}

/* ===========================================================
   COPY / PASTE FIELDS   (Ctrl+C / Ctrl+V)
   =========================================================== */
function _makeUniqueId() {
  let id;
  do {
    id = 'CPY' + String(++_fieldIdCounter).padStart(4, '0');
  } while (fields.some(f => f.id === id));
  return id;
}

function copySelected() {
  if (selectedIds.size === 0) return;
  const selected = [...selectedIds].map(i => JSON.parse(JSON.stringify(fields[i])));
  const minRow   = Math.min(...selected.map(f => f.row));
  const minCol   = Math.min(...selected.map(f => f.col));
  clipboard = selected.map(f => ({ ...f, relRow: f.row - minRow, relCol: f.col - minCol }));
}

function pasteFields() {
  if (!clipboard || !mapDef) return;

  function tryPasteAt(baseRow, baseCol) {
    for (const f of clipboard) {
      const row = baseRow + f.relRow;
      const col = baseCol + f.relCol;
      if (row < 0 || row >= mapDef.rows)               return false;
      if (col < 0 || col + f.length - 1 >= mapDef.cols) return false;
      if (fieldCollides({ row, col, length: f.length }, fields, null)) return false;
    }
    return true;
  }

  let pasteRow = null, pasteCol = null;

  // Try at current cursor position first
  if (lastMouseGrid && tryPasteAt(lastMouseGrid.row, lastMouseGrid.col)) {
    pasteRow = lastMouseGrid.row;
    pasteCol = lastMouseGrid.col;
  }

  // Fall back: scan grid row-by-row for first available position
  if (pasteRow === null) {
    outer: for (let r = 0; r < mapDef.rows; r++) {
      for (let c = 0; c < mapDef.cols; c++) {
        if (tryPasteAt(r, c)) { pasteRow = r; pasteCol = c; break outer; }
      }
    }
  }

  if (pasteRow !== null) {
    History.push();
    selectedIds.clear();
    clipboard.forEach(f => {
      const newField = { ...JSON.parse(JSON.stringify(f)), id: _makeUniqueId(),
                         row: pasteRow + f.relRow, col: pasteCol + f.relCol };
      delete newField.relRow;
      delete newField.relCol;
      fields.push(newField);
      selectedIds.add(fields.length - 1);
    });
    FieldRenderer.applyAll(mapDef, fields, cells, fillMode);
    updatePanelAndOverlay(mapDef, fields, cells);
    setPasteError('');
  } else {
    setPasteError("\u274c Can\u2019t paste: no space available for the copied fields");
  }
}

/* ===========================================================
   GENERATE BMS SOURCE
   Builds a simplified BMS assembler listing from current fields.
   =========================================================== */
function generateBmsSource() {
  if (!mapDef) return '';

  function normalizeMapName(name) {
    const cleaned = (name || '')
      .toUpperCase()
      .replace(/[^A-Z0-9@$#]/g, '')
      .slice(0, 7);
    return cleaned || 'MAPNAME';
  }

  function splitInitialText(text) {
    return String(text).replace(/'/g, "''");
  }

  function bmsLines(label, keyword, attrs) {
    const COL = 71;
    const lines = [];
    const firstPrefix = `${label.padEnd(8)} ${keyword} `;
    const nextPrefix = '               ';

    attrs.forEach((attr, index) => {
      const suffix = index < attrs.length - 1 ? ',' : '';
      const line = (index === 0 ? firstPrefix : nextPrefix) + attr + suffix;
      lines.push(index < attrs.length - 1 ? line.padEnd(COL) + 'X' : line);
    });

    if (lines.length === 0) {
      lines.push(firstPrefix.trimEnd());
    }

    return lines.join('\n');
  }

  const mapName = normalizeMapName(document.title || document.querySelector('.title')?.textContent || '');

  const out = [];
  out.push(bmsLines(mapName, 'DFHMSD', [
    'TYPE=&SYSPARM',
    'MODE=INOUT',
    'STORAGE=AUTO',
    'CTRL=FREEKB',
    'EXTATT=YES',
    'TERM=3270-2',
    'TIOAPFX=YES',
    'MAPATTS=(COLOR,HILIGHT,OUTLINE,PS,SOSI)',
    'DSATTS=(COLOR,HILIGHT,OUTLINE,PS,SOSI)',
  ]));
  out.push(bmsLines(mapName, 'DFHMDI', [
    `SIZE=(${mapDef.rows},${mapDef.cols})`,
    'COLUMN=1',
    'LINE=1',
  ]));

  fields.forEach((f, idx) => {
    const attrs = [];
    attrs.push(`POS=(${f.row + 1},${f.col + 1})`);
    attrs.push(`LENGTH=${f.length}`);

    // Build ATTRB list
    const attrb = [];
    if (f.askip) {
      attrb.push('ASKIP');
    } else if (f.prot) {
      attrb.push('PROT');
    } else if (f.unprot) {
      attrb.push('UNPROT');
    }
    if (f.brightness) attrb.push(f.brightness.toUpperCase()); // NORM/BRT/DRK; omit only when '—'
    if (f.numeric) attrb.push('NUM');
    if (f.ic)      attrb.push('IC');
    if (f.fset)    attrb.push('FSET');
    if (attrb.length > 0) {
      attrs.push(attrb.length === 1 ? `ATTRB=${attrb[0]}` : `ATTRB=(${attrb.join(',')})`);
    }

    // HILIGHT — always specified
    attrs.push(`HILIGHT=${(f.highlight || 'off').toUpperCase()}`);

    // COLOR — only if not default
    if (f.color && f.color !== 'default') {
      const cMap = { neutral: 'WHITE', turquoise: 'TURQUOISE', pink: 'PINK' };
      attrs.push('COLOR=' + (cMap[f.color] || f.color.toUpperCase()));
    }

    // OUTLINE — only if any side is active
    const outline = f.outline || [];
    if (outline.length > 0) {
      const allFour = ['OVER','UNDER','LEFT','RIGHT'].every(s => outline.includes(s));
      if (allFour) {
        attrs.push('OUTLINE=BOX');
      } else if (outline.length === 1) {
        attrs.push(`OUTLINE=${outline[0]}`);
      } else {
        attrs.push(`OUTLINE=(${outline.join(',')})`);
      }
    }

    // INITIAL
    if (f.initialText) attrs.push(`INITIAL='${splitInitialText(f.initialText)}'`);

    // Array fields: emit comment line before each member and use numbered label
    if (f.isArray) out.push(arrayCommentLine(f.arrayId));
    const fieldLabel = f.askip ? '' : (f.isArray ? arrLabel(f.arrayId, f.arrayIndex) : f.id);
    out.push(bmsLines(fieldLabel, 'DFHMDF', attrs));

    // Emit ASKIP stopper field after UNPROT fields with stopper=true,
    // unless the stopper position is off-grid or blocked by another field.
    if (f.unprot && f.stopper) {
      const stopperCol = f.col + f.length;
      if (stopperCol < mapDef.cols) {
        const blocked = fields.some((f2, i2) =>
          i2 !== idx &&
          f2.row === f.row && (
            (f2.col <= stopperCol && stopperCol < f2.col + f2.length) ||
            f2.col === stopperCol + 1
          )
        );
        if (!blocked) {
          out.push(bmsLines('', 'DFHMDF', [
            `POS=(${f.row + 1},${stopperCol + 1})`,
            'LENGTH=0',
            'ATTRB=ASKIP',
          ]));
        }
      }
    }
  });

  out.push(`${mapName.padEnd(8)} DFHMSD TYPE=FINAL `);
  out.push('        END');
  return out.join('\n');
}

/* ===========================================================
   MAIN: RENDER
   =========================================================== */
function renderBms() {
  History.clear();
  const parsed = Parser.parseMap(bmsSource);

  // Determine grid dimensions:
  //   - First ever render OR file→renderer update: let file's SIZE= (if present) initialize inputs.
  //   - All other calls (user changed input, user action): inputs are authoritative.
  const rIn = document.getElementById('gridRowsInput');
  const cIn = document.getElementById('gridColsInput');
  if (!_gridSizeInitialized || (_syncFromFile && parsed.sizeFromSource)) {
    _gridSizeInitialized = true;
    if (parsed.sizeFromSource) {
      if (rIn) rIn.value = parsed.rows;
      if (cIn) cIn.value = parsed.cols;
    }
  }
  const rows = (rIn && parseInt(rIn.value, 10) >= 1) ? parseInt(rIn.value, 10) : parsed.rows;
  const cols = (cIn && parseInt(cIn.value, 10) >= 1) ? parseInt(cIn.value, 10) : parsed.cols;
  mapDef = { ...parsed, rows, cols };

  // Update palette array count-input maxes to match the grid dimensions
  const _hCnt = document.querySelector('.palette-count-input[data-direction="horizontal"]');
  const _vCnt = document.querySelector('.palette-count-input[data-direction="vertical"]');
  if (_hCnt) _hCnt.max = cols;
  if (_vCnt) _vCnt.max = rows;

  fields = Parser.parseFields(bmsSource);
  cells  = GridBuilder.build(mapDef);

  FieldRenderer.applyAll(mapDef, fields, cells, fillMode);
  applyGridTheme();

  // Size cells to fit container, then draw overlay
  requestAnimationFrame(() => {
    fitGrid(mapDef);
  });

  // Attach interaction (re-init on every render)
  Interaction.init(mapDef, fields, cells);
}

/* ===========================================================
   MODULE: PALETTE DRAG
   Handles dragging items from the palette panel onto the BMS
   grid to create new fields.
   =========================================================== */
const PaletteDrag = (() => {
  let dragItem      = null;   // { ftype, length }
  let dropCoords    = null;   // { row, col } or null
  let dropCandidates = null;  // accepted candidate placements for array drops
  let dropColliding = false;
  let ghost         = null;
  let moveHandler   = null;
  let upHandler     = null;
  let fieldCounter  = 0;

  // Representative ghost text for each type + length
  const GHOST_TEXT = {
    label:       { 1: 'A', 8: 'Label   ', 16: 'Long Label      ' },
    inputtext:   { 1: '_', 8: '________', 16: '________________' },
    inputnum:    { 1: '0', 8: '00000000', 16: '0000000000000000' },
    outline:     { 1: '[A]', 8: '[Label  ]', 16: '[Long Label     ]' },
    'outline-var': { 8: '[________]', 16: '[________________]' },
  };

  function genId(ftype) {
    return ftype.slice(0, 3).toUpperCase() + String(++fieldCounter).padStart(4, '0');
  }

  function init() {
    ghost = document.getElementById('paletteDragGhost');
    document.querySelectorAll('.palette-item[data-ftype]').forEach(el => {
      el.addEventListener('mousedown', e => start(e, el));
    });
    // Count inputs beside array palette items: don't start a drag; clamp to valid range
    document.querySelectorAll('.palette-count-input').forEach(input => {
      input.addEventListener('mousedown', e => e.stopPropagation());
      input.addEventListener('change', () => {
        const max = parseInt(input.max, 10) || 99;
        const val = Math.max(1, Math.min(max, parseInt(input.value, 10) || 1));
        input.value = val;
      });
    });
  }

  function getArrayCandidates(row, col, length, count, dir) {
    const step = dir === 'h' ? length + 1 : 1;
    const candidates = [];
    for (let i = 0; i < count; i++) {
      const candidate = {
        row: dir === 'v' ? row + i : row,
        col: dir === 'h' ? col + i * step : col,
        length,
      };
      if (candidate.row >= mapDef.rows) break;
      if (candidate.col + length > mapDef.cols) break;
      if (fieldCollides(candidate, fields, null)) break;
      candidates.push(candidate);
    }
    return candidates;
  }

  function start(e, el) {
    if (e.button !== 0) return;
    const ftype = el.dataset.ftype;
    e.preventDefault();

    const length    = parseInt(el.dataset.length, 10);
    const direction = el.dataset.direction || 'horizontal';
    // Count is stored in a sibling .palette-count-input (outside the palette-item div)
    const _cntInp = el.parentElement && el.parentElement.querySelector('.palette-count-input');
    const count   = _cntInp ? Math.max(1, parseInt(_cntInp.value, 10) || 1) : parseInt(el.dataset.count || '3', 10);
    dragItem   = { ftype, length, count, direction };
    dropCoords = null;
    dropCandidates = null;

    // Clear grid selection so existing fields don't show collision state
    selectedIds.clear();
    SelectionOverlay.setColliding(false);

    let ghostTxt;
    if (ftype === 'array') {
      const dirArrow = direction === 'vertical' ? '↓' : '→';
      ghostTxt = `${dirArrow} Array ×${count} (${length})`;
    } else {
      ghostTxt = (GHOST_TEXT[ftype] || {})[length] || ftype + ':' + length;
    }
    ghost.textContent = ghostTxt;
    ghost.style.left    = e.clientX + 'px';
    ghost.style.top     = e.clientY + 'px';
    ghost.style.display = 'block';
    document.body.style.cursor = 'grabbing';

    moveHandler = onMove;
    upHandler   = onUp;
    window.addEventListener('mousemove', moveHandler);
    window.addEventListener('mouseup',   upHandler);
  }

  function onMove(e) {
    if (!dragItem || !mapDef) return;

    ghost.style.left = e.clientX + 'px';
    ghost.style.top  = e.clientY + 'px';

    const grid = document.getElementById('bmsGrid');
    if (!grid) return;
    const r = grid.getBoundingClientRect();

    if (e.clientX >= r.left && e.clientX <= r.right &&
        e.clientY >= r.top  && e.clientY <= r.bottom) {
      const cellW = r.width  / mapDef.cols;
      const cellH = r.height / mapDef.rows;
      const col   = Math.min(
        Math.max(0, Math.floor((e.clientX - r.left) / cellW)),
        mapDef.cols - dragItem.length
      );
      const row   = Math.min(
        Math.max(0, Math.floor((e.clientY - r.top) / cellH)),
        mapDef.rows - 1
      );
      dropCoords = { row, col };
      if (dragItem.ftype === 'array') {
        const dir = dragItem.direction === 'vertical' ? 'v' : 'h';
        dropCandidates = getArrayCandidates(row, col, dragItem.length, dragItem.count, dir);
        dropColliding = dropCandidates.length === 0;
        const preview = dropCandidates.length > 0 ? dropCandidates : [{ row, col, length: dragItem.length }];
        SelectionOverlay.setDropPreview(preview);
      } else {
        dropCandidates = null;
        dropColliding = fieldCollides({ row, col, length: dragItem.length }, fields, null);
        SelectionOverlay.setDropPreview({ row, col, length: dragItem.length });
      }
      SelectionOverlay.setColliding(dropColliding);
    } else {
      dropCoords    = null;
      dropCandidates = null;
      dropColliding = false;
      SelectionOverlay.setColliding(false);
      SelectionOverlay.setDropPreview(null);
    }

    SelectionOverlay.resize(mapDef);
    SelectionOverlay.draw(mapDef, selectedIds, fields);
  }

  function onUp() {
    window.removeEventListener('mousemove', moveHandler);
    window.removeEventListener('mouseup',   upHandler);

    ghost.style.display = 'none';
    document.body.style.cursor = '';
    SelectionOverlay.setColliding(false);
    SelectionOverlay.setDropPreview(null);

    if (dropCoords && dragItem && mapDef && !dropColliding) {
      const { row, col } = dropCoords;
      const { ftype, length, count, direction } = dragItem;

      if (ftype === 'array') {
        const dir   = direction === 'vertical' ? 'v' : 'h';
        const gid   = _newGroupId();
        const aid   = 'ARRAY' + String(++fieldCounter).padStart(2, '0');
        const accepted = dropCandidates && dropCandidates.length > 0
          ? dropCandidates
          : getArrayCandidates(row, col, length, count, dir);
        const actualCount = accepted.length;
        const newFields = [];
        for (let i = 0; i < actualCount; i++) {
          const candidate = accepted[i];
          newFields.push({
            id:         arrLabel(aid, i),
            row: candidate.row, col: candidate.col, length,
            initialText: '',
            prot:        false,
            unprot:      true,
            askip:       false,
            brightness:  'norm',
            numeric:     false,
            ic:          false,
            fset:        false,
            color:       'default',
            highlight:   'off',
            outline:     [],
            stopper:     true,
            groupId:     gid,
            isArray:     true,
            arrayId:     aid,
            arrayDir:    dir,
            arrayIndex:  i,
            arrayCols:   dir === 'h' ? actualCount : 1,
            arrayRows:   dir === 'v' ? actualCount : 1,
            colStep:     length + 1,
            rowStep:     1,
          });
        }
        if (newFields.length >= 1) {
          History.push();
          const startIdx = fields.length;
          newFields.forEach(f => fields.push(f));
          selectedIds.clear();
          for (let i = 0; i < newFields.length; i++) selectedIds.add(startIdx + i);
          FieldRenderer.applyAll(mapDef, fields, cells, fillMode);
          updatePanelAndOverlay(mapDef, fields, cells);
        } else if (mapDef) {
          SelectionOverlay.resize(mapDef);
          SelectionOverlay.draw(mapDef, selectedIds, fields);
        }
      } else {
        const isProt = ftype === 'label';
        const isOutline    = ftype === 'outline';
        const isOutlineVar = ftype === 'outline-var';
        const newField = {
          id:          genId(ftype),
          row,
          col,
          length,
          initialText: (ftype === 'label' || isOutline) ? 'Label' : '',
          prot:        ftype === 'label' || isOutline || isOutlineVar,
          unprot:      ftype === 'inputtext' || ftype === 'inputnum' || ftype === 'password' || isOutlineVar,
          askip:       ftype === 'label' || isOutline,
          brightness:  ftype === 'password' ? 'drk' : 'norm',
          numeric:     ftype === 'inputnum',
          ic:          false,
          fset:        false,
          color:       'default',
          highlight:   'off',
          outline:     (isOutline || isOutlineVar) ? ['OVER','UNDER','LEFT','RIGHT'] : [],
          stopper:     ftype === 'inputtext' || ftype === 'inputnum' || ftype === 'password' || isOutlineVar,
        };
        History.push();
        fields.push(newField);
        FieldRenderer.applyAll(mapDef, fields, cells, fillMode);
        selectedIds.clear();
        selectedIds.add(fields.length - 1);
        updatePanelAndOverlay(mapDef, fields, cells);
      }
    } else if (mapDef) {
      // Cancelled or colliding — wipe the drop-preview without adding anything
      SelectionOverlay.resize(mapDef);
      SelectionOverlay.draw(mapDef, selectedIds, fields);
    }

    dropCandidates = null;

    dragItem   = null;
    dropCoords = null;
  }

  return { init };
})();

/* ===========================================================
   BOOT
   =========================================================== */
document.addEventListener('DOMContentLoaded', () => {
  const renderView    = document.getElementById('renderView');
  const sourceView    = document.getElementById('sourceView');
  const renderBtn     = document.getElementById('renderBtn');
  const sourceBtn     = document.getElementById('sourceBtn');
  const fillSelect    = document.getElementById('fillSelect');
  const renderControls = document.getElementById('renderControls');
  const hlSel         = document.getElementById('panelHighlight');
  const rowInput      = document.getElementById('panelRow');
  const colInput      = document.getElementById('panelCol');
  const lenInput      = document.getElementById('panelLen');
  const autoSyncChk   = document.getElementById('autoSyncChk');
  const autoResizeChk = document.getElementById('autoResizeChk');

  // Apply persisted config before first render
  const initialConfig = getInitialConfig();
  if (typeof initialConfig.fill === 'string' && FILL_TYPES[initialConfig.fill]) {
    fillMode = initialConfig.fill;
    fillSelect.value = initialConfig.fill;
  }
  if (typeof initialConfig.theme === 'string' && GRID_THEMES[initialConfig.theme]) {
    gridTheme = initialConfig.theme;
  }
  if (typeof initialConfig.sync === 'boolean') {
    autoSyncChk.checked = initialConfig.sync;
  }
  if (typeof initialConfig.autoResize === 'boolean') {
    autoResize = initialConfig.autoResize;
    autoResizeChk.checked = initialConfig.autoResize;
  }

  sourceView.textContent = bmsSource;

  /* -- view toggle -- */
  renderBtn.addEventListener('click', () => {
    renderView.style.display    = 'block';
    sourceView.style.display    = 'none';
    renderControls.style.display = 'flex';
    renderBtn.classList.add('active');
    sourceBtn.classList.remove('active');
  });

  sourceBtn.addEventListener('click', () => {
    renderView.style.display    = 'none';
    sourceView.style.display    = 'block';
    renderControls.style.display = 'none';
    sourceBtn.classList.add('active');
    renderBtn.classList.remove('active');
  });

  /* -- save button -- */
  function doSave() {
    const vsc = getVsCodeApi();
    if (vsc) vsc.postMessage({ command: 'saveBms', content: generateBmsSource() });
  }
  document.getElementById('saveBtn').addEventListener('click', doSave);

  /* -- auto-sync: renderer→file, debounced 500 ms after any field mutation -- */
  // _syncFromFile is declared at module level
  let _syncDebounce = null;

  function scheduleSyncToFile() {
    if (!document.getElementById('autoSyncChk').checked) return;
    if (_syncFromFile) return;
    clearTimeout(_syncDebounce);
    _syncDebounce = setTimeout(() => {
      const vsc = getVsCodeApi();
      if (vsc) vsc.postMessage({ command: 'syncBms', content: generateBmsSource() });
    }, 500);
  }

  /* -- panel: brightness dropdown -- */
  document.getElementById('panelBright').addEventListener('change', e => {
    if (!e.target.value) return;
    History.push();
    selectedIds.forEach(idx => { fields[idx].brightness = e.target.value; });
    FieldRenderer.applyAll(mapDef, fields, cells, fillMode);
    SelectionOverlay.draw(mapDef, selectedIds, fields);
  });

  /* -- panel: type dropdown (UNPROT / PROT / ASKIP) -- */
  document.getElementById('panelType').addEventListener('change', e => {
    if (!e.target.value || selectedIds.size === 0) return;
    History.push();
    const val = e.target.value;
    selectedIds.forEach(idx => {
      const f = fields[idx];
      f.askip  = val === 'askip';
      f.prot   = val === 'prot' || val === 'askip';
      f.unprot = val === 'unprot';
      // Clear UNPROT-only attributes when switching away from UNPROT
      if (val !== 'unprot') { f.numeric = false; f.ic = false; f.fset = false; }
      // Clear initial text when switching away from ASKIP (labels only)
      if (val !== 'askip') f.initialText = '';
    });
    FieldRenderer.applyAll(mapDef, fields, cells, fillMode);
    SelectionOverlay.draw(mapDef, selectedIds, fields);
    updatePanel(fields);
  });

  /* -- panel: attribute checkboxes -- */
  function wireAttrCheck(id, prop) {
    document.getElementById(id).addEventListener('change', e => {
      if (selectedIds.size === 0) return;
      History.push();
      selectedIds.forEach(idx => { fields[idx][prop] = e.target.checked; });
      FieldRenderer.applyAll(mapDef, fields, cells, fillMode);
      SelectionOverlay.draw(mapDef, selectedIds, fields);
    });
  }
  wireAttrCheck('panelNum',     'numeric');
  wireAttrCheck('panelIc',      'ic');
  wireAttrCheck('panelFset',    'fset');
  // panelStopper: for arrays, propagate to the entire array so all members stay in sync
  document.getElementById('panelStopper').addEventListener('change', e => {
    if (selectedIds.size === 0) return;
    History.push();
    const selArr = [...selectedIds].map(i => fields[i]).filter(Boolean);
    const isSameArray = selArr.every(f => f.isArray) && new Set(selArr.map(f => f.arrayId)).size === 1;
    if (isSameArray) {
      const arrayId = selArr[0].arrayId;
      fields.forEach(f => { if (f.isArray && f.arrayId === arrayId) f.stopper = e.target.checked; });
    } else {
      selectedIds.forEach(idx => { fields[idx].stopper = e.target.checked; });
    }
    FieldRenderer.applyAll(mapDef, fields, cells, fillMode);
    SelectionOverlay.draw(mapDef, selectedIds, fields);
  });

  /* -- panel: outline side checkboxes -- */
  ['Over','Under','Left','Right'].forEach(side => {
    const el = document.getElementById('panelOl' + side);
    if (!el) return;
    el.addEventListener('change', ev => {
      if (selectedIds.size === 0) return;
      History.push();
      const sideUp = side.toUpperCase();
      selectedIds.forEach(idx => {
        const f = fields[idx];
        if (!f.outline) f.outline = [];
        if (ev.target.checked) {
          if (!f.outline.includes(sideUp)) f.outline.push(sideUp);
        } else {
          f.outline = f.outline.filter(s => s !== sideUp);
        }
      });
      FieldRenderer.applyAll(mapDef, fields, cells, fillMode);
      SelectionOverlay.draw(mapDef, selectedIds, fields);
      updatePanel(fields);
    });
  });

  /* -- panel text (label initial text) -- */
  let _panelTextPreEditState = null;

  // Capture state when the user focuses the text input so we can push one clean
  // undo entry (covering both the text change and any auto-resize) when done.
  document.getElementById('panelText').addEventListener('focus', () => {
    _panelTextPreEditState = JSON.parse(JSON.stringify(fields));
  });

  // Auto-resize: grow the label field as the user types (if enabled and no collision).
  // NOTE: we must NOT call updatePanel() here because it resets textInput.value
  // to f.initialText (the not-yet-committed old value), clobbering what the user typed.
  // Instead we only update the Len display directly.
  document.getElementById('panelText').addEventListener('input', e => {
    if (!autoResize || selectedIds.size !== 1) return;
    const idx = [...selectedIds][0];
    const f   = fields[idx];
    if (!f.askip) return;
    const txt = e.target.value;
    if (txt.length <= f.length) return;
    const newLen = txt.length;
    if (newLen <= mapDef.cols - f.col &&
        !fieldCollides({ row: f.row, col: f.col, length: newLen }, fields, new Set([idx]))) {
      f.length = newLen;
      // Update only the Len counter in the panel — do NOT call updatePanel() which
      // would reset the text input to the stale f.initialText.
      const lenEl = document.getElementById('panelLen');
      if (lenEl) lenEl.value = String(f.length);
      FieldRenderer.applyAll(mapDef, fields, cells, fillMode);
      SelectionOverlay.resize(mapDef);
      SelectionOverlay.draw(mapDef, selectedIds, fields);
    } else {
      // No room to grow — truncate input to the current field length
      e.target.value = txt.slice(0, f.length);
    }
  });

  document.getElementById('panelText').addEventListener('change', e => {
    if (selectedIds.size !== 1) return;
    const f = fields[[...selectedIds][0]];
    if (!f.prot) return;
    // Use pre-edit snapshot so undo reverts text + any auto-resize in one step
    if (_panelTextPreEditState) {
      History.pushRaw(_panelTextPreEditState);
      _panelTextPreEditState = null;
    } else {
      History.push();
    }
    f.initialText = e.target.value;
    FieldRenderer.applyAll(mapDef, fields, cells, fillMode);
    SelectionOverlay.resize(mapDef);
    SelectionOverlay.draw(mapDef, selectedIds, fields);
    updatePanel(fields);
  });

  /* -- fill change -- */
  fillSelect.addEventListener('change', () => {
    fillMode = fillSelect.value;
    FieldRenderer.applyAll(mapDef, fields, cells, fillMode);
    SelectionOverlay.draw(mapDef, selectedIds, fields);
    persistConfig();
  });

  /* -- delete button + Delete key shortcut + Ctrl+Z/Y undo/redo -- */
  document.getElementById('deleteBtn').addEventListener('click', deleteSelected);

  /* -- right-click context menu -- */
  const ctxMenu    = document.getElementById('ctxMenu');
  const ctxGoToDef = document.getElementById('ctxGoToDef');
  const ctxDelete  = document.getElementById('ctxDelete');
  const ctxGroup   = document.getElementById('ctxGroup');
  const ctxUngroup = document.getElementById('ctxUngroup');
  const ctxGroupSep = document.getElementById('ctxGroupSep');
  let ctxFieldIdx  = -1;

  // Returns true if all fields in the current selection already share one groupId
  function _selectionIsOneGroup() {
    if (selectedIds.size < 2) return false;
    const gids = new Set([...selectedIds].map(i => fields[i].groupId).filter(Boolean));
    if (gids.size !== 1) return false;
    const gid = [...gids][0];
    // All selected fields must carry that groupId
    return [...selectedIds].every(i => fields[i].groupId === gid);
  }

  function _selectionHasGroup() {
    return [...selectedIds].some(i => fields[i].groupId);
  }

  function _selectionHasArray() {
    return [...selectedIds].some(i => fields[i] && fields[i].isArray);
  }

  function showCtxMenu(x, y, fieldIdx) {
    ctxFieldIdx = fieldIdx;
    const hasField = fieldIdx >= 0;
    ctxGoToDef.classList.toggle('ctx-disabled', !hasField);

    const hasArray   = _selectionHasArray();
    // Group: 2+ fields, not already one group, and no array members mixed in
    const canGroup   = selectedIds.size >= 2 && !_selectionIsOneGroup() && !hasArray;
    const canUngroup = _selectionHasGroup();
    ctxGroup.style.display    = canGroup   ? '' : 'none';
    // Ungroup is shown but grayed out for arrays
    ctxUngroup.style.display  = canUngroup ? '' : 'none';
    ctxUngroup.classList.toggle('ctx-disabled', hasArray);
    ctxGroupSep.style.display = (canGroup || canUngroup) ? '' : 'none';

    ctxMenu.style.display = 'block';
    // Keep within viewport
    const mw = ctxMenu.offsetWidth  || 160;
    const mh = ctxMenu.offsetHeight || 60;
    ctxMenu.style.left = Math.min(x, window.innerWidth  - mw - 6) + 'px';
    ctxMenu.style.top  = Math.min(y, window.innerHeight - mh - 6) + 'px';
  }

  function hideCtxMenu() {
    ctxMenu.style.display = 'none';
    ctxFieldIdx = -1;
  }

  document.getElementById('bmsGrid').addEventListener('contextmenu', e => {
    if (!mapDef) return;
    e.preventDefault();
    const { col, row } = Interaction.getCellCoords(e, mapDef);
    const idx = Interaction.getFieldAtCell(fields, row, col);
    if (idx >= 0) {
      // Select the right-clicked field if not already in the selection
      if (!selectedIds.has(idx)) {
        selectedIds.clear();
        selectedIds.add(idx);
        updatePanelAndOverlay(mapDef, fields, cells);
      }
      showCtxMenu(e.clientX, e.clientY, idx);
    } else if (selectedIds.size > 0) {
      // Empty space but something was selected — offer Delete
      showCtxMenu(e.clientX, e.clientY, -1);
    }
  });

  ctxGoToDef.addEventListener('click', e => {
    e.stopPropagation();
    if (ctxFieldIdx < 0 || ctxFieldIdx >= fields.length) { hideCtxMenu(); return; }
    const fieldId = fields[ctxFieldIdx].id; // capture before hideCtxMenu clears ctxFieldIdx
    hideCtxMenu();
    const vsc = getVsCodeApi();
    if (vsc) vsc.postMessage({ command: 'revealField', fieldId });
  });

  ctxDelete.addEventListener('click', e => {
    e.stopPropagation();
    hideCtxMenu();
    if (selectedIds.size > 0) deleteSelected();
  });

  ctxGroup.addEventListener('click', e => {
    e.stopPropagation();
    hideCtxMenu();
    groupSelected();
  });

  ctxUngroup.addEventListener('click', e => {
    e.stopPropagation();
    hideCtxMenu();
    ungroupSelected();
  });

  // Dismiss context menu on outside click or Escape
  ctxMenu.addEventListener('click', e => e.stopPropagation());
  document.addEventListener('click', () => hideCtxMenu());

  document.addEventListener('keydown', e => {
    if (e.key === 'Escape') { hideCtxMenu(); return; }
    if ((e.ctrlKey || e.metaKey) && !e.shiftKey &&
        (e.key === 'z' || e.key === 'Z') && !InlineEditor.isActive()) {
      e.preventDefault(); History.undo(); return;
    }
    if ((e.ctrlKey || e.metaKey) &&
        (e.key === 'y' || e.key === 'Y') && !InlineEditor.isActive()) {
      e.preventDefault(); History.redo(); return;
    }
    // Copy selected fields
    if ((e.ctrlKey || e.metaKey) && (e.key === 'c' || e.key === 'C') &&
        selectedIds.size > 0 && !InlineEditor.isActive()) {
      const tag = document.activeElement && document.activeElement.tagName;
      if (tag === 'INPUT' || tag === 'SELECT' || tag === 'TEXTAREA') return;
      e.preventDefault(); copySelected(); return;
    }
    // Paste copied fields
    if ((e.ctrlKey || e.metaKey) && (e.key === 'v' || e.key === 'V') &&
        clipboard && !InlineEditor.isActive()) {
      const tag = document.activeElement && document.activeElement.tagName;
      if (tag === 'INPUT' || tag === 'SELECT' || tag === 'TEXTAREA') return;
      e.preventDefault(); pasteFields(); return;
    }
    // Delete/Backspace: only when no panel input is focused
    if ((e.key === 'Delete' || e.key === 'Backspace') &&
        !InlineEditor.isActive()) {
      const tag = document.activeElement && document.activeElement.tagName;
      if (tag === 'INPUT' || tag === 'SELECT' || tag === 'TEXTAREA') return;
      // Delete a selected stopper
      if (selectedStopperId !== null) {
        e.preventDefault();
        History.push();
        const _sf = fields[selectedStopperId];
        if (_sf && _sf.isArray) {
          // Propagate stopper removal to all members of the same array
          fields.forEach(f => { if (f.isArray && f.arrayId === _sf.arrayId) f.stopper = false; });
        } else if (_sf) {
          _sf.stopper = false;
        }
        selectedStopperId = null;
        FieldRenderer.applyAll(mapDef, fields, cells, fillMode);
        updatePanelAndOverlay(mapDef, fields, cells);
        return;
      }
      if (selectedIds.size > 0) {
        e.preventDefault();
        deleteSelected();
      }
    }
  });

  /* -- panel: ID rename -- */
  const panelIdInput = document.getElementById('panelId');
  panelIdInput.addEventListener('input', e => {
    const sanitized = sanitizeFieldId(e.target.value);
    if (e.target.value !== sanitized) {
      e.target.value = sanitized;
    }
  });

  document.getElementById('panelId').addEventListener('change', e => {
    const selArr = [...selectedIds].map(i => fields[i]);
    const isSameArray = selArr.length >= 1 &&
                        selArr.every(f => f.isArray) &&
                        new Set(selArr.map(f => f.arrayId)).size === 1;

    if (isSameArray) {
      // Rename the array: update arrayId on ALL fields with that arrayId and regenerate numbered labels
      const oldId = selArr[0].arrayId;
      const val   = sanitizeFieldId(e.target.value.trim().toUpperCase());
      if (!val) { e.target.value = oldId; return; }
      History.push();
      fields.forEach(f => {
        if (f.isArray && f.arrayId === oldId) {
          f.arrayId = val;
          f.id      = arrLabel(val, f.arrayIndex);
        }
      });
      e.target.value = val;
      FieldRenderer.applyAll(mapDef, fields, cells, fillMode);
      updatePanelAndOverlay(mapDef, fields, cells);
      return;
    }

    if (selectedIds.size !== 1) return;
    const val = sanitizeFieldId(e.target.value.trim());
    if (val) {
      History.push();
      fields[[...selectedIds][0]].id = val;
      e.target.value = val;
      FieldRenderer.applyAll(mapDef, fields, cells, fillMode);
      SelectionOverlay.draw(mapDef, selectedIds, fields);
    } else {
      e.target.value = fields[[...selectedIds][0]].id; // reject empty
    }
  });

  /* -- color picker -- */
  const cpTrigger = document.getElementById('colorPickerTrigger');
  const cpMenu    = document.getElementById('colorPickerMenu');

  cpTrigger.addEventListener('click', e => {
    e.stopPropagation();
    cpMenu.classList.toggle('open');
  });

  cpMenu.addEventListener('click', e => {
    e.stopPropagation();
    const opt = e.target.closest('.cp-option');
    if (!opt) return;
    cpMenu.classList.remove('open');
    setColorPickerValue(opt.dataset.value);
    History.push();
    selectedIds.forEach(idx => {
      fields[idx].color = opt.dataset.value;
      FieldRenderer.refreshField(mapDef, cells, fields[idx], idx, fillMode);
    });
    SelectionOverlay.draw(mapDef, selectedIds, fields);
  });

  document.addEventListener('click', () => cpMenu.classList.remove('open'));

  /* -- panel: highlight change -- */
  hlSel.addEventListener('change', () => {
    History.push();
    selectedIds.forEach(idx => {
      fields[idx].highlight = hlSel.value;
      FieldRenderer.refreshField(mapDef, cells, fields[idx], idx, fillMode);
    });
    SelectionOverlay.draw(mapDef, selectedIds, fields);
  });

  /* -- panel: row/col/len change reflects on render -- */
  function applyPositionFromPanel() {
    // Array selection: shift all members by position delta; resize all lengths together
    const _pSelArr = [...selectedIds].map(i => fields[i]);
    const _pIsArray = _pSelArr.length >= 1 &&
                      _pSelArr.every(f => f && f.isArray) &&
                      new Set(_pSelArr.map(f => f.arrayId)).size === 1;
    if (_pIsArray) {
      const arrayId      = _pSelArr[0].arrayId;
      const memberIdxs   = fields.map((f, i) => i).filter(i => fields[i].isArray && fields[i].arrayId === arrayId);
      const firstMIdx    = memberIdxs.find(i => fields[i].arrayIndex === 0) ?? memberIdxs[0];
      const firstF       = fields[firstMIdx];
      const r = parseInt(rowInput.value, 10);
      const c = parseInt(colInput.value, 10);
      const l = parseInt(lenInput.value, 10);
      const newFirstRow = (!isNaN(r) && r >= 1 && r <= mapDef.rows) ? r - 1 : firstF.row;
      const newFirstCol = (!isNaN(c) && c >= 1 && c <= mapDef.cols) ? c - 1 : firstF.col;
      let   newLen      = (!isNaN(l) && l >= 1)                     ? l     : firstF.length;
      const arrayCols   = firstF.arrayCols ?? (firstF.arrayDir === 'v' ? 1 : memberIdxs.length);
      let   colStep     = firstF.colStep   ?? (firstF.length + 1);
      const rowStep     = firstF.rowStep   ?? 1;
      // Auto-expand colStep if length would overlap the next element
      if (newLen >= colStep) colStep = newLen + 1;
      const memberIdxSet = new Set(memberIdxs);
      const newPos = memberIdxs.map(i => {
        const f  = fields[i];
        const ri = Math.floor(f.arrayIndex / arrayCols);
        const ci = f.arrayIndex % arrayCols;
        return { row: newFirstRow + ri * rowStep, col: newFirstCol + ci * colStep, length: newLen };
      });
      const valid = newPos.every(p =>
        p.row >= 0 && p.row < mapDef.rows &&
        p.col >= 0 && p.col + p.length <= mapDef.cols &&
        !fieldCollides(p, fields, memberIdxSet)
      );
      if (!valid) {
        rowInput.value = String(firstF.row + 1);
        colInput.value = String(firstF.col + 1);
        lenInput.value = String(firstF.length);
        return;
      }
      History.push();
      memberIdxs.forEach((fi, mi) => {
        fields[fi].row     = newPos[mi].row;
        fields[fi].col     = newPos[mi].col;
        fields[fi].length  = newLen;
        fields[fi].colStep = colStep;
      });
      FieldRenderer.applyAll(mapDef, fields, cells, fillMode);
      SelectionOverlay.resize(mapDef);
      SelectionOverlay.draw(mapDef, selectedIds, fields);
      updatePanel(fields);
      return;
    }

    if (selectedIds.size !== 1) return;
    const idx = [...selectedIds][0];
    const f   = fields[idx];
    const r   = parseInt(rowInput.value, 10);
    const c   = parseInt(colInput.value, 10);
    const l   = parseInt(lenInput.value, 10);

    // Compute candidate values, falling back to current field values if out of range
    const newRow = (!isNaN(r) && r >= 1 && r <= mapDef.rows)          ? r - 1 : f.row;
    const newCol = (!isNaN(c) && c >= 1 && c <= mapDef.cols)          ? c - 1 : f.col;
    const newLen = (!isNaN(l) && l >= 1 && l <= mapDef.cols - newCol) ? l     : f.length;

    // Reject if the new position/size would overlap another field
    if (fieldCollides({ row: newRow, col: newCol, length: newLen }, fields, new Set([idx]))) {
      rowInput.value = String(f.row + 1);
      colInput.value = String(f.col + 1);
      lenInput.value = String(f.length);
      return;
    }

    History.push();
    f.row    = newRow;
    f.col    = newCol;
    f.length = newLen;
    FieldRenderer.applyAll(mapDef, fields, cells, fillMode);
    SelectionOverlay.resize(mapDef);
    SelectionOverlay.draw(mapDef, selectedIds, fields);
  }
  rowInput.addEventListener('change', applyPositionFromPanel);
  colInput.addEventListener('change', applyPositionFromPanel);
  lenInput.addEventListener('change', applyPositionFromPanel);

  /* -- panel: array structure (cols, rows, H step, V step) -- */
  function applyArrayConfig() {
    const _ac = [...selectedIds].map(i => fields[i]);
    if (!(_ac.length >= 1 && _ac.every(f => f && f.isArray) &&
          new Set(_ac.map(f => f.arrayId)).size === 1) || !mapDef) return;
    const arrayId = _ac[0].arrayId;
    const firstF  = _ac.find(f => f.arrayIndex === 0) || _ac[0];
    const newCols    = Math.max(1, parseInt(document.getElementById('panelArrayCols').value,    10) || 1);
    const newRows    = Math.max(1, parseInt(document.getElementById('panelArrayRows').value,    10) || 1);
    const hGap = Math.max(0, parseInt(document.getElementById('panelArrayColStep').value, 10) || 0);
    const newColStep = firstF.length + hGap; // step = length + gap
    const vGap = Math.max(0, parseInt(document.getElementById('panelArrayRowStep').value, 10) || 0);
    const newRowStep = 1 + vGap; // step = 1 + gap
    if (!rebuildArray(arrayId, newCols, newRows, newColStep, newRowStep, firstF.length)) {
      updatePanel(fields); // restore inputs to actual values on failure
      return;
    }
    FieldRenderer.applyAll(mapDef, fields, cells, fillMode);
    updatePanelAndOverlay(mapDef, fields, cells);
  }
  document.getElementById('panelArrayCols').addEventListener('change',    applyArrayConfig);
  document.getElementById('panelArrayRows').addEventListener('change',    applyArrayConfig);
  document.getElementById('panelArrayColStep').addEventListener('change', applyArrayConfig);
  document.getElementById('panelArrayRowStep').addEventListener('change', applyArrayConfig);

  /* -- theme toggle -- */
  document.getElementById('themeBtn').addEventListener('click', () => {
    gridTheme = gridTheme === 'dark' ? 'light' : 'dark';
    applyGridTheme();
    FieldRenderer.applyAll(mapDef, fields, cells, fillMode);
    SelectionOverlay.draw(mapDef, selectedIds, fields);
    persistConfig();
  });

  /* -- refit grid on container resize -- */
  const resizeObserver = new ResizeObserver(() => {
    fitGrid(mapDef);
  });
  resizeObserver.observe(document.getElementById('canvasArea'));

  /* -- grid size inputs -- */
  function checkGridSizeWarn() {
    const r = parseInt(document.getElementById('gridRowsInput').value, 10) || 0;
    const c = parseInt(document.getElementById('gridColsInput').value, 10) || 0;
    setGridSizeWarn(r > 24 || c > 80);
  }

  document.getElementById('gridRowsInput').addEventListener('change', () => {
    if (mapDef) renderBms();
    checkGridSizeWarn();
  });
  document.getElementById('gridColsInput').addEventListener('change', () => {
    if (mapDef) renderBms();
    checkGridSizeWarn();
  });

  /* -- file→renderer: re-render when the source file is edited -- */
  window.addEventListener('message', event => {
    const msg = event.data;
    if (msg.command === 'updateSource') {
      _syncFromFile = true;
      bmsSource = msg.source;
      renderBms();
      setTimeout(() => { _syncFromFile = false; }, 0);
    }
  });

  /* -- renderer→file: wrap FieldRenderer so every mutation schedules a debounced sync -- */
  const _origApplyAll     = FieldRenderer.applyAll;
  const _origRefreshField = FieldRenderer.refreshField;
  FieldRenderer.applyAll     = function(...a) { _origApplyAll(...a);     scheduleSyncToFile(); };
  FieldRenderer.refreshField = function(...a) { _origRefreshField(...a); scheduleSyncToFile(); };

  /* -- auto-resize toggle -- */
  autoResizeChk.addEventListener('change', e => {
    autoResize = e.target.checked;
    persistConfig({ autoResize });
  });

  // Wire auto-sync toggle to persist changes too
  autoSyncChk.addEventListener('change', e => {
    persistConfig({ sync: e.target.checked });
  });

  /* -- click outside grid + field-panel: clear selection -- */
  document.addEventListener('mousedown', e => {
    if (!mapDef || (selectedIds.size === 0 && selectedStopperId === null)) return;
    const wrapper = document.getElementById('gridWrapper');
    const panel   = document.getElementById('fieldPanel');
    const ctxMenu = document.getElementById('ctxMenu');
    if (wrapper && wrapper.contains(e.target)) return;
    if (panel   && panel.contains(e.target))   return;
    if (ctxMenu && ctxMenu.contains(e.target)) return;
    // Commit any in-progress text edit before clearing
    const panelTextEl = document.getElementById('panelText');
    if (panelTextEl && document.activeElement === panelTextEl) {
      panelTextEl.dispatchEvent(new Event('change', { bubbles: true }));
    }
    selectedIds.clear();
    selectedStopperId = null;
    updatePanelAndOverlay(mapDef, fields, cells);
  });

  /* -- init overlay + render -- */
  SelectionOverlay.init();
  PaletteDrag.init();
  renderBms();
});
