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

let mapDef    = null;   // { rows, cols }
let fields    = [];     // parsed + augmented field objects
let cells     = [];     // flat array of cell DOM elements (row-major)

let selectedIds = new Set();  // field indices that are selected

let fillMode  = 'empty';

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
    // This correctly handles multi-line definitions where LENGTH= / POS= appear
    // on continuation lines (not just the first line of the macro statement).
    const labelRe = /^(\w{1,7})\s+DFHMDF\b/gm;
    const starts  = [];
    let m;
    while ((m = labelRe.exec(source)) !== null) {
      starts.push({ id: m[1], index: m.index });
    }

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
      });
    }
    return result;
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
  let dropPreview = null; // { row, col, length }
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
      const { row, col, length } = dropPreview;
      ctx.save();
      ctx.fillStyle   = isColliding ? 'rgba(220,50,50,0.25)'  : 'rgba(0,127,212,0.25)';
      ctx.strokeStyle = isColliding ? '#dd3333'               : '#007fd4';
      ctx.lineWidth   = 1;
      ctx.setLineDash([4, 3]);
      ctx.fillRect(col   * cellWpx, row * cellHpx, length * cellWpx, cellHpx);
      ctx.strokeRect(col * cellWpx, row * cellHpx, length * cellWpx, cellHpx);
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
        History.push();
        f.initialText = input.value;
      }
      input.remove();
      activeIdx = null;
      FieldRenderer.refreshField(mapDef, cells, f, idx, fillMode);
      SelectionOverlay.draw(mapDef, selectedIds, fields);
      updatePanel(fields);
    }

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
      default:       'default',
      field:         'pointer',
      move:          'grab',
      grabbing:      'grabbing',
      'resize-left': 'ew-resize',
      'resize-right':'ew-resize',
      lasso:         'crosshair',
      ctrlhover:     'alias',
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
    _hMouseleave = () => Cursor.set('default');
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

  function getResizeHandle(e, fieldIdx, fields, mapDef) {
    if (selectedIds.size !== 1) return null;
    const f = fields[fieldIdx];
    const grid  = document.getElementById('bmsGrid');
    const r     = grid.getBoundingClientRect();
    const cellW = r.width / mapDef.cols;
    const fx    = f.col * cellW + r.left;
    const fw    = f.length * cellW;
    const mx    = e.clientX;
    if (Math.abs(mx - fx)        < 8) return 'resize-left';
    if (Math.abs(mx - (fx + fw)) < 8) return 'resize-right';
    return null;
  }

  /* ---- hover cursor update ---- */
  function onGridHover(e, mapDef, fields) {
    if (isDragging || isLasso) return;
    const { col, row } = getCellCoords(e, mapDef);
    const idx = getFieldAtCell(fields, row, col);
    if (idx < 0) { Cursor.set('default'); return; }
    if (e.ctrlKey || e.metaKey) { Cursor.set('ctrlhover'); return; }
    const handle = getResizeHandle(e, idx, fields, mapDef);
    if (handle) { Cursor.set(handle); return; }
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

    const { col, row, px, py } = getCellCoords(e, mapDef);
    const idx = getFieldAtCell(fields, row, col);

    // Ctrl/Cmd + click: reveal in editor
    if ((e.ctrlKey || e.metaKey) && idx >= 0) {
      e.preventDefault();
      const f = fields[idx];
      const vscode = getVsCodeApi();
      if (vscode) vscode.postMessage({ command: 'revealField', fieldId: f.id });
      return;
    }

    // Clicked empty space — start lasso or clear selection
    if (idx < 0) {
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
      isDragging = true;
      dragInfo = {
        type: handle, fieldIdx: idx,
        startX: e.clientX, startY: e.clientY,
        origRow: fields[idx].row, origCol: fields[idx].col, origLen: fields[idx].length,
        origState: JSON.parse(JSON.stringify(fields)),
      };
      Cursor.set(handle);
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
      dragInfo.snapshot.forEach(snap => {
        const f = fields[snap.idx];
        f.row = Math.max(0, Math.min(mapDef.rows - 1, snap.row + dy));
        f.col = Math.max(0, Math.min(mapDef.cols - f.length, snap.col + dx));
      });
      dragInfo.colliding = anyCollision(
        dragInfo.snapshot.map(snap => fields[snap.idx]), fields, selectedIds
      );
    } else if (dragInfo.type === 'resize-right') {
      const f = fields[dragInfo.fieldIdx];
      f.length = Math.min(Math.max(1, dragInfo.origLen + dx), mapDef.cols - f.col);
      dragInfo.colliding = fieldCollides(f, fields, new Set([dragInfo.fieldIdx]));
    } else if (dragInfo.type === 'resize-left') {
      const f    = fields[dragInfo.fieldIdx];
      const newCol = Math.max(0, Math.min(dragInfo.origCol + dx, dragInfo.origCol + dragInfo.origLen - 1));
      const delta  = newCol - dragInfo.origCol;
      f.col    = newCol;
      f.length = Math.max(1, dragInfo.origLen - delta);
      dragInfo.colliding = fieldCollides(f, fields, new Set([dragInfo.fieldIdx]));
    }

    SelectionOverlay.setColliding(dragInfo.colliding || false);
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

    if (wasColliding && info) {
      // Revert to the positions that were recorded at drag start
      if (info.type === 'move') {
        info.snapshot.forEach(snap => {
          fields[snap.idx].row = snap.row;
          fields[snap.idx].col = snap.col;
        });
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

  if (selectedIds.size === 0) {
    panel.style.display = 'none';
    return;
  }

  panel.style.display = 'flex';
  const multi    = selectedIds.size > 1;
  const firstIdx = [...selectedIds][0];
  const first    = fields[firstIdx];

  idInput.disabled  = multi;  // editable when a single field is selected
  rowInput.disabled = multi;
  colInput.disabled = multi;
  lenInput.disabled = multi;

  idInput.value  = multi ? '(multiple)' : first.id;
  rowInput.value = multi ? '' : first.row + 1;
  colInput.value = multi ? '' : first.col + 1;
  lenInput.value = multi ? '' : first.length;

  // Text field: only for single-field ASKIP (label) selection
  const showText = !multi && first.askip;
  textGroup.style.display = showText ? 'flex' : 'none';
  if (showText) {
    textInput.disabled = false;
    textInput.value    = first.initialText || '';
  }

  // Type dropdown
  const selArr = [...selectedIds].map(i => fields[i]);
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
   GENERATE BMS SOURCE
   Builds a simplified BMS assembler listing from current fields.
   =========================================================== */
function generateBmsSource() {
  if (!mapDef) return '';

  // Wraps a long BMS macro line using continuation character X at col 72
  function bmsLines(label, keyword, attrs) {
    const COL = 71; // 0-indexed column of continuation char (col 72)
    const lines = [];
    // First line: "LABEL    KEYWORD attr1,"
    let cur = `${label.padEnd(8)} ${keyword} `;
    for (let i = 0; i < attrs.length; i++) {
      const part = attrs[i] + (i < attrs.length - 1 ? ',' : '');
      const candidate = cur + part;
      if (candidate.length <= COL) {
        cur = candidate;
      } else {
        // Flush current with continuation X
        lines.push(cur.padEnd(COL) + 'X');
        cur = '               ' + part; // 15-space indent for continuation
      }
    }
    lines.push(cur); // last line has no continuation
    return lines.join('\n');
  }

  const out = [];
  out.push(bmsLines('MAPSET', 'DFHMSD', ['TYPE=MAP', 'CTRL=FREEKB']));
  out.push(bmsLines('MAPNAME', 'DFHMDI',
    [`SIZE=(${mapDef.rows},${mapDef.cols})`, 'LINE=1', 'COLUMN=1']));

  fields.forEach(f => {
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

    // INITIAL
    if (f.initialText) attrs.push(`INITIAL='${f.initialText}'`);

    out.push(bmsLines(f.id, 'DFHMDF', attrs));
  });

  out.push(bmsLines('', 'DFHMSD', ['TYPE=FINAL']));
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
  let dropColliding = false;
  let ghost         = null;
  let moveHandler   = null;
  let upHandler     = null;
  let fieldCounter  = 0;

  // Representative ghost text for each type + length
  const GHOST_TEXT = {
    label:     { 1: 'A', 8: 'Label   ', 16: 'Long Label      ' },
    inputtext: { 1: '_', 8: '________', 16: '________________' },
    inputnum:  { 1: '0', 8: '00000000', 16: '0000000000000000' },
  };

  function genId(ftype) {
    return ftype.slice(0, 3).toUpperCase() + String(++fieldCounter).padStart(4, '0');
  }

  function init() {
    ghost = document.getElementById('paletteDragGhost');
    document.querySelectorAll('.palette-item[data-ftype]').forEach(el => {
      el.addEventListener('mousedown', e => start(e, el));
    });
  }

  function start(e, el) {
    if (e.button !== 0) return;
    const ftype = el.dataset.ftype;
    if (ftype === 'array') return;   // not yet implemented
    e.preventDefault();

    dragItem   = { ftype, length: parseInt(el.dataset.length, 10) };
    dropCoords = null;

    // Clear grid selection so existing fields don't show collision state
    selectedIds.clear();
    SelectionOverlay.setColliding(false);

    const gt = (GHOST_TEXT[ftype] || {})[dragItem.length] || ftype + ':' + dragItem.length;
    ghost.textContent = gt;
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
      dropCoords    = { row, col };
      dropColliding = fieldCollides({ row, col, length: dragItem.length }, fields, null);
      SelectionOverlay.setColliding(dropColliding);
      SelectionOverlay.setDropPreview({ row, col, length: dragItem.length });
    } else {
      dropCoords    = null;
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
      const { ftype, length } = dragItem;
      const isProt = ftype === 'label';
      const newField = {
        id:          genId(ftype),
        row,
        col,
        length,
        initialText: ftype === 'label' ? 'Label' : '',
        prot:        ftype === 'label',
        unprot:      ftype === 'inputtext' || ftype === 'inputnum' || ftype === 'password',
        askip:       ftype === 'label',
        brightness:  ftype === 'password' ? 'drk' : 'norm',
        numeric:     ftype === 'inputnum',
        ic:          false,
        fset:        false,
        color:       'default',
        highlight:   'off',
      };
      History.push();
      fields.push(newField);
      FieldRenderer.applyAll(mapDef, fields, cells, fillMode);
      selectedIds.clear();
      selectedIds.add(fields.length - 1);
      updatePanelAndOverlay(mapDef, fields, cells);
    } else if (mapDef) {
      // Cancelled or colliding — wipe the drop-preview without adding anything
      SelectionOverlay.resize(mapDef);
      SelectionOverlay.draw(mapDef, selectedIds, fields);
    }

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
  wireAttrCheck('panelNum',  'numeric');
  wireAttrCheck('panelIc',   'ic');
  wireAttrCheck('panelFset', 'fset');

  /* -- panel text (label initial text) -- */
  document.getElementById('panelText').addEventListener('change', e => {
    if (selectedIds.size !== 1) return;
    const f = fields[[...selectedIds][0]];
    if (!f.prot) return;
    History.push();
    f.initialText = e.target.value;
    FieldRenderer.applyAll(mapDef, fields, cells, fillMode);
    SelectionOverlay.resize(mapDef);
    SelectionOverlay.draw(mapDef, selectedIds, fields);
  });

  /* -- fill change -- */
  fillSelect.addEventListener('change', () => {
    fillMode = fillSelect.value;
    FieldRenderer.applyAll(mapDef, fields, cells, fillMode);
    SelectionOverlay.draw(mapDef, selectedIds, fields);
  });

  /* -- delete button + Delete key shortcut + Ctrl+Z/Y undo/redo -- */
  document.getElementById('deleteBtn').addEventListener('click', deleteSelected);

  /* -- right-click context menu -- */
  const ctxMenu    = document.getElementById('ctxMenu');
  const ctxGoToDef = document.getElementById('ctxGoToDef');
  const ctxDelete  = document.getElementById('ctxDelete');
  let ctxFieldIdx  = -1;

  function showCtxMenu(x, y, fieldIdx) {
    ctxFieldIdx = fieldIdx;
    const hasField = fieldIdx >= 0;
    ctxGoToDef.classList.toggle('ctx-disabled', !hasField);
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
    if ((e.key === 'Delete' || e.key === 'Backspace') &&
        selectedIds.size > 0 && !InlineEditor.isActive()) {
      e.preventDefault();
      deleteSelected();
    }
  });

  /* -- panel: ID rename (single selection only, max 20 chars) -- */
  document.getElementById('panelId').addEventListener('change', e => {
    if (selectedIds.size !== 1) return;
    const val = e.target.value.trim().slice(0, 20);
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

  /* -- theme toggle -- */
  document.getElementById('themeBtn').addEventListener('click', () => {
    gridTheme = gridTheme === 'dark' ? 'light' : 'dark';
    applyGridTheme();
    FieldRenderer.applyAll(mapDef, fields, cells, fillMode);
    SelectionOverlay.draw(mapDef, selectedIds, fields);
  });

  /* -- refit grid on container resize -- */
  const resizeObserver = new ResizeObserver(() => {
    fitGrid(mapDef);
  });
  resizeObserver.observe(document.getElementById('canvasArea'));

  /* -- grid size inputs -- */
  document.getElementById('gridRowsInput').addEventListener('change', () => {
    if (mapDef) renderBms();
  });
  document.getElementById('gridColsInput').addEventListener('change', () => {
    if (mapDef) renderBms();
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

  /* -- init overlay + render -- */
  SelectionOverlay.init();
  PaletteDrag.init();
  renderBms();
});
