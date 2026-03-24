/**
 * =====================================================
 * MAESTRO FURNITURE — CUTLIST OPTIMIZER
 * app.js — Main application logic
 * =====================================================
 *
 * Algorithm: Guillotine 2D Bin Packing
 *   - Iteratively places pieces in the most suitable
 *     free rectangle (Best Short Side Fit / BSSF).
 *   - Splits remaining free space via guillotine cuts.
 *   - Optionally rotates pieces for better fit.
 *   - Repeats on new boards until all pieces are placed.
 * =====================================================
 */

/* ────────────────────────────────────────────────
   STATE
──────────────────────────────────────────────── */
const state = {
  groups: [],   // [{ thickness, boards[] }] — one group per unique thickness
  parts: [],
  partColors: {},
  stock: [],   // [{ width, height, qty }] — stock board inventory
};

/* ────────────────────────────────────────────────
   DOM REFS
──────────────────────────────────────────────── */
const $ = id => document.getElementById(id);
const partsBody = $('partsBody');

/* ────────────────────────────────────────────────
   OPTIONS
──────────────────────────────────────────────── */
function getBoardGrain() {
  const el = $('boardGrainToggle');
  return el ? (el.dataset.value || 'H') : 'H';
}

function getOptions() {
  return {
    showLabels: ($('optLabels')?.checked ?? true),
    labelSize: parseFloat($('optLabelSize')?.value ?? 100) / 100,
    oneSheet: ($('optOneSheet')?.checked ?? false),
    considerMat: ($('optConsiderMaterial')?.checked ?? false),
    edgeBanding: ($('optEdgeBanding')?.checked ?? false),
    grainDirection: ($('optGrain')?.checked ?? true),
    boardGrain: getBoardGrain(),
  };
}

/* ────────────────────────────────────────────────
   UTILITY
──────────────────────────────────────────────── */
let _partId = 0;
const newId = () => `p${++_partId}`;

function toast(msg, duration = 2600) {
  const el = $('toast');
  el.textContent = msg;
  el.classList.add('show');
  clearTimeout(el._timer);
  el._timer = setTimeout(() => el.classList.remove('show'), duration);
}

function fmtMm2(mm2) {
  if (mm2 >= 1e6) return (mm2 / 1e6).toFixed(3) + ' m²';
  return mm2.toLocaleString() + ' mm²';
}

/** Deterministic pastel colour from string */
function colorForId(id) {
  if (state.partColors[id]) return state.partColors[id];
  // Generate a pleasing hue from the id's numeric part
  const idx = parseInt(id.slice(1), 10);
  const hue = (idx * 47 + 120) % 360; // spread hues
  const s = 52, l = 72;
  state.partColors[id] = `hsl(${hue},${s}%,${l}%)`;
  return state.partColors[id];
}

function hexFromHsl(id) { return colorForId(id); }   // alias for clarity

/* ────────────────────────────────────────────────
   PARTS TABLE
──────────────────────────────────────────────── */
function addPartRow(data = {}) {
  const id = data.id || newId();
  const tr = document.createElement('tr');
  tr.dataset.partId = id;
  const rowNum = partsBody.children.length + 1;
  const defaultThick = $('boardThickness').value || '18';

  const grainVal = data.grain ?? 'H';
  tr.innerHTML = `
    <td class="td-num">${rowNum}</td>
    <td><input class="td-input" type="text"   data-field="name"      value="${data.name ?? ''}"  placeholder="e.g. Side Panel" /></td>
    <td><input class="td-input" type="number" data-field="width"     value="${data.width ?? ''}"  placeholder="mm" min="1" /></td>
    <td><input class="td-input" type="number" data-field="height"    value="${data.height ?? ''}"  placeholder="mm" min="1" /></td>
    <td><input class="td-input" type="number" data-field="thickness" value="${data.thickness ?? defaultThick}" placeholder="mm" min="1" style="width:52px;" /></td>
    <td><input class="td-input" type="number" data-field="qty"       value="${data.qty ?? 1}"   min="1" style="width:52px;" /></td>
    <td style="text-align:center;">
      <div class="grain-toggle" data-field="grain" data-value="${grainVal}">
        <button type="button" class="grain-btn${grainVal === 'H' ? ' active' : ''}" data-dir="H" title="Horizontal grain">&#x2194;</button>
        <button type="button" class="grain-btn${grainVal === 'V' ? ' active' : ''}" data-dir="V" title="Vertical grain">&#x2195;</button>
      </div>
    </td>
    <td><button class="btn btn-danger" title="Remove row">✕</button></td>
  `;

  tr.querySelector('.btn-danger').addEventListener('click', () => {
    tr.remove();
    renumberRows();
    updateCount();
  });

  // Grain toggle click handler
  const grainToggle = tr.querySelector('.grain-toggle');
  grainToggle.querySelectorAll('.grain-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      const dir = btn.dataset.dir;
      grainToggle.dataset.value = dir;
      grainToggle.querySelectorAll('.grain-btn').forEach(b => b.classList.toggle('active', b.dataset.dir === dir));
    });
  });

  partsBody.appendChild(tr);
  updateCount();
  return tr;
}

function renumberRows() {
  [...partsBody.querySelectorAll('tr')].forEach((tr, i) => {
    tr.querySelector('.td-num').textContent = i + 1;
  });
}

function updateCount() {
  const n = partsBody.children.length;
  $('tableCount').textContent = `${n} part${n !== 1 ? 's' : ''}`;
}

function collectParts() {
  const rows = [...partsBody.querySelectorAll('tr')];
  const parts = [];
  const defaultThick = parseFloat($('boardThickness').value) || 18;
  for (const tr of rows) {
    const get = f => tr.querySelector(`[data-field="${f}"]`);
    const w = parseFloat(get('width').value);
    const h = parseFloat(get('height').value);
    if (!w || !h || w <= 0 || h <= 0) {
      toast('⚠️ All parts must have valid width and height.');
      return null;
    }
    const grainEl = tr.querySelector('[data-field="grain"]');
    parts.push({
      id: tr.dataset.partId,
      name: get('name').value.trim(),
      width: w,
      height: h,
      thickness: parseFloat(get('thickness').value) || defaultThick,
      qty: Math.max(1, parseInt(get('qty').value) || 1),
      rotate: false,  // rotate removed — grain direction controls orientation
      grain: grainEl ? grainEl.dataset.value : 'H',
    });
  }
  return parts;
}

/* ────────────────────────────────────────────────
   GUILLOTINE BIN PACKING ALGORITHM
──────────────────────────────────────────────── */

/**
 * Represents a free rectangle within a board.
 * x,y: top-left corner (in mm)
 */
function makeFreeRect(x, y, w, h) { return { x, y, w, h }; }

/**
 * Score of placing (pw x ph) in free rect fr.
 * Returns { score, rotated } or null if it doesn't fit.
 * Strategy: Best Short Side Fit (BSSF)
 *
 * Grain logic: when grainDirection is ON, the panel's grain must align
 * with the board's grain. If panel grain ≠ board grain, force rotation.
 * If they match, force no rotation.
 */
function scorePlacement(fr, pw, ph, canRotate, opts, partGrain) {
  if (opts && opts.grainDirection && partGrain) {
    const boardGrain = opts.boardGrain || 'H';
    const mustRotate = (partGrain !== boardGrain);
    if (mustRotate) {
      // Force rotated orientation: swap pw ↔ ph
      const fits = ph <= fr.w && pw <= fr.h;
      if (!fits) return null;
      const shortSide = Math.min(fr.w - ph, fr.h - pw);
      const longSide = Math.max(fr.w - ph, fr.h - pw);
      return { score: shortSide, long: longSide, rotated: true };
    } else {
      // Force original orientation
      const fits = pw <= fr.w && ph <= fr.h;
      if (!fits) return null;
      const shortSide = Math.min(fr.w - pw, fr.h - ph);
      const longSide = Math.max(fr.w - pw, fr.h - ph);
      return { score: shortSide, long: longSide, rotated: false };
    }
  }

  const fits = pw <= fr.w && ph <= fr.h;
  const fitRot = canRotate && ph <= fr.w && pw <= fr.h;
  if (!fits && !fitRot) return null;

  let best = null;

  if (fits) {
    const shortSide = Math.min(fr.w - pw, fr.h - ph);
    const longSide = Math.max(fr.w - pw, fr.h - ph);
    if (!best || shortSide < best.score || (shortSide === best.score && longSide < best.long)) {
      best = { score: shortSide, long: longSide, rotated: false };
    }
  }
  if (fitRot && canRotate) {
    const shortSide = Math.min(fr.w - ph, fr.h - pw);
    const longSide = Math.max(fr.w - ph, fr.h - pw);
    if (!best || shortSide < best.score || (shortSide === best.score && longSide < best.long)) {
      best = { score: shortSide, long: longSide, rotated: true };
    }
  }
  return best;
}

/**
 * Split free rect after placing pw x ph at fr's top-left.
 * Uses Longer Axis Guillotine splitting.
 */
function splitFreeRect(fr, pw, ph) {
  const result = [];
  // Space to the right:
  if (fr.w - pw > 0) result.push(makeFreeRect(fr.x + pw, fr.y, fr.w - pw, fr.h));
  // Space below:
  if (fr.h - ph > 0) result.push(makeFreeRect(fr.x, fr.y + ph, pw, fr.h - ph));
  return result;
}

/**
 * Remove free rects fully contained within the placed rect.
 */
function pruneFreeRects(freeRects, x, y, pw, ph) {
  return freeRects.filter(fr => {
    // Keep if not fully overlapped
    return !(fr.x >= x && fr.y >= y && fr.x + fr.w <= x + pw && fr.y + fr.h <= y + ph);
  });
}

/**
 * Run guillotine packing.
 * Returns array of board objects each with { width, height, kerf, placements[] }
 * placements[]: { id, name, x, y, w, h, rotated }
 *
 * If stockBoards is provided (array of {width,height,qty}), boards are drawn
 * from stock in order. If insufficient, packing stops and items are skipped.
 */
function guillotinePack({ boardW, boardH, kerf, edgeTrim, parts, opts, stockBoards }) {
  // Expand parts list by quantity, sort largest area first
  let items = [];
  for (const p of parts) {
    for (let i = 0; i < p.qty; i++) {
      items.push({ ...p, instanceIdx: i });
    }
  }
  // Sort by max dimension descending (better heuristic)
  items.sort((a, b) => Math.max(b.width, b.height) - Math.max(a.width, a.height));

  const boards = [];

  // Build a mutable stock pool if provided
  // Each entry: { width, height, remaining }
  const stock = stockBoards && stockBoards.length > 0
    ? stockBoards.map(s => ({ width: s.width, height: s.height, remaining: s.qty }))
    : null;

  const openNewBoard = () => {
    if (stock) {
      // Find a stock entry with remaining > 0
      const entry = stock.find(s => s.remaining > 0);
      if (!entry) return null; // out of stock
      entry.remaining--;
      return {
        width: entry.width,
        height: entry.height,
        kerf,
        placements: [],
        freeRects: [makeFreeRect(edgeTrim, edgeTrim, entry.width - 2 * edgeTrim, entry.height - 2 * edgeTrim)],
      };
    }
    // Unlimited mode (no stock)
    return {
      width: boardW,
      height: boardH,
      kerf,
      placements: [],
      freeRects: [makeFreeRect(edgeTrim, edgeTrim, boardW - 2 * edgeTrim, boardH - 2 * edgeTrim)],
    };
  };

  for (const item of items) {
    let placed = false;

    for (const board of boards) {
      // Try each free rect, pick best score
      let bestScore = null;
      let bestFrIdx = -1;
      let bestRotated = false;

      for (let i = 0; i < board.freeRects.length; i++) {
        const fr = board.freeRects[i];
        const scoreObj = scorePlacement(fr, item.width + kerf, item.height + kerf, item.rotate, opts, item.grain);
        if (scoreObj !== null) {
          if (bestScore === null || scoreObj.score < bestScore ||
            (scoreObj.score === bestScore && scoreObj.long < bestScore.long)) {
            bestScore = scoreObj.score;
            bestFrIdx = i;
            bestRotated = scoreObj.rotated;
          }
        }
      }

      if (bestFrIdx !== -1) {
        const fr = board.freeRects[bestFrIdx];
        const pw = bestRotated ? item.height : item.width;
        const ph = bestRotated ? item.width : item.height;
        const pkw = pw + kerf;
        const pkh = ph + kerf;

        const effectiveGrain = (opts && opts.grainDirection)
          ? (opts.boardGrain || 'H')
          : (item.grain || 'H');
        board.placements.push({
          id: item.id,
          name: item.name,
          x: fr.x,
          y: fr.y,
          w: pw,
          h: ph,
          rotated: bestRotated,
          grain: effectiveGrain,
        });

        const newRects = splitFreeRect(fr, pkw, pkh);
        board.freeRects.splice(bestFrIdx, 1);
        board.freeRects.push(...newRects);
        board.freeRects = pruneFreeRects(board.freeRects, fr.x, fr.y, pkw, pkh);

        placed = true;
        break;
      }
    }

    if (!placed) {
      // One-sheet mode: skip opening new boards
      if (opts && opts.oneSheet && boards.length >= 1) {
        toast(`⚠️ "${item.name}" skipped — one-sheet mode is on.`);
        continue;
      }

      // Open a new board from stock (or unlimited)
      const board = openNewBoard();
      if (!board) {
        toast(`❌ Out of stock! "${item.name}" and remaining parts could not be placed.`);
        break; // stop packing
      }
      boards.push(board);

      const fr = board.freeRects[0];
      const scoreObj = scorePlacement(fr, item.width + kerf, item.height + kerf, item.rotate, opts, item.grain);
      if (!scoreObj) {
        toast(`❌ "${item.name}" (${item.width}×${item.height}mm) is larger than the board and cannot be placed.`);
        continue;
      }
      const bestRotated = scoreObj.rotated;
      const pw = bestRotated ? item.height : item.width;
      const ph = bestRotated ? item.width : item.height;
      const pkw = pw + kerf;
      const pkh = ph + kerf;

      const effectiveGrain2 = (opts && opts.grainDirection)
        ? (opts.boardGrain || 'H')
        : (item.grain || 'H');
      board.placements.push({
        id: item.id,
        name: item.name,
        x: fr.x,
        y: fr.y,
        w: pw,
        h: ph,
        rotated: bestRotated,
        grain: effectiveGrain2,
      });

      const newRects = splitFreeRect(fr, pkw, pkh);
      board.freeRects = newRects;
    }
  }

  return boards;
}

/* ────────────────────────────────────────────────
   CANVAS RENDERER
──────────────────────────────────────────────── */
const CANVAS_PAD = 20;

/**
 * Renders all boards for all thickness groups into the right panel.
 * Clears the wrapper and rebuilds it entirely.
 */
function renderAllBoards() {
  const wrapper = $('canvasWrapper');
  // Remove everything except the empty-state div
  [...wrapper.children].forEach(el => {
    if (el.id !== 'emptyState') el.remove();
  });
  $('emptyState').style.display = 'none';

  const availW = wrapper.clientWidth - 24; // account for padding

  for (const group of state.groups) {
    // ── Group header (maroon bar with thickness) ──
    const header = document.createElement('div');
    header.className = 'board-group-header';

    // Deduplicate boards with identical layout fingerprints
    // Fingerprint = sorted list of "id:x:y:w:h" for all placements
    const deduped = []; // [{ board, count }]
    for (const board of group.boards) {
      const fp = board.placements
        .map(p => `${p.id}:${p.x}:${p.y}:${p.w}:${p.h}`)
        .sort().join('|');
      const existing = deduped.find(d => d.fp === fp);
      if (existing) {
        existing.count++;
      } else {
        deduped.push({ board, fp, count: 1 });
      }
    }

    const totalBoards = group.boards.length;
    const totalPieces = group.boards.reduce((s, b) => s + b.placements.length, 0);
    header.textContent =
      `${group.thickness}mm — ${totalBoards} board${totalBoards !== 1 ? 's' : ''} — ${totalPieces} piece${totalPieces !== 1 ? 's' : ''}`;
    wrapper.appendChild(header);

    // ── One card per unique board layout ──
    deduped.forEach(({ board, count }, bi) => {
      const scale = Math.min(availW / board.width, 340 / board.height);
      const cw = Math.round(board.width * scale + CANVAS_PAD * 2);
      const ch = Math.round(board.height * scale + CANVAS_PAD * 2);

      const item = document.createElement('div');
      item.className = 'board-canvas-item';

      const lbl = document.createElement('div');
      lbl.className = 'board-canvas-label';
      const countTag = count > 1 ? ` <span class="board-qty-badge">×${count}</span>` : '';
      lbl.innerHTML = `Board ${bi + 1} of ${deduped.length}  ·  ${board.placements.length} piece${board.placements.length !== 1 ? 's' : ''}${countTag}`;
      item.appendChild(lbl);

      const cvs = document.createElement('canvas');
      cvs.width = cw;
      cvs.height = ch;
      cvs.style.width = cw + 'px';
      cvs.style.height = ch + 'px';
      item.appendChild(cvs);
      wrapper.appendChild(item);

      drawBoardOnCanvas(cvs.getContext('2d'), board, bi, scale);
    });
  }
}

/**
 * Pure drawing function — works on any 2D context.
 * Used by renderAllBoards, buildPrintArea and exportPDF.
 */
function drawBoardOnCanvas(ctx, board, boardIdx, scale) {
  const cw = ctx.canvas.width;
  const ch = ctx.canvas.height;
  ctx.clearRect(0, 0, cw, ch);

  // Board background
  ctx.fillStyle = '#F0EDE8';
  ctx.strokeStyle = '#BBBBC8';
  ctx.lineWidth = 1.5;
  ctx.beginPath();
  ctx.roundRect(CANVAS_PAD, CANVAS_PAD, board.width * scale, board.height * scale, 4);
  ctx.fill();
  ctx.stroke();

  // Waste hatch
  drawWasteHatch(ctx, CANVAS_PAD, CANVAS_PAD, board.width * scale, board.height * scale);

  // Board grain lines (on the board background)
  const opts0 = state.opts || getOptions();
  if (opts0.grainDirection) {
    drawBoardGrainLines(ctx, CANVAS_PAD, CANVAS_PAD, board.width * scale, board.height * scale, opts0.boardGrain || 'H');
  }

  // Placements
  const opts = state.opts || getOptions();
  for (const p of board.placements) {
    const x = CANVAS_PAD + p.x * scale;
    const y = CANVAS_PAD + p.y * scale;
    const w = p.w * scale;
    const h = p.h * scale;

    ctx.fillStyle = hexFromHsl(p.id);
    ctx.beginPath();
    ctx.roundRect(x, y, w, h, 2);
    ctx.fill();
    ctx.strokeStyle = 'rgba(0,0,0,0.18)';
    ctx.lineWidth = 1;
    ctx.stroke();

    // Grain direction lines
    drawGrainLines(ctx, x, y, w, h, p.grain || 'H');

    // Edge banding: draw a coloured stripe on all 4 edges
    if (opts.edgeBanding) {
      const bw = Math.max(2, Math.min(5, w * 0.04));
      ctx.fillStyle = 'rgba(123,28,46,0.55)';
      ctx.fillRect(x, y, w, bw);           // top
      ctx.fillRect(x, y + h - bw, w, bw); // bottom
      ctx.fillRect(x, y, bw, h);           // left
      ctx.fillRect(x + w - bw, y, bw, h); // right
    }

    if (opts.showLabels) {
      const label = p.name ? (p.rotated ? `${p.name}*` : p.name) : '';
      drawLabel(ctx, x, y, w, h, label, `${p.w}×${p.h}`, opts.labelSize ?? 1);
    }
  }

  // Dimension labels
  ctx.font = 'bold 11px Inter, sans-serif';
  ctx.fillStyle = '#888';
  ctx.textAlign = 'center';
  ctx.fillText(`${board.width} mm`, CANVAS_PAD + (board.width * scale) / 2, CANVAS_PAD - 6);
  ctx.save();
  ctx.translate(CANVAS_PAD - 8, CANVAS_PAD + (board.height * scale) / 2);
  ctx.rotate(-Math.PI / 2);
  ctx.fillText(`${board.height} mm`, 0, 0);
  ctx.restore();

  // Board number badge
  ctx.fillStyle = '#7B1C2E';
  ctx.beginPath();
  ctx.roundRect(CANVAS_PAD, CANVAS_PAD, 64, 22, [0, 0, 4, 0]);
  ctx.fill();
  ctx.fillStyle = '#fff';
  ctx.font = 'bold 10px Inter, sans-serif';
  ctx.textAlign = 'left';
  ctx.fillText(`Board ${boardIdx + 1}`, CANVAS_PAD + 8, CANVAS_PAD + 15);
}

/**
 * Draw subtle grain direction lines on a panel.
 * grain: 'H' = horizontal lines, 'V' = vertical lines.
 */
function drawGrainLines(ctx, x, y, w, h, grain) {
  ctx.save();
  ctx.beginPath();
  ctx.rect(x + 1, y + 1, w - 2, h - 2);
  ctx.clip();
  ctx.strokeStyle = 'rgba(80,50,20,0.10)';
  ctx.lineWidth = 1;
  const spacing = 8;
  if (grain === 'V') {
    // Vertical lines
    for (let i = x + spacing; i < x + w; i += spacing) {
      ctx.beginPath(); ctx.moveTo(i, y + 1); ctx.lineTo(i, y + h - 1); ctx.stroke();
    }
  } else {
    // Horizontal lines (default)
    for (let i = y + spacing; i < y + h; i += spacing) {
      ctx.beginPath(); ctx.moveTo(x + 1, i); ctx.lineTo(x + w - 1, i); ctx.stroke();
    }
  }
  // Small arrow indicator in top-left corner
  ctx.strokeStyle = 'rgba(60,40,10,0.35)';
  ctx.lineWidth = 1.5;
  const aw = Math.min(14, w * 0.15);
  const ah = Math.min(14, h * 0.15);
  const mx = x + 4, my = y + 4;
  if (grain === 'V') {
    ctx.beginPath(); ctx.moveTo(mx, my); ctx.lineTo(mx, my + ah); ctx.stroke();
    ctx.beginPath(); ctx.moveTo(mx - 3, my + ah - 4); ctx.lineTo(mx, my + ah); ctx.lineTo(mx + 3, my + ah - 4); ctx.stroke();
  } else {
    ctx.beginPath(); ctx.moveTo(mx, my); ctx.lineTo(mx + aw, my); ctx.stroke();
    ctx.beginPath(); ctx.moveTo(mx + aw - 4, my - 3); ctx.lineTo(mx + aw, my); ctx.lineTo(mx + aw - 4, my + 3); ctx.stroke();
  }
  ctx.restore();
}

/**
 * Draw subtle grain lines on the board background (not on panels).
 * Uses wider spacing and lighter color than panel grain lines.
 */
function drawBoardGrainLines(ctx, x, y, w, h, grain) {
  ctx.save();
  ctx.beginPath();
  ctx.rect(x, y, w, h);
  ctx.clip();
  ctx.strokeStyle = 'rgba(140,110,70,0.08)';
  ctx.lineWidth = 1;
  const spacing = 14;
  if (grain === 'V') {
    for (let i = x + spacing; i < x + w; i += spacing) {
      ctx.beginPath(); ctx.moveTo(i, y); ctx.lineTo(i, y + h); ctx.stroke();
    }
  } else {
    for (let i = y + spacing; i < y + h; i += spacing) {
      ctx.beginPath(); ctx.moveTo(x, i); ctx.lineTo(x + w, i); ctx.stroke();
    }
  }
  ctx.restore();
}

function drawWasteHatch(ctx, x, y, w, h) {
  ctx.save();
  ctx.beginPath();
  ctx.rect(x, y, w, h);
  ctx.clip();
  ctx.strokeStyle = 'rgba(190,185,180,0.45)';
  ctx.lineWidth = 1;
  const spacing = 12;
  for (let i = -h; i < w + h; i += spacing) {
    ctx.beginPath();
    ctx.moveTo(x + i, y);
    ctx.lineTo(x + i + h, y + h);
    ctx.stroke();
  }
  ctx.restore();
}

function drawLabel(ctx, x, y, w, h, name, dims, sizeScale = 1) {
  const fontSize = Math.max(6, Math.min(13, w / 8, h / 4) * sizeScale);
  const dimFontSize = Math.max(5, (fontSize - 2));
  ctx.save();
  ctx.beginPath();
  ctx.rect(x + 1, y + 1, w - 2, h - 2);
  ctx.clip();
  ctx.textAlign = 'center';
  ctx.textBaseline = 'middle';
  const cx = x + w / 2;
  const cy = y + h / 2;
  if (h > 20 && w > 28) {
    if (name) {
      ctx.font = `600 ${fontSize}px Inter, sans-serif`;
      ctx.fillStyle = 'rgba(0,0,0,0.72)';
      ctx.fillText(truncate(name, w, ctx), cx, cy - dimFontSize * 0.7);
      ctx.font = `400 ${dimFontSize}px Inter, sans-serif`;
      ctx.fillStyle = 'rgba(0,0,0,0.5)';
      ctx.fillText(dims, cx, cy + fontSize * 0.6);
    } else {
      // No name — show only dimensions centered
      ctx.font = `500 ${fontSize}px Inter, sans-serif`;
      ctx.fillStyle = 'rgba(0,0,0,0.6)';
      ctx.fillText(dims, cx, cy);
    }
  } else if (h > 10 && w > 14) {
    ctx.font = `600 ${fontSize}px Inter, sans-serif`;
    ctx.fillStyle = 'rgba(0,0,0,0.72)';
    ctx.fillText(name ? truncate(name, w, ctx) : dims, cx, cy);
  }
  ctx.restore();
}

function truncate(text, maxPx, ctx2) {
  let t = text;
  while (ctx2.measureText(t).width > maxPx - 10 && t.length > 2) t = t.slice(0, -1);
  return t === text ? t : t + '…';
}

/* ────────────────────────────────────────────────
   OPTIMIZE — MAIN
──────────────────────────────────────────────── */
function optimize() {
  const boardW = parseFloat($('boardWidth').value);
  const boardH = parseFloat($('boardHeight').value);
  const boardThick = parseFloat($('boardThickness').value) || 18;
  const kerf = parseFloat($('bladeKerf').value) || 0;
  const edgeTrim = parseFloat($('edgeTrim').value) || 0;
  const opts = getOptions();

  if (!boardW || !boardH || boardW <= 0 || boardH <= 0) {
    toast('⚠️ Enter valid board dimensions.');
    return;
  }

  let parts = collectParts();
  if (!parts) return;
  if (parts.length === 0) {
    toast('⚠️ Add at least one part.');
    return;
  }

  // Consider material: only pack parts matching the active board thickness
  if (opts.considerMat) {
    const skipped = parts.filter(p => p.thickness !== boardThick).length;
    parts = parts.filter(p => p.thickness === boardThick);
    if (skipped > 0) toast(`ℹ️ ${skipped} part(s) skipped — thickness doesn't match board (${boardThick}mm).`);
    if (parts.length === 0) { toast('⚠️ No parts match the board thickness.'); return; }
  }

  // Get stock boards (if any)
  const stockBoards = state.stock.length > 0 ? state.stock.map(s => ({ ...s })) : null;

  // Validate stock has at least one board available
  if (stockBoards && stockBoards.every(s => s.qty === 0)) {
    toast('❌ All stock boards are out of stock. Add more stock first.');
    return;
  }

  // Group parts by their individual thickness
  const thicknessMap = new Map();
  for (const p of parts) {
    const t = p.thickness;
    if (!thicknessMap.has(t)) thicknessMap.set(t, []);
    thicknessMap.get(t).push(p);
  }

  // Run guillotine packing for each thickness group
  state.groups = [];
  state.opts = opts;   // store for re-render
  for (const [thickness, groupParts] of thicknessMap) {
    const boards = guillotinePack({ boardW, boardH, kerf, edgeTrim, parts: groupParts, opts, stockBoards });
    state.groups.push({ thickness, boards });
  }
  // Sort thickest-first
  state.groups.sort((a, b) => b.thickness - a.thickness);

  updateStats();
  renderAllBoards();

  $('statsSection').style.display = 'block';
  $('exportSection').style.display = 'block';
}

/* ────────────────────────────────────────────────
   STATS
──────────────────────────────────────────────── */
function updateStats() {
  const boardW = parseFloat($('boardWidth').value) || 0;
  const boardH = parseFloat($('boardHeight').value) || 0;
  const boardArea = boardW * boardH;

  let totalBoards = 0, usedArea = 0, totalParts = 0;
  for (const group of state.groups) {
    totalBoards += group.boards.length;
    for (const board of group.boards) {
      for (const p of board.placements) {
        usedArea += p.w * p.h;
        totalParts++;
      }
    }
  }

  const totalBoardArea = totalBoards * boardArea;
  const wasteArea = Math.max(0, totalBoardArea - usedArea);
  const wastePct = totalBoardArea > 0
    ? ((wasteArea / totalBoardArea) * 100).toFixed(1)
    : '0.0';

  $('statBoards').textContent = totalBoards;
  $('statUsed').textContent = fmtMm2(usedArea);
  $('statWaste').textContent = `${wastePct}%`;
  $('statParts').textContent = `${totalParts} pcs`;
}

/* ────────────────────────────────────────────────
   PDF EXPORT
──────────────────────────────────────────────── */
async function exportPDF() {
  if (state.groups.length === 0) { toast('Run Optimize first.'); return; }

  const { jsPDF } = window.jspdf;
  const pageW = 210, pageH = 297;
  const margin = 12;

  // ── Load logo ──
  let logoDataUrl = null;
  try {
    const resp = await fetch('logo.png');
    const blob = await resp.blob();
    logoDataUrl = await new Promise(res => {
      const r = new FileReader();
      r.onload = () => res(r.result);
      r.readAsDataURL(blob);
    });
  } catch (e) { /* logo optional */ }

  const opts = state.opts || getOptions();
  const material = $('materialName').value || '—';
  const boardW = $('boardWidth').value || '—';
  const boardH = $('boardHeight').value || '—';
  const dateStr = new Date().toLocaleDateString('en-GB');

  // Flatten boards
  const allBoards = [];
  for (const group of state.groups) {
    group.boards.forEach((board, bi) =>
      allBoards.push({ board, boardIdx: bi, groupLabel: `${group.thickness}mm`, totalInGroup: group.boards.length })
    );
  }

  // ── Layout constants (portrait A4) ──
  // Header: 18mm | two boards stacked, each ~118mm tall | gap 6mm | footer 8mm
  const HEADER_H = 18;
  const FOOTER_H = 8;
  const GAP = 5;
  const BOARD_H = Math.floor((pageH - HEADER_H - FOOTER_H - GAP * 3 - margin * 2) / 2); // ≈118mm
  const DRAW_W = pageW - margin * 2;

  // High-res render dimensions matching buildPrintArea
  const RENDER_W = 1800;
  const RENDER_H = Math.round(RENDER_W * (BOARD_H / DRAW_W));

  const doc = new jsPDF({ orientation: 'portrait', unit: 'mm', format: 'a4' });

  // ── Page header helper ──
  const drawPageHeader = () => {
    doc.setFillColor(123, 28, 46);
    doc.rect(0, 0, pageW, HEADER_H, 'F');

    if (logoDataUrl) {
      try { doc.addImage(logoDataUrl, 'PNG', margin, 1, 14, 14); } catch (e) { }
    }
    const tx = logoDataUrl ? margin + 17 : margin;
    doc.setTextColor(255, 255, 255);
    doc.setFontSize(10.5);
    doc.setFont('helvetica', 'bold');
    doc.text('MAESTRO FURNITURE — CutList Optimizer', tx, 8);
    doc.setFontSize(7.5);
    doc.setFont('helvetica', 'normal');
    doc.text(`${material}  |  ${boardW}×${boardH}mm  |  ${dateStr}`, tx, 14.5);
  };

  // ── Footer helper ──
  const drawPageFooter = (pageNum, totalPages) => {
    const fy = pageH - FOOTER_H + 3;
    doc.setDrawColor(200, 200, 210);
    doc.setLineWidth(0.3);
    doc.line(margin, fy - 1, pageW - margin, fy - 1);
    doc.setFontSize(7);
    doc.setFont('helvetica', 'normal');
    doc.setTextColor(150, 150, 165);
    doc.text('Maestro Furniture — CutList Optimizer', margin, fy + 3);
    doc.text(`Page ${pageNum} of ${totalPages}`, pageW - margin, fy + 3, { align: 'right' });
  };

  // ── Board render helper ──
  const renderBoard = (board, boardIdx, y, isFirst) => {
    const scale = Math.min(DRAW_W / board.width, BOARD_H / board.height);
    const cvs = document.createElement('canvas');
    cvs.width = Math.round(board.width * scale * 4 + CANVAS_PAD * 2);  // 4× for sharp output
    cvs.height = Math.round(board.height * scale * 4 + CANVAS_PAD * 2);
    drawBoardOnCanvas(cvs.getContext('2d'), board, boardIdx, scale * 4);

    // Compute actual mm dimensions to preserve aspect ratio
    const aspect = cvs.width / cvs.height;
    let imgW = DRAW_W;
    let imgH = imgW / aspect;
    if (imgH > BOARD_H) { imgH = BOARD_H; imgW = imgH * aspect; }
    const imgX = margin + (DRAW_W - imgW) / 2;

    doc.addImage(cvs.toDataURL('image/png'), 'PNG', imgX, y, imgW, imgH);

    // Small board label above the image
    doc.setFontSize(7.5);
    doc.setFont('helvetica', 'bold');
    doc.setTextColor(123, 28, 46);

    // Build label: edge banding total if enabled
    let lbl = `${isFirst ? '' : ''}Board ${boardIdx + 1} of ${allBoards.filter(b => b.groupLabel === allBoards.find(x => x.board === board)?.groupLabel).length}   ·   ${board.placements.length} piece${board.placements.length !== 1 ? 's' : ''}`;
    if (opts.edgeBanding) {
      const eb = board.placements.reduce((s, p) => s + 2 * (p.w + p.h), 0);
      lbl += `   ·   Edge banding: ${eb.toLocaleString()} mm`;
    }
    doc.text(lbl, margin, y - 1.5);
  };

  // ── Build pages: 2 boards per page ──
  const boardPages = [];
  for (let i = 0; i < allBoards.length; i += 2) {
    boardPages.push(allBoards.slice(i, i + 2));
  }
  const totalPages = boardPages.length + 1; // +1 for summary

  boardPages.forEach((pair, pi) => {
    if (pi > 0) doc.addPage();
    drawPageHeader();

    const y1 = HEADER_H + GAP;
    const y2 = y1 + BOARD_H + GAP;

    renderBoard(pair[0].board, pair[0].boardIdx, y1, true);
    if (pair[1]) renderBoard(pair[1].board, pair[1].boardIdx, y2, false);

    drawPageFooter(pi + 1, totalPages);
  });

  // ── Summary / Parts page ──
  doc.addPage();
  drawPageHeader();

  let totalEBLength = 0;
  for (const group of state.groups)
    for (const board of group.boards)
      for (const p of board.placements)
        totalEBLength += 2 * (p.w + p.h);

  let sy = HEADER_H + 8;

  doc.setTextColor(30, 30, 40);
  doc.setFontSize(10);
  doc.setFont('helvetica', 'bold');
  doc.text('Project Summary', margin, sy);
  sy += 8;

  const sumData = [
    ['Material', material],
    ['Board Size', `${boardW} × ${boardH} mm`],
    ['Thickness', `${$('boardThickness').value} mm`],
    ['Blade Kerf', `${$('bladeKerf').value} mm`],
    ['Boards Used', state.groups.reduce((s, g) => s + g.boards.length, 0).toString()],
    ['Total Parts', [...partsBody.querySelectorAll('tr')].reduce((a, tr) => a + (parseInt(tr.querySelector('[data-field="qty"]')?.value) || 1), 0).toString()],
    ['Waste', $('statWaste').textContent],
  ];
  if (opts.edgeBanding)
    sumData.push(['Edge Banding', `${totalEBLength.toLocaleString()} mm  (${(totalEBLength / 1000).toFixed(2)} m)`]);

  doc.setFontSize(8.5);
  for (const [k, v] of sumData) {
    doc.setFont('helvetica', 'bold'); doc.setTextColor(80, 80, 100); doc.text(k + ':', margin, sy);
    doc.setFont('helvetica', 'normal'); doc.setTextColor(30, 30, 40); doc.text(v, margin + 36, sy);
    sy += 6.5;
  }

  // Parts table
  sy += 4;
  doc.setFont('helvetica', 'bold'); doc.setFontSize(10); doc.setTextColor(30, 30, 40);
  doc.text('Full Parts List', margin, sy);
  sy += 7;

  const showEB = opts.edgeBanding;
  const cols = showEB
    ? { name: margin, w: margin + 50, h: margin + 72, qty: margin + 94, rot: margin + 110, eb: margin + 128 }
    : { name: margin, w: margin + 55, h: margin + 80, qty: margin + 108, rot: margin + 130, status: margin + 152 };
  const hdrs = showEB
    ? ['Part Name', 'Width', 'Height', 'Qty', 'Rot', 'Edge Band']
    : ['Part Name', 'Width', 'Height', 'Qty', 'Rot', 'Status'];
  const colArr = showEB
    ? [cols.name, cols.w, cols.h, cols.qty, cols.rot, cols.eb]
    : [cols.name, cols.w, cols.h, cols.qty, cols.rot, cols.status];

  doc.setFillColor(123, 28, 46);
  doc.rect(margin, sy - 5, pageW - margin * 2, 7, 'F');
  doc.setFontSize(7.5); doc.setFont('helvetica', 'bold'); doc.setTextColor(255, 255, 255);
  hdrs.forEach((h, i) => doc.text(h, colArr[i], sy));
  sy += 5;

  const allParts = collectParts();
  if (allParts) {
    let rowAlt = false;
    for (const p of allParts) {
      if (sy > pageH - FOOTER_H - 10) {
        drawPageFooter(totalPages, totalPages);
        doc.addPage(); drawPageHeader(); sy = HEADER_H + 10;
      }
      if (rowAlt) { doc.setFillColor(245, 244, 248); doc.rect(margin, sy - 4, pageW - margin * 2, 5.5, 'F'); }
      rowAlt = !rowAlt;
      doc.setFontSize(7.5); doc.setFont('helvetica', 'normal'); doc.setTextColor(30, 30, 40);
      const eb6 = showEB ? `${2 * (p.width + p.height)} mm ×${p.qty}` : '✓';
      doc.text(p.name.length > 26 ? p.name.slice(0, 25) + '…' : p.name, colArr[0], sy);
      doc.text(p.width.toString(), colArr[1], sy);
      doc.text(p.height.toString(), colArr[2], sy);
      doc.text(p.qty.toString(), colArr[3], sy);
      doc.text(p.rotate ? 'Yes' : 'No', colArr[4], sy);
      doc.text(eb6, colArr[5], sy);
      sy += 5.5;
    }
  }

  drawPageFooter(totalPages, totalPages);

  // ── Download ──
  const dateName = new Date().toLocaleDateString('en-GB').replace(/\//g, '-');
  const fileName = `Maestro-CutList-${dateName}.pdf`;
  const pdfBlob = doc.output('blob');
  const url = URL.createObjectURL(pdfBlob);
  const a = document.createElement('a');
  a.href = url; a.download = fileName;
  document.body.appendChild(a); a.click();
  setTimeout(() => { URL.revokeObjectURL(url); a.remove(); }, 3000);
  toast('📄 PDF exported!');
}



/* ────────────────────────────────────────────────
   FILE-BASED PROJECT SAVE / LOAD
   Uses File System Access API when available
   (Chrome / Edge on file:// fully supported).
   Falls back to download-link + file-input upload.
──────────────────────────────────────────────── */

const FILE_EXTENSION = '.mcl.json';
const FILE_TYPES = [{
  description: 'Maestro CutList Project',
  accept: { 'application/json': ['.json'] },
}];

function currentProjectData() {
  const rows = [...partsBody.querySelectorAll('tr')];
  const parts = rows.map(tr => {
    const get = f => tr.querySelector(`[data-field="${f}"]`);
    const grainEl = tr.querySelector('[data-field="grain"]');
    return {
      id: tr.dataset.partId,
      name: get('name').value.trim(),
      width: get('width').value,
      height: get('height').value,
      thickness: get('thickness').value,
      qty: get('qty').value,
      grain: grainEl ? grainEl.dataset.value : 'H',
    };
  });
  return {
    _maestroVersion: '1.4',
    savedAt: new Date().toISOString(),
    boardWidth: $('boardWidth').value,
    boardHeight: $('boardHeight').value,
    boardThickness: $('boardThickness').value,
    bladeKerf: $('bladeKerf').value,
    edgeTrim: $('edgeTrim').value,
    materialName: $('materialName').value,
    boardGrain: getBoardGrain(),
    options: getOptions(),
    stock: state.stock,
    parts,
  };
}

function applyProjectData(data) {
  $('boardWidth').value = data.boardWidth ?? 2440;
  $('boardHeight').value = data.boardHeight ?? 1220;
  $('boardThickness').value = data.boardThickness ?? 18;
  $('bladeKerf').value = data.bladeKerf ?? 3;
  $('edgeTrim').value = data.edgeTrim ?? 0;
  $('materialName').value = data.materialName ?? '';

  // Restore board grain
  setBoardGrainUI(data.boardGrain ?? data.options?.boardGrain ?? 'H');

  // Restore options if present
  if (data.options) {
    const o = data.options;
    if ($('optLabels')) $('optLabels').checked = o.showLabels ?? true;
    if ($('optOneSheet')) $('optOneSheet').checked = o.oneSheet ?? false;
    if ($('optConsiderMaterial')) $('optConsiderMaterial').checked = o.considerMat ?? false;
    if ($('optEdgeBanding')) $('optEdgeBanding').checked = o.edgeBanding ?? false;
    if ($('optGrain')) $('optGrain').checked = o.grainDirection ?? true;
  }

  // Restore stock
  state.stock = Array.isArray(data.stock) ? data.stock : [];
  renderStockUI();

  partsBody.innerHTML = '';
  _partId = 0;
  for (const p of (data.parts || [])) addPartRow(p);
  updateCount();
}

/* ── SAVE ── */
function openSaveModal() {
  $('saveProjectName').value = $('materialName').value || '';
  $('saveModal').classList.add('is-open');
  setTimeout(() => $('saveProjectName').focus(), 80);
}

function closeSaveModal() { $('saveModal').classList.remove('is-open'); }

async function saveCurrentProject() {
  const name = $('saveProjectName').value.trim();
  if (!name) { toast('Enter a project name first.'); return; }
  closeSaveModal();

  const data = currentProjectData();
  data.name = name;
  const json = JSON.stringify(data, null, 2);
  const safeName = name.replace(/[<>:"/\\|?*]/g, '_');

  // ── Path A: File System Access API (Chrome / Edge) ──
  if (typeof window.showSaveFilePicker === 'function') {
    try {
      const handle = await window.showSaveFilePicker({
        suggestedName: safeName + FILE_EXTENSION,
        types: FILE_TYPES,
      });
      const writable = await handle.createWritable();
      await writable.write(json);
      await writable.close();
      toast(`💾 Saved to disk: ${name}`);
      return;
    } catch (err) {
      if (err.name === 'AbortError') return; // user cancelled picker
      console.warn('showSaveFilePicker failed, falling back to download:', err);
    }
  }

  // ── Path B: Download fallback (Firefox / Safari) ──
  const blob = new Blob([json], { type: 'application/json' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = safeName + FILE_EXTENSION;
  document.body.appendChild(a);
  a.click();
  setTimeout(() => { URL.revokeObjectURL(url); a.remove(); }, 2000);
  toast(`💾 Downloaded: ${safeName}${FILE_EXTENSION}`);
}

/* ── LOAD ── */
function openProjectModal() {
  // If File System Access API is available use it directly (no modal needed)
  if (typeof window.showOpenFilePicker === 'function') {
    loadFromFilePicker();
    return;
  }
  // Fallback: trigger hidden file input
  loadFromFileInput();
}

async function loadFromFilePicker() {
  try {
    const [handle] = await window.showOpenFilePicker({ types: FILE_TYPES, multiple: false });
    const file = await handle.getFile();
    const text = await file.text();
    const data = JSON.parse(text);
    applyProjectData(data);
    const label = data.name || file.name;
    toast(`✅ Loaded: ${label}`);
  } catch (err) {
    if (err.name !== 'AbortError') toast('❌ Could not open file.');
  }
}

function loadFromFileInput() {
  let input = document.getElementById('_fileLoadInput');
  if (!input) {
    input = document.createElement('input');
    input.type = 'file';
    input.accept = '.json';
    input.id = '_fileLoadInput';
    input.style.display = 'none';
    document.body.appendChild(input);
  }
  // Reset so the same file can be re-picked
  input.value = '';
  input.onchange = async () => {
    const file = input.files[0];
    if (!file) return;
    try {
      const text = await file.text();
      const data = JSON.parse(text);
      applyProjectData(data);
      const label = data.name || file.name;
      toast(`✅ Loaded: ${label}`);
    } catch {
      toast('❌ Invalid project file.');
    }
  };
  input.click();
}

// Modal close helpers kept for the save modal
function closeProjectModal() { $('projectModal').classList.remove('is-open'); }

function escHtml(str) {
  return str.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

/**
 * Custom confirm — replaces native confirm() which is blocked on file:// in many browsers.
 * Usage: showConfirm('message', () => { /* confirmed *\/ });
 */
function showConfirm(message, onConfirm) {
  // Reuse saveModal structure but swap content temporarily
  const overlay = document.createElement('div');
  overlay.className = 'modal-overlay is-open';
  overlay.innerHTML = `
    <div class="modal modal--sm" style="animation:modalIn 0.18s ease;">
      <div class="modal-header"><h2>Confirm</h2></div>
      <div class="modal-body" style="padding:20px;font-size:14px;color:var(--grey-700);">${escHtml(message)}</div>
      <div class="modal-footer">
        <button class="btn btn-ghost" id="_cfnNo">Cancel</button>
        <button class="btn btn-primary" id="_cfnYes" style="background:var(--maroon);color:#fff;">Confirm</button>
      </div>
    </div>
  `;
  document.body.appendChild(overlay);
  const close = () => document.body.removeChild(overlay);
  overlay.querySelector('#_cfnNo').addEventListener('click', close);
  overlay.querySelector('#_cfnYes').addEventListener('click', () => { close(); onConfirm(); });
  overlay.addEventListener('click', e => { if (e.target === overlay) close(); });
}

function clearAll() {
  if (partsBody.children.length === 0) return;
  showConfirm('Clear all parts? This cannot be undone.', () => {
    partsBody.innerHTML = '';
    updateCount();
    toast('🗑 Parts cleared.');
  });
}

/* ────────────────────────────────────────────────
   STOCK BOARDS — UI & STATE
──────────────────────────────────────────────── */

function renderStockUI() {
  const body = $('stockBody');
  const empty = $('stockEmpty');
  body.innerHTML = '';

  if (state.stock.length === 0) {
    empty.style.display = 'block';
    return;
  }
  empty.style.display = 'none';

  state.stock.forEach((entry, idx) => {
    const row = document.createElement('div');
    row.className = 'stock-row';
    row.innerHTML = `
      <input class="stock-input" type="number" placeholder="W" min="1" value="${entry.width}" data-idx="${idx}" data-field="width" />
      <span class="stock-x">×</span>
      <input class="stock-input" type="number" placeholder="H" min="1" value="${entry.height}" data-idx="${idx}" data-field="height" />
      <input class="stock-input stock-qty" type="number" placeholder="Qty" min="0" value="${entry.qty}" data-idx="${idx}" data-field="qty" title="Quantity in stock" />
      <button class="btn btn-danger stock-del" data-idx="${idx}" title="Remove">✕</button>
    `;
    body.appendChild(row);
  });

  // Bind inputs
  body.querySelectorAll('.stock-input').forEach(inp => {
    inp.addEventListener('change', () => {
      const i = parseInt(inp.dataset.idx);
      const f = inp.dataset.field;
      state.stock[i][f] = Math.max(0, parseFloat(inp.value) || 0);
    });
  });

  // Bind delete buttons
  body.querySelectorAll('.stock-del').forEach(btn => {
    btn.addEventListener('click', () => {
      const i = parseInt(btn.dataset.idx);
      state.stock.splice(i, 1);
      renderStockUI();
    });
  });
}

function addStockRow() {
  // Default to the current board dimensions
  const w = parseFloat($('boardWidth').value) || 2440;
  const h = parseFloat($('boardHeight').value) || 1220;
  state.stock.push({ width: w, height: h, qty: 1 });
  renderStockUI();
  toast('📦 Stock board added.');
}

$('btnAddStock').addEventListener('click', addStockRow);

/* ────────────────────────────────────────────────
   EVENT LISTENERS
──────────────────────────────────────────────── */
$('btnOptimize').addEventListener('click', optimize);
$('btnAddPart').addEventListener('click', () => { addPartRow(); });
$('btnClearAll').addEventListener('click', clearAll);

$('btnExportPDF').addEventListener('click', exportPDF);

$('btnPrint').addEventListener('click', () => {
  if (state.groups.length === 0) { toast('Run Optimize first.'); return; }
  buildPrintArea();
  window.print();
});

/**
 * Renders all boards into #printArea for printing.
 * Always portrait A4 — 2 boards stacked per page.
 * Groups boards by thickness with a heading per group.
 */
function buildPrintArea() {
  const area = $('printArea');
  area.innerHTML = '';

  const material = $('materialName').value || '—';
  const boardW = $('boardWidth').value || '—';
  const boardH = $('boardHeight').value || '—';
  const dateStr = new Date().toLocaleDateString('en-GB');

  // Inject portrait @page rule
  let pageStyle = document.getElementById('_printPageStyle');
  if (!pageStyle) { pageStyle = document.createElement('style'); pageStyle.id = '_printPageStyle'; document.head.appendChild(pageStyle); }
  pageStyle.textContent = '@page { size: A4 portrait; margin: 10mm 12mm; }';

  // Portrait: full printable width ≈ 190mm — render at 1400px for sharpness
  // 2 boards stacked vertically per page, each ≤ 580px tall
  const PRINT_W = 1400;
  const PRINT_H = 580;

  let boardsOnPage = 0;
  let currentPageEl = null;

  const allBoards = []; // flat list: { board, boardIdx, groupLabel, totalInGroup }
  for (const group of state.groups) {
    group.boards.forEach((board, bi) => {
      allBoards.push({ board, boardIdx: bi, groupLabel: `${group.thickness}mm`, totalInGroup: group.boards.length });
    });
  }

  const totalAll = allBoards.length;

  allBoards.forEach(({ board, boardIdx, groupLabel, totalInGroup }, globalIdx) => {
    if (boardsOnPage === 0) {
      // Page header
      const headerEl = document.createElement('div');
      headerEl.className = 'print-page-header';
      headerEl.innerHTML = `
        <div class="print-brand">Maestro Furniture — CutList Optimizer</div>
        <div class="print-meta">${material} &nbsp;|&nbsp; ${boardW}×${boardH}mm &nbsp;|&nbsp; ${totalAll} boards &nbsp;|&nbsp; ${dateStr}</div>
      `;
      area.appendChild(headerEl);

      currentPageEl = document.createElement('div');
      currentPageEl.className = 'print-page portrait';
      area.appendChild(currentPageEl);
    }

    const scale = Math.min(PRINT_W / board.width, PRINT_H / board.height);
    const cw = Math.round(board.width * scale + CANVAS_PAD * 2);
    const ch = Math.round(board.height * scale + CANVAS_PAD * 2);

    const cvs = document.createElement('canvas');
    cvs.width = cw;
    cvs.height = ch;
    drawBoardOnCanvas(cvs.getContext('2d'), board, boardIdx, scale);

    const cell = document.createElement('div');
    cell.className = 'print-board-cell';
    const lbl = document.createElement('div');
    lbl.className = 'print-board-label';
    lbl.textContent = `${groupLabel} — Board ${boardIdx + 1} of ${totalInGroup}  ·  ${board.placements.length} piece${board.placements.length !== 1 ? 's' : ''}`;
    cell.appendChild(lbl);
    cell.appendChild(cvs);
    currentPageEl.appendChild(cell);

    boardsOnPage++;
    if (boardsOnPage === 2) boardsOnPage = 0; // reset after 2 per page
  });

  // Footer
  const footer = document.createElement('div');
  footer.className = 'print-footer';
  footer.innerHTML = `
    <span>Maestro Furniture — CutList Optimizer</span>
    <span>Waste: ${$('statWaste').textContent} &nbsp;|&nbsp; Material: ${$('statUsed').textContent} &nbsp;|&nbsp; ${dateStr}</span>
  `;
  area.appendChild(footer);
}

$('btnNewProject').addEventListener('click', () => {
  showConfirm('Start a new project? Unsaved changes will be lost.', () => {
    applyProjectData({ parts: [] });
    state.groups = [];
    const wrapper = $('canvasWrapper');
    [...wrapper.children].forEach(el => { if (el.id !== 'emptyState') el.remove(); });
    $('emptyState').style.display = 'flex';
    $('statsSection').style.display = 'none';
    $('exportSection').style.display = 'none';
    toast('🆕 New project started.');
  });
});


$('btnLoadProject').addEventListener('click', openProjectModal);
$('modalClose').addEventListener('click', closeProjectModal);
$('modalCancel').addEventListener('click', closeProjectModal);
$('projectModal').addEventListener('click', e => { if (e.target === $('projectModal')) closeProjectModal(); });

$('btnSaveProject').addEventListener('click', openSaveModal);
$('saveModalClose').addEventListener('click', closeSaveModal);
$('saveModalCancel').addEventListener('click', closeSaveModal);
$('saveModalConfirm').addEventListener('click', saveCurrentProject);
$('saveProjectName').addEventListener('keydown', e => { if (e.key === 'Enter') saveCurrentProject(); });
$('saveModal').addEventListener('click', e => { if (e.target === $('saveModal')) closeSaveModal(); });

/* Re-render on window resize */
let resizeTimer;
window.addEventListener('resize', () => {
  clearTimeout(resizeTimer);
  resizeTimer = setTimeout(() => {
    if (state.groups.length > 0) renderAllBoards();
  }, 160);
});

/* ── OPTIONS COLLAPSE TOGGLE ── */
$('optionsHeader').addEventListener('click', () => {
  $('optionsSection').classList.toggle('collapsed');
});

/* Re-render canvas when label/edge-banding toggles change (instant feedback) */
['optLabels', 'optEdgeBanding'].forEach(id => {
  $(id)?.addEventListener('change', () => {
    if (state.groups.length > 0) {
      state.opts = getOptions();
      renderAllBoards();
    }
  });
});

/* Label size slider — live update */
$('optLabelSize')?.addEventListener('input', () => {
  const val = $('optLabelSize').value;
  $('optLabelSizeVal').textContent = val + '%';
  if (state.groups.length > 0) {
    state.opts = getOptions();
    renderAllBoards();
  }
});

/* Show/hide label size slider based on labels toggle */
$('optLabels')?.addEventListener('change', () => {
  const row = $('labelSizeRow');
  if (row) row.style.opacity = $('optLabels').checked ? '1' : '0.4';
});

/* Board grain toggle handler */
(function initBoardGrainToggle() {
  const toggle = $('boardGrainToggle');
  if (!toggle) return;
  toggle.dataset.value = 'H';
  toggle.querySelectorAll('.grain-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      const dir = btn.dataset.dir;
      toggle.dataset.value = dir;
      toggle.querySelectorAll('.grain-btn').forEach(b => b.classList.toggle('active', b.dataset.dir === dir));
      if (state.groups.length > 0) {
        state.opts = getOptions();
        renderAllBoards();
      }
    });
  });
})();

function setBoardGrainUI(dir) {
  const toggle = $('boardGrainToggle');
  if (!toggle) return;
  toggle.dataset.value = dir;
  toggle.querySelectorAll('.grain-btn').forEach(b => b.classList.toggle('active', b.dataset.dir === dir));
}

/* ────────────────────────────────────────────────
   INIT — Load default parts
──────────────────────────────────────────────── */
(function init() {
  const defaultParts = [
    { name: 'Side Panel', width: 720, height: 400, qty: 2 },
    { name: 'Top Panel', width: 800, height: 400, qty: 1 },
    { name: 'Bottom Panel', width: 800, height: 400, qty: 1 },
    { name: 'Back Panel', width: 800, height: 700, qty: 1 },
    { name: 'Shelf', width: 760, height: 350, qty: 2 },
    { name: 'Door', width: 380, height: 700, qty: 2 },
  ];
  defaultParts.forEach(p => addPartRow(p));
  renderStockUI();
  initMobile();
})();

/* ────────────────────────────────────────────────
   MOBILE UI — tab switching + accordion sections
──────────────────────────────────────────────── */
function initMobile() {
  // ── Mobile tab switching ──
  const tabs = document.querySelectorAll('.mobile-tab');
  const panels = { panelLeft: $('panelLeft'), panelCenter: $('panelCenter'), panelRight: $('panelRight') };

  function activateTab(panelId) {
    tabs.forEach(t => t.classList.toggle('active', t.dataset.panel === panelId));
    Object.entries(panels).forEach(([id, el]) => {
      if (el) el.classList.toggle('mobile-active', id === panelId);
    });
    // Re-render boards when switching to layout tab on mobile
    if (panelId === 'panelRight' && state.groups.length > 0) {
      setTimeout(() => renderAllBoards(), 50);
    }
  }

  tabs.forEach(tab => {
    tab.addEventListener('click', () => activateTab(tab.dataset.panel));
  });

  // Set initial state
  const isMobile = () => window.innerWidth <= 768;
  function initPanelVisibility() {
    if (isMobile()) {
      // Show only Settings panel by default on mobile
      Object.values(panels).forEach(el => el && el.classList.remove('mobile-active'));
      if (panels.panelLeft) panels.panelLeft.classList.add('mobile-active');
      tabs.forEach(t => t.classList.toggle('active', t.dataset.panel === 'panelLeft'));
    } else {
      // On desktop, all panels visible (no mobile-active class needed)
      Object.values(panels).forEach(el => el && el.classList.remove('mobile-active'));
    }
  }

  initPanelVisibility();
  window.addEventListener('resize', () => {
    if (!isMobile()) {
      Object.values(panels).forEach(el => el && el.classList.remove('mobile-active'));
    } else {
      // Ensure at least one tab is visible
      const anyActive = [...tabs].some(t => t.classList.contains('active'));
      if (!anyActive) initPanelVisibility();
    }
  });

  // After optimize on mobile, auto-switch to Layout tab
  const origOptimize = window._mobileOrigOptimize;
  $('btnOptimize') && $('btnOptimize').addEventListener('click', () => {
    if (isMobile()) setTimeout(() => activateTab('panelRight'), 120);
  });
  $('mobileFabOptimize') && $('mobileFabOptimize').addEventListener('click', () => {
    optimize();
    if (isMobile()) setTimeout(() => activateTab('panelRight'), 120);
  });

  // ── Mobile collapsible sections (accordion) ──
  document.querySelectorAll('.mob-section-header').forEach(header => {
    const targetId = header.dataset.target;
    const body = document.getElementById(targetId);
    if (!body) return;

    header.addEventListener('click', () => {
      const isCollapsed = body.classList.contains('collapsed');
      body.classList.toggle('collapsed', !isCollapsed);
      header.classList.toggle('collapsed', !isCollapsed);
    });
  });

  // ── Mobile: New Project secondary button ──
  const btnNPM = $('btnNewProjectMobile');
  if (btnNPM) {
    btnNPM.addEventListener('click', () => $('btnNewProject') && $('btnNewProject').click());
  }
}

