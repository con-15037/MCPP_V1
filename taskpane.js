/* MCPP_V1 — PowerPoint task pane
 * ---------------------------------------------------------------
 * UNITS. The Office API works in POINTS. This panel works in the
 * PIXELS of a 1920x1080 design space. A standard 16:9 deck is
 * 960x540pt, so the factor is exactly 0.5. All conversion happens
 * at the boundary: P() going in, nothing coming out.
 *
 * CORNER RADIUS. PowerPoint does not store a radius. It stores an
 * adjustment fraction, and the drawn radius is:
 *
 *     radius = adjustment x min(width, height)
 *
 * which is why identical adjustments look different on differently
 * sized boxes. We invert it per shape.
 * --------------------------------------------------------------- */

'use strict';

/* ── config ─────────────────────────────────────────────────────
 * The one place to change the plugin's name and colours.
 *   name         shown in the panel header and the browser title.
 *                (mcpp_v1.xml has its own copy for the ribbon labels.)
 *   shapePrefix  start of every name this panel gives a shape, e.g.
 *                "MCPP box r1c1". The footer's duplicate check uses it.
 *   defaultFill  fill for generated boxes.
 *   swatches     primary = the four main chips; secondary = small dots. */
const CONFIG = {
  name: 'MCPP_V1',
  shapePrefix: 'MCPP',
  defaultFill: '#DDD5D0',
  swatches: {
    primary: [
      { name: 'Yellow',    hex: '#FEF317' },
      { name: 'Warm grey', hex: '#DDD5D0' },
      { name: 'Black',     hex: '#000000' },
      { name: 'White',     hex: '#FFFFFF' }
    ],
    secondary: [
      { name: 'Violet', hex: '#8844FF' },
      { name: 'Pink',   hex: '#FF45C4' },
      { name: 'Coral',  hex: '#FF5638' },
      { name: 'Teal',   hex: '#00A696' },
      { name: 'Blue',   hex: '#4D6AFF' }
    ]
  }
};

const SLIDE_W_PT = 960;
const SLIDE_H_PT = 540;
const PX_TO_PT   = 0.5;          // 1920px design space -> 960pt slide

const P  = px => px * PX_TO_PT;  // px -> pt
const el = id => document.getElementById(id);

let ADJ_SUPPORTED = false;       // PowerPointApi 1.10
let ADJ_STRATEGY  = null;        // which setter actually worked
let LAST_ERROR    = null;        // most recent failure, surfaced in the status line

/* ── state ──────────────────────────────────────────────────── */

function readState() {
  const n = (id, fallback) => {
    const v = parseFloat(el(id).value);
    return Number.isFinite(v) ? v : fallback;
  };
  return {
    mTop:    n('mTop', 60),
    mRight:  n('mRight', 60),
    mBottom: n('mBottom', 60),
    mLeft:   n('mLeft', 60),
    gutterH: n('gutterH', 60),
    gutterV: n('gutterV', 60),
    radius:  n('radius', 30),
    cols:    Math.max(1, Math.round(n('cols', 3))),
    rows:    Math.max(1, Math.round(n('rows', 1))),
    fill:    el('fill').value || CONFIG.defaultFill
  };
}

/** The live area inside the margins, in points. */
function marginBox(s) {
  const left   = P(s.mLeft);
  const top    = P(s.mTop);
  const right  = SLIDE_W_PT - P(s.mRight);
  const bottom = SLIDE_H_PT - P(s.mBottom);
  return { left, top, right, bottom, width: right - left, height: bottom - top };
}

/* ── corner radius maths ────────────────────────────────────── */

function radiusToAdjustment(radiusPx, widthPt, heightPt) {
  const minSide = Math.min(widthPt, heightPt);
  if (!(minSide > 0)) return { adj: 0, clamped: false };
  let adj = P(radiusPx) / minSide;
  const clamped = adj > 0.5;
  if (adj > 0.5) adj = 0.5;
  if (adj < 0) adj = 0;
  return { adj, clamped };
}

/* Confirmed against PowerPoint 16.112 (PowerPointApi 1.10): the Adjustments
 * object exposes count / get(i) / set(i, v). Nothing else is kept here —
 * assigning a[0] directly appears to succeed but only writes a JS property
 * on the proxy and never reaches the shape. */
function setAdjustmentValue(shape, value) {
  const a = shape.adjustments;
  if (!a) throw new Error('no-adjustments');

  if (typeof a.set === 'function')     { a.set(0, value);     ADJ_STRATEGY = 'set(0, v)';     return; }
  if (typeof a.setItem === 'function') { a.setItem(0, value); ADJ_STRATEGY = 'setItem(0, v)'; return; }
  throw new Error('no-setter');
}

function roundRectType() {
  const G = PowerPoint.GeometricShapeType || {};
  return G.roundRectangle || G.roundRect || 'RoundRectangle';
}

function isGeometric(type) {
  return type === (PowerPoint.ShapeType && PowerPoint.ShapeType.geometricShape) ||
         type === 'GeometricShape' || type === 'geometricShape';
}

/* ── shared helpers ─────────────────────────────────────────── */

function currentSlide(ctx) {
  return ctx.presentation.getSelectedSlides().getItemAt(0);
}

async function selectedShapes(ctx) {
  const sel = ctx.presentation.getSelectedShapes();
  sel.load('items/id,items/left,items/top,items/width,items/height,items/type');
  await ctx.sync();
  return sel.items;
}

/* ── 1. corner radius ───────────────────────────────────────── */

async function applyRadius(multiplier) {
  const s = readState();
  const target = Math.max(0, s.radius * multiplier);

  if (!ADJ_SUPPORTED) {
    return say(`Corner radius needs PowerPointApi 1.10. Your Office build is older — update it, then reopen the panel.`, 'bad');
  }

  let ids = [];
  try {
    await PowerPoint.run(async ctx => {
      const items = await selectedShapes(ctx);
      ids = items.map(sh => ({ id: sh.id, type: sh.type }));
    });
  } catch (e) {
    return say('Could not read the selection. Click a shape on the slide first.', 'bad');
  }

  if (!ids.length) return say('Nothing selected. Pick one or more shapes.', 'bad');

  let set = 0, rebuilt = 0, clamped = 0, skipped = 0;

  for (const { id, type } of ids) {
    const r = await setRadiusOnShape(id, target, type);
    if (r.rebuilt) rebuilt++;
    else if (r.ok) set++;
    else skipped++;
    if (r.clamped) clamped++;
  }

  const bits = [];
  if (set)     bits.push(`${set} set`);
  if (rebuilt) bits.push(`${rebuilt} rebuilt`);
  if (skipped) bits.push(`${skipped} skipped`);
  let msg = `${target}px radius — ${bits.join(', ')}.`;
  if (clamped) msg += ` ${clamped} hit the ceiling of half the short side.`;
  if (skipped && LAST_ERROR) msg += `  \u2014  ${LAST_ERROR.slice(0, 120)}`;
  say(msg, skipped ? 'bad' : 'good');
}

async function setRadiusOnShape(id, radiusPx, type) {
  try {
    return await PowerPoint.run(async ctx => {
      const shape = currentSlide(ctx).shapes.getItem(id);
      shape.load('width,height');
      await ctx.sync();
      const { adj, clamped } = radiusToAdjustment(radiusPx, shape.width, shape.height);
      setAdjustmentValue(shape, adj);
      await ctx.sync();
      return { ok: true, clamped, rebuilt: false };
    });
  } catch (e) {
    LAST_ERROR = `${type}: ${String(e.message || e)}`;
    // Rebuild only a plain geometric shape. NEVER a picture — image data can't be
    // re-embedded, so delete-and-recreate would destroy it.
    if (!isGeometric(type)) return { ok: false, rebuilt: false, clamped: false, type };
    return rebuildAsRounded(id, radiusPx);
  }
}

async function rebuildAsRounded(id, radiusPx) {
  const snap = { fill: null, alpha: null, line: null, text: null, font: null, vAlign: null };

  // geometry — always readable
  try {
    await PowerPoint.run(async ctx => {
      const sh = currentSlide(ctx).shapes.getItem(id);
      sh.load('left,top,width,height,rotation,name');
      await ctx.sync();
      snap.left = sh.left; snap.top = sh.top;
      snap.width = sh.width; snap.height = sh.height;
      snap.rotation = sh.rotation || 0; snap.name = sh.name;
    });
  } catch (e) {
    return { ok: false, rebuilt: false, clamped: false };
  }

  // fill + line — may legitimately not exist
  try {
    await PowerPoint.run(async ctx => {
      const sh = currentSlide(ctx).shapes.getItem(id);
      sh.fill.load('foregroundColor,transparency');
      sh.lineFormat.load('color,weight,visible');
      await ctx.sync();
      snap.fill  = sh.fill.foregroundColor;
      snap.alpha = sh.fill.transparency;
      snap.line  = { color: sh.lineFormat.color, weight: sh.lineFormat.weight, visible: sh.lineFormat.visible };
    });
  } catch (e) { /* leave defaults */ }

  // text
  try {
    await PowerPoint.run(async ctx => {
      const sh = currentSlide(ctx).shapes.getItem(id);
      sh.textFrame.textRange.load('text');
      sh.textFrame.textRange.font.load('name,size,color,bold,italic');
      sh.textFrame.load('verticalAlignment');
      await ctx.sync();
      snap.text = sh.textFrame.textRange.text;
      const f = sh.textFrame.textRange.font;
      snap.font = { name: f.name, size: f.size, color: f.color, bold: f.bold, italic: f.italic };
      snap.vAlign = sh.textFrame.verticalAlignment;
    });
  } catch (e) { /* no text */ }

  // delete, recreate, reapply
  try {
    return await PowerPoint.run(async ctx => {
      const slide = currentSlide(ctx);
      slide.shapes.getItem(id).delete();

      const shape = slide.shapes.addGeometricShape(roundRectType(), {
        left: snap.left, top: snap.top, width: snap.width, height: snap.height
      });

      if (snap.name) { try { shape.name = snap.name; } catch (e) {} }
      if (snap.rotation) { try { shape.rotation = snap.rotation; } catch (e) {} }
      if (snap.fill) { try { shape.fill.setSolidColor(snap.fill); } catch (e) {} }
      if (typeof snap.alpha === 'number') { try { shape.fill.transparency = snap.alpha; } catch (e) {} }
      if (snap.line) {
        try {
          if (snap.line.visible === false) shape.lineFormat.visible = false;
          else {
            if (snap.line.color)  shape.lineFormat.color  = snap.line.color;
            if (snap.line.weight) shape.lineFormat.weight = snap.line.weight;
          }
        } catch (e) {}
      }
      if (snap.text) {
        try {
          shape.textFrame.textRange.text = snap.text;
          const f = snap.font || {};
          if (f.name)  shape.textFrame.textRange.font.name  = f.name;
          if (f.size)  shape.textFrame.textRange.font.size  = f.size;
          if (f.color) shape.textFrame.textRange.font.color = f.color;
          if (typeof f.bold === 'boolean')   shape.textFrame.textRange.font.bold = f.bold;
          if (typeof f.italic === 'boolean') shape.textFrame.textRange.font.italic = f.italic;
          if (snap.vAlign) shape.textFrame.verticalAlignment = snap.vAlign;
        } catch (e) {}
      }

      const { adj, clamped } = radiusToAdjustment(radiusPx, snap.width, snap.height);
      try { setAdjustmentValue(shape, adj); } catch (e) {}

      await ctx.sync();
      return { ok: true, rebuilt: true, clamped };
    });
  } catch (e) {
    return { ok: false, rebuilt: false, clamped: false };
  }
}

/* ── 2. align ───────────────────────────────────────────────── */

/** Where each tool works: 'page' (inside the margins) or 'selection'. */
const MODE = { align: 'page', space: 'page', fill: 'page' };

/** Bounding box of a set of shapes, in points. */
function selectionBox(items) {
  const left   = Math.min(...items.map(s => s.left));
  const top    = Math.min(...items.map(s => s.top));
  const right  = Math.max(...items.map(s => s.left + s.width));
  const bottom = Math.max(...items.map(s => s.top + s.height));
  return { left, top, right, bottom, width: right - left, height: bottom - top };
}

async function align(mode) {
  try {
    await PowerPoint.run(async ctx => {
      const items = await selectedShapes(ctx);
      if (!items.length) throw new Error('none');

      const toSelection = MODE.align === 'selection';
      if (toSelection && items.length < 2) throw new Error('one');

      const box = toSelection ? selectionBox(items) : marginBox(readState());

      items.forEach(sh => {
        switch (mode) {
          case 'left':    sh.left = box.left; break;
          case 'right':   sh.left = box.right - sh.width; break;
          case 'centerH': sh.left = box.left + (box.width - sh.width) / 2; break;
          case 'top':     sh.top  = box.top; break;
          case 'bottom':  sh.top  = box.bottom - sh.height; break;
          case 'centerV': sh.top  = box.top + (box.height - sh.height) / 2; break;
        }
      });
      await ctx.sync();
      say(`${items.length} shape${items.length > 1 ? 's' : ''} aligned to ${toSelection ? 'the selection' : 'the page'}.`, 'good');
    });
  } catch (e) {
    const m = {
      none: 'Nothing selected. Pick one or more shapes.',
      one: 'Aligning to the selection needs at least two shapes.'
    };
    say(m[e.message] || 'Could not align that selection.', 'bad');
  }
}

/* ── 3. generate boxes ──────────────────────────────────────── */

async function generateGrid() {
  const s = readState();
  const box = marginBox(s);
  const bw = (box.width  - (s.cols - 1) * P(s.gutterH)) / s.cols;
  const bh = (box.height - (s.rows - 1) * P(s.gutterV)) / s.rows;

  if (bw <= 0 || bh <= 0) {
    return say('Margins and gutters leave no room. Reduce them or use fewer columns.', 'bad');
  }

  try {
    await PowerPoint.run(async ctx => {
      const slide = currentSlide(ctx);
      const RR = roundRectType();

      for (let r = 0; r < s.rows; r++) {
        for (let c = 0; c < s.cols; c++) {
          const shape = slide.shapes.addGeometricShape(RR, {
            left:  box.left + c * (bw + P(s.gutterH)),
            top:   box.top  + r * (bh + P(s.gutterV)),
            width: bw,
            height: bh
          });
          try { shape.name = `${CONFIG.shapePrefix} box r${r + 1}c${c + 1}`; } catch (e) {}
          try { shape.fill.setSolidColor(s.fill); } catch (e) {}
          try { shape.lineFormat.visible = false; } catch (e) {}
          try { setAdjustmentValue(shape, radiusToAdjustment(s.radius, bw, bh).adj); } catch (e) {}
        }
      }
      await ctx.sync();
      const n = s.cols * s.rows;
      say(`${n} box${n > 1 ? 'es' : ''} at ${Math.round(bw / PX_TO_PT)}\u00d7${Math.round(bh / PX_TO_PT)}px.`, 'good');
    });
  } catch (e) {
    say('Could not generate. Make sure a slide is open and try again.', 'bad');
  }
}

/* ── 4. fill to gutters ─────────────────────────────────────── */

/* Resizes the selected shapes so they fill a span with the gutter between
 * each. Page mode fills the area inside the margins; selection mode fills the
 * selection's own extent (its first and last edges don't move). */
async function fit(axis) {
  const s = readState();
  const horizontal = axis === 'row';
  const gutter = P(horizontal ? s.gutterH : s.gutterV);
  const toSelection = MODE.fill === 'selection';

  try {
    await PowerPoint.run(async ctx => {
      const items = await selectedShapes(ctx);
      if (!items.length) throw new Error('none');
      if (toSelection && items.length < 2) throw new Error('few');

      const ordered = items.slice().sort((a, b) => horizontal ? a.left - b.left : a.top - b.top);
      const n = ordered.length;

      let start, extent;
      if (toSelection) {
        start = horizontal ? ordered[0].left : ordered[0].top;
        const end = Math.max(...ordered.map(sh => horizontal ? sh.left + sh.width : sh.top + sh.height));
        extent = end - start;
      } else {
        const box = marginBox(s);
        start = horizontal ? box.left : box.top;
        extent = horizontal ? box.width : box.height;
      }

      const span = (extent - (n - 1) * gutter) / n;
      if (span <= 0) throw new Error('tight');

      ordered.forEach((sh, i) => {
        if (horizontal) { sh.left = start + i * (span + gutter); sh.width  = span; }   // height and top untouched
        else            { sh.top  = start + i * (span + gutter); sh.height = span; }   // width and left untouched
      });
      await ctx.sync();

      // Sizes changed, so the adjustment fraction now means a different
      // radius. Re-solve it. Square-cornered shapes just fall through.
      try {
        ordered.forEach(sh => {
          const w = horizontal ? span : sh.width;
          const h = horizontal ? sh.height : span;
          setAdjustmentValue(sh, radiusToAdjustment(s.radius, w, h).adj);
        });
        await ctx.sync();
      } catch (e) { /* some or all have square corners */ }

      say(`${n} shape${n > 1 ? 's' : ''} filled ${toSelection ? 'the selection' : 'the page'} at ${Math.round(span / PX_TO_PT)}px.`, 'good');
    });
  } catch (e) {
    const m = {
      none: 'Nothing selected. Pick one or more shapes.',
      few: 'Filling the selection needs at least two shapes.',
      tight: 'Gutters leave no room for that many shapes.'
    };
    say(m[e.message] || 'Could not fill that selection.', 'bad');
  }
}

/* ── 4b. space evenly ───────────────────────────────────────── */

/* Moves shapes only, never resizes them, so every gap is equal.
 * Selection mode: the first and last shapes stay put and the ones between
 * spread out evenly. Page mode: the page margins are the outer points, so the
 * gap before the first shape and after the last equals the gaps between. */
async function space(axis) {
  const horizontal = axis === 'row';
  const toSelection = MODE.space === 'selection';

  try {
    await PowerPoint.run(async ctx => {
      const items = await selectedShapes(ctx);
      if (!items.length) throw new Error('none');
      if (toSelection && items.length < 3) throw new Error('few');

      const ordered = items.slice().sort((a, b) => horizontal ? a.left - b.left : a.top - b.top);
      const n = ordered.length;
      const size = sh => horizontal ? sh.width : sh.height;

      let start, extent, gaps;
      if (toSelection) {
        start = horizontal ? ordered[0].left : ordered[0].top;
        const end = Math.max(...ordered.map(sh => horizontal ? sh.left + sh.width : sh.top + sh.height));
        extent = end - start;
        gaps = n - 1;
      } else {
        const box = marginBox(readState());
        start = horizontal ? box.left : box.top;
        extent = horizontal ? box.width : box.height;
        gaps = n + 1;
      }

      const total = ordered.reduce((sum, sh) => sum + size(sh), 0);
      const gap = (extent - total) / gaps;
      if (gap < 0) throw new Error('tight');

      let pos = toSelection ? start : start + gap;
      ordered.forEach(sh => {
        if (horizontal) sh.left = pos; else sh.top = pos;
        pos += size(sh) + gap;
      });
      await ctx.sync();

      say(`${n} shape${n > 1 ? 's' : ''} spaced evenly, ${Math.round(gap / PX_TO_PT)}px apart.`, 'good');
    });
  } catch (e) {
    const m = {
      none: 'Nothing selected. Pick one or more shapes.',
      few: 'Spacing the selection needs at least three shapes.',
      tight: 'The shapes are too big to space evenly in that area.'
    };
    say(m[e.message] || 'Could not space that selection.', 'bad');
  }
}

/* ── 5. fill selection ──────────────────────────────────────── */

async function setFill(hex) {
  try {
    await PowerPoint.run(async ctx => {
      const items = await selectedShapes(ctx);
      if (!items.length) throw new Error('none');
      items.forEach(sh => { try { sh.fill.setSolidColor(hex); } catch (e) {} });
      await ctx.sync();
      say(`${items.length} shape${items.length > 1 ? 's' : ''} filled ${hex}.`, 'good');
    });
  } catch (e) {
    say(e.message === 'none' ? 'Nothing selected. Pick one or more shapes.' : 'Could not apply that fill.', 'bad');
  }
}

/* ── 6. position readout, copy and paste ────────────────────── */

let CLIP = null;      // { x, y, w, h } in px, from the last Copy
let LAST_BOX = null;  // last values read from the selection, for relative maths
let polling = false;

const PROP = { posX: 'left', posY: 'top', posW: 'width', posH: 'height' };

let HAS_SELECTION = false;

function setReadout(box, count) {
  LAST_BOX = box;
  HAS_SELECTION = count > 0;
  const badge = el('selCount');
  badge.textContent = count > 1 ? `· ${count} selected, showing first` : (count === 1 ? '' : '· nothing selected');

  ['posX', 'posY', 'posW', 'posH'].forEach(id => {
    const input = el(id);
    if (document.activeElement === input) return;   // don't fight the user mid-type
    if (!box) { input.value = ''; input.classList.remove('live'); return; }
    const v = { posX: box.x, posY: box.y, posW: box.w, posH: box.h }[id];
    const rounded = String(Math.round(v));
    if (input.value !== rounded) input.value = rounded;
    input.classList.add('live');
  });
}

/* The readout is driven by PowerPoint's own selection event — no timers, no
 * background traffic. Two things to know about that event in a shared file:
 *
 *  - It also fires when OTHER people change their selection while co-authoring.
 *    That's by design, per Microsoft. Hence the debounce: a burst of remote
 *    activity collapses into one read instead of dozens.
 *  - It has regressed on some Mac builds in the past, so panel hover is kept as
 *    a fallback. It costs nothing — it only fires when you reach for the panel. */
let pollTimer = null;

function schedulePoll(delay) {
  if (pollTimer) clearTimeout(pollTimer);
  pollTimer = setTimeout(pollSelection, delay || 250);
}

async function pollSelection() {
  if (polling || document.hidden) return;
  polling = true;
  try {
    await PowerPoint.run(async ctx => {
      const items = await selectedShapes(ctx);
      if (!items.length) return setReadout(null, 0);
      const s = items[0];
      setReadout({
        x: s.left / PX_TO_PT, y: s.top / PX_TO_PT,
        w: s.width / PX_TO_PT, h: s.height / PX_TO_PT
      }, items.length);
    });
  } catch (e) {
    /* mid-transition between slides, or nothing open */
  } finally {
    polling = false;
  }
}

function copyPosition() {
  const n = id => { const v = fieldValue(id); return v === null ? NaN : v; };
  const box = { x: n('posX'), y: n('posY'), w: n('posW'), h: n('posH') };
  if (!Object.values(box).every(Number.isFinite)) {
    return say('Select a shape first, then Copy.', 'bad');
  }
  CLIP = box;
  el('pastePos').disabled = false;
  el('pasteSize').disabled = false;
  el('clipRow').hidden = false;
  el('clipPos').textContent = `${Math.round(box.x)}, ${Math.round(box.y)}`;
  el('clipSize').textContent = `${Math.round(box.w)} \u00d7 ${Math.round(box.h)}`;
  say(`Copied ${Math.round(box.x)},${Math.round(box.y)} at ${Math.round(box.w)}\u00d7${Math.round(box.h)}px.`, 'good');
}

/** props: subset of left/top/width/height to write, in px. */
async function applyToSelection(props, label) {
  const resize = 'width' in props || 'height' in props;
  const radius = readState().radius;

  try {
    await PowerPoint.run(async ctx => {
      const items = await selectedShapes(ctx);
      if (!items.length) throw new Error('none');

      items.forEach(sh => {
        Object.entries(props).forEach(([k, vPx]) => { sh[k] = P(vPx); });
      });
      await ctx.sync();

      if (resize) {
        try {
          items.forEach(sh => {
            const w = 'width'  in props ? P(props.width)  : sh.width;
            const h = 'height' in props ? P(props.height) : sh.height;
            setAdjustmentValue(sh, radiusToAdjustment(radius, w, h).adj);
          });
          await ctx.sync();
        } catch (e) { /* square corners, nothing to re-solve */ }
      }

      say(`${label} applied to ${items.length} shape${items.length > 1 ? 's' : ''}.`, 'good');
    });
  } catch (e) {
    say(e.message === 'none' ? 'Nothing selected. Pick one or more shapes.' : 'Could not apply that.', 'bad');
  }
}

function pastePosition() {
  if (!CLIP) return;
  applyToSelection({ left: CLIP.x, top: CLIP.y }, 'Position');
}

function pasteSize() {
  if (!CLIP) return;
  applyToSelection({ width: CLIP.w, height: CLIP.h }, 'Size');
}

/* Accepts plain numbers and arithmetic: 73, 43+30, 43 + 30, *2, /2, (40+8)*2.
 * A leading operator applies to the field's current value, so on a box of 43,
 * typing "+30" gives 73 — as does appending it to make "43+30".
 * A leading minus is therefore relative: "-30" subtracts. For an absolute
 * negative position, type "0-30". */
function evalExpr(raw, current) {
  let expr = String(raw).trim().replace(/\u2212/g, '-');   // tolerate a typographic minus
  if (!expr) return null;

  if (/^[+\-*/]/.test(expr)) {
    if (!Number.isFinite(current)) return null;
    expr = current + expr;
  }
  // Whitelist: digits, dot, whitespace, four operators, brackets. Nothing else can
  // reach the evaluator, so no identifiers or calls are possible.
  if (!/^[\d\s.+\-*/()]+$/.test(expr)) return null;

  try {
    const v = Function('"use strict";return (' + expr + ')')();
    return Number.isFinite(v) ? v : null;
  } catch (e) {
    return null;
  }
}

const BOX_KEY = { posX: 'x', posY: 'y', posW: 'w', posH: 'h' };

/** Current value of a position field, resolving any expression in it. */
function fieldValue(id) {
  const current = LAST_BOX ? LAST_BOX[BOX_KEY[id]] : NaN;
  return evalExpr(el(id).value, current);
}

function applyTypedField(id) {
  const current = LAST_BOX ? LAST_BOX[BOX_KEY[id]] : NaN;
  const v = evalExpr(el(id).value, current);
  if (v === null) return say(`Couldn't work out "${el(id).value}".`, 'bad');

  el(id).value = String(Math.round(v));       // show the result, not the sum
  const prop = PROP[id];
  applyToSelection({ [prop]: v }, prop.charAt(0).toUpperCase() + prop.slice(1));
}

/* ── 7. inset ───────────────────────────────────────────────── */

/* Moves the selection inward from the named side. Size is untouched, so
 * there's no radius to re-solve and nothing can collapse. */
const NUDGE = {
  top:    (d, sh) => { sh.top  += d; },
  bottom: (d, sh) => { sh.top  -= d; },
  left:   (d, sh) => { sh.left += d; },
  right:  (d, sh) => { sh.left -= d; }
};

async function inset(side) {
  const move = NUDGE[side];
  const amount = parseFloat(el('insetAmt').value);
  if (!Number.isFinite(amount)) return say('Enter an inset amount first.', 'bad');

  const d = P(amount);

  try {
    await PowerPoint.run(async ctx => {
      const items = await selectedShapes(ctx);
      if (!items.length) throw new Error('none');

      items.forEach(sh => move(d, sh));
      await ctx.sync();

      say(`${items.length} shape${items.length > 1 ? 's' : ''} moved ${amount}px in from the ${side}.`, 'good');
    });
  } catch (e) {
    say(e.message === 'none' ? 'Nothing selected. Pick one or more shapes.' : 'Could not move that selection.', 'bad');
  }
}

/* ── diagnostics ────────────────────────────────────────────── */

/* Probes, in the shape's own document, every plausible way of writing an
 * adjustment value — and reports which one PowerPoint actually accepts.
 * Output goes into the panel, so no developer console is needed. */
async function setTransparency(pct) {
  const value = Math.min(1, Math.max(0, pct / 100));
  let applied = 0, refused = 0;

  try {
    await PowerPoint.run(async ctx => {
      const items = await selectedShapes(ctx);
      if (!items.length) throw new Error('none');
      items.forEach(sh => {
        try { sh.fill.transparency = value; applied++; } catch (e) { refused++; }
      });
      await ctx.sync();
    });
  } catch (e) {
    return say(e.message === 'none'
      ? 'Nothing selected. Pick one or more shapes.'
      : 'Could not set transparency — the selection may have no fill.', 'bad');
  }

  let msg = `${applied} shape${applied === 1 ? '' : 's'} at ${pct}% transparent.`;
  if (refused) msg += ` ${refused} had no fill.`;
  say(msg, refused ? 'bad' : 'good');
}

async function removeFill() {
  try {
    await PowerPoint.run(async ctx => {
      const items = await selectedShapes(ctx);
      if (!items.length) throw new Error('none');
      items.forEach(sh => { try { sh.fill.clear(); } catch (e) {} });
      await ctx.sync();
      say(`Fill removed from ${items.length} shape${items.length === 1 ? '' : 's'}.`, 'good');
    });
  } catch (e) {
    say(e.message === 'none' ? 'Nothing selected. Pick one or more shapes.' : 'Could not remove the fill.', 'bad');
  }
}

async function removeStroke() {
  try {
    await PowerPoint.run(async ctx => {
      const items = await selectedShapes(ctx);
      if (!items.length) throw new Error('none');
      items.forEach(sh => { try { sh.lineFormat.visible = false; } catch (e) {} });
      await ctx.sync();
      say(`Stroke removed from ${items.length} shape${items.length === 1 ? '' : 's'}.`, 'good');
    });
  } catch (e) {
    say(e.message === 'none' ? 'Nothing selected. Pick one or more shapes.' : 'Could not remove the stroke.', 'bad');
  }
}

/* ── 8. footer ──────────────────────────────────────────────── */

const FOOTER_LEFT   = `${CONFIG.shapePrefix} footer section`;
const FOOTER_MID    = `${CONFIG.shapePrefix} footer info`;
const FOOTER_H      = 40;    // px band for the text boxes
const FOOTER_GAP    = 40;    // px between the three boxes

function styleFooterBox(shape, align) {
  const tf = shape.textFrame;
  const tr = tf.textRange;
  try { tr.font.name = 'Arial'; } catch (e) {}
  try { tr.font.size = 11; } catch (e) {}
  try { tr.font.bold = false; } catch (e) {}
  try { tr.font.italic = false; } catch (e) {}
  try { tr.paragraphFormat.horizontalAlignment = align; } catch (e) {}
  try { tf.verticalAlignment = 'Middle'; } catch (e) {}
  // zero the internal padding so the text sits exactly on the page margin
  ['leftMargin', 'rightMargin', 'topMargin', 'bottomMargin'].forEach(k => {
    try { tf[k] = 0; } catch (e) {}
  });
  try { shape.fill.clear(); } catch (e) {}
  try { shape.lineFormat.visible = false; } catch (e) {}
}

async function generateFooter() {
  const s = readState();

  // Vertical centre of the bottom margin band: half the margin up from the edge.
  const centreY = P(SLIDE_H_PT / PX_TO_PT - s.mBottom / 2);
  const top = centreY - P(FOOTER_H) / 2;
  const h = P(FOOTER_H);

  const usable = P(SLIDE_W_PT / PX_TO_PT - s.mLeft - s.mRight);
  const w = (usable - 2 * P(FOOTER_GAP)) / 3;
  if (w <= 0) return say('Margins leave no room for a footer.', 'bad');

  const slideW = SLIDE_W_PT;
  // Three-column geometry is kept even though only two boxes are made, so the
  // right slot stays free for PowerPoint's own slide-number placeholder.
  const boxes = [
    { name: FOOTER_LEFT, text: 'Section title',  align: 'Left',   left: P(s.mLeft) },
    { name: FOOTER_MID,  text: 'Secondary info', align: 'Center', left: (slideW - w) / 2 }
  ];

  let created = 0;

  try {
    await PowerPoint.run(async ctx => {
      const slide = currentSlide(ctx);
      slide.shapes.load('items/name');
      await ctx.sync();

      const existing = new Set(slide.shapes.items.map(sh => sh.name));

      boxes.forEach(b => {
        if (existing.has(b.name)) return;          // idempotent: never duplicates
        const shape = slide.shapes.addTextBox(b.text, {
          left: b.left, top, width: w, height: h
        });
        try { shape.name = b.name; } catch (e) {}
        styleFooterBox(shape, b.align);
        created++;
      });
      await ctx.sync();
    });
  } catch (e) {
    return say('Could not generate the footer. Make sure a slide is open.', 'bad');
  }

  say(created ? `Footer inserted (${created} boxes).` : 'Footer already on this slide.', 'good');
}

/* ── UI plumbing ────────────────────────────────────────────── */

function say(msg, tone) {
  const s = el('status');
  s.textContent = msg;
  s.dataset.tone = tone || '';
  // A successful edit means the readout is now stale. Defer so this lands
  // after the PowerPoint.run that called us has finished.
  if (tone === 'good') schedulePoll(150);
}

/* Ties a set of numeric fields together behind a chain-link toggle. */
function linkGroup(buttonId, fieldIds) {
  const btn = el(buttonId);
  let on = btn.classList.contains('link--on');
  const sync = src => {
    if (!on) return;
    fieldIds.forEach(id => { if (id !== src) el(id).value = el(src).value; });
  };
  btn.addEventListener('click', () => {
    on = !on;
    btn.classList.toggle('link--on', on);
    btn.setAttribute('aria-pressed', String(on));
    if (on) sync(fieldIds[0]);
  });
  fieldIds.forEach(id => el(id).addEventListener('input', () => sync(id)));
}

/* Draws the name, the swatches and the default fill from CONFIG. Runs before
 * the click handlers are bound, so the new buttons are picked up. */
function applyConfig() {
  document.title = CONFIG.name;
  el('appName').textContent = CONFIG.name;
  el('fill').value = CONFIG.defaultFill.toLowerCase();

  const main = el('swatchesMain');
  CONFIG.swatches.primary.forEach(sw => {
    const b = document.createElement('button');
    b.type = 'button';
    b.className = 'chip';
    b.dataset.fill = sw.hex;
    b.title = `${sw.name} ${sw.hex}`;
    b.style.setProperty('--c', sw.hex);
    b.appendChild(document.createElement('span'));
    const label = document.createElement('i');
    label.textContent = sw.name;
    b.appendChild(label);
    main.appendChild(b);
  });

  const more = el('swatchesMore');
  CONFIG.swatches.secondary.forEach(sw => {
    const b = document.createElement('button');
    b.type = 'button';
    b.className = 'dot';
    b.dataset.fill = sw.hex;
    b.title = `${sw.name} ${sw.hex}`;
    b.setAttribute('aria-label', `${sw.name} ${sw.hex}`);
    b.style.setProperty('--c', sw.hex);
    more.appendChild(b);
  });
}

function refreshLabels() {
  const s = readState();
  el('lblS').textContent = Math.round(s.radius * 0.5);
  el('lblM').textContent = Math.round(s.radius);
  el('lblL').textContent = Math.round(s.radius * 2);
  el('gridCount').textContent = s.cols * s.rows;
}

Office.onReady(info => {
  if (info.host !== Office.HostType.PowerPoint) {
    el('unsupported').hidden = false;
    return;
  }
  el('app').hidden = false;
  applyConfig();

  ADJ_SUPPORTED = Office.context.requirements.isSetSupported('PowerPointApi', '1.10');
  if (!ADJ_SUPPORTED) {
    say('Align and box generation are ready. Corner radius needs a newer Office build.', 'bad');
  }

  document.querySelectorAll('[data-align]').forEach(b =>
    b.addEventListener('click', () => align(b.dataset.align)));

  document.querySelectorAll('[data-mode-group]').forEach(b =>
    b.addEventListener('click', () => {
      const group = b.dataset.modeGroup;
      MODE[group] = b.dataset.mode;
      document.querySelectorAll(`[data-mode-group="${group}"]`).forEach(o =>
        o.classList.toggle('seg__btn--on', o === b));
    }));

  document.querySelectorAll('[data-radius]').forEach(b =>
    b.addEventListener('click', () => applyRadius(parseFloat(b.dataset.radius))));

  document.querySelectorAll('[data-fill]').forEach(b =>
    b.addEventListener('click', () => setFill(b.dataset.fill)));

  // 'input' updates the label only; 'change' fires once on release, so dragging
  // the slider doesn't write to the document on every pixel.
  el('alpha').addEventListener('input', () => {
    el('alphaVal').textContent = `${el('alpha').value}%`;
  });
  el('alpha').addEventListener('change', () => setTransparency(parseInt(el('alpha').value, 10)));

  el('noStroke').addEventListener('click', removeStroke);
  el('noFill').addEventListener('click', removeFill);

  // Info pop-ups: hover (CSS, after a delay) or click to pin open. Only one is
  // open at a time. Click elsewhere or Esc closes.
  const infoBoxes = Array.from(document.querySelectorAll('.info'));
  const closeInfo = except => infoBoxes.forEach(box => {
    if (box === except) return;
    box.classList.remove('info--open');
    box.querySelector('.info__btn').setAttribute('aria-expanded', 'false');
  });
  infoBoxes.forEach(box => {
    const btn = box.querySelector('.info__btn');
    btn.addEventListener('click', ev => {
      ev.stopPropagation();
      closeInfo(box);
      const open = !box.classList.contains('info--open');
      box.classList.toggle('info--open', open);
      btn.setAttribute('aria-expanded', String(open));
    });
  });
  document.addEventListener('click', () => closeInfo(null));
  document.addEventListener('keydown', ev => { if (ev.key === 'Escape') closeInfo(null); });

  document.querySelectorAll('[data-inset]').forEach(b =>
    b.addEventListener('click', () => inset(b.dataset.inset)));

  el('copyPos').addEventListener('click', copyPosition);
  el('pastePos').addEventListener('click', pastePosition);
  el('pasteSize').addEventListener('click', pasteSize);

  // Event-driven readout. No timers.
  try {
    Office.context.document.addHandlerAsync(
      Office.EventType.DocumentSelectionChanged,
      () => schedulePoll(250)
    );
  } catch (e) { /* fall back to hover below */ }

  // Fallback only — fires when you reach for the panel, never in the background.
  el('app').addEventListener('mouseenter', () => schedulePoll(0));

  pollSelection();

  el('makeGrid').addEventListener('click', generateGrid);
  el('makeFooter').addEventListener('click', generateFooter);
  el('fillRow').addEventListener('click', () => fit('row'));
  el('fillCol').addEventListener('click', () => fit('column'));
  el('spaceRow').addEventListener('click', () => space('row'));
  el('spaceCol').addEventListener('click', () => space('column'));

  linkGroup('linkMargins', ['mTop', 'mRight', 'mBottom', 'mLeft']);
  linkGroup('linkGutters', ['gutterH', 'gutterV']);

  el('applyRadius').addEventListener('click', () => applyRadius(1));

  /* Enter commits and drops focus everywhere. Where a field maps to a live
   * operation it also applies — but only when something is selected, so
   * tabbing through settings never throws an error at you. */
  document.querySelectorAll('.num input').forEach(input => {
    input.addEventListener('keydown', ev => {
      if (ev.key !== 'Enter') return;
      ev.preventDefault();
      if (PROP[input.id]) applyTypedField(input.id);
      else if (input.id === 'radius' && HAS_SELECTION) applyRadius(1);
      input.blur();
    });
  });

  ['radius', 'cols', 'rows'].forEach(id => el(id).addEventListener('input', refreshLabels));
  refreshLabels();
});
