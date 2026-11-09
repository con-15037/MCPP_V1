# MCPP_V1

A PowerPoint task pane for quick layout work. Margins, alignment, spacing, corner
radius, fills and grids of boxes, all in the design space you actually think in
(1920 × 1080).

It is plain HTML, CSS and JavaScript using Office.js, plus an XML manifest. There is no
build step and no framework. It runs in desktop PowerPoint for Mac (confirmed on 16.112).

---

## What it does

| Section | What it does |
|---|---|
| **Page** | Sets margins and gutters. Each has a link/unlink chain toggle. |
| **Align** | Aligns the selection to the page margins or to the selection's own bounds. |
| **Arrange** | Space horizontal, Space vertical, Fill horizontal, Fill vertical. |
| **Position** | Live X, Y, W, H readout. Fields accept arithmetic. Position and size can be copied and pasted. |
| **Corner radius** | Applies S, M or L (15, 30 or 60 px) with an Apply button. |
| **Inset** | Nudges a shape inward from the top, left, right or bottom. |
| **Fill** | Colour swatches, an alpha slider and a remove-stroke button. |
| **Generate boxes** | Builds a columns × rows grid of loose rounded boxes inside the margins. Also inserts a footer. |

Defaults: margins 60 px (top, left, right) and 80 px (bottom, unlinked); gutters 30 px,
linked; corner radius 30 px. Settings last for the session only. Nothing is saved.

---

## Install

### 1. Host the folder

The panel is a handful of static files, so any static host works. The simplest option is
Netlify:

- **Drag and drop:** go to **app.netlify.com/drop** and drop this folder onto the page.
- **Connected to this repo:** create a Netlify site from the repo with **no build
  command**, publishing the **repo root**. Every push then deploys automatically.

Either way you get a URL like `https://your-site-name.netlify.app`. The site must be
publicly readable, because PowerPoint loads the panel without signing in.

### 2. Point the manifest at it

Open `mcpp_v1.xml` and replace every `YOUR-SITE.netlify.app` with your host. There are
**nine**, each flagged with a comment. Find-and-replace does it in one go.

### 3. Sideload it (Mac)

Put the edited `mcpp_v1.xml` into:

```
~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef/
```

Create the `wef` folder if it isn't there. In Finder: **Go → Go to Folder**, then paste
the path.

### 4. Restart PowerPoint

**Home → Add-ins** and you will see **MCPP_V1 → Layout**. Click it and the panel opens.

### Updating later

Change the files and push (or drop the folder onto Netlify again), then close and reopen
the panel. The manifest only changes if the URL does, so there is no need to sideload
again.

### Sharing it

Send a teammate the edited `mcpp_v1.xml` and step 3. The code is already on your host.

---

## Security and permissions

- The manifest requests **ReadWriteDocument**: the panel can read and change the open
  deck.
- **Whoever controls the hosted files controls code that runs against open decks.** Turn on
  two-factor authentication for the hosting account, and for the account that can push to
  this repo.
- The panel makes **no network requests** of its own (no `fetch`, no XHR), so deck content
  never leaves the machine. The only remote load is `office.js` from Microsoft's CDN.
- Never put keys, tokens or passwords in this repo. Everything the panel loads is public.

---

## Debugging

Enable the web inspector once:

```
defaults write com.microsoft.Powerpoint OfficeWebAddinDeveloperExtras -bool true
```

Then right-click inside the panel → **Inspect Element**.

---

## Configuration

The top of `taskpane.js` has one `CONFIG` object:

- `name` is shown in the panel header and the browser title.
- `shapePrefix` starts the name of every shape the panel creates (`MCPP box r1c1`,
  `MCPP footer section`, `MCPP footer info`). The footer's no-duplicate check relies on it.
- `defaultFill` is the fill for generated boxes.
- `swatches` holds the four main colour chips and the small secondary dots.

To rename the tool, change `name` and `shapePrefix`, then the ribbon labels and display
name in `mcpp_v1.xml`.

---

## How it works

### Units

The Office API works in **points**. A standard 16:9 deck is 960 × 540 pt, which is a
1920 × 1080 design space at exactly **2 px = 1 pt**. Every input in the panel is in px;
conversion happens in one place (`P()` in `taskpane.js`).

For a deck that isn't standard widescreen, change `SLIDE_W_PT`, `SLIDE_H_PT` and
`PX_TO_PT` at the top of `taskpane.js`.

### Corner radius

PowerPoint doesn't store a radius. It stores an adjustment fraction, and the drawn radius
is:

```
radius = adjustment × min(width, height)
```

That is why the same adjustment draws differently on a wide box and a tall one. The panel
inverts it per shape:

```
adjustment = (radius × 0.5) / min(widthPt, heightPt)
```

- **The ceiling is half the short side.** The fraction caps at 0.5, so a 30 px radius on a
  44 px-tall box will clamp. The status line says when that happens.
- **Resizing a box changes its drawn radius.** PowerPoint recalculates from the fraction.
  Re-hit S/M/L after resizing. Fill and Space re-solve it for you.

The `Adjustments` object (PowerPointApi 1.10) exposes `count`, `get(i)` and `set(i, v)`.
Assigning `adjustments[0] = value` directly *appears* to work, because nothing throws, but
it only writes a property on the JavaScript proxy and never reaches the shape. Always use
`set(0, value)` and verify with `get(0)`.

### Square rectangles

A plain rectangle has no adjustment to set, and it can't be converted into a rounded
rectangle in place. When you hit S/M/L on one, the panel rebuilds it, carrying over its
position, size, rotation, name, fill, line, text and font.

Two caveats. The rebuilt shape lands at the **top of the z-order**, and only the first
font run is kept, so mixed formatting inside one box collapses to the first style. Only
plain geometric shapes are ever rebuilt. Pictures, tables, placeholders and text boxes are
skipped, not touched.

### Fill

The swatches apply a solid fill to the selection. The **Alpha** slider sets fill
transparency; it writes on release, not while dragging, so a drag is one document write
rather than a hundred. **Remove stroke** hides the outline.

Colour and transparency are independent: clicking a swatch doesn't reset alpha, and moving
the slider doesn't change colour. The slider is a control, not a readout. It doesn't
reflect the current selection's transparency, since that would mean an extra document read
on every selection change.

### Footer

**Insert footer** adds two Arial 11 pt text boxes, vertically centred in the bottom margin
band. The left box ("Section title") starts at the left margin; the centre box ("Secondary
info") is centred on the page. Internal text margins are zeroed so text sits exactly on
the page margin.

Widths are computed as three columns even though only two boxes are created, so the
right-hand slot stays clear for PowerPoint's own slide-number placeholder (Insert → Header
& Footer → Slide number, positioned in the Slide Master). That is a real self-updating
field.

**Why there is no page-number box:** `TextRange.InsertSlideNumber` exists in VBA and COM
only. Office.js has no equivalent, so an add-in can't create a field that updates itself,
only static text. Static numbers in a deck people reorder are worse than none.

The button is idempotent: run it twice on a slide and it won't duplicate. Boxes are named
`MCPP footer section` and `MCPP footer info`.

### Generated boxes

Loose shapes. No PowerPoint guides, grid snapping or placeholders are involved, just
rounded rectangles at computed coordinates, filled with `defaultFill` and no outline,
named `MCPP box r1c1` and so on.

### Inset

An amount plus **T / L / R / B**. Each button moves the selection inward from that side
without changing its size: T pushes down, B pushes up, L pushes right, R pushes left.
Negative amounts move the other way. Nothing is resized, so there is no radius to
re-solve.

### Align

The toggle in the Align header switches the reference frame. **Margins** aligns every
selected shape to the page margins. **Selection** aligns them to the bounding box of the
selection, which needs at least two shapes.

### Position readout

Shows X, Y, W and H in px, driven by `Office.EventType.DocumentSelectionChanged`. There
are no timers and no background traffic: the panel reads the document only when the
selection changes, plus once after any edit it makes itself.

This matters in shared decks. An earlier version polled every 600 ms, which competed with
co-authoring sync and produced spurious "not synced" warnings. **Don't reintroduce a
timer.**

Two quirks of that event:

- It **also fires when other people change their selection** while co-authoring. Microsoft
  says that is by design, so the handler is debounced 250 ms to collapse remote bursts
  into a single read.
- It has regressed on some Mac builds in the past, so panel hover is kept as a fallback.
  It costs nothing, since it only fires when you reach for the panel.

**Copy** stores the current box. **Paste position** moves every selected shape to that
origin; **Paste size** resizes them to match and re-solves the radius. Typing in any field
and pressing Enter applies just that one value to the whole selection. Fields accept
arithmetic (`43+30`, `*2`, `(40+8)*2`). A leading operator applies to the field's current
value, so `-30` subtracts; for an absolute negative position type `0-30`.

### Fill and Space

**Fill horizontal / vertical** span the page margins. Horizontal sets widths and
x-positions only; vertical sets heights and y-positions only. They work with a single
shape, which simply stretches to the margins.

**Space horizontal / vertical** work inside the selection's own extent instead. The outer
edges stay put, and the shapes are resized so every gap equals the gutter for that axis.

All four re-solve the corner radius afterwards.

---

## Limits worth recording

- **Line spacing can't be set from a task pane.** `PowerPoint.ParagraphFormat` exposes
  only `bulletFormat`, `horizontalAlignment` and `indentLevel`. There is no workaround.
  Use PowerPoint's Format Painter (Cmd+Shift+C / Cmd+Shift+V) to carry line spacing
  between boxes, and set boxes to Multiple rather than Exactly so spacing scales with font
  size. Text-style copy and paste was deliberately removed because Format Painter already
  does it.
- **One unsupported property in a batched `load()` fails the whole batch silently.** Keep
  property lists conservative, and fall back to loading one property at a time.
- **Inputs bind by wrapper class.** The Enter handler binds to `.num input` rather than
  `input[type="number"]`, so it covers text-type fields too.
- **Standard 16:9 only.** The slide size is assumed, not read from the document.

---

## Files

```
mcpp_v1.xml      the bit PowerPoint reads: edit the nine URLs
taskpane.html    panel markup
taskpane.css     panel styling
taskpane.js      all logic: config, units, radius maths, align, grid, fit
assets/          add-in icons (16, 32, 64, 80 px)
```
