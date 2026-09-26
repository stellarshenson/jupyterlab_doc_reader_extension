# Defects

Observed wrong behaviour in the document reader extension, with the trail of what was tried against it.

## Authors

- `@kj` Konrad Jelen

## Viewer widget `WIDGET`

Widget lifecycle, display elements and cleanup in `src/widget.ts`

- [-] `DEF-WIDGET-1` **PDF blob URL never revoked** - MINOR; the PDF object URL stays allocated after the document tab closes; fix: keep the embed URL and revoke it in `dispose()`; `src/widget.ts`
  - related: ACC-SLIDE-10 - the PPTX path keeps this leak after the DOCX change
  - repro: open a .pptx, copy the `embed` blob URL in devtools, close the tab, open the URL; the PDF still loads
  - root-cause: 2026-09-26T14:36:12Z @kj `src/widget.ts:109` creates the URL for the `<embed>` only; `dispose()` at line 272 checks `_iframe.src`, which never holds it
  - log: 2026-09-26T14:36:12Z @kj added
  - log: 2026-09-26T15:18:51Z @kj rejected: the PDF embed path was removed in 1.2.0; no blob URL is created any more
- [x] `DEF-WIDGET-7` **Disabled links give no reason** - MINOR; a DOCX or RTF link the guard disables still looks like a link and does nothing on click, with no explanation; `src/render.ts`
  - evidence: guardLinks sets title 'Link disabled: only web, mail and in-document links open'; jest 11/11 and galata 14/14 assert it
  - test-tags: UNIT, E2E
  - related: ACC-VIEW-16
  - repro: open a .docx holding a relative or file link, click it
  - root-cause: 2026-09-26T15:39:42Z @kj `guardLinks` removes `href` and adds nothing the user can read
  - log: 2026-09-26T15:39:42Z @kj added
  - log: 2026-09-26T16:10:58Z @kj edited test-tags added "UNIT, E2E"
  - log: 2026-09-26T16:10:58Z @kj closed

## DOCX rendering `DOCX`

Pages drawn by docx-preview inside the widget

- [x] `DEF-DOCX-2` **altChunk HTML runs script in JupyterLab** - CRITICAL; a DOCX with an altChunk HTML part runs that HTML, scripts included, in an iframe with no sandbox in the JupyterLab origin; fix: `renderAltChunks: false`; `src/render.ts`
  - evidence: render.ts sets renderAltChunks: false; galata altChunk test passes in 14/14 and fails with the option on; review rounds 2-4 found no failure
  - test-tags: E2E
  - related: ACC-VIEW-16
  - repro: open a DOCX whose altChunk HTML part holds a script, watch it run
  - root-cause: 2026-09-26T15:39:42Z @kj docx-preview 0.4.1 defaults `renderAltChunks: true`; `renderAltChunk` builds an unsandboxed iframe and sets `srcdoc` to the part
  - log: 2026-09-26T15:39:42Z @kj added
  - log: 2026-09-26T16:10:57Z @kj edited test-tags added "E2E"
  - log: 2026-09-26T16:10:57Z @kj closed
- [x] `DEF-DOCX-3` **Narrow panel clips the page** - MAJOR; in a panel narrower than the page, the left part of the page is cut off and cannot be scrolled to; fix: `align-items: safe center` on the wrapper; `style/base.css`
  - evidence: base.css align-items: safe center; review round 2 checked a page wider than the panel scrolls from its left edge; rounds 3-4 clean
  - test-tags: MANUAL
  - related: ACC-WORD-1
  - repro: open a .docx, drag the panel narrower than the page, scroll left
  - root-cause: 2026-09-26T15:39:42Z @kj docx-preview wrapper is a flex column with `align-items: center`; a centred item wider than its container overflows to the left, where scrolling cannot reach
  - log: 2026-09-26T15:39:42Z @kj added
  - log: 2026-09-26T16:10:57Z @kj edited test-tags added "MANUAL"
  - log: 2026-09-26T16:10:57Z @kj closed

## PPTX viewer `SLIDE`

Slides, thumbnails and toolbar driven by @aiden0z/pptx-renderer

- [x] `DEF-SLIDE-4` **Toolbar click stops slide keys** - MEDIUM; after clicking a slide toolbar button, the arrow, PageUp, PageDown, Home and End keys do nothing until the slide is clicked; fix: `noFocusOnClick` on the buttons; `src/widget.ts`
  - evidence: noFocusOnClick: true on slide toolbar buttons; one-off galata run 2026-09-26: ArrowRight and Home move slides right after Next slide and Zoom in clicks
  - test-tags: MANUAL
  - related: ACC-SLIDE-19
  - repro: open a .pptx, click Next slide, press ArrowRight
  - root-cause: 2026-09-26T15:39:42Z @kj ToolbarButton focuses itself on click; the key listener sits on the content node, a sibling of the toolbar
  - log: 2026-09-26T15:39:42Z @kj added
  - log: 2026-09-26T16:10:57Z @kj edited test-tags added "MANUAL"
  - log: 2026-09-26T16:10:57Z @kj closed
- [x] `DEF-SLIDE-5` **Deck leaks when tab closes during load** - MINOR; closing a PPTX tab before its first slide shows keeps the parsed deck and its media in memory until the page reloads; `src/widget.ts`
  - evidence: widget.ts _load disposes the deck and returns when the widget was disposed during SlideDeck.open; review round 4 found no failure
  - test-tags: MANUAL
  - repro: open a large .pptx, close the tab before the first slide appears
  - root-cause: 2026-09-26T15:39:42Z @kj `_load` assigns the deck after `await SlideDeck.open` with no disposed check; `dispose()` ran while the deck was still null
  - log: 2026-09-26T15:39:42Z @kj added
  - log: 2026-09-26T16:10:57Z @kj edited test-tags added "MANUAL"
  - log: 2026-09-26T16:10:57Z @kj closed
- [x] `DEF-SLIDE-6` **Slide changes not announced** - MEDIUM; screen readers are not told the find result, the slide position or which thumbnail is shown; `src/widget.ts`, `src/slides.ts`
  - evidence: role=status on slide position and find status, aria-current on thumbnails; galata asserts all three, 14/14 pass
  - test-tags: E2E
  - related: ACC-SLIDE-21
  - repro: open a .pptx with a screen reader, press Next slide, then find a word
  - root-cause: 2026-09-26T15:39:42Z @kj the position and find status spans have no live region and the active thumbnail is marked only by a CSS class
  - log: 2026-09-26T15:39:42Z @kj added
  - log: 2026-09-26T16:10:57Z @kj edited test-tags added "E2E"
  - log: 2026-09-26T16:10:57Z @kj closed

## Descriptions `DOCS`

Text that describes the extension: plugin metadata, README, comments, trackers

- [x] `DEF-DOCS-8` **Descriptions state removed behaviour** - MINOR; plugin description named DOC and omitted PPTX, README claimed Word page-break markers are used, the tracker said PPTX converts to PDF, a comment said PPTX links pass the link guard; fixed in text
  - evidence: plugin description, README, render.ts header and tracker text corrected in review rounds 1-2; rounds 3-4 found no false statement
  - test-tags: MANUAL
  - repro: read `src/index.ts` plugin description and README Caveats
  - log: 2026-09-26T15:39:42Z @kj added
  - log: 2026-09-26T16:10:58Z @kj edited test-tags added "MANUAL"
  - log: 2026-09-26T16:10:58Z @kj closed

