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
- [x] `DEF-WIDGET-9-TOOLBAR` **Toolbar buttons look and stay pressed** - MINOR; toolbar buttons in every viewer look like toggles and keep a pressed frame after a click; momentary actions expected; cause: ToolbarButton toggles its own pressed flag on click; fix: reset pressed after the action; `src/widget.ts`
  - evidence: Galata 'toolbar buttons are momentary' passes on build 1.2.4; 24/24 Galata and 14/14 Jest green; no aria-pressed true after Zoom in, Zoom out or Fit
  - root-cause: 2026-10-28T23:11:47Z @kj ToolbarButton in @jupyterlab/apputils 4.6.4 toggles its own pressed flag on every click; the shared helper at src/widget.ts:284 uses it for every viewer control, so each click leaves a pressed frame; reset pressed after the action
  - repro: open a DOCX, PPTX, XLSX, ODT or ODP file, look at -, +, Fit in the toolbar after a click
  - test-tags: FUNCTIONAL
  - log: 2026-10-28T23:10:23Z @kj added
  - log: 2026-10-28T23:24:34Z @kj edited text "toolbar buttons in every viewer look like toggles and keep a pressed frame after a click; momentary actions expected; cause under investigation; `src/widget.ts`" -> "toolbar buttons in every viewer look like toggles and keep a pressed frame after a click; momentary actions expected; cause: ToolbarButton toggles its own pressed flag on click; fix: reset pressed after the action; `src/widget.ts`"
  - log: 2026-10-28T23:24:34Z @kj closed: fixed: pressed reset in the shared addViewerToolbar helper
- [x] `DEF-WIDGET-10-DEAD-CODE` **Unused model and file type fields** - MINOR; src/model.ts (DocModel, 102 lines) has no importer and its comment says content is not loaded, while the factory uses the base64 model; contentType and fileFormat in FILE_TYPES are never passed to addFileType; both ship in the npm src and mislead a maintainer
  - evidence: src/model.ts removed, contentType and fileFormat removed from FILE_TYPES; build_test 2026-11-13: tsc clean, Jest 14, Galata 24 green
  - repro: grep -rn 'DocModel\|fileFormat\|contentType' src
  - test-tags: UNIT
  - root-cause: 2026-11-13T06:58:26Z @kj left from the server-PDF design; the widget moved to the built-in base64 model and the leftovers stayed
  - log: 2026-11-13T06:58:26Z @kj added
  - log: 2026-11-13T07:02:05Z @kj closed

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
- [x] `DEF-DOCS-11-USAGE` **README usage omits three formats** - MINOR; README Usage says to open a .docx, .pptx or .rtf file; .xlsx, .odt and .odp open too, so a reader of that section does not learn it
  - evidence: README Usage lists .docx, .pptx, .xlsx, .odt, .odp and .rtf
  - repro: read README.md section Usage
  - test-tags: MANUAL
  - root-cause: 2026-11-13T06:58:26Z @kj Usage was not updated when ODF and XLSX viewers were added
  - log: 2026-11-13T06:58:26Z @kj added
  - log: 2026-11-13T07:02:05Z @kj closed
- [x] `DEF-DOCS-12-STALE-TAB` **README omits that an open tab is not refreshed** - MINOR; an open viewer tab keeps showing the old document after the file is rewritten on disk, and Reload from Disk does not render it again; README Caveats did not say so
  - evidence: README Caveats states that an open tab shows the file as it was when opened and that Reload from Disk does not render it again
  - repro: open a .docx, rewrite the file on disk, choose File, Reload from Disk: the tab shows the old content
  - test-tags: MANUAL
  - root-cause: 2026-11-13T07:24:43Z @kj the widget calls _load once in its constructor (src/widget.ts:73) and subscribes to no content or file change signal; a re-render is not built, so the limit is documented
  - log: 2026-11-13T07:24:43Z @kj added
  - log: 2026-11-13T07:24:49Z @kj closed

