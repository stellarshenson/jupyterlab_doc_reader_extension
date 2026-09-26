# Acceptance Criteria

Behaviour the document reader extension must show when it opens Word, PowerPoint and legacy office files in JupyterLab. DOCX, PPTX and RTF render in the browser from the file bytes, with no PDF conversion and no server extension.

## Authors

- `@kj` Konrad Jelen

## Word rendering `WORD`

DOCX files displayed in the browser from the file bytes, without LibreOffice or PDF conversion

- [x] `ACC-WORD-1` **Client-side DOCX render** - CRITICAL; a .docx opens as rendered pages inside the widget, drawn by `docx-preview` in the browser
  - evidence: galata 14/14 on the 1.2.0 wheel, 2026-09-26: 'renders pages in the browser with no server call and no PDF' finds 2 sections for the 2-page sample.docx
  - test-tags: E2E
  - test: open a sample .docx, assert the widget holds one `section` element per page
  - mechanism: 2026-09-26T15:39:07Z @kj widget awaits `context.ready`, decodes `context.model.toString()` base64 to a `Uint8Array`, calls docx-preview `renderAsync` into a host element inside the widget
  - log: 2026-09-26T14:36:07Z @kj added
  - log: 2026-09-26T15:18:50Z @kj edited test-tags added "E2E"
  - log: 2026-09-26T15:39:07Z @kj mechanism updated "2026-09-26T14:36:07Z @kj widget awaits `context.ready`, decodes `context.model.toString()` base64 to a Blob, calls `renderAsync` into `this.node`" -> "widget awaits `context.ready`, decodes `context.model.toString()` base64 to a `Uint8Array`, calls docx-preview `renderAsync` into a host element inside the widget"
  - log: 2026-09-26T15:39:07Z @kj edited test "open a sample .docx, assert the widget node holds `section.docx` page elements" -> "open a sample .docx, assert the widget holds one `section` element per page"
  - log: 2026-09-26T16:11:07Z @kj closed
- [x] `ACC-WORD-2` **No server call for DOCX** - HIGH; opening a .docx sends no request to `/jupyterlab-doc-reader-extension/convert`
  - evidence: galata 14/14 on the 1.2.0 wheel, 2026-09-26: the same DOCX test records no request to jupyterlab-doc-reader-extension routes
  - test-tags: E2E
  - test: open a .docx with network capture on, assert no POST to `convert`
  - log: 2026-09-26T14:36:07Z @kj added
  - log: 2026-09-26T15:18:50Z @kj edited test-tags added "E2E"
  - log: 2026-09-26T16:11:07Z @kj closed
- [x] `ACC-WORD-3` **No PDF step for DOCX** - CRITICAL; the DOCX path creates no PDF blob and no `<embed type="application/pdf">` element
  - evidence: galata 14/14 on the 1.2.0 wheel, 2026-09-26: the same DOCX test finds no embed or iframe element in the widget
  - test-tags: E2E
  - test: open a .docx, assert the widget node holds no `embed` element
  - log: 2026-09-26T14:36:07Z @kj added
  - log: 2026-09-26T15:18:50Z @kj edited test-tags added "E2E"
  - log: 2026-09-26T16:11:07Z @kj closed
- [x] `ACC-WORD-4` **No LibreOffice** - CRITICAL; no dependency, subprocess or install step calls LibreOffice or `soffice`
  - evidence: grep for soffice and libreoffice in package.json, pyproject.toml, src/, the Python package, .github/ and Makefile: no match, 2026-09-26
  - test-tags: MANUAL
  - test: grep `package.json`, `pyproject.toml`, `src/` and the Python package for `soffice` and `libreoffice`, expect no match
  - log: 2026-09-26T14:36:07Z @kj added
  - log: 2026-09-26T15:18:50Z @kj edited test-tags added "MANUAL"
  - log: 2026-09-26T16:11:07Z @kj closed
- [-] `ACC-WORD-5` **Pages scroll** - MEDIUM; a document longer than the panel scrolls vertically inside the widget
  - test: open a 10-page .docx, scroll to the last page
  - log: 2026-09-26T14:36:07Z @kj added
  - log: 2026-09-26T14:53:05Z @kj rejected: rejected: detail below functional level; page layout is part of ACC-WORD-1
- [-] `ACC-WORD-6` **Dark theme: pages readable** - MEDIUM; in the JupyterLab dark theme each page keeps the document text colour on a white page
  - test: switch to JupyterLab Dark, open a .docx with black text, assert the text is visible
  - log: 2026-09-26T14:36:07Z @kj added
  - log: 2026-09-26T14:53:05Z @kj rejected: rejected: detail below functional level, criteria kept functional at @kj request
- [-] `ACC-WORD-7` **Edge: invalid DOCX** - HIGH; a .docx that is not a valid DOCX zip shows the widget error panel with the library error message
  - test: rename a .txt file to .docx, open it
  - log: 2026-09-26T14:36:07Z @kj added
  - log: 2026-09-26T14:53:05Z @kj rejected: rejected: folded into ACC-VIEW-15, which covers every format
- [-] `ACC-WORD-8` **Edge: empty file** - MEDIUM; a 0-byte .docx shows the widget error panel, never a blank widget
  - test: create an empty `empty.docx`, open it
  - log: 2026-09-26T14:36:07Z @kj added
  - log: 2026-09-26T14:53:06Z @kj rejected: rejected: folded into ACC-VIEW-15, which covers empty files
- [-] `ACC-WORD-9` **DOCX troubleshooting text** - LOW; the error panel for a DOCX failure does not tell the user to check `python-docx` or `reportlab`
  - test: open an invalid .docx, read the troubleshooting list
  - log: 2026-09-26T14:36:07Z @kj added
  - log: 2026-09-26T14:53:06Z @kj rejected: rejected: the server troubleshooting list goes away with the server extension (ACC-PACK-23)
- [-] `ACC-WORD-17` **Headings in Table of Contents** - MEDIUM; DOCX headings appear in the JupyterLab Table of Contents panel; clicking one scrolls the document to it
  - test: open a .docx with two headings, open the Table of Contents panel, click the second heading
  - mechanism: 2026-09-26T14:52:57Z @kj TOC factory registered with `ITableOfContentsRegistry`; headings read from paragraphs whose style name is `heading N` or `Title`, so localized style ids still match
  - log: 2026-09-26T14:52:57Z @kj added
  - log: 2026-09-26T14:59:32Z @kj rejected: docx-preview offers no table of contents; the extension adds no feature the viewer lacks

## Presentations `SLIDE`

PPTX slides rendered in the browser by `@aiden0z/pptx-renderer`, with the controls that viewer offers

- [-] `ACC-SLIDE-10` **PPTX keeps server PDF path** - HIGH; a .pptx still converts on the server with `python-pptx` and `reportlab` and shows as PDF in an `<embed>` element
  - test: open a sample .pptx, assert one POST to `convert` and one `embed` element in the widget
  - test-tags: UNIT
  - log: 2026-09-26T14:36:07Z @kj added
  - log: 2026-09-26T14:53:06Z @kj rejected: rejected: superseded, PPTX renders in the browser (ACC-SLIDE-18)
- [x] `ACC-SLIDE-18` **Client-side PPTX render** - CRITICAL; a .pptx opens as slides rendered in the browser by `@aiden0z/pptx-renderer`, with no PDF step
  - evidence: galata 14/14 on the 1.2.0 wheel, 2026-09-26: 'renders the first slide in the browser' finds Alpha slide text and no embed or iframe
  - test-tags: E2E
  - test: open a sample .pptx, assert the first slide text is in the widget DOM and no `embed` element exists
  - mechanism: 2026-09-26T14:52:57Z @kj `PptxViewer` in slide mode with `RECOMMENDED_ZIP_LIMITS`, loaded with `import()` only when a .pptx opens
  - log: 2026-09-26T14:52:57Z @kj added
  - log: 2026-09-26T15:18:51Z @kj edited test-tags added "E2E"
  - log: 2026-09-26T16:11:08Z @kj closed
- [x] `ACC-SLIDE-19` **Slide navigation** - HIGH; previous and next buttons and the arrow, PageUp, PageDown, Home and End keys change the shown slide
  - evidence: galata 14/14 on the 1.2.0 wheel, 2026-09-26: 'moves between slides with the buttons and keys' checks the counter after buttons, arrows, PageUp, PageDown, Home and End
  - test-tags: E2E
  - test: open a 3-slide .pptx, press next, ArrowRight, End, Home, assert the counter after each
  - log: 2026-09-26T14:52:58Z @kj added
  - log: 2026-09-26T15:18:51Z @kj edited test-tags added "E2E"
  - log: 2026-09-26T16:11:08Z @kj closed
- [x] `ACC-SLIDE-20` **Slide thumbnails** - MEDIUM; a thumbnail strip lists every slide and marks the shown one; clicking a thumbnail shows that slide
  - evidence: galata 14/14 on the 1.2.0 wheel, 2026-09-26: 'lists every slide as a thumbnail that opens it' finds 3 thumbnails; clicking the third shows 3 / 3
  - test-tags: E2E
  - test: open a 3-slide .pptx, assert 3 thumbnails, click the third, assert counter `3 / 3`
  - log: 2026-09-26T14:52:58Z @kj added
  - log: 2026-09-26T15:18:51Z @kj edited test-tags added "E2E"
  - log: 2026-09-26T16:11:08Z @kj closed
- [x] `ACC-SLIDE-21` **Find in slides** - HIGH; the toolbar find box searches every slide; Enter shows the slide with the next match and highlights the matched shape
  - evidence: galata 14/14 on the 1.2.0 wheel, 2026-09-26: 'finds text on any slide and highlights it': zebra shows 3 / 3, 1 of 1 and one highlight
  - test-tags: E2E
  - test: open a 3-slide .pptx, type a word found only on slide 3, press Enter, assert counter `3 / 3` and one highlight
  - mechanism: 2026-09-26T14:52:58Z @kj renderer `searchText` over the parsed model, then `goToSlide` and `highlightSearchResult`
  - log: 2026-09-26T14:52:58Z @kj added
  - log: 2026-09-26T15:18:51Z @kj edited test-tags added "E2E"
  - log: 2026-09-26T16:11:08Z @kj closed

## Legacy formats `LEGACY`

DOC and PPT binary files, which no browser renderer reads

- [x] `ACC-LEGACY-11` **DOC and PPT show unsupported message** - MEDIUM; opening a .doc or .ppt shows a panel saying the format is not supported and to save the file as DOCX or PPTX
  - evidence: galata 14/14 on the 1.2.0 wheel, 2026-09-26: 'DOC shows the unsupported message' and 'PPT shows the unsupported message' pass
  - test: open a sample .doc and .ppt, read the panel
  - test-tags: E2E
  - log: 2026-09-26T14:36:07Z @kj added
  - log: 2026-09-26T14:53:06Z @kj amended title "DOC shows unsupported message" -> "DOC and PPT show unsupported message"; text "opening a .doc shows the error panel with `Legacy .DOC format not supported`" -> "opening a .doc or .ppt shows a panel saying the format is not supported and to save the file as DOCX or PPTX"
  - log: 2026-09-26T14:53:06Z @kj edited test "open a sample .doc, read the error panel" -> "open a sample .doc and .ppt, read the panel"
  - log: 2026-09-26T15:18:51Z @kj edited test-tags "UNIT" -> "E2E"
  - log: 2026-09-26T16:11:08Z @kj closed
- [-] `ACC-LEGACY-12` **RTF shows unsupported message** - MEDIUM; opening a .rtf shows the error panel with `Legacy .RTF format not supported`
  - test: open a sample .rtf, read the error panel
  - test-tags: UNIT
  - log: 2026-09-26T14:36:07Z @kj added
  - log: 2026-09-26T14:53:06Z @kj rejected: rejected: superseded, RTF renders in the browser (ACC-RTF-22)

## Viewer `VIEW`

Toolbar, errors and link handling shared by every format

- [x] `ACC-VIEW-13` **Slide zoom and fit** - HIGH; for PPTX, toolbar zoom in and zoom out call the viewer `setZoom` in 25 percent steps and fit calls `setZoom(100)`, the width of the panel; DOCX and RTF get no zoom, their viewers offer none
  - evidence: galata 14/14 on the 1.2.0 wheel, 2026-09-26: 'zooms the slide and fits it back to the panel': width grows, shrinks and returns to the fitted width
  - test-tags: E2E
  - test: open a .pptx, press zoom in, zoom out and fit, assert the slide width changes each time
  - log: 2026-09-26T14:52:57Z @kj added
  - log: 2026-09-26T14:59:32Z @kj amended title "Zoom and fit" -> "Slide zoom and fit"; text "toolbar zoom in, zoom out and fit change the scale of the rendered pages or slides" -> "for PPTX, toolbar zoom in, zoom out and fit call the viewer `setZoom` and `setFitMode`; DOCX and RTF get no zoom, their viewers offer none"
  - log: 2026-09-26T14:59:32Z @kj edited test "open a .docx and a .pptx, press zoom in, zoom out and fit, assert the page or slide width changes each time" -> "open a .pptx, press zoom in, zoom out and fit, assert the slide width changes each time"
  - log: 2026-09-26T15:18:51Z @kj edited test-tags added "E2E"
  - log: 2026-09-26T15:39:07Z @kj amended text "for PPTX, toolbar zoom in, zoom out and fit call the viewer `setZoom` and `setFitMode`; DOCX and RTF get no zoom, their viewers offer none" -> "for PPTX, toolbar zoom in and zoom out call the viewer `setZoom` in 25 percent steps and fit calls `setZoom(100)`, the width of the panel; DOCX and RTF get no zoom, their viewers offer none"
  - log: 2026-09-26T16:11:08Z @kj closed
- [x] `ACC-VIEW-14` **Slide counter** - MEDIUM; toolbar shows the shown PPTX slide as `n / N` from the viewer `currentSlideIndex` and `slideCount`
  - evidence: galata 14/14 on the 1.2.0 wheel, 2026-09-26: the slide navigation test asserts n / N after every move, including 3 / 3 after End
  - test-tags: E2E
  - test: open a 3-slide .pptx, press End, assert `3 / 3`
  - log: 2026-09-26T14:52:57Z @kj added
  - log: 2026-09-26T14:59:32Z @kj amended title "Position counter" -> "Slide counter"; text "toolbar shows the current DOCX page or PPTX slide as `n / N` and updates it on scroll or slide change" -> "toolbar shows the shown PPTX slide as `n / N` from the viewer `currentSlideIndex` and `slideCount`"
  - log: 2026-09-26T14:59:32Z @kj edited test "open a 3-page .docx, scroll to the end, assert `3 / 3`" -> "open a 3-slide .pptx, press End, assert `3 / 3`"
  - log: 2026-09-26T15:18:51Z @kj edited test-tags added "E2E"
  - log: 2026-09-26T16:11:08Z @kj closed
- [x] `ACC-VIEW-15` **Unreadable file shows error** - HIGH; a file the renderer cannot read, an empty file included, shows the error panel with the renderer message, never a blank widget
  - evidence: galata 14/14 on the 1.2.0 wheel, 2026-09-26: 'shows the renderer error for a file that is not a DOCX' and 'shows an error for an empty file' pass
  - test-tags: E2E
  - test: open a .txt renamed to .docx and a 0-byte .pptx, assert the error panel each time
  - log: 2026-09-26T14:52:57Z @kj added
  - log: 2026-09-26T15:18:51Z @kj edited test-tags added "E2E"
  - log: 2026-09-26T16:11:08Z @kj closed
- [x] `ACC-VIEW-16` **Safe links** - HIGH; document links open only `http`, `https` and `mailto` targets, in a new tab; `#` links scroll inside the document; any other target is removed
  - evidence: jest 11/11 guardLinks tests and galata 14/14 on the 1.2.0 wheel, 2026-09-26: 'opens external links in a new tab and removes unsafe ones' pass
  - test-tags: UNIT, E2E
  - test: open a .docx holding a `javascript:` link and an external link, assert the first has no href and the second has `target="_blank"`
  - log: 2026-09-26T14:52:57Z @kj added
  - log: 2026-09-26T15:18:51Z @kj edited test-tags added "UNIT, E2E"
  - log: 2026-09-26T16:11:08Z @kj closed
- [x] `ACC-VIEW-24` **Features come from the viewer** - HIGH; every control the extension shows calls a feature of the chosen viewer library; the extension adds no table of contents, search or zoom of its own
  - evidence: review of src/, 2026-09-26: no heading scan, text search or CSS zoom code; each toolbar control calls a SlideDeck method that calls the pptx-renderer API
  - test-tags: MANUAL
  - test: review `src/`: each toolbar control maps to a viewer API call; no DOM search, heading scan or CSS zoom code
  - log: 2026-09-26T14:59:32Z @kj added
  - log: 2026-09-26T15:18:51Z @kj edited test-tags added "MANUAL"
  - log: 2026-09-26T16:11:08Z @kj closed
- [x] `ACC-VIEW-25` **Embedded HTML does not run** - HIGH; HTML a DOCX embeds as an altChunk part is not rendered, so its script never runs in the JupyterLab page
  - evidence: galata 14/14 on the 1.2.0 wheel, 2026-09-26: 'does not render HTML the file embeds as an altChunk' passes; the same test fails with renderAltChunks: true
  - related: DEF-DOCX-2 - the defect this covers
  - test: open altchunk.docx; assert the text around the chunk shows and no iframe exists
  - test-tags: E2E
  - mechanism: 2026-09-26T16:03:48Z @kj renderDocx passes renderAltChunks: false to docx-preview renderAsync
  - log: 2026-09-26T16:03:48Z @kj added
  - log: 2026-09-26T16:11:08Z @kj closed

## Rich Text `RTF`

RTF files rendered in the browser

- [x] `ACC-RTF-22` **Client-side RTF render** - HIGH; a .rtf opens as formatted text rendered in the browser by `rtf.js`
  - evidence: galata 14/14 on the 1.2.0 wheel, 2026-09-26: 'renders formatted text in the browser' finds the text and a bold run with font weight 700 or more
  - test-tags: E2E
  - test: open a sample .rtf with bold text, assert the text and a bold run are in the widget DOM
  - log: 2026-09-26T14:52:58Z @kj added
  - log: 2026-09-26T15:18:51Z @kj edited test-tags added "E2E"
  - log: 2026-09-26T16:11:08Z @kj closed

## Packaging `PACK`

What the Python and npm packages ship

- [x] `ACC-PACK-23` **No server extension** - HIGH; the package ships no Jupyter server extension and no route under `/jupyterlab-doc-reader-extension/`
  - evidence: wheel 1.2.0 installed 2026-09-26: jupyter server extension list does not name the package; the labextension is enabled and OK
  - test-tags: INTEGRATION
  - test: install the wheel, assert `jupyter server extension list` does not name the package
  - log: 2026-09-26T14:52:58Z @kj added
  - log: 2026-09-26T15:18:51Z @kj edited test-tags added "INTEGRATION"
  - log: 2026-09-26T16:11:08Z @kj closed

