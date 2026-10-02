# Changelog

<!-- <START NEW CHANGELOG ENTRY> -->

## [1.2.1] - 2026-10-02

### Added

- ODT and ODP viewing with `@opendocument/odr-core` in a sandboxed frame, with zoom, fit and find
- XLSX viewing with `@silurus/ooxml`: sheet tabs, zoom and find across sheets
- PPTX slide thumbnails, previous and next buttons, keyboard navigation, zoom, fit and find
- Galata functional tests that open every supported format in a browser

### Changed

- DOCX, PPTX and RTF files are rendered in the browser by `docx-preview`, `@aiden0z/pptx-renderer` and `rtf.js` instead of being converted to PDF on the server
- The Python server extension is removed; the Python package ships the prebuilt labextension only
- Legacy `.doc` and `.ppt` files show a message asking to save the file as DOCX or PPTX
- Document links open only for `http`, `https` and `mailto`; `#` links scroll inside the document
- Requirements are JupyterLab >= 4.6.0 and Python >= 3.10
- Build and lint tooling upgraded: `@jupyter/builder` 1.2 (Rspack) replaces `@jupyterlab/builder`, ESLint 10 with a flat config file, typescript-eslint 8, stylelint 17, TypeScript 5.9 and Prettier 3.9

### Fixed

- The isolated install test in CI runs on Python 3.13 with JupyterLab 4.6, which the extension requires

<!-- <END NEW CHANGELOG ENTRY> -->
