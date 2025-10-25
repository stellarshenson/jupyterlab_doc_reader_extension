# Journal

1. **Task - JupyterLab DOCX Reader Extension**: Implement JupyterLab extension to open and view DOCX files in the browser<br>
    **Result**: Created full-featured extension that converts DOCX files to PDF on-the-fly. Used python-docx and reportlab for pure Python conversion (no system dependencies). Implemented server-side API handler for document conversion with base64 encoding for transmission. Frontend displays PDF using embed tag for better browser compatibility. Supports paragraphs, headings, and tables with proper formatting. Includes comprehensive error handling with detailed tracebacks. Added debug and info level logging on server side for monitoring.

2. **Task - Polish Character Support**: Fix rendering of Polish characters (ą, ć, ę, ł, ń, ó, ś, ź, ż) in generated PDFs<br>
    **Result**: Implemented robust Unicode font detection that searches common Linux system paths for TrueType fonts. Extension now tries DejaVu, Liberation, and FreeSans fonts from system locations. Registers found fonts as 'UnicodeSans' and 'UnicodeSansBold' with proper logging. Falls back to Helvetica if no Unicode fonts available. Polish characters now render correctly instead of displaying as box symbols.
