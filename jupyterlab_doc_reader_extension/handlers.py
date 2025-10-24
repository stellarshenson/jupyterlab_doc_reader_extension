import json
import os
import tempfile
import base64
from pathlib import Path
from io import BytesIO

from jupyter_server.base.handlers import APIHandler
from jupyter_server.utils import url_path_join
import tornado

try:
    from docx import Document
    from reportlab.lib.pagesizes import letter
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import inch
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, PageBreak, Table, TableStyle
    from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT, TA_JUSTIFY
    from reportlab.lib import colors
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    from html import escape as html_escape
except ImportError as e:
    Document = None
    _import_error = str(e)


class DocumentConverterHandler(APIHandler):
    """Handler for converting documents (DOCX, DOC, RTF) to PDF"""

    @tornado.web.authenticated
    def post(self):
        """
        Convert a document file to PDF and return as base64-encoded data.
        Expects JSON payload: {"path": "/path/to/document.docx"}
        """
        try:
            data = self.get_json_body()
            file_path = data.get('path')

            self.log.info(f"Converting document: {file_path}")
            self.log.debug(f"Request data: {data}")

            if not file_path:
                self.set_status(400)
                self.finish(json.dumps({"error": "No file path provided"}))
                return

            # Get the full path to the file
            contents_manager = self.settings.get('contents_manager')
            if contents_manager:
                root_dir = contents_manager.root_dir
            else:
                root_dir = os.getcwd()
                self.log.warning(f"Contents manager not found, using cwd: {root_dir}")

            full_path = os.path.join(root_dir, file_path.lstrip('/'))
            self.log.debug(f"Full path: {full_path}")

            if not os.path.exists(full_path):
                self.set_status(404)
                self.finish(json.dumps({"error": f"File not found: {file_path}"}))
                return

            # Check file extension
            ext = Path(full_path).suffix.lower()
            if ext not in ['.docx', '.doc', '.rtf']:
                self.set_status(400)
                self.finish(json.dumps({"error": f"Unsupported file type: {ext}"}))
                return

            # Convert to PDF
            try:
                pdf_data = self._convert_to_pdf(full_path)
                self.log.debug(f"PDF size: {len(pdf_data)} bytes")

                # Encode as base64 for transmission
                pdf_base64 = base64.b64encode(pdf_data).decode('utf-8')

                response_data = {
                    "success": True,
                    "pdf_data": pdf_base64,
                    "filename": Path(file_path).stem + ".pdf"
                }

                self.finish(json.dumps(response_data))
                self.log.info(f"Conversion successful")

            except Exception as convert_error:
                import traceback
                error_traceback = traceback.format_exc()
                self.log.error(f"Conversion error: {str(convert_error)}\n{error_traceback}")
                self.set_status(500)
                self.finish(json.dumps({
                    "success": False,
                    "error": f"Conversion failed: {str(convert_error)}",
                    "error_type": type(convert_error).__name__,
                    "traceback": error_traceback,
                    "file_path": file_path,
                    "full_path": full_path
                }))

        except Exception as e:
            import traceback
            error_traceback = traceback.format_exc()
            self.log.error(f"Handler error: {str(e)}\n{error_traceback}")
            self.set_status(500)
            self.finish(json.dumps({
                "success": False,
                "error": str(e),
                "error_type": type(e).__name__,
                "traceback": error_traceback
            }))

    def _convert_to_pdf(self, input_path: str) -> bytes:
        """
        Convert DOCX document to PDF using python-docx + reportlab.
        Pure Python solution with no external system dependencies.
        Returns PDF data as bytes.
        """
        if Document is None:
            raise Exception(
                f"Required libraries not installed: {_import_error}\n"
                "Please install: pip install python-docx reportlab"
            )

        ext = Path(input_path).suffix.lower()

        # Only support DOCX for now (DOC and RTF need additional libraries)
        if ext != '.docx':
            raise Exception(
                f"Only DOCX format is currently supported. "
                f"Please convert your {ext.upper()} file to DOCX format."
            )

        try:
            # Read the DOCX file
            doc = Document(input_path)
            self.log.debug(f"Document loaded. Paragraphs: {len(doc.paragraphs)}, Tables: {len(doc.tables)}")

            # Register Unicode-supporting fonts from system paths
            # Try common Linux font locations for fonts that support Polish characters
            font_candidates = [
                # DejaVu fonts (most common, excellent Unicode support)
                ('/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf', 'UnicodeSans'),
                ('/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf', 'UnicodeSansBold'),
                # Liberation fonts (alternative)
                ('/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf', 'UnicodeSans'),
                ('/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf', 'UnicodeSansBold'),
                # FreeSans (GNU FreeFont)
                ('/usr/share/fonts/truetype/freefont/FreeSans.ttf', 'UnicodeSans'),
                ('/usr/share/fonts/truetype/freefont/FreeSansBold.ttf', 'UnicodeSansBold'),
            ]

            registered_fonts = set()
            for font_path, font_name in font_candidates:
                if os.path.exists(font_path) and font_name not in registered_fonts:
                    try:
                        pdfmetrics.registerFont(TTFont(font_name, font_path))
                        registered_fonts.add(font_name)
                        self.log.info(f"Registered font {font_name} from {font_path}")
                    except Exception as font_error:
                        self.log.debug(f"Failed to register {font_name} from {font_path}: {font_error}")

            if not registered_fonts:
                self.log.warning("No Unicode fonts found. Polish characters may not display correctly.")

            # Create PDF in memory
            pdf_buffer = BytesIO()
            pdf_doc = SimpleDocTemplate(
                pdf_buffer,
                pagesize=letter,
                rightMargin=72,
                leftMargin=72,
                topMargin=72,
                bottomMargin=72
            )

            # Get default styles
            styles = getSampleStyleSheet()

            # Determine which font to use (prefer Unicode fonts if registered)
            font_name = 'UnicodeSans' if 'UnicodeSans' in pdfmetrics.getRegisteredFontNames() else 'Helvetica'
            font_name_bold = 'UnicodeSansBold' if 'UnicodeSansBold' in pdfmetrics.getRegisteredFontNames() else 'Helvetica-Bold'

            # Create custom styles with better formatting
            normal_style = ParagraphStyle(
                'CustomNormal',
                parent=styles['Normal'],
                fontName=font_name,
                fontSize=11,
                leading=14,
                spaceAfter=12
            )

            heading1_style = ParagraphStyle(
                'CustomHeading1',
                parent=styles['Heading1'],
                fontName=font_name_bold,
                fontSize=18,
                leading=22,
                spaceAfter=12,
                spaceBefore=12,
                textColor=colors.HexColor('#2E3440')
            )

            heading2_style = ParagraphStyle(
                'CustomHeading2',
                parent=styles['Heading2'],
                fontName=font_name_bold,
                fontSize=14,
                leading=18,
                spaceAfter=10,
                spaceBefore=10,
                textColor=colors.HexColor('#3B4252')
            )

            # Build the story (content)
            story = []
            paragraph_count = 0
            non_empty_paragraphs = 0

            for paragraph in doc.paragraphs:
                paragraph_count += 1
                if paragraph.text.strip():
                    non_empty_paragraphs += 1
                    # Detect heading styles
                    if paragraph.style.name.startswith('Heading 1'):
                        story.append(Paragraph(paragraph.text, heading1_style))
                    elif paragraph.style.name.startswith('Heading'):
                        story.append(Paragraph(paragraph.text, heading2_style))
                    else:
                        story.append(Paragraph(paragraph.text, normal_style))
                else:
                    story.append(Spacer(1, 0.2*inch))

            self.log.debug(f"Processed {paragraph_count} paragraphs ({non_empty_paragraphs} non-empty)")

            # Handle tables
            table_count = 0
            for table in doc.tables:
                table_data = []
                for row in table.rows:
                    row_data = [cell.text for cell in row.cells]
                    table_data.append(row_data)

                if table_data:
                    table_count += 1
                    t = Table(table_data)
                    t.setStyle(TableStyle([
                        ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                        ('FONTNAME', (0, 0), (-1, 0), font_name_bold),
                        ('FONTNAME', (0, 1), (-1, -1), font_name),
                        ('FONTSIZE', (0, 0), (-1, 0), 12),
                        ('FONTSIZE', (0, 1), (-1, -1), 10),
                        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                        ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
                        ('GRID', (0, 0), (-1, -1), 1, colors.black)
                    ]))
                    story.append(t)
                    story.append(Spacer(1, 0.2*inch))

            self.log.debug(f"Processed {table_count} tables, total story elements: {len(story)}")

            # Build the PDF
            if not story:
                self.log.warning("Story is empty! Creating a placeholder PDF")
                story.append(Paragraph("Document appears to be empty or contains no readable content.", normal_style))

            pdf_doc.build(story)

            # Get the PDF bytes
            pdf_bytes = pdf_buffer.getvalue()
            pdf_buffer.close()

            return pdf_bytes

        except Exception as e:
            error_msg = str(e)
            if ext == '.doc':
                raise Exception(
                    "Legacy DOC format not supported. "
                    "Please convert to DOCX format for best results. "
                    f"Error: {error_msg}"
                )
            else:
                raise Exception(f"PDF generation error: {error_msg}")


def setup_handlers(web_app):
    host_pattern = ".*$"

    base_url = web_app.settings["base_url"]

    # Document converter endpoint
    converter_pattern = url_path_join(base_url, "jupyterlab-doc-reader-extension", "convert")
    handlers = [(converter_pattern, DocumentConverterHandler)]

    web_app.add_handlers(host_pattern, handlers)
