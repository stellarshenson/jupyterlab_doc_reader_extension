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
    from reportlab.lib.units import inch, Emu
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, PageBreak, Table, TableStyle, Image as RLImage
    from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT, TA_JUSTIFY
    from reportlab.lib import colors
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    from reportlab.pdfgen import canvas
    from html import escape as html_escape
    DOCX_AVAILABLE = True
except ImportError as e:
    Document = None
    DOCX_AVAILABLE = False
    _docx_import_error = str(e)

try:
    from pptx import Presentation
    from pptx.util import Inches, Pt, Emu as PptxEmu
    from pptx.enum.shapes import MSO_SHAPE_TYPE
    from pptx.dml.color import RGBColor
    PPTX_AVAILABLE = True
except ImportError as e:
    Presentation = None
    PPTX_AVAILABLE = False
    _pptx_import_error = str(e)

try:
    from PIL import Image as PILImage
    PIL_AVAILABLE = True
except ImportError:
    PIL_AVAILABLE = False


class DocumentConverterHandler(APIHandler):
    """Handler for converting documents (DOCX, DOC, RTF, PPTX, PPT) to PDF"""

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
            if ext not in ['.docx', '.doc', '.rtf', '.pptx', '.ppt']:
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
        Convert document to PDF using python-docx/python-pptx + reportlab.
        Pure Python solution with no external system dependencies.
        Returns PDF data as bytes.
        """
        ext = Path(input_path).suffix.lower()

        # Route to appropriate converter based on file type
        if ext == '.pptx':
            return self._convert_pptx_to_pdf(input_path)
        elif ext == '.ppt':
            raise Exception(
                "Legacy PPT format not supported. "
                "Please convert to PPTX format for best results."
            )
        elif ext == '.docx':
            return self._convert_docx_to_pdf(input_path)
        elif ext in ['.doc', '.rtf']:
            raise Exception(
                f"Legacy {ext.upper()} format not supported. "
                f"Please convert to DOCX format for best results."
            )
        else:
            raise Exception(f"Unsupported file type: {ext}")

    def _convert_docx_to_pdf(self, input_path: str) -> bytes:
        """
        Convert DOCX document to PDF using python-docx + reportlab.
        """
        if not DOCX_AVAILABLE:
            raise Exception(
                f"Required libraries not installed: {_docx_import_error}\n"
                "Please install: pip install python-docx reportlab"
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
            raise Exception(f"DOCX to PDF conversion error: {str(e)}")

    def _convert_pptx_to_pdf(self, input_path: str) -> bytes:
        """
        Convert PPTX presentation to PDF using python-pptx + reportlab.
        Renders each slide as a PDF page with text, shapes, and images.
        """
        if not PPTX_AVAILABLE:
            raise Exception(
                f"Required libraries not installed: {_pptx_import_error}\n"
                "Please install: pip install python-pptx"
            )

        try:
            # Load the presentation
            prs = Presentation(input_path)
            self.log.debug(f"Presentation loaded. Slides: {len(prs.slides)}")

            # Get slide dimensions (default is 10x7.5 inches for 4:3)
            slide_width_emu = prs.slide_width
            slide_height_emu = prs.slide_height

            # Convert EMU to points (1 inch = 914400 EMU, 1 inch = 72 points)
            slide_width_pt = slide_width_emu / 914400 * 72
            slide_height_pt = slide_height_emu / 914400 * 72

            self.log.debug(f"Slide size: {slide_width_pt:.1f} x {slide_height_pt:.1f} points")

            # Register Unicode fonts (same as DOCX handler)
            self._register_unicode_fonts()

            # Determine which font to use
            font_name = 'UnicodeSans' if 'UnicodeSans' in pdfmetrics.getRegisteredFontNames() else 'Helvetica'
            font_name_bold = 'UnicodeSansBold' if 'UnicodeSansBold' in pdfmetrics.getRegisteredFontNames() else 'Helvetica-Bold'

            # Create PDF in memory using canvas for precise positioning
            pdf_buffer = BytesIO()
            c = canvas.Canvas(pdf_buffer, pagesize=(slide_width_pt, slide_height_pt))

            for slide_idx, slide in enumerate(prs.slides):
                self.log.debug(f"Processing slide {slide_idx + 1}")

                # Draw slide background (white by default)
                c.setFillColor(colors.white)
                c.rect(0, 0, slide_width_pt, slide_height_pt, fill=1, stroke=0)

                # Try to get slide background color
                try:
                    if slide.background.fill.solid():
                        bg_color = slide.background.fill.fore_color.rgb
                        if bg_color:
                            c.setFillColor(colors.HexColor(f'#{bg_color}'))
                            c.rect(0, 0, slide_width_pt, slide_height_pt, fill=1, stroke=0)
                except Exception:
                    pass  # Use default white background

                # Process shapes on the slide
                for shape in slide.shapes:
                    try:
                        self._render_shape_to_canvas(
                            c, shape, slide_width_emu, slide_height_emu,
                            slide_width_pt, slide_height_pt,
                            font_name, font_name_bold
                        )
                    except Exception as shape_error:
                        self.log.debug(f"Error rendering shape: {shape_error}")
                        continue

                # Add slide number at bottom
                c.setFont(font_name, 8)
                c.setFillColor(colors.grey)
                c.drawCentredString(slide_width_pt / 2, 15, f"Slide {slide_idx + 1}")

                # Move to next page
                c.showPage()

            # Save the PDF
            c.save()
            pdf_bytes = pdf_buffer.getvalue()
            pdf_buffer.close()

            self.log.info(f"PPTX conversion complete: {len(prs.slides)} slides")
            return pdf_bytes

        except Exception as e:
            raise Exception(f"PPTX to PDF conversion error: {str(e)}")

    def _register_unicode_fonts(self):
        """Register Unicode-supporting fonts from system paths."""
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
            self.log.warning("No Unicode fonts found. International characters may not display correctly.")

    def _render_shape_to_canvas(self, c, shape, slide_width_emu, slide_height_emu,
                                 slide_width_pt, slide_height_pt, font_name, font_name_bold):
        """Render a single shape to the PDF canvas."""
        # Get shape position and size in points
        # Note: PDF coordinates start from bottom-left, PPTX from top-left
        left_pt = shape.left / 914400 * 72
        top_pt = shape.top / 914400 * 72
        width_pt = shape.width / 914400 * 72
        height_pt = shape.height / 914400 * 72

        # Convert to PDF coordinates (flip Y axis)
        pdf_y = slide_height_pt - top_pt - height_pt

        # Handle different shape types
        if shape.has_text_frame:
            self._render_text_frame(c, shape.text_frame, left_pt, pdf_y, width_pt, height_pt,
                                    slide_height_pt, font_name, font_name_bold)

        # Handle images
        if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
            self._render_picture(c, shape, left_pt, pdf_y, width_pt, height_pt)

        # Handle tables
        if shape.has_table:
            self._render_table(c, shape.table, left_pt, pdf_y, width_pt, height_pt,
                              font_name, font_name_bold)

    def _render_text_frame(self, c, text_frame, x, y, width, height, slide_height_pt,
                           font_name, font_name_bold):
        """Render a text frame to the canvas."""
        current_y = y + height - 5  # Start from top of text box

        for paragraph in text_frame.paragraphs:
            text = paragraph.text.strip()
            if not text:
                current_y -= 12  # Empty paragraph spacing
                continue

            # Determine font size and style
            font_size = 12  # Default
            current_font = font_name

            try:
                if paragraph.runs:
                    run = paragraph.runs[0]
                    if run.font.size:
                        font_size = run.font.size.pt
                    if run.font.bold:
                        current_font = font_name_bold
            except Exception:
                pass

            # Clamp font size to reasonable range
            font_size = max(6, min(72, font_size))

            # Set font color
            try:
                if paragraph.runs and paragraph.runs[0].font.color.rgb:
                    rgb = paragraph.runs[0].font.color.rgb
                    c.setFillColor(colors.HexColor(f'#{rgb}'))
                else:
                    c.setFillColor(colors.black)
            except Exception:
                c.setFillColor(colors.black)

            c.setFont(current_font, font_size)

            # Handle text alignment
            try:
                from pptx.enum.text import PP_ALIGN
                if paragraph.alignment == PP_ALIGN.CENTER:
                    c.drawCentredString(x + width / 2, current_y, text)
                elif paragraph.alignment == PP_ALIGN.RIGHT:
                    c.drawRightString(x + width, current_y, text)
                else:
                    c.drawString(x + 5, current_y, text)
            except Exception:
                c.drawString(x + 5, current_y, text)

            # Move to next line
            current_y -= font_size * 1.2

    def _render_picture(self, c, shape, x, y, width, height):
        """Render a picture shape to the canvas."""
        if not PIL_AVAILABLE:
            self.log.debug("PIL not available, skipping image")
            return

        try:
            # Get image blob from shape
            image = shape.image
            image_bytes = image.blob

            # Create PIL image and save to buffer
            pil_image = PILImage.open(BytesIO(image_bytes))

            # Convert to RGB if necessary (for PDF compatibility)
            if pil_image.mode in ('RGBA', 'P'):
                pil_image = pil_image.convert('RGB')

            # Save to buffer for reportlab
            img_buffer = BytesIO()
            pil_image.save(img_buffer, format='PNG')
            img_buffer.seek(0)

            # Draw image on canvas using ImageReader
            from reportlab.lib.utils import ImageReader
            img_reader = ImageReader(img_buffer)
            c.drawImage(img_reader, x, y, width=width, height=height,
                        preserveAspectRatio=True, mask='auto')
        except Exception as e:
            self.log.debug(f"Error rendering image: {e}")
            # Draw placeholder rectangle
            c.setStrokeColor(colors.grey)
            c.setFillColor(colors.lightgrey)
            c.rect(x, y, width, height, fill=1, stroke=1)
            c.setFillColor(colors.grey)
            c.setFont('Helvetica', 8)
            c.drawCentredString(x + width/2, y + height/2, "[Image]")

    def _render_table(self, c, table, x, y, width, height, font_name, font_name_bold):
        """Render a table to the canvas."""
        if not table.rows:
            return

        num_rows = len(table.rows)
        num_cols = len(table.columns)

        cell_width = width / num_cols
        cell_height = height / num_rows

        current_y = y + height  # Start from top

        for row_idx, row in enumerate(table.rows):
            current_x = x
            current_y -= cell_height

            for col_idx, cell in enumerate(row.cells):
                # Draw cell border
                c.setStrokeColor(colors.black)
                c.rect(current_x, current_y, cell_width, cell_height, fill=0, stroke=1)

                # Draw cell text
                text = cell.text.strip()
                if text:
                    # Use bold for header row
                    if row_idx == 0:
                        c.setFont(font_name_bold, 9)
                        c.setFillColor(colors.HexColor('#333333'))
                    else:
                        c.setFont(font_name, 9)
                        c.setFillColor(colors.black)

                    # Truncate text if too long
                    max_chars = int(cell_width / 5)
                    if len(text) > max_chars:
                        text = text[:max_chars-2] + '..'

                    c.drawString(current_x + 3, current_y + cell_height/2 - 3, text)

                current_x += cell_width


def setup_handlers(web_app):
    host_pattern = ".*$"

    base_url = web_app.settings["base_url"]

    # Document converter endpoint
    converter_pattern = url_path_join(base_url, "jupyterlab-doc-reader-extension", "convert")
    handlers = [(converter_pattern, DocumentConverterHandler)]

    web_app.add_handlers(host_pattern, handlers)
