# pdf_to_epub_final_complete.py
# Complete EPUB converter: full-page images for low-text pages + ALL embedded images + text always preserved

import argparse
import os
import re
import io
from pypdf import PdfReader
from ebooklib import epub
from PIL import Image, UnidentifiedImageError
import fitz  # PyMuPDF - for page rendering with BleedBox

def clean_chinese_text(text: str) -> str:
    text = re.sub(r'\s{2,}', ' ', text)
    chinese_pattern = r'[\u4e00-\u9fff\u3400-\u4dbf\uf900-\ufaff\u3000-\u303f\uff00-\uffef]'
    text = re.sub(f'({chinese_pattern}) +({chinese_pattern})', r'\1\2', text)
    text = re.sub(f'({chinese_pattern}) +({chinese_pattern})', r'\1\2', text)
    return text.strip()

def process_image(raw_data: bytes) -> tuple[bytes | None, str | None]:
    """Convert embedded image to RGB JPEG, safely skip unreadable ones"""
    try:
        img = Image.open(io.BytesIO(raw_data))
        if img.mode in ("CMYK", "CMYKA"):
            img = img.convert("RGB")
        elif img.mode != "RGB":
            img = img.convert("RGB")
        
        output = io.BytesIO()
        img.save(output, format="JPEG", quality=95)
        return output.getvalue(), "image/jpeg"
    except UnidentifiedImageError:
        return None, None
    except Exception as e:
        print(f"   [Warning] Skipping damaged/special image: {e}")
        return None, None

def render_page_as_image(page_fitz, dpi=300) -> bytes:
    """Render full page using BleedBox priority"""
    if not page_fitz.bleedbox.is_empty:
        rect = page_fitz.bleedbox
    elif not page_fitz.trimbox.is_empty:
        rect = page_fitz.trimbox
    elif not page_fitz.cropbox.is_empty:
        rect = page_fitz.cropbox
    else:
        rect = page_fitz.mediabox
    
    zoom = dpi / 72
    mat = fitz.Matrix(zoom, zoom)
    pix = page_fitz.get_pixmap(matrix=mat, clip=rect, colorspace=fitz.csRGB)
    
    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
    output = io.BytesIO()
    img.save(output, format="JPEG", quality=95)
    return output.getvalue()

def pdf_to_epub(input_pdf: str, output_epub: str, title: str = "Untitled Book", author: str = "Unknown Author", word_threshold: int = 100):
    reader = PdfReader(input_pdf)
    doc_fitz = fitz.open(input_pdf)
    
    book = epub.EpubBook()
    book.set_identifier('pdf-' + os.path.splitext(os.path.basename(input_pdf))[0])
    book.set_title(title)
    book.set_language('zh')
    book.add_author(author)
    
    # Enhanced CSS
    css_content = """
    body { text-align: justify; line-height: 1.8; font-family: serif; }
    h1, h2, h3 { text-align: center; page-break-after: avoid; margin: 1em 0; font-weight: bold; }
    p { margin: 0; padding: 0; }
    p.indent { text-indent: 2em; }
    p.noindent { text-indent: 0; }
    .fullpage-img {
        display: block;
        max-width: 100%;
        height: auto;
        margin: 2em auto;
        page-break-before: always;
        page-break-after: avoid;
        text-align: center;
    }
    .inline-img {
        display: block;
        max-width: 95%;
        height: auto;
        margin: 1.5em auto;
        text-align: center;
    }
    """
    style_css = epub.EpubItem(uid="style_css", file_name="style/style.css", media_type="text/css",
                              content=css_content.encode('utf-8'))
    book.add_item(style_css)
    
    chapter = epub.EpubHtml(title='Main Content', file_name='content.xhtml', lang='zh')
    chapter.add_item(style_css)
    
    html_content = f'<h1>{title}</h1>\n'
    
    is_first_para_after_heading = True
    embedded_image_counter = 0
    fullpage_image_counter = 0
    
    for page_num, (page_pypdf, page_fitz) in enumerate(zip(reader.pages, doc_fitz), 1):
        print(f"Processing page {page_num}...")
        
        text = page_pypdf.extract_text(extraction_mode="layout") or ""
        word_count = len(re.findall(r'\w+', text))
        
        page_html = ""
        
        # === 1. Low-text page → add full-page rendered image ===
        #if word_count < word_threshold:
        fullpage_image_counter += 1
        img_data = render_page_as_image(page_fitz, dpi=300)
        filename = f"images/fullpage_{page_num:03d}.jpg"
        uid = f"fullpage_{fullpage_image_counter}"

        epub_img = epub.EpubItem(uid=uid, file_name=filename, media_type="image/jpeg", content=img_data)
        book.add_item(epub_img)

        page_html += f'<p class="fullpage-img"><img src="{filename}" alt="Full-page image page {page_num}"/></p>\n'
        print(f"   → Added full-page image (low text: {word_count} words)")
        
        # === 2. Always extract and add ALL embedded images ===
        embedded_html = []
        if "/Resources" in page_pypdf and "/XObject" in page_pypdf["/Resources"]:
            xobjects = page_pypdf["/Resources"]["/XObject"].get_object()
            dir(xobjects)
            for obj_name in xobjects:
                xobj = xobjects[obj_name].get_object()
                if xobj.get("/Subtype") != "/Image":
                    continue
                try:
                    raw_data = xobj._data
                    if not raw_data:
                        continue
                except:
                    continue
                
                img_data, media_type = process_image(raw_data)
                if img_data is None:
                    continue
                
                embedded_image_counter += 1
                filename = f"images/embedded_{page_num:03d}_{embedded_image_counter:03d}.jpg"
                uid = f"embedded_{embedded_image_counter}"
                
                epub_img = epub.EpubItem(uid=uid, file_name=filename, media_type=media_type, content=img_data)
                book.add_item(epub_img)
                
                embedded_html.append(f'<p class="inline-img"><img src="../{filename}" alt="Embedded Image {embedded_image_counter}"/></p>\n')
        
        if embedded_html:
            print(f"   → Added {len(embedded_html)} embedded image(s)")
        
        # === 3. Always add extracted text (captions, etc.) ===
        if text.strip():
            raw_paragraphs = [p.strip() for p in text.split('\n\n') if p.strip()]
            for para in raw_paragraphs:
                lines = [line.strip() for line in para.split('\n') if line.strip()]
                if not lines:
                    continue
                combined = ' '.join(lines)
                combined = clean_chinese_text(combined)
                
                if len(combined) < 150 and combined.isupper():
                    if len(combined) < 80:
                        page_html += f'<h2>{combined}</h2>\n'
                    else:
                        page_html += f'<h3>{combined}</h3>\n'
                    is_first_para_after_heading = True
                else:
                    if is_first_para_after_heading:
                        page_html += f'<p class="noindent">{combined}</p>\n'
                        is_first_para_after_heading = False
                    else:
                        page_html += f'<p class="indent">{combined}</p>\n'
        
        # === 4. Append embedded images after text ===
        page_html += ''.join(embedded_html)
        
        # Add this page's content
        html_content += page_html
    
    # Finalize chapter
    chapter.content = f"""
    <html xmlns="http://www.w3.org/1999/xhtml">
    <head><title>{title}</title></head>
    <body>{html_content}</body>
    </html>
    """.encode('utf-8')
    
    book.add_item(chapter)
    book.toc = (epub.Link('content.xhtml', 'Main Content', 'content'),)
    book.add_item(epub.EpubNav())
    book.add_item(epub.EpubNcx())
    book.spine = ['nav', chapter]
    
    epub.write_epub(output_epub, book)
    print(f"\n=== EPUB Completed ===")
    print(f"File: {output_epub}")
    print(f"Full-page Images（non-text pages）: {fullpage_image_counter} images")
    print(f"Embedded Images: {embedded_image_counter} images")
    print(f"All texts are preserved")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="PDF to EPUB (Full-page Image + All embedded images + All preserved texts)")
    parser.add_argument("input_pdf", help="Input PDF File")
    parser.add_argument("output_epub", nargs="?", help="Output EPUB File (Same filename as default)")
    parser.add_argument("--title", default="Untitled Book", help="Book Title")
    parser.add_argument("--author", default="Unknown Author", help="Author")
    parser.add_argument("--threshold", type=int, default=100, help="Full-page image / less than 100 chars (Default: 100)")
    
    args = parser.parse_args()
    if args.output_epub is None:
        args.output_epub = os.path.splitext(args.input_pdf)[0] + ".epub"
    
    pdf_to_epub(args.input_pdf, args.output_epub, args.title, args.author, args.threshold)
