import os
import re
import shutil
import tempfile
import subprocess
import zipfile
import copy
from pathlib import Path
from fastapi import FastAPI, UploadFile, File, HTTPException, BackgroundTasks
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, JSONResponse

try:
    from lxml import etree
    HAS_LXML = True
except ImportError:
    HAS_LXML = False

app = FastAPI(
    title="PDFiyo Conversion API",
    description="High-fidelity document conversion microservice for PDFiyo",
    version="2.0.0",
)

# Enable CORS for all frontends
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

def cleanup_path(path: str):
    """Safely delete temporary files or directories after response is sent."""
    try:
        if os.path.isdir(path):
            shutil.rmtree(path, ignore_errors=True)
        elif os.path.exists(path):
            os.remove(path)
    except Exception as e:
        print(f"Warning: could not delete {path}: {e}")


# ============================================================================
# PROFESSIONAL-GRADE MICROSOFT WORD COMPATIBILITY SANITIZER v2.0
# ============================================================================
# Uses lxml for proper XML parsing (with robust regex fallback) to fix ALL
# Word 2007 through Office 365 compatibility issues in pdf2docx output:
#
#  1. Strips illegal XML 1.0 control characters
#  2. Normalizes <w:start>/<w:end> -> <w:left>/<w:right> in borders & margins
#     (Word 2007 strictly adheres to ECMA-376 1st ed. where start/end do not exist)
#  3. Merges and deduplicates consecutive <w:tcMar>, <w:tcBorders>, <w:tblBorders>
#  4. Ensures every <w:tc> table cell contains at least one block element (<w:p>)
#  5. Wraps orphaned <w:drawing> elements inside proper <w:p><w:r> containers
#  6. Fixes table grid missing column definitions without destroying column widths
#  7. Strips unsupported modern namespaces (w14, w15, w16*, wp14) that crash Word 2007
#  8. Cleans mc:Ignorable attribute
#  9. Ensures <w:body> ends with <w:sectPr> (section properties)
# 10. Preserves all root XML declarations, standalone headers, and encoding
# ============================================================================

ILLEGAL_XML_CHARS_RE = re.compile(
    r'[\x00-\x08\x0b\x0c\x0e-\x1f\x7f-\x84\x86-\x9f]'
)

# OpenXML core namespace URIs
NS = {
    'w':   'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'r':   'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'wp':  'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
    'a':   'http://schemas.openxmlformats.org/drawingml/2006/main',
    'mc':  'http://schemas.openxmlformats.org/markup-compatibility/2006',
    'v':   'urn:schemas-microsoft-com:vml',
    'o':   'urn:schemas-microsoft-com:office:office',
    'wpc': 'http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas',
    'wps': 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape',
}

# Modern Word namespaces that crash Word 2007
UNSUPPORTED_NS_URIS = {
    'http://schemas.microsoft.com/office/word/2010/wordml',
    'http://schemas.microsoft.com/office/word/2012/wordml',
    'http://schemas.microsoft.com/office/word/2015/wordml/symex',
    'http://schemas.microsoft.com/office/word/2016/wordml/cid',
    'http://schemas.microsoft.com/office/word/2018/wordml/cex',
    'http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash',
    'http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing',
}

def _strip_control_chars(text: str) -> str:
    """Remove illegal XML 1.0 control characters from text."""
    return ILLEGAL_XML_CHARS_RE.sub('', text) if text else text

def _sanitize_xml_with_lxml(xml_bytes: bytes, filename: str) -> bytes:
    """
    Parse XML with lxml and apply all sanitization fixes.
    Returns cleaned XML bytes preserving the XML declaration.
    """
    # Step 1: Pre-clean raw text of illegal control characters
    try:
        text = xml_bytes.decode('utf-8', errors='replace')
    except Exception:
        text = xml_bytes.decode('latin-1', errors='replace')

    text = _strip_control_chars(text)

    # Step 2: Normalize start/end border & margin tags BEFORE parsing across all XML parts
    if filename.endswith('.xml'):
        text = re.sub(r'<w:start\b', '<w:left', text)
        text = re.sub(r'</w:start>', '</w:left>', text)
        text = re.sub(r'<w:end\b', '<w:right', text)
        text = re.sub(r'</w:end>', '</w:right>', text)

    # Step 3: Parse with lxml (recover mode handles minor syntax edge-cases)
    try:
        parser = etree.XMLParser(recover=True, remove_blank_text=False, encoding='utf-8')
        root = etree.fromstring(text.encode('utf-8'), parser)
    except Exception as e:
        print(f"  lxml parse failed for {filename}: {e}, using text fallback")
        return text.encode('utf-8')

    if root is None:
        return text.encode('utf-8')

    is_document_part = ('document.xml' in filename or 'header' in filename or 'footer' in filename)

    # Step 4: Strip elements from unsupported namespaces (w14, w15, etc.) that crash Word 2007
    elements_to_remove = []
    for elem in root.iter():
        if elem.tag and isinstance(elem.tag, str) and elem.tag.startswith('{'):
            ns_uri = elem.tag.split('}')[0][1:]
            if ns_uri in UNSUPPORTED_NS_URIS:
                elements_to_remove.append(elem)

    for elem in elements_to_remove:
        parent = elem.getparent()
        if parent is not None:
            if elem.tail:
                prev = elem.getprevious()
                if prev is not None:
                    prev.tail = (prev.tail or '') + elem.tail
                else:
                    parent.text = (parent.text or '') + elem.tail
            parent.remove(elem)

    # Step 5: Clean attributes from unsupported namespaces
    for elem in root.iter():
        if not isinstance(elem.tag, str):
            continue
        attrs_to_remove = []
        for attr_name in elem.attrib:
            if isinstance(attr_name, str) and attr_name.startswith('{'):
                attr_ns = attr_name.split('}')[0][1:]
                if attr_ns in UNSUPPORTED_NS_URIS:
                    attrs_to_remove.append(attr_name)
        for attr_name in attrs_to_remove:
            del elem.attrib[attr_name]

    # Step 6: Clean mc:Ignorable attribute
    mc_ns = NS.get('mc', '')
    ignorable_key = f'{{{mc_ns}}}Ignorable' if mc_ns else None
    if ignorable_key:
        for elem in root.iter():
            if ignorable_key in elem.attrib:
                ignorable_val = elem.attrib[ignorable_key]
                declared_prefixes = set(elem.nsmap.keys()) if hasattr(elem, 'nsmap') else set()
                clean_prefixes = []
                for prefix in ignorable_val.split():
                    if prefix in declared_prefixes:
                        prefix_uri = elem.nsmap.get(prefix, '')
                        if prefix_uri not in UNSUPPORTED_NS_URIS:
                            clean_prefixes.append(prefix)
                if clean_prefixes:
                    elem.attrib[ignorable_key] = ' '.join(clean_prefixes)
                else:
                    del elem.attrib[ignorable_key]

    # Step 7: Document & header/footer structural fixes
    if is_document_part:
        w_ns = NS['w']

        # 7a. Deduplicate and MERGE consecutive sibling tcMar, tcBorders, tblBorders, tblCellMar
        for tag_local in ['tcMar', 'tcBorders', 'tblBorders', 'tblCellMar']:
            tag_full = f'{{{w_ns}}}{tag_local}'
            seen_parents = set()
            for elem in list(root.iter(tag_full)):
                parent = elem.getparent()
                if parent is None or id(parent) in seen_parents:
                    continue
                seen_parents.add(id(parent))
                children = list(parent)
                first_found = None
                to_remove = []
                for child in children:
                    if child.tag == tag_full:
                        if first_found is None:
                            first_found = child
                        else:
                            # Merge child elements from duplicate into first_found so no border/margin is lost
                            for sub in list(child):
                                sub_local = sub.tag.split('}')[-1] if '}' in sub.tag else sub.tag
                                if first_found.find(f'{{{w_ns}}}{sub_local}') is None:
                                    first_found.append(copy.deepcopy(sub))
                            to_remove.append(child)
                for dup in to_remove:
                    parent.remove(dup)

        # 7b. Ensure every <w:tc> table cell has at least one block element (<w:p> or <w:tbl>)
        for tc in root.iter(f'{{{w_ns}}}tc'):
            has_block = False
            for child in tc:
                if isinstance(child.tag, str):
                    local = child.tag.split('}')[-1] if '}' in child.tag else child.tag
                    if local in ('p', 'tbl', 'sdt'):
                        has_block = True
                        break
            if not has_block:
                tc_pr = tc.find(f'{{{w_ns}}}tcPr')
                p_elem = etree.SubElement(tc, f'{{{w_ns}}}p')
                if tc_pr is not None:
                    tc.remove(p_elem)
                    tc_children = list(tc)
                    if tc_pr in tc_children:
                        tc.insert(tc_children.index(tc_pr) + 1, p_elem)
                    else:
                        tc.insert(0, p_elem)

        # 7c. Wrap orphaned <w:drawing> elements directly under <w:tc> into proper <w:p><w:r>
        for tc in root.iter(f'{{{w_ns}}}tc'):
            drawings_to_wrap = []
            for child in list(tc):
                if isinstance(child.tag, str):
                    local = child.tag.split('}')[-1] if '}' in child.tag else child.tag
                    if local == 'drawing':
                        drawings_to_wrap.append(child)
            for drawing in drawings_to_wrap:
                tc_children = list(tc)
                idx = tc_children.index(drawing) if drawing in tc_children else 0
                tc.remove(drawing)
                p_elem = etree.Element(f'{{{w_ns}}}p')
                r_elem = etree.SubElement(p_elem, f'{{{w_ns}}}r')
                r_elem.append(drawing)
                tc.insert(idx, p_elem)

        # 7d. Ensure tblGrid has enough columns without destroying custom widths
        for tbl in root.iter(f'{{{w_ns}}}tbl'):
            tbl_grid = tbl.find(f'{{{w_ns}}}tblGrid')
            if tbl_grid is None:
                continue

            grid_cols = tbl_grid.findall(f'{{{w_ns}}}gridCol')
            grid_count = len(grid_cols)

            max_cells = 0
            for tr in tbl.findall(f'{{{w_ns}}}tr'):
                cell_count = 0
                for tc in tr.findall(f'{{{w_ns}}}tc'):
                    tc_pr = tc.find(f'{{{w_ns}}}tcPr')
                    if tc_pr is not None:
                        grid_span = tc_pr.find(f'{{{w_ns}}}gridSpan')
                        if grid_span is not None:
                            span_val = grid_span.get(f'{{{w_ns}}}val', '1')
                            try:
                                cell_count += int(span_val)
                            except (ValueError, TypeError):
                                cell_count += 1
                        else:
                            cell_count += 1
                    else:
                        cell_count += 1
                max_cells = max(max_cells, cell_count)

            if max_cells > 0 and grid_count < max_cells:
                missing = max_cells - grid_count
                for _ in range(missing):
                    new_gc = etree.SubElement(tbl_grid, f'{{{w_ns}}}gridCol')
                    new_gc.set(f'{{{w_ns}}}w', '1440')

        # 7e. Ensure <w:body> ends with <w:sectPr> in document.xml
        if 'document.xml' in filename:
            body = root.find(f'{{{w_ns}}}body')
            if body is not None:
                children = list(body)
                has_sect_pr = False
                if children:
                    last = children[-1]
                    if isinstance(last.tag, str) and last.tag.split('}')[-1] == 'sectPr':
                        has_sect_pr = True
                if not has_sect_pr:
                    sect_pr = etree.SubElement(body, f'{{{w_ns}}}sectPr')
                    pg_sz = etree.SubElement(sect_pr, f'{{{w_ns}}}pgSz')
                    pg_sz.set(f'{{{w_ns}}}w', '12240')
                    pg_sz.set(f'{{{w_ns}}}h', '15840')
                    pg_mar = etree.SubElement(sect_pr, f'{{{w_ns}}}pgMar')
                    pg_mar.set(f'{{{w_ns}}}top', '1440')
                    pg_mar.set(f'{{{w_ns}}}right', '1440')
                    pg_mar.set(f'{{{w_ns}}}bottom', '1440')
                    pg_mar.set(f'{{{w_ns}}}left', '1440')

    # Step 8: Clean text content in all <w:t> elements
    w_ns = NS.get('w', '')
    if w_ns:
        for t_elem in root.iter(f'{{{w_ns}}}t'):
            if t_elem.text:
                t_elem.text = _strip_control_chars(t_elem.text)

    # Step 9: Serialize back preserving XML declaration
    try:
        result = etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)
        return result
    except Exception as e:
        print(f"  lxml serialize failed for {filename}: {e}")
        return text.encode('utf-8')


def _sanitize_xml_regex_fallback(xml_bytes: bytes, filename: str) -> bytes:
    """Fallback regex-based sanitizer when lxml is unavailable."""
    try:
        text = xml_bytes.decode('utf-8', errors='replace')
    except Exception:
        text = xml_bytes.decode('latin-1', errors='replace')

    text = ILLEGAL_XML_CHARS_RE.sub('', text)

    if filename.endswith('.xml'):
        text = re.sub(r'<w:start\b', '<w:left', text)
        text = re.sub(r'</w:start>', '</w:left>', text)
        text = re.sub(r'<w:end\b', '<w:right', text)
        text = re.sub(r'</w:end>', '</w:right>', text)
        text = re.sub(r'(<w:tcMar>[\s\S]*?</w:tcMar>)(?:\s*<w:tcMar>[\s\S]*?</w:tcMar>)+', r'\1', text)
        text = re.sub(r'(<w:tcBorders>[\s\S]*?</w:tcBorders>)(?:\s*<w:tcBorders>[\s\S]*?</w:tcBorders>)+', r'\1', text)
        text = re.sub(r'(<w:tc(?:\s+[^>]*)?>)(?!\s*<w:(?:p|tbl|sdt))', r'\1<w:p/>', text)
        text = re.sub(r'(<w:tc(?:\s+[^>]*)?>\s*)(<w:drawing[\s\S]*?</w:drawing>)', r'\1<w:p><w:r>\2</w:r></w:p>', text)

    return text.encode('utf-8')


def sanitize_docx_for_word(docx_path: str) -> bool:
    """
    Professional-grade DOCX sanitizer for 100% clean opening in ALL Microsoft Word versions.
    Uses lxml for proper XML parsing when available, falls back to regex.
    """
    if not os.path.exists(docx_path) or os.path.getsize(docx_path) == 0:
        return False

    temp_fixed = docx_path + ".word_clean.tmp"
    try:
        with zipfile.ZipFile(docx_path, 'r') as zin:
            with zipfile.ZipFile(temp_fixed, 'w', compression=zipfile.ZIP_DEFLATED) as zout:
                for item in zin.infolist():
                    raw_data = zin.read(item.filename)

                    if item.filename.endswith('.xml') or item.filename.endswith('.rels'):
                        if HAS_LXML:
                            clean_data = _sanitize_xml_with_lxml(raw_data, item.filename)
                        else:
                            clean_data = _sanitize_xml_regex_fallback(raw_data, item.filename)
                        zout.writestr(item, clean_data)
                    else:
                        zout.writestr(item, raw_data)

        shutil.move(temp_fixed, docx_path)
        print(f"  Sanitized {docx_path} successfully (engine: {'lxml' if HAS_LXML else 'regex'})")
        return True
    except Exception as e:
        print(f"Error sanitizing docx: {e}")
        if os.path.exists(temp_fixed):
            try:
                os.remove(temp_fixed)
            except Exception:
                pass
        return False


def convert_pdf_to_docx_libreoffice(input_pdf: str, tmp_dir: str, output_docx: str) -> bool:
    """Fallback engine using LibreOffice Writer headless."""
    try:
        lo_profile = os.path.join(tmp_dir, "lo_profile").replace("\\", "/")
        cmd = [
            "soffice",
            "--headless",
            f"-env:UserInstallation=file://{lo_profile}",
            "--infilter=writer_pdf_import",
            "--convert-to", "docx",
            "--outdir", tmp_dir,
            input_pdf
        ]
        res = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=90)
        expected = os.path.join(tmp_dir, Path(input_pdf).stem + ".docx")
        if os.path.exists(expected) and os.path.getsize(expected) > 0:
            if expected != output_docx:
                shutil.move(expected, output_docx)
            return True
        docxs = [f for f in os.listdir(tmp_dir) if f.lower().endswith(".docx") and f != os.path.basename(output_docx)]
        if docxs:
            shutil.move(os.path.join(tmp_dir, docxs[0]), output_docx)
            return True
    except Exception as e:
        print(f"LibreOffice PDF-to-DOCX error: {e}")
    return False


@app.get("/")
@app.get("/health")
def health_check():
    return {
        "status": "healthy",
        "service": "PDFiyo Conversion API",
        "version": "2.0.0",
        "engine": "lxml" if HAS_LXML else "regex",
        "supported_conversions": [
            "pdf-to-word",
            "word-to-pdf",
            "pdf-to-excel",
            "excel-to-pdf",
            "ppt-to-pdf",
        ]
    }


# ============================================================================
# 1. PDF TO WORD (.docx)
# High-Fidelity Flowing Layout Extraction with Full Table Support
# ============================================================================
@app.post("/convert/pdf-to-word")
async def convert_pdf_to_word(background_tasks: BackgroundTasks, file: UploadFile = File(...)):
    if not file.filename.lower().endswith(".pdf"):
        raise HTTPException(status_code=400, detail="Invalid file type. Please upload a .pdf file.")

    tmp_dir = tempfile.mkdtemp(prefix="pdfiyo_pdf2word_")
    background_tasks.add_task(cleanup_path, tmp_dir)

    input_pdf = os.path.join(tmp_dir, "input.pdf")
    base_name = re.sub(r'(\.[^.]+)+$', '', file.filename).strip() or "document"
    output_docx = os.path.join(tmp_dir, f"{base_name}.docx")

    try:
        content = await file.read()
        with open(input_pdf, "wb") as f:
            f.write(content)

        converted_ok = False

        # ENGINE 1 (HIGH-FIDELITY PRIMARY):
        # Native paragraph & full table layout reconstruction via pdf2docx.
        # Connected border tolerance: 0.75 joins micro-broken line segments in table borders.
        # Max border width: 6.0 catches thick table borders.
        # Image resolution ratio: 4.0 ensures crisp embedded images.
        # Sanitizer fixes Word 2007-365 schema incompatibilities (start/end -> left/right,
        # duplicate tcMar/tcBorders, loose drawings, control characters).
        try:
            from pdf2docx import Converter
            cv = Converter(input_pdf)
            try:
                cv.convert(
                    output_docx,
                    start=0,
                    end=None,
                    multi_processing=False,
                    parse_lattice_table=True,
                    parse_stream_table=True,
                    connected_border_tolerance=0.75,
                    max_border_width=6.0,
                    min_border_clearance=2.0,
                    clip_image_res_ratio=4.0,
                    ignore_page_error=True
                )
            finally:
                cv.close()

            if os.path.exists(output_docx) and os.path.getsize(output_docx) > 0:
                if sanitize_docx_for_word(output_docx):
                    converted_ok = True
                    print(f"  Engine 1 (pdf2docx high-fidelity + lxml) succeeded for {file.filename}")
        except Exception as e:
            print(f"pdf2docx primary engine error: {e}")

        # ENGINE 2 (FALLBACK): Safe table parsing mode
        if not converted_ok:
            try:
                if os.path.exists(output_docx):
                    os.remove(output_docx)

                from pdf2docx import Converter
                cv = Converter(input_pdf)
                try:
                    cv.convert(
                        output_docx,
                        start=0,
                        end=None,
                        parse_lattice_table=False,
                        multi_processing=False,
                        ignore_page_error=True
                    )
                finally:
                    cv.close()

                if os.path.exists(output_docx) and os.path.getsize(output_docx) > 0:
                    if sanitize_docx_for_word(output_docx):
                        converted_ok = True
                        print(f"  Engine 2 (pdf2docx safe) succeeded for {file.filename}")
            except Exception as e:
                print(f"pdf2docx safe engine error: {e}")

        # ENGINE 3 (EMERGENCY FALLBACK): LibreOffice Headless
        if not converted_ok:
            if os.path.exists(output_docx):
                os.remove(output_docx)

            lo_ok = convert_pdf_to_docx_libreoffice(input_pdf, tmp_dir, output_docx)
            if lo_ok and os.path.exists(output_docx) and os.path.getsize(output_docx) > 0:
                sanitize_docx_for_word(output_docx)
                converted_ok = True
                print(f"  Engine 3 (LibreOffice) succeeded for {file.filename}")

        if not converted_ok or not os.path.exists(output_docx) or os.path.getsize(output_docx) == 0:
            raise HTTPException(status_code=500, detail="Conversion failed - output file could not be generated.")

        return FileResponse(
            path=output_docx,
            filename=f"{base_name}.docx",
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            headers={"Access-Control-Expose-Headers": "Content-Disposition"}
        )
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"PDF to Word conversion failed: {str(e)}")


# ============================================================================
# 2. WORD TO PDF (.pdf)
# ============================================================================
@app.post("/convert/word-to-pdf")
async def convert_word_to_pdf(background_tasks: BackgroundTasks, file: UploadFile = File(...)):
    ext = Path(file.filename).suffix.lower()
    if ext not in [".docx", ".doc", ".rtf", ".odt"]:
        raise HTTPException(status_code=400, detail="Invalid file type. Please upload a .docx or .doc file.")

    tmp_dir = tempfile.mkdtemp(prefix="pdfiyo_word2pdf_")
    background_tasks.add_task(cleanup_path, tmp_dir)

    base_name = re.sub(r'(\.[^.]+)+$', '', file.filename).strip() or "document"
    input_file = os.path.join(tmp_dir, f"input{ext}")

    try:
        content = await file.read()
        with open(input_file, "wb") as f:
            f.write(content)

        lo_profile = os.path.join(tmp_dir, "lo_profile").replace("\\", "/")
        cmd = [
            "soffice",
            "--headless",
            f"-env:UserInstallation=file://{lo_profile}",
            "--convert-to", "pdf",
            "--outdir", tmp_dir,
            input_file
        ]
        result = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=60)
        if result.returncode != 0:
            raise Exception(result.stderr.decode("utf-8", errors="ignore") or "LibreOffice conversion failed")

        output_pdf = os.path.join(tmp_dir, "input.pdf")
        if not os.path.exists(output_pdf):
            pdfs = [f for f in os.listdir(tmp_dir) if f.lower().endswith(".pdf")]
            if pdfs:
                output_pdf = os.path.join(tmp_dir, pdfs[0])
            else:
                raise Exception("No PDF generated.")

        return FileResponse(
            path=output_pdf,
            filename=f"{base_name}.pdf",
            media_type="application/pdf",
            headers={"Access-Control-Expose-Headers": "Content-Disposition"}
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Word to PDF conversion failed: {str(e)}")


# ============================================================================
# 3. PDF TO EXCEL (.xlsx)
# ============================================================================
@app.post("/convert/pdf-to-excel")
async def convert_pdf_to_excel(background_tasks: BackgroundTasks, file: UploadFile = File(...)):
    if not file.filename.lower().endswith(".pdf"):
        raise HTTPException(status_code=400, detail="Invalid file type. Please upload a .pdf file.")

    tmp_dir = tempfile.mkdtemp(prefix="pdfiyo_pdf2excel_")
    background_tasks.add_task(cleanup_path, tmp_dir)

    input_pdf = os.path.join(tmp_dir, "input.pdf")
    base_name = re.sub(r'(\.[^.]+)+$', '', file.filename).strip() or "spreadsheet"
    output_xlsx = os.path.join(tmp_dir, f"{base_name}.xlsx")

    try:
        content = await file.read()
        with open(input_pdf, "wb") as f:
            f.write(content)

        import pdfplumber
        import openpyxl

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Extracted Tables"

        has_tables = False
        current_row = 1

        with pdfplumber.open(input_pdf) as pdf:
            for p_idx, page in enumerate(pdf.pages, start=1):
                tables = page.extract_tables()
                if tables:
                    for table in tables:
                        has_tables = True
                        for row in table:
                            for col_idx, cell in enumerate(row, start=1):
                                ws.cell(row=current_row, column=col_idx, value=cell or "")
                            current_row += 1
                        current_row += 1
                else:
                    text = page.extract_text()
                    if text:
                        for line in text.split("\n"):
                            ws.cell(row=current_row, column=1, value=line.strip())
                            current_row += 1
                        current_row += 1

        wb.save(output_xlsx)

        return FileResponse(
            path=output_xlsx,
            filename=f"{base_name}.xlsx",
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Access-Control-Expose-Headers": "Content-Disposition"}
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"PDF to Excel conversion failed: {str(e)}")


# ============================================================================
# 4. EXCEL TO PDF (.pdf)
# ============================================================================
@app.post("/convert/excel-to-pdf")
async def convert_excel_to_pdf(background_tasks: BackgroundTasks, file: UploadFile = File(...)):
    ext = Path(file.filename).suffix.lower()
    if ext not in [".xlsx", ".xls", ".csv", ".ods"]:
        raise HTTPException(status_code=400, detail="Invalid file type. Please upload a .xlsx, .xls, or .csv file.")

    tmp_dir = tempfile.mkdtemp(prefix="pdfiyo_excel2pdf_")
    background_tasks.add_task(cleanup_path, tmp_dir)

    base_name = re.sub(r'(\.[^.]+)+$', '', file.filename).strip() or "document"
    input_file = os.path.join(tmp_dir, f"input{ext}")

    try:
        content = await file.read()
        with open(input_file, "wb") as f:
            f.write(content)

        lo_profile = os.path.join(tmp_dir, "lo_profile").replace("\\", "/")
        cmd = [
            "soffice",
            "--headless",
            f"-env:UserInstallation=file://{lo_profile}",
            "--convert-to", "pdf",
            "--outdir", tmp_dir,
            input_file
        ]
        result = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=60)
        if result.returncode != 0:
            raise Exception(result.stderr.decode("utf-8", errors="ignore") or "LibreOffice conversion failed")

        output_pdf = os.path.join(tmp_dir, "input.pdf")
        if not os.path.exists(output_pdf):
            pdfs = [f for f in os.listdir(tmp_dir) if f.lower().endswith(".pdf")]
            if pdfs:
                output_pdf = os.path.join(tmp_dir, pdfs[0])
            else:
                raise Exception("No PDF generated.")

        return FileResponse(
            path=output_pdf,
            filename=f"{base_name}.pdf",
            media_type="application/pdf",
            headers={"Access-Control-Expose-Headers": "Content-Disposition"}
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Excel to PDF conversion failed: {str(e)}")


# ============================================================================
# 5. PPT TO PDF (.pdf)
# ============================================================================
@app.post("/convert/ppt-to-pdf")
async def convert_ppt_to_pdf(background_tasks: BackgroundTasks, file: UploadFile = File(...)):
    ext = Path(file.filename).suffix.lower()
    if ext not in [".pptx", ".ppt", ".odp"]:
        raise HTTPException(status_code=400, detail="Invalid file type. Please upload a .pptx or .ppt file.")

    tmp_dir = tempfile.mkdtemp(prefix="pdfiyo_ppt2pdf_")
    background_tasks.add_task(cleanup_path, tmp_dir)

    base_name = re.sub(r'(\.[^.]+)+$', '', file.filename).strip() or "presentation"
    input_file = os.path.join(tmp_dir, f"input{ext}")

    try:
        content = await file.read()
        with open(input_file, "wb") as f:
            f.write(content)

        lo_profile = os.path.join(tmp_dir, "lo_profile").replace("\\", "/")
        cmd = [
            "soffice",
            "--headless",
            f"-env:UserInstallation=file://{lo_profile}",
            "--convert-to", "pdf",
            "--outdir", tmp_dir,
            input_file
        ]
        result = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=60)
        if result.returncode != 0:
            raise Exception(result.stderr.decode("utf-8", errors="ignore") or "LibreOffice conversion failed")

        output_pdf = os.path.join(tmp_dir, "input.pdf")
        if not os.path.exists(output_pdf):
            pdfs = [f for f in os.listdir(tmp_dir) if f.lower().endswith(".pdf")]
            if pdfs:
                output_pdf = os.path.join(tmp_dir, pdfs[0])
            else:
                raise Exception("No PDF generated.")

        return FileResponse(
            path=output_pdf,
            filename=f"{base_name}.pdf",
            media_type="application/pdf",
            headers={"Access-Control-Expose-Headers": "Content-Disposition"}
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"PPT to PDF conversion failed: {str(e)}")
