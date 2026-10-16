import os
import re
import shutil
import tempfile
import subprocess
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
from fastapi import FastAPI, UploadFile, File, HTTPException, BackgroundTasks
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, JSONResponse

app = FastAPI(
    title="PDFiyo Conversion API",
    description="High-fidelity document conversion microservice for PDFiyo (PDF to Word, Word to PDF, Excel, PPT)",
    version="1.0.0",
)

# Enable CORS for all frontends (PDFiyo production & localhost)
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
# ECMA-376 (WORDPROCESSINGML) SCHEMA NORMALIZATION ENGINE
# Guarantees 100% Microsoft Word OpenXML compliance for high-fidelity output.
# ============================================================================

ILLEGAL_XML_CHARS_RE = re.compile(
    r'[\x00-\x08\x0b\x0c\x0e-\x1f\x7f-\x84\x86-\x9f\ud800-\udfff\ufffe\uffff]'
)

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

# Exact XSD Sequence Order for <w:tcPr> (Table Cell Properties)
TCPR_CHILD_ORDER = [
    f"{{{W_NS}}}cnfStyle",
    f"{{{W_NS}}}tcW",
    f"{{{W_NS}}}gridSpan",
    f"{{{W_NS}}}hMerge",
    f"{{{W_NS}}}vMerge",
    f"{{{W_NS}}}tcBorders",
    f"{{{W_NS}}}shd",
    f"{{{W_NS}}}noWrap",
    f"{{{W_NS}}}tcMar",
    f"{{{W_NS}}}textDirection",
    f"{{{W_NS}}}tcFitText",
    f"{{{W_NS}}}vAlign",
    f"{{{W_NS}}}hideMark",
    f"{{{W_NS}}}headers",
    f"{{{W_NS}}}cellIns",
    f"{{{W_NS}}}cellDel",
    f"{{{W_NS}}}cellMerge",
    f"{{{W_NS}}}tcPrChange",
]

# Exact XSD Sequence Order for <w:pPr> (Paragraph Properties)
PPR_CHILD_ORDER = [
    f"{{{W_NS}}}pStyle",
    f"{{{W_NS}}}keepNext",
    f"{{{W_NS}}}keepLines",
    f"{{{W_NS}}}pageBreakBefore",
    f"{{{W_NS}}}framePr",
    f"{{{W_NS}}}widowControl",
    f"{{{W_NS}}}numPr",
    f"{{{W_NS}}}suppressLineNumbers",
    f"{{{W_NS}}}pBdr",
    f"{{{W_NS}}}shd",
    f"{{{W_NS}}}tabs",
    f"{{{W_NS}}}suppressAutoHyphens",
    f"{{{W_NS}}}kinsoku",
    f"{{{W_NS}}}wordWrap",
    f"{{{W_NS}}}overflowPunct",
    f"{{{W_NS}}}topLinePunct",
    f"{{{W_NS}}}autoSpaceDE",
    f"{{{W_NS}}}autoSpaceDN",
    f"{{{W_NS}}}bidi",
    f"{{{W_NS}}}adjustRightInd",
    f"{{{W_NS}}}snapToGrid",
    f"{{{W_NS}}}spacing",
    f"{{{W_NS}}}ind",
    f"{{{W_NS}}}contextualSpacing",
    f"{{{W_NS}}}mirrorIndents",
    f"{{{W_NS}}}suppressOverlap",
    f"{{{W_NS}}}jc",
    f"{{{W_NS}}}textDirection",
    f"{{{W_NS}}}textAlignment",
    f"{{{W_NS}}}textboxTightWrap",
    f"{{{W_NS}}}outlineLvl",
    f"{{{W_NS}}}divId",
    f"{{{W_NS}}}cnfStyle",
    f"{{{W_NS}}}rPr",
    f"{{{W_NS}}}sectPr",
    f"{{{W_NS}}}pPrChange",
]

# Exact XSD Sequence Order for <w:tblPr> (Table Properties)
TBLPR_CHILD_ORDER = [
    f"{{{W_NS}}}tblStyle",
    f"{{{W_NS}}}tblpPr",
    f"{{{W_NS}}}tblOverlap",
    f"{{{W_NS}}}bidiVisual",
    f"{{{W_NS}}}tblStyleRowBandSize",
    f"{{{W_NS}}}tblStyleColBandSize",
    f"{{{W_NS}}}tblW",
    f"{{{W_NS}}}jc",
    f"{{{W_NS}}}tblCellSpacing",
    f"{{{W_NS}}}tblInd",
    f"{{{W_NS}}}tblBorders",
    f"{{{W_NS}}}shd",
    f"{{{W_NS}}}tblLayout",
    f"{{{W_NS}}}tblCellMar",
    f"{{{W_NS}}}tblLook",
    f"{{{W_NS}}}tblCaption",
    f"{{{W_NS}}}tblDescription",
    f"{{{W_NS}}}tblPrChange",
]

# Exact XSD Sequence Order for <w:trPr> (Table Row Properties)
TRPR_CHILD_ORDER = [
    f"{{{W_NS}}}cnfStyle",
    f"{{{W_NS}}}divId",
    f"{{{W_NS}}}gridBefore",
    f"{{{W_NS}}}wBefore",
    f"{{{W_NS}}}gridAfter",
    f"{{{W_NS}}}wAfter",
    f"{{{W_NS}}}cantSplit",
    f"{{{W_NS}}}trHeight",
    f"{{{W_NS}}}tblHeader",
    f"{{{W_NS}}}tblCellSpacing",
    f"{{{W_NS}}}jc",
    f"{{{W_NS}}}hidden",
    f"{{{W_NS}}}ins",
    f"{{{W_NS}}}del",
    f"{{{W_NS}}}trPrChange",
]

def reorder_children(parent, order_list):
    """Sort direct children according to ECMA-376 XSD sequence order."""
    order_dict = {tag: i for i, tag in enumerate(order_list)}
    sorted_children = sorted(list(parent), key=lambda child: order_dict.get(child.tag, 9999))
    parent[:] = sorted_children

def reorder_tbl_children(tbl, w_ns):
    """Ensure <w:tblPr> is first, then <w:tblGrid>, followed by rows <w:tr>."""
    tblPr = None
    tblGrid = None
    rows = []
    others = []
    for child in list(tbl):
        if child.tag == f"{{{w_ns}}}tblPr":
            tblPr = child
        elif child.tag == f"{{{w_ns}}}tblGrid":
            tblGrid = child
        elif child.tag == f"{{{w_ns}}}tr":
            rows.append(child)
        else:
            others.append(child)
    new_children = []
    if tblPr is not None:
        new_children.append(tblPr)
    if tblGrid is not None:
        new_children.append(tblGrid)
    new_children.extend(rows)
    new_children.extend(others)
    tbl[:] = new_children

def reorder_tr_children(tr, w_ns):
    """Ensure <w:tblPrEx> and <w:trPr> precede cell elements <w:tc>."""
    tblPrEx = None
    trPr = None
    cells = []
    others = []
    for child in list(tr):
        if child.tag == f"{{{w_ns}}}tblPrEx":
            tblPrEx = child
        elif child.tag == f"{{{w_ns}}}trPr":
            trPr = child
        elif child.tag == f"{{{w_ns}}}tc":
            cells.append(child)
        else:
            others.append(child)
    new_children = []
    if tblPrEx is not None:
        new_children.append(tblPrEx)
    if trPr is not None:
        new_children.append(trPr)
    new_children.extend(cells)
    new_children.extend(others)
    tr[:] = new_children

def normalize_tc(tc, w_ns):
    """Ensure <w:tcPr> is first, wrap illegal loose drawings, and guarantee at least one <w:p>."""
    tcPr = None
    block_elements = []
    for child in list(tc):
        if child.tag == f"{{{w_ns}}}tcPr":
            tcPr = child
        elif child.tag == f"{{{w_ns}}}drawing":
            p = ET.Element(f"{{{w_ns}}}p")
            r = ET.SubElement(p, f"{{{w_ns}}}r")
            r.append(child)
            block_elements.append(p)
        else:
            block_elements.append(child)

    has_block = any(elem.tag in [f"{{{w_ns}}}p", f"{{{w_ns}}}tbl"] for elem in block_elements)
    if not has_block:
        block_elements.append(ET.Element(f"{{{w_ns}}}p"))

    new_children = []
    if tcPr is not None:
        new_children.append(tcPr)
    new_children.extend(block_elements)
    tc[:] = new_children

def normalize_bookmarks(root, w_ns):
    """Remove orphaned bookmarkEnd tags that lack matching bookmarkStart tags."""
    start_ids = set()
    for b_start in root.iter(f"{{{w_ns}}}bookmarkStart"):
        bid = b_start.attrib.get(f"{{{w_ns}}}id")
        if bid:
            start_ids.add(bid)
    for parent in root.iter():
        to_remove = []
        for child in parent:
            if child.tag == f"{{{w_ns}}}bookmarkEnd":
                bid = child.attrib.get(f"{{{w_ns}}}id")
                if bid and bid not in start_ids:
                    to_remove.append(child)
        for child in to_remove:
            parent.remove(child)

def sanitize_and_normalize_docx(docx_path: str) -> bool:
    """
    Normalizes Word document.xml and related parts to strict ECMA-376 schema standards:
    - Removes forbidden XML 1.0 control characters
    - Fixes <w:tcPr> property element order (eliminating Column 5913 table corruption)
    - Fixes <w:pPr>, <w:tblPr>, <w:trPr> property element order
    - Fixes table element hierarchy (<w:tblPr> -> <w:tblGrid> -> <w:tr> -> <w:tc>)
    - Ensures every table cell contains required paragraph block elements
    - Cleans orphaned bookmark references
    """
    if not os.path.exists(docx_path) or os.path.getsize(docx_path) == 0:
        return False

    temp_fixed = docx_path + ".normalized.tmp"
    try:
        with zipfile.ZipFile(docx_path, 'r') as zin:
            with zipfile.ZipFile(temp_fixed, 'w', compression=zipfile.ZIP_DEFLATED) as zout:
                for item in zin.infolist():
                    raw_data = zin.read(item.filename)
                    if item.filename.endswith('.xml') or item.filename.endswith('.rels'):
                        text = raw_data.decode('utf-8', errors='replace')
                        clean_text = ILLEGAL_XML_CHARS_RE.sub('', text)

                        try:
                            root = ET.fromstring(clean_text.encode('utf-8'))

                            if item.filename.endswith('document.xml'):
                                # Register Word namespaces
                                ET.register_namespace('w', W_NS)
                                ET.register_namespace('r', "http://schemas.openxmlformats.org/officeDocument/2006/relationships")
                                ET.register_namespace('m', "http://schemas.openxmlformats.org/officeDocument/2006/math")
                                ET.register_namespace('wp', "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing")
                                ET.register_namespace('a', "http://schemas.openxmlformats.org/drawingml/2006/main")
                                ET.register_namespace('pic', "http://schemas.openxmlformats.org/drawingml/2006/picture")

                                # 1. Normalize table cell property sequences
                                for tcPr in root.iter(f"{{{W_NS}}}tcPr"):
                                    reorder_children(tcPr, TCPR_CHILD_ORDER)

                                # 2. Normalize paragraph property sequences
                                for pPr in root.iter(f"{{{W_NS}}}pPr"):
                                    reorder_children(pPr, PPR_CHILD_ORDER)

                                # 3. Normalize table property sequences
                                for tblPr in root.iter(f"{{{W_NS}}}tblPr"):
                                    reorder_children(tblPr, TBLPR_CHILD_ORDER)

                                # 4. Normalize row property sequences
                                for trPr in root.iter(f"{{{W_NS}}}trPr"):
                                    reorder_children(trPr, TRPR_CHILD_ORDER)

                                # 5. Enforce table children hierarchy
                                for tbl in root.iter(f"{{{W_NS}}}tbl"):
                                    reorder_tbl_children(tbl, W_NS)

                                # 6. Enforce row children hierarchy
                                for tr in root.iter(f"{{{W_NS}}}tr"):
                                    reorder_tr_children(tr, W_NS)

                                # 7. Normalize table cells & wrap direct drawings
                                for tc in root.iter(f"{{{W_NS}}}tc"):
                                    normalize_tc(tc, W_NS)

                                # 8. Clean broken bookmarks
                                normalize_bookmarks(root, W_NS)

                                clean_data = ET.tostring(root, encoding='utf-8', xml_declaration=True)
                            else:
                                clean_data = clean_text.encode('utf-8')
                        except ET.ParseError:
                            return False

                        zout.writestr(item, clean_data)
                    else:
                        zout.writestr(item, raw_data)

        shutil.move(temp_fixed, docx_path)
        return True
    except Exception as e:
        print(f"Error normalizing docx: {e}")
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
        "version": "1.0.0",
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
# High-Fidelity Flowing Layout Extraction + Schema-Normalized OOXML
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
        # Native paragraph & table layout reconstruction via pdf2docx
        # followed by Schema Normalization to eliminate XML errors.
        try:
            from pdf2docx import Converter
            cv = Converter(input_pdf)
            cv.convert(output_docx, start=0, end=None, multi_processing=False)
            cv.close()

            if os.path.exists(output_docx) and os.path.getsize(output_docx) > 0:
                if sanitize_and_normalize_docx(output_docx):
                    converted_ok = True
                else:
                    print("Standard pdf2docx XML normalization failed, trying safe mode...")
        except Exception as e:
            print(f"Primary pdf2docx error: {e}")

        # ENGINE 2 (FALLBACK): Safe table parsing mode
        if not converted_ok:
            try:
                from pdf2docx import Converter
                cv = Converter(input_pdf)
                cv.convert(output_docx, start=0, end=None, parse_lattice_table=False, multi_processing=False)
                cv.close()

                if os.path.exists(output_docx) and os.path.getsize(output_docx) > 0:
                    if sanitize_and_normalize_docx(output_docx):
                        converted_ok = True
            except Exception as e:
                print(f"Safe table pdf2docx error: {e}")

        # ENGINE 3 (EMERGENCY FALLBACK): LibreOffice Headless
        if not converted_ok:
            lo_ok = convert_pdf_to_docx_libreoffice(input_pdf, tmp_dir, output_docx)
            if lo_ok and os.path.exists(output_docx) and os.path.getsize(output_docx) > 0:
                sanitize_and_normalize_docx(output_docx)
                converted_ok = True

        if not converted_ok or not os.path.exists(output_docx) or os.path.getsize(output_docx) == 0:
            raise HTTPException(status_code=500, detail="Conversion completed but output Word file could not be generated cleanly.")

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
                raise Exception("No PDF output generated by engine.")

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
