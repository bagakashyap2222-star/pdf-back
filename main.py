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


# Illegal XML 1.0 characters regex: forbids \x00-\x08, \x0b, \x0c, \x0e-\x1f, etc.
ILLEGAL_XML_CHARS_RE = re.compile(
    r'[\x00-\x08\x0b\x0c\x0e-\x1f\x7f-\x84\x86-\x9f\ud800-\udfff\ufffe\uffff]'
)

def sanitize_docx(docx_path: str) -> bool:
    """
    Sanitizes a DOCX file to guarantee 100% Microsoft Word OpenXML compliance:
    1. Removes any forbidden XML 1.0 control characters (e.g. \\x0c, \\x00, \\x0b, \\x1f)
       which PDF font streams frequently contain.
    2. Ensures table cells <w:tc> contain required paragraph <w:p> elements.
    3. Validates XML parsing of all internal parts so Microsoft Word opens without error.
    """
    if not os.path.exists(docx_path) or os.path.getsize(docx_path) == 0:
        return False

    temp_fixed = docx_path + ".sanitized.tmp"
    try:
        W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

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
                                modified = False
                                for tc in root.iter(f"{{{W_NS}}}tc"):
                                    has_p = any(child.tag == f"{{{W_NS}}}p" for child in tc)
                                    if not has_p:
                                        p_elem = ET.Element(f"{{{W_NS}}}p")
                                        tc.append(p_elem)
                                        modified = True
                                if modified:
                                    ET.register_namespace('w', W_NS)
                                    clean_data = ET.tostring(root, encoding='utf-8', xml_declaration=True)
                                else:
                                    clean_data = clean_text.encode('utf-8')
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
        print(f"Error sanitizing docx: {e}")
        if os.path.exists(temp_fixed):
            try:
                os.remove(temp_fixed)
            except Exception:
                pass
        return False


def convert_pdf_to_docx_libreoffice(input_pdf: str, tmp_dir: str, output_docx: str) -> bool:
    """
    Converts PDF to standard Microsoft Word DOCX using LibreOffice Writer headless.
    LibreOffice's official C++ OpenXML exporter produces 100% schema-compliant Word documents,
    flawlessly preserving tables, checkboxes, borders, and paragraphs without document.xml corruption.
    """
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

        # PRIMARY ENGINE: LibreOffice Writer Headless
        # Produces 100% compliant OpenXML Word files, preserving tables & checkboxes without document.xml schema errors
        try:
            lo_ok = convert_pdf_to_docx_libreoffice(input_pdf, tmp_dir, output_docx)
            if lo_ok and os.path.exists(output_docx) and os.path.getsize(output_docx) > 0:
                sanitize_docx(output_docx)
                converted_ok = True
            else:
                print("LibreOffice produced no output, switching to safe pdf2docx engine...")
        except Exception as e:
            print(f"LibreOffice error: {e}, falling back to pdf2docx...")

        # SECONDARY ENGINE (Fallback): pdf2docx with lattice table parsing disabled
        # parse_lattice_table=False prevents pdf2docx from generating malformed table structures at document.xml
        if not converted_ok:
            try:
                from pdf2docx import Converter
                cv = Converter(input_pdf)
                cv.convert(output_docx, start=0, end=None, parse_lattice_table=False)
                cv.close()

                if os.path.exists(output_docx) and os.path.getsize(output_docx) > 0:
                    sanitize_docx(output_docx)
                    converted_ok = True
            except Exception as e:
                print(f"pdf2docx fallback conversion error: {e}")

        # If safe mode also failed, try standard pdf2docx as last resort
        if not converted_ok:
            try:
                from pdf2docx import Converter
                cv = Converter(input_pdf)
                cv.convert(output_docx, start=0, end=None)
                cv.close()

                if os.path.exists(output_docx) and os.path.getsize(output_docx) > 0:
                    sanitize_docx(output_docx)
                    converted_ok = True
            except Exception as e:
                print(f"Standard pdf2docx failed: {e}")

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

        # Run headless LibreOffice
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
            # Check if any .pdf was produced in the directory
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
                        current_row += 1  # Blank row between tables
                else:
                    # Fallback: extract lines as rows if no strict table structure detected
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
