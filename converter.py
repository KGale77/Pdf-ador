import os
import sys
import subprocess
import shutil
import winreg

# Extension mappings
WORD_EXTS = {".doc", ".docx", ".docm", ".dot", ".dotx", ".rtf", ".odt"}
EXCEL_EXTS = {".xls", ".xlsx", ".xlsm", ".xlsb", ".csv", ".ods"}
PPT_EXTS = {".ppt", ".pptx", ".pptm", ".pot", ".potx", ".pps", ".ppsx", ".odp"}
SUPPORTED_EXTS = WORD_EXTS.union(EXCEL_EXTS).union(PPT_EXTS)

def find_libreoffice():
    """Attempts to find the soffice executable in common locations."""
    # 1. Check if in PATH
    soffice_path = shutil.which("soffice")
    if soffice_path:
        return soffice_path

    # 2. Check default Windows installation paths
    default_paths = [
        os.path.join(os.environ.get("ProgramFiles", "C:\\Program Files"), "LibreOffice", "program", "soffice.exe"),
        os.path.join(os.environ.get("ProgramFiles(x86)", "C:\\Program Files (x86)"), "LibreOffice", "program", "soffice.exe"),
    ]
    for path in default_paths:
        if os.path.exists(path):
            return path

    # 3. Check Windows Registry
    try:
        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\soffice.exe") as key:
            path, _ = winreg.QueryValueEx(key, "")
            if os.path.exists(path):
                return path
    except OSError:
        pass

    return None

def convert_with_office_com(input_path, output_path, ext):
    """Converts a document to PDF using Microsoft Office COM Automation."""
    import win32com.client
    import pythoncom

    # Ensure pythoncom is initialized for the current thread
    pythoncom.CoInitialize()
    
    input_path = os.path.abspath(input_path)
    output_path = os.path.abspath(output_path)
    
    # Ensure parent output directory exists
    os.makedirs(os.path.dirname(output_path), exist_ok=True)

    app = None
    doc = None
    try:
        if ext in WORD_EXTS:
            app = win32com.client.Dispatch("Word.Application")
            app.Visible = False
            app.DisplayAlerts = False
            doc = app.Documents.Open(input_path, ReadOnly=True)
            # wdFormatPDF = 17
            doc.SaveAs(output_path, FileFormat=17)
            doc.Close(SaveChanges=False)
            return True
            
        elif ext in EXCEL_EXTS:
            app = win32com.client.Dispatch("Excel.Application")
            app.Visible = False
            app.DisplayAlerts = False
            app.ScreenUpdating = False
            doc = app.Workbooks.Open(input_path, ReadOnly=True)
            # xlTypePDF = 0
            # xlQualityStandard = 0
            doc.ExportAsFixedFormat(0, output_path)
            doc.Close(SaveChanges=False)
            return True
            
        elif ext in PPT_EXTS:
            app = win32com.client.Dispatch("PowerPoint.Application")
            # To prevent windows from showing up in PowerPoint, we must open presentation with WithWindow=False
            # ppSaveAsPDF = 32
            # Some versions of PowerPoint PPT require we do not touch App.Visible before Open
            doc = app.Presentations.Open(input_path, ReadOnly=True, WithWindow=False)
            doc.SaveAs(output_path, 32)
            doc.Close()
            return True
        else:
            raise ValueError(f"Unsupported extension for COM conversion: {ext}")
            
    finally:
        # Clean up
        try:
            if doc:
                doc.Close()
        except Exception:
            pass
            
        try:
            if app:
                # Excel/Word require Quit, PowerPoint can be quit if it is not already running or we own it
                app.Quit()
        except Exception:
            pass
            
        pythoncom.CoUninitialize()

def convert_with_libreoffice(soffice_path, input_path, output_path):
    """Converts a document to PDF using LibreOffice headless command line."""
    input_path = os.path.abspath(input_path)
    output_path = os.path.abspath(output_path)
    
    # LibreOffice output directory
    output_dir = os.path.dirname(output_path)
    os.makedirs(output_dir, exist_ok=True)
    
    # LibreOffice will create a PDF named like the input file but with .pdf in the outdir.
    # So we'll let it do that, and then rename it to the exact output_path the user wants.
    input_name_no_ext = os.path.splitext(os.path.basename(input_path))[0]
    expected_pdf_path = os.path.join(output_dir, f"{input_name_no_ext}.pdf")
    
    # Remove any pre-existing temp output file to avoid conflicts
    if os.path.exists(expected_pdf_path):
        try:
            os.remove(expected_pdf_path)
        except Exception:
            pass

    cmd = [
        soffice_path,
        "--headless",
        "--convert-to", "pdf",
        "--outdir", output_dir,
        input_path
    ]
    
    # Run headless conversion
    startupinfo = None
    if sys.platform == "win32":
        startupinfo = subprocess.STARTUPINFO()
        startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
        
    result = subprocess.run(
        cmd,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
        check=True,
        startupinfo=startupinfo
    )
    
    if os.path.exists(expected_pdf_path):
        if os.path.abspath(expected_pdf_path) != os.path.abspath(output_path):
            if os.path.exists(output_path):
                os.remove(output_path)
            os.rename(expected_pdf_path, output_path)
        return True
        
    raise RuntimeError(f"LibreOffice finished but output PDF was not found. Stderr: {result.stderr}")

def convert_to_pdf(input_path, output_path=None, engine="auto", libreoffice_path=None):
    """
    Converts any supported MS Office file to PDF.
    
    Engines:
    - 'auto': Attempts MS Office COM first. If it fails or is not available, falls back to LibreOffice.
    - 'office': Exclusively uses MS Office COM automation.
    - 'libreoffice': Exclusively uses LibreOffice.
    """
    if not os.path.exists(input_path):
        raise FileNotFoundError(f"Input file not found: {input_path}")
        
    ext = os.path.splitext(input_path)[1].lower()
    if ext not in SUPPORTED_EXTS:
        raise ValueError(f"Unsupported file format: {ext}")
        
    if not output_path:
        output_path = os.path.splitext(input_path)[0] + ".pdf"
        
    errors = []
    
    # Try Microsoft Office COM
    if engine in ("auto", "office"):
        try:
            success = convert_with_office_com(input_path, output_path, ext)
            if success:
                return "office"
        except Exception as e:
            errors.append(f"MS Office COM error: {e}")
            if engine == "office":
                raise RuntimeError("; ".join(errors))

    # Try LibreOffice
    if engine in ("auto", "libreoffice"):
        soffice = libreoffice_path or find_libreoffice()
        if soffice:
            try:
                success = convert_with_libreoffice(soffice, input_path, output_path)
                if success:
                    return "libreoffice"
            except Exception as e:
                errors.append(f"LibreOffice error: {e}")
        else:
            errors.append("LibreOffice executable (soffice) not found.")
            
    # If both failed or target engine failed
    raise RuntimeError(f"Failed to convert file. Details:\n" + "\n".join(errors))

if __name__ == "__main__":
    # Quick CLI interface for testing
    if len(sys.argv) < 2:
        print("Usage: python converter.py <input_file> [output_file]")
        sys.exit(1)
        
    infile = sys.argv[1]
    outfile = sys.argv[2] if len(sys.argv) > 2 else None
    
    try:
        print(f"Converting '{infile}'...")
        used_engine = convert_to_pdf(infile, outfile)
        print(f"Success! Converted using engine: {used_engine}")
    except Exception as e:
        print(f"Error: {e}")
        sys.exit(1)
