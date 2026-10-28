
# Count-PDF-Pages.ps1
# Counts PDF pages in a specified folder and all its subfolders.
# Creates an Excel-compatible CSV report.

$ErrorActionPreference = "Stop"

Write-Host ""
Write-Host "====================================="
Write-Host "       PDF PAGE COUNT TOOL"
Write-Host "====================================="
Write-Host ""

$folder = Read-Host "Paste the full path to the folder containing your PDFs"
$folder = $folder.Trim().Trim('"')

if (-not (Test-Path -LiteralPath $folder -PathType Container)) {
    Write-Host ""
    Write-Host "ERROR: That folder could not be found."
    Write-Host "Please check the path and try again."
    Read-Host "Press Enter to close"
    exit 1
}

try {
    $null = & py -c "import pypdf" 2>&1
    if ($LASTEXITCODE -ne 0) {
        throw "The Python pypdf library is not installed. Run: py -m pip install --user pypdf"
    }
}
catch {
    Write-Host ""
    Write-Host $_.Exception.Message
    Read-Host "Press Enter to close"
    exit 1
}

$output = Join-Path $folder "PDF_Page_Count_Report.csv"

$env:PDF_SCAN_FOLDER = (Resolve-Path -LiteralPath $folder).Path
$env:PDF_SCAN_OUTPUT = $output

$pythonCode = @'
import os
import csv
from pathlib import Path
from pypdf import PdfReader

root = Path(os.environ["PDF_SCAN_FOLDER"])
output = Path(os.environ["PDF_SCAN_OUTPUT"])

pdfs = sorted(
    (p for p in root.rglob("*")
     if p.is_file() and p.suffix.lower() == ".pdf"
     and p.resolve() != output.resolve()),
    key=lambda p: str(p).lower()
)

print(f"Found {len(pdfs)} PDF files.")
rows = []

for i, path in enumerate(pdfs, start=1):
    print(f"[{i}/{len(pdfs)}] {path.name}")
    try:
        reader = PdfReader(str(path), strict=False)

        if reader.is_encrypted:
            result = reader.decrypt("")
            if result == 0:
                raise ValueError("Password-protected PDF")

        pages = len(reader.pages)
        status = "OK"
    except Exception as e:
        pages = ""
        status = f"ERROR: {e}"

    rows.append({
        "File Name": path.name,
        "Pages": pages,
        "Size (MB)": round(path.stat().st_size / (1024 * 1024), 3),
        "Full Path": str(path),
        "Status": status
    })

with output.open("w", newline="", encoding="utf-8-sig") as f:
    writer = csv.DictWriter(
        f, fieldnames=["File Name", "Pages", "Size (MB)",
                       "Full Path", "Status"]
    )
    writer.writeheader()
    writer.writerows(rows)

successful = sum(1 for row in rows if row["Status"] == "OK")
print()
print(f"Successfully counted: {successful} of {len(pdfs)} PDFs")
print(f"Report saved to: {output}")
'@

Write-Host ""
Write-Host "Scanning folder and subfolders..."
Write-Host ""

try {
    $pythonCode | & py -
    if ($LASTEXITCODE -ne 0) {
        Write-Host "Python encountered an error. See the messages above."
    }
}
finally {
    Remove-Item Env:\PDF_SCAN_FOLDER -ErrorAction SilentlyContinue
    Remove-Item Env:\PDF_SCAN_OUTPUT -ErrorAction SilentlyContinue
}

Write-Host ""
Read-Host "Finished. Press Enter to close"
