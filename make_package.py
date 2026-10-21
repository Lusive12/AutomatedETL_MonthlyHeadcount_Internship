"""
make_package.py - Script Otomatis Pembuat Paket Distribusi Monthly HR Report

Cara menjalankan:
  .venv\\Scripts\\python.exe make_package.py
"""
import os
import shutil
import pathlib
import zipfile

def build_package():
    root = pathlib.Path(__file__).resolve().parent
    # Jika dijalankan dari folder Monthly HR Report Automation atau ETL Group
    if not (root / "dist").exists() and (root.parent / "Monthly HR Report Automation").exists():
        root = root.parent / "PHA Monthly HR Report Automationv1.6"

    dist_pkg = root / "dist" / "PHA_Monthly_HR_Automation_Packagev1.6"

    print("[1/6] Menyiapkan folder paket distribusi...")
    if dist_pkg.exists():
        shutil.rmtree(dist_pkg)
    dist_pkg.mkdir(parents=True, exist_ok=True)

    print("[2/6] Menyalin hasil kompilasi PyInstaller (_internal & .exe)...")
    src_dist = root / "dist" / "Run_HR_Automation"
    if not src_dist.exists():
        print(f"ERROR: Folder {src_dist} tidak ditemukan!")
        print("Jalankan kompilasi terlebih dahulu: .venv\\Scripts\\pyinstaller --noconfirm Run_HR_Automation.spec")
        return

    for item in src_dist.iterdir():
        dest = dist_pkg / item.name
        if item.is_dir():
            shutil.copytree(item, dest)
        else:
            shutil.copy2(item, dest)

    print("[3/6] Menyalin folder config dan input...")
    shutil.copytree(root / "config", dist_pkg / "config")
    shutil.copytree(root / "input", dist_pkg / "input")

    # Bersihkan file temporer Excel (~$*)
    for temp_file in dist_pkg.glob("input/**/~$*"):
        try:
            temp_file.unlink()
        except Exception:
            pass

    print("[4/6] Membuat folder output/ dan logs/ kosong...")
    (dist_pkg / "output").mkdir(exist_ok=True)
    (dist_pkg / "logs").mkdir(exist_ok=True)

    print("[5/6] Membuat file HR_USER_GUIDE.txt, RUN_ME.bat, dan menyalin Panduan Word...")
    guide_content = """# HR Data Automation - Quick Start Guide (PHA Server)
Selamat datang! Aplikasi ini mengonsolidasikan export data bulanan SAP ke dalam Laporan Akhir (Final HR Topic Report).

---
## Cara Menjalankan Aplikasi (3 Langkah Mudah):
1. Langkah 1: Buka folder 'input/' dan masukkan file-file SAP terbaru ke dalam subfolder yang sesuai.
2. Langkah 2: Klik ganda file 'Run_HR_Automation.exe' (atau 'RUN_ME.bat').
3. Langkah 3: Setelah proses selesai, buka folder 'output/' untuk mengambil:
   - FINAL_HR_Topic_Report_YYYYMMDD.xlsx (Laporan gabungan 46 kolom)
   - 00_processing_summary.xlsx (Ringkasan audit & status data)

---
## Referensi Subfolder Input:
- input/Shared/
    -> PHA_PQAH.xlsx, PHA_Headings.xlsx, PHA_Direct_Spv.xlsx, PHA_PA0105.xlsx, PHA_Job_Layer.xlsx, PHA_Employee_Level.xlsx, PHA_PA0021.xlsx
- input/Contract Type & Contract End Data/
    -> PHA_PA0016.xlsx, Contract Type_PHA_PA0016.xlsx
- input/Education, Institute, Branch of Study/
    -> PHA_PA0022.xlsx, Order_PHA_PA0022.xlsx, PHA_Branch of Study.xlsx
- input/Entity, Division, Area, Function/
    -> PHA_PA9000.xlsx
- input/General Cost Center & Cost Center Text/
    -> PHA_PA0027_01.xlsx, PHA_PA0027_02.xlsx, PHA_CSKT.xlsx
- input/Join Date & Year of Service/
    -> PHA_PA0041.xlsx
- input/Position Effective Date & Current Tenure/
    -> PHA_HRP1001.xlsx

---
## Catatan Penting:
- Pastikan semua file Excel ditutup sebelum menjalankan aplikasi.
- Cek folder 'logs/' jika ada kendala untuk melihat histori log dan file audit integritas data (<timestamp>_insight.xlsx).
- Buku panduan lengkap (.docx) tersedia pada file: PANDUAN_OPERASIONAL_HR_AUTOMATION.docx
"""
    (dist_pkg / "HR_USER_GUIDE.txt").write_text(guide_content, encoding="utf-8")

    # Launcher batch script
    bat_content = "@echo off\r\ncd /d \"%~dp0\"\r\ntitle Monthly HR Report Automation\r\necho Starting Monthly HR Report Automation...\r\nRun_HR_Automation.exe\r\npause\r\n"
    (dist_pkg / "RUN_ME.bat").write_text(bat_content, encoding="utf-8")

    # Copy Word Guidebook if present
    word_doc = root / "PHA_PANDUAN_OPERASIONAL_HR_AUTOMATION.docx"
    if word_doc.exists():
        shutil.copy2(word_doc, dist_pkg / "PHA_PANDUAN_OPERASIONAL_HR_AUTOMATION.docx")

    print("[6/6] Mengompres folder menjadi file ZIP siap kirim...")
    zip_path = root / "dist" / "Monthly_HR_Automation_Package.zip"
    if zip_path.exists():
        zip_path.unlink()

    with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED, compresslevel=9) as zipf:
        for file_path in dist_pkg.rglob("*"):
            archive_name = file_path.relative_to(dist_pkg.parent)
            zipf.write(file_path, archive_name)

    size_mb = zip_path.stat().st_size / (1024 * 1024)
    print("=" * 60)
    print("BERHASIL DIBUAT!")
    print(f"Folder Paket : {dist_pkg}")
    print(f"File ZIP     : {zip_path} ({size_mb:.2f} MB)")
    print("=" * 60)

if __name__ == "__main__":
    build_package()
