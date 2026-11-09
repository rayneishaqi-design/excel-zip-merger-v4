import csv
import io
import os
import shutil
import tempfile
import zipfile
from collections import Counter
from contextlib import closing
from pathlib import PurePosixPath

import streamlit as st
from openpyxl import load_workbook
import xlsxwriter
import pdfplumber

from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib import colors
from reportlab.lib.utils import simpleSplit

# =========================================================
# 2. KONFIGURASI APLIKASI
# =========================================================

st.set_page_config(
    page_title="Penggabung Data Besar",
    page_icon="📦",
    layout="wide"
)

st.title("📦 Penggabung ZIP, Excel, CSV, dan PDF — Data Besar")

st.write(
    "Gabungkan banyak ZIP, Excel, CSV, dan PDF "
    "dengan pemrosesan bertahap."
)

st.warning(
    "Jangan unggah data perusahaan rahasia ke layanan publik "
    "tanpa izin. File diproses pada server aplikasi."
)

st.caption(
    "Hasil penggabungan dibuat menjadi satu PDF dan otomatis "
    "berlanjut ke halaman berikutnya jika data banyak."
)


MAX_CSV_ROWS = 700_000
MAX_XLSX_ROWS = 1_000_000

EXTS = (".csv", ".xlsx", ".xlsm", ".pdf")


# =========================================================
# 3. FUNGSI DASAR
# =========================================================

def clean_header(vals):
    result = []
    seen = Counter()

    for i, value in enumerate(vals, 1):
        head = str(value).strip() if value is not None else ""
        head = head or f"Kolom_{i}"

        seen[head] += 1

        if seen[head] > 1:
            head = f"{head}__{seen[head]}"

        result.append(head)

    return result


def find_header(row):
    return any(
        v is not None and str(v).strip()
        for v in row
    )


def text_cell(v):
    if v is None:
        return ""

    if hasattr(v, "isoformat"):
        return (
            v.isoformat(sep=" ")
            if hasattr(v, "hour")
            else v.isoformat()
        )

    return str(v)


# =========================================================
# 4. MEMBACA FILE ZIP
# =========================================================

def keep_safe_zip(info):
    name = info.filename.replace("\\", "/")
    parts = PurePosixPath(name).parts

    if (
        info.is_dir()
        or ".." in parts
        or name.startswith("/")
        or "__MACOSX" in parts
    ):
        return False

    return name.lower().endswith(EXTS)


def copy_to_spool(source, threshold=16 * 1024 * 1024):
    f = tempfile.SpooledTemporaryFile(
        max_size=threshold,
        mode="w+b"
    )

    shutil.copyfileobj(
        source,
        f,
        length=1024 * 1024
    )

    f.seek(0)
    return f


def iter_upload_members(uploads):
    for uploaded in uploads:
        uploaded.seek(0)

        if uploaded.name.lower().endswith(".zip"):
            with zipfile.ZipFile(uploaded) as outer:
                for info in outer.infolist():

                    if not keep_safe_zip(info):
                        continue

                    if info.flag_bits & 1:
                        raise ValueError(
                            f"ZIP berpassword: {info.filename}"
                        )

                    with outer.open(info) as member:
                        with copy_to_spool(member) as f:
                            yield (
                                f"{uploaded.name}/{info.filename}",
                                f
                            )

        elif uploaded.name.lower().endswith(EXTS):
            yield uploaded.name, uploaded


# =========================================================
# 5. MEMBACA CSV
# =========================================================

def csv_rows(source):
    source.seek(0)

    head = source.read(4096)
    source.seek(0)

    enc = "utf-8-sig"

    for name in ("utf-8-sig", "cp1252", "latin-1"):
        try:
            head.decode(name, errors="strict")
            enc = name
            break
        except UnicodeDecodeError:
            pass

    wrapper = io.TextIOWrapper(
        source,
        encoding=enc,
        errors="replace",
        newline=""
    )

    try:
        sample = wrapper.read(16_384)
        wrapper.seek(0)

        try:
            dialect = csv.Sniffer().sniff(
                sample,
                delimiters=",;\t|"
            )
        except csv.Error:
            dialect = csv.excel

        yield from csv.reader(wrapper, dialect)

    finally:
        try:
            wrapper.detach()
        except (ValueError, OSError):
            pass


# =========================================================
# 6. MEMBACA PDF
# =========================================================

def pdf_rows(source):
    source.seek(0)

    with pdfplumber.open(source) as pdf:
        for page_number, page in enumerate(
            pdf.pages,
            start=1
        ):

            tables = page.extract_tables() or []

            for table_number, table in enumerate(
                tables,
                start=1
            ):

                cleaned_rows = []

                for row in table:
                    cleaned = [
                        " ".join(
                            str(cell)
                            .replace("\n", " ")
                            .split()
                        )
                        if cell is not None else ""
                        for cell in row
                    ]

                    if any(cleaned):
                        cleaned_rows.append(cleaned)

                if cleaned_rows:
                    sheet_name = (
                        f"Halaman_{page_number}_"
                        f"Tabel_{table_number}"
                    )

                    yield sheet_name, iter(cleaned_rows)


# =========================================================
# 7. MEMBACA EXCEL, CSV, DAN PDF
# =========================================================

def iterate_sheets(
    file_name,
    source,
    mode,
    wanted_sheet
):

    lower_name = file_name.lower()

    # CSV
    if lower_name.endswith(".csv"):
        yield "CSV", csv_rows(source)
        return

    # PDF
    if lower_name.endswith(".pdf"):
        yield from pdf_rows(source)
        return

    # EXCEL
    source.seek(0)

    with closing(
        load_workbook(
            source,
            read_only=True,
            data_only=True
        )
    ) as wb:

        if mode == "Sheet pertama":
            sheets = wb.worksheets[:1]

        elif mode == "Nama sheet tertentu":
            sheets = [
                s for s in wb.worksheets
                if s.title.lower() == wanted_sheet.lower()
            ]

        else:
            sheets = wb.worksheets

        for ws in sheets:
            ws.reset_dimensions()

            yield ws.title, ws.iter_rows(
                values_only=True
            )


# =========================================================
# 8. MEMERIKSA HEADER DAN KOLOM
# =========================================================

def scan_columns(
    uploads,
    mode,
    wanted_sheet,
    metadata
):

    headers = []
    seen = set()
    checked = 0
    problems = []

    for path, handle in iter_upload_members(uploads):
        try:
            for sheet, rows in iterate_sheets(
                path,
                handle,
                mode,
                wanted_sheet
            ):

                try:
                    for first in rows:
                        if find_header(first):
                            for col in clean_header(first):
                                if col not in seen:
                                    seen.add(col)
                                    headers.append(col)

                            checked += 1
                            break
                finally:
                    if hasattr(rows, "close"):
                        rows.close()

        except Exception as exc:
            problems.append([
                path,
                "",
                f"Gagal pemeriksaan header: {type(exc).__name__}: {exc}"
            ])

    if metadata:
        for field in ("__asal_file", "__asal_sheet"):
            if field in seen:
                raise ValueError(
                    f"Kolom sumber bernama {field} sudah ada."
                )

        headers = ["__asal_file", "__asal_sheet"] + headers

    return headers, checked, problems


# =========================================================
# 9. OUTPUT WRITER
# =========================================================

class OutputWriter:

    def __init__(self, fmt, out_folder, headers):
        self.fmt = fmt
        self.folder = out_folder
        self.headers = headers
        self.count = 0
        self.part = 1
        self.local_row = 0

        self.path = os.path.join(out_folder, "hasil_gabungan.pdf")
        self.page_width, self.page_height = landscape(A4)
        self.margin = 24
        self.bottom_margin = 30
        self.col_width = (self.page_width - 2 * self.margin) / max(1, len(headers))

        n = len(headers)
        self.font_size = 8 if n <= 5 else 7 if n <= 9 else 6 if n <= 13 else 5

        self.pdf = canvas.Canvas(
            self.path,
            pagesize=landscape(A4),
            pageCompression=1
        )
        self.pdf.setTitle("Hasil Penggabungan Data")
        self.pdf.setAuthor("Penggabung Data Besar")
        self.draw_page_header()

    def safe_text(self, value):
        return text_cell(value).encode(
            "latin-1", errors="replace"
        ).decode("latin-1")

    def draw_page_header(self):
        self.pdf.setFillColor(colors.HexColor("#17365D"))
        self.pdf.rect(
            0, self.page_height - 55,
            self.page_width, 55,
            fill=1, stroke=0
        )

        self.pdf.setFillColor(colors.white)
        self.pdf.setFont("Helvetica-Bold", 14)
        self.pdf.drawString(
            self.margin, self.page_height - 34,
            "HASIL PENGGABUNGAN DATA"
        )

        self.pdf.setFont("Helvetica", 7)
        self.pdf.drawRightString(
            self.page_width - self.margin,
            self.page_height - 34,
            f"Halaman {self.part}"
        )

        self.y = self.page_height - 75
        self.draw_table_header()

    def draw_table_header(self):
        row_height = 25
        y_bottom = self.y - row_height

        self.pdf.setFillColor(colors.HexColor("#DCE6F1"))
        self.pdf.rect(
            self.margin, y_bottom,
            self.page_width - 2 * self.margin,
            row_height, fill=1, stroke=0
        )

        for i, header in enumerate(self.headers):
            x = self.margin + i * self.col_width

            self.pdf.setStrokeColor(colors.HexColor("#8A9AAF"))
            self.pdf.rect(
                x, y_bottom, self.col_width,
                row_height, fill=0, stroke=1
            )

            self.pdf.setFillColor(colors.HexColor("#17365D"))
            self.pdf.setFont("Helvetica-Bold", self.font_size)

            lines = simpleSplit(
                self.safe_text(header),
                "Helvetica-Bold",
                self.font_size,
                self.col_width - 6
            )

            line_y = self.y - 10
            for line in lines[:2]:
                self.pdf.drawString(x + 3, line_y, line)
                line_y -= self.font_size + 1

        self.y = y_bottom

    def add(self, values):
        wrapped = []

        for value in values:
            lines = simpleSplit(
                self.safe_text(value),
                "Helvetica",
                self.font_size,
                max(10, self.col_width - 6)
            )
            wrapped.append(lines or [""])

        max_lines = max(1, max(len(v) for v in wrapped))
        line_height = self.font_size + 2
        row_height = max(19, max_lines * line_height + 6)

        if self.y - row_height < self.bottom_margin:
            self.pdf.showPage()
            self.part += 1
            self.draw_page_header()

        y_bottom = self.y - row_height

        for i, lines in enumerate(wrapped):
            x = self.margin + i * self.col_width

            if self.count % 2 == 1:
                self.pdf.setFillColor(colors.HexColor("#F5F7FA"))
                self.pdf.rect(
                    x, y_bottom, self.col_width,
                    row_height, fill=1, stroke=0
                )

            self.pdf.setStrokeColor(colors.HexColor("#B7C2D0"))
            self.pdf.rect(
                x, y_bottom, self.col_width,
                row_height, fill=0, stroke=1
            )

            self.pdf.setFillColor(colors.HexColor("#202A35"))
            self.pdf.setFont("Helvetica", self.font_size)

            line_y = self.y - line_height - 3
            for line in lines:
                if line_y < y_bottom + 2:
                    break
                self.pdf.drawString(x + 3, line_y, line)
                line_y -= line_height

        self.y = y_bottom
        self.count += 1
        self.local_row += 1

    def close(self):
        self.pdf.save()


# 10. PROSES PENGGABUNGAN
# =========================================================

def merge(
    uploads,
    mode,
    wanted_sheet,
    metadata,
    fmt,
    progress
):

    headers, inspected, errors = scan_columns(
        uploads,
        mode,
        wanted_sheet,
        metadata
    )

    if (
        not headers
        or (metadata and len(headers) == 2)
        or inspected == 0
    ):
        raise ValueError(
            "Tidak ada header/sheet yang terbaca. "
            "Periksa isi file."
        )

    folder = tempfile.mkdtemp(
        prefix="merger_large_"
    )

    output = OutputWriter(
        fmt,
        folder,
        headers
    )

    total = 0
    report = []

    try:

        for number, (path, source) in enumerate(
            iter_upload_members(uploads),
            1
        ):

            progress.caption(
                f"Memproses file ke-{number}: {path[-90:]}"
            )

            try:

                found_sheet = False

                for sheet, rows in iterate_sheets(
                    path,
                    source,
                    mode,
                    wanted_sheet
                ):

                    found_sheet = True
                    count = 0
                    it = iter(rows)

                    first = next(
                        (
                            row for row in it
                            if find_header(row)
                        ),
                        None
                    )

                    if first is None:
                        report.append([
                            path,
                            sheet,
                            0,
                            "Sheet kosong"
                        ])
                        continue

                    current_headers = clean_header(first)

                    mapper = {
                        col: i
                        for i, col in enumerate(current_headers)
                    }

                    fields = (
                        headers[2:]
                        if metadata
                        else headers
                    )

                    idx = [
                        mapper.get(col)
                        for col in fields
                    ]

                    for row in it:

                        if not find_header(row):
                            continue

                        values = [
                            row[j]
                            if j is not None and j < len(row)
                            else ""
                            for j in idx
                        ]

                        if metadata:
                            values = [
                                path,
                                sheet
                            ] + values

                        output.add(values)
                        count += 1

                    total += count

                    report.append([
                        path,
                        sheet,
                        count,
                        "OK"
                    ])

                    if hasattr(rows, "close"):
                        rows.close()

                if not found_sheet:
                    errors.append([
                        path,
                        "",
                        "Tidak ada tabel atau sheet yang cocok"
                    ])

            except Exception as exc:
                errors.append([
                    path,
                    "",
                    f"{type(exc).__name__}: {exc}"
                ])

        output.close()

        report_path = os.path.join(
            folder,
            "laporan.csv"
        )

        with open(
            report_path,
            "w",
            encoding="utf-8-sig",
            newline=""
        ) as f:

            w = csv.writer(f)

            w.writerow([
                "Asal file",
                "Asal sheet",
                "Baris tergabung",
                "Status"
            ])

            w.writerows(report)

            w.writerow([
                "TOTAL",
                "",
                total,
                ""
            ])

            w.writerow([])

            w.writerow([
                "FILE BERMASALAH",
                "SHEET",
                "KETERANGAN"
            ])

            w.writerows(errors)

        return {
            "folder": folder,
            "output": output.path,
            "report": report_path,
            "total": total,
            "parts": output.part,
            "nerrors": len(errors),
            "nfiles": len(report),
            "headers": len(headers)
        }

    except Exception:

        try:
            output.close()
        except Exception:
            pass

        shutil.rmtree(
            folder,
            ignore_errors=True
        )

        raise


# =========================================================
# 11. TAMPILAN UPLOAD
# =========================================================

files = st.file_uploader(
    "Pilih beberapa ZIP / Excel / CSV / PDF (bisa dicampur)",

    type=[
        "zip",
        "xlsx",
        "xlsm",
        "csv",
        "pdf"
    ],

    accept_multiple_files=True
)

st.caption(
    "PDF yang didukung adalah PDF dengan teks/tabel "
    "yang dapat diekstrak. PDF hasil scan/foto memerlukan OCR."
)

col1, col2 = st.columns(2)

with col1:

    mode = st.selectbox(
        "Sheet yang diproses",
        [
            "Semua sheet",
            "Sheet pertama",
            "Nama sheet tertentu"
        ]
    )

with col2:

    sheet = st.text_input(
        "Nama sheet",
        "orders",
        disabled=mode != "Nama sheet tertentu"
    )

metadata = st.checkbox(
    "Tambahkan asal file dan sheet",
    value=True
)

fmt = "PDF (.pdf)"

st.info(
    "Format hasil akhir: PDF. Semua data yang terbaca "
    "akan digabungkan menjadi satu dokumen PDF."
)


# 12. TOMBOL PENGGABUNGAN
# =========================================================

if st.button(
    "Gabungkan semua file",
    type="primary",
    disabled=not files
):

    if "last" in st.session_state:

        shutil.rmtree(
            st.session_state.last["folder"],
            ignore_errors=True
        )

        del st.session_state.last

    state = st.empty()

    try:

        with st.spinner(
            "Memproses file secara bertahap..."
        ):

            choice = "PDF (.pdf)"

            result = merge(
                files,
                mode,
                sheet,
                metadata,
                choice,
                state
            )

        st.session_state.last = result

    except Exception as exc:

        st.error(
            f"Proses gagal: {type(exc).__name__}: {exc}"
        )


# =========================================================
# 13. HASIL PENGGABUNGAN
# =========================================================

if "last" in st.session_state:

    data = st.session_state.last

    a, b, c, d = st.columns(4)

    a.metric(
        "Baris gabungan",
        f"{data['total']:,}"
    )

    b.metric(
        "Sheet sumber terbaca",
        data["nfiles"]
    )

    c.metric(
        "Bagian hasil",
        data["parts"]
    )

    d.metric(
        "Masalah",
        data["nerrors"]
    )

    if data["total"] == 0:

        st.error(
            "Tidak ada baris data yang terbaca! "
            "Periksa laporan dan file sumber."
        )

    else:

        st.success(
            "Proses selesai. Periksa laporan sebelum "
            "menggunakan hasil penggabungan."
        )

        with open(data["output"], "rb") as f:

            st.download_button(
                "📥 Download gabungan",

                f,

                file_name=os.path.basename(
                    data["output"]
                ),

                mime="application/pdf"
            )

    with open(data["report"], "rb") as f:

        st.download_button(
            "📋 Download laporan baris per file",

            f,

            file_name="laporan_penggabungan.csv",

            mime="text/csv"
        )

    st.caption(
        "Hasil utama tersedia dalam format PDF. "
        "Laporan pemeriksaan file dapat diunduh secara terpisah."
    )
