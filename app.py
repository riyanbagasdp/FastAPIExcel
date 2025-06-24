from flask import Flask, render_template, request, send_file
import pandas as pd
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, PageBreak, Image, Flowable
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.units import cm
from reportlab.lib.utils import ImageReader
import os
from docx import Document

app = Flask(__name__)
UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

LOGO_PATH = "static/logo_bi.jpg"
LOGO_PATH3 = "static/logo.jpg"
WATERMARK_PATH = "static/private.png"

# === Reusable Components ===

def make_cover(nama, nip, periode):
    try:
        nama = str(nama).strip().upper()
        if not nama or nama == "0":
            nama = "(NAMA TIDAK VALID)"
    except:
        nama = "(NAMA ERROR)"

    try:
        nip_int = int(float(nip))
        if nip_int == 0:
            raise ValueError
        nip = str(nip_int)
    except:
        nip = "(NIP TIDAK VALID)"

    data = [
        [Image(LOGO_PATH, width=2.3*cm, height=1.4*cm)],
        ['DATA RINCIAN POTONGAN LAIN-LAIN PADA SLIP GAJI PEGAWAI'],
        [f'BULAN : {periode}'],
        [''],
        [nama],
        [f'NIP : {nip}'],
        [''],
    ]
    table = Table(data, colWidths=[12 * cm])
    table.setStyle(TableStyle([
        ('SPAN', (0, 0), (-1, 0)),
        ('SPAN', (0, 1), (-1, 1)),
        ('SPAN', (0, 2), (-1, 2)),
        ('FONTNAME', (0, 0), (-1, -1), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 8),
        ('ALIGN', (0, 0), (-1, 2), 'CENTER'),
        ('ALIGN', (0, 3), (-1, 3), 'CENTER'),
        ('ALIGN', (0, 4), (1, 6), 'CENTER'),
        ('FONTNAME', (0, 4), (-1, 5), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 4), (-1, 5), 12),
        ('BOX', (0, 0), (-1, -1), 0.75, colors.black),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
    ]))
    return TableWithWatermark (table, WATERMARK_PATH)

def make_logo_box():
    data = [
        [''],
        [''],
        [Image(LOGO_PATH3, width=1 * cm, height=1 * cm)],
        [''],
        ['KANTOR PERWAKILAN BANK INDONESIA SOLO'],
        ['Jl. Jend. Sudirman No.15, Kp. Baru, Kec. Ps. Kliwon, Kota Surakarta, Jawa Tengah 57111'],
        [''],
        [''],
    ]
    table = Table(data, colWidths=[12 * cm])
    table.setStyle(TableStyle([
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('FONTNAME', (0, 0), (-1, -1), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 8),
        ('FONTSIZE', (0, 5), (0, 5), 4),
        ('BOX', (0, 0), (-1, -1), 0.5, colors.black),
    ]))
    return table

class TableWithWatermark(Flowable):
    def __init__(self, table, watermark_path):
        super().__init__()
        self.table = table
        self.watermark_path = watermark_path

    def wrap(self, availWidth, availHeight):
        self.width, self.height = self.table.wrap(availWidth, availHeight)
        return self.width, self.height

    def draw(self):
        self.canv.saveState()
        self.canv.translate(self.width / 2, self.height / 2)
        self.canv.rotate(17)
        img = ImageReader(self.watermark_path)
        wm_width = self.width * 0.9
        wm_height = self.height * 0.9
        self.canv.drawImage(
            img, -wm_width / 2, -wm_height / 2,
            width=wm_width, height=wm_height,
            mask='auto', preserveAspectRatio=True
        )
        self.canv.restoreState()
        self.table.drawOn(self.canv, 0, 0)

def make_detail(nama, nip, pipebi, ipebi, kopebi, zistabungan, pot, jumlah, periode):
    content_table_data = [
        ['', ''],
        ['DATA RINCIAN POTONGAN LAIN-LAIN PADA SLIP GAJI PEGAWAI', ''],
        ['KANTOR PERWAKILAN BANK INDONESIA SOLO', ''],
        ['', ''],
        ['NIP', f': {int(nip)}'],
        ['Nama', f': {nama}'],
        ['Bulan', f': {periode}'],
        ['', ''],
        ['1. PIPEBI Solo', f": Rp {pipebi:,.0f}".replace(",", ".")],
        ['2. IPEBI Solo', f": Rp {ipebi:,.0f}".replace(",", ".")],
        ['3. KOPEBI Solo', f": Rp {kopebi:,.0f}".replace(",", ".")],
        ['4. ZIS & Tabungan', f": Rp {zistabungan:,.0f}".replace(",", ".")],
        ['5. Potongan Lain', f": Rp {pot:,.0f}".replace(",", ".")],
        ['', ''],
        ['JUMLAH POTONGAN LAIN-LAIN', f"Rp {jumlah:,.0f}".replace(",", ".")],
        ['', '']
    ]

    content_table = Table(content_table_data, colWidths=[6 * cm, 6 * cm])
    content_table.setStyle(TableStyle([
        ('SPAN', (0, 0), (1, 0)),
        ('SPAN', (0, 1), (1, 1)),
        ('SPAN', (0, 2), (1, 2)),
        ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
        ('FONTSIZE', (0, 0), (-1, -1), 8),
        ('LEADING', (0, 0), (-1, -1), 8),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 1),
        ('TOPPADDING', (0, 0), (-1, -1), 1),
        ('ALIGN', (0, 0), (1, 2), 'CENTER'),
        ('ALIGN', (0, 3), (1, 5), 'LEFT'),
        ('ALIGN', (0, 7), (-1, -1), 'LEFT'),
        ('LEFTPADDING', (0, 3), (0, 6), 35),
        ('LEFTPADDING', (0, 7), (0, 13), 45),
        ('LEFTPADDING', (0, 14), (-1, 15), 50),
        ('FONTNAME', (0, 1), (-1, 2), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 1), (-1, 2), 9),
        ('BOX', (0, 0), (-1, -1), 0.5, colors.black),
    ]))

    return TableWithWatermark(content_table, WATERMARK_PATH)

# === PDF Generation ===
def generate_docx(bulan, tahun, file_path, output_docx):
    periode = f"{str(bulan).upper()} {str(tahun)}"
    df = pd.read_excel(file_path, sheet_name="REKAP", header=[7, 8], nrows=59, engine="xlrd")
    df.columns = [' '.join([str(i).strip() for i in col if pd.notna(i)]).upper() for col in df.columns]
    df = df[['NIP UNNAMED: 1_LEVEL_1', 'NAMA UNNAMED: 2_LEVEL_1',
             'POTONGAN PIPEBI', 'POTONGAN IPEBI', 'POTONGAN KOPEBI',
             'POTONGAN TABUNGAN & ZIS', 'POTONGAN POT KESEHATAN']]
    df.columns = ['NIP', 'NAMA', 'PIPEBI', 'IPEBI', 'KOPEBI', 'TABUNGAN & ZIS', 'POT KESEHATAN']
    df.fillna(0, inplace=True)
    df['NIP'] = df['NIP'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
    df['NAMA'] = df['NAMA'].astype(str).str.strip()
    df = df[(df['NIP'].str.isdigit()) & (df['NIP'] != '0') & (df['NAMA'] != '') &
            (df['NAMA'].str.upper() != '0') & (~df['NAMA'].str.upper().str.contains("TOTAL|JUMLAH", na=False))]
    df["JUMLAH"] = df[['PIPEBI', 'IPEBI', 'KOPEBI', 'TABUNGAN & ZIS', 'POT KESEHATAN']].sum(axis=1)

    document = Document()
    document.add_heading(f'SLIP GAJI PEGAWAI - {periode}', 0)

    for _, row in df.iterrows():
        document.add_paragraph()
        document.add_paragraph(f'NIP: {row["NIP"]}')
        document.add_paragraph(f'Nama: {row["NAMA"]}')
        document.add_paragraph(f'Periode: {periode}')
        document.add_paragraph(f'1. PIPEBI Solo: Rp {row["PIPEBI"]:,.0f}'.replace(",", "."))
        document.add_paragraph(f'2. IPEBI Solo: Rp {row["IPEBI"]:,.0f}'.replace(",", "."))
        document.add_paragraph(f'3. KOPEBI Solo: Rp {row["KOPEBI"]:,.0f}'.replace(",", "."))
        document.add_paragraph(f'4. ZIS & Tabungan: Rp {row["TABUNGAN & ZIS"]:,.0f}'.replace(",", "."))
        document.add_paragraph(f'5. Potongan Lain: Rp {row["POT KESEHATAN"]:,.0f}'.replace(",", "."))
        document.add_paragraph(f'JUMLAH: Rp {row["JUMLAH"]:,.0f}'.replace(",", "."))
        document.add_paragraph('-' * 40)

    document.save(output_docx)
    
def generate_pdf(bulan, tahun, file_path, output_pdf):
    periode = f"{str(bulan).upper()} {str(tahun)}"
    df = pd.read_excel(file_path, sheet_name="REKAP", header=[7, 8], nrows=59, engine="xlrd")
    df.columns = [' '.join([str(i).strip() for i in col if pd.notna(i)]).upper() for col in df.columns]
    df = df[['NIP UNNAMED: 1_LEVEL_1', 'NAMA UNNAMED: 2_LEVEL_1',
             'POTONGAN PIPEBI', 'POTONGAN IPEBI', 'POTONGAN KOPEBI',
             'POTONGAN TABUNGAN & ZIS', 'POTONGAN POT KESEHATAN']]
    df.columns = ['NIP', 'NAMA', 'PIPEBI', 'IPEBI', 'KOPEBI', 'TABUNGAN & ZIS', 'POT KESEHATAN']
    df.fillna(0, inplace=True)
    df['NIP'] = df['NIP'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
    df['NAMA'] = df['NAMA'].astype(str).str.strip()
    df = df[(df['NIP'].str.isdigit()) & (df['NIP'] != '0') & (df['NAMA'] != '') &
            (df['NAMA'].str.upper() != '0') & (~df['NAMA'].str.upper().str.contains("TOTAL|JUMLAH", na=False))]
    df["JUMLAH"] = df[['PIPEBI', 'IPEBI', 'KOPEBI', 'TABUNGAN & ZIS', 'POT KESEHATAN']].sum(axis=1)

    doc = SimpleDocTemplate(output_pdf, pagesize=landscape(A4),
                            leftMargin=1.2*cm, rightMargin=1.2*cm, topMargin=1.2*cm, bottomMargin=1.2*cm)
    elements = []
    rows = df.to_dict(orient="records")

    for i in range(0, len(rows), 2):
        data = []
        for j in range(2):
            if i + j < len(rows):
                row = rows[i + j]
                vertical_stack = Table([
                    [make_cover(row['NAMA'], row['NIP'], periode)],
                    [make_logo_box()],
                    [make_detail(row['NAMA'], row['NIP'],
                                 row['PIPEBI'], row['IPEBI'], row['KOPEBI'],
                                 row['TABUNGAN & ZIS'], row['POT KESEHATAN'], row['JUMLAH'],
                                 periode)]
                ], colWidths=[13.5 * cm])
                vertical_stack.setStyle(TableStyle([('VALIGN', (0, 0), (-1, -1), 'TOP')]))
                data.append(vertical_stack)
            else:
                data.append([])

        row_table = Table([data], colWidths=[13.5 * cm, 13.5 * cm])
        row_table.setStyle(TableStyle([('VALIGN', (0, 0), (-1, -1), 'TOP')]))
        elements.append(row_table)
        elements.append(PageBreak())

    doc.build(elements)

# === Flask Routes ===

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        bulan = request.form.get('bulan')
        tahun = request.form.get('tahun')
        file = request.files.get('file')

        if not bulan or not tahun or not file:
            return "Mohon lengkapi bulan, tahun, dan file."

        save_path = os.path.join(UPLOAD_FOLDER, file.filename)
        file.save(save_path)

        output_pdf = os.path.join("downloads", f"Slip_Gaji_{bulan}_{tahun}.pdf")
        os.makedirs("downloads", exist_ok=True)

        try:
            generate_pdf(bulan, tahun, save_path, output_pdf)
            print(f"✅ PDF berhasil digenerate: {output_pdf}")
        except Exception as e:
            print(f"❌ Gagal generate PDF: {e}")
            return f"Gagal generate PDF: {e}"

        if not os.path.exists(output_pdf):
            return "❌ PDF tidak ditemukan."

        return send_file(output_pdf, as_attachment=True)

    return render_template('index.html')


if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port)

