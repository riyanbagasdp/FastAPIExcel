from flask import Flask, render_template, request, send_file
import pandas as pd
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, PageBreak, Image, Flowable
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.units import cm
from reportlab.lib.utils import ImageReader
from reportlab.platypus import Table, TableStyle, SimpleDocTemplate, PageBreak
from dotenv import load_dotenv
import os
from docx import Document
from reportlab.platypus import Paragraph
import zipfile
from io import BytesIO
from flask import send_file
import smtplib
from email.message import EmailMessage
import mimetypes


app = Flask(__name__)

LOGO_PATH = "static/logo_bi.jpg"
LOGO_PATH3 = "static/logo.jpg"
WATERMARK_PATH = "static/private.png"

# === Reusable Components ===

def make_cover(nama, nip, periode):
    try:
        nama = str(nama).strip().upper()
        if not nama or nama in ["0", "NONE"]:
            nama = "(NAMA TIDAK VALID)"
    except:
        nama = "(NAMA ERROR)"

    try:
        if pd.isna(nip):
            raise ValueError
        nip_int = int(float(nip))
        if nip_int == 0:
            raise ValueError
        nip = str(nip_int)
    except:
        nip = "(NIP TIDAK VALID)"

    # ✅ Gambar aman
    try:
        logo_img = Image(LOGO_PATH, width=2.3*cm, height=1.4*cm)
    except:
        logo_img = Paragraph("LOGO BI TIDAK DITEMUKAN", None)

    data = [
        [logo_img],
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
    return TableWithWatermark(table, WATERMARK_PATH)

def make_logo_box():
    try:
        logo_img = Image(LOGO_PATH3, width=1 * cm, height=1 * cm)
    except:
        logo_img = Paragraph("LOGO TIDAK DITEMUKAN", None)

    data = [
        [''],
        [''],
        [logo_img],
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
        try:
            img = ImageReader(self.watermark_path)
            wm_width = self.width * 0.9
            wm_height = self.height * 0.9
            self.canv.drawImage(
                img, -wm_width / 2, -wm_height / 2,
                width=wm_width, height=wm_height,
                mask='auto', preserveAspectRatio=True
            )
        except Exception as e:
            print(f"❌ Gagal load watermark: {e}")
        self.canv.restoreState()
        self.table.drawOn(self.canv, 0, 0)


def make_detail(nama, nip, pipebi, ipebi, kopebi, zistabungan, pot, jumlah, periode):
    # Pastikan semua angka tidak None
    values = [pipebi, ipebi, kopebi, zistabungan, pot, jumlah]
    pipebi, ipebi, kopebi, zistabungan, pot, jumlah = [
        0 if pd.isna(v) else v for v in values
    ]

    # Validasi nama dan periode
    if not isinstance(nama, str) or not nama.strip():
        nama = "(NAMA TIDAK VALID)"
    if not isinstance(periode, str) or not periode.strip():
        periode = "(PERIODE TIDAK VALID)"

    # Validasi dan format NIP
    try:
        nip_display = f"{int(float(nip))}"
    except:
        nip_display = "(NIP TIDAK VALID)"

    content_table_data = [
        ['', ''],
        ['DATA RINCIAN POTONGAN LAIN-LAIN PADA SLIP GAJI PEGAWAI', ''],
        ['KANTOR PERWAKILAN BANK INDONESIA SOLO', ''],
        ['', ''],
        ['NIP', f': {nip_display}'],
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
def generate_pdf(bulan, tahun, file_path, output_buffer):
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

    doc = SimpleDocTemplate(output_buffer, pagesize=landscape(A4),
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
    

def send_email(to_email, subject, body, attachment_path, sender_email, sender_password):
    msg = EmailMessage()
    msg["Subject"] = subject
    msg["From"] = sender_email
    msg["To"] = to_email
    msg.set_content(body)

    with open(attachment_path, 'rb') as f:
        file_data = f.read()
        mime_type, _ = mimetypes.guess_type(attachment_path)
        maintype, subtype = mime_type.split("/") if mime_type else ("application", "octet-stream")
        msg.add_attachment(file_data, maintype=maintype, subtype=subtype,
                           filename=os.path.basename(attachment_path))

    with smtplib.SMTP("smtp.gmail.com", 587) as server:
        server.starttls()
        server.login(sender_email, sender_password)
        server.send_message(msg)

# === Generate PDF and Send ===
def generate_pdf_single(bulan, tahun, file_path, sender_email, sender_password):
    periode = f"{bulan.upper()} {tahun}"
    df = pd.read_excel(file_path, sheet_name="REKAP", header=[7, 8])

    df.columns = [' '.join([str(i).strip() for i in col if pd.notna(i)]).upper() for col in df.columns]
    colmap = {}

    for col in df.columns:
        if "NIP" in col and "NAMA" not in col:
            colmap["NIP"] = col
        elif "NAMA" in col:
            colmap["NAMA"] = col
        elif "PIPEBI" in col:
            colmap["PIPEBI"] = col
        elif "IPEBI" in col:
            colmap["IPEBI"] = col
        elif "KOPEBI" in col:
            colmap["KOPEBI"] = col
        elif "ZIS" in col or "TABUNGAN" in col:
            colmap["ZIS"] = col
        elif "KESEHATAN" in col:
            colmap["POT"] = col
        elif "EMAIL" in col.upper() and "EMAIL" not in colmap:
            colmap["EMAIL"] = col

    df = df[[colmap[k] for k in ["NIP", "NAMA", "PIPEBI", "IPEBI", "KOPEBI", "ZIS", "POT", "EMAIL"]]]
    df.columns = ["NIP", "NAMA", "PIPEBI", "IPEBI", "KOPEBI", "ZIS", "POT", "EMAIL"]

    df.fillna(0, inplace=True)
    df["NIP"] = df["NIP"].astype(str).str.replace(r"\.0$", "", regex=True).str.strip()
    df["EMAIL"] = df["EMAIL"].astype(str).str.strip()
    df["NAMA"] = df["NAMA"].astype(str).str.strip()
    df = df[df["EMAIL"].str.contains("@", na=False)]

    pot_cols = ["PIPEBI", "IPEBI", "KOPEBI", "ZIS", "POT"]
    df[pot_cols] = df[pot_cols].apply(pd.to_numeric, errors="coerce").fillna(0)
    df["JUMLAH"] = df[pot_cols].sum(axis=1)

    for _, row in df.iterrows():
        buffer = BytesIO()
        doc = SimpleDocTemplate(buffer, pagesize=landscape(A4),
                                leftMargin=1.5 * cm, rightMargin=1.5 * cm,
                                topMargin=1.5 * cm, bottomMargin=1.5 * cm)

        slip_content = Table([
            [make_cover(row['NAMA'], row['NIP'], periode)],
            [make_logo_box()],
            [make_detail(row['NAMA'], row['NIP'],
                         row['PIPEBI'], row['IPEBI'], row['KOPEBI'],
                         row['ZIS'], row['POT'], row['JUMLAH'], periode)]
        ], colWidths=[13.5 * cm])

        layout = Table([[slip_content]], colWidths=[27 * cm], rowHeights=[16.5 * cm])
        layout.setStyle(TableStyle([
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ]))

        doc.build([layout])
        buffer.seek(0)

        filename = f"{row['NIP']}_{row['NAMA'].replace(' ', '_')}.pdf"

        try:
            send_email_with_buffer(
                to_email=row["EMAIL"],
                subject=f"Slip Gaji Anda - {periode}",
                body=f"Yth. {row['NAMA']},\n\nBerikut terlampir slip gaji Anda.\n\nSalam,\nUnit Management Intern Kantor Perwakilan Bank Indonesia Solo",
                pdf_buffer=buffer,
                filename=filename,
                sender_email=sender_email,
                sender_password=sender_password
            )
            print(f"✅ Email terkirim ke {row['EMAIL']}")
        except Exception as e:
            print(f"❌ Gagal kirim email ke {row['EMAIL']}: {e}")

def send_email_with_buffer(to_email, subject, body, pdf_buffer, filename, sender_email, sender_password):
    msg = EmailMessage()
    msg["Subject"] = subject
    msg["From"] = sender_email
    msg["To"] = to_email
    msg.set_content(body)

    pdf_data = pdf_buffer.getvalue()
    msg.add_attachment(pdf_data, maintype="application", subtype="pdf", filename=filename)

    with smtplib.SMTP("smtp.gmail.com", 587) as server:
        server.starttls()
        server.login(sender_email, sender_password)
        server.send_message(msg)


# === Flask Routes ===
@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        bulan = request.form.get('bulan')
        tahun = request.form.get('tahun')
        file = request.files.get('file')
        action = request.form.get('action')

        if not bulan or not tahun or not file:
            return "❌ Mohon lengkapi bulan, tahun, dan file."

        # Baca file Excel ke memori
        excel_file = BytesIO(file.read())

        # Ambil variabel lingkungan (dari .env jika lokal, dari Railway jika hosting)
        if os.environ.get("RAILWAY_STATIC_URL") is None:
            from dotenv import load_dotenv
            load_dotenv()

        sender_email = os.getenv("EMAIL_SENDER")
        sender_password = os.getenv("EMAIL_PASSWORD")

        if not sender_email or not sender_password:
            return "❌ Email pengirim belum dikonfigurasi di .env atau environment hosting"

        # === Kirim slip satu per satu ke email masing-masing
        if action == "single":
            try:
                generate_pdf_single(bulan, tahun, excel_file, sender_email, sender_password)
                return "✅ Slip berhasil dikirim ke email masing-masing karyawan!"
            except Exception as e:
                return f"❌ Gagal mengirim email: {e}"

        # === Gabungkan semua slip dalam satu file PDF untuk diunduh
        else:
            try:
                output_buffer = BytesIO()
                generate_pdf(bulan, tahun, excel_file, output_buffer)
                output_buffer.seek(0)

                filename = f"Slip_Gaji_{bulan}_{tahun}_gabungan.pdf"
                return send_file(
                    output_buffer,
                    as_attachment=True,
                    download_name=filename,
                    mimetype="application/pdf"
                )
            except Exception as e:
                return f"❌ Gagal generate PDF gabungan: {e}"

    return render_template('index.html')


if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port)