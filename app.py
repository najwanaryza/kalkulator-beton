"""
CIVIL CONCRETE - Aplikasi Perhitungan Beton Teknik Sipil
Aplikasi web berbasis Streamlit untuk kebutuhan akademik.

Jalankan lokal:
    pip install -r requirements.txt
    streamlit run app.py

Catatan: estimasi material menggunakan komposisi per m3 yang dimasukkan pengguna,
bukan mix design SNI otomatis. Verifikasi seluruh hasil terhadap SNI dan
spesifikasi proyek yang berlaku sebelum digunakan dalam pekerjaan konstruksi.
"""
import math
from io import BytesIO
from datetime import datetime

import pandas as pd
import streamlit as st

st.set_page_config(page_title="CIVIL CONCRETE", page_icon="🏗️", layout="wide")

# ---------- FUNGSI PERHITUNGAN ----------
def valid_number(value, label, allow_zero=False):
    try:
        number = float(value)
    except (TypeError, ValueError):
        raise ValueError(f"Masukkan angka yang valid untuk {label}.")
    if not math.isfinite(number):
        raise ValueError(f"Nilai {label} harus berupa angka yang valid.")
    if number < 0 or (not allow_zero and number == 0):
        op = "tidak boleh negatif" if allow_zero else "harus lebih besar dari nol"
        raise ValueError(f"{label.capitalize()} {op}.")
    return number

def calculate_strength(shape, dimension, load_kn):
    d = valid_number(dimension, "dimensi benda uji")
    load = valid_number(load_kn, "beban maksimum")
    area = math.pi * d**2 / 4 if shape == "Silinder" else d**2
    return area, load * 1000 / area

def calculate_volume(element, a, b, c, qty, waste):
    a = valid_number(a, "dimensi 1")
    b = valid_number(b, "dimensi 2")
    c = valid_number(c, "dimensi 3")
    qty = valid_number(qty, "jumlah elemen")
    waste = valid_number(waste, "cadangan", allow_zero=True)
    if element == "Silinder / tiang bulat":
        one_volume = math.pi * (a / 2) ** 2 * b
        formula = "π × (diameter/2)² × tinggi"
    else:
        one_volume = a * b * c
        formula = "panjang × lebar × tinggi/tebal"
    base = one_volume * qty
    total = base * (1 + waste / 100)
    return base, total, formula

def calculate_material(volume, cement, sand, gravel, water, waste):
    volume = valid_number(volume, "volume beton")
    waste = valid_number(waste, "cadangan", allow_zero=True)
    rates = [
        valid_number(cement, "semen", allow_zero=True),
        valid_number(sand, "pasir", allow_zero=True),
        valid_number(gravel, "kerikil", allow_zero=True),
        valid_number(water, "air", allow_zero=True),
    ]
    volume_final = volume * (1 + waste / 100)
    names = ["Semen (kg)", "Pasir (kg)", "Kerikil (kg)", "Air (liter)"]
    return pd.DataFrame({
        "Material": names,
        "Kebutuhan per m3": rates,
        "Volume dihitung (m3)": [volume_final] * 4,
        "Total kebutuhan": [round(volume_final * x, 2) for x in rates],
    })

def make_excel(dataframes):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        wrote = False
        for sheet, frame in dataframes.items():
            if frame is not None and not frame.empty:
                frame.to_excel(writer, index=False, sheet_name=sheet[:31])
                wrote = True
        if not wrote:
            pd.DataFrame({"Keterangan": ["Belum ada hasil perhitungan"]}).to_excel(
                writer, index=False, sheet_name="Rekap"
            )
    output.seek(0)
    return output.getvalue()

def make_pdf(dataframes):
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib.styles import getSampleStyleSheet
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
    output = BytesIO()
    doc = SimpleDocTemplate(output, pagesize=landscape(A4),
                            rightMargin=28, leftMargin=28, topMargin=28, bottomMargin=28)
    styles = getSampleStyleSheet()
    story = [
        Paragraph("LAPORAN CIVIL CONCRETE", styles["Title"]),
        Paragraph("Rekapitulasi perhitungan beton teknik sipil", styles["Normal"]),
        Paragraph("Waktu: " + datetime.now().strftime("%d-%m-%Y %H:%M"), styles["Normal"]),
        Spacer(1, 14),
    ]
    has_data = False
    for title, frame in dataframes.items():
        if frame is None or frame.empty:
            continue
        has_data = True
        story.append(Paragraph(title.replace("_", " "), styles["Heading2"]))
        rows = [list(map(str, frame.columns))]
        rows += [[str(v) for v in row] for row in frame.astype(str).values.tolist()]
        table = Table(rows, repeatRows=1)
        table.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#17365D")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("GRID", (0, 0), (-1, -1), 0.4, colors.grey),
            ("FONTSIZE", (0, 0), (-1, -1), 7),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ]))
        story.extend([table, Spacer(1, 12)])
    if not has_data:
        story.append(Paragraph("Belum ada hasil perhitungan.", styles["Normal"]))
    story.append(Paragraph(
        "Catatan: hasil merupakan estimasi berdasarkan data dan asumsi input pengguna. "
        "Verifikasi terhadap SNI, spesifikasi proyek, dan tenaga ahli.",
        styles["Italic"]))
    doc.build(story)
    output.seek(0)
    return output.getvalue()

# ---------- TAMPILAN ----------
st.title("🏗️ CIVIL CONCRETE")
st.markdown("**Aplikasi perhitungan beton untuk pekerjaan infrastruktur teknik sipil**")
st.warning(
    "Aplikasi ini merupakan alat bantu akademik dan estimasi awal. "
    "Perhitungan kebutuhan material memakai komposisi input pengguna, bukan "
    "mix design SNI otomatis dan bukan pengganti desain struktur oleh tenaga ahli."
)

if "hasil" not in st.session_state:
    st.session_state.hasil = {}

tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "🧪 Uji Kuat Tekan", "📐 Volume Beton", "🧱 Kebutuhan Material",
    "📏 Pengujian Slump", "📊 Biaya & Ekspor"
])

with tab1:
    st.subheader("Perhitungan kuat tekan beton")
    c1, c2, c3 = st.columns(3)
    shape = c1.selectbox("Bentuk benda uji", ["Silinder", "Kubus"])
    dimension = c2.number_input("Diameter silinder / sisi kubus (mm)", min_value=0.1, value=150.0)
    load = c3.number_input("Beban maksimum (kN)", min_value=0.1, value=500.0)
    target = st.number_input("Mutu beton target (MPa)", min_value=0.1, value=25.0)
    if st.button("Hitung Kuat Tekan", type="primary"):
        try:
            area, fc = calculate_strength(shape, dimension, load)
            status = "Memenuhi target input" if fc >= target else "Di bawah target input"
            st.session_state.hasil["Uji_Kuat_Tekan"] = pd.DataFrame([{
                "Bentuk benda uji": shape, "Dimensi (mm)": dimension,
                "Luas tekan (mm2)": round(area, 2), "Beban maksimum (kN)": load,
                "Kuat tekan (MPa)": round(fc, 3), "Mutu target (MPa)": target,
                "Evaluasi": status
            }])
            st.metric("Kuat tekan beton", f"{fc:.3f} MPa")
            st.write(f"Luas tekan: **{area:,.2f} mm²**")
            st.info(status)
            st.dataframe(st.session_state.hasil["Uji_Kuat_Tekan"], use_container_width=True)
        except ValueError as e:
            st.error(str(e))

with tab2:
    st.subheader("Perhitungan volume elemen beton")
    element = st.selectbox("Jenis elemen", [
        "Pelat / lantai / perkerasan", "Balok", "Kolom",
        "Fondasi telapak", "Silinder / tiang bulat"
    ])
    x1, x2, x3 = st.columns(3)
    if element == "Silinder / tiang bulat":
        a = x1.number_input("Diameter (m)", min_value=0.001, value=0.5)
        b = x2.number_input("Tinggi (m)", min_value=0.001, value=3.0)
        c = x3.number_input("Tidak digunakan", min_value=0.0, value=1.0, disabled=True)
    else:
        a = x1.number_input("Panjang (m)", min_value=0.001, value=10.0)
        b = x2.number_input("Lebar (m)", min_value=0.001, value=5.0)
        c = x3.number_input("Tebal / tinggi (m)", min_value=0.001, value=0.15)
    q1, q2 = st.columns(2)
    qty = q1.number_input("Jumlah elemen", min_value=1.0, value=1.0)
    waste = q2.number_input("Cadangan (%)", min_value=0.0, value=5.0)
    if st.button("Hitung Volume", type="primary"):
        try:
            base, total, formula = calculate_volume(element, a, b, c, qty, waste)
            st.session_state.hasil["Volume_Beton"] = pd.DataFrame([{
                "Elemen": element, "Volume dasar (m3)": round(base, 3),
                "Cadangan (%)": waste, "Jumlah akhir (m3)": round(total, 3),
                "Jumlah elemen": qty, "Rumus": formula
            }])
            m1, m2 = st.columns(2)
            m1.metric("Volume geometris", f"{base:.3f} m³")
            m2.metric("Volume + cadangan", f"{total:.3f} m³")
            st.dataframe(st.session_state.hasil["Volume_Beton"], use_container_width=True)
        except ValueError as e:
            st.error(str(e))

with tab3:
    st.subheader("Estimasi kebutuhan material")
    st.caption("Masukkan kebutuhan material per 1 m³ berdasarkan mix design atau spesifikasi yang telah disetujui.")
    m1, m2 = st.columns(2)
    volume = m1.number_input("Volume beton (m³)", min_value=0.001, value=10.0)
    mwaste = m2.number_input("Cadangan material (%)", min_value=0.0, value=5.0)
    m1, m2, m3, m4 = st.columns(4)
    cement = m1.number_input("Semen (kg/m³)", min_value=0.0, value=350.0)
    sand = m2.number_input("Pasir (kg/m³)", min_value=0.0, value=700.0)
    gravel = m3.number_input("Kerikil (kg/m³)", min_value=0.0, value=1050.0)
    water = m4.number_input("Air (liter/m³)", min_value=0.0, value=190.0)
    if st.button("Hitung Kebutuhan Material", type="primary"):
        try:
            frame = calculate_material(volume, cement, sand, gravel, water, mwaste)
            st.session_state.hasil["Kebutuhan_Material"] = frame
            st.success(f"Volume yang dihitung termasuk cadangan: {volume * (1 + mwaste/100):.3f} m³")
            st.dataframe(frame, use_container_width=True)
        except ValueError as e:
            st.error(str(e))

with tab4:
    st.subheader("Evaluasi nilai slump")
    s1, s2, s3 = st.columns(3)
    slump = s1.number_input("Nilai slump (mm)", min_value=0.0, value=100.0)
    smin = s2.number_input("Batas bawah target (mm)", min_value=0.0, value=75.0)
    smax = s3.number_input("Batas atas target (mm)", min_value=0.0, value=125.0)
    if st.button("Evaluasi Slump", type="primary"):
        if smax < smin:
            st.error("Batas atas harus lebih besar atau sama dengan batas bawah.")
        else:
            status = "Dalam rentang target input" if smin <= slump <= smax else "Di luar rentang target input"
            st.session_state.hasil["Pengujian_Slump"] = pd.DataFrame([{
                "Nilai slump (mm)": slump, "Batas bawah (mm)": smin,
                "Batas atas (mm)": smax, "Evaluasi": status
            }])
            st.metric("Hasil evaluasi", status)
            st.caption("Rentang target harus ditentukan berdasarkan spesifikasi pekerjaan.")

with tab5:
    st.subheader("Analisis biaya material")
    material_df = st.session_state.hasil.get("Kebutuhan_Material")
    if material_df is None:
        st.info("Hitung kebutuhan material pada tab sebelumnya terlebih dahulu.")
    else:
        st.dataframe(material_df, use_container_width=True)
        st.caption("Masukkan harga semen, pasir, dan kerikil per kg; air per liter.")
        p1, p2, p3, p4 = st.columns(4)
        pc = p1.number_input("Harga semen (Rp/kg)", min_value=0.0, value=1500.0)
        ps = p2.number_input("Harga pasir (Rp/kg)", min_value=0.0, value=300.0)
        pg = p3.number_input("Harga kerikil (Rp/kg)", min_value=0.0, value=350.0)
        pw = p4.number_input("Harga air (Rp/liter)", min_value=0.0, value=10.0)
        if st.button("Hitung Biaya Material", type="primary"):
            prices = [pc, ps, pg, pw]
            quantities = material_df["Total kebutuhan"].astype(float).tolist()
            costs = [q * p for q, p in zip(quantities, prices)]
            cost_df = pd.DataFrame({
                "Material": material_df["Material"], "Jumlah": quantities,
                "Harga satuan (Rp)": prices, "Subtotal (Rp)": costs
            })
            st.session_state.hasil["Biaya_Material"] = cost_df
            st.metric("Total estimasi biaya material", f"Rp {sum(costs):,.0f}")
            st.dataframe(cost_df, use_container_width=True)

    st.divider()
    st.subheader("Ekspor laporan")
    available = {k: v for k, v in st.session_state.hasil.items() if v is not None and not v.empty}
    if available:
        st.write("Data yang akan dimasukkan ke laporan:")
        st.write(", ".join(available.keys()))
        try:
            st.download_button(
                "⬇️ Unduh Excel",
                data=make_excel(available),
                file_name="Laporan_Civil_Concrete.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            st.download_button(
                "⬇️ Unduh PDF",
                data=make_pdf(available),
                file_name="Laporan_Civil_Concrete.pdf",
                mime="application/pdf"
            )
        except Exception as e:
            st.error(f"Gagal menyiapkan file laporan: {e}")
    else:
        st.info("Belum ada hasil perhitungan untuk diekspor. Jalankan salah satu modul terlebih dahulu.")

st.divider()
st.caption("CIVIL CONCRETE | Versi akademik awal • Periksa kembali satuan, asumsi, dan ketentuan SNI yang berlaku.")
