import streamlit as st
import pandas as pd
import os
import sqlite3
import jdatetime
import io
import re
import qrcode
import plotly.express as px
import folium
from streamlit_folium import st_folium
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from fpdf import FPDF

st.set_page_config(page_title="مدیریت کارگاه علائم راهداری مازندران", layout="wide")

st.markdown("""
    <style>
        .stApp {
            direction: rtl;
            text-align: right;
        }
        label, .stTextInput, .stNumberInput, .stSelectbox, .stMultiSelect {
            direction: rtl;
            text-align: right;
        }
        h1, h2, h3, h4, h5, h6, .stMarkdown {
            direction: rtl;
            text-align: right;
        }
    </style>
""", unsafe_allow_html=True)

CDR_DIR = "uploaded_cdr_files"
QR_DIR = "qr_codes"
if not os.path.exists(CDR_DIR):
    os.makedirs(CDR_DIR)
if not os.path.exists(QR_DIR):
    os.makedirs(QR_DIR)

DB_FILE = "tablo_database.db"

def init_db():
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS signs_table (
            کد_ثبت_سامانه TEXT PRIMARY KEY,
            کد_تابلو TEXT,
            نوع_تابلو TEXT,
            نوع_تولید TEXT,
            تاریخ_ثبت_سفارش TEXT,
            سال_سفارش INTEGER,
            فرد_سفارش_دهنده TEXT,
            مسیر_فایل_کورل TEXT,
            تعداد_تابلو INTEGER,
            طول_سانتی_متر REAL,
            عرض_سانتی_متر REAL,
            مساحت_متر_مربع REAL,
            ارتفاع_اخطاری_انتظامی_سانتی REAL,
            تاریخ_تولید_تحویل_انبار TEXT,
            تاریخ_تحویل_شهرستان TEXT,
            شهرستان_تحویل_گیرنده TEXT,
            شرح_سایر_شهرستان TEXT,
            محور_و_مکان_نصب TEXT,
            کیلومتر_محور TEXT,
            عرض_جغرافیایی_lat REAL,
            طول_جغرافیایی_lon REAL,
            نصب_شده BOOLEAN
        )
    ''')
    
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS materials_table (
            کد_ماده TEXT PRIMARY KEY,
            نوع_مصالح TEXT,
            نوع_شبرنگ TEXT,
            برند_شبرنگ TEXT,
            رنگ TEXT,
            مقدار_مترمربع REAL,
            نوع_حرکت TEXT,
            محل_تامین TEXT,
            تاریخ TEXT,
            توضیحات TEXT
        )
    ''')
    
    try:
        cursor.execute("ALTER TABLE materials_table ADD COLUMN برند_شبرنگ TEXT")
    except:
        pass

    conn.commit()
    conn.close()

init_db()

def load_data():
    conn = sqlite3.connect(DB_FILE)
    try:
        df = pd.read_sql("SELECT * FROM signs_table", conn)
    except Exception:
        df = pd.DataFrame(columns=[
            "کد_ثبت_سامانه", "کد_تابلو", "نوع_تابلو", "نوع_تولید", "تاریخ_ثبت_سفارش", "سال_سفارش", "فرد_سفارش_دهنده",
            "مسیر_فایل_کورل", "تعداد_تابلو", "طول_سانتی_متر", "عرض_سانتی_متر", "مساحت_متر_مربع",
            "ارتفاع_اخطاری_انتظامی_سانتی", "تاریخ_تولید_تحویل_انبار", "تاریخ_تحویل_شهرستان",
            "شهرستان_تحویل_گیرنده", "شرح_سایر_شهرستان", "محور_و_مکان_نصب", "کیلومتر_محور", 
            "عرض_جغرافیایی_lat", "طول_جغرافیایی_lon", "نصب_شده"
        ])
    conn.close()
    return df

def save_data(df):
    conn = sqlite3.connect(DB_FILE)
    df.to_sql("signs_table", conn, if_exists="replace", index=False)
    conn.close()

def load_materials():
    conn = sqlite3.connect(DB_FILE)
    try:
        df_mat = pd.read_sql("SELECT * FROM materials_table", conn)
    except Exception:
        df_mat = pd.DataFrame(columns=[
            "کد_ماده", "نوع_مصالح", "نوع_شبرنگ", "برند_شبرنگ", "رنگ", "مقدار_مترمربع", 
            "نوع_حرکت", "محل_تامین", "تاریخ", "توضیحات"
        ])
    conn.close()
    return df_mat

def save_materials(df_mat):
    conn = sqlite3.connect(DB_FILE)
    df_mat.to_sql("materials_table", conn, if_exists="replace", index=False)
    conn.close()

def generate_qr_code(text, code_name):
    qr = qrcode.QRCode(version=1, box_size=10, border=2)
    qr.add_data(text)
    qr.make(fit=True)
    img = qr.make_image(fill_color="black", back_color="white")
    path = os.path.join(QR_DIR, f"QR_{code_name}.png")
    img.save(path)
    return path

def generate_kml(df_subset):
    kml_header = """<?xml version="1.0" encoding="UTF-8"?>
<kml xmlns="http://www.opengis.net/kml/2.2">
  <Document>
    <name>Mazandaran Traffic Signs GIS</name>
    <description>خروجی مختصات جغرافیایی تابلوهای راهداری استان مازندران</description>
"""
    kml_footer = """  </Document>
</kml>"""
    
    placemarks = ""
    for _, row in df_subset.iterrows():
        lat = row.get("عرض_جغرافیایی_lat", 36.5656)
        lon = row.get("طول_جغرافیایی_lon", 53.0586)
        sys_code = row.get("کد_ثبت_سامانه", "-")
        sign_code = row.get("کد_تابلو", "-")
        t_type = row.get("نوع_تابلو", "-")
        county = row.get("شهرستان_تحویل_گیرنده", "-")
        axis = row.get("محور_و_مکان_نصب", "-")
        km = row.get("کیلومتر_محور", "-")
        
        desc = f"کد سامانه: {sys_code} | کد تابلو: {sign_code} | نوع: {t_type} | شهرستان: {county} | محور: {axis} (کیلومتر {km})"
        
        placemarks += f"""    <Placemark>
      <name>{sys_code} ({sign_code})</name>
      <description>{desc}</description>
      <Point>
        <coordinates>{lon},{lat},0</coordinates>
      </Point>
    </Placemark>
"""
    return kml_header + placemarks + kml_footer

def generate_warehouse_delivery_word(row):
    doc = Document()
    doc.add_heading('حواله خروج از کارگاه و تحویل تجهیزات ایمنی', level=1)
    
    p = doc.add_paragraph()
    p.add_run(f"تاریخ صدور: {jdatetime.datetime.now().strftime('%Y/%m/%d')}\n").bold = True
    p.add_run(f"شماره حواله: WH-DEL-{row['کد_ثبت_سامانه']}\n")
    
    doc.add_heading('مشخصات اقلام تحویلی:', level=2)
    table = doc.add_table(rows=1, cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = 'شرح مشخصات'
    hdr_cells[1].text = 'مقدار / جزئیات'
    
    data_rows = [
        ("کد ثبت سامانه", str(row['کد_ثبت_سامانه'])),
        ("کد تابلو (مرجع)", str(row['کد_تابلو'])),
        ("نوع تابلو", str(row['نوع_تابلو'])),
        ("نوع تولید", str(row['نوع_تولید'])),
        ("تعداد", str(row['تعداد_تابلو'])),
        ("مساحت (مترمربع)", str(row['مساحت_متر_مربع'])),
        ("شهرستان تحویل‌گیرنده", str(row['شهرستان_تحویل_گیرنده']) if row['شهرستان_تحویل_گیرنده'] != 'سایر...' else str(row['شرح_سایر_شهرستان'])),
        ("محور و مکان نصب", str(row['محور_و_مکان_نصب'])),
        ("کیلومتر محور", str(row.get('کیلومتر_محور', '-'))),
        ("تاریخ تحویل به شهرستان", str(row['تاریخ_تحویل_شهرستان']))
    ]
    
    for title, val in data_rows:
        row_cells = table.add_row().cells
        row_cells[0].text = title
        row_cells[1].text = val
        
    doc.add_paragraph('\n\nنام و امضای تحویل‌دهنده (کارگاه): _______________        نام و امضای نماینده شهرستان: _______________')
    
    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio.getvalue()

def generate_labels_pdf(df_labels):
    pdf = FPDF(orientation='P', unit='mm', format='A4')
    pdf.add_page()
    pdf.set_font("helvetica", size=10)
    
    type_map = {"اطلاعاتی": "Information", "تودلی": "Todoli", "اخطاری": "Warning", "انتظامی": "Regulatory"}
    prod_map = {"جدید": "New", "مرمت": "Repair"}
    
    y_offset = 10
    for _, r in df_labels.iterrows():
        if y_offset > 240:
            pdf.add_page()
            y_offset = 10
            
        sys_c = str(r["کد_ثبت_سامانه"])
        sign_c = str(r["کد_تابلو"])
        t_type = type_map.get(str(r["نوع_تابلو"]), str(r["نوع_تابلو"]))
        p_type = prod_map.get(str(r["نوع_تولید"]), str(r["نوع_تولید"]))
        
        county = str(r["شهرستان_تحویل_گیرنده"])
        county_safe = county.encode('ascii', 'ignore').decode('ascii')
        if not county_safe: county_safe = "County"
        
        axis = str(r["محور_و_مکان_نصب"])
        axis_safe = axis.encode('ascii', 'ignore').decode('ascii')
        if not axis_safe: axis_safe = "Axis"
        
        date_ord = str(r["تاریخ_ثبت_سفارش"])
        
        pdf.rect(10, y_offset, 190, 35)
        
        pdf.set_xy(12, y_offset + 3)
        pdf.cell(186, 6, "Mazandaran Road Maintenance Dept - Label", 0, 1, 'L')
        
        pdf.set_xy(12, y_offset + 10)
        pdf.cell(186, 6, f"SysCode: {sys_c}  |  SignCode: {sign_c}", 0, 1, 'L')
        
        pdf.set_xy(12, y_offset + 17)
        pdf.cell(186, 6, f"Type: {t_type} ({p_type})  |  Destination: {county_safe}", 0, 1, 'L')
        
        pdf.set_xy(12, y_offset + 24)
        pdf.cell(186, 6, f"Axis: {axis_safe}  |  Date: {date_ord}", 0, 1, 'L')
        
        q_text = f"SysCode:{sys_c}|SignCode:{sign_c}|Type:{r['نوع_تابلو']}|County:{r['شهرستان_تحویل_گیرنده']}"
        q_path = generate_qr_code(q_text, str(sys_c).replace('/', '_'))
        if os.path.exists(q_path):
            pdf.image(q_path, x=168, y=y_offset + 2, w=30)
        
        y_offset += 40
        
    return bytes(pdf.output())

MAZANDARAN_COUNTIES = [
    "گلوگاه", "بهشهر", "نکا", "میاندرود", "ساری", "قائمشهر", 
    "جویبار", "بابل", "بابلسر", "سوادکوه", "نور", "آمل", 
    "لاریجان", "نوشهر", "چالوس", "عباس آباد", "تنکابن", 
    "رامسر", "محمودآباد", "فریدونکنار"
]
MAZANDARAN_COUNTIES_FULL = MAZANDARAN_COUNTIES + ["سایر..."]
ORDER_PEOPLE = ["رئیس شهرستان", "کارشناس ایمنی", "رئیس اداره ایمنی", "معاون راهداری", "مدیر کل"]

STD_COLORS = ["سفید", "نارنجی", "آبی", "قهوه ای", "سبز"]
PHOSPHOR_COLORS = ["سفید", "نارنجی", "فسفری", "آبی", "قهوه ای", "سبز"]
DAY_COLORS = ["مشکی", "سفید"]

SHABRANG_CLASSES = ["۳ ساله", "۷ ساله", "۱۰ ساله", "روز رنگ"]
SHABRANG_BRANDS = ["3M", "Avery Dennison", "Kiwalite", "Orafol / Oralite", "Avery", "سایر / متفرقه"]

if 'df' not in st.session_state:
    st.session_state.df = load_data()
if 'df_mat' not in st.session_state:
    st.session_state.df_mat = load_materials()

df = st.session_state.df
df_mat = st.session_state.df_mat
today_j_str = jdatetime.datetime.now().strftime("%Y%m%d")
current_j_year = jdatetime.datetime.now().year

st.title("🛡️ سامانه مدیریت کارگاه علائم راهداری مازندران")
st.markdown("---")

tabs = st.tabs([
    "فلش‌کارتها", 
    "ثبت تولید کارگاه", 
    "ویرایش داده", 
    "انبار مواد اولیه",
    "نقشه",
    "انبار تابلو"
])

with tabs[0]:
    st.subheader("📊 موجودی انبار")
    
    def calculate_current_balance(m_type, r_type=None, col=None):
        sub = df_mat.copy()
        if sub.empty:
            return 0.0
        sub = sub[sub["نوع_مصالح"] == m_type]
        if r_type:
            sub = sub[sub["نوع_شبرنگ"] == r_type]
        if col:
            sub = sub[sub["رنگ"] == col]
            
        in_val = sub[sub["نوع_حرکت"] == "ورود به انبار"]["مقدار_مترمربع"].sum()
        out_val = sub[sub["نوع_حرکت"] == "خروج از انبار"]["مقدار_مترمربع"].sum()
        return in_val - out_val

    st.markdown("##### 📦 شبرنگ")
    m_cols = st.columns(4)

    with m_cols[0]:
        t_3_val = calculate_current_balance("شبرنگ", "۳ ساله")
        lines_3 = []
        for c in STD_COLORS:
            val = calculate_current_balance('شبرنگ', '۳ ساله', c)
            c_style = "color:#fca5a5;" if val < 0 else ""
            val_str = f"{val:,.2f}"
            lines_3.append(f"<div style='font-size: 11px; margin-top: 2px;' dir='rtl'>▪️ {c}: <b style='{c_style}' dir='ltr' style='display:inline-block;'>{val_str}</b></div>")
        sub_html_3 = "".join(lines_3)
        color_style_3 = "#ef4444" if t_3_val < 0 else "#93c5fd"
        t_3_str = f"{t_3_val:,.2f}"
        
        st.markdown(f"""
            <div style="background: linear-gradient(135deg, #1f2937, #374151); padding: 12px; border-radius: 10px; color: white; border-right: 4px solid {color_style_3}; box-shadow: 0 2px 5px rgba(0,0,0,0.1);">
                <h6 style="margin:0 0 4px 0; font-size:13px; border-bottom:1px solid #4b5563; padding-bottom:3px;">🔹 شبرنگ ۳ ساله</h6>
                <p style="font-size:16px; font-weight:bold; margin:2px 0; color:{color_style_3};" dir="rtl">کل: <span dir="ltr" style="display:inline-block;">{t_3_str}</span> m²</p>
                {sub_html_3}
            </div>
        """, unsafe_allow_html=True)

    with m_cols[1]:
        t_7_val = calculate_current_balance("شبرنگ", "۷ ساله")
        lines_7 = []
        for c in PHOSPHOR_COLORS:
            val = calculate_current_balance('شبرنگ', '۷ ساله', c)
            c_style = "color:#fca5a5;" if val < 0 else ""
            val_str = f"{val:,.2f}"
            lines_7.append(f"<div style='font-size: 11px; margin-top: 2px;' dir='rtl'>▪️ {c}: <b style='{c_style}' dir='ltr' style='display:inline-block;'>{val_str}</b></div>")
        sub_html_7 = "".join(lines_7)
        color_style_7 = "#ef4444" if t_7_val < 0 else "#6ee7b7"
        t_7_str = f"{t_7_val:,.2f}"
        
        st.markdown(f"""
            <div style="background: linear-gradient(135deg, #1f2937, #374151); padding: 12px; border-radius: 10px; color: white; border-right: 4px solid {color_style_7}; box-shadow: 0 2px 5px rgba(0,0,0,0.1);">
                <h6 style="margin:0 0 4px 0; font-size:13px; border-bottom:1px solid #4b5563; padding-bottom:3px;">🔹 شبرنگ ۷ ساله</h6>
                <p style="font-size:16px; font-weight:bold; margin:2px 0; color:{color_style_7};" dir="rtl">کل: <span dir="ltr" style="display:inline-block;">{t_7_str}</span> m²</p>
                {sub_html_7}
            </div>
        """, unsafe_allow_html=True)

    with m_cols[2]:
        t_10_val = calculate_current_balance("شبرنگ", "۱۰ ساله")
        lines_10 = []
        for c in PHOSPHOR_COLORS:
            val = calculate_current_balance('شبرنگ', '۱۰ ساله', c)
            c_style = "color:#fca5a5;" if val < 0 else ""
            val_str = f"{val:,.2f}"
            lines_10.append(f"<div style='font-size: 11px; margin-top: 2px;' dir='rtl'>▪️ {c}: <b style='{c_style}' dir='ltr' style='display:inline-block;'>{val_str}</b></div>")
        sub_html_10 = "".join(lines_10)
        color_style_10 = "#ef4444" if t_10_val < 0 else "#fde047"
        t_10_str = f"{t_10_val:,.2f}"
        
        st.markdown(f"""
            <div style="background: linear-gradient(135deg, #1f2937, #374151); padding: 12px; border-radius: 10px; color: white; border-right: 4px solid {color_style_10}; box-shadow: 0 2px 5px rgba(0,0,0,0.1);">
                <h6 style="margin:0 0 4px 0; font-size:13px; border-bottom:1px solid #4b5563; padding-bottom:3px;">🔹 شبرنگ ۱۰ ساله</h6>
                <p style="font-size:16px; font-weight:bold; margin:2px 0; color:{color_style_10};" dir="rtl">کل: <span dir="ltr" style="display:inline-block;">{t_10_str}</span> m²</p>
                {sub_html_10}
            </div>
        """, unsafe_allow_html=True)

    with m_cols[3]:
        t_day_val = calculate_current_balance("شبرنگ", "روز رنگ")
        lines_day = []
        for c in DAY_COLORS:
            val = calculate_current_balance('شبرنگ', 'روز رنگ', c)
            c_style = "color:#fca5a5;" if val < 0 else ""
            val_str = f"{val:,.2f}"
            lines_day.append(f"<div style='font-size: 11px; margin-top: 2px;' dir='rtl'>▪️ {c}: <b style='{c_style}' dir='ltr' style='display:inline-block;'>{val_str}</b></div>")
        sub_html_day = "".join(lines_day)
        color_style_day = "#ef4444" if t_day_val < 0 else "#f472b6"
        t_day_str = f"{t_day_val:,.2f}"
        
        st.markdown(f"""
            <div style="background: linear-gradient(135deg, #1f2937, #374151); padding: 12px; border-radius: 10px; color: white; border-right: 4px solid {color_style_day}; box-shadow: 0 2px 5px rgba(0,0,0,0.1);">
                <h6 style="margin:0 0 4px 0; font-size:13px; border-bottom:1px solid #4b5563; padding-bottom:3px;">🔹 شبرنگ روز رنگ</h6>
                <p style="font-size:16px; font-weight:bold; margin:2px 0; color:{color_style_day};" dir="rtl">کل: <span dir="ltr" style="display:inline-block;">{t_day_str}</span> m²</p>
                {sub_html_day}
            </div>
        """, unsafe_allow_html=True)

    st.markdown("<div style='margin: 10px 0;'></div>", unsafe_allow_html=True)
    
    st.markdown("##### 🛠️ ورق (ریل)")
    r_cols = st.columns(2)
    with r_cols[0]:
        rail_new_bal = calculate_current_balance("ریل نو")
        bg_new = '#b91c1c' if rail_new_bal < 0 else '#2f855a'
        rail_new_str = f"{rail_new_bal:,.2f}"
        st.markdown(f"""
            <div style="background: {bg_new}; padding: 12px; border-radius: 10px; color: white; text-align: center; box-shadow: 0 2px 5px rgba(0,0,0,0.1);">
                <h6 style="margin:0 0 4px 0;">🟩 موجودی ورق/ریل‌های نو</h6>
                <p style="font-size: 18px; font-weight: bold; margin: 0;" dir="rtl"><span dir="ltr" style="display:inline-block;">{rail_new_str}</span> m²</p>
            </div>
        """, unsafe_allow_html=True)
    with r_cols[1]:
        rail_old_bal = calculate_current_balance("ریل فرسوده")
        bg_old = '#b91c1c' if rail_old_bal < 0 else '#c05621'
        rail_old_str = f"{rail_old_bal:,.2f}"
        st.markdown(f"""
            <div style="background: {bg_old}; padding: 12px; border-radius: 10px; color: white; text-align: center; box-shadow: 0 2px 5px rgba(0,0,0,0.1);">
                <h6 style="margin:0 0 4px 0;">🟧 موجودی ورق/ریل‌های فرسوده</h6>
                <p style="font-size: 18px; font-weight: bold; margin: 0;" dir="rtl"><span dir="ltr" style="display:inline-block;">{rail_old_str}</span> m²</p>
            </div>
        """, unsafe_allow_html=True)

    st.markdown("<hr style='margin: 15px 0;'>", unsafe_allow_html=True)

    dc1, dc2, dc3 = st.columns(3)
    with dc1:
        all_dash_years = sorted(df["سال_سفارش"].dropna().unique().tolist()) if not df.empty and "سال_سفارش" in df.columns else [current_j_year]
        selected_dash_year = st.selectbox("انتخاب سال ثبت سفارش:", options=["همه سال‌ها"] + all_dash_years)
    with dc2:
        selected_dash_county = st.selectbox("فیلتر شهرستان:", options=["همه شهرستان‌ها"] + MAZANDARAN_COUNTIES_FULL, key="dash_cnt")
    with dc3:
        dash_date_mode = st.selectbox("نوع فیلتر تاریخ:", options=["بدون فیلتر تاریخ", "بازه تاریخ شمسی"])

    filtered_dash = df.copy()
    if selected_dash_year != "همه سال‌ها":
        filtered_dash = filtered_dash[filtered_dash["سال_سفارش"] == selected_dash_year]
    if selected_dash_county != "همه شهرستان‌ها":
        filtered_dash = filtered_dash[filtered_dash["شهرستان_تحویل_گیرنده"] == selected_dash_county]
        
    if dash_date_mode == "بازه تاریخ شمسی":
        dc_s1, dc_s2 = st.columns(2)
        with dc_s1: d_start = st.text_input("از تاریخ (مثال 14050101):", value="")
        with dc_s2: d_end = st.text_input("تا تاریخ (مثال 14051229):", value="")
        if d_start and d_end:
            filtered_dash = filtered_dash[
                (filtered_dash["تاریخ_ثبت_سفارش"].astype(str).str.replace('.', '') >= d_start) &
                (filtered_dash["تاریخ_ثبت_سفارش"].astype(str).str.replace('.', '') <= d_end)
            ]

    st.markdown("##### 🌟 تولید و مرمت تابلو")

    info_df = filtered_dash[filtered_dash["نوع_تابلو"] == "اطلاعاتی"]
    info_new = info_df[info_df["نوع_تولید"] == "جدید"]["مساحت_متر_مربع"].sum() if not info_df.empty else 0.0
    info_repair = info_df[info_df["نوع_تولید"] == "مرمت"]["مساحت_متر_مربع"].sum() if not info_df.empty else 0.0
    info_total_sqm = info_new + info_repair
    info_delivered = info_df[info_df["نصب_شده"] == True]["مساحت_متر_مربع"].sum() if not info_df.empty else 0.0
    info_remaining = info_total_sqm - info_delivered

    todoli_df = filtered_dash[filtered_dash["نوع_تابلو"] == "تودلی"]
    todoli_total_sqm = todoli_df["مساحت_متر_مربع"].sum() if not todoli_df.empty else 0.0
    todoli_delivered = todoli_df[todoli_df["نصب_شده"] == True]["مساحت_متر_مربع"].sum() if not todoli_df.empty else 0.0
    todoli_remaining = todoli_total_sqm - todoli_delivered

    warning_df = filtered_dash[filtered_dash["نوع_تابلو"].isin(["اخطاری", "انتظامی"])]
    warn_count = int(warning_df["تعداد_تابلو"].sum()) if not warning_df.empty else 0
    warn_delivered = int(warning_df[warning_df["نصب_شده"] == True]["تعداد_تابلو"].sum()) if not warning_df.empty else 0
    warn_in_warehouse = warn_count - warn_delivered

    ccols = st.columns(3)
    with ccols[0]:
        st.markdown(f"""
            <div style="background: linear-gradient(135deg, #1e3a8a, #2563eb); padding: 18px; border-radius: 14px; color: white; box-shadow: 0 4px 10px rgba(0,0,0,0.15); margin-bottom: 10px; border-right: 6px solid #93c5fd;">
                <h4 style="margin: 0 0 6px 0; font-size: 16px;">🟦 تابلوهای اطلاعاتی</h4>
                <p style="margin: 2px 0; font-size: 22px; font-weight: bold;" dir="rtl">کل: <span dir="ltr" style="display:inline-block;">{info_total_sqm:,.2f}</span> <span style="font-size: 13px;">متر مربع</span></p>
                <div style="font-size: 12px; background: rgba(255,255,255,0.15); padding: 6px; border-radius: 6px; margin-top: 6px;">
                    🆕 جدید: <b dir="ltr" style="display:inline-block;">{info_new:,.2f}</b> m² | 🛠️ مرمت: <b dir="ltr" style="display:inline-block;">{info_repair:,.2f}</b> m²
                </div>
                <div style="font-size: 12px; margin-top: 6px; opacity: 0.95;">
                    🚚 تحویل‌شده (نصب): <b dir="ltr" style="display:inline-block;">{info_delivered:,.2f}</b> m²<br>
                    📦 مانده در انبار: <b dir="ltr" style="display:inline-block;">{info_remaining:,.2f}</b> m²
                </div>
            </div>
        """, unsafe_allow_html=True)

    with ccols[1]:
        st.markdown(f"""
            <div style="background: linear-gradient(135deg, #7c2d12, #ea580c); padding: 18px; border-radius: 14px; color: white; box-shadow: 0 4px 10px rgba(0,0,0,0.15); margin-bottom: 10px; border-right: 6px solid #ffedd5;">
                <h4 style="margin: 0 0 6px 0; font-size: 16px;">🟧 تابلوهای تودلی</h4>
                <p style="margin: 2px 0; font-size: 22px; font-weight: bold;" dir="rtl">کل: <span dir="ltr" style="display:inline-block;">{todoli_total_sqm:,.2f}</span> <span style="font-size: 13px;">متر مربع</span></p>
                <div style="font-size: 12px; margin-top: 28px; opacity: 0.95;">
                    🚚 تحویل‌شده (نصب): <b dir="ltr" style="display:inline-block;">{todoli_delivered:,.2f}</b> m²<br>
                    📦 مانده در انبار: <b dir="ltr" style="display:inline-block;">{todoli_remaining:,.2f}</b> m²
                </div>
            </div>
        """, unsafe_allow_html=True)

    with ccols[2]:
        st.markdown(f"""
            <div style="background: linear-gradient(135deg, #854d0e, #ca8a04); padding: 18px; border-radius: 14px; color: white; box-shadow: 0 4px 10px rgba(0,0,0,0.15); margin-bottom: 10px; border-right: 6px solid #fef08a;">
                <h4 style="margin: 0 0 6px 0; font-size: 16px;">⚠️ اخطاری و انتظامی</h4>
                <p style="margin: 2px 0; font-size: 22px; font-weight: bold;" dir="rtl">کل: <span dir="ltr" style="display:inline-block;">{warn_count:,}</span> <span style="font-size: 13px;">عدد</span></p>
                <div style="font-size: 12px; margin-top: 28px; opacity: 0.95;">
                    🚚 تحویل‌شده (نصب): <b dir="ltr" style="display:inline-block;">{warn_delivered:,}</b> عدد<br>
                    📦 مانده در انبار: <b dir="ltr" style="display:inline-block;">{warn_in_warehouse:,}</b> عدد
                </div>
            </div>
        """, unsafe_allow_html=True)

    st.markdown("---")
    st.markdown("##### 📈 نمودار ۳ میله‌ای مقایسه‌ای عملکرد تحویلی به شهرستان‌ها (شروع محور عمودی از صفر)")
    
    chart_rows = []
    for county in MAZANDARAN_COUNTIES:
        c_sub = filtered_dash[filtered_dash["شهرستان_تحویل_گیرنده"] == county] if not filtered_dash.empty else pd.DataFrame()
        info_sqm = c_sub[c_sub["نوع_تابلو"] == "اطلاعاتی"]["مساحت_متر_مربع"].sum() if not c_sub.empty else 0.0
        warn_cnt = int(c_sub[c_sub["نوع_تابلو"].isin(["اخطاری", "انتظامی"])]["تعداد_تابلو"].sum()) if not c_sub.empty else 0
        todoli_cnt = int(c_sub[c_sub["نوع_تابلو"] == "تودلی"]["تعداد_تابلو"].sum()) if not c_sub.empty else 0
        
        chart_rows.append({"شهرستان": county, "شاخص": "مساحت اطلاعاتی (m²)", "مقدار": float(info_sqm)})
        chart_rows.append({"شهرستان": county, "شاخص": "تعداد اخطاری/انتظامی (عدد)", "مقدار": float(warn_cnt)})
        chart_rows.append({"شهرستان": county, "شاخص": "تعداد تودلی (عدد)", "مقدار": float(todoli_cnt)})
        
    df_chart_long = pd.DataFrame(chart_rows)
    if not df_chart_long.empty:
        fig = px.bar(
            df_chart_long, 
            x="شهرستان", 
            y="مقدار", 
            color="شاخص", 
            barmode="group",
            color_discrete_map={
                "مساحت اطلاعاتی (m²)": "#2563eb",
                "تعداد اخطاری/انتظامی (عدد)": "#ea580c",
                "تعداد تودلی (عدد)": "#16a34a"
            }
        )
        fig.update_layout(
            xaxis_title="شهرستان",
            yaxis_title="مقدار تحویلی",
            legend_title="نوع شاخص",
            font=dict(family="Tahoma, Arial", size=12),
            margin=dict(l=20, r=20, t=20, b=20),
            yaxis=dict(rangemode="tozero")
        )
        st.plotly_chart(fig, use_container_width=True)

with tabs[1]:
    st.subheader("➕ ثبت تولید")
    
    f1, f2, f3, f4 = st.columns(4)
    with f1:
        sign_code = st.text_input("کد تابلو (اختیاری - مرجع طرح، مثلا W-1):", value="")
    with f2:
        tablo_type = st.selectbox("نوع تابلو:", ["اطلاعاتی", "تودلی", "اخطاری", "انتظامی"])
    with f3:
        prod_type = st.selectbox("نوع تولید:", ["جدید", "مرمت"])
    with f4:
        order_date = st.text_input("تاریخ ثبت سفارش (مثال 14050215):", value=today_j_str)

    if prod_type == "مرمت":
        st.warning("⚠️ **هشدار سیستم:** برای تابلوهای نوع **«مرمت»** (اطلاعاتی، تودلی، اخطاری و انتظامی)، ثبت مساحت و کسر ورق لحاظ نمی‌شود.")

    with st.form("tablo_form"):
        f5, f6 = st.columns(2)
        with f5:
            order_person = st.selectbox("فرد سفارش‌دهنده:", ORDER_PEOPLE)
        with f6:
            tablo_count = st.number_input("تعداد تابلو:", min_value=1, value=1)

        width_cm = 0.0
        height_cm = 0.0
        area_sqm = 0.0

        if tablo_type == "اطلاعاتی":
            col_b1, col_b2 = st.columns(2)
            with col_b1:
                width_cm = st.number_input("طول (سانتی‌متر):", min_value=0.0, value=100.0)
            with col_b2:
                height_cm = st.number_input("عرض (سانتی‌متر):", min_value=0.0, value=60.0)
            area_sqm = round((width_cm * height_cm) / 10000.0, 2)

        rail_material_choice = "--- بدون کسر ورق ---"
        if tablo_type == "اطلاعاتی" and prod_type == "جدید":
            rail_material_choice = st.selectbox("انتخاب ریل ورق:", ["ریل نو", "ریل فرسوده"])

        shabrang_inputs = []
        for i in range(1, 7):
            s_col1, s_col2, s_col3, s_col4, s_col5 = st.columns([1, 2, 2, 2, 2])
            with s_col1:
                st.markdown(f"**شبرنگ {i}:**")
            with s_col2:
                s_cls = st.selectbox(f"نوع شبرنگ {i}:", ["--- غیرفعال ---"] + SHABRANG_CLASSES, key=f"s_cls_{i}")
            with s_col3:
                s_brd = st.selectbox(f"برند {i}:", SHABRANG_BRANDS, key=f"s_brd_{i}")
            with s_col4:
                s_clr = st.selectbox(f"رنگ {i}:", STD_COLORS, key=f"s_clr_{i}")
            with s_col5:
                s_sqm = st.number_input(f"مساحت (m²) {i}:", min_value=0.0, value=0.0, step=0.1, key=f"s_sqm_{i}")
            
            if s_cls != "--- غیرفعال ---" and s_sqm > 0:
                shabrang_inputs.append({"class": s_cls, "brand": s_brd, "color": s_clr, "sqm": s_sqm})

        g1, g2, g3 = st.columns(3)
        with g1:
            install_location = st.text_input("محور و مکان پیشنهادی نصب:")
        with g2:
            km_location = st.text_input("کیلومتر در محور:", value="-")
        with g3:
            lat_val = st.number_input("عرض جغرافیایی (Lat) [اختیاری]:", value=0.0, format="%.6f")
            lon_val = st.number_input("طول جغرافیایی (Lon) [اختیاری]:", value=0.0, format="%.6f")

        f7, f8, f9 = st.columns(3)
        with f7:
            prod_delivery_date = st.text_input("تاریخ تولید و تحویل به انبار:", value=today_j_str)
        with f8:
            county_delivery_date = st.text_input("تاریخ تحویل به شهرستان:", value=today_j_str)
        with f9:
            uploaded_cdr = st.file_uploader("آپلود فایل طراحی کورل (.cdr یا .zip):", type=["cdr", "zip"])

        receiver_county = st.selectbox("شهرستان تحویل‌گیرنده:", MAZANDARAN_COUNTIES_FULL)
        other_county_desc = st.text_input("نام سازمان یا دهیاری تحویل‌گیرنده (در صورت سایر):") if receiver_county == "سایر..." else "-"
        is_installed = st.checkbox("نصب شده است؟", value=False)

        submitted_order = st.form_submit_button("💾 ثبت سفارش و کسر مصالح از انبار")
        if submitted_order:
            clean_sign_code = sign_code.strip() if sign_code.strip() else "-"
            saved_cdr_path = "-"
            base_time_id = int(pd.Timestamp.now().timestamp())
            
            if uploaded_cdr is not None:
                file_ext = uploaded_cdr.name.split('.')[-1]
                saved_cdr_path = os.path.join(CDR_DIR, f"{clean_sign_code}_{base_time_id}.{file_ext}")
                with open(saved_cdr_path, "wb") as f:
                    f.write(uploaded_cdr.getbuffer())

            new_rows = []
            mat_out_rows = []
            for i in range(1, int(tablo_count) + 1):
                sys_code = f"SYS-{base_time_id}-{i}" if tablo_count > 1 else f"SYS-{base_time_id}"
                new_rows.append({
                    "کد_ثبت_سامانه": sys_code, "کد_تابلو": clean_sign_code, "نوع_تابلو": tablo_type,
                    "نوع_تولید": prod_type, "تاریخ_ثبت_سفارش": order_date, "سال_سفارش": int(order_date[:4]) if len(order_date)>=4 else current_j_year,
                    "فرد_سفارش_دهنده": order_person, "مسیر_فایل_کورل": saved_cdr_path, "تعداد_تابلو": 1,
                    "طول_سانتی_متر": width_cm, "عرض_سانتی_متر": height_cm, "مساحت_متر_مربع": area_sqm,
                    "ارتفاع_اخطاری_انتظامی_سانتی": 0.0, "تاریخ_تولید_تحویل_انبار": prod_delivery_date,
                    "تاریخ_تحویل_شهرستان": county_delivery_date, "شهرستان_تحویل_گیرنده": receiver_county,
                    "شرح_سایر_شهرستان": other_county_desc, "محور_و_مکان_نصب": install_location or "-",
                    "کیلومتر_محور": km_location, "عرض_جغرافیایی_lat": float(lat_val), "طول_جغرافیایی_lon": float(lon_val),
                    "نصب_شده": is_installed
                })

                if tablo_type == "اطلاعاتی" and prod_type == "جدید" and rail_material_choice != "--- بدون کسر ورق ---":
                    mat_out_rows.append({
                        "کد_ماده": f"MAT-OUT-RAIL-{base_time_id}-{i}", "نوع_مصالح": rail_material_choice,
                        "نوع_شبرنگ": "-", "برند_شبرنگ": "-", "رنگ": "-", "مقدار_مترمربع": float(area_sqm),
                        "نوع_حرکت": "خروج از انبار", "محل_تامین": f"ساخت تابلو {clean_sign_code}", "تاریخ": order_date, "توضیحات": sys_code
                    })

                for idx, sh_item in enumerate(shabrang_inputs):
                    mat_out_rows.append({
                        "کد_ماده": f"MAT-OUT-{base_time_id}-{i}-{idx+1}", "نوع_مصالح": "شبرنگ",
                        "نوع_شبرنگ": sh_item["class"], "برند_شبرنگ": sh_item["brand"], "رنگ": sh_item["color"],
                        "مقدار_مترمربع": float(sh_item["sqm"]), "نوع_حرکت": "خروج از انبار", "محل_تامین": f"ساخت تابلو {clean_sign_code}",
                        "تاریخ": order_date, "توضیحات": sys_code
                    })

            df = pd.concat([df, pd.DataFrame(new_rows)], ignore_index=True)
            st.session_state.df = df
            save_data(df)

            if mat_out_rows:
                df_mat = pd.concat([df_mat, pd.DataFrame(mat_out_rows)], ignore_index=True)
                st.session_state.df_mat = df_mat
                save_materials(df_mat)

            st.success("✅ سفارش با موفقیت ثبت شد و متریال از انبار کسر گردید!")

with tabs[2]:
    st.subheader("✏️ ویرایش داده‌ها و حذف رکورد")
    if not df.empty:
        srch_c1, srch_c2 = st.columns(2)
        with srch_c1:
            query_sys = st.text_input("جستجو بر اساس کد سامانه یا کد تابلو:", value="", key="search_edit_sys")
        with srch_c2:
            query_cnt = st.selectbox("فیلتر شهرستان:", options=["همه شهرستان‌ها"] + MAZANDARAN_COUNTIES_FULL, key="filter_edit_cnt")

        filtered_mgmt = df.copy()
        if query_sys.strip():
            filtered_mgmt = filtered_mgmt[filtered_mgmt["کد_ثبت_سامانه"].astype(str).str.contains(query_sys.strip(), na=False) | filtered_mgmt["کد_تابلو"].astype(str).str.contains(query_sys.strip(), na=False)]
        if query_cnt != "همه شهرستان‌ها":
            filtered_mgmt = filtered_mgmt[filtered_mgmt["شهرستان_تحویل_گیرنده"] == query_cnt]

        if not filtered_mgmt.empty:
            st.markdown("##### 🗑️ بخش حذف رکورد:")
            del_query = st.text_input("جستجوی کد سامانه جهت حذف:", value="", key="del_search_box")
            if del_query.strip():
                del_matched = filtered_mgmt[filtered_mgmt["کد_ثبت_سامانه"].astype(str).str.contains(del_query.strip(), na=False)]
                if not del_matched.empty:
                    st.dataframe(del_matched, use_container_width=True)
                    del_target_code = st.selectbox("تایید کد ثبت سامانه جهت حذف قطعی:", options=del_matched["کد_ثبت_سامانه"].tolist(), key="del_select_box")
                    if st.button("🗑️ حذف رکورد انتخابی", type="primary", key="del_exec_btn"):
                        df = df[df["کد_ثبت_سامانه"] != del_target_code]
                        st.session_state.df = df
                        save_data(df)
                        st.success("🗑️ رکورد مورد نظر با موفقیت حذف شد!")
                        st.rerun()
                else:
                    st.warning("رکوردی با این کد سامانه یافت نشد.")

            st.markdown("---")
            st.markdown("##### 📋 جدول ویرایش اطلاعات (مسیر فایل کورل غیرقابل ویرایش است):")
            
            # غیرفعال کردن ویرایش ستون‌های کلیدی و مسیر فایل کورل
            edited_df = st.data_editor(filtered_mgmt, use_container_width=True, disabled=["کد_ثبت_سامانه", "مسیر_فایل_کورل"])
            
            if st.button("💾 ذخیره تغییرات جدول", key="save_edit_table_btn"):
                for col in ["سال_سفارش", "تعداد_تابلو"]:
                    if col in edited_df.columns:
                        edited_df[col] = pd.to_numeric(edited_df[col], errors='coerce').fillna(0).astype(int)
                for col in ["طول_سانتی_متر", "عرض_سانتی_متر", "مساحت_متر_مربع", "ارتفاع_اخطاری_انتظامی_سانتی", "عرض_جغرافیایی_lat", "طول_جغرافیایی_lon"]:
                    if col in edited_df.columns:
                        edited_df[col] = pd.to_numeric(edited_df[col], errors='coerce').fillna(0.0)

                for _, row in edited_df.iterrows():
                    sys_id = row["کد_ثبت_سامانه"]
                    for col in df.columns:
                        if col in row and col != "مسیر_فایل_کورل":
                            df.loc[df["کد_ثبت_سامانه"] == sys_id, col] = row[col]
                            
                st.session_state.df = df
                save_data(df)
                st.success("✅ تغییرات با موفقیت ذخیره شد!")
                st.rerun()
        else:
            st.warning("رکوردی با این مشخصات یافت نشد.")
    else:
        st.info("رکوردی برای ویرایش یا حذف وجود ندارد.")

with tabs[3]:
    st.subheader("📦 انبار مواد اولیه و مصالح")
    with st.form("materials_form"):
        m1, m2, m3, m4 = st.columns(4)
        with m1: 
            mat_type = st.selectbox("نوع مصالح:", ["شبرنگ", "ریل نو", "ریل فرسوده"])
        with m2: 
            movement_type = st.selectbox("نوع تراکنش:", ["ورود به انبار (خرید)", "خروج از انبار (مصرف)"])
        with m3: 
            mat_date = st.text_input("تاریخ:", value=today_j_str)
        with m4: 
            mat_amount = st.number_input("مقدار (m²):", min_value=0.0, value=10.0)

        is_shabrang = (mat_type == "شبرنگ")

        m5, m6, m7 = st.columns(3)
        with m5:
            mat_shabrang_class = st.selectbox("نوع شبرنگ:", SHABRANG_CLASSES, disabled=not is_shabrang) if is_shabrang else "-"
        with m6:
            mat_brand = st.selectbox("برند شبرنگ:", SHABRANG_BRANDS, disabled=not is_shabrang) if is_shabrang else "-"
        with m7:
            mat_color = st.selectbox("رنگ شبرنگ:", STD_COLORS, disabled=not is_shabrang) if is_shabrang else "-"

        m8, m9 = st.columns(2)
        with m8:
            party_name = st.text_input("تامین‌کننده (در ورود) / تحویل‌گیرنده (در خروج):", value="-")
        with m9:
            mat_desc = st.text_input("توضیحات:", value="-")

        submitted_mat = st.form_submit_button("💾 ثبت تراکنش مصالح")
        if submitted_mat:
            is_entry = "ورود" in movement_type
            new_mat_row = {
                "کد_ماده": f"MAT-{int(pd.Timestamp.now().timestamp())}", 
                "نوع_مصالح": mat_type,
                "نوع_شبرنگ": mat_shabrang_class if is_shabrang else "-", 
                "برند_شبرنگ": mat_brand if is_shabrang else "-",
                "رنگ": mat_color if is_shabrang else "-", 
                "مقدار_مترمربع": float(mat_amount),
                "نوع_حرکت": "ورود به انبار" if is_entry else "خروج از انبار",
                "محل_تامین": party_name, 
                "تاریخ": mat_date, 
                "توضیحات": mat_desc
            }
            df_mat = pd.concat([df_mat, pd.DataFrame([new_mat_row])], ignore_index=True)
            st.session_state.df_mat = df_mat
            save_materials(df_mat)
            st.success("✅ تراکنش مصالح با موفقیت ثبت شد!")

    st.markdown("---")
    st.subheader("🔍 جستجو و حذف تراکنش‌های انبار مصالح")
    if not df_mat.empty:
        mat_search = st.text_input("جستجو بر اساس کد ماده یا تامین‌کننده/تحویل‌گیرنده:", value="", key="mat_search_input")
        df_mat_filtered = df_mat.copy()
        if mat_search.strip():
            df_mat_filtered = df_mat_filtered[
                df_mat_filtered["کد_ماده"].astype(str).str.contains(mat_search.strip(), na=False) |
                df_mat_filtered["محل_تامین"].astype(str).str.contains(mat_search.strip(), na=False)
            ]

        if not df_mat_filtered.empty and mat_search.strip():
            st.dataframe(df_mat_filtered, use_container_width=True)
            del_mat_code = st.selectbox("تایید کد ماده جهت حذف قطعی:", options=df_mat_filtered["کد_ماده"].tolist(), key="del_mat_select")
            if st.button("🗑️ حذف تراکنش مصالح", type="primary", key="del_mat_btn"):
                df_mat = df_mat[df_mat["کد_ماده"] != del_mat_code]
                st.session_state.df_mat = df_mat
                save_materials(df_mat)
                st.success("🗑️ تراکنش مصالح با موفقیت حذف شد!")
                st.rerun()
        else:
            st.dataframe(df_mat, use_container_width=True)
    else:
        st.info("رکوردی در انبار مصالح ثبت نشده است.")

with tabs[4]:
    st.subheader("🗺️ نقشه جغرافیایی")
    if not df.empty:
        mf1, mf2, mf3, mf4 = st.columns(4)
        with mf1:
            all_map_years = sorted(df["سال_سفارش"].dropna().unique().tolist()) if "سال_سفارش" in df.columns else [current_j_year]
            map_sel_year = st.selectbox("فیلتر سال سفارش:", options=["همه سال‌ها"] + all_map_years, key="map_yr")
        with mf2:
            map_sel_county = st.selectbox("فیلتر شهرستان:", options=["همه شهرستان‌ها"] + MAZANDARAN_COUNTIES_FULL, key="map_cnt")
        with mf3:
            map_sel_type = st.selectbox("فیلتر نوع تابلو:", options=["همه انواع تابلوها", "اطلاعاتی", "اخطاری", "انتظامی", "تودلی"], key="map_tp")
        with mf4:
            map_status_filter = st.radio("وضعیت نصب:", ["همه تابلوها", "فقط نصب‌شده", "فقط موجود در انبار"], horizontal=True, key="map_st")

        map_filtered_df = df.copy()
        if map_sel_year != "همه سال‌ها":
            map_filtered_df = map_filtered_df[map_filtered_df["سال_سفارش"] == map_sel_year]
        if map_sel_county != "همه شهرستان‌ها":
            map_filtered_df = map_filtered_df[map_filtered_df["شهرستان_تحویل_گیرنده"] == map_sel_county]
        if map_sel_type != "همه انواع تابلوها":
            map_filtered_df = map_filtered_df[map_filtered_df["نوع_تابلو"] == map_sel_type]
            
        if map_status_filter == "فقط نصب‌شده":
            map_filtered_df = map_filtered_df[map_filtered_df["نصب_شده"] == True]
        elif map_status_filter == "فقط موجود در انبار":
            map_filtered_df = map_filtered_df[map_filtered_df["نصب_شده"] == False]

        st.markdown("---")
        if not map_filtered_df.empty:
            map_plot_df = map_filtered_df.copy()
            map_plot_df["lat"] = pd.to_numeric(map_plot_df["عرض_جغرافیایی_lat"], errors="coerce").fillna(36.5656)
            map_plot_df["lon"] = pd.to_numeric(map_plot_df["طول_جغرافیایی_lon"], errors="coerce").fillna(53.0586)

            avg_lat = float(map_plot_df["lat"].mean())
            avg_lon = float(map_plot_df["lon"].mean())

            m = folium.Map(location=[avg_lat, avg_lon], zoom_start=9, tiles="OpenStreetMap")

            for _, row in map_plot_df.iterrows():
                circle_color = "green" if row["نصب_شده"] else "red"
                sign_code_lbl = str(row["کد_تابلو"])
                sys_code_lbl = str(row["کد_ثبت_سامانه"])
                
                popup_html = f"<b>کد سامانه:</b> {sys_code_lbl}<br><b>کد تابلو:</b> {sign_code_lbl}<br><b>نوع:</b> {row['نوع_تابلو']}<br><b>شهرستان:</b> {row['شهرستان_تحویل_گیرنده']}"
                
                folium.CircleMarker(
                    location=[row["lat"], row["lon"]],
                    radius=8,
                    color=circle_color,
                    fill=True,
                    fill_color=circle_color,
                    fill_opacity=0.8,
                    popup=folium.Popup(popup_html, max_width=200)
                ).add_to(m)

                folium.map.Marker(
                    [row["lat"], row["lon"]],
                    icon=folium.DivIcon(
                        icon_size=(150,36),
                        icon_anchor=(0,0),
                        html=f'<div style="font-size: 11pt; color: black; font-weight: bold; background-color: rgba(255,255,255,0.7); padding: 2px 4px; border-radius: 3px; display: inline-block;">{sign_code_lbl}</div>'
                    )
                ).add_to(m)

            st_folium(m, width="100%", height=500)

            kml_content = generate_kml(map_filtered_df)
            st.download_button("📥 دانلود فایل KML فیلترشده برای Google Earth", data=kml_content, file_name=f"Mazandaran_Signs_{today_j_str}.kml", mime="application/vnd.google-earth.kml+xml")
        else:
            st.warning("⚠️ هیچ رکوردی با فیلترهای انتخاب‌شده روی نقشه پیدا نشد.")
    else:
        st.info("رکوردی برای نمایش روی نقشه وجود ندارد.")

with tabs[5]:
    st.subheader("📦 انبار تابلوها، جستجوی تک‌کد و دانلود PDF تکی یا تجمیعی")
    if not df.empty:
        def parse_date_parts(date_str):
            clean = re.sub(r'[^\d]', '', str(date_str).strip())
            return (clean[:4] if len(clean)>=4 else "-", clean[4:6] if len(clean)>=6 else "-", clean[6:8] if len(clean)>=8 else "-")

        df_inv = df.copy()
        df_inv["_سال"] = df_inv["تاریخ_تولید_تحویل_انبار"].apply(lambda x: parse_date_parts(x)[0])
        df_inv["_ماه"] = df_inv["تاریخ_تولید_تحویل_انبار"].apply(lambda x: parse_date_parts(x)[1])
        df_inv["_روز"] = df_inv["تاریخ_تولید_تحویل_انبار"].apply(lambda x: parse_date_parts(x)[2])

        c1, c2, c3, c4, c5 = st.columns(5)
        with c1: f_cnt = st.selectbox("شهرستان:", options=["همه شهرستان‌ها"] + MAZANDARAN_COUNTIES_FULL)
        with c2: f_yr = st.selectbox("سال تولید:", options=["همه سال‌ها"] + sorted(df_inv["_سال"].unique().tolist()))
        with c3: f_mo = st.selectbox("ماه تولید:", options=["همه ماه‌ها"] + sorted(df_inv["_ماه"].unique().tolist()))
        with c4: f_dy = st.selectbox("روز تولید:", options=["همه روزها"] + sorted(df_inv["_روز"].unique().tolist()))
        with c5: f_txt = st.text_input("جستجوی تک‌کد سامانه/تابلو:")

        filtered_inv = df_inv.copy()
        if f_cnt != "همه شهرستان‌ها": filtered_inv = filtered_inv[filtered_inv["شهرستان_تحویل_گیرنده"] == f_cnt]
        if f_yr != "همه سال‌ها": filtered_inv = filtered_inv[filtered_inv["_سال"] == f_yr]
        if f_mo != "همه ماه‌ها": filtered_inv = filtered_inv[filtered_inv["_ماه"] == f_mo]
        if f_dy != "همه روزها": filtered_inv = filtered_inv[filtered_inv["_روز"] == f_dy]
        if f_txt.strip(): filtered_inv = filtered_inv[filtered_inv["کد_ثبت_سامانه"].str.contains(f_txt, na=False) | filtered_inv["کد_تابلو"].str.contains(f_txt, na=False)]

        st.markdown("---")
        if not filtered_inv.empty:
            sel_sys = st.selectbox("انتخاب تابلوی مد نظر برای دانلود تکی:", options=filtered_inv["کد_ثبت_سامانه"].tolist())
            row_single = df[df["کد_ثبت_سامانه"] == sel_sys].iloc[0]

            col_s1, col_s2 = st.columns(2)
            with col_s1:
                pdf_single = generate_labels_pdf(pd.DataFrame([row_single]))
                st.download_button("📥 دانلود لیبل PDF (تکی)", data=pdf_single, file_name=f"Label_{sel_sys}.pdf", mime="application/pdf", use_container_width=True)
            with col_s2:
                word_single = generate_warehouse_delivery_word(row_single)
                st.download_button("📥 دانلود حواله Word (تکی)", data=word_single, file_name=f"Delivery_{sel_sys}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)

            st.markdown("---")
            total_filtered_count = len(filtered_inv)
            pdf_multi = generate_labels_pdf(filtered_inv)
            st.download_button(
                label=f"📥 دانلود PDF تجمیعی لیبل‌ها (شامل تمامی {total_filtered_count} مورد فیلترشده)",
                data=pdf_multi,
                file_name=f"Aggregated_Labels_{today_j_str}.pdf",
                mime="application/pdf",
                use_container_width=True
            )
        else:
            st.warning("موردی با این مشخصات یافت نشد.")
    else:
        st.info("رکوردی وجود ندارد.")

