import streamlit as st
import openpyxl
from datetime import datetime
import os
import re
import pdfplumber

st.set_page_config(
    page_title="أداة تفريغ العهود - 2026",
    page_icon="🏛️",
    layout="centered"
)

st.title("🏛️ أداة تفريغ العهود وإصدار أمر الصرف")
st.caption("الجمعية الخيرية لتحفيظ القرآن ببريدة - نموذج الإدارة العامة 2026")
st.divider()

API_KEY = st.secrets.get("GEMINI_API_KEY", "")

def convert_arabic_digits(text):
    """تحويل كافة الأرقام العربية الشرقية (٠١٢٣٤٥٦٧٨٩) إلى أرقام إنجليزية (0123456789)"""
    if not text:
        return ""
    arabic_digits = "٠١٢٣٤٥٦٧٨٩"
    english_digits = "0123456789"
    translation_table = str.maketrans(arabic_digits, english_digits)
    return str(text).translate(translation_table)

def tafqeet_arabic(amount):
    """تحويل المبلغ تلقائياً إلى كلمات عربية (تفقيط)"""
    try:
        if amount <= 0:
            return ""
        riyal = int(amount)
        halalas = int(round((amount - riyal) * 100))
        text = f"{riyal} ريال سعودي"
        if halalas > 0:
            text += f" و {halalas} هللة"
        return text + " فقط لا غير"
    except Exception:
        return ""

def get_arabic_month(month_num):
    months = {
        1: "يناير", 2: "فبراير", 3: "مارس", 4: "أبريل",
        5: "مايو", 6: "يونيو", 7: "يوليو", 8: "أغسطس",
        9: "سبتمبر", 10: "أكتوبر", 11: "نوفمبر", 12: "ديسمبر"
    }
    return months.get(month_num, "")

def parse_pdf_invoice_dynamic(uploaded_file):
    """استخراج ديناميكي مرن لرقم الفاتورة المتغير ورقم الحساب والتاريخ والمبلغ"""
    data = {
        "inv_no": "",
        "account_no": "",
        "inv_date": "",
        "amount": 0.0,
        "tax_no": "300000361310003",
        "supplier": "الشركة السعودية للكهرباء"
    }
    try:
        uploaded_file.seek(0)
        with pdfplumber.open(uploaded_file) as pdf:
            full_text_raw = ""
            for page in pdf.pages:
                text = page.extract_text() or ""
                full_text_raw += text + "\n"
        
        # تحويل النص كاملاً لأرقام إنجليزية لسهولة المطابقة
        full_text = convert_arabic_digits(full_text_raw)
        
        # 1. استخراج رقم الفاتورة الديناميكي المتغير
        inv_match = (
            re.search(r'رقم\s*الفاتورة\s*:\s*(\d+)', full_text) or
            re.search(r'رقم\s*الفاتورة[^\d]*(\d{10,16})', full_text) or
            re.search(r'الفاتورة\s*:\s*(\d{10,16})', full_text)
        )
        if inv_match:
            data["inv_no"] = inv_match.group(1)
            
        # 2. استخراج رقم الحساب المتغير
        acc_match = (
            re.search(r'رقم\s*الحساب[^\d]*(\d{10,12})', full_text) or
            re.search(r'الحساب[^\d]*(\d{10,12})', full_text) or
            re.search(r'(\d{11})', full_text)
        )
        if acc_match:
            data["account_no"] = acc_match.group(1)
            
        # 3. استخراج تاريخ الفاتورة / بداية الاستحقاق
        date_match = re.search(r'(\d{4}[-/]\d{2}[-/]\d{2})', full_text) or re.search(r'(\d{2}[-/]\d{2}[-/]\d{4})', full_text)
        if date_match:
            d_str = date_match.group(1).replace('/', '-')
            parts = d_str.split('-')
            if len(parts[0]) == 4:
                data["inv_date"] = d_str
            else:
                data["inv_date"] = f"{parts[2]}-{parts[1]}-{parts[0]}"
                
        # 4. استخراج الرقم الضريبي للمورد
        tax_match = re.search(r'(3\d{14})', full_text)
        if tax_match:
            data["tax_no"] = tax_match.group(1)

        # 5. استخراج المبلغ غير شامل الضريبة ديناميكياً
        amt_match = re.search(r'غير\s*شامل[^\d]*(\d+[\.,]\d{2})', full_text) or re.search(r'الخاضع\s*للضريبة[^\d]*(\d+[\.,]\d{2})', full_text)
        if amt_match:
            raw_amt = amt_match.group(1).replace(',', '.')
            try:
                data["amount"] = float(raw_amt)
            except Exception:
                pass
        else:
            # البحث عن جميع أرقام المبالغ العشرية في المستند واختيار المبلغ الصحيح
            amounts = re.findall(r'(\d+\.\d{2})', full_text)
            possible_amounts = []
            for a in amounts:
                val = float(a)
                if 10.0 <= val <= 50000.0 and val not in [1765.0, 2180.13]:
                    possible_amounts.append(val)
            if possible_amounts:
                data["amount"] = possible_amounts[0]

    except Exception:
        pass
    return data

st.subheader("1️⃣ رفع الفواتير (حتى 5 فواتير)")
uploaded_files = st.file_uploader(
    "التقط صور الفواتير أو اسحب ملفات الـ PDF هنا (حتى 5 فواتير معاً)", 
    type=["pdf", "png", "jpg", "jpeg"],
    accept_multiple_files=True
)

if uploaded_files:
    if len(uploaded_files) > 5:
        st.warning("⚠️ تم تحديد أكثر من 5 فواتير. سيتم معالجة أول 5 فواتير فقط.")
        uploaded_files = uploaded_files[:5]
        
    st.success(f"✅ تم استلام {len(uploaded_files)} فاتورة بنجاح!")
    st.divider()
    
    st.subheader("2️⃣ مراجعة وتحديد البيانات")
    
    dar_list = [
        "إدارة الدور النسائية",
        "دار حصة",
        "دار التقوى",
        "دار الصفا",
        "دار لمياء بنت الخياط",
        "دار منيرة المانع",
        "دار الفتاة",
        "دار الفايزية",
        "دار الهدى",
        "دار الضياء",
        "دار زينب",
        "دار العليا",
        "دار الخنساء",
        "التاليات",
        "خولة بنت حكيم",
        "غراس",
        "سيدة بنيجات",
        "الأفق",
        "الوقار",
        "دار الطرفية"
    ]
    
    dar_name = st.selectbox("اختر اسم الدار", dar_list)
    
    invoices_data = []
    
    for idx, file in enumerate(uploaded_files):
        st.markdown(f"### 📄 الفاتورة رقم ({idx + 1}): `{file.name}`")
        
        with st.spinner(f"🔍 جاري استخراج البيانات تلقائياً للفاتورة #{idx + 1}..."):
            ext_data = parse_pdf_invoice_dynamic(file)
            
        col1, col2 = st.columns(2)
        with col1:
            supplier = st.text_input(f"اسم المورد #{idx+1}", value=ext_data.get("supplier", "الشركة السعودية للكهرباء"), key=f"sup_{idx}")
            inv_no = st.text_input(f"رقم الفاتورة #{idx+1}", value=str(ext_data.get("inv_no", "")), key=f"inv_{idx}")
            account_no = st.text_input(f"رقم الحساب #{idx+1}", value=str(ext_data.get("account_no", "")), key=f"acc_{idx}")
            inv_date = st.text_input(f"تاريخ الفاتورة #{idx+1}", value=str(ext_data.get("inv_date", "")), key=f"date_{idx}")
            
        with col2:
            tax_no = st.text_input(f"الرقم الضريبي #{idx+1}", value=str(ext_data.get("tax_no", "300000361310003")), key=f"tax_{idx}")
            amount = st.number_input(f"المبلغ بدون الضريبة #{idx+1}", value=float(ext_data.get("amount", 0.0) or 0.0), step=0.01, key=f"amt_{idx}")
            statement = st.text_input(f"البيان #{idx+1}", value=f"سداد فاتورة كهرباء - {dar_name}", key=f"st_{idx}")
            category = st.selectbox(f"بند الحساب #{idx+1}", [
                "الكهرباء", 
                "المياة", 
                "الهاتف والبريد",
                "الصيانة",
                "قرطاسيات ومطبوعات",
                "نظافة",
                "أخرى"
            ], key=f"cat_{idx}")
            
            estimated_total = amount * 1.15
            written_amount = tafqeet_arabic(estimated_total)
            written_text = st.text_input(f"المبلغ كتابةً (تفقيط) #{idx+1}", value=written_amount, key=f"writ_{idx}")

        invoices_data.append({
            "supplier": supplier,
            "inv_no": inv_no,
            "account_no": account_no,
            "inv_date": inv_date,
            "amount": amount,
            "tax_no": tax_no,
            "statement": statement,
            "category": category,
            "written_amount": written_text
        })
        st.markdown("---")

    st.subheader("3️⃣ المعالجة والإرسال")
    
    if st.button("🚀 تفريغ العهدة والتجهيز لمنصة فنار"):
        excel_path = "نموذج عهدة الادارة العامة 2026.xlsx.xlsx"
        
        if os.path.exists(excel_path):
            try:
                wb = openpyxl.load_workbook(excel_path)
                ws = wb.active
                
                # تحديث عنوان الخلية A2 بشهر الفاتورة واسم الدار
                first_inv_date_str = invoices_data[0]["inv_date"]
                try:
                    parsed_date = datetime.strptime(first_inv_date_str, '%Y-%m-%d')
                except Exception:
                    parsed_date = datetime.today()
                    
                inv_month_name = get_arabic_month(parsed_date.month)
                inv_year = parsed_date.year
                
                ws['A2'] = f"تفريغ عهدة : {dar_name} شهر {inv_month_name} {inv_year} م -"
                
                start_row = 6
                
                cat_col_map = {
                    "الهاتف والبريد": "I",
                    "الكهرباء": "J",
                    "المياة": "K",
                    "الصيانة": "O",
                    "قرطاسيات ومطبوعات": "P",
                    "نظافة": "R",
                    "أخرى": "W"
                }
                
                for idx, item in enumerate(invoices_data):
                    current_row = start_row + idx
                    
                    ws[f'C{current_row}'] = item["statement"]
                    ws[f'D{current_row}'] = item["account_no"]
                    ws[f'E{current_row}'] = int(item["inv_no"]) if str(item["inv_no"]).isdigit() else item["inv_no"]
                    ws[f'F{current_row}'] = item["inv_date"]
                    ws[f'G{current_row}'] = item["tax_no"]
                    ws[f'G{current_row}'].number_format = '@'
                    ws[f'H{current_row}'] = item["supplier"]
                    
                    col_letter = cat_col_map.get(item["category"], "J")
                    ws[f'{col_letter}{current_row}'] = item["amount"]
                
                output_filename = f"تفريغ_عهدة_{dar_name}_2026.xlsx"
                wb.save(output_filename)
                
                st.success(f"🎉 تم تفريغ البيانات لـ ({dar_name}) وتصشيح رقم الفاتورة وشهر الفاتورة بنجاح!")
                
                with open(output_filename, "rb") as file
