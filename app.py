import streamlit as st
import openpyxl
from datetime import datetime
import os
import json
import tempfile
import google.generativeai as genai

st.set_page_config(
    page_title="أداة تفريغ العهود - 2026",
    page_icon="🏛️",
    layout="centered"
)

st.title("🏛️ أداة تفريغ العهود وإصدار أمر الصرف")
st.caption("الجمعية الخيرية لتحفيظ القرآن ببريدة - نموذج الإدارة العامة 2026")
st.divider()

API_KEY = st.secrets.get("GEMINI_API_KEY", "")

def tafqeet_arabic(amount):
    """تحويل المبلغ إلى كلمات عربية (تفقيط)"""
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

def convert_arabic_digits(text):
    """تحويل الأرقام العربية الشرقية إلى أرقام إنجليزية standard"""
    arabic_digits = "٠١٢٣٤٥٦٧٨٩"
    english_digits = "0123456789"
    translation_table = str.maketrans(arabic_digits, english_digits)
    return str(text).translate(translation_table)

def extract_invoice_data(uploaded_file):
    """استخراج أرقام الفواتير والحسابات والتواريخ والمبالغ بالضبط"""
    if not API_KEY:
        return {}
    try:
        genai.configure(api_key=API_KEY)
        model = genai.GenerativeModel('gemini-1.5-flash')
        
        suffix = ".pdf" if uploaded_file.type == "application/pdf" else ".png"
        with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp_file:
            uploaded_file.seek(0)
            tmp_file.write(uploaded_file.read())
            tmp_path = tmp_file.name

        sample_file = genai.upload_file(path=tmp_path)
        
        prompt = """
        أنت دقيق جداً في مراجعة فواتير الكهرباء والخدمات السعودية. قم بفحص المستند المرفق بدقة واستخرج القيم التالية بالأرقام الإنجليزية (0-9):

        1. inv_no: رقم الفاتورة المكتوب بوضوح (مثال: 252003458719).
        2. account_no: رقم الحساب المكتوب في ملخص الفاتورة (مثال: 30105906941).
        3. inv_date: تاريخ بداية الاستحقاق / تاريخ الفاتورة المكتوب في خانة تاريخ الفاتورة بصيغة YYYY-MM-DD (مثال: 2026-10-17).
        4. amount: المبلغ الخاضع للضريبة / إجمالي المبلغ غير شامل ضريبة القيمة المضافة (مثال: 407.42).
        5. tax_no: الرقم الضريبي للمورد (مثال: 300000361310003).
        6. supplier: اسم الشركة الموردة (مثل: الشركة السعودية للكهرباء).

        قم بإرجاع النتيجة بصيغة JSON حصرية فقط بدون أي كلام خارجي:
        {
            "inv_no": "252003458719",
            "account_no": "30105906941",
            "inv_date": "2026-10-17",
            "amount": 407.42,
            "tax_no": "300000361310003",
            "supplier": "الشركة السعودية للكهرباء"
        }
        """
        
        response = model.generate_content([sample_file, prompt])
        
        os.remove(tmp_path)
        genai.delete_file(sample_file.name)
        
        clean_json = response.text.replace("```json", "").replace("```", "").strip()
        data = json.loads(clean_json)
        
        # تحويل الأرقام لاستبعاد أي خطأ بالأرقام العربية
        if "inv_no" in data: data["inv_no"] = convert_arabic_digits(data["inv_no"])
        if "account_no" in data: data["account_no"] = convert_arabic_digits(data["account_no"])
        if "tax_no" in data: data["tax_no"] = convert_arabic_digits(data["tax_no"])
        if "inv_date" in data: data["inv_date"] = convert_arabic_digits(data["inv_date"])
        
        return data
    except Exception as e:
        return {}

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
    
    if not API_KEY:
        st.error("⚠️ مفتاح GEMINI_API_KEY غير مفعّل في Secrets.")
        
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
        
        ext_data = {}
        if API_KEY:
            with st.spinner(f"🔍 جاري استخراج البيانات الدقيقة للمستند #{idx + 1}..."):
                ext_data = extract_invoice_data(file)
            
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
                
                st.success(f"🎉 تم تفريغ البيانات لـ ({dar_name}) وتعديل شهر الإصدار إلى ({inv_month_name} {inv_year}) بنجاح!")
                
                with open(output_filename, "rb") as file:
                    st.download_button(
                        label="📥 تحميل ملف Excel المفرغ جاهز للرفع",
                        data=file,
                        file_name=output_filename,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                
                st.link_button("🌐 الذهاب إلى منصة فنار لإنشاء أمر الصرف", "https://fanar.gqb.org.sa/requests/home")
                
            except Exception as e:
                st.error(f"حدث خطأ أثناء معالجة الملف: {e}")
        else:
            st.error(f"⚠️ لم يتم العثور على الملف ({excel_path}) داخل المستودع.")
