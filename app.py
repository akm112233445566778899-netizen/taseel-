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

def extract_invoice_with_ai(uploaded_file):
    """استخراج ذكي وشامل لجميع أنواع واختلافات الفواتير عبر Gemini AI"""
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
        أنت محاسب قانوني تدقق في فواتير الكهرباء والخدمات السعودية. قم بفحص المستند المرفق بدقة متناهية واستخرج البيانات التالية بالظبط:

        1. inv_no: رقم الفاتورة (Invoice Number) المكون غالباً من 12 رقم يبدأ بـ 20 أو 25 (مثال: 208026326251 أو 252003458719).
        2. account_no: رقم الحساب / رقم المشترك (Account Number) المكون من 11 رقم يبدأ بـ 30 (مثال: 30136196833 أو 30105906941).
        3. inv_date: تاريخ بداية الاستحقاق أو تاريخ الفاتورة المكتوب في ملخص الفاتورة بصيغة YYYY-MM-DD (احذر: لا تأخذ تاريخ بداية الفترة أو نهاية الفترة، بل تاريخ بداية الاستحقاق أو تاريخ الفاتورة مثل: 2026-10-11 أو 2026-10-17).
        4. amount: المبلغ الخاضع للضريبة / إجمالي المبلغ غير شامل ضريبة القيمة المضافة (أرقام فقط كعدد عشري مثل: 1821.40 أو 407.42).
        5. tax_no: الرقم الضريبي للمورد (غالباً: 300000361310003).
        6. supplier: اسم المورد (الشركة السعودية للكهرباء).

        ضع النتائج في كائن JSON حصراً بدون أي كلام إضافي:
        {
            "inv_no": "رقم الفاتورة",
            "account_no": "رقم الحساب",
            "inv_date": "YYYY-MM-DD",
            "amount": 0.00,
            "tax_no": "الرقم الضريبي",
            "supplier": "اسم الشركة"
        }
        """
        
        response = model.generate_content([sample_file, prompt])
        
        os.remove(tmp_path)
        try:
            genai.delete_file(sample_file.name)
        except Exception:
            pass
            
        clean_json = response.text.replace("```json", "").replace("```", "").strip()
        return json.loads(clean_json)
    except Exception:
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
    
    dar_name = st.selectbox("
