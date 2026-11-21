import streamlit as st
import openpyxl
from datetime import datetime
import os
import io
import json
import hashlib
import base64
import urllib.request
import urllib.error

st.set_page_config(
    page_title="نظام تسهيل | تفريغ العهود 2026",
    page_icon="⚡",
    layout="centered"
)

# --- عرض الشعار واسم البرنامج "تسهيل" ---
if os.path.exists("logo.png"):
    col_logo, col_title = st.columns([1, 3.5])
    with col_logo:
        st.image("logo.png", width=120)
    with col_title:
        st.title("⚡ نظام تسهيل")
        st.subheader("أداة تفريغ العهود وإصدار أمر الصرف")
        st.caption("الجمعية الخيرية لتحفيظ القرآن ببريدة - الإدارة العامة 2026")
else:
    st.title("⚡ نظام تسهيل")
    st.subheader("أداة تفريغ العهود وإصدار أمر الصرف")
    st.caption("الجمعية الخيرية لتحفيظ القرآن ببريدة - الإدارة العامة 2026")

st.divider()

# تهيئة ذاكرة الجلسة لمنع إعادة القراءة والتعليق
if "parsed_cache" not in st.session_state:
    st.session_state.parsed_cache = {}

API_KEY = st.secrets.get("GEMINI_API_KEY", "").strip()

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

def get_file_hash(file_bytes):
    return hashlib.md5(file_bytes).hexdigest()

def extract_any_invoice(file_bytes, mime_type):
    """قراءة ذكية لأي فاتورة في السوق (PDF أو صورة) تدعم جميع صيغ مفاتيح Gemini"""
    if not API_KEY:
        return {"error": "مفتاح GEMINI_API_KEY غير مضاف في Secrets"}

    try:
        b64_data = base64.b64encode(file_bytes).decode("utf-8")

        prompt = """
        أنت محاسب سعودي خبير. افحص الفاتورة المرفقة (سواء كانت فاتورة كهرباء، مياه، محل تجاري، سباكة، قرطاسية، أو أي مؤسسة) واستخرج البيانات بدقة 100%:

        1. supplier: اسم المحل أو الشركة الموردة كما هو مكتوب في أعلى الفاتورة أو الختم (مثال: الشركة السعودية للكهرباء، مكتبة جرير، مؤسسة كذا...).
        2. tax_no: الرقم الضريبي للمورد المكون من 15 رقم (يبدأ وينتهي بـ 3 غالباً). إذا لم يوجد اتركه فارغاً "".
        3. inv_no: رقم الفاتورة (Invoice Number / Bill No). في فواتير الكهرباء يكون 12 رقم يبدأ بـ 20 أو 25. في فواتير المحلات هو رقم الفاتورة الضريبية.
        4. account_no: رقم الحساب أو رقم المشترك أو رقم العميل إن وجد (في الكهرباء 11 رقم يبدأ بـ 30)، وإلا اتركه فارغاً "".
        5. inv_date: تاريخ إصدار الفاتورة أو تاريخ بداية الاستحقاق بصيغة YYYY-MM-DD (تجنب تاريخ بداية/نهاية فترة الاستهلاك في الكهرباء).
        6. amount_pre_tax: المبلغ الإجمالي قبل الضريبة (المبلغ الخاضع للضريبة / غير شامل الضريبة) كرقم عشري. ملاحظة هامة: إذا كانت الفاتورة تعرض المبلغ الإجمالي شامل الضريبة (15%) فقط ولا تذكر المبلغ قبل الضريبة، فاقسم الإجمالي على 1.15 وقرّبه لأقرب منزلتين عشريتين.
        7. category: اختر أنسب بند من هذه القائمة حصراً بناءً على نشاط المحل أو محتوى الفاتورة:
           ["الكهرباء", "المياة", "الهاتف والبريد", "الصيانة", "قرطاسيات ومطبوعات", "نظافة", "أخرى"]
        8. description: وصف مختصر جداً لمحتوى الفاتورة (مثال: فاتورة كهرباء، أدوات سباكة وصيانة، منظفات، أوراق وقرطاسية).

        أعد النتيجة بصيغة JSON صالحة فقط بدون أي نص إضافي:
        {
            "supplier": "",
            "tax_no": "",
            "inv_no": "",
            "account_no": "",
            "inv_date": "YYYY-MM-DD",
            "amount_pre_tax": 0.0,
            "category": "أخرى",
            "description": ""
        }
        """

        url = "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent"
        headers = {"Content-Type": "application/json"}

        if API_KEY.startswith("AIza"):
            url += f"?key={API_KEY}"
        else:
            headers["Authorization"] = f"Bearer {API_KEY}"

        payload = {
            "contents": [{
                "parts": [
                    {"inline_data": {"mime_type": mime_type, "data": b64_data}},
                    {"text": prompt}
                ]
            }],
            "generationConfig": {
                "response_mime_type": "application/json",
                "temperature": 0.0
            }
        }

        req = urllib.request.Request(
            url,
            data=json.dumps(payload).encode("utf-8"),
            headers=headers,
            method="POST"
        )
        with urllib.request.urlopen(req, timeout=60) as resp:
            result = json.loads(resp.read().decode("utf-8"))
            text_out = result["candidates"][0]["content"]["parts"][0]["text"]
            return json.loads(text_out)

    except urllib.error.HTTPError as e:
        err_body = e.read().decode("utf-8", errors="ignore")
        return {"error": f"خطأ من خادم Google ({e.code}): {err_body}"}
    except Exception as e:
        return {"error": str(e)}

# --- القسم الأول: رفع الفواتير ---
st.subheader("1️⃣ رفع الفواتير (أي فاتورة PDF أو صورة)")
uploaded_files = st.file_uploader(
    "ارفع فواتير الكهرباء أو فواتير المحلات والمشتريات (حتى 5 فواتير معاً)",
    type=["pdf", "png", "jpg", "jpeg", "webp"],
    accept_multiple_files=True
)

if uploaded_files:
    if len(uploaded_files) > 5:
        st.warning("⚠️ تم تحديد أكثر من 5 فواتير. سيتم معالجة أول 5 فواتير فقط.")
        uploaded_files = uploaded_files[:5]

    if not API_KEY:
        st.error("⚠️ تنبيه: مفتاح GEMINI_API_KEY غير موجود في إعدادات Secrets في Streamlit.")

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

    dar_name = st.selectbox("اختر اسم الدار المعنية", dar_list)
    categories_list = [
        "الكهرباء",
        "المياة",
        "الهاتف والبريد",
        "الصيانة",
        "قرطاسيات ومطبوعات",
        "نظافة",
        "أخرى"
    ]

    invoices_data = []

    for idx, file in enumerate(uploaded_files):
        file_bytes = file.getvalue()
        f_hash = get_file_hash(file_bytes)

        # القراءة مرة واحدة فقط وحفظها في الذاكرة لمنع التعليق
        if f_hash not in st.session_state.parsed_cache:
            with st.spinner(f"🔍 جاري قراءة الفاتورة #{idx + 1} ({file.name})..."):
                mime_type = file.type or "application/pdf"
                st.session_state.parsed_cache[f_hash] = extract_any_invoice(file_bytes, mime_type)

        ext_data = st.session_state.parsed_cache[f_hash]

        st.markdown(f"### 📄 الفاتورة رقم ({idx + 1}): `{file.name}`")
        if "error" in ext_data:
            st.warning(f"تعذرت القراءة التلقائية: {ext_data['error']} (يمكنك الإدخال يدوياً)")

        detected_cat = ext_data.get("category", "أخرى")
        default_cat_index = categories_list.index(detected_cat) if detected_cat in categories_list else 6
        default_desc = ext_data.get("description", "سداد فاتورة")
        default_supplier = ext_data.get("supplier", "")

        col1, col2 = st.columns(2)
        with col1:
            supplier = st.text_input(
                f"اسم المورد / المحل #{idx+1}",
                value=str(default_supplier),
                key=f"sup_{f_hash}_{idx}"
            )
            inv_no = st.text_input(
                f"رقم الفاتورة #{idx+1}",
                value=str(ext_data.get("inv_no", "")),
                key=f"inv_{f_hash}_{idx}"
            )
            account_no = st.text_input(
                f"رقم الحساب / المرجع #{idx+1}",
                value=str(ext_data.get("account_no", "")),
                key=f"acc_{f_hash}_{idx}"
            )
            inv_date = st.text_input(
                f"تاريخ الفاتورة #{idx+1}",
                value=str(ext_data.get("inv_date", "")),
                key=f"date_{f_hash}_{idx}"
            )

        with col2:
            tax_no = st.text_input(
                f"الرقم الضريبي #{idx+1}",
                value=str(ext_data.get("tax_no", "")),
                key=f"tax_{f_hash}_{idx}"
            )
            amount = st.number_input(
                f"المبلغ بدون الضريبة #{idx+1}",
                value=float(ext_data.get("amount_pre_tax", 0.0) or 0.0),
                step=0.01,
                key=f"amt_{f_hash}_{idx}"
            )
            statement = st.text_input(
                f"البيان #{idx+1}",
                value=f"{default_desc} - {dar_name}",
                key=f"st_{f_hash}_{idx}"
            )
            category = st.selectbox(
                f"بند الحساب #{idx+1}",
                categories_list,
                index=default_cat_index,
                key=f"cat_{f_hash}_{idx}"
            )

            estimated_total = amount * 1.15
            written_amount = tafqeet_arabic(estimated_total)
            written_text = st.text_input(
                f"المبلغ كتابةً (تفقيط) #{idx+1}",
                value=written_amount,
                key=f"writ_{f_hash}_{idx}"
            )

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

                first_inv_date_str = invoices_data[0]["inv_date"]
                try:
                    parsed_date = datetime.strptime(first_inv_date_str, '%Y-%m-%d')
                except Exception:
                    parsed_date = datetime.today()

                inv_month_name = get_arabic_month(parsed_date.month)
                inv_year = parsed_date.year

                ws['
