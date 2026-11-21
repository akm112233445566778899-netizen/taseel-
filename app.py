import streamlit as st
import openpyxl
from datetime import datetime
import os
import io
import json
import hashlib
from google import genai
from google.genai import types

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

def get_file_hash(file_bytes):
    return hashlib.md5(file_bytes).hexdigest()

def extract_any_invoice(file_bytes, mime_type):
    """قراءة ذكية لأي فاتورة في السوق (PDF أو صورة) باستخدام Gemini 2.5 Flash"""
    if not API_KEY:
        return {"error": "مفتاح GEMINI_API_KEY غير مضاف في Secrets"}

    try:
        client = genai.Client(api_key=API_KEY)

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

        response = client.models.generate_content(
            model="gemini-2.5-flash",
            contents=[
                types.Part.from_bytes(data=file_bytes, mime_type=mime_type),
                prompt
            ],
            config=types.GenerateContentConfig(
                response_mime_type="application/json",
                temperature=0.0
            )
        )

        return json.loads(response.text)
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

                    raw_inv = str(item["inv_no"]).strip()
                    ws[f'E{current_row}'] = int(raw_inv) if raw_inv.isdigit() else raw_inv

                    ws[f'F{current_row}'] = item["inv_date"]
                    ws[f'G{current_row}'] = item["tax_no"]
                    ws[f'G{current_row}'].number_format = '@'
                    ws[f'H{current_row}'] = item["supplier"]

                    col_letter = cat_col_map.get(item["category"], "W")
                    ws[f'{col_letter}{current_row}'] = item["amount"]

                # حفظ الملف في الذاكرة (RAM) وتفريغ الكائنات لمنع امتلاء السيرفر
                output_buffer = io.BytesIO()
                wb.save(output_buffer)
                wb.close()
                output_buffer.seek(0)

                output_filename = f"تفريغ_عهدة_{dar_name}_2026.xlsx"

                st.success(f"🎉 تم تفريغ البيانات لـ ({dar_name}) بنجاح!")

                st.download_button(
                    label="📥 تحميل ملف Excel المفرغ جاهز للرفع",
                    data=output_buffer,
                    file_name=output_filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

                st.link_button("🌐 الذهاب إلى منصة فنار لإنشاء أمر الصرف", "https://fanar.gqb.org.sa/requests/home")

            except Exception as e:
                st.error(f"حدث خطأ أثناء معالجة الملف: {e}")
        else:
            st.error(f"⚠️ لم يتم العثور على الملف ({excel_path}) داخل المستودع.")

# زر لتنظيف الذاكرة عند الانتهاء من مجموعة كبيرة من الدور
if st.session_state.parsed_cache:
    if st.sidebar.button("🧹 مسح ذاكرة الفواتير السابقة"):
        st.session_state.parsed_cache.clear()
        st.rerun()
