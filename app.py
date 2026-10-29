import io
import os
import re
import zipfile
from collections import Counter
from datetime import datetime

import openpyxl
import pdfplumber
import streamlit as st

st.set_page_config(
    page_title="نظام تسهيل - تفريغ العهود 2026",
    page_icon="⚡",
    layout="centered",
)

# --- عرض الشعار واسم البرنامج "تسهيل" أعلى الواجهة ---
if os.path.exists("logo.png"):
    col_logo, col_title = st.columns([1, 3.5])
    with col_logo:
        st.image("logo.png", width=120)
    with col_title:
        st.title("⚡ نظام تسهيل")
        st.subheader("أداة تفريغ العهود وإصدار أمر الصرف")
        st.caption(
            "الجمعية الخيرية لتحفيظ القرآن ببريدة - نموذج الإدارة العامة 2026"
        )
else:
    st.title("⚡ نظام تسهيل")
    st.subheader("أداة تفريغ العهود وإصدار أمر الصرف")
    st.caption(
        "الجمعية الخيرية لتحفيظ القرآن ببريدة - نموذج الإدارة العامة 2026"
    )

st.divider()

# ----------------------------------------------------------------------
# الإعدادات العامة
# ----------------------------------------------------------------------
VAT_RATE = 0.15
START_ROW = 6  # أول صف بيانات في جميع النماذج
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

# النموذج الأساسي (لكل البنود التي ليس لها ملف مستقل)
MAIN_TEMPLATE = "نموذج عهدة الادارة العامة 2026.xlsx.xlsx"

# القيم الافتراضية لفواتير الكهرباء
DEFAULT_SUPPLIER = "الشركة السعودية للكهرباء"
DEFAULT_TAX_NO = "300000361310003"

CATEGORIES = [
    "الكهرباء",
    "المياة",
    "الهاتف والبريد",
    "الصيانة",
    "قرطاسيات ومطبوعات",
    "نظافة",
    "أخرى",
]

# عمود كل بند في النموذج الأساسي
MAIN_CAT_COL_MAP = {
    "الهاتف والبريد": "I",
    "الكهرباء": "J",
    "المياة": "K",
    "الصيانة": "O",
    "قرطاسيات ومطبوعات": "P",
    "نظافة": "R",
    "أخرى": "W",
}

# ----------------------------------------------------------------------
# البنود التي يُخرج لها ملف إكسل مستقل
# لإضافة بند جديد (مثلاً المياة) يكفي إضافة عنصر هنا + وضع ملف النموذج
# الخاص به في مجلد التطبيق.
#   template          : اسم ملف النموذج
#   title             : عنوان التفريغ في الخلية A2
#   file_name         : اسم الملف الناتج
#   default_statement : البيان الافتراضي
#   columns           : الأعمدة المتاحة للمبلغ في هذا النموذج (اسم -> حرف العمود)
# ----------------------------------------------------------------------
SEPARATE_FILES = {
    "الصيانة": {
        "template": "نموذج عهدة الصيانة 2026.xlsx",
        "title": "تفريغ عهدة : صيانة {dar} {period} م -",
        "file_name": "تفريغ_عهدة_صيانة_{dar}.xlsx",
        "default_statement": "قيمة صيانة",
        "columns": {
            "الصيانة": "J",
            "صيانة مباني": "O",
            "صيانة سيارات": "V",
        },
    },
}

ARABIC_QUARTERS = {1: "الأول", 2: "الثاني", 3: "الثالث", 4: "الرابع"}


# ----------------------------------------------------------------------
# دوال مساعدة
# ----------------------------------------------------------------------
def convert_arabic_digits(text):
    """تحويل الأرقام العربية الشرقية (٠١٢٣٤٥٦٧٨٩) إلى أرقام إنجليزية"""
    if not text:
        return ""
    arabic_digits = "٠١٢٣٤٥٦٧٨٩"
    english_digits = "0123456789"
    return str(text).translate(str.maketrans(arabic_digits, english_digits))


def tafqeet_arabic(amount):
    """تحويل المبلغ إلى كلمات (صيغة مبسطة: أرقام + فقط لا غير)"""
    try:
        if amount <= 0:
            return ""
        amount = round(amount, 2)
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
        1: "يناير",
        2: "فبراير",
        3: "مارس",
        4: "أبريل",
        5: "مايو",
        6: "يونيو",
        7: "يوليو",
        8: "أغسطس",
        9: "سبتمبر",
        10: "أكتوبر",
        11: "نوفمبر",
        12: "ديسمبر",
    }
    return months.get(month_num, "")


def to_float(s):
    return float(convert_arabic_digits(s).replace(",", "").strip())


def parse_date(s):
    """يحوّل نص بصيغة YYYY-MM-DD إلى تاريخ، أو None إذا كان غير صالح"""
    try:
        return datetime.strptime(str(s).strip(), "%Y-%m-%d")
    except Exception:
        return None


def guess_quarter_label(items):
    """
    يقترح فترة التفريغ (مثل: الربع الثالث 2026) حسب الربع الذي تقع فيه
    أغلب الفواتير. عند التعادل يؤخذ الأقدم. المستخدم يقدر يعدّلها يدوياً.
    """
    quarters = []
    for it in items:
        dt = parse_date(it["inv_date"])
        if dt:
            quarters.append((dt.year, (dt.month - 1) // 3 + 1))
    if not quarters:
        today = datetime.today()
        quarters = [(today.year, (today.month - 1) // 3 + 1)]
    counts = Counter(quarters)
    top = max(counts.values())
    year, quarter = min(k for k, v in counts.items() if v == top)
    return f"الربع {ARABIC_QUARTERS[quarter]} {year}"


def safe_filename(text):
    """اسم صالح للملفات: المسافات تصير _ وتُحذف الرموز الممنوعة"""
    text = re.sub(r'[\\/:*?"<>|]', "", str(text)).strip()
    return re.sub(r"\s+", "_", text)


def write_invoice_row(ws, row, item, date_format=None):
    """يكتب فاتورة واحدة في صف من النموذج (الخانات الفارغة تبقى فارغة)"""
    ws[f"C{row}"] = item["statement"]

    if item["account_no"]:
        ws[f"D{row}"] = item["account_no"]

    raw_inv = str(item["inv_no"]).strip()
    if raw_inv.isdigit():
        ws[f"E{row}"] = int(raw_inv)
    elif raw_inv:
        ws[f"E{row}"] = raw_inv

    # حفظ التاريخ كتاريخ حقيقي في الإكسل (وليس نص)
    inv_dt = parse_date(item["inv_date"])
    if inv_dt:
        ws[f"F{row}"] = inv_dt
        if date_format:
            ws[f"F{row}"].number_format = date_format
    elif str(item["inv_date"]).strip():
        ws[f"F{row}"] = item["inv_date"]

    if item["tax_no"]:
        ws[f"G{row}"] = item["tax_no"]
        ws[f"G{row}"].number_format = "@"
    if item["supplier"]:
        ws[f"H{row}"] = item["supplier"]

    ws[f"{item['amount_col']}{row}"] = float(item["amount"])


def build_workbook(template_path, title, items, date_format=None):
    """يفتح النموذج، يكتب العنوان والفواتير، ويرجع الملف كبايتات"""
    wb = openpyxl.load_workbook(template_path)
    ws = wb.active
    ws["A2"] = title
    for idx, item in enumerate(items):
        write_invoice_row(ws, START_ROW + idx, item, date_format)
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ----------------------------------------------------------------------
# استخراج التاريخ
# ----------------------------------------------------------------------
def extract_invoice_date(full_text):
    """
    يستخرج تاريخ الفاتورة بدون الاعتماد على الكلمات العربية
    (لأن خط الـ PDF أحياناً يطلع حروف مشوهة).

    الفاتورة فيها 4 تواريخ:
      بداية الفترة < نهاية الفترة <= تاريخ الفاتورة < تاريخ بداية الاستحقاق
    فتاريخ الفاتورة هو ثاني أكبر تاريخ.
    يرجع نص بصيغة YYYY-MM-DD أو "" إذا ما لقى.
    """
    # الطريقة الأولى (الأدق): تاريخ الفاتورة يكون في نفس سطر رقم الحساب،
    # أما تاريخ الاستحقاق فيكون في سطر مستقل تحته، فلا يُؤخذ أبداً.
    for line in full_text.split("\n"):
        if re.search(r"(?<!\d)30\d{9}(?!\d)", line):
            line_dates = []
            for d, m, y in re.findall(
                r"(?<!\d)(\d{2})[-/](\d{2})[-/](\d{4})(?!\d)", line
            ):
                try:
                    dt = datetime(int(y), int(m), int(d))
                    if 2000 <= dt.year <= 2100:
                        line_dates.append(dt)
                except ValueError:
                    pass
            if len(line_dates) == 1:
                return line_dates[0].strftime("%Y-%m-%d")

    # الطريقة الاحتياطية: ثاني أكبر تاريخ في الفاتورة
    found = set()

    # DD-MM-YYYY أو DD/MM/YYYY
    for d, m, y in re.findall(r"(?<!\d)(\d{2})[-/](\d{2})[-/](\d{4})(?!\d)", full_text):
        try:
            dt = datetime(int(y), int(m), int(d))
            if 2000 <= dt.year <= 2100:  # يستبعد التواريخ الهجرية
                found.add(dt)
        except ValueError:
            pass

    # YYYY-MM-DD أو YYYY/MM/DD
    for y, m, d in re.findall(r"(?<!\d)(\d{4})[-/](\d{2})[-/](\d{2})(?!\d)", full_text):
        try:
            dt = datetime(int(y), int(m), int(d))
            if 2000 <= dt.year <= 2100:
                found.add(dt)
        except ValueError:
            pass

    dates = sorted(found)
    if not dates:
        return ""
    if len(dates) == 1:
        return dates[0].strftime("%Y-%m-%d")
    return dates[-2].strftime("%Y-%m-%d")


# ----------------------------------------------------------------------
# استخراج المبلغ
# ----------------------------------------------------------------------
def extract_net_amount(full_text):
    """
    يبحث عن ثلاثية متطابقة حسابياً:
      صافي + ضريبة (15%) = إجمالي
    ويأخذ أكبر صافي يحقق الشرط (وهو إجمالي الفاتورة، مو بنود الجدول).
    كمية الاستهلاك (ك.و.س) ما تحقق المعادلة فتتجاهل تلقائياً.
    """
    amt_re = r"(?<![\d,.])(?:\d{1,3}(?:,\d{3})+|\d+)\.\d{2}(?!\d)"
    nums = set()
    for x in re.findall(amt_re, full_text):
        try:
            nums.add(to_float(x))
        except ValueError:
            pass

    best = 0.0
    for net in nums:
        if net < 1:
            continue
        vat = round(net * VAT_RATE, 2)
        total = round(net + vat, 2)
        has_vat = any(abs(vat - x) < 0.02 for x in nums)
        has_total = any(abs(total - x) < 0.02 for x in nums)
        if has_vat and has_total and net > best:
            best = net
    return best


# ----------------------------------------------------------------------
# تحليل الفاتورة
# ----------------------------------------------------------------------
def parse_pdf_invoice_robust(uploaded_file):
    data = {
        "inv_no": "",
        "account_no": "",
        "inv_date": "",
        "amount": 0.0,
        "tax_no": DEFAULT_TAX_NO,
        "supplier": DEFAULT_SUPPLIER,
        "raw_text": "",
        "error": "",
    }

    if not uploaded_file.name.lower().endswith(".pdf"):
        data["error"] = "الاستخراج التلقائي يعمل مع ملفات PDF فقط، أدخل البيانات يدوياً."
        return data

    try:
        uploaded_file.seek(0)
        full_text_raw = ""
        with pdfplumber.open(uploaded_file) as pdf:
            for page in pdf.pages:
                full_text_raw += (page.extract_text() or "") + "\n"

        data["raw_text"] = full_text_raw
        full_text = convert_arabic_digits(full_text_raw)

        # 1. رقم الفاتورة (12 رقم يبدأ بـ 20 أو 25)
        m = re.search(r"(?<!\d)(2[05]\d{10})(?!\d)", full_text)
        if m:
            data["inv_no"] = m.group(1)

        # 2. رقم الحساب (11 رقم يبدأ بـ 30)
        m = re.search(r"(?<!\d)(30\d{9})(?!\d)", full_text)
        if m:
            data["account_no"] = m.group(1)

        # 3. الرقم الضريبي للمورد (15 رقم يبدأ بـ 3000)
        m = re.search(r"(?<!\d)(3000\d{11})(?!\d)", full_text)
        if m:
            data["tax_no"] = m.group(1)

        # 4. تاريخ الفاتورة
        data["inv_date"] = extract_invoice_date(full_text)

        # 5. المبلغ بدون ضريبة
        data["amount"] = extract_net_amount(full_text)

    except Exception as e:
        data["error"] = f"تعذر قراءة الملف: {e}"

    return data


# ----------------------------------------------------------------------
# الواجهة
# ----------------------------------------------------------------------
dar_list = [
    "إدارة الدور النسائية",
    "دار حصة",
    "دار التقوى",
    "دار الصفا",
    "دار سمية بنت الخياط",
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
    "دار بينات",
    "الأفق",
    "الوقار",
    "مركز محكم",
    "دار الهداية",
    "الرابية",
    "دار الطرفية",
]

OTHER_DAR = "➕ دار أخرى (اكتب الاسم)"

st.subheader("اختر الدار")
dar_choice = st.selectbox("اختر اسم الدار المعنية", dar_list + [OTHER_DAR])
if dar_choice == OTHER_DAR:
    dar_name = st.text_input("اسم الدار", key="custom_dar_name").strip()
    if not dar_name:
        st.info("اكتب اسم الدار للمتابعة.")
        st.stop()
else:
    dar_name = dar_choice

st.subheader("1️⃣ رفع الفواتير (حتى 5 فواتير)")
uploaded_files = st.file_uploader(
    "التقط صور الفواتير أو اسحب ملفات الـ PDF هنا (حتى 5 فواتير معاً)",
    type=["pdf", "png", "jpg", "jpeg"],
    accept_multiple_files=True,
)

if uploaded_files:
    if len(uploaded_files) > 5:
        st.warning("⚠️ تم تحديد أكثر من 5 فواتير. سيتم معالجة أول 5 فواتير فقط.")
        uploaded_files = uploaded_files[:5]

    st.success(f"✅ تم استلام {len(uploaded_files)} فاتورة بنجاح!")
    st.divider()

    st.subheader("2️⃣ مراجعة وتحديد البيانات")

    invoices_data = []

    for idx, file in enumerate(uploaded_files):
        st.markdown(f"### 📄 الفاتورة رقم ({idx + 1}): `{file.name}`")

        with st.spinner(f"🔍 جاري استخراج البيانات للفاتورة #{idx + 1}..."):
            ext_data = parse_pdf_invoice_robust(file)

        if ext_data.get("error"):
            st.warning(f"⚠️ {ext_data['error']}")
        else:
            if not ext_data["inv_date"]:
                st.warning("⚠️ لم يتم العثور على تاريخ الفاتورة، أدخله يدوياً.")
            if not ext_data["amount"]:
                st.warning("⚠️ لم يتم العثور على المبلغ، أدخله يدوياً.")

        with st.expander(f"عرض النص المستخرج للفاتورة #{idx + 1} (للفحص)"):
            st.text(ext_data.get("raw_text", ""))

        # البند يُحدَّد أولاً لأن القيم الافتراضية تعتمد عليه
        category = st.selectbox(
            f"بند الحساب #{idx+1}",
            CATEGORIES,
            key=f"cat_{idx}",
        )
        sep_cfg = SEPARATE_FILES.get(category)

        if sep_cfg:
            st.caption(f"📁 فواتير «{category}» تُفرَّغ في ملف إكسل مستقل.")
            # فواتير هذا البند من موردين مختلفين، فلا نفترض مورداً معيّناً
            default_supplier = ""
            default_tax_no = ""
            default_statement = sep_cfg["default_statement"]
        else:
            default_supplier = ext_data.get("supplier", DEFAULT_SUPPLIER)
            default_tax_no = str(ext_data.get("tax_no", DEFAULT_TAX_NO))
            default_statement = f"سداد فاتورة كهرباء - {dar_name}"

        col1, col2 = st.columns(2)
        with col1:
            supplier = st.text_input(
                f"اسم المورد #{idx+1}",
                value=default_supplier,
                key=f"sup_{idx}_{category}",
            )
            inv_no = st.text_input(
                f"رقم الفاتورة #{idx+1}",
                value=str(ext_data.get("inv_no", "")),
                key=f"inv_{idx}",
            )
            account_no = st.text_input(
                f"رقم الحساب #{idx+1}",
                value=str(ext_data.get("account_no", "")),
                key=f"acc_{idx}",
            )
            inv_date = st.text_input(
                f"تاريخ الفاتورة #{idx+1} (YYYY-MM-DD)",
                value=str(ext_data.get("inv_date", "")),
                key=f"date_{idx}",
            )

        with col2:
            tax_no = st.text_input(
                f"الرقم الضريبي #{idx+1}",
                value=default_tax_no,
                key=f"tax_{idx}_{category}",
            )

            extracted_net = float(ext_data.get("amount", 0.0) or 0.0)

            # فواتير الملفات المستقلة غالباً يكون المعلوم عندك هو الإجمالي،
            # فنتيح إدخاله شامل الضريبة ونحسب المبلغ قبل الضريبة تلقائياً
            if sep_cfg:
                amount_mode = st.radio(
                    f"المبلغ المُدخل #{idx+1}",
                    ["قبل الضريبة", "شامل الضريبة"],
                    horizontal=True,
                    key=f"mode_{idx}",
                )
            else:
                amount_mode = "قبل الضريبة"

            if amount_mode == "شامل الضريبة":
                gross = st.number_input(
                    f"المبلغ شامل الضريبة #{idx+1}",
                    value=round(extracted_net * (1 + VAT_RATE), 2),
                    step=0.01,
                    key=f"amt_{idx}_gross",
                )
                amount = round(gross / (1 + VAT_RATE), 2)
                estimated_total = round(gross, 2)
                st.caption(
                    f"المبلغ قبل الضريبة المحسوب: {amount:,.2f} "
                    "(هذا الرقم اللي ينكتب في الملف)"
                )
            else:
                amount = st.number_input(
                    f"المبلغ بدون الضريبة #{idx+1}",
                    value=extracted_net,
                    step=0.01,
                    key=f"amt_{idx}",
                )
                estimated_total = round(amount + round(amount * VAT_RATE, 2), 2)

            statement = st.text_input(
                f"البيان #{idx+1}",
                value=default_statement,
                key=f"st_{idx}_{dar_name}_{category}",
            )

            # عمود المبلغ في النموذج
            if sep_cfg:
                cols_map = sep_cfg["columns"]
                if len(cols_map) > 1:
                    sub_label = st.selectbox(
                        f"نوع {category} #{idx+1}",
                        list(cols_map),
                        key=f"sub_{idx}_{category}",
                    )
                    amount_col = cols_map[sub_label]
                else:
                    amount_col = next(iter(cols_map.values()))
            else:
                amount_col = MAIN_CAT_COL_MAP.get(category, "J")

            written_amount = tafqeet_arabic(estimated_total)
            written_text = st.text_input(
                f"المبلغ كتابةً (تفقيط) #{idx+1}",
                value=written_amount,
                key=f"writ_{idx}_{estimated_total}",
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
            "amount_col": amount_col,
            "written_amount": written_text,
        })
        st.markdown("---")

    # تقسيم الفواتير: ملفات مستقلة حسب البند + الباقي للنموذج الأساسي
    main_items = [it for it in invoices_data if it["category"] not in SEPARATE_FILES]
    sep_groups = {}
    for cat in SEPARATE_FILES:
        items = [it for it in invoices_data if it["category"] == cat]
        if items:
            sep_groups[cat] = items

    st.subheader("3️⃣ المعالجة والإرسال")

    # فترة التفريغ لكل ملف مستقل (تظهر في عنوان الملف) — تُقترح تلقائياً وتتعدّل
    period_labels = {}
    for cat, items in sep_groups.items():
        default_period = guess_quarter_label(items)
        period_labels[cat] = st.text_input(
            f"فترة تفريغ {cat} (تظهر في عنوان الملف)",
            value=default_period,
            key=f"period_{cat}_{default_period}",
        )

    plan = []
    if main_items:
        plan.append(f"الملف الأساسي ({len(main_items)} فاتورة)")
    for cat, items in sep_groups.items():
        plan.append(f"ملف {cat} ({len(items)} فاتورة)")
    st.caption("سيتم إنشاء: " + " + ".join(plan))

    # بصمة المدخلات: نعرض نتيجة التفريغ فقط ما دام ما تغيّر شي بعدها
    fingerprint = repr((dar_name, invoices_data, period_labels))

    if st.button("🚀 تفريغ العهدة والتجهيز لمنصة فنار"):
        files = []

        # --- الملف الأساسي ---
        if main_items:
            if os.path.exists(MAIN_TEMPLATE):
                try:
                    parsed_date = parse_date(main_items[0]["inv_date"])
                    if parsed_date is None:
                        st.warning(
                            "⚠️ تاريخ الفاتورة الأولى غير صالح (الصيغة المطلوبة"
                            " YYYY-MM-DD)، تم استخدام تاريخ اليوم في عنوان التفريغ."
                        )
                        parsed_date = datetime.today()

                    inv_month_name = get_arabic_month(parsed_date.month)
                    inv_year = parsed_date.year

                    data = build_workbook(
                        MAIN_TEMPLATE,
                        f"تفريغ عهدة : {dar_name} شهر {inv_month_name} {inv_year} م -",
                        main_items,
                        date_format="yyyy-mm-dd",
                    )
                    files.append({
                        "label": "ملف العهدة الأساسي",
                        "file_name": f"تفريغ_عهدة_{dar_name}_2026.xlsx",
                        "data": data,
                    })
                except Exception as e:
                    st.error(f"حدث خطأ أثناء معالجة ملف الإكسل: {e}")
            else:
                st.error(
                    f"لم يتم العثور على ملف النموذج المعتمد: `{MAIN_TEMPLATE}` في مجلد"
                    " التطبيق!"
                )

        # --- الملفات المستقلة (الصيانة وغيرها) ---
        for cat, items in sep_groups.items():
            cfg = SEPARATE_FILES[cat]
            if not os.path.exists(cfg["template"]):
                st.error(
                    f"لم يتم العثور على نموذج {cat}: `{cfg['template']}` في مجلد"
                    " التطبيق!"
                )
                continue
            try:
                title = cfg["title"].format(
                    dar=dar_name, period=period_labels[cat].strip()
                )
                data = build_workbook(cfg["template"], title, items)
                files.append({
                    "label": f"ملف {cat}",
                    "file_name": cfg["file_name"].format(dar=safe_filename(dar_name)),
                    "data": data,
                })
            except Exception as e:
                st.error(f"حدث خطأ أثناء معالجة ملف {cat}: {e}")

        if files:
            st.session_state["fanar_result"] = {
                "fingerprint": fingerprint,
                "dar": dar_name,
                "files": files,
            }
        else:
            st.session_state.pop("fanar_result", None)

    # عرض النتيجة (خارج شرط الزر حتى ما تختفي أزرار التحميل بعد أول تحميل)
    result = st.session_state.get("fanar_result")
    if result and result["fingerprint"] == fingerprint:
        st.success(f"🎉 تم تفريغ البيانات لـ ({result['dar']}) بنجاح!")

        for i, out in enumerate(result["files"]):
            st.download_button(
                label=f"📥 تحميل {out['label']}",
                data=out["data"],
                file_name=out["file_name"],
                mime=XLSX_MIME,
                key=f"dl_{i}",
            )

        if len(result["files"]) > 1:
            zip_buf = io.BytesIO()
            with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
                for out in result["files"]:
                    zf.writestr(out["file_name"], out["data"])
            st.download_button(
                label="📦 تحميل كل الملفات (ZIP)",
                data=zip_buf.getvalue(),
                file_name=f"تفريغ_عهدة_{safe_filename(result['dar'])}.zip",
                mime="application/zip",
                key="dl_zip",
            )

        st.link_button(
            "🌐 الذهاب إلى منصة فنار لإنشاء أمر الصرف",
            "https://fanar.gqb.org.sa/requests/home",
        )
