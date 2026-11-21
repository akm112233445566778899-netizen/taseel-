import base64
import io
import json
import os
import re
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

VAT_RATE = 0.15

# ----------------------------------------------------------------------
# أنواع الفواتير: لإضافة نوع جديد أضف سطر هنا فقط
#   category : بند الحساب (لازم يكون من CATEGORIES)
#   statement: نص البيان ({dar} يتبدل باسم الدار)
#   supplier : اسم المورد الافتراضي
#   tax_no   : الرقم الضريبي الافتراضي
#   electric : True = يستخدم الاستخراج المخصص لفواتير الكهرباء
# ----------------------------------------------------------------------
CATEGORIES = [
    "الكهرباء",
    "المياة",
    "الهاتف والبريد",
    "الصيانة",
    "قرطاسيات ومطبوعات",
    "نظافة",
    "أخرى",
]

INVOICE_TYPES = {
    "فاتورة كهرباء": {
        "category": "الكهرباء",
        "statement": "سداد فاتورة كهرباء - {dar}",
        "supplier": "الشركة السعودية للكهرباء",
        "tax_no": "300000361310003",
        "electric": True,
    },
    "فاتورة مياه": {
        "category": "المياة",
        "statement": "سداد فاتورة مياه - {dar}",
        "supplier": "شركة المياه الوطنية",
        "tax_no": "",
        "electric": False,
    },
    "فاتورة هاتف / إنترنت": {
        "category": "الهاتف والبريد",
        "statement": "سداد فاتورة هاتف وإنترنت - {dar}",
        "supplier": "",
        "tax_no": "",
        "electric": False,
    },
    "فاتورة صيانة": {
        "category": "الصيانة",
        "statement": "سداد فاتورة صيانة - {dar}",
        "supplier": "",
        "tax_no": "",
        "electric": False,
    },
    "قرطاسيات ومطبوعات": {
        "category": "قرطاسيات ومطبوعات",
        "statement": "شراء قرطاسيات ومطبوعات - {dar}",
        "supplier": "",
        "tax_no": "",
        "electric": False,
    },
    "فاتورة نظافة": {
        "category": "نظافة",
        "statement": "سداد فاتورة نظافة - {dar}",
        "supplier": "",
        "tax_no": "",
        "electric": False,
    },
    "أخرى": {
        "category": "أخرى",
        "statement": "سداد فاتورة - {dar}",
        "supplier": "",
        "tax_no": "",
        "electric": False,
    },
}


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


def find_all_dates(full_text):
    """كل التواريخ الميلادية الصالحة في النص، بترتيب ظهورها"""
    result = []
    for d, m, y in re.findall(r"(?<!\d)(\d{2})[-/](\d{2})[-/](\d{4})(?!\d)", full_text):
        try:
            dt = datetime(int(y), int(m), int(d))
            if 2000 <= dt.year <= 2100:
                result.append(dt)
        except ValueError:
            pass
    for y, m, d in re.findall(r"(?<!\d)(\d{4})[-/](\d{2})[-/](\d{2})(?!\d)", full_text):
        try:
            dt = datetime(int(y), int(m), int(d))
            if 2000 <= dt.year <= 2100:
                result.append(dt)
        except ValueError:
            pass
    return result


# ----------------------------------------------------------------------
# استخراج التاريخ (كهرباء)
# ----------------------------------------------------------------------
def extract_invoice_date(full_text):
    """
    يستخرج تاريخ فاتورة الكهرباء بدون الاعتماد على الكلمات العربية
    (لأن خط الـ PDF أحياناً يطلع حروف مشوهة).

    الفاتورة فيها 4 تواريخ:
      بداية الفترة < نهاية الفترة <= تاريخ الفاتورة < تاريخ بداية الاستحقاق
    فتاريخ الفاتورة هو ثاني أكبر تاريخ.
    يرجع نص بصيغة YYYY-MM-DD أو "" إذا ما لقى.
    """
    # الطريقة الأولى (الأدق): تاريخ الفاتورة يكون في نفس سطر رقم الحساب
    for line in full_text.split("\n"):
        if re.search(r"(?<!\d)30\d{9}(?!\d)", line):
            line_dates = find_all_dates(line)
            if len(line_dates) == 1:
                return line_dates[0].strftime("%Y-%m-%d")

    # الطريقة الاحتياطية: ثاني أكبر تاريخ في الفاتورة
    dates = sorted(set(find_all_dates(full_text)))
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
# القراءة المجانية من الصورة (OCR) + استخراج عام
# لفواتير الـ PDF اللي نصها تالف (أرقام معكوسة) أو الصور الممسوحة
# ----------------------------------------------------------------------
# الرقم الضريبي للجمعية (المشتري) لاستبعاده عند البحث عن رقم المورد
BUYER_TAX = "310834779300003"


def ocr_text(uploaded_file):
    """يحوّل الصفحة لصورة ويقرأها بـ Tesseract. يرجع "" إذا الأداة غير متوفرة"""
    try:
        import pytesseract
        from PIL import Image

        uploaded_file.seek(0)
        images = []
        if uploaded_file.name.lower().endswith(".pdf"):
            with pdfplumber.open(uploaded_file) as pdf:
                for page in pdf.pages[:3]:
                    images.append(page.to_image(resolution=300).original.convert("L"))
        else:
            images.append(Image.open(uploaded_file).convert("L"))

        out = ""
        for img in images:
            for psm in (4, 3):  # وضعين يكملون بعض في الجداول
                for lang in ("ara+eng", "eng"):
                    try:
                        out += pytesseract.image_to_string(
                            img, lang=lang, config=f"--psm {psm}"
                        ) + "\n"
                        break
                    except pytesseract.TesseractError:
                        continue
        return out
    except Exception:
        return ""


def ocr_header(uploaded_file):
    """يقرأ أعلى الصفحة بحجم مكبّر (فيه غالباً رقم الفاتورة والتاريخ)"""
    try:
        import pytesseract
        from PIL import Image

        uploaded_file.seek(0)
        if uploaded_file.name.lower().endswith(".pdf"):
            with pdfplumber.open(uploaded_file) as pdf:
                img = pdf.pages[0].to_image(resolution=300).original.convert("L")
        else:
            img = Image.open(uploaded_file).convert("L")
        w, h = img.size
        out = ""
        for a, b in ((0.15, 0.24), (0.17, 0.21)):
            crop = img.crop((0, int(h * a), w, int(h * b)))
            crop = crop.resize((crop.width * 2, crop.height * 2))
            out += pytesseract.image_to_string(crop, lang="eng", config="--psm 4") + "\n"
        return out
    except Exception:
        return ""


def generic_extract(text, text_layer="", header=""):
    """استخراج عام من أي نص فاتورة. القيم الناقصة تبقى فاضية"""
    text = convert_arabic_digits(text)
    res = {"inv_no": "", "inv_date": "", "tax_no": "", "supplier": "",
           "amount": 0.0, "note": ""}

    # المبلغ: الأفضل ثلاثية صافي + 15% = إجمالي
    res["amount"] = extract_net_amount(text)
    if not res["amount"]:
        # احتياطي: أكبر مبلغ غالباً هو الإجمالي بعد الضريبة
        amt_re = r"(?<![\d,.])(?:\d{1,3}(?:,\d{3})+|\d+)\.\d{2}(?!\d)"
        vals = []
        for x in re.findall(amt_re, text):
            try:
                vals.append(to_float(x))
            except ValueError:
                pass
        if vals:
            res["amount"] = round(max(vals) / (1 + VAT_RATE), 2)
            res["note"] = "تم حساب المبلغ قبل الضريبة من الإجمالي (÷1.15)، تأكد منه."

    # التاريخ: أول تاريخ ميلادي (من النص كله ثم من أعلى الصفحة)
    dates = find_all_dates(text) or find_all_dates(convert_arabic_digits(header))
    if dates:
        res["inv_date"] = dates[0].strftime("%Y-%m-%d")

    # الرقم الضريبي للمورد (نستبعد رقم الجمعية)
    for t in re.findall(r"(?<!\d)(3\d{14})(?!\d)", text):
        if t != BUYER_TAX:
            res["tax_no"] = t
            break

    # رقم الفاتورة
    m = re.search(r"Invoice\s*(?:Number|No\.?)\D{0,15}(\d{3,15})", text, re.I)
    if m:
        res["inv_no"] = m.group(1)
    elif header:
        # سطر لوحده فيه أرقام فقط (6-10 خانات) في أعلى الصفحة
        m = re.search(r"^\s*(\d{6,10})\s*$", convert_arabic_digits(header), re.M)
        if m:
            res["inv_no"] = m.group(1)

    # اسم المورد: من نص OCR، وإذا ما لقى من نص الـ PDF بعد عكس السطور
    # (بعض الفواتير تخزّن العربي معكوس)
    sources = text.split("\n") + [ln[::-1] for ln in (text_layer or "").split("\n")]
    for line in sources:
        m = re.search(r"(?:مؤسسة|شركة|مصنع|مكتب|مركز)[^\n\d|:]{3,50}", line)
        if not m:
            continue
        name = re.sub(r"\s*الاسم\s*", " ", m.group(0)).strip()
        if "الجمعية" not in name and "الدور" not in name and len(name) > 6:
            res["supplier"] = name
            break

    return res


# ----------------------------------------------------------------------
# تحليل الفاتورة
# ----------------------------------------------------------------------
def parse_pdf_invoice_robust(uploaded_file, inv_type):
    cfg = INVOICE_TYPES[inv_type]
    data = {
        "inv_no": "",
        "account_no": "",
        "inv_date": "",
        "amount": 0.0,
        "tax_no": cfg["tax_no"],
        "supplier": cfg["supplier"],
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

        if cfg["electric"]:
            # فواتير الكهرباء: استخراج مخصص
            m = re.search(r"(?<!\d)(2[05]\d{10})(?!\d)", full_text)
            if m:
                data["inv_no"] = m.group(1)

            m = re.search(r"(?<!\d)(30\d{9})(?!\d)", full_text)
            if m:
                data["account_no"] = m.group(1)

            m = re.search(r"(?<!\d)(3000\d{11})(?!\d)", full_text)
            if m:
                data["tax_no"] = m.group(1)

            data["inv_date"] = extract_invoice_date(full_text)

            # لو ما لقى شي (نص تالف أو الفاتورة مو كهرباء) نجرب القراءة من الصورة
            if not data["inv_date"] and not extract_net_amount(full_text):
                ocr = ocr_text(uploaded_file)
                if ocr.strip():
                    g = generic_extract(ocr, full_text_raw, ocr_header(uploaded_file))
                    data["raw_text"] = (
                        "--- نص OCR ---\n" + ocr + "\n--- نص PDF ---\n" + full_text_raw
                    )
                    data["inv_no"] = g["inv_no"] or data["inv_no"]
                    data["inv_date"] = g["inv_date"]
                    data["tax_no"] = g["tax_no"]
                    data["supplier"] = g["supplier"]
                    data["amount"] = g["amount"]
                    data["note"] = g["note"]
                    return data
        else:
            # باقي الأنواع: نقرأ الصورة بالـ OCR إذا متوفر (أدق من نص PDF التالف)
            ocr = ocr_text(uploaded_file)
            g = generic_extract(
                ocr if ocr.strip() else full_text,
                full_text_raw,
                ocr_header(uploaded_file) if ocr.strip() else "",
            )
            data["raw_text"] = (
                "--- نص OCR ---\n" + ocr + "\n--- نص PDF ---\n" + full_text_raw
                if ocr.strip() else full_text_raw
            )
            data["inv_no"] = g["inv_no"]
            data["inv_date"] = g["inv_date"]
            data["tax_no"] = g["tax_no"] or cfg["tax_no"]
            data["supplier"] = g["supplier"] or cfg["supplier"]
            data["amount"] = g["amount"]
            data["note"] = g["note"]
            return data

        data["amount"] = extract_net_amount(full_text)

    except Exception as e:
        data["error"] = f"تعذر قراءة الملف: {e}"

    return data


def detect_type_free(uploaded_file):
    """بدون ذكاء اصطناعي: يميّز فاتورة الكهرباء من نمط أرقامها، وغيرها يرجع "أخرى" """
    key = f"dt_{uploaded_file.name}_{uploaded_file.size}"
    if key not in st.session_state:
        t = "أخرى"
        try:
            uploaded_file.seek(0)
            if uploaded_file.name.lower().endswith(".pdf"):
                with pdfplumber.open(uploaded_file) as pdf:
                    txt = convert_arabic_digits(
                        "\n".join((pg.extract_text() or "") for pg in pdf.pages[:2])
                    )
                if re.search(r"(?<!\d)(2[05]\d{10})(?!\d)", txt) and re.search(
                    r"(?<!\d)30\d{9}(?!\d)", txt
                ):
                    t = "فاتورة كهرباء"
        except Exception:
            pass
        st.session_state[key] = t
    return st.session_state[key]


# ----------------------------------------------------------------------
# القراءة بالذكاء الاصطناعي (تقرأ أي فاتورة: PDF أو صورة)
# تحتاج مفتاح ANTHROPIC_API_KEY في Secrets، وإذا ما وُجد يرجع للاستخراج القديم
# ----------------------------------------------------------------------
AI_MODEL = "claude-sonnet-5-5"


def ai_extract(uploaded_file):
    try:
        import anthropic

        api_key = st.secrets["ANTHROPIC_API_KEY"]
    except Exception:
        return {"error": "no_key"}

    try:
        uploaded_file.seek(0)
        raw = uploaded_file.read()
        b64 = base64.standard_b64encode(raw).decode()
        name = uploaded_file.name.lower()

        if name.endswith(".pdf"):
            block = {
                "type": "document",
                "source": {"type": "base64", "media_type": "application/pdf", "data": b64},
            }
        else:
            media = "image/png" if name.endswith(".png") else "image/jpeg"
            block = {
                "type": "image",
                "source": {"type": "base64", "media_type": media, "data": b64},
            }

        types_list = " | ".join(INVOICE_TYPES.keys())
        prompt = f"""اقرأ هذه الفاتورة (سعودية غالباً) وأرجع JSON فقط بدون أي شرح أو علامات ```.
المفاتيح:
- "supplier": اسم المورد/الشركة البائعة
- "inv_no": رقم الفاتورة
- "account_no": رقم الحساب أو العقد أو الاشتراك إن وُجد، وإلا ""
- "inv_date": تاريخ الفاتورة الميلادي بصيغة YYYY-MM-DD (تجاهل التواريخ الهجرية وتواريخ الاستحقاق وفترة الاستهلاك)
- "tax_no": الرقم الضريبي للمورد (15 رقم) إن وُجد، وإلا ""
- "amount": المبلغ قبل ضريبة القيمة المضافة (رقم فقط)
- "total": الإجمالي بعد الضريبة (رقم فقط)
- "type": نوع الفاتورة، اختر واحداً بالضبط من: {types_list}
حوّل الأرقام العربية (٠١٢٣) إلى إنجليزية. إذا لم تجد قيمة اتركها "" أو 0، ولا تخمّن."""

        client = anthropic.Anthropic(api_key=api_key)
        resp = client.messages.create(
            model=AI_MODEL,
            max_tokens=1000,
            messages=[{"role": "user", "content": [block, {"type": "text", "text": prompt}]}],
        )
        text = "".join(b.text for b in resp.content if b.type == "text")
        text = text.replace("```json", "").replace("```", "").strip()
        return json.loads(text)
    except Exception as e:
        return {"error": f"تعذرت القراءة بالذكاء الاصطناعي: {e}"}


def get_ai_data(uploaded_file):
    """يحفظ النتيجة عشان ما تنعاد القراءة (وتنحسب تكلفتها) مع كل تحديث للصفحة"""
    key = f"ai_{uploaded_file.name}_{uploaded_file.size}"
    if key not in st.session_state:
        st.session_state[key] = ai_extract(uploaded_file)
    return st.session_state[key]


CAT_COL_MAP = {
    "الهاتف والبريد": "I",
    "الكهرباء": "J",
    "المياة": "K",
    "الصيانة": "O",
    "قرطاسيات ومطبوعات": "P",
    "نظافة": "R",
    "أخرى": "W",
}


def fill_workbook(wb, dar_name, group):
    """يعبّي ملف إكسل بفواتير مجموعة (تصنيف واحد). يرجع True إذا التاريخ صالح"""
    ws = wb.active
    date_ok = True
    try:
        parsed_date = datetime.strptime(group[0]["inv_date"].strip(), "%Y-%m-%d")
    except Exception:
        parsed_date = datetime.today()
        date_ok = False

    ws["A2"] = (
        f"تفريغ عهدة : {dar_name} شهر {get_arabic_month(parsed_date.month)}"
        f" {parsed_date.year} م -"
    )

    start_row = 6
    for idx, item in enumerate(group):
        r = start_row + idx
        ws[f"C{r}"] = item["statement"]
        ws[f"D{r}"] = item["account_no"]

        raw_inv = str(item["inv_no"]).strip()
        ws[f"E{r}"] = int(raw_inv) if raw_inv.isdigit() else raw_inv

        try:
            ws[f"F{r}"] = datetime.strptime(item["inv_date"].strip(), "%Y-%m-%d")
            ws[f"F{r}"].number_format = "yyyy-mm-dd"
        except Exception:
            ws[f"F{r}"] = item["inv_date"]

        ws[f"G{r}"] = item["tax_no"]
        ws[f"G{r}"].number_format = "@"
        ws[f"H{r}"] = item["supplier"]

        col_letter = CAT_COL_MAP.get(item["category"], "J")
        ws[f"{col_letter}{r}"] = float(item["amount"])
    return date_ok


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

st.subheader("اختر الدار")
dar_name = st.selectbox("اختر اسم الدار المعنية", dar_list)

st.subheader("1️⃣ رفع الفواتير")
uploaded_files = st.file_uploader(
    "التقط صور الفواتير أو اسحب ملفات الـ PDF هنا",
    type=["pdf", "png", "jpg", "jpeg"],
    accept_multiple_files=True,
)

if uploaded_files:
    st.success(f"✅ تم استلام {len(uploaded_files)} فاتورة بنجاح!")
    st.divider()

    st.subheader("2️⃣ مراجعة وتحديد البيانات")

    invoices_data = []
    type_names = list(INVOICE_TYPES.keys())

    for idx, file in enumerate(uploaded_files):
        st.markdown(f"### 📄 الفاتورة رقم ({idx + 1}): `{file.name}`")

        with st.spinner(f"🔍 جاري قراءة الفاتورة #{idx + 1}..."):
            ai = get_ai_data(file)
        ai_ok = bool(ai) and not ai.get("error")

        # النوع يُقترح تلقائياً من الذكاء الاصطناعي وتقدر تغيره
        detected = ai.get("type") if ai_ok else None
        if detected not in INVOICE_TYPES:
            detected = detect_type_free(file) if not ai_ok else "أخرى"

        inv_type = st.selectbox(
            f"نوع الفاتورة #{idx+1}",
            type_names,
            index=type_names.index(detected),
            key=f"type_{idx}_{file.name}",
        )
        cfg = INVOICE_TYPES[inv_type]

        if ai_ok:
            try:
                amt = float(ai.get("amount") or 0)
                tot = float(ai.get("total") or 0)
            except (TypeError, ValueError):
                amt, tot = 0.0, 0.0
            ext_data = {
                "supplier": ai.get("supplier") or cfg["supplier"],
                "inv_no": ai.get("inv_no") or "",
                "account_no": ai.get("account_no") or "",
                "inv_date": ai.get("inv_date") or "",
                "tax_no": ai.get("tax_no") or cfg["tax_no"],
                "amount": amt,
                "raw_text": json.dumps(ai, ensure_ascii=False, indent=2),
                "error": "",
            }
            if amt and tot and abs(round(amt * (1 + VAT_RATE), 2) - tot) > 0.05:
                st.warning("⚠️ المبلغ + الضريبة 15% لا يطابق الإجمالي في الفاتورة، راجع المبلغ.")
        else:
            if ai and ai.get("error") not in (None, "no_key"):
                st.warning(f"⚠️ {ai['error']} (تم استخدام الاستخراج العادي)")
            ext_data = parse_pdf_invoice_robust(file, inv_type)
            if ext_data.get("note"):
                st.info(f"ℹ️ {ext_data['note']}")

        if ext_data.get("error"):
            st.warning(f"⚠️ {ext_data['error']}")
        else:
            if not ext_data["inv_date"]:
                st.warning("⚠️ لم يتم العثور على تاريخ الفاتورة، أدخله يدوياً.")
            if not ext_data["amount"]:
                st.warning("⚠️ لم يتم العثور على المبلغ، أدخله يدوياً.")

        with st.expander(f"عرض النص المستخرج للفاتورة #{idx + 1} (للفحص)"):
            st.text(ext_data.get("raw_text", ""))

        # المفاتيح تحتوي على نوع الفاتورة عشان القيم الافتراضية تتحدث عند تغيير النوع
        col1, col2 = st.columns(2)
        with col1:
            supplier = st.text_input(
                f"اسم المورد #{idx+1}",
                value=ext_data.get("supplier", ""),
                key=f"sup_{idx}_{inv_type}",
            )
            inv_no = st.text_input(
                f"رقم الفاتورة #{idx+1}",
                value=str(ext_data.get("inv_no", "")),
                key=f"inv_{idx}_{inv_type}",
            )
            account_no = st.text_input(
                f"رقم الحساب #{idx+1}",
                value=str(ext_data.get("account_no", "")),
                key=f"acc_{idx}_{inv_type}",
            )
            inv_date = st.text_input(
                f"تاريخ الفاتورة #{idx+1} (YYYY-MM-DD)",
                value=str(ext_data.get("inv_date", "")),
                key=f"date_{idx}_{inv_type}",
            )

        with col2:
            tax_no = st.text_input(
                f"الرقم الضريبي #{idx+1}",
                value=str(ext_data.get("tax_no", "")),
                key=f"tax_{idx}_{inv_type}",
            )
            amount = st.number_input(
                f"المبلغ بدون الضريبة #{idx+1}",
                value=float(ext_data.get("amount", 0.0) or 0.0),
                step=0.01,
                key=f"amt_{idx}_{inv_type}",
            )

            statement = st.text_input(
                f"البيان #{idx+1}",
                value=cfg["statement"].format(dar=dar_name),
                key=f"st_{idx}_{dar_name}_{inv_type}",
            )

            category = st.selectbox(
                f"بند الحساب #{idx+1}",
                CATEGORIES,
                index=CATEGORIES.index(cfg["category"]),
                key=f"cat_{idx}_{inv_type}",
            )

            estimated_total = round(amount + round(amount * VAT_RATE, 2), 2)
            written_amount = tafqeet_arabic(estimated_total)
            written_text = st.text_input(
                f"المبلغ كتابةً (تفقيط) #{idx+1}",
                value=written_amount,
                key=f"writ_{idx}_{amount}",
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
            "written_amount": written_text,
        })
        st.markdown("---")

    st.subheader("3️⃣ المعالجة والإرسال")

    # تصنيفات الفواتير المرفوعة (بالترتيب)
    cats_used = []
    for item in invoices_data:
        if item["category"] not in cats_used:
            cats_used.append(item["category"])

    st.caption(
        "اختر ملف الإكسل لكل تصنيف. إذا ما اخترت، يُستخدم النموذج الافتراضي."
    )
    template_uploads = {}
    for cat in cats_used:
        n = sum(1 for it in invoices_data if it["category"] == cat)
        template_uploads[cat] = st.file_uploader(
            f"📂 ملف الإكسل لتصنيف «{cat}» ({n} فاتورة)",
            type=["xlsx"],
            key=f"tpl_{cat}",
        )

    if st.button("🚀 تفريغ العهدة والتجهيز لمنصة فنار"):
        default_path = "نموذج عهدة الادارة العامة 2026.xlsx.xlsx"
        outputs = []

        for cat in cats_used:
            group = [it for it in invoices_data if it["category"] == cat]
            up = template_uploads.get(cat)
            try:
                if up is not None:
                    up.seek(0)
                    wb = openpyxl.load_workbook(io.BytesIO(up.read()))
                elif os.path.exists(default_path):
                    wb = openpyxl.load_workbook(default_path)
                else:
                    st.error(
                        f"تصنيف «{cat}»: ما اخترت ملف إكسل، وملف النموذج الافتراضي"
                        f" `{default_path}` غير موجود في مجلد التطبيق."
                    )
                    continue

                if not fill_workbook(wb, dar_name, group):
                    st.warning(
                        f"⚠️ تصنيف «{cat}»: تاريخ الفاتورة الأولى غير صالح (الصيغة"
                        " YYYY-MM-DD)، تم استخدام تاريخ اليوم في العنوان."
                    )

                buf = io.BytesIO()
                wb.save(buf)
                suffix = f"_{cat}" if len(cats_used) > 1 else ""
                outputs.append(
                    (f"تفريغ_عهدة_{dar_name}{suffix}_2026.xlsx", buf.getvalue())
                )
            except Exception as e:
                st.error(f"حدث خطأ أثناء معالجة ملف تصنيف «{cat}»: {e}")

        # نحفظها في session_state عشان تبقى أزرار التحميل بعد أي تحديث للصفحة
        st.session_state["outputs"] = outputs

    outputs = st.session_state.get("outputs", [])
    if outputs:
        st.success(f"🎉 تم تجهيز {len(outputs)} ملف!")
        for i, (fname, data) in enumerate(outputs):
            st.download_button(
                label=f"📥 تحميل {fname}",
                data=data,
                file_name=fname,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key=f"dl_{i}",
            )
        st.link_button(
            "🌐 الذهاب إلى منصة فنار لإنشاء أمر الصرف",
            "https://fanar.gqb.org.sa/requests/home",
        )
