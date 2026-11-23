import base64
import io
import zlib
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
#   statement: (غير مستخدم، البيان صار «سداد فاتورة» + اسم المورد)
#   supplier : اسم المورد الافتراضي
#   tax_no   : الرقم الضريبي الافتراضي
#   electric : True = يستخدم الاستخراج المخصص لفواتير الكهرباء
# ----------------------------------------------------------------------
CATEGORIES = [
    "الهاتف والبريد",
    "الكهرباء",
    "المياة",
    "الأصول",
    "اثاث وتجهيزات مكتبية",
    "الحاسب الالي",
    "صيانة مباني",
    "قرطاسيات ومطبوعات",
    "دعاية وإعلان",
    "نظافة",
    "ضيافة و بوفية",
    "هدايا",
    "المحروقات",
    "صيانة سيارات",
    "أخرى",
]


def _t(category, statement, supplier="", tax_no="", electric=False):
    return {"category": category, "statement": statement + " - {dar}",
            "supplier": supplier, "tax_no": tax_no, "electric": electric}


INVOICE_TYPES = {
    "فاتورة كهرباء": _t("الكهرباء", "سداد فاتورة كهرباء",
                        "الشركة السعودية للكهرباء", "300000361310003", True),
    "فاتورة مياه": _t("المياة", "سداد فاتورة مياه", "شركة المياه الوطنية"),
    "فاتورة هاتف / إنترنت": _t("الهاتف والبريد", "سداد فاتورة هاتف وإنترنت"),
    "فاتورة صيانة": _t("صيانة مباني", "سداد فاتورة صيانة"),
    "صيانة سيارات": _t("صيانة سيارات", "سداد فاتورة صيانة سيارات"),
    "قرطاسيات ومطبوعات": _t("قرطاسيات ومطبوعات", "شراء قرطاسيات ومطبوعات"),
    "فاتورة نظافة": _t("نظافة", "سداد فاتورة نظافة"),
    "ضيافة وبوفية": _t("ضيافة و بوفية", "شراء ضيافة وبوفية"),
    "هدايا": _t("هدايا", "شراء هدايا"),
    "دعاية وإعلان": _t("دعاية وإعلان", "سداد فاتورة دعاية وإعلان"),
    "محروقات": _t("المحروقات", "شراء محروقات"),
    "أثاث وتجهيزات مكتبية": _t("اثاث وتجهيزات مكتبية", "شراء أثاث وتجهيزات مكتبية"),
    "حاسب آلي وملحقاته": _t("الحاسب الالي", "شراء حاسب آلي وملحقاته"),
    "أصول": _t("الأصول", "شراء أصول"),
    "أخرى": _t("أخرى", "سداد فاتورة"),
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


MONTH_NAMES = {
    "يناير": 1, "فبراير": 2, "مارس": 3, "أبريل": 4, "ابريل": 4, "مايو": 5,
    "يونيو": 6, "يوليو": 7, "أغسطس": 8, "اغسطس": 8, "سبتمبر": 9,
    "أكتوبر": 10, "اكتوبر": 10, "نوفمبر": 11, "ديسمبر": 12,
    "jan": 1, "feb": 2, "mar": 3, "apr": 4, "may": 5, "jun": 6,
    "jul": 7, "aug": 8, "sep": 9, "oct": 10, "nov": 11, "dec": 12,
}


def find_all_dates(full_text):
    """كل التواريخ الميلادية الصالحة في النص، بترتيب ظهورها"""
    result = []
    full_text = re.sub(r"(?<=\d)\s*/\s*(?=\d)", "/", full_text)

    def add(y, m, d):
        try:
            dt = datetime(int(y), int(m), int(d))
            if 2000 <= dt.year <= 2100:
                result.append(dt)
        except ValueError:
            pass

    # DD-MM-YYYY أو DD/MM/YYYY أو DD.MM.YYYY
    for d, m, y in re.findall(r"(?<!\d)(\d{1,2})[-/.](\d{1,2})[-/.](\d{4})(?!\d)", full_text):
        add(y, m, d)
    # YYYY-MM-DD أو YYYY/MM/DD أو YYYY.MM.DD
    for y, m, d in re.findall(r"(?<!\d)(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})(?!\d)", full_text):
        add(y, m, d)
    # 24 أكتوبر 2026 / 24-Oct-2026
    names = "|".join(sorted(MONTH_NAMES, key=len, reverse=True))
    for d, name, y in re.findall(
        rf"(?<!\d)(\d{{1,2}})\s*[-/ ]?\s*({names})\w*\.?\s*[-/,]?\s*(\d{{4}})(?!\d)",
        full_text, re.I,
    ):
        add(y, MONTH_NAMES[name.lower()], d)
    # Oct 24, 2026
    for name, d, y in re.findall(
        rf"(?<![\w])({names})\w*\.?\s*(\d{{1,2}}),?\s*(\d{{4}})(?!\d)",
        full_text, re.I,
    ):
        add(y, MONTH_NAMES[name.lower()], d)
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


# موردين معروفين: الرقم الضريبي -> (الاسم، نوع الفاتورة). أضف موردينك هنا
KNOWN_SUPPLIERS = {
    "300000157210003": ("شركة الاتصالات السعودية", "فاتورة هاتف / إنترنت"),
    "300000361310003": ("الشركة السعودية للكهرباء", "فاتورة كهرباء"),
}


def decode_cid(text):
    """بعض الـ PDF (مثل فواتير STC) تخزّن الأرقام بأكواد (cid:NN). هنا نفك الأرقام منها"""
    if not text or "(cid:" not in text:
        return text or ""

    def r(m):
        n = int(m.group(1))
        if 19 <= n <= 28:
            return str(n - 19)
        return {15: ".", 16: "-", 17: ".", 18: "/", 3: " "}.get(n, "\u2022")

    out = re.sub(r"\(cid:(\d+)\)", r, text)
    return re.sub(r"\u2022+", " ", out)


def normalize_decimals(text):
    """الفاصلة العشرية (409,00) -> نقطة (409.00). ما تلمس الآلاف (1,000)"""
    return re.sub(r"(?<=\d),(?=\d{2}(?!\d))", ".", text or "")


BIZ_KEYS = r"(?:مؤسسة|شركة|مصنع|مكتب|مركز|معرض|ورشة|مطبعة|مكتبة|محلات|محل|للتجارة|للمقاولات|لحلول)"
NOT_SUPPLIER = ("الجمعية", "جمعية", "الدور", "العميل", "فاتورة", "ضريبية", "تقرير", "الضمان", "الملاحظات")
CUR = r"(?:\bSAR\b|\bSR\b|ر\.\s?س|ريال)"


def generic_extract(text, text_layer="", header=""):
    """استخراج عام من أي نص فاتورة. القيم الناقصة تبقى فاضية"""
    text = convert_arabic_digits(text)
    layer = normalize_decimals(convert_arabic_digits(decode_cid(text_layer or "")))
    res = {"inv_no": "", "account_no": "", "inv_date": "", "tax_no": "",
           "supplier": "", "amount": 0.0, "note": "", "vat_free": False}

    # مصادر النص بالترتيب: OCR، أعلى الصفحة، نص الـ PDF معكوس، نص الـ PDF
    # (عكس السطر يقلب الأرقام أيضاً، فنرجّع كل رقم لوضعه الصحيح)
    rev_layer = "\n".join(
        re.sub(r"[0-9][0-9./:\-]*", lambda m: m.group(0)[::-1], ln[::-1])
        for ln in (text_layer or "").split("\n")
    )
    srcs = [text, convert_arabic_digits(header), convert_arabic_digits(rev_layer), layer]

    # الرقم الضريبي للمورد (نستبعد رقم الجمعية)
    for t in re.findall(r"(?<!\d)(3\d{14})(?!\d)", text + "\n" + layer):
        if t != BUYER_TAX:
            res["tax_no"] = t
            break
    if not res["tax_no"] and re.search(r"stc\.com\.sa", text + layer, re.I):
        res["tax_no"] = "300000157210003"
    known = KNOWN_SUPPLIERS.get(res["tax_no"])

    # المبلغ قبل الضريبة
    # 1) ثلاثية متطابقة: صافي + 15% = إجمالي
    amount = extract_net_amount(text) or extract_net_amount(layer)
    if not amount:
        # 2) مبالغ مكتوبة بجانب عملة (SAR 110 / 110 ريال / ر.س)، بدون اشتراط كسور
        cur_vals = []
        for src in (text, layer):
            for a, b in re.findall(rf"{CUR}\s*:?\s*(\d[\d,]*(?:\.\d+)?)|(\d[\d,]*(?:\.\d+)?)\s*{CUR}", src, re.I):
                try:
                    cur_vals.append(to_float(a or b))
                except ValueError:
                    pass
        # 3) أكبر مبلغ بكسور
        dec_vals = []
        amt_re = r"(?<![\d,.])(?:\d{1,3}(?:,\d{3})+|\d+)\.\d{2}(?!\d)"
        for x in re.findall(amt_re, text):
            try:
                dec_vals.append(to_float(x))
            except ValueError:
                pass
        total = max(cur_vals) if cur_vals else (max(dec_vals) if dec_vals else 0.0)
        if total:
            if res["tax_no"]:
                amount = round(total / (1 + VAT_RATE), 2)
                res["note"] = "تم حساب المبلغ قبل الضريبة من الإجمالي (÷1.15)، تأكد منه."
            else:
                amount = total
                res["vat_free"] = True
                res["note"] = "الفاتورة بدون رقم ضريبي: اعتُمد الإجمالي كمبلغ بدون ضريبة، تأكد منه."
    res["amount"] = amount

    # التاريخ: أول مصدر فيه تاريخ ميلادي
    for src in srcs:
        dates = find_all_dates(src)
        if dates:
            res["inv_date"] = dates[0].strftime("%Y-%m-%d")
            break

    # رقم الفاتورة: الأدق أولاً (OCR ثم أعلى الصفحة)، وبعدها نص الـ PDF (قد يكون تالف)
    inv_patterns = [
        r"Invoice\s*(?:Number|No\.?|#)\D{0,15}(\d{3,15})",
        r"(?:رقم\s*الفاتورة|فاتورة\s*رقم|رقم\s*الفاتوره)\D{0,15}(\d{1,15})",
    ]

    def find_inv(sources):
        for src in sources:
            for pat in inv_patterns:
                m = re.search(pat, src, re.I)
                if m:
                    return m.group(1)
        return ""

    res["inv_no"] = find_inv(srcs[:2])
    if not res["inv_no"] and header:
        m = re.search(r"^\s*(\d{6,10})\s*$", convert_arabic_digits(header), re.M)
        if m:
            res["inv_no"] = m.group(1)
    if not res["inv_no"]:
        # رقم تسلسلي بصيغة حساب-رقم مثل 634205470-182 (فواتير الاتصالات)
        m = re.search(r"(?<!\d)(\d{9})\s*-\s*(\d{1,4})(?!\d)", layer)
        if m:
            res["inv_no"], res["account_no"] = f"{m.group(1)}-{m.group(2)}", m.group(1)
        else:
            m = re.search(r"(?<!\d)(\d{1,4})\s*-\s*(\d{9})(?!\d)", layer)
            if m:
                res["inv_no"], res["account_no"] = f"{m.group(2)}-{m.group(1)}", m.group(2)
    if not res["inv_no"]:
        res["inv_no"] = find_inv(srcs[2:])

    # اسم المورد
    def clean_name(n):
        return re.sub(r"\s*الاسم\s*", " ", n).strip()

    def ok_name(n):
        return len(n) > 6 and not any(w in n for w in NOT_SUPPLIER)

    if known:
        res["supplier"] = known[0]
        return res
    # نص الـ PDF أدق من OCR إذا كان سليم، فنجرّبه أول (معكوس ثم عادي)، وبعدها OCR
    lines = ([ln[::-1] for ln in (text_layer or "").split("\n")]
             + (text_layer or "").split("\n") + text.split("\n"))
    # أ) سطر فيه كلمة نشاط (مؤسسة/شركة/لحلول...): نأخذ السطر كامل إذا قصير ونظيف
    for line in lines:
        mk = re.search(BIZ_KEYS, line)
        if not mk:
            continue
        if len(line.split()) <= 8 and not re.search(r"[\d:|]", line):
            cand = clean_name(line)
        else:
            sl = re.search(BIZ_KEYS + r"[^\n\d|:]{3,50}", line)
            cand = clean_name(sl.group(0)) if sl else ""
        if cand and ok_name(cand):
            res["supplier"] = cand
            break
    # ب) احتياطي: أول سطر عربي نظيف في أعلى نص الـ PDF (غالباً اسم المنشأة)
    if not res["supplier"]:
        for line in (text_layer or "").split("\n")[:4]:
            ln = line.strip()
            if re.fullmatch(r"[\u0600-\u06FF\s]+", ln) and 2 <= len(ln.split()) <= 8 and ok_name(ln):
                res["supplier"] = ln
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
                    data["account_no"] = g["account_no"] or data["account_no"]
                    data["inv_date"] = g["inv_date"]
                    data["tax_no"] = g["tax_no"]
                    data["supplier"] = g["supplier"]
                    data["amount"] = g["amount"]
                    data["note"] = g["note"]
                    data["vat_free"] = g["vat_free"]
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
            data["account_no"] = g["account_no"]
            data["inv_date"] = g["inv_date"]
            data["tax_no"] = g["tax_no"] or cfg["tax_no"]
            data["supplier"] = g["supplier"] or cfg["supplier"]
            data["amount"] = g["amount"]
            data["note"] = g["note"]
            data["vat_free"] = g["vat_free"]
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
                        "\n".join((pg.extract_text() or "") for pg in pdf.pages[:4])
                    )
                dec = convert_arabic_digits(decode_cid(txt))
                hit = next((v[1] for k, v in KNOWN_SUPPLIERS.items() if k in dec), None)
                if hit:
                    t = hit
                elif re.search(r"(?<!\d)(2[05]\d{10})(?!\d)", txt) and re.search(
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


# تخطيط النموذج القديم (احتياطي إذا ما انلقى صف عناوين في الملف)
OLD_LAYOUT = {
    "start": 6,
    "fields": {"statement": "C", "account_no": "D", "inv_no": "E",
               "inv_date": "F", "tax_no": "G", "supplier": "H"},
    "cats": {"الهاتف والبريد": "I", "الكهرباء": "J", "المياة": "K",
             "صيانة مباني": "O", "قرطاسيات ومطبوعات": "P", "نظافة": "R", "أخرى": "W"},
    "tax_col": None,
}

# أسماء بديلة لعناوين الأعمدة
HEADER_ALIASES = {"الصيانة": "صيانة مباني", "المياه": "المياة"}


def norm(t):
    """توحيد النص العربي للمقارنة (مسافات، همزات، ة/ه)"""
    t = re.sub(r"\s+", "", str(t or ""))
    for ch in "أإآ":
        t = t.replace(ch, "ا")
    return t.replace("ة", "ه").replace("ى", "ي")


def detect_layout(ws):
    """يقرأ صف العناوين (اللي فيه «البيان») ويحدد عمود كل حقل وكل بند"""
    for hr in range(1, 15):
        row_map = {}
        for c in ws[hr]:
            if c.value is not None:
                row_map[norm(c.value)] = c.column_letter
        if "البيان" not in row_map:
            continue
        lay = {"start": hr + 1, "fields": {}, "cats": {}, "tax_col": None}
        for key, name in (("statement", "البيان"), ("inv_no", "رقمالفاتوره"),
                          ("inv_date", "التاريخ"), ("tax_no", "الرقمالضريبي"),
                          ("supplier", "اسمالمورد"), ("account_no", "رقمالحساب")):
            if name in row_map:
                lay["fields"][key] = row_map[name]
        for cat in CATEGORIES:
            if norm(cat) in row_map:
                lay["cats"][cat] = row_map[norm(cat)]
        for alias, cat in HEADER_ALIASES.items():
            if norm(alias) in row_map and cat not in lay["cats"]:
                lay["cats"][cat] = row_map[norm(alias)]
        for k, v in row_map.items():
            if "ضريبه" in k and "مشتريات" in k:
                lay["tax_col"] = v
        return lay
    return None


def _data_end_row(ws, lay):
    """آخر صف بيانات: الصفوف المرقّمة في عمود A، وإلا آخر صف في الورقة"""
    start = lay["start"]
    if isinstance(ws[f"A{start}"].value, (int, float)):
        r = start
        while isinstance(ws[f"A{r + 1}"].value, (int, float)):
            r += 1
        return r
    return ws.max_row


def fill_workbook(wb, dar_name, group, kind="تصفية"):
    """يضيف فواتير تصنيف واحد بعد البنود الموجودة في الملف. يرجع True إذا التاريخ صالح"""
    ws = wb.active
    lay = detect_layout(ws) or OLD_LAYOUT
    fields, cats = lay["fields"], lay["cats"]
    end = _data_end_row(ws, lay)
    check_cols = list(fields.values()) + list(cats.values())

    def is_empty(r):
        return all(ws[f"{c}{r}"].value in (None, "") for c in check_cols)

    date_ok = True
    try:
        parsed = datetime.strptime(group[0]["inv_date"].strip(), "%Y-%m-%d")
    except Exception:
        parsed = datetime.today()
        date_ok = False
    month_name = get_arabic_month(parsed.month)

    # العنوان (A2): النموذج الجديد له صيغة، والقديم صيغة ثانية
    a2 = str(ws["A2"].value or "")
    if "عهدة دار" in a2 or "طلب - تصفية" in a2:
        ws["A2"] = f"عهدة دار ({dar_name}) : ({kind}) شهر ({month_name}) {parsed.year} م - "
    else:
        ws["A2"] = f"تفريغ عهدة : {dar_name} شهر {month_name} {parsed.year} م -"

    row = lay["start"]
    for item in group:
        while row <= end and not is_empty(row):
            row += 1
        if row > end:
            raise ValueError(
                f"الملف ما فيه صفوف فاضية كافية (آخر صف بيانات {end}). "
                "استخدم ملف ثاني أو قسّم الفواتير."
            )

        statement = item["statement"]
        # النموذج الجديد ما فيه عمود رقم الحساب، فنضيفه للبيان عشان ما يضيع
        if item["account_no"] and "account_no" not in fields:
            statement = f"{statement} - حساب رقم {item['account_no']}"

        def put(key, value):
            col = fields.get(key)
            if col:
                ws[f"{col}{row}"] = value
            return col

        put("statement", statement)
        put("account_no", item["account_no"])

        raw_inv = str(item["inv_no"]).strip()
        put("inv_no", int(raw_inv) if raw_inv.isdigit() else raw_inv)

        col = fields.get("inv_date")
        if col:
            cell = ws[f"{col}{row}"]
            try:
                cell.value = datetime.strptime(item["inv_date"].strip(), "%Y-%m-%d")
                if cell.number_format == "General":
                    cell.number_format = "yyyy-mm-dd"
            except Exception:
                cell.value = item["inv_date"]

        col = fields.get("tax_no")
        if col:
            cell = ws[f"{col}{row}"]
            cell.value = str(item["tax_no"])
            if cell.number_format == "General":
                cell.number_format = "@"

        put("supplier", item["supplier"])

        cat_col = cats.get(item["category"]) or cats.get("أخرى")
        if not cat_col:
            raise ValueError(f"ما لقيت عمود لبند «{item['category']}» في الملف.")
        ws[f"{cat_col}{row}"] = float(item["amount"])

        # عمود الضريبة في النموذج معادلة (15% دائماً)، فللفاتورة بدون ضريبة نصفّره
        if item.get("vat_free") and lay.get("tax_col"):
            ws[f"{lay['tax_col']}{row}"] = 0

        row += 1
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
        fk = f"{idx}_{file.size}_{zlib.crc32(file.name.encode('utf-8'))}"

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
                key=f"sup_{fk}_{inv_type}",
            )
            inv_no = st.text_input(
                f"رقم الفاتورة #{idx+1}",
                value=str(ext_data.get("inv_no", "")),
                key=f"inv_{fk}_{inv_type}",
            )
            account_no = st.text_input(
                f"رقم الحساب #{idx+1}",
                value=str(ext_data.get("account_no", "")),
                key=f"acc_{fk}_{inv_type}",
            )
            inv_date = st.text_input(
                f"تاريخ الفاتورة #{idx+1} (YYYY-MM-DD)",
                value=str(ext_data.get("inv_date", "")),
                key=f"date_{fk}_{inv_type}",
            )

        with col2:
            tax_no = st.text_input(
                f"الرقم الضريبي #{idx+1}",
                value=str(ext_data.get("tax_no", "")),
                key=f"tax_{fk}_{inv_type}",
            )
            amount = st.number_input(
                f"المبلغ بدون الضريبة #{idx+1}",
                value=float(ext_data.get("amount", 0.0) or 0.0),
                step=0.01,
                key=f"amt_{fk}_{inv_type}",
            )

            # البيان: "سداد فاتورة" + اسم المورد (يتحدث لما تغيّر اسم المورد)
            statement = st.text_input(
                f"البيان #{idx+1}",
                value=f"سداد فاتورة {supplier.strip()}".strip(),
                key=f"st_{fk}_{zlib.crc32(supplier.strip().encode('utf-8'))}",
            )

            category = st.selectbox(
                f"بند الحساب #{idx+1}",
                CATEGORIES,
                index=CATEGORIES.index(cfg["category"]),
                key=f"cat_{fk}_{inv_type}",
            )

            vat_free = st.checkbox(
                f"الفاتورة بدون ضريبة #{idx+1}",
                value=bool(ext_data.get("vat_free", False)),
                key=f"vf_{fk}_{inv_type}",
            )
            vat_part = 0.0 if vat_free else round(amount * VAT_RATE, 2)
            estimated_total = round(amount + vat_part, 2)
            if estimated_total > 1500:
                st.warning("⚠️ الفاتورة تزيد عن 1500 ريال، لازم ترفق عروض أسعار.")
            written_amount = tafqeet_arabic(estimated_total)
            written_text = st.text_input(
                f"المبلغ كتابةً (تفقيط) #{idx+1}",
                value=written_amount,
                key=f"writ_{fk}_{amount}_{vat_free}",
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
            "vat_free": vat_free,
        })
        st.markdown("---")

    st.subheader("3️⃣ المعالجة والإرسال")

    # تصنيفات الفواتير المرفوعة (بالترتيب)
    cats_used = []
    for item in invoices_data:
        if item["category"] not in cats_used:
            cats_used.append(item["category"])

    # ملفات الإكسل الموجودة في مجلد التطبيق (نفس مكان app.py على GitHub)
    DEFAULT_TEMPLATE = "نموذج عهدة الادارة العامة 2026.xlsx.xlsx"
    xlsx_files = sorted(
        f for f in os.listdir(".")
        if f.lower().endswith(".xlsx") and not f.startswith("~$")
    )
    if not xlsx_files:
        st.error("ما فيه أي ملف إكسل في مجلد التطبيق. أضف النموذج على GitHub بجانب app.py.")

    # الأفضل نموذج «طلب صرف عهدة / تصفية عهدة» إذا موجود، وإلا النموذج القديم
    default_idx = 0
    for i, f in enumerate(xlsx_files):
        if "طلب" in f and "صرف" in f:
            default_idx = i
            break
    else:
        if DEFAULT_TEMPLATE in xlsx_files:
            default_idx = xlsx_files.index(DEFAULT_TEMPLATE)

    kind = st.radio("نوع الطلب (للنموذج الجديد)", ["تصفية", "طلب"], horizontal=True)

    st.caption("اختر ملف الإكسل لكل تصنيف من الملفات الموجودة في التطبيق.")
    chosen_files = {}
    for cat in cats_used:
        n = sum(1 for it in invoices_data if it["category"] == cat)
        if xlsx_files:
            chosen_files[cat] = st.selectbox(
                f"📂 ملف الإكسل لتصنيف «{cat}» ({n} فاتورة)",
                xlsx_files,
                index=default_idx,
                key=f"tpl_{cat}",
            )

    if st.button("🚀 تفريغ العهدة والتجهيز لمنصة فنار"):
        outputs = []

        for cat in cats_used:
            group = [it for it in invoices_data if it["category"] == cat]
            path = chosen_files.get(cat)
            try:
                if not path or not os.path.exists(path):
                    st.error(f"تصنيف «{cat}»: ملف الإكسل المختار غير موجود في مجلد التطبيق.")
                    continue
                wb = openpyxl.load_workbook(path)

                if not fill_workbook(wb, dar_name, group, kind):
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
