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
    page_title="نظام تسهيل | 2026",
    page_icon="⚡",
    layout="centered"
)

if os.path.exists("logo.png"):
    c1, c2 = st.columns([1, 3.5])
    with c1:
        st.image("logo.png", width=120)
    with c2:
        st.title("⚡ نظام تسهيل")
        st.subheader("تفريغ العهود وأوامر الصرف")
        st.caption("تحفيظ القرآن ببريدة - 2026")
else:
    st.title("⚡ نظام تسهيل")
    st.subheader("تفريغ العهود وأوامر الصرف")
    st.caption("تحفيظ القرآن ببريدة - 2026")

st.divider()

if "parsed_cache" not in st.session_state:
    st.session_state.parsed_cache = {}

API_KEY = st.secrets.get("GEMINI_API_KEY", "").strip()

def tafqeet_arabic(amount):
    try:
        if amount <= 0:
            return ""
        riyal = int(amount)
        halalas = int(round((amount - riyal) * 100))
        t = f"{riyal} ريال سعودي"
        if halalas > 0:
            t += f" و {halalas} هللة"
        return t + " فقط لا غير"
    except Exception:
        return ""

def get_arabic_month(m):
    months = {
        1: "يناير", 2: "فبراير", 3: "مارس",
        4: "أبريل", 5: "مايو", 6: "يونيو",
        7: "يوليو", 8: "أغسطس", 9: "سبتمبر",
        10: "أكتوبر", 11: "نوفمبر", 12: "ديسمبر"
    }
    return months.get(m, "")

def get_file_hash(file_bytes):
    return hashlib.md5(file_bytes).hexdigest()

def extract_any_invoice(file_bytes, mime_type):
    if not API_KEY:
        return {"error": "مفتاح GEMINI_API_KEY غير مضاف"}

    try:
        b64_data = base64.b64encode(file_bytes).decode("utf-8")

        prompt = (
            "أنت محاسب سعودي. استخرج من الفاتورة بصيغة JSON فقط:\n"
            "1. supplier: اسم المورد أو المحل.\n"
            "2. tax_no: الرقم الضريبي (15 رقم) أو فارغ.\n"
            "3. inv_no: رقم الفاتورة.\n"
            "4. account_no: رقم الحساب إن وجد أو فارغ.\n"
            "5. inv_date: تاريخ الفاتورة YYYY-MM-DD.\n"
            "6. amount_pre_tax: المبلغ قبل الضريبة (اقسم الإجمالي على 1.15 إذا كان شامل الضريبة).\n"
            "7. category: اختر من [الكهرباء, المياة, الهاتف والبريد, الصيانة, قرطاسيات ومطبوعات, نظافة, أخرى].\n"
            "8. description: وصف مختصر للفاتورة.\n"
            '{"supplier":"","tax_no":"","inv_no":"","account_no":"",'
            '"inv_date":"YYYY-MM-DD","amount_pre_tax":0.0,'
            '"category":"أخرى","description":""}'
        )

        base_url = "https://generativelanguage.googleapis.com/v1beta/models/"
        url = base_url + "gemini-2.5-flash:generateContent"
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

        data_bytes = json.dumps(payload).encode("utf-8")
        req = urllib.request.Request(
            url, data=data_bytes, headers=headers, method="POST"
        )
        with urllib.request.urlopen(req, timeout=60) as resp:
            res = json.loads(resp.read().decode("utf-8"))
            txt = res["candidates"][0]["content"]["parts"][0]["text"]
            return json.loads(txt)

    except urllib.error.HTTPError as e:
        err_body = e.read().decode("utf-8", errors="ignore")
        return {"error": f"خطأ Google ({e.code}): {err_body}"}
    except Exception as e:
        return {"error": str(e)}

st.subheader("1️⃣ رفع الفواتير (PDF أو صورة)")
uploaded_files = st.file_uploader(
    "ارفع الفواتير (حتى 5 فواتير معاً)",
    type=["pdf", "png", "jpg", "jpeg", "webp"],
    accept_multiple_files=True
)

if uploaded_files:
    if len(uploaded_files) > 5:
        st.warning("⚠️ سيتم معالجة أول 5 فواتير فقط.")
        uploaded_files = uploaded_files[:5]

    if not API_KEY:
        st.error("⚠️ مفتاح GEMINI_API_KEY غير موجود في Secrets.")

    st.divider()
    st.subheader("2️⃣ مراجعة البيانات")

    dar_list = [
        "إدارة الدور النسائية", "دار حصة", "دار التقوى",
        "دار الصفا", "دار لمياء بنت الخياط", "دار منيرة المانع",
        "دار الفتاة", "دار الفايزية", "دار الهدى",
        "دار الضياء", "دار زينب", "دار العليا",
        "دار الخنساء", "التاليات", "خولة بنت حكيم",
        "غراس", "سيدة بنيجات", "الأفق",
        "الوقار", "دار الطرفية"
    ]

    dar_name = st.selectbox("اختر اسم الدار", dar_list)
    categories_list = [
        "الكهرباء", "المياة", "الهاتف والبريد",
        "الصيانة", "قرطاسيات ومطبوعات", "نظافة", "أخرى"
    ]

    invoices_data = []

    for idx, file in enumerate(uploaded_files):
        f_bytes = file.getvalue()
        f_hash = get_file_hash(f_bytes)

        if f_hash not in st.session_state.parsed_cache:
            with st.spinner(f"🔍 جاري قراءة الفاتورة #{idx + 1}..."):
                m_type = file.type or "application/pdf"
                st.session_state.parsed_cache[f_hash] = extract_any_invoice(
                    f_bytes, m_type
                )

        ext = st.session_state.parsed_cache[f_hash]

        st.markdown(f"### 📄 الفاتورة ({idx + 1}): `{file.name}`")
        if "error" in ext:
            st.warning(f"تعذرت القراءة: {ext['error']}")

        d_cat = ext.get("category", "أخرى")
        cat_idx = categories_list.index(d_cat) if d_cat in categories_list else 6
        d_desc = ext.get("description", "سداد فاتورة")
        d_sup = ext.get("supplier", "")

        col1, col2 = st.columns(2)
        with col1:
            supplier = st.text_input(
                f"اسم المورد #{idx+1}",
                value=str(d_sup),
                key=f"sup_{f_hash}_{idx}"
            )
            inv_no = st.text_input(
                f"رقم الفاتورة #{idx+1}",
                value=str(ext.get("inv_no", "")),
                key=f"inv_{f_hash}_{idx}"
            )
            account_no = st.text_input(
                f"رقم الحساب #{idx+1}",
                value=str(ext.get("account_no", "")),
                key=f"acc_{f_hash}_{idx}"
            )
            inv_date = st.text_input(
                f"تاريخ الفاتورة #{idx+1}",
                value=str(ext.get("inv_date", "")),
                key=f"date_{f_hash}_{idx}"
            )

        with col2:
            tax_no = st.text_input(
                f"الرقم الضريبي #{idx+1}",
                value=str(ext.get("tax_no", "")),
                key=f"tax_{f_hash}_{idx}"
            )
            amount = st.number_input(
                f"المبلغ بدون الضريبة #{idx+1}",
                value=float(ext.get("amount_pre_tax", 0.0) or 0.0),
                step=0.01,
                key=f"amt_{f_hash}_{idx}"
            )
            statement = st.text_input(
                f"البيان #{idx+1}",
                value=f"{d_desc} - {dar_name}",
                key=f"st_{f_hash}_{idx}"
            )
            category = st.selectbox(
                f"بند الحساب #{idx+1}",
                categories_list,
                index=cat_idx,
                key=f"cat_{f_hash}_{idx}"
            )

            est_total = amount * 1.15
            w_amount = tafqeet_arabic(est_total)
            w_text = st.text_input(
                f"التفقيط #{idx+1}",
                value=w_amount,
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
            "written_amount": w_text
        })
        st.markdown("---")

    st.subheader("3️⃣ المعالجة والإرسال")

    if st.button("🚀 تفريغ العهدة والتجهيز لمنصة فنار"):
        excel_path = "نموذج عهدة الادارة العامة 2026.xlsx.xlsx"

        if os.path.exists(excel_path):
            try:
                wb = openpyxl.load_workbook(excel_path)
                ws = wb.active

                first_dt = invoices_data[0]["inv_date"]
                try:
                    p_dt = datetime.strptime(first_dt, "%Y-%m-%d")
                except Exception:
                    p_dt = datetime.today()

                m_name = get_arabic_month(p_dt.month)
                yr = p_dt.year

                ws["A2"] = f"تفريغ عهدة : {dar_name} شهر {m_name} {yr} م -"

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
                    r = start_row + idx
                    ws[f"C{r}"] = item["statement"]
                    ws[f"D{r}"] = item["account_no"]

                    raw_inv = str(item["inv_no"]).strip()
                    ws[f"E{r}"] = int(raw_inv) if raw_inv.isdigit() else raw_inv

                    ws[f"F{r}"] = item["inv_date"]
                    ws[f"G{r}"] = item["tax_no"]
                    ws[f"G{r}"].number_format = "@"
                    ws[f"H{r}"] = item["supplier"]

                    col_l = cat_col_map.get(item["category"], "W")
                    ws[f"{col_l}{r}"] = item["amount"]

                out_buf = io.BytesIO()
                wb.save(out_buf)
                wb.close()
                out_buf.seek(0)

                out_name = f"تفريغ_عهدة_{dar_name}_2026.xlsx"
                st.success(f"🎉 تم تفريغ البيانات لـ ({dar_name}) بنجاح!")

                st.download_button(
                    label="📥 تحميل ملف Excel المفرغ",
                    data=out_buf,
                    file_name=out_name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

                st.link_button(
                    "🌐 الذهاب إلى منصة فنار",
                    "https://fanar.gqb.org.sa/requests/home"
                )

            except Exception as e:
                st.error(f"حدث خطأ: {e}")
        else:
            st.error(f"⚠️ الملف غير موجود: {excel_path}")

if st.session_state.parsed_cache:
    if st.sidebar.button("🧹 مسح ذاكرة الفواتير"):
        st.session_state.parsed_cache.clear()
        st.rerun()
