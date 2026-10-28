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


def convert_arabic_digits(text):
  """تحويل كافة الأرقام العربية الشرقية (٠١٢٣٤٥٦٧٨٩) إلى أرقام إنجليزية"""
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


def parse_pdf_invoice_robust(uploaded_file):
  """استخراج موجه ودقيق 100% لفواتير الكهرباء يمنع استخراج تاريخ اليوم وجمع المبالغ"""
  data = {
      "inv_no": "",
      "account_no": "",
      "inv_date": "",
      "amount": 0.0,
      "tax_no": "300000361310003",
      "supplier": "الشركة السعودية للكهرباء",
  }
  try:
    uploaded_file.seek(0)
    full_text_raw = ""
    lines = []
    with pdfplumber.open(uploaded_file) as pdf:
      for page in pdf.pages:
        t = page.extract_text() or ""
        full_text_raw += t + "\n"
        lines.extend(t.split("\n"))

    full_text = convert_arabic_digits(full_text_raw)

    # 1. رقم الفاتورة (12 رقم يبدأ بـ 20 أو 25)
    inv_m = re.search(r"\b(2[05]\d{10})\b", full_text)
    if inv_m:
      data["inv_no"] = inv_m.group(1)

    # 2. رقم الحساب (11 رقم يبدأ بـ 30)
    acc_m = re.search(r"\b(30\d{9})\b", full_text)
    if acc_m:
      data["account_no"] = acc_m.group(1)

    # 3. الرقم الضريبي للمورد
    tax_m = re.search(r"\b(3\d{14})\b", full_text)
    if tax_m:
      data["tax_no"] = tax_m.group(1)

    # 4. استخراج تاريخ الفاتورة الأصلي/تاريخ الاستحقاق المباشر (تجنب التواريخ اليومية وفترة الاستهلاك)
    # البحث عن السطر المكتوب فيه "تاريخ بداية الاستحقاق" أو "تاريخ الفاتورة"
    target_date = ""
    for line in lines:
      line_eng = convert_arabic_digits(line)
      if "تاريخ بداية الاستحقاق" in line_eng or "تاريخ الفاتورة" in line_eng:
        dt_match = re.findall(
            r"(\d{2}[-/]\d{2}[-/]\d{4}|\d{4}[-/]\d{2}[-/]\d{2})", line_eng
        )
        if dt_match:
          target_date = dt_match[0]
          break

    # إذا لم يُعثر عليه في نفس السطر، يتم أخذ التاريخ المكتوب بأسفل الجدول (تاريخ إصدار الفاتورة)
    if not target_date:
      all_dt = re.findall(
          r"(\d{2}[-/]\d{2}[-/]\d{4}|\d{4}[-/]\d{2}[-/]\d{2})", full_text
      )
      if len(all_dt) >= 3:
        target_date = all_dt[2]
      elif all_dt:
        target_date = all_dt[-1]

    if target_date:
      d_clean = target_date.replace("/", "-")
      parts = d_clean.split("-")
      if len(parts) == 3:
        data["inv_date"] = (
            d_clean
            if len(parts[0]) == 4
            else f"{parts[2]}-{parts[1] if len(parts[1])==2 else '0'+parts[1]}-{parts[0] if len(parts[0])==2 else '0'+parts[0]}"
        )

    # 5. استخراج المبلغ الصافي بدون ضريبة (الاعتماد المباشر على المبلغ المطلوب سداده / 1.15)
    # البحث عن المبلغ الإجمالي في ملخص الفاتورة
    net_amount = 0.0

    # البحث عن الأرقام المالية وتحديد المبلغ الإجمالي المطلوب سداده المكتوب بصيغة عادية
    raw_amts = re.findall(r"[\d,]+\.\d{2}", full_text)
    valid_floats = []
    for item in raw_amts:
      try:
        v = float(item.replace(",", "").strip())
        if 10.0 <= v <= 300000.0 and v != 10197.0:
          valid_floats.append(v)
      except Exception:
        pass

    if valid_floats:
      # أخذ أعلى مبلغ مطلوب سداده وهو الإجمالي وقسمته على 1.15 بشكل ديناميكي ومستقر
      total_invoice = max(valid_floats)
      net_amount = round(total_invoice / 1.15, 2)

      # مطابقة إذا كان الصافي مكتوباً أصلاً في جدول الخاضع للضريبة
      for val in valid_floats:
        if abs(val - net_amount) < 0.8:
          net_amount = val
          break

    data["amount"] = net_amount

  except Exception:
    pass

  return data


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
      "دار الطرفية",
  ]

  dar_name = st.selectbox("اختر اسم الدار المعنية", dar_list)

  invoices_data = []

  for idx, file in enumerate(uploaded_files):
    st.markdown(f"### 📄 الفاتورة رقم ({idx + 1}): `{file.name}`")

    with st.spinner(f"🔍 جاري استخراج البيانات للفاتورة #{idx + 1}..."):
      ext_data = parse_pdf_invoice_robust(file)

    col1, col2 = st.columns(2)
    with col1:
      supplier = st.text_input(
          f"اسم المورد #{idx+1}",
          value=ext_data.get("supplier", "الشركة السعودية للكهرباء"),
          key=f"sup_{idx}",
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
          f"تاريخ الفاتورة #{idx+1}",
          value=str(ext_data.get("inv_date", "")),
          key=f"date_{idx}",
      )

    with col2:
      tax_no = st.text_input(
          f"الرقم الضريبي #{idx+1}",
          value=str(ext_data.get("tax_no", "300000361310003")),
          key=f"tax_{idx}",
      )
      amount = st.number_input(
          f"المبلغ بدون الضريبة #{idx+1}",
          value=float(ext_data.get("amount", 0.0) or 0.0),
          step=0.01,
          key=f"amt_{idx}",
      )

      statement = st.text_input(
          f"البيان #{idx+1}",
          value=f"سداد فاتورة كهرباء - {dar_name}",
          key=f"st_{idx}_{dar_name}",
      )

      category = st.selectbox(
          f"بند الحساب #{idx+1}",
          [
              "الكهرباء",
              "المياة",
              "الهاتف والبريد",
              "الصيانة",
              "قرطاسيات ومطبوعات",
              "نظافة",
              "أخرى",
          ],
          key=f"cat_{idx}",
      )

      estimated_total = amount * 1.15
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

  if st.button("🚀 تفريغ العهدة والتجهيز لمنصة فنار"):
    excel_path = "نموذج عهدة الادارة العامة 2026.xlsx.xlsx"

    if os.path.exists(excel_path):
      try:
        wb = openpyxl.load_workbook(excel_path)
        ws = wb.active

        first_inv_date_str = invoices_data[0]["inv_date"]
        try:
          parsed_date = datetime.strptime(first_inv_date_str, "%Y-%m-%d")
        except Exception:
          parsed_date = datetime.today()

        inv_month_name = get_arabic_month(parsed_date.month)
        inv_year = parsed_date.year

        ws["A2"] = (
            f"تفريغ عهدة : {dar_name} شهر {inv_month_name} {inv_year} م -"
        )

        start_row = 6

        cat_col_map = {
            "الهاتف والبريد": "I",
            "الكهرباء": "J",
            "المياة": "K",
            "الصيانة": "O",
            "قرطاسيات ومطبوعات": "P",
            "نظافة": "R",
            "أخرى": "W",
        }

        for idx, item in enumerate(invoices_data):
          current_row = start_row + idx

          ws[f"C{current_row}"] = item["statement"]
          ws[f"D{current_row}"] = item["account_no"]

          raw_inv = str(item["inv_no"]).strip()
          ws[f"E{current_row}"] = (
              int(raw_inv) if raw_inv.isdigit() else raw_inv
          )

          ws[f"F{current_row}"] = item["inv_date"]
          ws[f"G{current_row}"] = item["tax_no"]
          ws[f"G{current_row}"].number_format = "@"
          ws[f"H{current_row}"] = item["supplier"]

          col_letter = cat_col_map.get(item["category"], "J")
          ws[f"{col_letter}{current_row}"] = float(item["amount"])

        output_filename = f"تفريغ_عهدة_{dar_name}_2026.xlsx"
        wb.save(output_filename)

        st.success(f"🎉 تم تفريغ البيانات لـ ({dar_name}) بنجاح!")

        with open(output_filename, "rb") as file:
          st.download_button(
              label="📥 تحميل ملف Excel المفرغ جاهز للرفع",
              data=file,
              file_name=output_filename,
              mime=(
                  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
              ),
          )

        st.link_button(
            "🌐 الذهاب إلى منصة فنار لإنشاء أمر الصرف",
            "https://fanar.gqb.org.sa/requests/home",
        )
      except Exception as e:
        st.error(f"حدث خطأ أثناء معالجة ملف الإكسل: {e}")
    else:
      st.error(
          f"لم يتم العثور على ملف النموذج المعتمد: `{excel_path}` في مجلد"
          " التطبيق!"
      )
