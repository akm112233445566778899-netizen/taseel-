import streamlit as st
import openpyxl
from datetime import datetime
import os
import json
from PIL import Image
import google.generativeai as genai

st.set_page_config(
    page_title="أداة تفريغ العهود - 2026",
    page_icon="🏛️",
    layout="centered"
)

st.title("🏛️ أداة تفريغ العهود وإصدار أمر الصرف")
st.caption("الجمعية الخيرية لتحفيظ القرآن ببريدة - نموذج الإدارة العامة 2026")
st.divider()

# مفتاح API الخاص بـ Gemini لقراءة الفواتير
API_KEY = st.secrets.get("GEMINI_API_KEY", "")

def extract_invoice_data(image_file):
    """دالة تقرأ الفاتورة واستخراج رقم الفاتورة والبيانات تلقائياً"""
    if not API_KEY:
        return None
    try:
        genai.configure(api_key=API_KEY)
        model = genai.GenerativeModel('gemini-1.5-flash')
        image = Image.open(image_file)
        
        prompt = """
        استخرج البيانات التالية من الفاتورة بدقة وأرجعها بصيغة JSON فقط بهذا الشكل:
        {
            "inv_no": "رقم الفاتورة",
            "supplier": "اسم المورد",
            "tax_no": "الرقم الضريبي للمورد",
            "amount": "المبلغ قبل الضريبة كعدد",
            "tax_amount": "مبلغ الضريبة كعدد"
        }
        إذا لم تجد أي عنصر اترك قيمته نص فارغ "". لا تضف أي كلام إضافي خارج الـ JSON.
        """
        
        response = model.generate_content([image, prompt])
        clean_json = response.text.replace("```json", "").replace("```", "").strip()
        data = json.loads(clean_json)
        return data
    except Exception as e:
        return None

st.subheader("1️⃣ رفع الفاتورة")
uploaded_file = st.file_uploader("التقط صورة للفاتورة أو اسحب ملف الفاتورة هنا", type=["png", "jpg", "jpeg"])

if uploaded_file:
    st.success("✅ تم استلام الفاتورة بنجاح!")
    
    # محاولة قراءة رقم الفاتورة والبيانات تلقائياً
    extracted_data = None
    with st.spinner("🔍 جاري قراءة واستخراج رقم الفاتورة والبيانات تلقائياً..."):
        extracted_data = extract_invoice_data(uploaded_file)
    
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
    
    # تعبئة القيم المستخرجة تلقائياً إذا تم العثور عليها
    default_inv = extracted_data.get("inv_no", "") if extracted_data else ""
    default_supplier = extracted_data.get("supplier", "شركة الكهرباء السعودية") if extracted_data else "شركة الكهرباء السعودية"
    default_tax_no = extracted_data.get("tax_no", "") if extracted_data else ""
    default_amount = float(extracted_data.get("amount", 0.0)) if (extracted_data and extracted_data.get("amount")) else 0.0
    default_tax = float(extracted_data.get("tax_amount", 0.0)) if (extracted_data and extracted_data.get("tax_amount")) else 0.0

    col1, col2 = st.columns(2)
    with col1:
        supplier = st.text_input("اسم المورد", value=default_supplier)
        inv_no = st.text_input("رقم الفاتورة (مستخرج تلقائياً)", value=default_inv)
        amount = st.number_input("المبلغ (بدون الضريبة)", value=default_amount, step=0.01)
        tax_amount = st.number_input("مبلغ الضريبة", value=default_tax, step=0.01)
    
    with col2:
        tax_no = st.text_input("الرقم الضريبي", value=default_tax_no)
        statement = st.text_input("البيان", value=f"سداد فاتورة - {dar_name}")
        
        category = st.selectbox("بند الحساب (العمود)", [
            "الكهرباء", 
            "المياة", 
            "الهاتف والبريد",
            "الصيانة",
            "قرطاسيات ومطبوعات",
            "نظافة",
            "أخرى"
        ])

    st.divider()
    st.subheader("3️⃣ المعالجة والإرسال")
    
    if st.button("🚀 تفريغ العهدة والتجهيز لمنصة فنار"):
        excel_path = "نموذج_العهدة_2026.xlsx"
        
        if os.path.exists(excel_path):
            try:
                wb = openpyxl.load_workbook(excel_path)
                ws = wb.active
                
                row = 5
                
                ws[f'C{row}'] = statement      # البيان
                ws[f'D{row}'] = int(inv_no) if str(inv_no).isdigit() else inv_no # رقم الفاتورة
                ws[f'E{row}'] = datetime.today().strftime('%Y-%m-%d')        # التاريخ
                ws[f'F{row}'] = tax_no         # الرقم الضريبي
                ws[f'G{row}'] = supplier       # اسم المورد
                
                cat_col_map = {
                    "الكهرباء": "H",
                    "المياة": "I",
                    "الهاتف والبريد": "J",
                    "الصيانة": "M",
                    "قرطاسيات ومطبوعات": "N",
                    "نظافة": "P",
                    "أخرى": "U"
                }
                
                col_letter = cat_col_map.get(category, "H")
                ws[f'{col_letter}{row}'] = amount        # المبلغ في عمود البند
                ws[f'V{row}'] = tax_amount               # مبلغ الضريبة
                
                output_filename = f"تفريغ_عهدة_{dar_name}_2026.xlsx"
                wb.save(output_filename)
                
                st.success(f"🎉 تم تفريغ البيانات لـ ({dar_name}) ورقم الفاتورة ({inv_no}) بنجاح!")
                
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
            st.error(f"⚠️ لم يتم العثور على الملف ({excel_path}) داخل المستودع، يُرجى التأكد من رفعه بنفس الاسم.")
