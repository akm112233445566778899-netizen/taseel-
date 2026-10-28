import streamlit as st
import openpyxl
from datetime import datetime
import os
import json
from PIL import Image
import google.generativeai as genai
import pypdf
import io

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
    """دالة بسيطة لتحويل المبلغ إلى نص عربي (تفقيط)"""
    try:
        if amount <= 0:
            return ""
        
        units = ["", "واحد", "اثنان", "ثلاثة", "أربعة", "خمسة", "ستة", "سبعة", "ثمانية", "تسعة", "عشرة"]
        
        riyal = int(amount)
        halalas = int(round((amount - riyal) * 100))
        
        text = f"{riyal} ريال سعودي"
        if halalas > 0:
            text += f" و {halalas} هللة"
        text += " فقط لا غير"
        return text
    except Exception:
        return ""

def extract_invoice_data(uploaded_file):
    """استخراج بيانات الفاتورة بدقة عالية من الصور والـ PDF"""
    if not API_KEY:
        return {}
    try:
        genai.configure(api_key=API_KEY)
        model = genai.GenerativeModel('gemini-1.5-flash')
        
        inputs = []
        
        if uploaded_file.type == "application/pdf":
            # قراءة PDF كنص أو استخراج الصور منه
            reader = pypdf.PdfReader(uploaded_file)
            pdf_text = ""
            for page in reader.pages:
                pdf_text += page.extract_text() or ""
            
            if len(pdf_text.strip()) > 20:
                inputs.append(f"محتوى ملف الـ PDF النصي:\n{pdf_text}")
            else:
                # إذا كان PDF مصور، نأخذ أول صفحة كصورة
                for page in reader.pages:
                    for img_obj in page.images:
                        img_bytes = img_obj.data
                        inputs.append(Image.open(io.BytesIO(img_bytes)))
                        break
                    if inputs:
                        break
        else:
            inputs.append(Image.open(uploaded_file))
            
        if not inputs:
            return {}
            
        prompt = """
        أنت محاسب تدقق في الفواتير السعودية. قم بفحص المرفق بدقة واستخرج البيانات التالية بصيغة JSON حصرية فقط:
        {
            "inv_no": "رقم الفاتورة بالضبط",
            "tax_no": "الرقم الضريبي للمورد (غالباً 15 رقم يبدأ بـ 3)",
            "supplier": "اسم الشركة أو المورد",
            "amount": 0.00,
            "tax_amount": 0.00
        }
        
        شروط مهمة:
        - استخرج رقم الفاتورة بدون حذف أي أرقام منه.
        - الرقم الضريبي يبدأ بـ 3 وينتهي بـ 3 ويتكون من 15 رقم.
        - لا تضف أي نص خارج JSON.
        """
        inputs.append(prompt)
        
        response = model.generate_content(inputs)
        clean_json = response.text.replace("```json", "").replace("```", "").strip()
        return json.loads(clean_json)
    except Exception as e:
        return {}

st.subheader("1️⃣ رفع الفواتير (حتى 5 فواتير)")
uploaded_files = st.file_uploader(
    "التقط صور الفواتير أو اسحب ملفات الـ PDF هنا (يمكنك تحديد حتى 5 فواتير معاً)", 
    type=["pdf", "png", "jpg", "jpeg"],
    accept_multiple_files=True
)

if uploaded_files:
    if len(uploaded_files) > 5:
        st.warning("⚠️ تم تحديد أكثر من 5 فواتير. سيتم معالجة أول 5 فواتير فقط.")
        uploaded_files = uploaded_files[:5]
        
    st.success(f"✅ تم استلام {len(uploaded_files)} فاتورة بنجاح!")
    
    if not API_KEY:
        st.error("⚠️ مفتاح GEMINI_API_KEY غير معرف في إعدادات Secrets!")
        
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
            with st.spinner(f"🔍 جاري قراءة واستخراج بيانات الفاتورة #{idx + 1}..."):
                ext_data = extract_invoice_data(file)
            
        col1, col2 = st.columns(2)
        with col1:
            supplier = st.text_input(f"اسم المورد #{idx+1}", value=ext_data.get("supplier", "شركة الكهرباء السعودية"), key=f"sup_{idx}")
            inv_no = st.text_input(f"رقم الفاتورة #{idx+1}", value=str(ext_data.get("inv_no", "")), key=f"inv_{idx}")
            amount = st.number_input(f"المبلغ بدون الضريبة #{idx+1}", value=float(ext_data.get("amount", 0.0) or 0.0), step=0.01, key=f"amt_{idx}")
            tax_amount = st.number_input(f"مبلغ الضريبة #{idx+1}", value=float(ext_data.get("tax_amount", 0.0) or 0.0), step=0.01, key=f"tax_amt_{idx}")
            
        with col2:
            tax_no = st.text_input(f"الرقم الضريبي #{idx+1}", value=str(ext_data.get("tax_no", "")), key=f"tax_{idx}")
            statement = st.text_input(f"البيان #{idx+1}", value=f"سداد فاتورة - {dar_name}", key=f"st_{idx}")
            category = st.selectbox(f"بند الحساب #{idx+1}", [
                "الكهرباء", 
                "المياة", 
                "الهاتف والبريد",
                "الصيانة",
                "قرطاسيات ومطبوعات",
                "نظافة",
                "أخرى"
            ], key=f"cat_{idx}")
            
            # حساب الإجمالي وتحويله إلى كلمات (تفقيط) تلقائياً
            total_amt = amount + tax_amount
            written_amount = tafqeet_arabic(total_amt)
            written_text = st.text_input(f"المبلغ كتابةً (تفقيط) #{idx+1}", value=written_amount, key=f"writ_{idx}")

        invoices_data.append({
            "supplier": supplier,
            "inv_no": inv_no,
            "amount": amount,
            "tax_amount": tax_amount,
            "tax_no": tax_no,
            "statement": statement,
            "category": category,
            "written_amount": written_text
        })
        st.markdown("---")

    st.subheader("3️⃣ المعالجة والإرسال")
    
    if st.button("🚀 تفريغ العهدة والتجهيز لمنصة فنار"):
        excel_path = "نموذج عهدة الادارة العامة 2026.xlsx"
        
        if os.path.exists(excel_path):
            try:
                wb = openpyxl.load_workbook(excel_path)
                ws = wb.active
                
                start_row = 5
                
                cat_col_map = {
                    "الكهرباء": "H",
                    "المياة": "I",
                    "الهاتف والبريد": "J",
                    "الصيانة": "M",
                    "قرطاسيات ومطبوعات": "N",
                    "نظافة": "P",
                    "أخرى": "U"
                }
                
                for idx, item in enumerate(invoices_data):
                    current_row = start_row + idx
                    
                    ws[f'C{current_row}'] = item["statement"]
                    ws[f'D{current_row}'] = int(item["inv_no"]) if str(item["inv_no"]).isdigit() else item["inv_no"]
                    ws[f'E{current_row}'] = datetime.today().strftime('%Y-%m-%d')
                    ws[f'F{current_row}'] = item["tax_no"]
                    ws[f'G{current_row}'] = item["supplier"]
                    
                    col_letter = cat_col_map.get(item["category"], "H")
                    ws[f'{col_letter}{current_row}'] = item["amount"]
                    ws[f'V{current_row}'] = item["tax_amount"]
                
                output_filename = f"تفريغ_عهدة_{dar_name}_2026.xlsx"
                wb.save(output_filename)
                
                st.success(f"🎉 تم تفريغ {len(invoices_data)} فاتورة لـ ({dar_name}) بنجاح!")
                
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
