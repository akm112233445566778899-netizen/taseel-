import streamlit as st
import openpyxl
from datetime import datetime
import os

# ضبط إعدادات الصفحة لتناسب الجوال والكمبيوتر
st.set_page_config(
    page_title="أداة تفريغ العهود - منصة فنار",
    page_icon="🏛️",
    layout="centered"
)

# العنوان الرئيسي
st.title("🏛️ أداة تفريغ العهود وإصدار أمر الصرف")
st.caption("الجمعية الخيرية لتحفيظ القرآن ببريدة - دار الرابية")
st.divider()

# الخطوة الأولى: رفع الفاتورة
st.subheader("1️⃣ رفع الفاتورة")
uploaded_file = st.file_uploader("التقط صورة للفاتورة أو اسحب ملف الـ PDF هنا", type=["pdf", "png", "jpg", "jpeg"])

if uploaded_file:
    st.success("✅ تم استلام الفاتورة بنجاح!")
    st.divider()
    
    # الخطوة الثانية: مراجعة البيانات
    st.subheader("2️⃣ بيانات الفاتورة المفرغة")
    
    col1, col2 = st.columns(2)
    with col1:
        supplier = st.text_input("اسم المورد", value="شركة الكهرباء السعودية")
        inv_no = st.text_input("رقم الفاتورة", value="262003286782")
        amount = st.number_input("المبلغ الإجمالي (ريال)", value=248.53, step=0.01)
    
    with col2:
        tax_no = st.text_input("الرقم الضريبي", value="300000361310003")
        account_no = st.text_input("رقم الحساب / العداد", value="30044598537")
        category = st.selectbox("بند الحساب", [
            "كهرباء (4230100505)", 
            "مياه (4230100506)", 
            "هاتف وبريد (4230100502)"
        ])

    st.divider()
    st.subheader("3️⃣ المعالجة والإرسال")
    
    if st.button("🚀 تفريغ العهدة والتجهيز لمنصة فنار"):
        excel_path = "تفريغ عهدة سداد دار الرابية.xlsx"
        
        if os.path.exists(excel_path):
            try:
                # فتح ملف Excel وتحديث البيانات
                wb = openpyxl.load_workbook(excel_path)
                ws = wb.active
                
                # كتابة البيانات المفرغة في الصف المخصص
                ws['C6'] = "الرابية"
                ws['D6'] = account_no
                ws['E6'] = int(inv_no) if inv_no.isdigit() else inv_no
                ws['F6'] = datetime.today().strftime('%Y-%m-%d')
                ws['G6'] = tax_no
                ws['H6'] = supplier
                ws['J6'] = amount
                
                output_filename = "تفريغ_عهدة_محدث.xlsx"
                wb.save(output_filename)
                
                st.success("🎉 تم تفريغ الفاتورة بنجاح في ملف Excel العهدة!")
                
                # خيار تحميل الملف المفرغ
                with open(output_filename, "rb") as file:
                    st.download_button(
                        label="📥 تحميل ملف Excel المفرغ جاهز للرفع",
                        data=file,
                        file_name=output_filename,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                
                st.info("💡 اضغط على الزر أدناه للانتقال المباشر لإنشاء أمر الصرف في منصة فنار:")
                st.link_button("🌐 الذهاب إلى منصة فنار لإنشاء أمر الصرف", "https://fanar.gqb.org.sa/requests/home")
                
            except Exception as e:
                st.error(f"حدث خطأ أثناء معالجة ملف Excel: {e}")
        else:
            st.error("⚠️ ملف النموذج الأصلي (تفريغ عهدة سداد دار الرابية.xlsx) غير موجود في المستودع!")
