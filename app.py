import streamlit as st
import openpyxl
from datetime import datetime
import os

st.set_page_config(
    page_title="أداة تفريغ العهود - 2026",
    page_icon="🏛️",
    layout="centered"
)

st.title("🏛️ أداة تفريغ العهود وإصدار أمر الصرف")
st.caption("الجمعية الخيرية لتحفيظ القرآن ببريدة - نموذج الإدارة العامة 2026")
st.divider()

st.subheader("1️⃣ رفع الفاتورة")
uploaded_file = st.file_uploader("التقط صورة للفاتورة أو اسحب ملف الـ PDF هنا", type=["pdf", "png", "jpg", "jpeg"])

if uploaded_file:
    st.success("✅ تم استلام الفاتورة بنجاح!")
    st.divider()
    
    st.subheader("2️⃣ مراجعة البيانات المستخرجة")
    
    col1, col2 = st.columns(2)
    with col1:
        supplier = st.text_input("اسم المورد", value="شركة الكهرباء السعودية")
        inv_no = st.text_input("رقم الفاتورة", value="262003286782")
        amount = st.number_input("المبلغ (بدون الضريبة)", value=216.11, step=0.01)
        tax_amount = st.number_input("مبلغ الضريبة", value=32.42, step=0.01)
    
    with col2:
        tax_no = st.text_input("الرقم الضريبي", value="300000361310003")
        statement = st.text_input("البيان", value="سداد فاتورة كهرباء")
        category = st.selectbox("بند الحساب (العمود)", [
            "الكهرباء", 
            "المياة", 
            "الهاتف والبريد",
            "قرطاسيات ومطبوعات",
            "نظافة",
            "أخرى"
        ])

    st.divider()
    st.subheader("3️⃣ المعالجة والإرسال")
    
    if st.button("🚀 تفريغ العهدة والتجهيز لمنصة فنار"):
        # اسم ملف Excel الجديد المرفوع على GitHub
        excel_path = "نموذج_العهدة_2026.xlsx"
        
        if os.path.exists(excel_path):
            try:
                wb = openpyxl.load_workbook(excel_path)
                ws = wb.active
                
                # كتابة البيانات في السطر الأول المتاح (الصف 5)
                row = 5
                
                ws[f'C{row}'] = statement      # البيان
                ws[f'D{row}'] = int(inv_no) if inv_no.isdigit() else inv_no # رقم الفاتورة
                ws[f'E{row}'] = datetime.today().strftime('%Y-%m-%d')        # التاريخ
                ws[f'F{row}'] = tax_no         # الرقم الضريبي
                ws[f'G{row}'] = supplier       # اسم المورد
                
                # تحديد العمود بناءً على البند المختار
                cat_col_map = {
                    "الكهرباء": "H",
                    "المياة": "I",
                    "الهاتف والبريد": "J",
                    "قرطاسيات ومطبوعات": "N",
                    "نظافة": "P",
                    "أخرى": "U"
                }
                
                col_letter = cat_col_map.get(category, "H")
                ws[f'{col_letter}{row}'] = amount        # المبلغ في عمود البند
                ws[f'V{row}'] = tax_amount               # مبلغ الضريبة
                
                output_filename = "تفريغ_عهدة_2026_محدث.xlsx"
                wb.save(output_filename)
                
                st.success("🎉 تم تفريغ البيانات في النموذج الجديد بنجاح!")
                
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
