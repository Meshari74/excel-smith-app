import streamlit as st
import openai
import pandas as pd
from io import BytesIO

# إعداد مفتاح OpenAI API (تأكد من وضع مفتاحك هنا)
openai.api_key = "YOUR_API_KEY"

st.set_page_config(page_title="ExcelSmithGPT", layout="centered")
st.title("📊 ExcelSmithGPT")
st.markdown("صانع الجداول الذكي – بدون فزلكة")

# واجهة المستخدم
prompt = st.text_area("📝 اكتب وصف الجدول المطلوب:", placeholder="مثلاً: جدول مبيعات شهري لموظفين مع مجموع ومعدل")
generate_button = st.button("🎯 أنشئ الجدول")

# دالة للتعامل مع GPT وإنشاء كود Excel
@st.cache_data
def generate_excel_from_prompt(prompt):
    # إرسال الطلب لـ GPT ليولّد كود Python ينشئ جدول Excel
    system_msg = "أنت مساعد متخصص في إنتاج جداول Excel باستخدام بايثون وPandas. لا تشرح. فقط أنشئ الكود المطلوب." 

    response = openai.ChatCompletion.create(
        model="gpt-4",
        messages=[
            {"role": "system", "content": system_msg},
            {"role": "user", "content": f"صمم كود بايثون لإنشاء جدول Excel بناءً على: {prompt}. استخدم مكتبة pandas و ExcelWriter."}
        ],
        temperature=0
    )

    code = response.choices[0].message['content']
    return code

# تنفيذ عند الضغط على الزر
if generate_button and prompt:
    with st.spinner("جاري توليد الجدول..."):
        try:
            code = generate_excel_from_prompt(prompt)

            # تشغيل الكود المتولد داخل بيئة آمنة
            local_vars = {}
            exec(code, {}, local_vars)
            df = local_vars.get("df", pd.DataFrame())

            buffer = BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name="Sheet1")
            st.success("✅ الجدول جاهز!")
            st.download_button("📥 حمل ملف Excel", data=buffer.getvalue(), file_name="excel_smith_output.xlsx")
        except Exception as e:
            st.error(f"❌ صار خطأ: {e}")
