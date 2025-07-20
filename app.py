import streamlit as st
import pandas as pd
import zipfile
import io
from pathlib import Path
from docx import Document
import docx.shared
import base64
from fpdf import FPDF
from PIL import Image
import tempfile
import os

st.set_page_config(page_title="SAT 오답노트 생성기", layout="centered")
st.title("📝 SAT 오답노트 생성기")

# 예시 엑셀
example_df = pd.DataFrame({
    "이름": ["홍길동", "김민지"],
    "문서제목": ["25 SAT MATH S2 만점반 Mock3", "25 SAT MATH S2 만점반 Mock3"],
    "Module1": ["1,3,5", ""],
    "Module2": ["", "2,4"]
})

# 예시 엑셀 다운로드 링크
def get_example_excel_download():
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
        example_df.to_excel(writer, index=False, sheet_name='오답노트')
    buffer.seek(0)
    b64 = base64.b64encode(buffer.read()).decode()
    href = f'<a href="data:application/octet-stream;base64,{b64}" download="예시_오답노트_양식.xlsx">📥 예시 엑셀파일 다운로드</a>'
    return href

# 표시
st.markdown("### 📊 예시 엑셀 양식")
st.dataframe(example_df)
st.markdown(get_example_excel_download(), unsafe_allow_html=True)

# 업로드
st.markdown("### 📦 오답노트 파일 업로드")
uploaded_zip = st.file_uploader("M1, M2 폴더 포함된 ZIP 파일 업로드", type="zip")
uploaded_excel = st.file_uploader("오답노트 엑셀 파일 업로드 (.xlsx)", type=["xlsx"])

# PDF 생성 함수
def create_pdf(name, title, m1_images, m2_images, m1_nums, m2_nums):
    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()
    pdf.set_font("Arial", 'B', 14)
    pdf.cell(0, 10, f"{name}_{title}", ln=True)

    def add_images(section_title, nums, image_dict):
        if nums:
            pdf.set_font("Arial", 'B', 12)
            pdf.cell(0, 10, section_title, ln=True)
            for num in nums:
                possible_names = [f"{num}.png", f"{num}.jpg", f"{num}.jpeg"]
                found_img = None
                for name in possible_names:
                    if name in image_dict:
                        found_img = image_dict[name]
                        break
                if found_img:
                    pdf.set_font("Arial", '', 11)
                    pdf.cell(0, 8, f"문항 {num}", ln=True)
                    with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp_img:
                        tmp_img.write(found_img.read())
                        tmp_img_path = tmp_img.name
                    pdf.image(tmp_img_path, w=170)
                    os.remove(tmp_img_path)

    add_images("Module 1", m1_nums, m1_images)
    add_images("Module 2", m2_nums, m2_images)

    output = io.BytesIO()
    pdf.output(output)
    output.seek(0)
    return output

# PDF 미리보기용 base64 생성
def pdf_preview_base64(pdf_bytes: bytes):
    b64 = base64.b64encode(pdf_bytes).decode()
    pdf_display = f'<iframe src="data:application/pdf;base64,{b64}" width="100%" height="500px" type="application/pdf"></iframe>'
    return pdf_display

# 오답노트 생성
if uploaded_zip and uploaded_excel:
    with zipfile.ZipFile(uploaded_zip) as z:
        m1_imgs = {}
        m2_imgs = {}
        for f in z.namelist():
            filename = Path(f).name
            if f.startswith("M1/") and filename.lower().endswith((".png", ".jpg", ".jpeg")):
                m1_imgs[filename] = io.BytesIO(z.read(f))
            elif f.startswith("M2/") and filename.lower().endswith((".png", ".jpg", ".jpeg")):
                m2_imgs[filename] = io.BytesIO(z.read(f))

    df = pd.read_excel(uploaded_excel)
    zip_buffer = io.BytesIO()

    with zipfile.ZipFile(zip_buffer, mode="w", compression=zipfile.ZIP_DEFLATED) as zipf:
        for _, row in df.iterrows():
            name = str(row['이름'])
            title = str(row['문서제목'])

            m1_list = str(row['Module1']).split(",") if pd.notna(row['Module1']) else []
            m2_list = str(row['Module2']).split(",") if pd.notna(row['Module2']) else []

            m1_nums = [num.strip() for num in m1_list if num.strip()]
            m2_nums = [num.strip() for num in m2_list if num.strip()]

            if not m1_nums and not m2_nums:
                continue

            pdf_io = create_pdf(name, title, m1_imgs, m2_imgs, m1_nums, m2_nums)
            zipf.writestr(f"{name}_{title}.pdf", pdf_io.getvalue())

            # ✅ 미리보기
            with st.expander(f"📄 {name} PDF 미리보기"):
                st.markdown(pdf_preview_base64(pdf_io.getvalue()), unsafe_allow_html=True)

    zip_buffer.seek(0)
    st.markdown("### 📦 전체 ZIP 파일 다운로드")
    st.download_button(
        label="📥 ZIP 파일 다운로드",
        data=zip_buffer,
        file_name="오답노트_모음.zip",
        mime="application/zip"
    )
