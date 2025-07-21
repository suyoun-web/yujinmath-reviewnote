import streamlit as st
import pandas as pd
import zipfile
import os
import io
from PIL import Image
from fpdf import FPDF
from datetime import datetime

# PDF 생성용 폰트 경로
FONT_REGULAR = "fonts/NanumGothic.ttf"
FONT_BOLD = "fonts/NanumGothicBold.ttf"
pdf_font_name = "NanumGothic"

if os.path.exists(FONT_REGULAR) and os.path.exists(FONT_BOLD):
    class KoreanPDF(FPDF):
        def __init__(self):
            super().__init__()
            self.add_font(pdf_font_name, '', FONT_REGULAR, uni=True)
            self.add_font(pdf_font_name, 'B', FONT_BOLD, uni=True)
            self.set_font(pdf_font_name, size=10)
else:
    st.error("⚠️ 한글 PDF 생성을 위해 fonts 폴더에 NanumGothic.ttf 와 NanumGothicBold.ttf 모두 필요합니다.")

# 예시 엑셀 다운로드용 버퍼 생성
def get_example_excel():
    output = io.BytesIO()
    example_df = pd.DataFrame({
        '이름': ['홍길동', '김철수'],
        'Module1': ['1,3,5', '2,4'],
        'Module2': ['2,6', '1,3']
    })
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        example_df.to_excel(writer, index=False)
    output.seek(0)
    return output

def extract_zip_to_dict(zip_file):
    m1_imgs, m2_imgs = {}, {}
    with zipfile.ZipFile(zip_file) as z:
        for file in z.namelist():
            if file.lower().endswith(('png', 'jpg', 'jpeg')):
                folder = file.split('/')[0].lower()
                q_num = os.path.splitext(os.path.basename(file))[0]
                with z.open(file) as f:
                    img = Image.open(f).convert("RGB")
                    if folder == "m1":
                        m1_imgs[q_num] = img
                    elif folder == "m2":
                        m2_imgs[q_num] = img
    return m1_imgs, m2_imgs

def create_student_pdf(name, m1_imgs, m2_imgs, doc_title, output_dir):
    pdf = KoreanPDF()
    pdf.set_margins(left=25.4, top=30.0, right=25.4)  # cm → mm: 2.54cm = 25.4mm, 3cm = 30.0mm
    pdf.add_page()
    pdf.set_font(pdf_font_name, style='B', size=10)
    pdf.cell(0, 8, txt=f"<{name}_{doc_title}>", ln=True)

    def add_images(title, images):
        img_est_height = 100
        module_title = "<Module1>" if title == "Module 1" else "<Module2>"

        if title == "Module 2" and pdf.get_y() + 10 + (img_est_height if images else 0) > pdf.page_break_trigger:
            pdf.add_page()

        pdf.set_font(pdf_font_name, size=10)
        pdf.cell(0, 8, txt=module_title, ln=True)
        if images:
            for img in images:
                img_path = f"temp_{datetime.now().timestamp()}.jpg"
                img.save(img_path)
                pdf.image(img_path, w=180)
                os.remove(img_path)
                pdf.ln(8)
        else:
            pdf.ln(8)  # 이미지가 없더라도 공간 확보용

    add_images("Module 1", m1_imgs)
    add_images("Module 2", m2_imgs)

    pdf_path = os.path.join(output_dir, f"{name}_{doc_title}.pdf")
    pdf.output(pdf_path)
    return pdf_path

st.set_page_config(page_title="SAT 오답노트 생성기", layout="centered")
st.title("📝 SAT 오답노트 생성기")

st.header("📊 예시 엑셀 양식")
with st.expander("예시 엑셀파일 열기"):
    st.dataframe(pd.read_excel(get_example_excel()))
example = get_example_excel()
st.download_button("📥 예시 엑셀파일 다운로드", example, file_name="예시_오답노트_양식.xlsx")

st.header("📄 문서 제목 입력")
doc_title = st.text_input("문서 제목 (예: 25 SAT MATH S2 만점반 Mock3)", value="SAT 오답노트")

st.header("📦 오답노트 파일 업로드")
st.caption("M1, M2 폴더 포함된 ZIP 파일 업로드")
img_zip = st.file_uploader("", type="zip")

st.caption("오답노트 엑셀 파일 업로드 (.xlsx)")
excel_file = st.file_uploader("", type="xlsx")

generated_files = []
generate = st.button("📎 오답노트 생성")

if generate and img_zip and excel_file:
    try:
        m1_imgs, m2_imgs = extract_zip_to_dict(img_zip)
        df = pd.read_excel(excel_file)
        output_dir = "generated_pdfs"
        os.makedirs(output_dir, exist_ok=True)

        for _, row in df.iterrows():
            name = row['이름']
            m1_nums = str(row['Module1']).split(',') if pd.notna(row['Module1']) else []
            m2_nums = str(row['Module2']).split(',') if pd.notna(row['Module2']) else []
            m1_list = [m1_imgs[num.strip()] for num in m1_nums if num.strip() in m1_imgs]
            m2_list = [m2_imgs[num.strip()] for num in m2_nums if num.strip() in m2_imgs]
            pdf_path = create_student_pdf(name, m1_list, m2_list, doc_title, output_dir)
            generated_files.append((name, pdf_path))

        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "w") as zipf:
            for name, path in generated_files:
                zipf.write(path, os.path.basename(path))
        zip_buffer.seek(0)

        st.success("✅ 오답노트 PDF 생성 완료!")
        st.download_button("📁 ZIP 파일 다운로드", zip_buffer, file_name="오답노트_모음.zip")

    except Exception as e:
        st.error(f"오류 발생: {e}")

if generated_files:
    st.markdown("---")
    st.header("👁️ 개별 PDF 미리보기")
    selected = st.selectbox("학생 선택", [name for name, _ in generated_files])
    if selected:
        generated_dict = {name: path for name, path in generated_files}
        selected_path = generated_dict[selected]
        with open(selected_path, "rb") as f:
            st.download_button(f"📄 {selected} PDF 다운로드", f, file_name=f"{selected}.pdf")
