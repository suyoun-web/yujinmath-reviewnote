import streamlit as st
import zipfile
import os
import tempfile
import pandas as pd
from PIL import Image
import io
from fpdf import FPDF
import base64
import shutil

def load_example_excel():
    with open("예시_오답노트_양식_수정본.xlsx", "rb") as f:
        return f.read()

def save_uploaded_file(uploaded_file, save_path):
    with open(save_path, "wb") as f:
        f.write(uploaded_file.read())

def extract_images(zip_file, temp_dir):
    with zipfile.ZipFile(zip_file, 'r') as zip_ref:
        zip_ref.extractall(temp_dir)

def collect_images(folder_path):
    image_extensions = ['.png', '.jpg', '.jpeg']
    images = []
    for root, _, files in os.walk(folder_path):
        for file in sorted(files):
            if any(file.lower().endswith(ext) for ext in image_extensions):
                images.append(os.path.join(root, file))
    return images

def create_student_pdf(name, module1_imgs, module2_imgs, doc_title, output_dir):
    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=15)

    if module1_imgs:
        pdf.add_page()
        pdf.set_font("Arial", 'B', 16)
        pdf.cell(0, 10, f"{name} - {doc_title} - Module 1", ln=True)
        for img_path in module1_imgs:
            pdf.add_page()
            pdf.image(img_path, x=10, y=20, w=pdf.w - 20)

    if module2_imgs:
        pdf.add_page()
        pdf.set_font("Arial", 'B', 16)
        pdf.cell(0, 10, f"{name} - {doc_title} - Module 2", ln=True)
        for img_path in module2_imgs:
            pdf.add_page()
            pdf.image(img_path, x=10, y=20, w=pdf.w - 20)

    pdf_output = os.path.join(output_dir, f"{name}_{doc_title}.pdf")
    pdf.output(pdf_output)
    return pdf_output

def create_download_link(zip_path):
    with open(zip_path, "rb") as f:
        bytes_data = f.read()
    b64 = base64.b64encode(bytes_data).decode()
    href = f'<a href="data:application/zip;base64,{b64}" download="오답노트_전체.zip">📦 전체 ZIP 다운로드</a>'
    return href

def main():
    st.title("📝 SAT 오답노트 생성기")

    st.subheader("📊 예시 엑셀 양식")
    with st.expander("예시 엑셀파일 양식 보기"):
        st.write("아래 예시와 같이 이름, Module1, Module2 컬럼만 포함해주세요.")
        example_df = pd.read_excel(io.BytesIO(load_example_excel()))
        st.dataframe(example_df)

    st.markdown("[📥 예시 엑셀파일 다운로드](sandbox:/mnt/data/예시_오답노트_양식_수정본.xlsx)")

    st.subheader("📦 오답노트 파일 업로드")
    st.caption("M1, M2 폴더 포함된 ZIP 파일 업로드")
    image_zip_file = st.file_uploader("Drag and drop file here", type="zip")

    st.caption("오답노트 엑셀 파일 업로드 (.xlsx)")
    excel_file = st.file_uploader("Drag and drop file here", type="xlsx")

    doc_title = st.text_input("🖋️ 문서 제목 (예: 25 SAT MATH S2 만점반 Mock3)")

    if st.button("📎 오답노트 자동 생성"):
        if not image_zip_file or not excel_file or not doc_title:
            st.error("모든 파일과 문서 제목을 입력해주세요.")
            return

        with tempfile.TemporaryDirectory() as temp_dir:
            zip_path = os.path.join(temp_dir, "images.zip")
            save_uploaded_file(image_zip_file, zip_path)
            extract_images(zip_path, temp_dir)

            m1_path = os.path.join(temp_dir, "M1")
            m2_path = os.path.join(temp_dir, "M2")

            module1_images = collect_images(m1_path)
            module2_images = collect_images(m2_path)

            df = pd.read_excel(excel_file)
            output_dir = os.path.join(temp_dir, "output")
            os.makedirs(output_dir, exist_ok=True)

            student_pdfs = []

            for _, row in df.iterrows():
                name = str(row['이름'])
                m1_nums = str(row['Module1']).split(',') if pd.notna(row['Module1']) else []
                m2_nums = str(row['Module2']).split(',') if pd.notna(row['Module2']) else []

                m1_imgs = [img for img in module1_images if any(img.lower().endswith(f"{num.strip()}.png") or img.lower().endswith(f"{num.strip()}.jpg") or img.lower().endswith(f"{num.strip()}.jpeg") for num in m1_nums)]
                m2_imgs = [img for img in module2_images if any(img.lower().endswith(f"{num.strip()}.png") or img.lower().endswith(f"{num.strip()}.jpg") or img.lower().endswith(f"{num.strip()}.jpeg") for num in m2_nums)]

                pdf_path = create_student_pdf(name, m1_imgs, m2_imgs, doc_title, output_dir)

                with open(pdf_path, "rb") as f:
                    st.download_button(f"📄 {name} 오답노트 미리보기", f, file_name=os.path.basename(pdf_path), mime="application/pdf")

                student_pdfs.append(pdf_path)

            final_zip = os.path.join(temp_dir, "오답노트_전체.zip")
            with zipfile.ZipFile(final_zip, 'w') as zipf:
                for pdf_path in student_pdfs:
                    zipf.write(pdf_path, os.path.basename(pdf_path))

            st.markdown(create_download_link(final_zip), unsafe_allow_html=True)

if __name__ == "__main__":
    main()
