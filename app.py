import streamlit as st
import pandas as pd
import zipfile
import os
import io
import re
from PIL import Image
from fpdf import FPDF
from datetime import datetime

# ==============================
# 공통 설정
# ==============================
st.set_page_config(page_title="SAT 오답노트 & 통계 생성기", layout="centered")

# PDF 생성용 폰트 경로
FONT_REGULAR = "fonts/NanumGothic.ttf"
FONT_BOLD = "fonts/NanumGothicBold.ttf"
pdf_font_name = "NanumGothic"

if os.path.exists(FONT_REGULAR) and os.path.exists(FONT_BOLD):
    class KoreanPDF(FPDF):
        def __init__(self):
            super().__init__()
            # 좌/우 2.54cm(25.4mm), 위 3.0cm(30mm) 여백
            self.set_margins(25.4, 30, 25.4)
            self.add_font(pdf_font_name, '', FONT_REGULAR, uni=True)
            self.add_font(pdf_font_name, 'B', FONT_BOLD, uni=True)
            self.set_font(pdf_font_name, size=10)
else:
    st.error("⚠️ 한글 PDF 생성을 위해 fonts 폴더에 NanumGothic.ttf 와 NanumGothicBold.ttf 모두 필요합니다.")

# ==============================
# 유틸(오답노트) : 예시 엑셀
# ==============================
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

# ==============================
# 유틸(오답노트) : ZIP 파싱
# ==============================
def extract_zip_to_dict(zip_file):
    m1_imgs, m2_imgs = {}, {}
    with zipfile.ZipFile(zip_file) as z:
        for file in z.namelist():
            if file.lower().endswith(('png', 'jpg', 'jpeg', 'webp')):
                parts = file.split('/')
                if len(parts) < 2:
                    continue
                folder = parts[0].lower()
                q_num = os.path.splitext(os.path.basename(file))[0]
                with z.open(file) as f:
                    img = Image.open(f).convert("RGB")
                    if folder == "m1":
                        m1_imgs[q_num] = img
                    elif folder == "m2":
                        m2_imgs[q_num] = img
    return m1_imgs, m2_imgs

# ==============================
# 유틸(오답노트) : PDF 생성
# ==============================
def create_student_pdf(name, m1_imgs, m2_imgs, doc_title, output_dir):
    pdf = KoreanPDF()
    pdf.add_page()
    pdf.set_font(pdf_font_name, style='B', size=10)
    pdf.cell(0, 8, txt=f"<{name}_{doc_title}>", ln=True)

    def add_images(title, images):
        img_est_height = 100
        # Module2 제목이 바닥에 걸리면 제목+이미지를 다음 페이지에 붙여 시작
        if title == "<Module2>" and pdf.get_y() + 10 + (img_est_height if images else 0) > pdf.page_break_trigger:
            pdf.add_page()

        pdf.set_font(pdf_font_name, size=10)
        pdf.cell(0, 8, txt=title, ln=True)
        if images:
            for img in images:
                img_path = f"temp_{datetime.now().timestamp()}.jpg"
                img.save(img_path)
                pdf.image(img_path, w=180)
                try:
                    os.remove(img_path)
                except:
                    pass
                pdf.ln(8)

    # 이미지가 없어도 모듈 제목은 항상 출력
    add_images("<Module1>", m1_imgs)
    add_images("<Module2>", m2_imgs)

    os.makedirs(output_dir, exist_ok=True)
    pdf_path = os.path.join(output_dir, f"{name}_{doc_title}.pdf")
    pdf.output(pdf_path)
    return pdf_path

# ==============================
# 유틸(통계) : 입력 파싱 & 오답률 계산
# ==============================
def robust_parse_wrong_list(cell):
    """None/빈칸 -> None(미응시), 'X' -> [] (응시/오답 0), '1,2,5' -> [1,2,5]"""
    if pd.isna(cell) or str(cell).strip() == "":
        return None
    s = str(cell).strip()
    if s.lower() == "x":
        return []
    s = s.replace("，", ",").replace(";", ",")
    tokens = [t.strip() for t in s.split(",") if t.strip() != ""]
    nums = []
    for t in tokens:
        if re.fullmatch(r"\d+", t):
            nums.append(int(t))
    return nums

def compute_module_rates(series, total_questions):
    """오답률(%) = (틀린 학생 수 / 응시자 수) * 100  (응시자: None이 아닌 학생)"""
    attempted = series.apply(lambda v: v is not None).sum()
    wrong_counts = {q: 0 for q in range(1, total_questions+1)}
    for v in series:
        if isinstance(v, list):
            for q in v:
                if 1 <= q <= total_questions:
                    wrong_counts[q] += 1

    rows = []
    for q in range(1, total_questions+1):
        wrong = wrong_counts[q]
        rate = round((wrong / attempted) * 100, 1) if attempted > 0 else 0.0
        rows.append({"문제 번호": q, "오답률(%)": rate, "틀린 학생 수": int(wrong)})
    return pd.DataFrame(rows)

# ==============================
# UI - 탭 구성
# ==============================
tab1, tab2 = st.tabs(["📝 오답노트 생성기", "📊 오답률 통계 생성기"])

# =========================================================
# 탭 1: 오답노트 생성기 (기존 버전 기반)
# =========================================================
with tab1:
    st.title("📝 SAT 오답노트 생성기")

    st.header("📊 예시 엑셀 양식")
    with st.expander("예시 엑셀파일 열기"):
        st.dataframe(pd.read_excel(get_example_excel()))
    example = get_example_excel()
    st.download_button("📥 예시 엑셀파일 다운로드", example, file_name="예시_오답노트_양식.xlsx")

    st.header("📄 문서 제목 입력")
    doc_title = st.text_input("문서 제목 (예: 25 S2 SAT MATH 만점반 Mock Test1)", value="25 S2 SAT MATH 만점반 Mock Test1")

    st.header("📦 오답노트 파일 업로드")
    st.caption("M1, M2 폴더 포함된 ZIP 파일 업로드")
    img_zip = st.file_uploader("문제 ZIP 파일", type="zip")

    st.caption("오답노트 엑셀 파일 업로드 (.xlsx) — 열 이름은 '이름', 'Module1', 'Module2'")
    excel_file = st.file_uploader("오답 현황 엑셀", type="xlsx")

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

                # Module1 또는 Module2 중 하나라도 비어 있으면 건너뜀 (요청사항)
                if pd.isna(row['Module1']) or pd.isna(row['Module2']):
                    continue

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
        st.header("👁️ 개별 PDF 다운로드")
        selected = st.selectbox("학생 선택", [name for name, _ in generated_files])
        if selected:
            generated_dict = {name: path for name, path in generated_files}
            selected_path = generated_dict[selected]
            with open(selected_path, "rb") as f:
                st.download_button(f"📄 {selected} PDF 다운로드", f, file_name=f"{selected}.pdf")

# =========================================================
# 탭 2: 오답률 통계 생성기 (새 탭)
# =========================================================
with tab2:
    st.title("📊 오답률 통계 생성기")

    # 예시 엑셀/CSV 제공 (앱에서 보기/복사/다운로드)
    def example_df():
        return pd.DataFrame({
            "이름": ["홍길동", "김철수", "이영희", "박민수"],
            "Module1": ["1,3,5", "X", "2,4,7", ""],   # "" 또는 NaN = 미응시
            "Module2": ["2,6", "1,3", "X", "5"]
        })

    with st.expander("🧾 예시 입력 파일 보기 / 복사 / 다운로드"):
        ex = example_df()
        st.caption("열 이름은 반드시 **이름, Module1, Module2** 입니다. 값은 `1,3,5` 처럼 콤마로 구분하고, 오답이 없으면 `X`, 미응시는 빈칸으로 두세요.")
        st.dataframe(ex, use_container_width=True)
        csv_text = ex.to_csv(index=False)
        st.text_area("복사용 CSV", csv_text, height=160)
        buf_ex = io.BytesIO()
        with pd.ExcelWriter(buf_ex, engine="xlsxwriter") as w:
            ex.to_excel(w, index=False, sheet_name="예시")
        buf_ex.seek(0)
        st.download_button("📥 예시 엑셀 다운로드", buf_ex, file_name="예시_오답현황_양식.xlsx")

    # 통계 입력
    exam_title = st.text_input("통계 제목 입력 (예: 8월 Final mock 1)", value="8월 Final mock 1")
    col1, col2 = st.columns(2)
    with col1:
        m1_total = st.number_input("Module1 문제 수", min_value=1, value=22)
    with col2:
        m2_total = st.number_input("Module2 문제 수", min_value=1, value=22)

    stat_file = st.file_uploader("📂 학생 오답 현황 엑셀 업로드 (.xlsx)", type="xlsx", key="stats_uploader")

    if stat_file:
        try:
            df_stat = pd.read_excel(stat_file)
            # 컬럼 정규화/검증
            df_stat.columns = [str(c).strip() for c in df_stat.columns]
            required_cols = {"이름", "Module1", "Module2"}
            if not required_cols.issubset(df_stat.columns):
                st.error(f"엑셀에 {required_cols} 컬럼이 모두 있어야 합니다.")
                st.stop()

            # 파싱
            df_stat["M1_parsed"] = df_stat["Module1"].apply(robust_parse_wrong_list)
            df_stat["M2_parsed"] = df_stat["Module2"].apply(robust_parse_wrong_list)

            # 통계 계산
            m1_stats = compute_module_rates(df_stat["M1_parsed"], int(m1_total))
            m2_stats = compute_module_rates(df_stat["M2_parsed"], int(m2_total))
            m1_stats["문제 번호"] = m1_stats["문제 번호"].apply(lambda x: f"m1-{x}")
            m2_stats["문제 번호"] = m2_stats["문제 번호"].apply(lambda x: f"m2-{x}")

            combined = pd.concat([m1_stats, m2_stats], ignore_index=True)[["문제 번호", "오답률(%)", "틀린 학생 수"]]

            st.subheader("미리보기")
            st.dataframe(combined, use_container_width=True)

            # 엑셀 저장 (제목행 + 가운데정렬 + 오답률≥30% 강조)
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
                sheet_name = "오답률 통계"
                combined.to_excel(writer, index=False, sheet_name=sheet_name, startrow=2)
                wb = writer.book
                ws = writer.sheets[sheet_name]

                # 제목 행
                title_fmt = wb.add_format({"bold": True, "align": "center", "valign": "vcenter"})
                ws.merge_range(0, 0, 0, 2, f"<{exam_title}>", title_fmt)

                # 헤더
                header_fmt = wb.add_format({"bold": True, "align": "center", "valign": "vcenter"})
                ws.write(2, 0, "문제 번호", header_fmt)
                ws.write(2, 1, "오답률(%)", header_fmt)
                ws.write(2, 2, "틀린 학생 수", header_fmt)

                # 가운데 정렬
                center_fmt = wb.add_format({"align": "center", "valign": "vcenter"})
                ws.set_column(0, 2, 14, center_fmt)

                # 오답률 30% 이상 강조 (Bold + 폰트 15)
                cond_fmt = wb.add_format({"bold": True, "font_size": 15, "align": "center", "valign": "vcenter"})
                if len(combined) > 0:
                    ws.conditional_format(3, 1, 3 + len(combined) - 1, 1, {
                        "type": "cell", "criteria": ">=", "value": 30, "format": cond_fmt
                    })

            output.seek(0)
            st.download_button(
                "📥 오답률 통계 다운로드",
                data=output,
                file_name=f"오답률_통계_{exam_title}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            st.success("✅ 통계 엑셀을 생성했습니다.")
            st.info("오답률 = (틀린 학생 수) / (해당 모듈을 푼 학생 수)\n- 'X'는 응시했지만 오답 0개로 처리됩니다.\n- 빈 칸/NaN은 미응시로 간주되어 분모에서 제외됩니다.")
        except Exception as e:
            st.error(f"처리 중 오류가 발생했습니다: {e}")
    else:
        st.info("예시를 참고해 엑셀을 준비한 뒤 업로드하면 통계가 생성됩니다.")
