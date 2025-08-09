# app.py
import streamlit as st
import pandas as pd
import zipfile
import os
import io
import re
from PIL import Image
from fpdf import FPDF
from datetime import datetime

# ---------------------------
# 공통 설정
# ---------------------------
st.set_page_config(page_title="SAT 오답노트 & 통계 생성기", layout="centered")
FONT_REGULAR = "fonts/NanumGothic.ttf"
FONT_BOLD = "fonts/NanumGothicBold.ttf"
pdf_font_name = "NanumGothic"

# ---------------------------
# PDF 클래스 (오답노트용, 한글 폰트 + 여백)
# ---------------------------
class KoreanPDF(FPDF):
    def __init__(self):
        super().__init__()
        # 좌/우 2.54cm(25.4mm), 상 3cm(30mm), 하 2.54cm
        self.set_margins(25.4, 30.0, 25.4)
        self.set_auto_page_break(auto=True, margin=25.4)
        if os.path.exists(FONT_REGULAR) and os.path.exists(FONT_BOLD):
            self.add_font(pdf_font_name, '', FONT_REGULAR, uni=True)
            self.add_font(pdf_font_name, 'B', FONT_BOLD, uni=True)
            self.set_font(pdf_font_name, size=10)

# ---------------------------
# 유틸: 예시 엑셀(입력용)
# ---------------------------
def get_example_input_excel():
    output = io.BytesIO()
    example_df = pd.DataFrame({
        '이름': ['홍길동', '김철수', '이영희'],
        'Module1': ['1,3,5', 'X', None],   # X=응시/오답0, None=미응시
        'Module2': ['2,6', '1,3', 'X']
    })
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        example_df.to_excel(writer, index=False, sheet_name="예시")
    output.seek(0)
    return output

# ---------------------------
# 유틸: ZIP에서 이미지 읽기 (오답노트용)
# ---------------------------
def extract_zip_to_dict(zip_file):
    m1_imgs, m2_imgs = {}, {}
    with zipfile.ZipFile(zip_file) as z:
        for file in z.namelist():
            if file.lower().endswith(('.png', '.jpg', '.jpeg', '.webp')):
                parts = file.split('/')
                if len(parts) < 2:
                    continue
                folder = parts[0].lower()
                q_num = os.path.splitext(parts[-1])[0]
                with z.open(file) as f:
                    img = Image.open(f).convert("RGB")
                    if folder == "m1":
                        m1_imgs[q_num] = img
                    elif folder == "m2":
                        m2_imgs[q_num] = img
    return m1_imgs, m2_imgs

# ---------------------------
# 유틸: 학생 PDF 생성 (오답노트)
# ---------------------------
def create_student_pdf(name, m1_imgs, m2_imgs, doc_title, output_dir):
    pdf = KoreanPDF()
    pdf.add_page()
    if os.path.exists(FONT_REGULAR) and os.path.exists(FONT_BOLD):
        pdf.set_font(pdf_font_name, style='B', size=10)
    pdf.cell(0, 8, txt=f"<{name}_{doc_title}>", ln=True)

    def add_images(module_tag, images):
        # Module2 제목이 바닥에 걸리면 제목+이미지를 다음 페이지로
        est_img_h = 100
        if module_tag == "<Module2>" and pdf.get_y() + 10 + est_img_h > pdf.page_break_trigger:
            pdf.add_page()

        if os.path.exists(FONT_REGULAR):
            pdf.set_font(pdf_font_name, size=10)
        pdf.cell(0, 8, txt=module_tag, ln=True)

        if images:
            for img in images:
                tmp = f"/tmp/{datetime.now().timestamp()}.jpg"
                img.save(tmp)
                pdf.image(tmp, w=180)  # 여백 고려한 폭
                try:
                    os.remove(tmp)
                except:
                    pass
                pdf.ln(8)
        else:
            pdf.ln(8)

    add_images("<Module1>", m1_imgs)
    add_images("<Module2>", m2_imgs)

    os.makedirs(output_dir, exist_ok=True)
    path = os.path.join(output_dir, f"{name}_{doc_title}.pdf")
    pdf.output(path)
    return path

# ---------------------------
# 유틸: 모듈 셀 파싱 (공통)
# None/빈칸 -> None(미응시), 'X' -> [] (응시/오답0), '1,2,5' -> [1,2,5]
# ---------------------------
def parse_wrong_list(cell):
    if pd.isna(cell) or (isinstance(cell, str) and cell.strip() == ""):
        return None
    s = str(cell).strip()
    if s.lower() == "x":
        return []
    nums = []
    for tok in s.split(","):
        tok = tok.strip()
        if re.fullmatch(r"\d+", tok):
            nums.append(int(tok))
    return nums

# ---------------------------
# 세션 상태
# ---------------------------
if 'generated_files' not in st.session_state:
    st.session_state.generated_files = []
if 'zip_buffer' not in st.session_state:
    st.session_state.zip_buffer = None

# ---------------------------
# UI: 탭 구성
# ---------------------------
tab1, tab2 = st.tabs(["📝 오답노트 생성", "📊 문제별 오답률 (별도 생성)"])

# =========================================================
# 탭 1: 오답노트 생성 (기존처럼 독립 동작)
# =========================================================
with tab1:
    st.subheader("문서 제목")
    doc_title = st.text_input("예: 25 S2 SAT MATH 만점반 Mock Test1", value="25 S2 SAT MATH 만점반 Mock Test1")

    st.subheader("문제 ZIP / 오답 Excel 업로드")
    st.caption("ZIP은 최상단에 M1, M2 폴더를 포함하고, 각 폴더에 문제 이미지(파일명=문항번호)가 있어야 합니다.")
    img_zip = st.file_uploader("문제 ZIP 파일", type="zip")

    st.caption("엑셀 열: '이름', 'Module1', 'Module2'  | 값: 1,3,5 (콤마 구분) / 오답 없음= 'X' / 미응시=빈칸")
    excel_file = st.file_uploader("오답 현황 엑셀 (.xlsx)", type="xlsx")

    st.caption("예시 엑셀 미리보기/다운로드")
    with st.expander("입력 예시 보기"):
        st.dataframe(pd.read_excel(get_example_input_excel()))
    st.download_button("📥 입력 예시 엑셀 다운로드", get_example_input_excel(), file_name="예시_오답현황_양식.xlsx")

    if st.button("📎 오답노트 생성"):
        if not img_zip or not excel_file:
            st.warning("ZIP 파일과 엑셀 파일을 모두 업로드해주세요.")
        else:
            try:
                m1_imgs_all, m2_imgs_all = extract_zip_to_dict(img_zip)
                df = pd.read_excel(excel_file)

                out_dir = "generated_pdfs"
                os.makedirs(out_dir, exist_ok=True)
                st.session_state.generated_files = []

                for _, row in df.iterrows():
                    name = row['이름']

                    # Module1 또는 Module2가 비어있으면 생성 스킵
                    if pd.isna(row['Module1']) or pd.isna(row['Module2']):
                        continue

                    m1_nums = parse_wrong_list(row['Module1'])
                    m2_nums = parse_wrong_list(row['Module2'])

                    m1_list = [m1_imgs_all.get(str(n)) for n in (m1_nums or []) if str(n) in m1_imgs_all]
                    m2_list = [m2_imgs_all.get(str(n)) for n in (m2_nums or []) if str(n) in m2_imgs_all]

                    pdf_path = create_student_pdf(name, m1_list, m2_list, doc_title, out_dir)
                    st.session_state.generated_files.append((name, pdf_path))

                buf = io.BytesIO()
                with zipfile.ZipFile(buf, "w") as zipf:
                    for name, path in st.session_state.generated_files:
                        zipf.write(path, os.path.basename(path))
                buf.seek(0)
                st.session_state.zip_buffer = buf

                st.success("✅ 오답노트 PDF 생성 완료!")
            except Exception as e:
                st.error(f"오류 발생: {e}")

    if st.session_state.zip_buffer:
        st.download_button("📁 전체 ZIP 다운로드", st.session_state.zip_buffer, file_name="오답노트_모음.zip")

    if st.session_state.generated_files:
        st.markdown("---")
        st.subheader("개별 PDF 다운로드")
        selected = st.selectbox("학생 선택", [name for name, _ in st.session_state.generated_files])
        if selected:
            target = dict(st.session_state.generated_files)[selected]
            with open(target, "rb") as f:
                st.download_button(f"📄 {selected} PDF 다운로드", f, file_name=os.path.basename(target))

# =========================================================
# 탭 2: 문제별 오답률 (별도 생성/다운로드)
# =========================================================
with tab2:
    st.subheader("오답률 통계 생성")
    exam_title = st.text_input("통계 제목 입력 (예: 8월 Final mock 1)", value="8월 Final mock 1")

    st.caption("엑셀 열: '이름', 'Module1', 'Module2'  | 값: 1,3,7 / 오답 없음= 'X' / 미응시=빈칸")
    stat_file = st.file_uploader("통계용 엑셀 업로드 (.xlsx)", type=["xlsx"], key="stats_uploader")

    with st.sidebar:
        st.markdown("### ⚙️ 통계 설정")
        m1_total = st.number_input("Module1 총 문항 수", 1, 200, 22, 1, key="m1_total")
        m2_total = st.number_input("Module2 총 문항 수", 1, 200, 22, 1, key="m2_total")

    def compute_module_rates(series, total_questions):
        # 응시자(분모): None이 아닌 학생
        attempted = series.apply(lambda v: v is not None).sum()
        wrong_counts = {q: 0 for q in range(1, total_questions+1)}
        for v in series:
            if isinstance(v, list):
                for q in v:
                    if 1 <= q <= total_questions:
                        wrong_counts[q] += 1
        rows = []
        for q in range(1, total_questions+1):
            w = wrong_counts[q]
            rate = (w / attempted) if attempted > 0 else 0.0
            rows.append({"문제 번호": q, "오답률(%)": round(rate*100, 2), "틀린 학생 수": int(w)})
        return pd.DataFrame(rows)

    if stat_file is not None:
        try:
            stat_df = pd.read_excel(stat_file)
            stat_df["M1_parsed"] = stat_df["Module1"].apply(parse_wrong_list)
            stat_df["M2_parsed"] = stat_df["Module2"].apply(parse_wrong_list)

            m1_tbl = compute_module_rates(stat_df["M1_parsed"], int(m1_total))
            m2_tbl = compute_module_rates(stat_df["M2_parsed"], int(m2_total))

            # m1-1..m1-22, m2-1..m2-22 한 시트에 이어 붙이기
            m1_tbl = m1_tbl.rename(columns={"문제 번호": "문제 번호"})
            m1_tbl.insert(0, "문제 번호", m1_tbl["문제 번호"].apply(lambda x: f"m1-{x}"))
            m2_tbl.insert(0, "문제 번호", m2_tbl["문제 번호"].apply(lambda x: f"m2-{x}"))

            combined = pd.concat([m1_tbl[["문제 번호", "오답률(%)", "틀린 학생 수"]],
                                  m2_tbl[["문제 번호", "오답률(%)", "틀린 학생 수"]]],
                                  ignore_index=True)

            st.dataframe(combined, use_container_width=True)

            # 엑셀로 내보내기 (제목 행 + 가운데 정렬 + 조건부서식: 오답률>=30 bold+size15)
            out = io.BytesIO()
            with pd.ExcelWriter(out, engine="xlsxwriter") as writer:
                sheet_name = "오답률 통계"
                combined.to_excel(writer, index=False, sheet_name=sheet_name, startrow=2)

                wb = writer.book
                ws = writer.sheets[sheet_name]

                # 제목 행 (A1에 <제목> 형태)
                title_text = f"<{exam_title}>"
                title_fmt = wb.add_format({"bold": True, "align": "center", "valign": "vcenter"})
                ws.write(0, 0, title_text, title_fmt)
                # 제목 행: A1~C1 병합 + 가운데 정렬
                ws.merge_range(0, 0, 0, 2, title_text, title_fmt)

                # 헤더 행 포맷
                header_fmt = wb.add_format({"bold": True, "align": "center", "valign": "vcenter"})
                ws.write(2, 0, "문제 번호", header_fmt)
                ws.write(2, 1, "오답률(%)", header_fmt)
                ws.write(2, 2, "틀린 학생 수", header_fmt)

                # 데이터 가운데 정렬
                center_fmt = wb.add_format({"align": "center", "valign": "vcenter"})
                # 전체 열 가운데 정렬
                ws.set_column(0, 2, 14, center_fmt)

                # 조건부 서식: 오답률(%) >= 30 → Bold + font size 15
                cond_fmt = wb.add_format({"bold": True, "font_size": 15, "align": "center", "valign": "vcenter"})
                start_row = 3  # 데이터 시작(0-index)
                end_row = 3 + len(combined) - 1
                # 오답률(%) 열 = 컬럼 1
                if len(combined) > 0:
                    ws.conditional_format(start_row, 1, end_row, 1, {
                        "type": "cell",
                        "criteria": ">=",
                        "value": 30,
                        "format": cond_fmt
                    })

            out.seek(0)
            st.download_button(
                "📥 오답률 통계 엑셀 다운로드",
                data=out,
                file_name=f"오답률_통계_{exam_title}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

            st.success("✅ 통계 엑셀을 생성했습니다.")
            st.info("오답률 = (틀린 학생 수) / (해당 모듈을 푼 학생 수)\n- 'X'는 응시했지만 오답 0개로 처리됩니다.\n- 빈 칸/NaN은 미응시로 간주되어 분모에서 제외됩니다.")

        except Exception as e:
            st.error(f"처리 중 오류가 발생했습니다: {e}")
    else:
        st.caption("통계를 따로 만들고 싶을 때 이 탭에서 엑셀을 업로드해 주세요.")

