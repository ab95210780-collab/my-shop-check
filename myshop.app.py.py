import streamlit as st
from openpyxl import Workbook
from openpyxl.drawing.image import Image
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from PIL import Image as PILImage
import io

st.set_page_config(page_title="세로형 점검 리포트", layout="centered")
st.title("📄 세로형 입회 점검 리포트")

with st.form("inspection_form"):
    visit_date = st.date_input("📅 입회날짜")
    business_type = st.selectbox("🏬 업태", ["GMS", "SM", "APP", "기타"])
    company_name = st.text_input("🏢 기업명")
    store_name = st.text_input("🏪 점포명")
    sv_name = st.text_input("👤 SV명")
    visit_purpose = st.selectbox("🎯 입회목적", ["업무확인", "신규오픈", "SV지도", "기타"])
    visit_result = st.text_area("📝 입회결과")
    uploaded_files = st.file_uploader("📸 사진 업로드 (여러 장 가능)", type=["jpg", "jpeg", "png"], accept_multiple_files=True)
    
    submit_button = st.form_submit_button("✨ 세로형 엑셀 생성")

if submit_button:
    if store_name and company_name:
        wb = Workbook()
        ws = wb.active
        ws.title = "점검리포트"

        # 스타일 정의
        label_fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid") # 연녹색 배경
        border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
        bold_font = Font(bold=True)
        center_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
        left_align = Alignment(horizontal="left", vertical="center", wrap_text=True)

        # 세로형 데이터 구성 (항목 이름, 데이터)
        content = [
            ("입회날짜", str(visit_date)),
            ("업태", business_type),
            ("기업명", company_name),
            ("점포명", store_name),
            ("SV명", sv_name),
            ("입회목적", visit_purpose),
            ("입회결과", visit_result)
        ]

        # 1. 항목 및 데이터 쓰기
        for i, (label, value) in enumerate(content, 1):
            # 항목명 (A열)
            ws.cell(row=i, column=1).value = label
            ws.cell(row=i, column=1).fill = label_fill
            ws.cell(row=i, column=1).font = bold_font
            ws.cell(row=i, column=1).alignment = center_align
            ws.cell(row=i, column=1).border = border
            
            # 내용 (B열)
            ws.cell(row=i, column=2).value = value
            ws.cell(row=i, column=2).alignment = left_align
            ws.cell(row=i, column=2).border = border
            
            # 행 높이 조절 (입회결과 부분은 더 높게)
            if label == "입회결과":
                ws.row_dimensions[i].height = 100
            else:
                ws.row_dimensions[i].height = 30

        ws.column_dimensions['A'].width = 20
        ws.column_dimensions['B'].width = 60

        # 2. 사진 배치 (데이터 아래쪽으로 세로로 쌓기)
        if uploaded_files:
            current_row = len(content) + 1
            ws.cell(row=current_row, column=1).value = "현장사진"
            ws.cell(row=current_row, column=1).fill = label_fill
            ws.cell(row=current_row, column=1).border = border
            ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=2)
            
            current_row += 1
            for uploaded_file in uploaded_files:
                img = PILImage.open(uploaded_file)
                # 가로 비율 유지하며 크기 조정
                img.thumbnail((400, 400)) 
                
                img_byte_arr = io.BytesIO()
                img.save(img_byte_arr, format='PNG')
                img_for_excel = Image(img_byte_arr)
                
                ws.add_image(img_for_excel, f'B{current_row}')
                ws.row_dimensions[current_row].height = 250
                current_row += 1

        excel_data = io.BytesIO()
        wb.save(excel_data)
        excel_data.seek(0)
        
        st.success("✅ 세로형 보고서가 생성되었습니다!")
        st.download_button(label="📁 세로형 엑셀 다운로드", data=excel_data, file_name=f"입회보고서_{store_name}.xlsx")
