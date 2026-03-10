import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.drawing.image import Image
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from PIL import Image as PILImage
import io

# 앱 제목 및 설정
st.set_page_config(page_title="점포 입회 리포트", layout="centered")
st.title("🏪 프리미엄 점포 점검 리포트")
st.write("정보를 입력하면 디자인된 엑셀 파일이 생성됩니다.")

# 1. 정보 입력 섹션
with st.form("check_form"):
    col1, col2 = st.columns(2)
    with col1:
        store_name = st.text_input("📍 점포 이름", placeholder="예: 강남역점")
    with col2:
        check_date = st.date_input("📅 점검 일자")
    
    check_content = st.text_area("📝 점검 내용", placeholder="점검 특이사항을 적어주세요.")
    uploaded_file = st.file_uploader("📸 점검 사진 업로드", type=["jpg", "jpeg", "png"])
    submit_button = st.form_submit_button("✨ 디자인 엑셀 생성")

# 2. 디자인 엑셀 생성 로직
if submit_button:
    if store_name and check_content:
        wb = Workbook()
        ws = wb.active
        ws.title = "점검결과"

        # 스타일 정의 (색상, 글꼴, 테두리)
        header_fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
        header_font = Font(color="FFFFFF", bold=True, size=12)
        center_alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), 
                            top=Side(style='thin'), bottom=Side(style='thin'))

        # 헤더 작성 및 스타일 적용
        headers = ["점검 일자", "점포 이름", "점검 내용", "현장 사진"]
        for col_num, header in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col_num)
            cell.value = header
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = center_alignment
            cell.border = thin_border

        # 데이터 입력 및 스타일 적용
        data = [str(check_date), store_name, check_content]
        for col_num, value in enumerate(data, 1):
            cell = ws.cell(row=2, column=col_num)
            cell.value = value
            cell.alignment = center_alignment
            cell.border = thin_border

        # 열 너비 조절
        ws.column_dimensions['A'].width = 15
        ws.column_dimensions['B'].width = 20
        ws.column_dimensions['C'].width = 40
        ws.column_dimensions['D'].width = 35

        # 사진 처리
        if uploaded_file is not None:
            img = PILImage.open(uploaded_file)
            img = img.resize((250, 200)) # 사진 크기 조절
            
            img_byte_arr = io.BytesIO()
            img.save(img_byte_arr, format='PNG')
            img_for_excel = Image(img_byte_arr)
            
            ws.add_image(img_for_excel, 'D2')
            ws.row_dimensions[2].height = 160 # 행 높이 확보
            
            # 사진 셀 테두리
            ws.cell(row=2, column=4).border = thin_border

        # 메모리에 저장
        excel_data = io.BytesIO()
        wb.save(excel_data)
        excel_data.seek(0)

        st.success("✅ 리포트 디자인이 완료되었습니다!")
        st.download_button(
            label="📁 디자인된 엑셀 다운로드",
            data=excel_data,
            file_name=f"{store_name}_점검리포트_{check_date}.xlsx"
        )
    else:
        st.error("⚠️ 필수 항목(점포명, 내용)을 입력해주세요.")