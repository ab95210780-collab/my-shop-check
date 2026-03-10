import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.drawing.image import Image
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from PIL import Image as PILImage
import io

# 페이지 설정
st.set_page_config(page_title="점포 입회 점검 시스템", layout="centered")
st.title("📋 점포 입회 점검 리포트")

# 1. 정보 입력 섹션 (UI)
with st.form("inspection_form"):
    col1, col2 = st.columns(2)
    
    with col1:
        visit_date = st.date_input("📅 입회날짜")
        business_type = st.selectbox("🏬 업태", ["GMS", "SM", "APP", "기타"])
        company_name = st.text_input("🏢 기업명")
    
    with col2:
        store_name = st.text_input("🏪 점포명")
        sv_name = st.text_input("👤 SV명")
        visit_purpose = st.selectbox("🎯 입회목적", ["SV지도", "신규기업", "BC지도", "기타"])

    visit_result = st.text_area("📝 입회결과", placeholder="점검 결과 및 특이사항을 입력하세요.")
    uploaded_file = st.file_uploader("📸 현장 사진 업로드", type=["jpg", "jpeg", "png"])
    
    submit_button = st.form_submit_button("✨ 엑셀 리포트 생성")

# 2. 엑셀 생성 로직
if submit_button:
    if store_name and company_name and visit_result:
        wb = Workbook()
        ws = wb.active
        ws.title = "입회점검결과"

        # 디자인 스타일 설정
        header_fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
        header_font = Font(color="FFFFFF", bold=True)
        center_style = Alignment(horizontal="center", vertical="center", wrap_text=True)
        thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

        # 헤더 설정
        headers = ["입회날짜", "업태", "기업명", "점포명", "SV명", "입회목적", "입회결과", "현장사진"]
        for col_num, header in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col_num)
            cell.value = header
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = center_style
            cell.border = thin_border

        # 데이터 입력
        data = [str(visit_date), business_type, company_name, store_name, sv_name, visit_purpose, visit_result]
        for col_num, value in enumerate(data, 1):
            cell = ws.cell(row=2, column=col_num)
            cell.value = value
            cell.alignment = center_style
            cell.border = thin_border

        # 열 너비 설정
        widths = {'A': 15, 'B': 12, 'C': 20, 'D': 20, 'E': 12, 'F': 15, 'G': 45, 'H': 35}
        for col, width in widths.items():
            ws.column_dimensions[col].width = width

        # 사진 처리 (H열)
        if uploaded_file is not None:
            img = PILImage.open(uploaded_file)
            img = img.resize((250, 200))
            img_byte_arr = io.BytesIO()
            img.save(img_byte_arr, format='PNG')
            img_for_excel = Image(img_byte_arr)
            
            ws.add_image(img_for_excel, 'H2')
            ws.row_dimensions[2].height = 160
            ws.cell(row=2, column=8).border = thin_border

        # 메모리 저장 및 다운로드
        excel_data = io.BytesIO()
        wb.save(excel_data)
        excel_data.seek(0)
        
        st.success(f"✅ {store_name} 리포트 생성 완료!")
        st.download_button(
            label="📁 엑셀 파일 다운로드",
            data=excel_data,
            file_name=f"입회리포트_{store_name}_{visit_date}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.error("⚠️ 기업명, 점포명, 입회결과는 필수 입력 항목입니다.")
