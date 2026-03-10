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

       # 2. 사진 배치 (더 안전한 방식으로 수정)
        if uploaded_files:
            current_row = len(content) + 2 # 데이터 아래 한 줄 띄우고 시작
            
            for uploaded_file in uploaded_files:
                # 라벨 표시
                ws.cell(row=current_row, column=1).value = "현장사진"
                ws.cell(row=current_row, column=1).fill = label_fill
                ws.cell(row=current_row, column=1).border = border
                
                # 이미지 처리
                img = PILImage.open(uploaded_file)
                # 엑셀에 적당한 크기로 리사이즈 (가로 약 300픽셀)
                img.thumbnail((300, 300))
                
                img_byte_arr = io.BytesIO()
                # PNG 포맷으로 확실히 지정해서 저장
                img.save(img_byte_arr, format='PNG')
                img_byte_arr.seek(0) # 커서를 맨 앞으로
                
                img_for_excel = Image(img_byte_arr)
                
                # B열 해당 행에 사진 삽입
                ws.add_image(img_for_excel, f'B{current_row}')
                
                # 사진이 보일 수 있게 행 높이를 대폭 키움 (중요!)
                ws.row_dimensions[current_row].height = 230 
                ws.cell(row=current_row, column=2).border = border
                
                current_row += 1 # 다음 사진은 아래 행에
