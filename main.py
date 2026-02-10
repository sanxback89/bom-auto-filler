"""
BOM PDF → Excel Template Auto Filler
메인 실행 파일

모듈 구조:
  utils.py          - 텍스트 정제 유틸리티
  models.py         - BomRow 데이터 모델
  image_handler.py  - 이미지 추출/삽입
  pdf_parser.py     - PDF 파싱 (Master, BOM Details, ColorMatrix)
  excel_template.py - 엑셀 템플릿 탐색/스타일 헬퍼
  excel_writer.py   - fill_template 메인 로직
  gui.py            - tkinter GUI
"""
from gui import App


if __name__ == "__main__":
    app = App()
    app.mainloop()
