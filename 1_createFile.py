from openpyxl import Workbook

wb = Workbook()         # 새 워크북 생성 (새 엑셀 오픈. 저장되지 않은 상태)
ws = wb.active          # 현재 활성화된 sheet를 가져옴. 이 위에서 작업!
ws.title = "HeraSheet"  # Sheet명 변경
wb.save("sample.xlsx")  # 저장
wb.close                # 워크북 종료