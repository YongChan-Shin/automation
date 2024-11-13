from openpyxl import Workbook
from openpyxl import load_workbook

import json

wb = load_workbook("판매데이터.xlsx")
ws = wb["상품사진삽입"]

first_row = 2
last_row = ws.max_row + 1

# JSON 파일로 저장
jsonData = {}
jsonData["data"] = []
  
for i in range(first_row, last_row):
  ws.cell(i, 5).value = '"prdName": "{}", "salesCnt": {}'.format(ws.cell(i, 1).value, ws.cell(i, 2).value)
  jsonData["data"].append({
    "prdName": ws.cell(i, 1).value,
    "salesCnt": ws.cell(i, 2).value
  })
  
print(jsonData)

with open("salesData.json", "w", encoding="UTF-8") as outfile:
  json.dump(jsonData, outfile, indent=2, ensure_ascii=False)