from openpyxl import Workbook
from openpyxl import load_workbook
from openpyxl.drawing.image import Image
from openpyxl.styles import PatternFill
from openpyxl.styles import Alignment
from openpyxl.styles.fonts import Font
from openpyxl.utils import get_column_letter
import os
from os import listdir
from os.path import exists
from os import makedirs
import productsCode

# 폴더 내 엑셀 파일 검색
currPath = os.getcwd()
files = listdir(currPath)
excelFileList = []

for i in files:
  if(i.split('.')[-1] == 'xlsx'):
    if(not i.startswith('~')):
      excelFileList.append(i)
      
productCode = productsCode.productCode

for file in excelFileList:

  wb = load_workbook(currPath + '\\' + file)
  wb.create_sheet('상품사진삽입')

  ws_name = wb.get_sheet_names()

  sheet1 = wb[str(ws_name[0])]
  sheet2 = wb[str(ws_name[1])]
  
  wb.active = wb['상품사진삽입']
  
  fillAlignment = Alignment(horizontal='center')
  fillFont = Font(bold=True)
  
  sheet2.cell(1, 1).value = '주문건수(상품기준)'
  sheet2.cell(1, 2).value = '판매량'
  sheet2.cell(1, 3).value = '대표이미지'
  sheet2.cell(1, 1).alignment = fillAlignment
  sheet2.cell(1, 2).alignment = fillAlignment
  sheet2.cell(1, 3).alignment = fillAlignment
  sheet2.cell(1, 1).font = fillFont
  sheet2.cell(1, 2).font = fillFont
  sheet2.cell(1, 3).font = fillFont

  sheet2.column_dimensions['A'].width = 40
  sheet2.column_dimensions['B'].width = 10
  sheet2.column_dimensions['C'].width = 12

  fillData = PatternFill(fill_type='solid', start_color='FFCCCC', end_color='FFCCCC')
  fillData2 = PatternFill(fill_type='solid', start_color='FFFF00', end_color='FFFF00')
  
  sheet2["A1"].fill = fillData
  sheet2["B1"].fill = fillData
  sheet2["C1"].fill = fillData

  for sheet in wb:
    if sheet.title == '상품사진삽입':
      sheet.sheet_view.tabSelected = True
    else:
      sheet.sheet_view.tabSelected = False
      
  first_row = 2
  last_row = sheet1.max_row + 1  
  
  for i in range(first_row, last_row):
    if sheet1.cell(row=i, column=4).value == None or sheet1.cell(row=i, column=4).value == '':
      continue
    else:
      sheet2["A" + str(i)].alignment = Alignment(vertical='center')
      sheet2["B" + str(i)].alignment = Alignment(vertical='center')
      print(sheet1.cell(row=i, column=4).value + " / 판매수량 : " + str(sheet1.cell(row=i, column=5).value))
      sheet2.cell(row=i, column=1).value = sheet1.cell(row=i, column=4).value
      sheet2.cell(row=i, column=2).value = sheet1.cell(row=i, column=5).value
      try:
        sheet2.row_dimensions[i].height = 75
        # image_path = 'https://gi.esmplus.com/jja6806/thumbnail/{}.jpg'.format(sheet2.cell(row=i, column=3).value)
        image_path = '.\\data\\images\\' + str(productCode[sheet2.cell(row=i, column=1).value]) + '.jpg'
        image = Image(image_path)
        image.width = 100
        image.height = 100
        sheet2.add_image(image, anchor='C'+str(i))
      except:
        sheet2.cell(i, 1).fill = fillData2
        sheet2.cell(i, 2).fill = fillData2
        sheet2.cell(i, 3).fill = fillData2
  
  wb.save(currPath + '\\' + file)