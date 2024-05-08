from openpyxl import load_workbook
import os

wb = load_workbook('data.xlsx')
sheet1Name = '상품리스트'
sheet2Name = '사이즈추가'

currPath = os.getcwd()

firstCell = 2
lastCell = wb[sheet1Name].max_row + 1

wb.active = wb[sheet2Name]

cnt = 2

for i in range(firstCell, lastCell):
  if wb[sheet1Name].cell(i, 4).value == None or wb[sheet1Name].cell(i, 4).value == '':
    continue
  colorList = wb[sheet1Name].cell(i, 4).value.split('/')
  for color in colorList:
    sizeList = wb[sheet1Name].cell(i, 5).value.split(',')
    for size in sizeList:
      for j in range(1, 16):
        wb[sheet2Name].cell(cnt, j).value = wb[sheet1Name].cell(i, j).value
      wb[sheet2Name].cell(cnt, 4).value = color
      wb[sheet2Name].cell(cnt, 5).value = size
      cnt += 1

wb.save(currPath + '\\data_사이즈 추가본.xlsx')