import os
from os import listdir
from os import makedirs
from shutil import copyfile

path = os.getcwd()
path_copy = path + "\\copy"
print(path_copy)
files = listdir(path)
makedirs(path_copy)

for file in files:
  file_name = file.split('.')[0]
  file_ext = file.split('.')[-1]
  if file_ext == 'jpg':
    for i in range(1, 5):
      copyfile(file, path_copy + '\\' + file_name + '_상세_' + str(i) + '.' + file_ext)
      # copyfile(file, path_copy + '\\' + file_name + '_라벨' + str(i) + '.' + file_ext)






# import openpyxl

# wb = openpyxl.Workbook()
# excel_file_name = path_copy + "\\text.xlsx"
# wb.save(excel_file_name)

# files = listdir(path_copy)

# cnt = 1
# text = ""

# for file in files:
#   if "_1.jpg" in file:
#     text = ", ".join(file)
#     if cnt % 4 == 0:
#       print(text)
#       text = ""
#     if cnt % 5 == 0:
#       print(file)
#       cnt = 1
#   cnt += 1
      

