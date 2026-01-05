from PIL import Image
import os
from os import listdir
from openpyxl import load_workbook

############# 사전 준비 사항 #############
# data 폴더 생성 : 해당 폴더로 원본 이미지 복사
# img 폴더 생성 : 수정된 이미지 저장용 폴더

currPath = os.getcwd() + '\\data\\'
files = listdir(currPath)

print(currPath)
print(files)

for i in files:
  image1 = Image.open(currPath + i)
  new_image = Image.new('RGB', (image1.width, image1.height), color=(255, 255, 255))
  new_image.paste(image1, (0, 0))
  new_image.save('img\\' + i, 'JPEG', quality=95, subsampling=0)