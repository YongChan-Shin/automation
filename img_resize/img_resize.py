from PIL import Image
import os
from os import listdir
from openpyxl import load_workbook

currPath = os.getcwd() + '\\data\\'
files = listdir(currPath)

for i in files:
  image1 = Image.open(currPath + i)
  new_image = Image.new('RGB', (image1.width, image1.height), color=(255, 255, 255))
  new_image.paste(image1, (0, 0))
  new_image.save('img\\' + i, 'JPEG', quality=95, subsampling=0)