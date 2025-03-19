from PIL import Image
import os
from os import listdir
from os.path import exists
from os import makedirs

currPath = os.getcwd()
files = listdir(currPath)

if not exists(currPath + '\\resize'):
  makedirs(currPath + '\\resize')

for i in files:
  if i.split('.')[-1] == 'jpg':
    image1 = Image.open(i)
    image1 = image1.resize((100, 100))
    new_image = Image.new('RGB', (image1.width, image1.height), color=(255, 255, 255))
    new_image.paste(image1, (0, 0))
    new_image.save('resize\\' + i, 'JPEG')