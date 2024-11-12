import os
from os import listdir
import json

# 폴더 내 json 파일 검색
currPath = os.getcwd()
files = listdir(currPath + "/trackingJSON")
jsonFileList = []

trackingInfo = {}
trackingInfo["data"] = []

for i in files:
  if(i.split(".")[-1] == "json"):
    if(not i.startswith("~")):
      jsonFileList.append(i)

idx = 0

for i in jsonFileList:
  with open(currPath + "/trackingJSON/" + i, "r", encoding="UTF-8") as f:
    jsonData = json.load(f)
    for prd in jsonData["data"]:
      trackingInfo["data"].append(prd)

with open("./trackingJSON/result/trackingTotal.json", "w", encoding="UTF-8") as outfile:
  json.dump(trackingInfo, outfile, indent=2, ensure_ascii=False)