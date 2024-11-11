import os
from os import listdir
import json

# 폴더 내 json 파일 검색
currPath = os.getcwd()
files = listdir(currPath + "/trackingJSON")
jsonFileList = []

trackingInfo = {}

for i in files:
  if(i.split(".")[-1] == "json"):
    if(not i.startswith("~")):
      jsonFileList.append(i)

idx = 0

for i in jsonFileList:
  with open(currPath + "/trackingJSON/" + i, "r", encoding="UTF-8") as f:
    jsonData = json.load(f)
    # print(jsonData["data"])
    for prd in jsonData["data"]:
      if prd["kidscomoPrdName"] not in trackingInfo:
        trackingInfo[prd["kidscomoPrdName"]] = {"kidscomoPrdName": prd["kidscomoPrdName"], "samsonyPrdName": prd["samsonyPrdName"], "kidscomoPrice": prd["kidscomoPrice"], "samsonyPrice": prd["samsonyPrice"], "priceGap": prd["samsonyPrice"] - prd["kidscomoPrice"]}
      else:
        if trackingInfo[prd["kidscomoPrdName"]]["samsonyPrice"] > prd["samsonyPrice"]:
          trackingInfo[prd["kidscomoPrdName"]]["samsonyPrice"] = prd["samsonyPrice"]
          trackingInfo[prd["kidscomoPrdName"]]["priceGap"] = trackingInfo[prd["kidscomoPrdName"]]["samsonyPrice"] - trackingInfo[prd["kidscomoPrdName"]]["kidscomoPrice"]
        else:
          pass

with open("./trackingJSON/result/trackingTotal.json", "w", encoding="UTF-8") as outfile:
  json.dump(trackingInfo, outfile, indent=2, ensure_ascii=False)