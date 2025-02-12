import shutil
import sqlite3
con = sqlite3.connect('./productsData.db')
cur = con.cursor()
cur.execute("CREATE TABLE IF NOT EXISTS ProductsData(PrdName TEXT UNIQUE, Color TEXT UNIQUE, Size TEXT UNIQUE);")

# DB파일 백업
from_dbFile_path = './productsData.db'
to_dbFile_path = './dbBackup/productsData_backup.db'
shutil.copy(from_dbFile_path, to_dbFile_path)


query = "DELETE FROM ProductsData"
cur.execute(query)

import productseData
prdInfo = productseData.product_list
colorInfo = productseData.color_list
sizeInfo = productseData.size_list

for i in prdInfo:
  query = "INSERT OR IGNORE INTO ProductsData (PrdName) VALUES (?)"
  cur.execute(query, (i,))

for idx, color in enumerate(colorInfo):
  print(idx, color)
  query = "UPDATE OR IGNORE ProductsData SET Color=? WHERE rowid = ?"
  cur.execute(query, (color, idx + 1))

for idx, size in enumerate(sizeInfo):
  print(idx, size)
  query = "UPDATE OR IGNORE ProductsData SET Size=? WHERE rowid = ?"
  cur.execute(query, (size, idx + 1))

con.commit()

con.close()