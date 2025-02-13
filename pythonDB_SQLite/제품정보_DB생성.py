import shutil
import sqlite3
import productsData

# SQLite 연결 객체 생성
con = sqlite3.connect('D:/1.업무/10.기타자료/Development/db/productsData.db')

# 커서 객체 생성
cur = con.cursor()

# 테이블 생성
cur.execute("CREATE TABLE IF NOT EXISTS ProductsData(PrdName TEXT UNIQUE, Color TEXT UNIQUE, Size TEXT UNIQUE, Cap TEXT UNIQUE);")

# 기존 DB파일 백업
from_dbFile_path = 'D:/1.업무/10.기타자료/Development/db/productsData.db'
to_dbFile_path = 'D:/1.업무/10.기타자료/Development/db/backup/productsData_backup.db'
shutil.copy(from_dbFile_path, to_dbFile_path)

# DB파일 기존 데이터 초기화
# query = "DELETE FROM ProductsData"
# cur.execute(query)

prdInfo = productsData.product_list
# colorInfo = productsData.color_list
# sizeInfo = productsData.size_list
# capInfo = productsData.capList

for i in prdInfo:
  # query = "INSERT OR IGNORE INTO ProductsData (PrdName) VALUES (?)"
  try:
    query = "INSERT INTO ProductsData (PrdName) VALUES (?)"
    cur.execute(query, (i,))
  except Exception as e:
    print('{} / {}'.format(i, e))

# for i in prdInfo:
#   query = "INSERT OR IGNORE INTO ProductsData (PrdName) VALUES (?)"
#   cur.execute(query, (i,))

# for idx, color in enumerate(colorInfo):
#   print(idx, color)
#   query = "UPDATE OR IGNORE ProductsData SET Color=? WHERE rowid = ?"
#   cur.execute(query, (color, idx + 1))

# for idx, size in enumerate(sizeInfo):
#   print(idx, size)
#   query = "UPDATE OR IGNORE ProductsData SET Size=? WHERE rowid = ?"
#   cur.execute(query, (size, idx + 1))

# for idx, cap in enumerate(capInfo):
#   print(idx, cap)
#   query = "UPDATE OR IGNORE ProductsData SET Cap=? WHERE rowid = ?"
#   cur.execute(query, (cap, idx + 1))

# DB 변경사항 저장
con.commit()

# cur.execute("SELECT Color from ProductsData WHERE Color IS NOT NULL ORDER BY rowid")
# data = cur.fetchall()
# for i in data:
#   print(i[0])

# DB연결 종료
con.close()