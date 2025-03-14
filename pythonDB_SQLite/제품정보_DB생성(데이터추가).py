import shutil
import sqlite3
import productsData

# SQLite 연결 객체 생성
con = sqlite3.connect('D:/1.업무/10.기타자료/Development/db/productsData.db')

# 커서 객체 생성
cur = con.cursor()

# 기존 DB파일 백업
from_dbFile_path = 'D:/1.업무/10.기타자료/Development/db/productsData.db'
to_dbFile_path = 'D:/1.업무/10.기타자료/Development/db/backup/productsData_backup.db'
shutil.copy(from_dbFile_path, to_dbFile_path)

excPrdInfo = productsData.excPrd_List

# 데이터 추가 기준 초기행 지정
initRow = 4
for i in excPrdInfo:
  try:
    query = "UPDATE ProductsData SET ExcProducts=? WHERE rowid = ?"
    cur.execute(query, (i, initRow))
    initRow += 1
  except Exception as e:
    print('{} / {}'.format(i, e))

# DB 변경사항 저장
con.commit()

# DB연결 종료
con.close()