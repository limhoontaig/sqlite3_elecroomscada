import sqlite3
import pandas as pd
import os
from tkinter import Tk, filedialog

# tkinter 창 숨기기
root = Tk()
root.withdraw()

# 1. 파일 선택 창 띄우기 (DB 파일 선택)
db_path = filedialog.askopenfilename(
    title="백업할 SQLite DB 파일(.db)을 선택하세요",
    filetypes=[("SQLite Database", "*.db"), ("All Files", "*.*")]
)

# 파일 선택을 취소한 경우 종료
if not db_path:
    print("파일 선택이 취소되었습니다.")
    exit()

# DB 파일이 있는 디렉토리 경로 추출
base_dir = os.path.dirname(db_path)

# 2. DB 연결
conn = sqlite3.connect(db_path)
cursor = conn.cursor()

# 3. 모든 테이블 이름 가져오기
cursor.execute("SELECT name FROM sqlite_master WHERE type='table';")
tables = [row[0] for row in cursor.fetchall()]

print(f"선택한 DB: {db_path}")
print(f"발견된 테이블: {tables}\n")

# 4. 각 테이블을 CSV로 저장 (DB와 같은 폴더에 저장)
for table_name in tables:
    print(f"[{table_name}] 테이블 내보내는 중...")
    
    df = pd.read_sql_query(f"SELECT * FROM {table_name}", conn)
    
    # DB와 동일한 디렉토리에 파일 경로 지정
    csv_file_name = f"{table_name}_dump.csv"
    csv_file_path = os.path.join(base_dir, csv_file_name)
    
    # 한글 깨짐 방지용 utf-8-sig
    df.to_csv(csv_file_path, index=False, encoding='utf-8-sig')
    print(f" -> 저장 완료: {csv_file_path}")

conn.close()
print("\n✨ 모든 테이블이 해당 폴더에 성공적으로 CSV로 백업되었습니다!")