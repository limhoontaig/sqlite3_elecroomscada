import sqlite3
import pandas as pd
import os
from tkinter import Tk, filedialog

# tkinter 창 숨기기
root = Tk()
root.withdraw()

# 1. 복수 개의 CSV 파일 선택 창 띄우기 (Ctrl이나 Shift로 다중 선택 가능)
csv_paths = filedialog.askopenfilenames(
    title="업로드할 CSV 파일들을 모두 선택하세요 (다중 선택 가능)",
    filetypes=[("CSV Files", "*.csv")]
)

if not csv_paths:
    print("CSV 파일 선택이 취소되었습니다.")
    exit()

print(f"선택된 CSV 파일 ({len(csv_paths)}개):")
for path in csv_paths:
    print(f" - {os.path.basename(path)}")

# 2. 데이터를 저장할 DB 파일 선택 (새로 만들려면 파일명을 직접 쳐도 됩니다)
db_path = filedialog.asksaveasfilename(
    title="데이터를 저장할 SQLite DB 파일(.db)을 지정하세요",
    filetypes=[("SQLite Database", "*.db")],
    defaultextension=".db"
)

if not db_path:
    print("DB 파일 지정이 취소되었습니다.")
    exit()

# 3. DB 연결
conn = sqlite3.connect(db_path)

# 4. 선택한 CSV 파일들을 순회하며 DB에 업로드
print(f"\n[{db_path}]에 데이터 업로드 시작...")
for csv_path in csv_paths:
    # 파일명에서 '_dump' 또는 '.csv'를 제거하여 깔끔한 테이블 이름 생성
    file_name = os.path.basename(csv_path)
    table_name = file_name.replace('_dump.csv', '').replace('.csv', '')
    
    print(f"[{file_name}] 읽어서 [{table_name}] 테이블로 업로드 중...")
    
    # CSV 파일 읽기
    df = pd.read_csv(csv_path)
    
    # DB에 테이블로 저장 (기존 테이블이 있다면 삭제 후 재생성)
    df.to_sql(table_name, conn, if_exists='replace', index=False)
    print(f" -> [{table_name}] 테이블 업로드 완료!")

conn.close()
print("\n✨ 선택한 모든 CSV 파일이 데이터베이스에 성공적으로 업로드되었습니다!")