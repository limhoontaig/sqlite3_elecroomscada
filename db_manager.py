# db_manager.py
import os
import sqlite3
import struct
from datetime import datetime, timedelta

# 💡 사용자의 AppData/Local/ElecRoomSCADA 폴더 내에 생성
DB_DIR = os.path.join(os.environ['LOCALAPPDATA'], 'ElecRoomSCADA')
DB_NAME = os.path.join(DB_DIR, "plc_logging_real.db")

def ensure_db_directory():
    if not os.path.exists(DB_DIR):
        os.makedirs(DB_DIR)

ensure_db_directory()

# DB_NAME = "plc_logging_real.db"
DATA_LABELS = [
    "실내온도", "외기온도", "SF운전시간", "EF운전시간", 
    "KEP_V_R", "KEP_V_S", "KEP_V_T", "KEP_V_R_S", "KEP_V_S_T", "KEP_V_T_R", 
    "KEP_A_R", "KEP_A_S", "KEP_A_T", 
    "KEP_frequency", "KEP_P_kW", "KEP_P_kWh", 
    "Tr1_A_R", "Tr1_A_S", "Tr1_A_T", "Tr1_V_R", "Tr1_V_S", "Tr1_V_T", "Tr1_V_R_S", "Tr1_V_S_T", "Tr1_V_T_R", "Tr1_P_kW", "Tr1_Temp",
    "Tr2_A_R", "Tr2_A_S", "Tr2_A_T", "Tr2_V_R", "Tr2_V_S", "Tr2_V_T", "Tr2_V_R_S", "Tr2_V_S_T", "Tr2_V_T_R", "Tr2_P_kW", "Tr2_Temp",
    "Tr3_A_R", "Tr3_A_S", "Tr3_A_T", "Tr3_V_R", "Tr3_V_S", "Tr3_V_T", "Tr3_V_R_S", "Tr3_V_S_T", "Tr3_V_T_R", "Tr3_P_kW", "Tr3_Temp"
]

COLUMN_LABELS = ["날짜", "시간/구분"] + DATA_LABELS

def init_db():
    conn = sqlite3.connect(DB_NAME)
    c = conn.cursor()
    cols = ", ".join([f'"{name}" REAL' for name in DATA_LABELS])
    
    # 1. 실시간 데이터 테이블 (날짜와 시간을 묶어서 PRIMARY KEY 지정)
    c.execute(f'CREATE TABLE IF NOT EXISTS raw_data (log_date DATE, log_time TIME, {cols}, PRIMARY KEY (log_date, log_time))')
    
    # 🌟 [중요 수정] 평균 데이터 테이블에도 중복 저장을 막기 위해 PRIMARY KEY를 지정합니다.
    c.execute(f'CREATE TABLE IF NOT EXISTS hourly_avg (log_date DATE, log_time TIME, {cols}, PRIMARY KEY (log_date, log_time))')
    
    # 3. 최고/최저 데이터 테이블
    c.execute(f'CREATE TABLE IF NOT EXISTS daily_extremes (log_date DATE, extreme_type TEXT, {cols}, PRIMARY KEY (log_date, extreme_type))')
    
    # 🌟 [인덱스 개선] 날짜(log_date)와 시간(log_time)을 함께 인덱스로 잡으면 조회 속도가 갈라집니다.
    # 엑셀 보고서 출력이나 그래프 그릴 때 검색 속도가 수십 배 빨라집니다.
    c.execute('CREATE INDEX IF NOT EXISTS idx_raw_data_date_time ON raw_data (log_date, log_time);')
    c.execute('CREATE INDEX IF NOT EXISTS idx_hourly_avg_date_time ON hourly_avg (log_date, log_time);')

    create_field_inspection_table()
    conn.commit()
    conn.close()



# ==========================================
# 3. 데이터 분석 연산부
# ==========================================
def calculate_hourly_avg():
    try:
        conn = sqlite3.connect(DB_NAME, timeout=30)
        c = conn.cursor()
        now = datetime.now()
        last_hour = (now - timedelta(hours=1))
        target_date = last_hour.strftime('%Y-%m-%d')
        target_hour = last_hour.strftime('%H')
        
        avg_select = ", ".join([f'AVG("{name}")' for name in DATA_LABELS])
        col_names = ", ".join([f'"{name}"' for name in DATA_LABELS])
        
        query = f"SELECT {avg_select} FROM raw_data WHERE log_date = ? AND log_time LIKE ?"
        c.execute(query, (target_date, f"{target_hour}:%"))
        result = c.fetchone()

        if result and result[0] is not None:
            rounded_result = [round(val, 1) for val in result]
            placeholders = ", ".join(["?"] * len(DATA_LABELS))
            insert_query = f"INSERT OR REPLACE INTO hourly_avg (log_date, log_time, {col_names}) VALUES (?, ?, {placeholders})"
            c.execute(insert_query, [target_date, f"{target_hour}:00:00"] + rounded_result)
            conn.commit()
    except Exception as e:
        print(f"평균 계산 오류: {e}")
    finally:
        conn.close()


def calculate_daily_extremes(target_date):
    try:
        conn = sqlite3.connect(DB_NAME, timeout=30)
        c = conn.cursor()
        max_select = ", ".join([f'MAX("{name}")' for name in DATA_LABELS])
        min_select = ", ".join([f'MIN("{name}")' for name in DATA_LABELS])
        col_names = ", ".join([f'"{name}"' for name in DATA_LABELS])
        placeholders = ", ".join(["?"] * len(DATA_LABELS))
        
        c.execute(f'SELECT {max_select} FROM raw_data WHERE log_date = ?', (target_date,))
        max_res = c.fetchone()
        if max_res and max_res[0] is not None:
            c.execute(f'INSERT OR REPLACE INTO daily_extremes (log_date, extreme_type, {col_names}) VALUES (?, ?, {placeholders})', [target_date, 'MAX'] + list(max_res))
            
        c.execute(f'SELECT {min_select} FROM raw_data WHERE log_date = ?', (target_date,))
        min_res = c.fetchone()
        if min_res and min_res[0] is not None:
            c.execute(f'INSERT OR REPLACE INTO daily_extremes (log_date, extreme_type, {col_names}) VALUES (?, ?, {placeholders})', [target_date, 'MIN'] + list(min_res))
        conn.commit()
    except Exception as e:
        print(f"최고/최저 계산 오류: {e}")
    finally:
        conn.close()


# 계량기 검침내역 관리 db table 및 관련 함수들 

METER_FIELDS = [
    "main_active", "main_reactive", 
    "ind_mid", "ind_max", "ind_light", 
    "street_mid", "street_max", "street_light", 
    "geo_1", "geo_2", "geo_3"
]

def create_manual_meter_table():
    """3개 계량장치의 일일 지침을 저장할 독자적인 테이블을 생성합니다."""
    conn = sqlite3.connect(DB_NAME)
    c = conn.cursor()
    c.execute('''
        CREATE TABLE IF NOT EXISTS manual_meter_logs (
            log_date DATE PRIMARY KEY,
            main_active REAL,
            main_reactive REAL,
            ind_mid REAL,
            ind_max REAL,
            ind_light REAL,
            street_mid REAL,
            street_max REAL,
            street_light REAL,
            geo_1 REAL,
            geo_2 REAL,
            geo_3 REAL
        )
    ''')
    conn.commit()
    conn.close()

def get_manual_meter_data(target_date):
    """지정된 날짜의 수동 검침 지침 데이터를 딕셔너리 형태로 불러옵니다."""
    create_manual_meter_table() # 테이블 없으면 자동 생성
    conn = sqlite3.connect(DB_NAME)
    c = conn.cursor()
    
    fields_str = ", ".join(METER_FIELDS)
    c.execute(f'SELECT {fields_str} FROM manual_meter_logs WHERE log_date = ?', (target_date,))
    res = c.fetchone()
    conn.close()
    
    data_dict = {f: "" for f in METER_FIELDS}
    if res:
        for idx, field in enumerate(METER_FIELDS):
            data_dict[field] = str(res[idx]) if res[idx] is not None else ""
    return data_dict

def save_manual_meter_data(target_date, data_dict):
    """지정된 날짜에 11개 필드 데이터를 저장하거나 기존 데이터가 있으면 수정(UPDATE)합니다."""
    create_manual_meter_table()
    conn = sqlite3.connect(DB_NAME)
    c = conn.cursor()
    
    # INSERT OR REPLACE 구문을 활용하여 데이터 유무에 따라 저장/수정 자동 처리
    fields_str = "log_date, " + ", ".join(METER_FIELDS)
    placeholders = ", ".join(["?"] * (len(METER_FIELDS) + 1))
    
    values = [target_date]
    for field in METER_FIELDS:
        val = data_dict.get(field, "")
        values.append(float(val) if val.strip() != "" else None)
        
    c.execute(f'INSERT OR REPLACE INTO manual_meter_logs ({fields_str}) VALUES ({placeholders})', values)
    conn.commit()
    conn.close()

def get_manual_meter_log_for_table(target_date):
    """메인 화면 테이블 표기용으로 해당 날짜의 수동 검침 데이터를 리스트 형태로 반환합니다."""
    create_manual_meter_table()
    conn = sqlite3.connect(DB_NAME)
    c = conn.cursor()
    
    fields_str = ", ".join(METER_FIELDS)
    c.execute(f"SELECT log_date, {fields_str} FROM manual_meter_logs WHERE log_date = ?", (target_date,))
    row = c.fetchone()
    conn.close()
    
    if row:
        # 데이터를 UI 테이블에 출력하기 좋은 정밀도나 스트링 형태로 변환하여 반환
        return [row[0]] + [f"{v:.1f}" if isinstance(v, float) else str(v) if v is not None else "-" for v in row[1:]]
    else:
        # 데이터가 없을 경우 날짜와 함께 빈 대시(-) 채우기
        return [target_date] + ["-"] * len(METER_FIELDS)

# ==========================================
# [신규 추가] 현장 점검 관리 DB 및 함수들
# ==========================================

def create_field_inspection_table():
    """현장 점검 데이터를 기록할 테이블을 생성합니다."""
    conn = sqlite3.connect(DB_NAME)
    c = conn.cursor()
    # 날짜와 차수(1, 2, 3차)의 조합이 중복되지 않도록 복합 기본키(PRIMARY KEY) 설정
    c.execute('''
        CREATE TABLE IF NOT EXISTS field_inspection (
            inspection_date DATE,
            inspection_round INTEGER,
            inspector_name TEXT NOT NULL,
            inspected_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            PRIMARY KEY (inspection_date, inspection_round)
        )
    ''')
    conn.commit()
    conn.close()

def save_field_inspection(target_date, round_val, inspector_name):
    """지정된 날짜와 차수에 점검자 정보를 저장합니다. 이미 있으면 수정(UPDATE)합니다."""
    create_field_inspection_table()
    conn = sqlite3.connect(DB_NAME)
    c = conn.cursor()
    try:
        c.execute('''
            INSERT OR REPLACE INTO field_inspection (inspection_date, inspection_round, inspector_name, inspected_at)
            VALUES (?, ?, ?, datetime('now', 'localtime'))
        ''', (target_date, round_val, inspector_name))
        conn.commit()
        return True
    except Exception as e:
        print(f"현장 점검 저장 오류: {e}")
        return False
    finally:
        conn.close()

def get_field_inspections_for_date(target_date):
    """특정 날짜의 1, 2, 3차 점검 내역을 딕셔너리 형태로 반환합니다."""
    create_field_inspection_table()
    conn = sqlite3.connect(DB_NAME)
    c = conn.cursor()
    c.execute('''
        SELECT inspection_round, inspector_name, strftime('%H:%M', inspected_at) 
        FROM field_inspection 
        WHERE inspection_date = ?
    ''', (target_date,))
    rows = c.fetchall()
    conn.close()
    
    # 기본값 설정 (데이터가 없을 경우 빈 문자열 처리)
    result = {
        1: {"name": "", "time": ""},
        2: {"name": "", "time": ""},
        3: {"name": "", "time": ""}
    }
    for row in rows:
        round_num, name, time_str = row
        if round_num in result:
            result[round_num] = {"name": name, "time": time_str}
    return result        