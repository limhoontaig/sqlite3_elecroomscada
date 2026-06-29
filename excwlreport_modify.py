# =================================================================
# excel_report.py 내의 get_stat 함수 수정본
# =================================================================

    c = conn.cursor()
    
    def get_stat(func, col):
        try:
            c.execute(f'SELECT {func}("{col}") FROM raw_data WHERE log_date = ?', (selected_date,))
            res = c.fetchone()[0]
            
            # DB에서 가져온 값이 숫자가 맞는지 한 번 더 꼼꼼하게 체크합니다.
            if res is not None and isinstance(res, (int, float)):
                return round(float(res), 1)
            else:
                return "-"  # 데이터가 비어있거나 None이면 안전하게 문자열 대시 반환
        except Exception as e:
            print(f"통계 연산 중 오류 발생 ({col}): {e}")
            return "-"  # 에러가 나도 프로그램이 죽지 않도록 방어