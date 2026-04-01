import sys
import os
from datetime import date

# 自作モジュールを読み込めるようにパスを通す
sys.path.append(os.getcwd())

from utils.date_utils import get_next_closing_date, calculate_payment_date, next_business_day

def test_logic(label, order_date, closing_day, term_months, payment_day):
    print(f"--- {label} ---")
    closing_date = get_next_closing_date(order_date, closing_day)
    due_date = calculate_payment_date(closing_date, term_months, payment_day)
    due_date = next_business_day(due_date)
    
    print(f"購入日: {order_date}")
    print(f"計算された締め日: {closing_date}")
    print(f"計算されたお支払い期限: {due_date}")
    return due_date

# ケース1: 中村様の例 (4/1購入 -> 4/30締め -> 5/31払い)
test_logic("ケース1: 中村様の例", date(2026, 4, 1), 31, 1, 31)

# ケース2: 締め日当日 (4/30購入 -> 4/30締め -> 5/31払い)
test_logic("ケース2: 締め日当日", date(2026, 4, 30), 31, 1, 31)

# ケース3: 締め日翌日 (5/1購入 -> 5/31締め -> 6/30払い)
test_logic("ケース3: 締め日翌日", date(2026, 5, 1), 31, 1, 31)
