import calendar
from datetime import date, timedelta
from dateutil.relativedelta import relativedelta

def get_effective_date(year: int, month: int, day: int) -> date:
    """
    指定された年、月の「day日」を返します。
    day=31の場合、その月の最終日（30日や28日など）に自動調整します。
    """
    last_day = calendar.monthrange(year, month)[1]
    actual_day = min(day, last_day)
    return date(year, month, actual_day)

def calculate_payment_date(base_date: date, term_months: int, payment_day: int) -> date:
    """
    注文日(base_date)から見て、指定された支払月(term_months後)と支払日(payment_day)を計算します。
    payment_day=31の場合は、その月の最終日に調整します。
    """
    # 指定された月数分後の日付を取得
    target_month_date = base_date + relativedelta(months=term_months)
    
    # その月の指定日を取得（月末調整込み）
    return get_effective_date(target_month_date.year, target_month_date.month, payment_day)

def is_closing_day(check_date: date, closing_day: int) -> bool:
    """
    check_date が指定された締め日(closing_day)に該当するか判定します。
    closing_day=31の場合、その月の最終日であれば True を返します。
    """
    last_day = calendar.monthrange(check_date.year, check_date.month)[1]
    
    # 締め日が31の場合、今日が月末なら True
    if closing_day >= 31:
        return check_date.day == last_day
    
    # それ以外は日付が一致するか（ただし、30日しかない月に 31日締め設定の人が漏れないように min を使う）
    effective_closing = min(closing_day, last_day)
    return check_date.day == effective_closing

def next_business_day(target_date: date) -> date:
    """
    土日を避けて翌営業日を返します。（祝日は未対応ですが、簡易的に土日のみ対応）
    """
    # 0:Mon, 1:Tue, ..., 5:Sat, 6:Sun
    while target_date.weekday() >= 5:
        target_date += timedelta(days=1)
    return target_date

def get_next_closing_date(base_date: date, closing_day: int) -> date:
    """
    base_date 以降で最も近い締め日を返します。
    closing_day=None または 0 の場合は、base_date をそのまま締め日として扱います（都度締め）。
    """
    if not closing_day:
        return base_date
        
    # その月の暫定締め日（31日の場合は月末調整含む）
    potential_closing = get_effective_date(base_date.year, base_date.month, closing_day)
    
    # すでにその月の締め日を過ぎている場合、翌月の締め日を返す
    if base_date > potential_closing:
        next_month = base_date + relativedelta(months=1)
        return get_effective_date(next_month.year, next_month.month, closing_day)
    
    return potential_closing
