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

def is_holiday(target_date: date) -> bool:
    """日本の祝日判定 (2026-2027年対応)"""
    # 主要な祝日リスト（銀行休業日基準）
    holidays = [
        # 2026年
        "2026-01-01", "2026-01-02", "2026-01-03", "2026-01-12", "2026-02-11", "2026-02-23",
        "2026-03-20", "2026-04-29", "2026-05-03", "2026-05-04", "2026-05-05", "2026-05-06",
        "2026-07-20", "2026-08-11", "2026-09-21", "2026-09-22", "2026-09-23", "2026-10-12",
        "2026-11-03", "2026-11-23", "2026-12-31",
        # 2027年
        "2027-01-01", "2027-01-02", "2027-01-03", "2027-01-11", "2027-02-11", "2027-02-23",
        "2027-03-21", "2027-03-22", "2027-04-29", "2027-05-03", "2027-05-04", "2027-05-05",
        "2027-07-19", "2027-08-11", "2027-09-20", "2027-09-23", "2027-10-11", "2027-11-03",
        "2027-11-23", "2027-12-31"
    ]
    return target_date.strftime("%Y-%m-%d") in holidays

def next_business_day(target_date: date) -> date:
    """
    土日および祝日を避けて翌営業日を返します。
    """
    # 0:Mon, 1:Tue, ..., 5:Sat, 6:Sun
    # 土日、または祝日の間は日付を進める
    while target_date.weekday() >= 5 or is_holiday(target_date):
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
