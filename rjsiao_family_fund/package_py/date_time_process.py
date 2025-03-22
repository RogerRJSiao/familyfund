import datetime as dt
# from datetime import datetime, timedelta

def generate_month_periods(start: str, end: str):
    start_date = dt.datetime.strptime(start, "%Y%m")
    end_date = dt.datetime.strptime(end, "%Y%m")

    period_list = []

    curr_date = start_date
    while curr_date <= end_date:
        period_list.append(curr_date.strftime("%Y%m"))
        #嘗試計算下一個月
        curr_date = curr_date + dt.timedelta(days=32)
        curr_date = curr_date.replace(day=1)

    return tuple(period_list)