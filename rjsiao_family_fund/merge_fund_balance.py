import sys
import gspread
from oauth2client.service_account import ServiceAccountCredentials

import datetime
import pandas as pd


# MAIN
if __name__ == "__main__":
    print("Start: merge_fund_balance.py")

    # 取出傳入參數
    if len(sys.argv) > 1:
        for index, arg in enumerate(sys.argv[1:], start=1):
            print(f"Argument {index}: {arg}")
    else:
        print("No arguments passed")
    yyyymm_start = sys.argv[1]  #開始年月
    yyyymm_end = sys.argv[2]    #結束年月

    # 取得當天日期，修正 yyyymm_end
    today = datetime.date.today()
    yyyymmdd = today.strftime("%Y%m%d")
    if yyyymmdd < yyyymm_end:
        yyyymm_end = str(int(yyyymmdd[0:6])-1) if int(yyyymmdd[4:6]) > 1 else str(yyyymmdd[0:4]-1)+"12"
    if yyyymmdd < yyyymm_start:
        yyyymm_start = str(int(yyyymmdd[0:6])-1) if int(yyyymmdd[4:6]) > 1 else str(yyyymmdd[0:4]-1)+"12"
    print(f'Today: {yyyymmdd}\nyyyymm_start: {yyyymm_start}\nyyyymm_end: {yyyymm_end}')

    # 取得 Google Sheet 資料

    # Google Sheet 認證


