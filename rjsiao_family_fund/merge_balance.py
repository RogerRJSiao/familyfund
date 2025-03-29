import sys

from package_py import google_files_process as gfp
from package_py import error_handle as eh
from package_py import date_time_process as dtp

import datetime as dt
import pandas as pd

def set_src_permission(year, month, src):
  # 注意檔案路徑，相對於當前執行的 py 檔案位置
  file_google_sheetid = 'src_data_fund.xlsx'
  credentials = 'api/rogersiaoaiot-myfamilyfund-250329.json'
  
  # 讀取對照檔
  sheet_name = 'Y' + str(year)
  df_src = pd.read_excel(file_google_sheetid, sheet_name=sheet_name)
  # 讀取替身
  auth = df_src[df_src['唯一值'] == 'auth']['Google_sheet_id'].values[0]
  # 讀取 Google Sheet ID (不可以是 .xlsx id)
  match src:
    case 'form':  #對照組=單頭/單身
      sheet_id = df_src[df_src['唯一值'] == 'form']['Google_sheet_id'].values[0]
    case 'A'|'B'|'C'|'D'|'E'|'Z':  #即時檔
      key = str(year) + src
      sheet_id = df_src[df_src['唯一值'] == key]['Google_sheet_id'].values[0]
    case 'monthly':  #月結檔
      key = str(year) + str(month)  
      sheet_id = df_src[df_src['唯一值'] == key]['Google_sheet_id'].values[0]
    case _:
      sheet_id = ''
      auth = ''
      credentials = ''

  # print(f"auth: {auth}, \nsheet_id: {sheet_id}, \ncredentials: {credentials}")
  return {'sheet_id': sheet_id, 'auth': auth, 'credential': credentials}

def processing_monthly_amount(df_mydata, tbl_fmt, method):
  df_calculated = pd.DataFrame()
  match method:
    case 'agg':       # 聚合函數
      df_calculated = df_mydata.groupby(tbl_fmt[0:2], as_index=False).agg({tbl_fmt[2]: 'sum'})
      df_calculated = df_calculated.sort_values(tbl_fmt[1],ascending=True)
    case 'pvt'|'_':   # 樞紐分析
      df_mydata = df_mydata[tbl_fmt]
      df_calculated = df_mydata.pivot_table(values=tbl_fmt[2], index=tbl_fmt[0], columns=tbl_fmt[1], 
                                            aggfunc='sum', fill_value=0, margins=True, margins_name='小計') #缺值補0，加小計
  
  # print(df_calculated)
  return df_calculated

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
    arr_yyyymm = tuple()        #月結年月列表

    # 取得當天日期，修正 yyyymm_end
    today = dt.date.today()
    yyyymmdd = today.strftime("%Y%m%d")
    now = dt.datetime.now()
    formatted_dt = now.strftime("%Y%m%d_%H%M%S")

    if yyyymmdd < yyyymm_end:
        yyyymm_end = str(int(yyyymmdd[0:6])-1) if int(yyyymmdd[4:6]) > 1 else str(yyyymmdd[0:4]-1)+"12"
    if yyyymmdd < yyyymm_start:
        yyyymm_start = str(int(yyyymmdd[0:6])-1) if int(yyyymmdd[4:6]) > 1 else str(yyyymmdd[0:4]-1)+"12"
        # 開始實施--202501月結
        yyyymm_start = yyyymm_start if int(yyyymm_start) >= 202501 else "202501"
    print(f'Today: {yyyymmdd}\nyyyymm_start: {yyyymm_start}\nyyyymm_end: {yyyymm_end}')

    # 定義月別列表
    arr_yyyymm = dtp.generate_month_periods(yyyymm_start, yyyymm_end)
    print(f'YearMonths yyyymm: {arr_yyyymm}.')

    # 定義檢查帳別範圍
    arr_check_fund = ['A','B','C','D','E','Z']
    print(f"Family funds: {arr_check_fund}.")
    
    # 提取資料集：月結檔
    df_mydata = pd.DataFrame()
    for yyyymm in arr_yyyymm:
      year = yyyymm[0:4]
      month = yyyymm[4:6]
      conn_sheets = "月結檔"
      access = set_src_permission(year, month, 'monthly')
      sheets = gfp.get_google_sheets(access['sheet_id'], access['auth'], access['credential'], conn_sheets)

      for fund in arr_check_fund:
        df_monthly_fund = pd.DataFrame()
        print(f'Now is fetching data from: {month}{fund}.')
        if not sheets:
          msg = f"檔案-{year}{month}{conn_sheets} 活頁-{fund} 連線異常，無資料!!!"
          eh.create_today_log(msg)
          continue

        sheet_name = fund # 活頁名稱：帳別
        arr = gfp.get_sheet_data(sheets, conn_sheets, sheet_name)
        if not arr:
          msg = f"檔案-{year}{month}{conn_sheets} 活頁-{fund} 無資料!!!"
          eh.create_today_log(msg)
          continue
        
        df_monthly_fund = pd.DataFrame(arr)
        df_mydata = pd.concat([df_mydata, df_monthly_fund], axis=0, ignore_index=True)
        print(f"資料累計(列,欄): {df_mydata.shape}")
    
    # 資料預處理
    df_mydata_ori = df_mydata.copy()
    df_mydata = df_mydata.sort_values(by=['申報個帳','月結年月','核對處理日','認列碼'], 
                                      ascending=[True,True,True,True], ignore_index=True)
    print(f"資料總計(列): {len(df_mydata)}")
    
    # 計算月結年月x收支
    tbl_fmt = ['收支','月結年月','認列金額'] #[列,欄,值]
    df_mydata_rpt1 = processing_monthly_amount(df_mydata_ori, tbl_fmt, 'pvt')
    # 計算月結年月x中分類
    tbl_fmt = ['中分類','月結年月','認列金額'] #[列,欄,值]
    df_mydata_rpt2 = processing_monthly_amount(df_mydata_ori, tbl_fmt, 'pvt')

    # 準備路徑檔名
    min_month = arr_yyyymm[0]
    max_month = arr_yyyymm[-1]
    file_name_output = f"家庭記帳彙總_{min_month}_{max_month}_更新{formatted_dt}.xlsx"
    file_path = "../mydata/" + file_name_output
    # 將資料集寫檔至指定活頁
    with pd.ExcelWriter(file_path) as writer:
      df_mydata.to_excel(writer, sheet_name='總表', index=False)
      df_mydata_rpt1.to_excel(writer, sheet_name='每月收支別', index=True)
      df_mydata_rpt2.to_excel(writer, sheet_name='每月中分類別', index=True)

    print(f"檔案-{file_name_output} 已下載完成!")
