from package_py import google_files_process as gfp
from package_py import error_handle as eh

import csv
import datetime
import pandas as pd

def process_sheet_data(date):
  # 定義檢查月別範圍(每月5日之前會檢查前一月，1月例外)
  curr_yr = int(date[0:4])
  curr_mon = int(date[5:7])
  last_mon = curr_mon - 1 if curr_mon > 1 else 12
  curr_mon = date[5:7]
  last_mon = last_mon if len(str(last_mon)) == 2 else ('0' + str(last_mon))
  curr_day = int(date[8:10])
  arr_check_month = [curr_mon] if curr_day >= 6 or last_mon == 12 else [last_mon, curr_mon]
  print(f"Today is {date}, and this report will show data within {arr_check_month} month.")

  # 定義檢查帳別範圍
  arr_check_fund = ['A','B','C','D','E','Z']
  print(f"Family funds include {arr_check_fund}.")

  # 提取資料集：表單form-單頭、單身
  conn_sheets = "單頭單身"
  access = set_src_permission(curr_yr, '', 'form')
  sheets = gfp.get_google_sheets(access['sheet_id'], access['auth'], access['credential'], conn_sheets)
  if not sheets:
    msg = f"檔案-{conn_sheets} 連線異常，無資料!!!"
    eh.create_today_log(msg)
    return

  sheet_name = "單頭"  
  mydata_form_head = gfp.get_sheet_data(sheets, conn_sheets, sheet_name)
  sheet_name = "單身"
  mydata_form_body = gfp.get_sheet_data(sheets, conn_sheets, sheet_name)
  if not mydata_form_head or not mydata_form_body:
    msg = f"檔案-{conn_sheets} 無資料!!!"
    eh.create_today_log(msg)
    return

  # 標記form資料集各列來源
  mydata_form_head = tag_each_record(mydata_form_head, curr_yr, '', '單頭')
  mydata_form_body = tag_each_record(mydata_form_body, curr_yr, '', '單身')
  mydata_form = {'h': mydata_form_head, 'd': mydata_form_body}

  # 提取資料集：請款fund
  mydata_fund = []
  reviewed = {'finished':[], 'error':[]}
  for fund in arr_check_fund:     # 帳別
    for month in arr_check_month: # 月別
      print(f'Now is fetching data from: {month}{fund}.')
      conn_sheets = fund + "帳" # 檔案名稱：帳別
      access = set_src_permission(curr_yr, month, fund)
      sheets = gfp.get_google_sheets(access['sheet_id'], access['auth'], access['credential'], conn_sheets)
      if not sheets:
        msg = f"檔案-{conn_sheets} 活頁-{month} 連線異常，無資料!!!"
        eh.create_today_log(msg)
        reviewed['error'].append(conn_sheets)
        continue
      
      sheet_name = month # 活頁名稱：月別 
      arr = gfp.get_sheet_data(sheets, conn_sheets, sheet_name)
      if not arr:
        msg = f"檔案-{conn_sheets} 活頁-{month} 無資料!!!"
        eh.create_today_log(msg)
        reviewed['error'].append(month + fund)
        continue

      # 標記fund資料集各列來源
      arr = tag_each_record(arr, curr_yr, month, fund)
      # 新讀取帳別(檔案)、月別(活頁)加入資料集
      mydata_fund = mydata_fund + arr
      reviewed['finished'].append(month + fund)
      print(f'Fetched data records from {month}{fund}: ' + str(len(mydata_fund)))

  # 比對單頭、單身、會計(單頭、單身對照組是 form)
  data_mismatched = get_proofread_records(mydata_form, mydata_fund, arr_check_month, arr_check_fund)
  mydata = {
    'ori_h': mydata_form_head, 'ori_b': mydata_form_body,
    'mis_h': data_mismatched['h'], 'mis_b': data_mismatched['b'], 
    'mis_a': data_mismatched['a'], 'mis_s': data_mismatched['s'], 
  }

  return mydata, reviewed


def set_src_permission(year, month, src):
  # 注意檔案路徑，相對於當前執行的 py 檔案位置
  file_google_sheetid = 'src_data_fund.xlsx'
  credentials = 'api/credentials_family_fund_form.json'
  
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

  # print(f"auth: {auth}, sheet_id: {sheet_id}, credentials: {credentials}")
  return {'sheet_id': sheet_id, 'auth': auth, 'credential': credentials}


def tag_each_record(data, year, month, fund):
  untaged_records = data
  taged_records = []
  curr_cnt = 0
  for record in untaged_records:
    curr_cnt = curr_cnt + 1
    record['帳別'] = fund if fund else ''
    record['年'] = year if year else ''
    record['月'] = month if month else ''
    record['列次'] = curr_cnt + 1 # form 和 fund 都有標題列(第1列)
    taged_records.append(record)

  return taged_records


def get_proofread_records(data_form, data_fund, arr_months, arr_funds):
  ### fund 資料表表頭
  # 單頭 head：Timestamp*	店名	申報號碼**	日期	支付方式	申報者	申報順序	
  # 單身 body：Timestamp*	申報號碼**	品項或說明	產品簡稱	數量	計價單位	折價	金額	申報個帳	
  # 會計 account：月結年月	核對處理日	收支	中分類	認列金額	認列碼	戶名	認列說明

  # form 資料清洗：去除狀態碼為刪除者
  df = pd.DataFrame(data_form['h'])
  df_formhead = df[(df['狀態碼'] != '刪除')]
  df = pd.DataFrame(data_form['d'])
  df_formbody = df[(df['狀態碼'] != '刪除')]

  # fund 資料清洗：去除收入帳者
  df = pd.DataFrame(data_fund)
  df_fund = df[(df['收支'] == '支') | (df['收支'] == '支退')]

  # 從 fund 取值 f 並逐列檢查 (form 對照組)
  data_mismatched_head = []
  data_mismatched_body = []
  data_mismatched_account = []
  for row in df_fund.itertuples(): 
    # 單頭資料
    h_stamp = row.Timestamp_單頭
    h_shop = row.店名	
    h_invoice = row.申報號碼_單頭
    h_date = row.日期	
    h_pay = row.支付方式	
    h_obj = row.申報者	
    h_seq = row.申報順序

    # 單身資料
    b_stamp = row.Timestamp_單身	
    b_invoice = row.申報號碼_單身
    b_item = row.品項或說明
    b_abbr = row.產品簡稱
    b_qty = row.數量
    b_unit = row.計價單位
    b_dicount = row.折價
    b_amount = row.金額
    b_fund = row.申報個帳

    # 會計資料
    a_sort = row.中分類	
    a_date = row.核對處理日	
    a_code = row.認列碼	
    a_obj = row.戶名	
    a_inout = row.收支	
    a_amount = row.認列金額	
    a_remark = row.認列說明	
    a_yyyymm = row.月結年月
    x_fund = row.帳別
    x_year = row.年
    x_month = row.月
    x_rownum = row.列次

    if h_stamp and b_stamp: # 個帳手動輸入(期初、收)不列入判斷
      try: 
        # 比對是否存在單頭資料
        h_query = ' Timestamp==@h_stamp and 店名==@h_shop and 申報號碼==@h_invoice and 日期==@h_date and 支付方式==@h_pay ' 
        h_query = h_query + ' and 申報者==@h_obj and 申報順序==@h_seq '
        h_query_df = df_formhead.query(h_query)
        if len(h_query_df) == 0:
          arr = modify_record_element(row, 'head', f'單頭異常!! 這一筆在 {x_fund} 帳的單頭 其中一欄與表單(單頭)輸入不一致')
          data_mismatched_head.append(arr)

        # 比對是否存在單身資料
        h_query = ' Timestamp==@b_stamp and 申報號碼==@b_invoice and 品項或說明==@b_item and 產品簡稱==@b_abbr and 數量==@b_qty ' 
        h_query = h_query + ' and 計價單位==@b_unit and 折價==@b_dicount and 金額==@b_amount and 申報個帳==@b_fund '
        h_query_df = df_formbody.query(h_query)
        if len(h_query_df) == 0:
          arr = modify_record_element(row, 'body', f'單身異常!! 這一筆在 {x_fund} 帳的單身 其中一欄與表單(單身)輸入不一致')
          data_mismatched_body.append(arr)
          # 比對單身字數：品項或說明
          h_query = ' Timestamp==@b_stamp and 申報號碼==@b_invoice '
          h_query_df = df_formbody.query(h_query)
          if len(h_query_df) >= 1:
            len_item_fund = len(b_item)
            len_item_form = len(h_query_df.iloc[0, 2])  # 品項或說明
            if len_item_form != len_item_fund:
              arr = modify_record_element(row, 'body', f'單身異常!! 品項或說明 在個帳({len_item_fund}字)、表單({len_item_form}字)不一致 (檢查左右有無留空)')
              data_mismatched_body.append(arr)
      except:
        arr = modify_record_element(row, 'common', '異常!! 單頭單身的欄名可能有異動')
        data_mismatched_account.append(arr)


    # 驗證會計資料
    try:
      if a_sort[:1] != x_fund:
        arr = modify_record_element(row, 'account', '異常!! 中分類 第一碼是帳別')
        data_mismatched_account.append(arr)
      if (a_date[5:7] != x_month and int(a_date[5:7]) != int(x_month) + 1) or not is_valid_iso_date(a_date):
        arr = modify_record_element(row, 'account', '異常!! 核對處理日 格式yyyy-mm-dd，且mm是當月或次月')
        data_mismatched_account.append(arr)
      if a_code[:2] != x_month or a_code[2:3] != x_fund: 
        arr = modify_record_element(row, 'account', '異常!! 認列碼 是月別mm+帳別+流水號，且mm是當月')
        data_mismatched_account.append(arr)
      if a_obj not in ['A','B','C','D','E','Z','蕭寶郎','簡慧卿','蕭瑞展','蕭宛柔']:
        arr = modify_record_element(row, 'account', '異常!! 戶名 是帳別名或註冊人名，且與店名或申報者相同')
        data_mismatched_account.append(arr)
      if a_inout not in ['期初','收','支','支退']:
        arr = modify_record_element(row, 'account', '異常!! 收支 是期初、收、支、支退')
        data_mismatched_account.append(arr)
      if a_amount > 0 or type(a_amount) == 'float':
        arr = modify_record_element(row, 'account', '異常!! 認列金額 對應支出應為負數，且不為小數')
        data_mismatched_account.append(arr)
      if abs(a_amount) > b_amount:
        arr = modify_record_element(row, 'account', '異常!! 認列金額 不大於實際報帳金額')
        data_mismatched_account.append(arr)
      if int(str(a_yyyymm)[:4]) != int(x_year) or int(str(a_yyyymm)[4:6]) != int(x_month):
        arr = modify_record_element(row, 'account', '異常!! 月結年月 應為當年度的當月份')
        data_mismatched_account.append(arr)
    except:
      msg = f"{x_month}{x_fund} 最後一列下方儲存格有多餘資料未刪，或右列會計輸入有誤、有缺漏值"
      eh.create_today_log(msg)

  # 檢查跳號
  data_mismatched_serial = [] # 存取跳號流水號
  for fund in arr_funds:
    for month in arr_months:
      df_serial = df_fund['認列碼'][df['認列碼'].str[:3] == (month + fund)] # 只取認列碼前3碼
      print(month + fund)
      print(len(df_serial))
      if len(df_serial) != 0:
        arr_serials = df_serial.tolist()
        arr_serials = list(dict.fromkeys(arr_serials)) # 移除重複值
        if len(arr_serials) != int(max(arr_serials)[-2:]):
          arr_serials.sort() # 由小到大排序
          for i in range(0, len(arr_serials)):
            print(f"{i + 1} -- {int(arr_serials[i][-2:])}") 
            if i + 1 != int(arr_serials[i][-2:]):
              arr = modify_record_element(row, 'serial', f'異常!! 認列碼 出現跳號：{arr_serials[i]}')
              data_mismatched_serial.append(arr)
    
    data_organized = {
      'h': data_mismatched_head, 'b': data_mismatched_body, 
      'a': data_mismatched_account, 's': data_mismatched_serial
    }
    
  return data_organized


def modify_record_element(row, tbl_name, msg):
  record_refined = []
  show_msg = msg
  # convert tuple (row) to list
  match tbl_name:
    case 'common':
      record_refined = [show_msg]
    case 'account':
      record_refined = list(row[-4:]) + [show_msg] + list(row[1:-4])
    case 'head':
      record_refined = list(row[-4:]) + [show_msg] + list(row[1:8])
    case 'body':
      record_refined = list(row[-4:]) + [show_msg] + list(row[8:17])
    case 'serial':
      record_refined = [show_msg]
    case _:
      record_refined = []

  return record_refined


def is_valid_iso_date(date_string):
    try:
        datetime.date.fromisoformat(date_string)
        return True
    except ValueError:
        return False


# MAIN
if __name__ == "__main__":
  # 連線sheet並取得資料
  today = datetime.date.today()
  iso_date = today.isoformat()
  print(iso_date) #yyyy-mm-dd
  data, reviewed = process_sheet_data(iso_date)
  delimiter_comma = "," # set a delimiter — using comma for illustration
  reviewed_finished = delimiter_comma.join(reviewed['finished']) if len(reviewed['finished']) > 0 else '無'
  reviewed_error = delimiter_comma.join(reviewed['error']) if len(reviewed['error']) > 0 else '無'
  print("完成比對"+ delimiter_comma.join(reviewed_finished))
  print("讀取資料異常"+ delimiter_comma.join(reviewed_error))

  if not data:
    msg = f"無法取得資料!!!"
    eh.create_today_log(msg)
  else:
    # 格式化日期時間為字串
    now = datetime.datetime.now()
    formatted_date = now.strftime("%Y%m%d_%H%M%S")

    # 設定儲存路徑、檔名(用時戳)
    file_name = '家庭記帳比對_' + formatted_date + '.csv'
    file_path = "../mydata/" + file_name

    title_check = ['帳別','年','月','個帳列次','check']
    title_head = ['Timestamp_單頭','店名','申報號碼_單頭','日期','支付方式','申報者','申報順序']
    title_body = ['Timestamp_單身','申報號碼_單身','品項或說明','產品簡稱','數量','計價單位','折價','金額','申報個帳']	
    title_account = ['月結年月','核對處理日','收支','中分類','認列金額','認列碼','戶名','認列說明']
    title_fund = title_head + title_body + title_account

    # Write data to CSV file
    with open(file_path, mode='w', newline='') as file:
      writer = csv.writer(file)
      # 報表頁首
      writer.writerows([['=== 家庭共帳異常報表 ===']])
      writer.writerows([['製表時刻：'+ str(now)[:19] ]])
      writer.writerows([['完成比對：'+ reviewed_finished]])
      writer.writerows([['讀取資料異常：'+ reviewed_error]])
      writer.writerows([[]])

      # 寫入單頭異常
      writer.writerows([title_check + title_head])
      cnt = len(data['mis_h'])
      if cnt != 0:
        writer.writerows(data['mis_h'])
        writer.writerows([['==> 筆數：'+ str(cnt)]])
      else: 
        writer.writerows([['單頭無異常!']])
      writer.writerows([[]])

      # 寫入單身異常
      writer.writerows([title_check + title_body])
      cnt = len(data['mis_b'])
      if cnt != 0:
        writer.writerows(data['mis_b'])
        writer.writerows([['==> 筆數：'+ str(cnt)]])
      else:
        writer.writerows([['單身無異常!']])
      writer.writerows([[]])

      # 寫入會計異常
      writer.writerows([title_check + title_fund])
      cnt = len(data['mis_a'])
      if cnt != 0:
        writer.writerows(data['mis_a'])
        writer.writerows([['==> 筆數：'+ str(cnt)]])
      else: 
        writer.writerows([['會計無異常!']])
      writer.writerows([[]])

      # 寫入認列異常
      writer.writerows([['認列碼編碼']])
      cnt = len(data['mis_s'])
      if cnt != 0:
        writer.writerows(data['mis_s'])
        writer.writerows([['==> 筆數：'+ str(cnt)]])
      else: 
        writer.writerows([['認列碼無異常!']])
      writer.writerows([[]])

      # 報表頁尾
      writer.writerows([['=== 以下空白 ===']])

    print(f"Data has been written to {file_path}")