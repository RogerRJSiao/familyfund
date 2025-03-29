import gspread
from oauth2client.service_account import ServiceAccountCredentials

from package_py import error_handle as eh
# from .error_handle import create_today_log

def get_google_sheets(spreadsheet_id, auth, credentials_file_path, conn_sheets):
  # credentials_file_path 引用的 json 寫在 "client_email" (讀取替身)，需要設定為 viewer 權限。
  # spreadsheet_id 必須是 Google sheet id (長度較長)，不可為自行上傳的 .xlsx id (出現 sheets.worksheet(str) Error)。
  try:
    scopes = ["https://spreadsheets.google.com/feeds"]
    credentials = ServiceAccountCredentials.from_json_keyfile_name(credentials_file_path, scopes)
    client = gspread.authorize(credentials)
    sheets = client.open_by_key(spreadsheet_id)
  except:
    errmsg = f'Google sheets 檔案-{conn_sheets} 連線發生異常，請先檢查：'
    eh.create_today_log(errmsg)
    errmsg = f'1. 檔案-{conn_sheets} 的 google sheet id 是否為【{spreadsheet_id}】？'
    eh.create_today_log(errmsg)
    errmsg = f'2. 是否已把 {auth} 以檢視者 (viewer) 權限加入這個 google sheet 檔案-{conn_sheets}？'
    eh.create_today_log(errmsg)
    errmsg = f'3. 憑證 json ({credentials_file_path}) 是否放在本地端指定資料夾？'
    eh.create_today_log(errmsg)
    sheets = ''

  return sheets


def get_sheet_data(sheets, filename, sheetname):
  try:
    worksheet = sheets.worksheet(sheetname)
    mydata = worksheet.get_all_records()
  except:
    worksheet_list = sheets.worksheets()
    if sheetname in worksheet_list:
      errmsg = f'下載問題! 請確認 Google sheet 檔案-{filename} 活頁-{sheetname} 內的所有欄名唯一沒有重複名稱，且右側空白處無任何資料。'
    else:
      errmsg = f'下載問題! 請確認 Google sheet 檔案-{filename} 活頁-{sheetname} 是否存在'
    eh.create_today_log(errmsg)
    mydata = ''
  
  return mydata


