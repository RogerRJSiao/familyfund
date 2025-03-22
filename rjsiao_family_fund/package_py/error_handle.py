import datetime as dt

def create_today_log(msg):
  # 格式化日期時間為字串
  now = dt.datetime.now()
  formatted_date = now.strftime("%Y%m%d")
  formatted_datetime = now.strftime("%Y%m%d_%H%M%S")

  # 設定儲存路徑、檔名(用日期戳)
  file_name = '家庭記帳比對_err_' + formatted_date + '.txt'
  file_path = "../mydata/err_log/" + file_name

  errmsg = f'{formatted_datetime} | ERROR: {msg}'
  print(errmsg)
  f = open(file_path, "a")  #a:追加,w:覆寫,r:讀取
  f.write(errmsg + "\n")
  f.close()