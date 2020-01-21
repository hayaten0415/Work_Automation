#ライブラリ
import win32com.client
import datetime
import pathlib


# 欲しいファイルをprogress_report.pyと同じフォルダにコピーし
# コピーした先のコピーファイルのプルパスを返す
def fileCopy(filename):
  import shutil
  shutil.copy(filename, pathlib.Path(__file__).parent)
  return pathlib.Path.joinpath(pathlib.Path(__file__).parent, filename)

#Excel表読み込み
def readSheet(filename, sheetname):
  xl = win32com.client.Dispatch("Excel.Application")
  wb = xl.Workbooks.Open(Filename=filename,ReadOnly=1,UpdateLinks=False)
  xl.Visible = False
  sheet = wb.Worksheets(sheetname)
  sheet.Range("C4:I43").Copy()

def callOutlook():
  # コピーしたいファイルのフォルダ
  excelfolder = pathlib.Path(r'')
  # コピーしたいファイル
  excelfile = pathlib.Path.joinpath(excelfolder, 'hoge')
  copiedfile = fileCopy(excelfile)
  #Outlookのオブジェクト設定(この2行は固定)
  outlook = win32com.client.Dispatch("Outlook.Application")
  mymail = outlook.CreateItem(0)
  #メールの設定
  ##BodyFormatの値
  #1：テキストメール
  #2：HTMLメール
  #3：リッチテキストメール
  mymail.BodyFormat = 2
  mymail.To = "hoge"
  mymail.cc = "hoge"
  # メールの件名
  mymail.Subject = " "
  # メールの本文
  mymail.Body = """

""" + "\n"

  #Excel表読み込み
  readSheet(copiedfile, 'hoge')
  #出来上がったメール確認
  mymail.Display(True)


callOutlook()


#https://towel-memo.com/python/email_python/
#確認せず送信する場合は、mymail.Display(True)をコメントアウトして、下記コードのコメントアウトを外す
#mymail.Send()