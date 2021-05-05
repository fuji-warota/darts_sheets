import openpyxl
#import glob
import os

wb = openpyxl.Workbook()
sheet = wb.active
sheet.title = 'score_sheet'
#スコア項目作成
sheet['A1'] = "PLAYER NAME"
sheet['J1'] = "SCORE"
count = 1
for i in range(66,74):
  col = chr(i)+"1"
  sheet[col] = str(count)+"R"
  count += 1

#名前入力&セル設定
print("player name please")
print("input example(exp exp exp)\n")
name = list(map(str, input().split()))
#例外処理
if len(name) == 0:
  print("ERROR:No input")
  exit()
for i in range(len(name)):
  row = "A"+str(i+2)
  sheet[row] = name[i]
  row = "K"+str(i+2)
  com_area = "B"+str(i+2)+":J"+str(i+2)
  sheet[row] = '=SUM('+com_area+')'
#シート保存して終了
#wb.save('test.xlsx')
wb.save(os.environ['HOME']+'/Desktop/score_sheet.xlsx')
