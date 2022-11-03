import warnings

warnings.simplefilter(("ignore"))
import openpyxl
import pandas as pd

# 000_KNU_New_Vdic2.xlsx 파일 넣기

Stockfilefolder = input("종목시세폴더입력: ")
Stockfilename = input("시세파일이름입력:")
fileStock = "C:/Users/yangj/PycharmProjects/pythonProject1/뉴스키워드/" + Stockfilefolder + "/" + Stockfilename + ".xlsx"
Stockfile = openpyxl.load_workbook(fileStock)  # 파일이름입력
stock_ws = Stockfile.active
Stock_data = []  # list 타입
i = 0
for row in stock_ws.rows:
    Stock_data.append([])
    for cell in row:
        if cell.value != None:
            Stock_data[i].append(cell.value)
    i += 1
del Stock_data[0]
for i in range(len(Stock_data)):
    del Stock_data[i][0]
#print(Stock_data)


vert_p = []  # 수직 중복 삭제
for i in range(len(Stock_data)):
   vert_p.append([])
   for j in range(len(Stock_data[i])):
       vert_p[i].append(Stock_data[i][j])  # 단어만 넣기
print(vert_p)

vert_p.sort(key=lambda x: x[0])  # 단어 기준으로 정렬
for i in range(len(vert_p) - 2):  # 단어 비교해서 같으면 누적, 다르면 값 바꾸기
   for j in range(i + 1, len(vert_p)):
       if vert_p[i][0] == vert_p[j][0] :
           vert_p[i][1] += vert_p[j][1]
           vert_p[j] = ['0', 0]
       if str.isalnum(vert_p[i][0]) == False:
           vert_p[i] =['0', 0]

vert_p = [i for i in vert_p if not '0' in i]  # '0'들어간 열 제거
df_ver = pd.DataFrame(vert_p)
df_ver.to_excel(Stockfilename + ' Stock_dictionary2.xlsx', sheet_name='sheet1')
####사전 완성####
