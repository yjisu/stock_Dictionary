import warnings
#########5/23일작성중 #########
warnings.simplefilter(("ignore"))
import openpyxl
import pandas as pd

####### [날짜, 뉴스단어 한개] 구성 만드는 파일 차트 분석 시 count함수 이욜 할 때 참조 자료
# ex)hmm뉴스키워드날짜뉴스모으고특수삭제.xlsx 파일 넣음<- DayNewsMerge.py 중간에 주석처리 된 부분 해제하고 결과 얻기
Stockfilename = input("키워드파일이름입력:")
fileStock = "C:/Users/yangj/PycharmProjects/pythonProject1/샘플/" + Stockfilename + ".xlsx"
Stockfile = openpyxl.load_workbook(fileStock)  # 파일이름입력
stock_ws = Stockfile.active
Stock_data = []  # list 타입
date=[]
i = 0
for row in stock_ws.rows:
    Stock_data.append([])
    date.append(row[1].value)
    for cell in row:
        if cell.value != None :
            Stock_data[i].append(cell.value)
    i += 1
del Stock_data[0]
del date[0]
for i in range(len(Stock_data)):
    del Stock_data[i][0] #각 열의 첫번째 행 삭제
for i in range(len(Stock_data)):
    del Stock_data[i][0] #각 열의 첫번째 행 삭제
print(Stock_data)
print(date)
a=[] #
print(len(date),len(Stock_data))
for j in range(len(Stock_data)):
    for k in range(len(Stock_data[j])):
        a.append([date[j],Stock_data[j][k]])
print(a)
df_SourTar = pd.DataFrame(a)
df_SourTar.to_excel(Stockfilename+'countif.xlsx',sheet_name='sheet1')