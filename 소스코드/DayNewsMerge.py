import warnings

warnings.simplefilter(("ignore"))
import openpyxl
import pandas as pd

#######Gephi에 사용할 edge파일 만들기 전에 필요한 자료 만드는 과정
####6개월치 키워드 합친 키워드 파일 넣기

Stockfilename = input("키워드파일이름입력:")
fileStock = "C:/Users/yangj/PycharmProjects/pythonProject1/" + Stockfilename + ".xlsx"
Stockfile = openpyxl.load_workbook(fileStock)  # 파일이름입력
stock_ws = Stockfile.active
Stock_data = []  # list 타입
date=[]
i = 0
for row in stock_ws.rows:
    Stock_data.append([])
    date.append(row[1].value)
    for cell in row:
        if cell.value != None:
            Stock_data[i].append(cell.value)
    i += 1
del Stock_data[0] #첫번째 의미없는 열 삭제
del date[0]
for i in range(len(Stock_data)):
    del Stock_data[i][0] #각 열의 첫번째 행 삭제
for i in range(len(Stock_data)):
    del Stock_data[i][0] #각 열의 날짜 행 삭제


Tdata=[]

for x in range(len(Stock_data)):
    Tdata.append([])
    for y in range(len(Stock_data[x])):
        if str.isalnum(Stock_data[x][y]) == True:
                  Tdata[x].append(Stock_data[x][y])

result = { '날짜':date, '단어':Tdata }

df = pd.DataFrame(result)
#print(df)
list_df=df.values.tolist() #dataframe list로 변경
print(list_df)
#print(list_df[0][0]) 날짜 2021.01.01.

new_date = [] # 날짜 중복 삭제
for v in date:
    if v not in new_date:
        new_date.append(v)
#print(new_date)

Setlist =[]# 날짜별 키워드 넣기
for v in range(len(new_date)):
    Setlist.append([])
    Setlist[v].append(new_date[v])
    for i in range(len(list_df)):
        for j in range(len(list_df[i][1])):
            if new_date[v] == list_df[i][0] :
                Setlist[v].append(list_df[i][1][j])
print(Setlist)
df_ver= pd.DataFrame(Setlist)
df_ver.to_excel(Stockfilename+' 날짜뉴스모으고특수삭제.xlsx',sheet_name='sheet1')
SourceTarget=[]
for i in range(len(list_df)):
    SourceTarget.append([])
    for j in range(len(list_df[i][1])-1):
        SourceTarget.append([list_df[i][0],list_df[i][1][j],list_df[i][1][j+1],1])
print(SourceTarget)
SourceTarget = [v for v in SourceTarget if v]
df_SourTar = pd.DataFrame(SourceTarget)
#df_SourTar.to_excel(Stockfilename+'Edge3.xlsx',sheet_name='sheet1')