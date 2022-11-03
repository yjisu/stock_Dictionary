# KNU 한국어 감성사전
# 작성자 : 온병원, 박상민, 나철원
# 소속 : 군산대학교 소프트웨어융합공학과 Data Intelligence Lab
# 홈페이지 : dilab.kunsan.ac.kr
# 작성일 : 2018.05.14
# 뜻풀이 데이터 출처 : https://github.com/mrchypark/stdkor
# 신조어 데이터 출처 : https://ko.wikipedia.org/wiki/%EB%8C%80%ED%95%9C%EB%AF%BC%EA%B5%AD%EC%9D%98_%EC%9D%B8%ED%84%B0%EB%84%B7_%EC%8B%A0%EC%A1%B0%EC%96%B4_%EB%AA%A9%EB%A1%9D
# 이모티콘 데이터 출처: https://ko.wikipedia.org/wiki/%EC%9D%B4%EB%AA%A8%ED%8B%B0%EC%BD%98
# SentiWordNet_3.0.0_20130122 데이터 출처 : http://sentiwordnet.isti.cnr.it/
# SenticNet-5.0 데이터 출처 : http://sentic.net/
# 감정단어사전0603 데이터 출처 : http://datascience.khu.ac.kr/board/bbs/board.php?bo_table=05_01&wr_id=91 
# 김은영, “국어 감정동사 연구”, 2004.02, 학위논문(박사) - 전남대학교 국어국문학과 대학원

#-*-coding:utf-8-*-
import collections
import json

import warnings
warnings.simplefilter(("ignore"))
import openpyxl
import pandas as pd
import re
from datetime import datetime

############종목 감성 판단 ex)hmm뉴스키워드.xlsx 파일 넣는 과정
class KnuSL():
    
    def data_list(wordname):
        with open('KnuSentiLex/data/SentiWord_info.json', encoding='utf-8-sig', mode='r') as f:
            data = json.load(f)
            result = [0,0]
            
        for i in range(0, len(data)):
            if data[i]['word'] == wordname:
                result.pop()
                result.pop()
                result.append(data[i]['word_root'])
                result.append(int(data[i]['polarity']))

        r_word = result[0] #어근
        s_word = result[1] #극성
        
        return s_word

if __name__ == "__main__":

    ksl = KnuSL
    
    print("\nKNU 한국어 감성사전입니다~ :)")
    print("사전에 단어가 없는 경우 결과가 None으로 나타납니다!!!")
    print("종료하시려면 #을 입력해주세요!!!")
    print("-2:매우 부정, -1:부정, 0:중립 or Unkwon, 1:긍정, 2:매우 긍정")
    print("\n")
#########
Newsfilefolder = input("종목폴더입력: ")
Newsfilename=input("파일이름입력:")
Newsfilepos = "C:/Users/yangj/PycharmProjects/pythonProject1/뉴스키워드/"+Newsfilefolder+"/" + Newsfilename + ".xlsx"
Newsfile = openpyxl.load_workbook(Newsfilepos)#파일이름입력
ws=Newsfile.active
data=[]
date=[]
i=0
for row in ws.rows:
    data.append([])
    date.append(row[1].value)
    for cell in row:
        if cell.value != None:
            data[i].append(cell.value)
    i += 1
del data[0] #첫번째 의미없는 열 삭제
del date[0]
for i in range(len(data)):
    del data[i][0] #각 열의 첫번째 행 삭제
for i in range(len(data)):
    del data[i][0] #각 열의 날짜 행 삭제

KNUdata=[]
Tdata=[]

for x in range(len(data)):
    KNUdata.append([])
    Tdata.append([])
    for y in range(len(data[x])):
        KNUdata[x].append(ksl.data_list(data[x][y]))
        Tdata[x].append([data[x][y], KNUdata[x][y]])

result = { '날짜':date, '단어, 극성':Tdata }

df = pd.DataFrame(result)
#print(df)
list_df=df.values.tolist() #dataframe list로 변경
#print(list_df)
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
            if new_date[v] == list_df[i][0]:
                Setlist[v].append(list_df[i][1][j])
print(Setlist)
print(Setlist[0][0]) #2021.01.01
print(type(Setlist[0][0]))
print(Setlist[0][0].split('-'))
print(Setlist[0][1][1]) #극성 0
print(type(Setlist[0][1][1])) #극성 모든 타입 int

#print(list_df[0][1][0]) 키워드와 극성 ['HMM…"체질개선해', 'X']
#print(list_df[0][1][0][1]) 극성 x
#print(list_df[0][0].split('.')[:3]) ['2021', '01', '01']
#df.to_excel(Newsfilename+' KNU.xlsx',sheet_name='sheet1')

Stockfilefolder = input("종목시세폴더입력: ")
Stockfilename=input("시세파일이름입력:")
fileStock = "C:/Users/yangj/PycharmProjects/pythonProject1/종목별시세/"+Stockfilefolder+"/" + Stockfilename + ".xlsx"
Stockfile = openpyxl.load_workbook(fileStock)#파일이름입력
stock_ws=Stockfile.active
Stock_data=[] #list 타입
i=0
for row in stock_ws.rows:
    Stock_data.append([])
    for cell in row:
        if cell.value != None:
            Stock_data[i].append(cell.value)
    i += 1
del Stock_data[0]
for i in range(len(Stock_data)):
    del Stock_data[i][2] # 대비 삭제
for i in range(len(Stock_data)):
    del Stock_data[i][7] #거래대금 삭제
for i in range(len(Stock_data)):
    del Stock_data[i][7] #시가 총액 삭제
for i in range(len(Stock_data)):
    del Stock_data[i][7] #상장주식 수 삭제 / 결과:'일자', '종가', '등락률', '시가', '고가', '저가', '거래량'
#print(Stock_data)

def Calpercentage(a,b): #시초가 대비 고점/저점 비율
    return abs(a-b)/a*100
####아래로 수정 필요 (미완성)####

i=0
for k in range(len(Setlist)):
    if( Stock_data[i][0].split('/') == Setlist[k][0].split('.')[:3]): # 날짜 비교 날짜가 같다면
          if Calpercentage(Stock_data[i][3],Stock_data[i][4]) > 2 : #당일 시가 대비 고가가 2퍼 높을때
             for j in range(1,len(Setlist[k])):
                 if Setlist[k][j][1] == 0:
                    Setlist[k][j][1] = 1
                 else:
                    Setlist[k][j][1] += 1
          elif Calpercentage(Stock_data[i][3],Stock_data[i][5]) < -2 : #당일 시가 대비 저가가 2퍼 낮을 때
            for j in range(1,len(Setlist[k])):
                if Setlist[k][j][1] == 0:
                    Setlist[k][j][1] = -1
                else:
                     Setlist[k][j][1] -= 1
          else:
                if Stock_data[i+1][2] > 0:          # 다음날 주가 등락률이 양수면
                    for j in range(1,len(Setlist[k])):   #어제뉴스는 호재 취급
                        if Setlist[k][j][1] == 0:
                            Setlist[k][j][1] = 1
                        else:
                            Setlist[k][j][1] += 1
                elif Stock_data[i+1][2] < 0:
                    for j in range(1,len(Setlist[k])):  # 음수면 어제 뉴스는 악재 취급
                        if Setlist[k][j][1] == 0:
                            Setlist[k][j][1] = -1
                        else:
                            Setlist[k][j][1] -= 1
          i += 1

    else:
        if Calpercentage(Stock_data[i][3], Stock_data[i][4]) > 2:  # 당일 시가 대비 고가가 2퍼 높을때
            for j in range(1, len(Setlist[k])):
                if Setlist[k][j][1] == 0:
                    Setlist[k][j][1] = 1
                else:
                    Setlist[k][j][1] += 1
        elif Calpercentage(Stock_data[i][3], Stock_data[i][5]) < -2:  # 당일 시가 대비 저가가 2퍼 낮을 때
            for j in range(1, len(Setlist[k])):
                if Setlist[k][j][1] == 0:
                    Setlist[k][j][1] = -1
                else:
                    Setlist[k][j][1] -= 1
        else:
            if Stock_data[i + 1][2] > 0:  # 다음날 주가 등락률이 양수면
                for j in range(1, len(Setlist[k])):  # 어제뉴스는 호재 취급
                    if Setlist[k][j][1] == 0:
                        Setlist[k][j][1] = 1
                    else:
                        Setlist[k][j][1] += 1
            elif Stock_data[i + 1][2] < 0:
                for j in range(1, len(Setlist[k])):  # 음수면 어제 뉴스는 악재 취급
                    if Setlist[k][j][1] == 0:
                        Setlist[k][j][1] = -1
                    else:
                        Setlist[k][j][1] -= 1

        i += 1
        #<이거 삭제서 hmm한번 더 돌려보기

print(Setlist)

#df_Setlist = pd.DataFrame(Setlist)
#df_Setlist.to_excel(Stockfilename+' KNU_New.xlsx',sheet_name='sheet1')

Setlist_w = []
for i in range(len(Setlist)):
    Setlist_w.append([])
    for j in range(1, len(Setlist[i])):
        Setlist_w[i].append(Setlist[i][j][0])  # 극성 제외 단어만 추출

counter = {}
for i in range(len(Setlist_w)):
    counter[i] = collections.Counter(Setlist_w[i])  # 누적치

for i in range(len(Setlist_w)):
    Setlist_w[i] = list(zip(counter[i].keys(), counter[i].values()))  # 튜플 리스트화 [(값, 값)]

Plist = []
for i in range(len(Setlist_w)):
    Plist.append([])
    for j in range(len(Setlist_w[i])):
        Plist[i].append(list(Setlist_w[i][j]))  # 튜플 -> 리스트화 [[값, 값]]

for i in range(len(Plist)):
    for j in range(len(Plist[i])):
        Plist[i][j][1] = 0  # 극성 0으로 초기화

for i in range(len(Setlist)):
    for j in range(1, len(Setlist[i])):
            for h in range(len(Plist[i])):
                if Setlist[i][j][0] == Plist[i][h][0]:
                    Plist[i][h][1] += Setlist[i][j][1] #누적치
vert_p=[] #수직 중복 삭제
for i in range(len(Plist)):
    for j in range(len(Plist[i])):
        vert_p.append(Plist[i][j]) #단어만 넣기
#print(vert_p)
vert_p.sort(key=lambda x:x[0]) #단어 기준으로 정렬
for i in range(len(vert_p)-2): #단어 비교해서 같으면 누적 다르면 값 바꾸기
    for j in range(i+1,len(vert_p)):
        if vert_p[i][0] == vert_p[j][0]:
            vert_p[i][1]+=vert_p[j][1]
            vert_p[j]=['0',0]
print(vert_p)
vert_p=[i for i in vert_p if not '0' in i] #'0'들어간 열 제거
df_ver= pd.DataFrame(vert_p)
df_ver.to_excel(Stockfilename+' KNU_New_Vdic2.xlsx',sheet_name='sheet1')

####키워드파일 월별로 돌려서 그 나온 결과 파일들을 합쳐서 Merge_dictionay.py에 넣어서 사전 만들기 ####


