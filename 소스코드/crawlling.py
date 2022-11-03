import warnings
warnings.simplefilter(("ignore"))
import konlpy
from konlpy.tag import *
import openpyxl
import pandas as pd
from math import  log10

#형태소분석라이브러리
#okt = Okt()
hannanum = Hannanum()
#filename= input("분석할 파일이름 입력:") #파일명
filefolder = input("종목폴더입력: ")
filename=input("파일이름입력:")
filepos = "C:/Users/yangj/PycharmProjects/pythonProject1/뉴스크롤링/"+filefolder+"/" + filename + ".xlsx"
kfile = openpyxl.load_workbook(filepos)#파일이름입력
sheet=kfile.worksheets[0]#sheet1에 있는 데이터 가죠오기
#print(sheet)
data=[]
for row in sheet.rows: #data에 크롤링한 뉴스 제목들 저장
    data.append(
            row[1].value
    )
#print(data)
#print(type(data[1])) #str
#newData=[]
newData2=[]
#for i in range(len(data)):
#    newData.append(okt.nouns(data[i])) #명사만 추출okt
#print(newData)
for i in range(len(data)-1):
    newData2.append(hannanum.nouns(data[i+1])) #명사만 추출hannanum가 okt보다 성능좋음
#print(newData2)
#print(type(newData2))#newData2 데이터 형식은 list
#df= pd.DataFrame.from_records(newData2)#newData2 dataframe으로 변환
#df.to_excel(filename+'_명사추출'+'.xlsx') #파일명의 엑셀로 변환

# -- TF-IDF function

def f(t, d):
    # d is document == tokens
    return d.count(t)

def tf(t, d):
    # d is document == tokens
    return 0.5 + 0.5*f(t,d)/max([f(w,d) for w in d])

def idf(t, D):
    # D is documents == document list
    numerator = len(D)
    denominator = 1 + len([ True for d in D if t in d])
    return log10(numerator/denominator)

def tfidf(t, d, D):
    return tf(t,d)*idf(t, D)

def tfidfScorer(D):
    result = []
    for d in D:
        result.append([(t, tfidf(t, d, D)) for t in d] )
    return result

#newData2는 명사추출을 통해 분리되어있음

if __name__ == '__main__':
    corpus=[]
    for i in range(len(newData2)):
        corpus.append(newData2[i])
    TfIf=[] #결과저장
    for i, result in enumerate(tfidfScorer(corpus)):
        print('====== document[%d] ======' % i)
        print(result)

        #TfIf.append(result)

#df= pd.DataFrame.from_records(TfIf)#TfIf dataframe으로 변환
#df.to_excel(filename+'_가중치추출'+'.xlsx')