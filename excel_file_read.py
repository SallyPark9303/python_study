# 1. excel 파일에서 필요한 데이터 추출
# 2. db에 저장

import pandas as pd
import os
from tkinter import filedialog
from tkinter import messagebox

list_file = []                                          #파일 목록 담을 리스트 생성
files = filedialog.askopenfilenames(initialdir="/",\
                 title = "파일을 선택 해 주세요",\
                    filetypes = (("*.xlsx","*xlsx"),("*.xls","*xls"),("*.csv","*csv")))
#files 변수에 선택 파일 경로 넣기

if files == '':
    messagebox.showwarning("경고", "파일을 추가 하세요")    #파일 선택 안했을 때 메세지 출력

print(files)
df = pd.read_excel(files[0], engine='openpyxl')
df.fillna(0)
# pymssql 패키지 import
import pymssql
 
# MSSQL 접속
conn = pymssql.connect(host='DESKTOP-6VEJ2T7', database='shop_prj', charset='utf8')
 
# Connection 으로부터 Cursor 생성
cursor = conn.cursor()
 
for item in df.values:
     val = ''
  
     for idx,i in enumerate(item):
          print("i내용", i)
         
          if  pd.isna(i): 
          #  i =''
            item[idx] = ''
     
     # db에 저장
     # SQL문 실행
     sql = 'insert into [dbo].[megaItems] ([itemcode] ,[channelNo],[addr1],[addr2],[productName],[price]) values (%s,%s,%s,%s,%s,%s)'
     
     val = (item[0],str(item[1]),(item[2]),item[3],item[4],item[5])
     cursor.execute(sql,val)
     conn.commit()

# 연결 끊기
conn.close()