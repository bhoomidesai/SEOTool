import openpyxl
import re
import urllib.request
from bs4 import BeautifulSoup
import urllib.request
import xlsxwriter
import sqlite3

fi=open("ignore.txt")

ignore=[]
linesi=str( fi.readlines())
l=linesi.split()

for line in l:
    ignore.append((line))

print(str(ignore))
    
fp=open("urls.txt")

L=[]
lines= fp.readlines()

 # OUTPUT SHEET INFO
workbook = xlsxwriter.Workbook('Analysis.xlsx')
heading = workbook.add_format({'bold':True,'font_color':'blue'})

conn = sqlite3.connect('mywords.db')

for line in lines:
    L.append((line))
   
for lee in L:
    link = lee
    print(link)
    req = urllib.request.Request( link, data=None, 
             headers={ 'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_9_3) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/35.0.1916.47 Safari/537.36' } 
                                      ) 
    page = urllib.request.urlopen(req) 
    soup = BeautifulSoup(page,"html.parser")    
    print(soup.title)			
    print(soup.title.string)
    for script in soup(["script", "style"]): 
	     script.extract() 
    text = soup.get_text()
    wordsset = []
    lineu = (line1.strip() for line1 in text.split())
    for line1 in lineu:
        print (line1)
        wordsset.append((line1))

    original=set(wordsset)
    ignoreset=set(ignore)
    myset={}
    myset=original-ignoreset
    totalwords=len(original)
    print("My set : ",myset)


    #adding worksheet with table headings
    worksheet = workbook.add_worksheet()
    worksheet.write("A1","Kewords",heading)
    title=str(soup.title.string)
    worksheet.write("A2",title,heading)
    worksheet.write("B1","No of Occurence",heading)
    worksheet.write("C1","Weightage",heading)
    mylist=list(myset)
    text=str(mylist)
    count=2
    for word in mylist:
        print(word,(wordsset.count(word)/totalwords)*100)
        temp=(wordsset.count(word)/totalwords)*100
        count+=1
        conn.execute("INSERT INTO WDENCITY (WORDS,DENCITY) \
          VALUES (?,?)",(word,temp))
       # cursor = conn.execute("SELECT WORDS, DENCITY from WDENCITY")
       # for row in cursor:
        #   print ("Words = ", row[0])
         #  print ("Dencity = ", row[1])
        keyword1 = re.findall(word,text)
        worksheet.write("A"+str(count),word)
        worksheet.write("B"+str(count),len(keyword1))
        worksheet.write("C"+str(count),(wordsset.count(word)/totalwords)*100)
    conn.commit()
    chart1 = workbook.add_chart({'type':'column'})
    s_name = worksheet.name
    formula = '='+s_name+'!A1:C'+str(count)
    print(formula)
    #adding Chart to Worksheet
    chart1.add_series({'values':formula})
    worksheet.insert_chart("F8",chart1)
    word=[]
    text=""
    #------------loop over
conn.close()    
workbook.close()

