#!/usr/bin/env python3
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import pandas as pd
import xlsxwriter
import xlrd
import datetime

#Server File location
loc2 = ('S:\OPERATIONS & ENGINEERING\Well DR Pressure\Dashboard - Wells.xlsm')
#Local file location for off network
#loc2 =('C:\\Users\camoruso\Desktop\Dashboard - Wells.xlsm')

sheetname = "WS Alt Data"
#login Info
F1 = open("C:\\Users\camoruso\Desktop\password.txt", "r")
contents = F1.readlines()
username = contents[0]
password = contents[1]
F1.close()

#Open Worksheet with Pandas
dfexcel = pd.read_excel(loc2, sheet_name = sheetname)
numrows = len(dfexcel.index)
#print(numrows, "numrows")
columns=dfexcel.columns
#print(columns)
column = columns[0]
#print(column)
Datelist = dfexcel[column].tolist()
#print(c)
#print(Datelist)
last = len(Datelist)
lastdate = Datelist[last-1]
#print(lastdate.date())

today = datetime.date.today()
Yesterday = datetime.date.today() + datetime.timedelta(days=-1)
if lastdate.date() < Yesterday:
	Begdate = lastdate.date() + datetime.timedelta(days=1)
	Begdate = Begdate.strftime('%m/%d/%Y')
	print(Begdate)
	Enddate = Yesterday
	Enddate = Enddate.strftime('%m/%d/%Y')
	print(Enddate)
else:
	Begdate = "05/02/2019"
	Enddate= "05/07/2019"

driver = webdriver.Chrome()
#"C:\Users\camoruso\Documents\Helpful Docs\Programming")
driver.get('https://www.gasstorage.net/WORSHAMSTEED/Operator/index.cfm')
element = driver.find_element_by_xpath('//*[@id="username"]')
element.send_keys(username)

element = driver.find_element_by_xpath('//*[@id="password"]')
element.send_keys(password)

element = driver.find_element_by_xpath('//*[@id="frmLogin"]/table[1]/tbody/tr[3]/td/input')
element.send_keys(Keys.RETURN)

element = driver.find_element_by_xpath('//*[@id="navmenu"]/table/tbody/tr[14]/td/a')
element.send_keys(Keys.RETURN)

element = driver.find_element_by_xpath('//*[@id="NetFacilStart_Disp"]')
element.clear()
element.send_keys(Begdate)

element = driver.find_element_by_xpath('//*[@id="NetFacilEnd_Disp"]')
element.clear()
element.send_keys(Enddate)

element = driver.find_element_by_xpath('//*[@id="content"]/div/div[3]/fieldset[3]/form/table/tbody/tr/td[5]/input')
element.send_keys(Keys.RETURN)

table = driver.find_element_by_xpath('//*[@id="content"]/div/div[3]/table/tbody')
table2 = driver.find_elements_by_xpath('//*[@id="content"]/div/div[3]/table/tbody/tr')
rows = table.find_elements_by_tag_name("tr")
print(rows)
x= []
for row in rows:
	print([td.text for td in row.find_elements_by_tag_name("td")])
	a = [td.text for td in row.find_elements_by_tag_name("td")]
	#dictionary = dict(zip(a,a))
	#print(dictionary)	
	x.append(a)

df = pd.DataFrame(x)
#Count number of rows copied over
numrows2 = len(df.index)
print(df[1:numrows2-1])
#Change dataframe to chop off header and total row
df2 = df[1:numrows2-1]
driver.close()
driver.quit()

#output
#print(df[1:numrows2])
#writer = pd.Excelwriter(localloc)
#Write to excel after the number of entries from origina numrows (Excel sheet count)
df2.to_excel(loc2,sheetname,startrow=numrows+2)
