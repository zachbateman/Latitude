#!/usr/bin/env python3
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import pandas as pd

Begdate = "05/02/2019"
Enddate= "05/07/2019"
driver = webdriver.Chrome()
#"C:\Users\camoruso\Documents\Helpful Docs\Programming")
driver.get('https://www.gasstorage.net/WORSHAMSTEED/Operator/index.cfm')
element = driver.find_element_by_xpath('//*[@id="username"]')
element.send_keys("camoruso")

element = driver.find_element_by_xpath('//*[@id="password"]')
element.send_keys("*****")

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

#element = driver.find_element_by_xpath('//*[@id="content"]/div/div[3]/table/tbody')
#for row in driver.find_element_by_xpath('//*[@id="content"]/div/div[3]/table'):
#	tds = tr.find_elements_by_tag_name('td')
#	print([td.text for td in tds])

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
print(df)

