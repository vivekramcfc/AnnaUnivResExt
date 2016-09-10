from mechanize import Browser
from bs4 import BeautifulSoup
import xlwt

result=xlwt.Workbook(encoding="utf-8")

s1=result.add_sheet("Sheet 1")
row1=["S.No.","REG. NO.","NAME","CS6001","CS6601","CS6611","CS6612","CS6659","CS6660","GE6674","IT6502","IT6601"]

for i,s in enumerate(row1):
	s1.write(0,i,s)

stud=[]
for i in range(0,58):
	stud.append("81001310"+str(input()))
	
print "Done"

for i,reg in enumerate(stud):
	browser=Browser()
	browser.open("http://aucoe.annauniv.edu/result/134679852/cgrade.html")

	browser.select_form(nr=0)
	browser['regno']=reg

	response=browser.submit()

	content=response.read()

	soup=BeautifulSoup(content)
	table=soup.find('table')
	header=[]
	for strong in table.find_all('strong'):
		header.append(str(strong.text).replace('\n','').replace('\r',''))
	s1.write(i+1,0,str(i+1))
	s1.write(i+1,1,header[0])
	s1.write(i+1,2,header[1])
	j=3;
	for s in row1[3:]:
		pos=-1
		pos=header.index(s) if s in header else None;
		if pos>=0:
			s1.write(i+1,j,header[pos+1])
		else:
			s1.write(i+1,j,"-")
		j+=1

result.save("result2.xls")
