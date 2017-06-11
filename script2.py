#! python2

from mechanize import Browser
from bs4 import BeautifulSoup
import xlwt

result=xlwt.Workbook(encoding="utf-8")

s1=result.add_sheet("Sheet 1")
row1=["S.No.","REG. NO.","NAME","CS6010","CS6801","CS6811","GE6075","CS6008"]

for i,s in enumerate(row1):
	s1.write(0,i,s)

stud=[]
for i in range(0,10):
	stud.append("81001310400"+str(i))
	
for i in range(10,100):
	stud.append("8100131040"+str(i))

for i in range(100,120):
	stud.append("810013104"+str(i))
	
for i in range(301,325):
	stud.append("810013104"+str(i))
	
for i in range(701,717):
	stud.append("810013104"+str(i))

	
	
print "Done"

for i,reg in enumerate(stud):
	browser=Browser()
	browser.set_handle_equiv(True)
	browser.set_handle_gzip(True)
	browser.set_handle_redirect(True)
	browser.set_handle_referer(True)
	browser.set_handle_robots(False)
	browser.addheaders = [('User-Agent', 'Mozilla/5.0 (X11; U; Linux i686; en-US; rv:1.9.0.1) Gecko/2008071615 Fedora/3.0.1-1.fc9 Firefox/3.0.1'), ('Accept', '*/*')]
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