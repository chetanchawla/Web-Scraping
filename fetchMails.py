from bs4 import BeautifulSoup
import urllib.request
import csv
from tablepyxl import tablepyxl
from tablepyxl.tablepyxl import get_Tables, write_rows
from premailer import Premailer

def document_to_one_sheet_xl(doc, filename):
    wb = document_to_one_sheet_workbook(doc)
    wb.save(filename)

def document_to_one_sheet_workbook(doc):
    wb = tablepyxl.Workbook()
    sheet = wb.active

    inline_styles_doc = Premailer(doc, remove_classes=False).transform()
    tables = get_Tables(inline_styles_doc)

    row = 1
    for table in tables:
        if table.head:
            row = write_rows(sheet, table.head, row)
        if table.body:
            row = write_rows(sheet, table.body, row)

    return wb
#Text file containing links to all colleges Faculty informatons
ins_data= open('C:/Users/asd/Desktop/insLinks2.txt','r')
allLinks=[]
i=0
totalTable=""
for line in ins_data:
    allLinks.append(line)
    i=i+1
ins_data.close()

for college in allLinks:
	url = college
	try:
		content = urllib.request.urlopen(url).read()
	except:
		print("unable to parse data from the link")
		continue

	soup = BeautifulSoup(content, "lxml")

	#Finding all tables on the given page
	tables=soup.find_all('table')

        #Getting the header from the table
	headerTag=soup.find('div',id='inst-name')
	#converting to string format
	headerList=[]
	for h in headerTag:
	    headerList.append(str(h))
	#extracting only the name and ignoring everything else
	header= ''.join([char for char in headerList[1] if(char.isupper() or char is " ")])
	#Selecting the desired table with faculty details

	try: table=tables[3]
	except:
		print("data not present for", header)
		continue

	#Getting Table to excel
	#Converting table data to list
	tableDash=[]
	for x in table:
	    tableDash.append(str(x))
	#fetching only table in string format and adding header of College
	tableDash[1]= "<table>" + "<thead>" +"<tr>" + "<th>" + header + "</th>" + "</tr>" + "</thead>"+ tableDash[1] + "</table>"
	totalTable= totalTable + tableDash[1]
	print("appended for ", header)
#converting to csv format
totalTable=totalTable
f = open('table1.html','w')
f.write(totalTable)
f.close()
document_to_one_sheet_xl(totalTable, "C:/Users/asd/Desktop/table2.csv")
tablepyxl.document_to_xl(totalTable, "C:/Users/asd/Desktop/table1.csv")

