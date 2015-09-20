import random
from bs4 import BeautifulSoup
import time
import xlwt
import os.path
import xlrd    

abspath = os.path.abspath(__file__)
dname = os.path.dirname(abspath)
os.chdir(dname)

def importList():
    f = open('./UMMS M1 Lunch Love.xls', 'rt')
    soup = BeautifulSoup(f, "html.parser")
    table = soup.find("table")
    datasets = []
    for row in table.find_all("tr")[1:]:
        dataset = zip(td.get_text() for td in row.find_all("td"))
        datasets.append(''.join(dataset[1]) + ', ' + ''.join(dataset[2]))
    return datasets

def pastPairs():    
    pastpairs = []
    book = xlrd.open_workbook("./pastpairs.xls")
    sheet = book.sheet_by_name("sheet1")
    row = sheet.row(0)        
    for row_idx in range(1, sheet.nrows):
        pastpairs.append([sheet.cell_value(row_idx, 0), sheet.cell_value(row_idx, 1)])
    return pastpairs
    
def popRandom(lst):
    idx = random.randrange(0,len(lst))
    return lst.pop(idx)

def createPairs():
    pairs = []
    lst = importList()   
    while len(lst)>0:
        rand1 = popRandom(lst)
        rand2 = popRandom(lst)
        pair = [rand1, rand2]
        pairs.append(pair)
    pairs.insert(0,['M1 Lunch Love', (time.strftime("%m/%d/%Y"))])
    return pairs

def checkPairs(pairs):
    pastpairs = pastPairs()
    for idp in range(0, len(pairs)):
        if sorted(pairs[idp]) in sorted(pastpairs):
            return True
        
def writeExl(dataset, name):
	from tempfile import TemporaryFile
	book = xlwt.Workbook()
	sheet1 = book.add_sheet('sheet1')
	for i,e in enumerate(list(zip(*dataset)[0])):
		sheet1.write(i,0,e)
	for i,e in enumerate(list(zip(*dataset)[1])):
		sheet1.write(i,1,e)
	pathname = "./" + name
	book.save(pathname)
	book.save(TemporaryFile())

def main():
    pairs = createPairs()
    if os.path.isfile("../pastpairs.xls"):
        while True == checkPairs(pairs):
            pairs = createPairs()
    writeExl(pairs, "pairs.xls")
    if os.path.isfile("./pastpairs.xls"):
        writeExl(pairs[1:] + pastPairs(), "pastpairs.xls")
    else:
        writeExl(pairs[1:], "pastpairs.xls")
    print 'Complete'

main()
