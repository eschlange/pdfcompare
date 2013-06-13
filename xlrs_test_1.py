from mmap import mmap,ACCESS_READ

from xlrd import open_workbook

print open_workbook('test.xls')

with open('test.xls','rb') as f:

 print open_workbook(

 file_contents=mmap(f.fileno(),0,access=ACCESS_READ)

 )

aString = open('test.xls','rb').read()

print open_workbook(file_contents=aString)
