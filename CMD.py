import os
import xlrd
from xlutils.copy import copy

pre_command = r"C:\Users\Autobio-A3517\Desktop\UnitTest\UnitTest.exe 192.168.2.249 DSPRO-zcy test 7895123"
wb = xlrd.open_workbook(r'C:\Users\Autobio-A3517\Desktop\test\123.xlsx')
wb2 = copy(wb)
sheet = wb.sheet_by_index(0)
n_rows = sheet.nrows
n_cols = sheet.ncols
i = 1
while i <= n_rows-1:
    para1 = sheet.cell(i,1).value
    para2 = sheet.cell(i,3).value
    para3 = sheet.cell(i,5).value
    post_comamnd = (str(para1)+' '+str(para2)+' '+str(para3))
    command = pre_command +' '+post_comamnd
    response = os.popen(command)
    f = response.readlines()
    for line in f:
        if line.startswith('1'):
            a = line.split('|')
            b = a[1].strip()+'|'+a[2].strip()+'|'+a[3].strip()          
            table = wb2.get_sheet(0)
            result = table.write(i,n_cols+1,b)
            wb2.save(r'C:\Users\Autobio-A3517\Desktop\test\1234.xls')
    i +=1





