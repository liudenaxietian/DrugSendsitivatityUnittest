import psycopg2
import xlrd
import xlwt
from xlutils.copy import copy
import argparse

conn = psycopg2.connect(database="DSPRO-zcy", user="test", password="7895123", host="192.168.2.249", port="5432")
cur = conn.cursor()
extract_para = cur.execute(r"SELECT * FROM public.test_agentmodifyr")
paras = cur.fetchall()

def Drug_all_report_test(file_name):
    n = 0
    while n <= len(paras)-1:
        para = paras[n]
        command = r"SELECT * from public.f_get_expert_rules(array["+str(para[0])+r"], array["+str(para[1])  +r"])"
        # print(command)
        # print(para[1])
        test_value = cur.execute(command)
        test_result = cur.fetchone()
        # print(test_result[13])
        str1 = str(para[0])+str('all')
        str2 = str(test_result[0])+str(test_result[11])
        # print(str1)
        # print(str2)
        wb = xlrd.open_workbook(file_name)
        write_excel = copy(wb)
        sheet = write_excel.get_sheet('ReportAll_1')
        if str1 == str2:
            if str(para[5]) == 'True' and test_result[13][0] == para[1]:
                table = sheet.write(n,1,str(para[0]))
                table = sheet.write(n,2,str(para[1]))
                table = sheet.write(n,3,str(str(para[0])+' | '+str('all')+' | '+'True'))
                table = sheet.write(n,4,str(str(test_result[0])+' | '+str(test_result[11])))
                table = sheet.write(n,5,'Pass')
            elif str(para[5]) == 'False' and str(test_result[13]) == 'None':
                table = sheet.write(n,1,str(para[0]))
                table = sheet.write(n,2,str(para[1]))
                table = sheet.write(n,3,str(str(para[0])+' | '+str('all')+' | '+'False'))
                table = sheet.write(n,4,str(str(test_result[0])+' | '+str(test_result[11])))
                table = sheet.write(n,5,'Pass')
            else:
                table = sheet.write(n,1,str(para[0]))
                table = sheet.write(n,2,str(para[1]))
                table = sheet.write(n,3,str(str(para[0])+' | '+str('all')))
                table = sheet.write(n,4,str(str(test_result[0])+' | '+str(test_result[11])))
                table = sheet.write(n,5,'Fail')
        else:
            table = sheet.write(n,1,str(para[0]))
            table = sheet.write(n,2,str(para[1]))
            table = sheet.write(n,3,str(str(para[0])+' | '+str('all')))
            table = sheet.write(n,4,str(str(test_result[0])+' | '+str(test_result[11])))
            table = sheet.write(n,5,'Fail')
        n +=1
        write_excel.save(file_name)   


def get_parser():
    parser = argparse.ArgumentParser(description="used for directly calling in the commandline for drug expeter test")
    parser.add_argument('file_name',type = str,nargs = 1,help = 'the work directory where the work excel in')

    return parser

def main():
    parser = get_parser()
    arg = vars(parser.parse_args())
    file_name = arg['file_name'][0]

    Drug_all_report_test(file_name)
              
if __name__=="__main__":
    # Drug_all_report_test(r"C:\Users\Autobio-A3517\Desktop\ExpertReportAll123.xls")
    main()





