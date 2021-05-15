import os
import pandas as pd
from xlrd import open_workbook
from xlutils.copy import copy
import xlwt

folder = '/<Folder location>'
files = os.listdir(folder)

complete_list = []

count_success = 0
count_failure = 0

for x in files:
    try:
        individual_list = []
        wb = open_workbook(folder+x)
        sheet = wb.sheet_by_index(0)
        val_a = sheet.cell_value(1,0)
        val_b = sheet.cell_value(1,1)
        start = 86
        for i in range(0,8):
            if(sheet.cell_value(start+i,0)=="<Value to be searched>"):
                val_c = sheet.cell_value(start+i+1, 0)
                val_d = sheet.cell_value(start+i+1, 1)
                val_e = sheet.cell_value(start+i+1, 2)
                break
            else:
                continue
        individual_list.append(val_a)
        individual_list.append(val_b)
        individual_list.append(val_c)
        individual_list.append(val_d)
        individual_list.append(val_e)
        complete_list.append(individual_list)
        count_success+=1
        print(x+" - Success")
    except:
        count_failure+=1
        print(x+ "- Failed")

print("Total files successfully scanned : "+str(count_success))
print("Total files that were not scanned: "+str(count_failure))

def tabulate_data():
    excel_file='stats.xlsx'
    wb=open_workbook(excel_file)
    sheet=wb.sheet_by_name('Sheet1')
    data = copy(wb)
    current_row = 1
    worksheet = data.get_sheet('Sheet1')
    for x in complete_list:
        worksheet.write(current_row,0,x[0])
        worksheet.write(current_row,1,x[1])
        worksheet.write(current_row,2,x[2])
        worksheet.write(current_row,3,x[3])
        worksheet.write(current_row,4,x[4])
        current_row += 1
    data.save(excel_file)
    print("Saved to file")
    
tabulate_data()
