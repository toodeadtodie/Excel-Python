# Excel-Python
Collecting data from different excel sheet to new one using python

# Below is the code


import glob
import openpyxl
import xlsxwriter
import os
location = 'Location of the file with name'
excel_files = glob.glob(location)
excel_files.sort(key=os.path.getmtime)
workbook = xlsxwriter.Workbook('Name of excel file that you want to create')
worksheet1 = workbook.add_worksheet('Ballast Data')
worksheet1.write('A1', 'Date')
worksheet1.write('B1', 'LT')
worksheet1.write('C1', 'UTC')
worksheet1.write('D1', 'RPM N/N')
worksheet1.write('E1', 'BF')
worksheet1.write('F1', 'Steaming Time')
worksheet1.write('G1', 'Distance')
worksheet1.write('H1', 'ME Consp.')
worksheet1.write('I1', 'AE Consp.')
row = 1
column = 0
count = 0
count1 = 0

for file in excel_files:
    book = openpyxl.load_workbook(file)
    sheet = book.active

    Dt = sheet['G5'].value.strftime("%m/%d/%y")
    LT = sheet['G6'].value.strftime("%H:%M")
    UTC = sheet['G7'].value.strftime("%H:%M")
    ST = sheet['I12'].value
    Dist = sheet['I11'].value
    remark = sheet['N34'].value
    ME1 = sheet['I16'].value
    ME2 = sheet['I17'].value
    AX1 = sheet['I18'].value
    AX2 = sheet['I19'].value
    AX3 = sheet['I20'].value
    AX4 = sheet['I21'].value
    AX5 = sheet['I22'].value
    AX6 = sheet['I23'].value
    BF = sheet.cell(row=45+count, column=21).value
    RPM = sheet.cell(row=45+count, column=16).value
    Main_Engine_consp = ME1+ME2
    Auxillary_Engine_consp = AX1+AX2+AX3+AX4+AX5+AX6
    count+=1
    if count > 8:
        BF2 = sheet.cell(row=45+count1,column=21).value
        RPM2 = sheet.cell(row=45+count1,column=16).value
        count1+=1
        BF=BF2
        RPM=RPM2
    content = [Dt,LT,UTC,RPM,BF,ST,Dist,Main_Engine_consp,Auxillary_Engine_consp]
    for n in content:
        worksheet1.write(row,column, n)
        column+=1
    row+=1
    column = 0
    
     
worksheet3 = workbook.add_worksheet('Ballast Events')
worksheet3.write('A1', 'Operation')
worksheet3.write('B1', 'From')
worksheet3.write('D1', 'To')
worksheet3.write('F1', 'Distance')
worksheet3.write('G1', 'ME Consp')
worksheet3.write('H1', 'AE Consp')
worksheet3.write('J1', 'From')
worksheet3.write('L1', 'DTG')
worksheet3.write('M1', 'ROB F.O')
worksheet3.write('N1', 'ROB D.O')
worksheet3.write('O1', 'Remark')

row =1
column = 0
r=1
c=9

for file in excel_files:
    book = openpyxl.load_workbook(file)
    sheet = book.active
    for i in sheet.iter_rows(min_row=15,min_col=14,max_row=24,max_col=25,values_only=True):
        my_list = list(i)
        if type(my_list[0]) == str:
            Event = my_list[0]
            from_date = my_list[2].strftime("%m/%d/%y")
            from_time = my_list[3].strftime("%H:%M")
            to_date = my_list[4].strftime("%m/%d/%y")
            to_time = my_list[5].strftime("%H:%M")
            distane = my_list[7]
            ME_consp = my_list[8]+my_list[10]
            AX_consp = my_list[9]+my_list[11]
            new_list = [Event,from_date,from_time,to_date,to_time,distane,ME_consp,AX_consp]
            for n in new_list:
                worksheet3.write(row,column, n)
                column+=1
            row+=1
            column = 0

    for j in sheet.iter_rows(min_row=30,min_col=15,max_row=32,max_col=23,values_only=True): 
        my_list1 = list(j)
        if type(my_list1[5]) == int:
            from_date1 = my_list1[1].strftime("%m/%d/%y")
            from_time1 = my_list1[2].strftime("%H:%M")
            DTG = my_list1[5]
            ROB_FO = my_list1[6]
            ROB_DO = my_list1[7]
            remark = my_list1[8]
            new_list1 = [from_date1,from_time1,DTG,ROB_FO,ROB_DO,remark]
            
            for m in new_list1:
                worksheet3.write(r,c,m)
                c+=1
            r+=1
            c=9

location = 'Location of file with name'
excel_files = glob.glob(location)
excel_files.sort(key=os.path.getmtime)
worksheet2 = workbook.add_worksheet('Laden Data')
worksheet2.write('A1', 'Date')
worksheet2.write('B1', 'LT')
worksheet2.write('C1', 'UTC')
worksheet2.write('D1', 'RPM N/N')
worksheet2.write('E1', 'BF')
worksheet2.write('F1', 'Steaming Time')
worksheet2.write('G1', 'Distance')
worksheet2.write('H1', 'ME Consp.')
worksheet2.write('I1', 'AE Consp.')
row = 1
column = 0
count = 0

for file in excel_files:
    book = openpyxl.load_workbook(file)
    sheet = book.active

    Dt = sheet['G5'].value.strftime("%m/%d/%y")
    LT = sheet['G6'].value.strftime("%H:%M")
    UTC = sheet['G7'].value.strftime("%H:%M")
    ST = sheet['I12'].value
    Dist = sheet['I11'].value
    remark = sheet['N34'].value
    ME1 = sheet['I16'].value
    ME2 = sheet['I17'].value
    AX1 = sheet['I18'].value
    AX2 = sheet['I19'].value
    AX3 = sheet['I20'].value
    AX4 = sheet['I21'].value
    AX5 = sheet['I22'].value
    AX6 = sheet['I23'].value
    BF = sheet.cell(row=45+count, column=21).value
    RPM = sheet.cell(row=45+count, column=16).value
    Main_Engine_consp = ME1+ME2
    Auxillary_Engine_consp = AX1+AX2+AX3+AX4+AX5+AX6
    count+=1
    content = [Dt,LT,UTC,RPM,BF,ST,Dist,Main_Engine_consp,Auxillary_Engine_consp]
    for n in content:
        worksheet2.write(row,column, n)
        column+=1
    row+=1
    column = 0
    if count > 16:
        count = 0

worksheet4 = workbook.add_worksheet('Laden Events')
worksheet4.write('A1', 'Operation')
worksheet4.write('B1', 'From')
worksheet4.write('D1', 'To')
worksheet4.write('F1', 'Distance')
worksheet4.write('G1', 'ME Consp')
worksheet4.write('H1', 'AE Consp')
worksheet4.write('J1', 'From')
worksheet4.write('L1', 'DTG')
worksheet4.write('M1', 'ROB F.O')
worksheet4.write('N1', 'ROB D.O')
worksheet4.write('O1', 'Remark')

row =1
column = 0
r=1
c=9

for file in excel_files:
    book = openpyxl.load_workbook(file)
    sheet = book.active
    for i in sheet.iter_rows(min_row=15,min_col=14,max_row=24,max_col=25,values_only=True):
        my_list = list(i)
        if type(my_list[0]) == str:
            Event = my_list[0]
            from_date = my_list[2].strftime("%m/%d/%y")
            from_time = my_list[3].strftime("%H:%M")
            to_date = my_list[4].strftime("%m/%d/%y")
            to_time = my_list[5].strftime("%H:%M")
            distane = my_list[7]
            ME_consp = my_list[8]+my_list[10]
            AX_consp = my_list[9]+my_list[11]
            new_list = [Event,from_date,from_time,to_date,to_time,distane,ME_consp,AX_consp]
            for n in new_list:
                worksheet4.write(row,column, n)
                column+=1
            row+=1
            column = 0
 
    for j in sheet.iter_rows(min_row=30,min_col=15,max_row=32,max_col=23,values_only=True): 
        my_list1 = list(j)
        if type(my_list1[5]) == int:
            from_date1 = my_list1[1].strftime("%m/%d/%y")
            from_time1 = my_list1[2].strftime("%H:%M")
            DTG = my_list1[5]
            ROB_FO = my_list1[6]
            ROB_DO = my_list1[7]
            remark = my_list1[8]
            new_list1 = [from_date1,from_time1,DTG,ROB_FO,ROB_DO,remark]
            
            for m in new_list1:
                worksheet4.write(r,c,m)
                c+=1
            r+=1
            c=9

workbook.close()
