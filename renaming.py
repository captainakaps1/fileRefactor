import glob
import os
import xlrd
import shutil
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
import re

# create folder new_folder
# open excel file and parse through lines
#   take file names from C NAME index 1
#   take folder name from
#       LIS_CTRY index 3
#       C NAME index 1
# open folder matching folder name without capitalization
# search through file names for match
# if match:
#   copy file into new_folder
#   rename file as CompanyX_Year
#   highlight line in excel file
# else:
#   move to next line

ttt = []
count = 1
hold = ''
try:
    os.mkdir('./new_folder')
except:
    print('folder already exists')

loc = './excel.xlsx'
wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_name('Sheet1')

for p in os.listdir('.'):
    if os.path.isdir('./' + p):
        if p == '.idea' or p == 'new_folder':
            continue

        for i in range(sheet.nrows):
            if i == 0:
                continue

            folder_name_outer = sheet.cell_value(i, 3)
            folder_name_inner = sheet.cell_value(i, 1)

            if p.capitalize() == folder_name_outer.capitalize():
                try:
                    lsdir_o = os.listdir('./' + p)
                    try:
                        try:
                            lsdir_o.index(folder_name_inner)
                        except:
                            if lsdir_o[0].istitle():
                                folder_name_inner = folder_name_inner.capitalize()
                            elif lsdir_o[0].isupper():
                                folder_name_inner = folder_name_inner.upper()
                            elif lsdir_o[0].islower():
                                folder_name_inner = folder_name_inner.lower()
                        basepath = './' + folder_name_outer + '/' + folder_name_inner
                        try:
                            os.listdir(basepath)
                        except:
                            e = 0
                            while e < lsdir_o.__len__():
                                f = folder_name_inner.split(' ')
                                truth = False
                                for op in f:
                                    if op.upper() == folder_name_outer.upper():
                                        continue
                                    if lsdir_o[e].find(op):
                                        truth = True

                                if truth:
                                    folder_name_inner = lsdir_o[e]
                                e = e + 1
                        basepath = './' + folder_name_outer + '/' + folder_name_inner
                        for file_name in os.listdir(basepath):

                            q = ''
                            k = [2014, 2015, 2016, 2017]
                            print(q.join(re.findall("[0-9]", file_name)), i)
                            try:

                                k.index(int(q.join(re.findall("[0-9]", file_name))))

                                if os.path.isfile(os.path.join(basepath, file_name)):
                                    try:
                                        k.index(int(sheet.cell_value(i, 4)))
                                        # if sheet.row(x)[1] == folder_name_inner and sheet.row(x)[4] == q.join(
                                        # re.findall("[0-9]", file_name)):
                                        try:
                                            ttt.index(i)
                                        except:
                                            ttt.append(i)
                                            print(i, "ooooooo")
                                    except:
                                        print(i)

                                    src = basepath + '/' + file_name
                                    dst = './new_folder'
                                    shutil.copy2(src, dst)
                                    s = ""
                                    folder_name_inner = folder_name_inner.replace(' ', '')
                                    folder_name_inner = folder_name_inner.replace('-', '')

                                    os.rename(dst + '/' + file_name, dst + '/' + folder_name_inner + '_' + s.join(
                                        re.findall("[0-9]", file_name)) + '.pdf')



                            except:
                                print('hello')

                    except:
                        print('jello')

                except:
                    print('the folder was not found A')

file = 'excel.xlsx'
wb = load_workbook(filename=file)
ws = wb['Sheet1']

alpha = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H']
# Enumerate the cells in the second row
print(ttt)
print(ttt.__len__())
print(ws.max_column)
for row in ttt:
    col = 0
    while col <= ws.max_column:
        cell = ws[alpha[col] + str(row)]
        cell.fill = PatternFill(start_color="008000", end_color="008000", fill_type="solid")
        col = col + 1

wb.save(filename=file)
