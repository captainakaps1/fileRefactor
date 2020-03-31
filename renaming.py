import glob
import os
import xlrd
import shutil
import zipfile
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
#

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
            elif hold == sheet.cell_value(i, 1):
                continue

            hold = sheet.cell_value(i, 1)

            folder_name_outer = sheet.cell_value(i, 3)
            folder_name_inner = sheet.cell_value(i, 1)

            if p.capitalize() == folder_name_outer.capitalize():

                try:
                    lsdir_o = os.listdir('./' + p)
                    h = 0
                    while h < lsdir_o.__len__():
                        lsdir_o[h] = lsdir_o[h].upper()
                        h = h + 1
                    try:
                        #print(lsdir_o)

                        if lsdir_o[0].istitle():
                            folder_name_inner =  folder_name_inner.capitalize()
                        elif lsdir_o[0].isupper():
                            folder_name_inner = folder_name_inner.upper()
                        elif lsdir_o[0].islower():
                            folder_name_inner = folder_name_inner.lower()


                        basepath = './' + folder_name_outer + '/' + folder_name_inner
                        print(os.listdir(basepath))
                        for file_name in os.listdir(basepath):

                            q = ''
                            k = [2014, 2015, 2016, 2017]
                            print(q.join(re.findall("[0-9]", file_name)))
                            if k.index(int(q.join(re.findall("[0-9]", file_name)))):

                                if os.path.isfile(os.path.join(basepath, file_name)):
                                    src = basepath + '/' + file_name
                                    dst = './new_folder'

                                    typeZ = glob.iglob(os.path.join(src, "*.zip"))
                                    typeR = glob.iglob(os.path.join(src, "*.rar"))
                                    for t in typeZ:
                                        print(typeZ)
                                        if t == 'zip':
                                            extract = zipfile.ZipFile(src)
                                            src = os.path.splitext(src)[0]
                                            extract.extractall(src)

                                    print('l', src, dst)
                                    print(shutil.copy2(src, dst))
                                    print('o')

                                    print(file_name, folder_name_inner)
                                    s = ""
                                    folder_name_inner = folder_name_inner.replace(' ', '')
                                    folder_name_inner = folder_name_inner.replace('-', '')

                                    os.rename(dst+'/'+file_name, dst+'/' + folder_name_inner + '_' + s.join(re.findall("[0-9]", file_name))+'.pdf')

                    except:
                        print('jello')

                except:
                    print('the folder was not found A')


        # deal with zips
