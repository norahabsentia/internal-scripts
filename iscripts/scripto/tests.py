import xlrd
from xlsxwriter import Workbook
import time

start = time.time()
loc = 'test_files/Export _  _ Wednesday, October 3, 2018.xlsx'
wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(0)
print('reading file',time.time() - start)
# Extracting number of rows and columns
rows = sheet.nrows
columns = sheet.ncols

dic = {}
for i in range(columns):
    dic[sheet.col_values(i)[0]] = sheet.col_values(i)
print('data conversion to dictionary',time.time() - start)

remap = ['visitType', 'visitIp', 'pagesCount', 'serverDate', 'serverTime', 'visitDurationPretty',
         'vcf downloaded']
for zz in range(51):
    remap.append('type (actionDetails %d)' % zz)
    remap.append('url (actionDetails %d)' % zz)

vcf_li = ['vcf downloaded']
for r in range(rows):
    vcf_check = []
    for c in range(51):
        vcf_check.append(dic['url (actionDetails %d)' % c][r][-4:])
    if '.vcf' in vcf_check:
        vcf_li.append('Yes')
    else:
        vcf_li.append('No')
dic['vcf downloaded'] = vcf_li

visitType = {}
visitType_li = ['visitType']
for j in dic['visitIp']:
    if j != 'visitIp' and j not in visitType.keys():
        visitType[j] = ['Unique']
        visitType_li.append('Unique')
    elif j != 'visitIp':
        visitType[j] += ['Return']
        visitType_li.append('Return')
dic['visitType'] = visitType_li
# seperate server date and time
date_li = ['serverDate']
time_li = ['serverTime']
for l in dic['serverTimePretty (actionDetails 0)']:
    if l != 'serverTimePretty (actionDetails 0)':
        if l != '':
            datetime = l.split()
            date = " ".join(datetime[:3])
            time = datetime[-1]
            date_li.append(date)
            time_li.append(time)
        else:
            date_li.append(l)
            time_li.append(l)
dic['serverDate'] = date_li
dic['serverTime'] = time_li

# counting number of pages visited
count_li = ['pagesCount']
for xx in range(rows - 1):
    c = 0
    for x in range(columns):
        if x < 51:
            if dic['serverTimePretty (actionDetails %d)' % x][xx + 1] != '':
                c += 1
    count_li.append(c)
dic['pagesCount'] = count_li
# print(len(dic['pagesCount']), dic['pagesCount'])
# contains list of all required data
alldata = []
for z in remap:
    if z in dic.keys():
        alldata.append(dic[z])
# for i in alldata:
#     print(i)
end = time.time() - start
print(end)

# print('alldata', time.time() - start)

# """
# Code to write data to output.xls file
# """
# workbook = Workbook('output/script2/output.xlsx', {'strings_to_urls': False})
# worksheet = workbook.add_worksheet()
# co = 0
# for i in alldata:
#     for index, value in enumerate(i):
#         worksheet.write(index, co, value)
#     co += 1
# workbook.close()