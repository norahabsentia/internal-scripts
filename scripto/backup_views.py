from django.http import HttpResponseRedirect, HttpResponse, Http404
from django.shortcuts import render
from django.conf import settings
from .forms import UploadFileForm
from django.core.files.storage import FileSystemStorage
import pandas as pd
import xlrd
import os
import glob
from xlwt import Workbook
from xlsxwriter import Workbook
from zipfile import ZipFile
import shutil

def home(request):
    return render(request, 'home.html')

def zipit(dpath):
    shutil.make_archive(os.path.join(os.path.dirname(__file__),"../output/script4/Output"), 'zip', dpath)
# script 1 - Piwik Analytics / App Preview / ReportingScript1
def script1(request):
    try:
        if request.method == 'POST':
            form = UploadFileForm(request.POST, request.FILES)
            if form.is_valid():
                input_file = request.FILES['file']
                fs = FileSystemStorage()
                filename = fs.save(input_file.name, input_file)
                uploaded_file_url = fs.path(filename)
                # Code to read data from input excel file
                loc = ("%s" % uploaded_file_url)
                # # To open Workbook
                wb = xlrd.open_workbook(loc)
                sheet = wb.sheet_by_index(0)
                # Extracting number of rows and columns
                rows = sheet.nrows
                columns = sheet.ncols

                input_format = []
                for i in range(sheet.ncols):
                    input_format.append(sheet.cell_value(0, i))

                remap = ['visitType', 'visitIp', 'pagesCount', 'serverDate', 'serverTime', 'visitDurationPretty']
                for zinga in range(51):
                    remap.append('pageTitle (actionDetails %d)' % zinga)

                dic = {}
                for i in range(columns):
                    dic[sheet.col_values(i)[0]] = sheet.col_values(i)

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
                        date = l[:-8]
                        time = l[-8:]
                        date_li.append(date)
                        time_li.append(time)
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
                # contains list of all required data
                alldata = []
                for z in remap:
                    if z in dic.keys():
                        alldata.append(dic[z])
                """
                Code to write data to output.xls file
                """
                workbook = Workbook('output/script1/output.xlsx', {'strings_to_urls': False})
                worksheet = workbook.add_worksheet()
                co = 0
                for i in alldata:
                    for index, value in enumerate(i):
                        worksheet.write(index, co, value)
                    co += 1
                workbook.close()
                fs.delete(filename)
                dout = os.path.join(os.path.dirname(__file__), "../output/script1/output.xlsx")
                down = download(request, dout)
                return down
            return render(request, 'scripto/fail.html')
        else:
            form = UploadFileForm()
        return render(request, 'scripto/upload/uploadFile.html', {'form': form})
    except:
        raise Http404

def script2(request):
    try:
        if request.method == 'POST':
            form = UploadFileForm(request.POST, request.FILES)
            if form.is_valid():
                input_file = request.FILES['file']
                fs = FileSystemStorage()
                filename = fs.save(input_file.name, input_file)
                uploaded_file_url = fs.path(filename)
                loc = uploaded_file_url

                df = pd.read_excel(loc, index_col=None)
                a = list(df.columns)

                visitType = []
                visitTyped = {}
                for i in df['visitIp']:
                    if i not in visitTyped.keys():
                        visitTyped[i] = 1
                        visitType.append('Unique')
                    else:
                        visitType.append('Return')
                df['visitType'] = visitType

                all_urls = [_a for _a in a if 'url (actionDetails' in _a]

                pagesCount = []
                vcf = []
                for _r in range(df.shape[0]):
                    c = 0
                    count = 0
                    for _j in all_urls:
                        if type(df[_j][_r]) != float and '.vcf' in df[_j][_r]:
                            c = 1
                        if type(df[_j][_r]) != float:
                            count += 1
                    if c == 0:
                        vcf.append('No')
                    else:
                        vcf.append('Yes')
                    pagesCount.append(count)
                df['pagesCount'] = pagesCount
                df['vcf downloaded'] = vcf

                remap = ['visitType', 'visitIp', 'pagesCount', 'serverDatePretty', 'serverTimePretty',
                         'visitDurationPretty', 'vcf downloaded']
                remap += all_urls
                df.to_excel('output/script2/output.xlsx', index=False, columns=remap)




                wb = xlrd.open_workbook(loc)
                sheet = wb.sheet_by_index(0)
                # Extracting number of rows and columns
                rows = sheet.nrows
                columns = sheet.ncols

                dic = {}
                for i in range(columns):
                    dic[sheet.col_values(i)[0]] = sheet.col_values(i)

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
                print(len(dic['pagesCount']), dic['pagesCount'])
                # contains list of all required data
                alldata = []
                for z in remap:
                    if z in dic.keys():
                        alldata.append(dic[z])
                for i in alldata:
                    print(i)

                """
                Code to write data to output.xls file
                """
                workbook = Workbook('output/script2/output.xlsx', {'strings_to_urls': False})
                worksheet = workbook.add_worksheet()
                co = 0
                for i in alldata:
                    for index, value in enumerate(i):
                        worksheet.write(index, co, value)
                    co += 1
                workbook.close()


                fs.delete(filename)
                dout = os.path.join(os.path.dirname(__file__), "../output/script2/output.xlsx")
                down = download(request, dout)
                return down
        else:
            form = UploadFileForm()
        return render(request, 'scripto/upload/uploadFile.html', {'form': form})
    except:
        raise Http404

# Script 3 - FIITJEE Demo / Norah Analytics / Notifications Demo / XlsxToJson
def script3(request):
    try:
        if request.method == 'POST':
            form = UploadFileForm(request.POST, request.FILES)
            if form.is_valid():
                input_file = request.FILES['file']
                fs = FileSystemStorage()
                filename = fs.save(input_file.name, input_file)
                uploaded_file_url = fs.path(filename)
                loc = uploaded_file_url
                # To open Workbook
                wb = xlrd.open_workbook(loc)
                sheet = wb.sheet_by_index(0)
                # Extracting number of rows and columns
                rows = sheet.nrows
                columns = sheet.ncols

                input_format = []
                for i in range(sheet.ncols):
                    input_format.append(sheet.cell_value(0, i))

                dic = {}
                for i in range(1,rows):
                    z = sheet.row_values(i)
                    if z[0] not in dic.keys():
                        dic[z[0]] = { "notification":[{"text":"", "title":"", "data":[]}]}
                        if dic[z[0]]["notification"][-1]["text"] != z[2]:
                            dic[z[0]]["notification"][-1]["text"] = z[2]
                            dic[z[0]]["notification"][-1]["title"] = z[1]
                            dic[z[0]]["notification"][-1]["data"].append({"id":z[4], "text":z[3]})
                        elif dic[z[0]]["notification"][-1]["text"] == z[2]:
                            dic[z[0]]["notification"][-1]["data"].append({"id":z[4], "text":z[3]})
                    else:
                        if dic[z[0]]["notification"][-1]["text"] != z[2]:
                            dic[z[0]]["notification"].append({"text":z[2], "title":z[1], "data":[{"id":z[4], "text":z[3]}]})
                        elif dic[z[0]]["notification"][-1]["text"] == z[2]:
                            dic[z[0]]["notification"][-1]["data"].append({"id":z[4], "text":z[3]})
                finaldata = {"profiles":dic}
                """
                writing data to json or text file
                """
                import json
                # Serialize the list of dicts to JSON
                j = json.dumps(finaldata, indent=4)
                # Write to file
                with open('output/script3/output.json', 'w') as f:
                    f.write(j)

                fs.delete(filename)
                dout = os.path.join(os.path.dirname(__file__), "../output/script3/output.json")
                down = download(request, dout)
                return down
            return render(request, 'scripto/fail.html')
        else:
            form = UploadFileForm()
        return render(request, 'scripto/upload/uploadFile.html', {'form': form})
    except:
        raise Http404
# Script 4 - FIITJEE Demo / Norah Analytics / Reporting Dashboard / XlsxToJson
def script4(request):
    try:
        if request.method == 'POST':
            form = UploadFileForm(request.POST, request.FILES)
            if form.is_valid():
                input_file = request.FILES['file']
                fs = FileSystemStorage()
                filename = fs.save(input_file.name, input_file)
                uploaded_file_url = fs.path(filename)
                print(input_file.name)

                zip_ref = ZipFile(uploaded_file_url, 'r')
                zip_ref.extractall(os.path.join(os.path.dirname(__file__),"../output/script4"))
                zip_ref.close()
                fs.delete(filename)

                for filename in glob.iglob(os.path.join(os.path.dirname(__file__),"../output/script4/*/**/*.xlsx"), recursive=True):
                    loc = ("%s" % filename)
                    wb = xlrd.open_workbook(loc)
                    sheet = wb.sheet_by_index(0)

                    rows = sheet.nrows
                    columns = sheet.ncols

                    input_format = []
                    input_format2 = []
                    for i in range(sheet.ncols):
                        input_format.append(sheet.cell_value(0, i))
                        input_format2.append(sheet.cell_value(1, i))

                    dic = {}
                    if input_format[0] == input_format2[0]:
                        for i in range(1, rows):
                            z = sheet.row_values(i)
                            dic2 = {}
                            for j in range(1, columns):
                                dic2[input_format[j]] = z[j]
                            if z[0] not in dic.keys():
                                dic[z[0]] = [dic2]
                            else:
                                dic[z[0]].append(dic2)
                    else:
                        for i in range(1, rows):
                            z = sheet.row_values(i)
                            for j in range(columns):
                                dic[input_format[j]] = z[j]
                    """
                    writing data to json or text file
                    """
                    import json
                    zamura = filename[:-5]
                    j = json.dumps(dic, indent=4)
                    with open('%s.json' % zamura, 'w') as f:
                        f.write(j)
                din = os.path.join(os.path.dirname(__file__),"../output/script4/%s" % input_file.name[:-4])
                shutil.make_archive(os.path.join(os.path.dirname(__file__), "../output/script4/output"), 'zip', din)
                shutil.rmtree(os.path.join(os.path.dirname(__file__),"../output/script4/%s" % input_file.name[:-4]))
                dout = os.path.join(os.path.dirname(__file__), "../output/script4/output.zip")
                down = download(request, dout)
                return down
            return render(request, 'scripto/fail.html')
        else:
            form = UploadFileForm()
        return render(request, 'scripto/upload/uploadFile.html', {'form': form})
    except:
        raise Http404

# script 5 - Pictor Live Campaigns / HDFC / PL BL Durga Puja / DataPre-processing
def script5(request):
    try:
        if request.method == 'POST':
            form = UploadFileForm(request.POST, request.FILES)
            if form.is_valid():
                input_file = request.FILES['file']
                flink1 = request.POST['link1']
                fvideo = request.POST['video']
                fvideoLink = request.POST['videoLink']
                fendposter = request.POST['endposter']
                fs = FileSystemStorage()
                filename = fs.save(input_file.name, input_file)
                uploaded_file_url = fs.path(filename)
                # Code to read data from input excel file
                loc = ("%s"% uploaded_file_url)
                wb = xlrd.open_workbook(loc)
                sheet = wb.sheet_by_index(0)
                rows = sheet.nrows
                columns = sheet.ncols

                input_format = []
                for i in range(sheet.ncols):
                    input_format.append(sheet.cell_value(0, i))

                dic = {}
                for i in range(columns):
                    dic[sheet.col_values(i)[0]] = sheet.col_values(i, 1)

                required_data = {}
                name_li = ['name']
                loan_li = ['loan']
                link_li = ['link']
                link1_li = ['link1']
                video_li = ['video']
                videoLink_li = ['videoLink']
                endposter_li = ['endposter']

                for row in range(rows - 1):
                    first_name = []
                    for alpha in dic['Customer Name'][row].split():
                        if len(alpha) <= 2:
                            first_name.append(alpha)
                        else:
                            first_name.append(alpha)
                            break
                    if len(first_name) == 1:
                        name_li.append(first_name[0])
                    else:
                        name_li.append(' '.join(first_name))
                    rupee = dic['Pre approved amount'][row]
                    loan_li.append(rupee)
                    link = dic['CTA Link'][row]
                    link_li.append(link)
                    link1_li.append(flink1)
                    video_li.append(fvideo)
                    videoLink_li.append(fvideoLink)
                    endposter_li.append(fendposter)

                required_data['name'] = name_li
                required_data['loan'] = loan_li
                required_data['link'] = link_li
                required_data['link1'] = link1_li
                required_data['video'] = video_li
                required_data['videoLink'] = videoLink_li
                required_data['endposter'] = endposter_li

                map = ['name', 'loan', 'link', 'link1', 'video', 'videoLink', 'endposter']
                # contains list of all required data
                alldata = []
                for z in map:
                    if z in required_data.keys():
                        alldata.append(required_data[z])
                # Code to write data to output.xlsx file
                o_filename = filename[:-5] + '_db_import.xlsx'
                dpath = os.path.join(os.path.dirname(__file__), "../output/script5/%s"% o_filename)
                workbook = Workbook(dpath, {'strings_to_urls': False})
                worksheet = workbook.add_worksheet()
                co = 0
                for i in alldata:
                    for index, value in enumerate(i):
                        worksheet.write(index, co, value)
                    co += 1
                workbook.close()
                fs.delete(filename)
                down = download(request, dpath)
                return down
            return render(request, 'scripto/fail.html')
        else:
            form = UploadFileForm()
        return render(request, 'scripto/upload/script5uploadFile.html', {'form': form})
    except:
        raise Http404

def download(request, path):
    path = path
    file_path = os.path.join(settings.MEDIA_ROOT, path)
    if os.path.exists(file_path):
        with open(file_path, 'rb') as fh:
            response = HttpResponse(fh.read(), content_type="application/vnd.ms-excel")
            response['Content-Disposition'] = 'inline; filename=' + os.path.basename(file_path)
            return response
    raise Http404
