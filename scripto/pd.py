import pandas as pd
import time
import glob
import os
import json

# uploaded_file_url = fs.path(filename)
# print(input_file.name)

# zip_ref = ZipFile(uploaded_file_url, 'r')
# zip_ref.extractall(os.path.join(os.path.dirname(__file__), "../output/script4"))
# zip_ref.close()
# fs.delete(filename)

for filename in glob.iglob(os.path.join(os.path.dirname(__file__), "test_files/D2/**/*.xlsx"), recursive=True):
    loc = ("%s" % filename)
    # print(loc)
    df = pd.read_excel(loc)
    all_columns = list(df.columns)
    print(all_columns)
    print(df)
    dic = {}

    # if df[all_columns[0]][0] == df[all_columns[0]][1]:


    j = json.dumps(df, indent=4)
    with open('s.json', 'w') as f:
        f.write(j)

    break
    # input_format = []
    # input_format2 = []
    # for i in range(sheet.ncols):
    #     input_format.append(sheet.cell_value(0, i))
    #     input_format2.append(sheet.cell_value(1, i))
    #
    # dic = {}
    # if input_format[0] == input_format2[0]:
    #     for i in range(1, rows):
    #         z = sheet.row_values(i)
    #         dic2 = {}
    #         for j in range(1, columns):
    #             dic2[input_format[j]] = z[j]
    #         if z[0] not in dic.keys():
    #             dic[z[0]] = [dic2]
    #         else:
    #             dic[z[0]].append(dic2)
    # else:
    #     for i in range(1, rows):
    #         z = sheet.row_values(i)
    #         for j in range(columns):
    #             dic[input_format[j]] = z[j]
#     """
#     writing data to json or text file
#     """
#     import json
#
#     zamura = filename[:-5]
#     j = json.dumps(dic, indent=4)
#     with open('%s.json' % zamura, 'w') as f:
#         f.write(j)
# din = os.path.join(os.path.dirname(__file__), "../output/script4/%s" % input_file.name[:-4])
# shutil.make_archive(os.path.join(os.path.dirname(__file__), "../output/script4/output"), 'zip', din)
# shutil.rmtree(os.path.join(os.path.dirname(__file__), "../output/script4/%s" % input_file.name[:-4]))
# dout = os.path.join(os.path.dirname(__file__), "../output/script4/output.zip")
# down = download(request, dout)


