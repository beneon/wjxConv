import pandas as pd
from path_definition import xlsfile, xlstitle
import re
import os
from generate_report_with_data import Report
from datetime import date, timedelta, datetime

best_date = datetime(2021,4,29)
due_date = best_date+timedelta(days=7)

df = pd.read_excel(
    xlsfile,
    engine='openpyxl',
    dtype=str
)

reName = re.compile(r'姓名')
reClass = re.compile(r'班级')
reId = re.compile(r'学号')
reShangchuan = re.compile('上传')


# 实现方式：首先找到姓名、班级和学号列，尝试找到文件上传列。中间的就都是答题列了
colname_index_tuple_name = [(i,e) for i,e in enumerate(df.columns) if reName.search(e)]
colname_index_tuple_class = [(i,e) for i,e in enumerate(df.columns) if reClass.search(e)]
colname_index_tuple_id = [(i,e) for i,e in enumerate(df.columns) if reId.search(e)]
colname_index_tuple_upload = [(i,e) for i,e in enumerate(df.columns) if reShangchuan.search(e)]
assert len(colname_index_tuple_name)==1 & len(colname_index_tuple_id)==1 & len(colname_index_tuple_class)==1, 'column for name and id is not found in datafile'

if colname_index_tuple_id[0][0]-colname_index_tuple_name[0][0]!=2:
    print('index diff for id column and name column is not 2')

colname_index_tuple_name = colname_index_tuple_name[0]
colname_index_tuple_id = colname_index_tuple_id[0]
colname_index_tuple_class = colname_index_tuple_class[0]

if len(colname_index_tuple_upload)==0:
    print('no upload file field was found in data file')
    colname_index_tuple_upload = None
elif len(colname_index_tuple_upload)!=1:
    print('more than 1 upload file field was found in data file')
    colname_index_tuple_upload = None
else:
    colname_index_tuple_upload = colname_index_tuple_upload[0]

if colname_index_tuple_upload:
    df_refined = df.iloc[:,colname_index_tuple_name[0]:colname_index_tuple_upload[0]]
else:
    df_refined = df.iloc[:,colname_index_tuple_name[0]:]

info_columns_dict = {
    'name':colname_index_tuple_name[1],
    'class':colname_index_tuple_class[1],
    'id':colname_index_tuple_id[1],
}


def generate_report(data_entry):
    if xlstitle:
        report = Report(xlstitle,data_entry,info_columns_dict)
        report.report_save()
    else:
        raise Exception(f'xls title do not exist for {xlsfile}')


df_refined.apply(lambda r:generate_report(r),axis=1)

df_conv = pd.read_excel(
    os.path.splitext(xlsfile)[0]+'_conv.xlsx',
    engine='openpyxl',
)

df_conv['提交答卷时间'] = pd.to_datetime(df_conv['提交答卷时间'])

def gen_report_link(data_entry):
    file_name = f"{data_entry[info_columns_dict['id']]}_{data_entry[info_columns_dict['name']]}.docx"
    return f'=HYPERLINK("./{file_name}","作业链接")'

df_conv['report_link'] = df_conv.apply(lambda r:gen_report_link(r),axis=1)

# overdue flag
df_conv['提交答卷时间'] = pd.to_datetime(df_conv['提交答卷时间'])
df_conv['overdued'] = df_conv['提交答卷时间']>due_date
df_conv['docx_score']=100
df_conv['upload_file_score']=100
df_conv['total_score']="([@[docx_score]]+[@[upload_file_score]])/2*(1-IF([@overdued],0.2,0))"



def ind2excel_col_num(ind):
    """

    :param ind: number index
    :return: converted to excel col number like 1 for A, max num support for 26*26
    """
    alphabet_list = list('ABCDEFGHIJKLMNOPQRSTUVWXYZ')
    dig1 = ind-1 % 26
    dig2 = (ind - 1 - dig1)/26
    alpha1 = alphabet_list[dig1]
    if dig2==0:
        alpha2=""
    else:
        alpha2=alphabet_list[dig2]

    return f'{alpha2}{alpha1}'



writer = pd.ExcelWriter(os.path.splitext(xlsfile)[0]+'_conv.xlsx', engine='xlsxwriter')
df_conv.to_excel(writer,'Sheet1',index=False)
writer.save()

