import os
import shutil
import zipfile
import pandas as pd
import urllib.parse as ups
import snoop
import re

# 问卷数据下载后是xls文件, 附件压缩包是zip文件, 两者和代码放在同一个文件夹中
# 列出当前文件夹下面所有的文件
os.chdir('data')
filelist = os.listdir()
# 下面分别是获取xls文件和zip文件的文件名
[xlsfile] = [e for e in filelist if os.path.splitext(e)[1] == '.xlsx']
[zipfilepath] = [e for e in filelist if os.path.splitext(e)[1] == '.zip']

# 先将zip包解压缩
# 首先使用zipfile库读取附件文件, 建立ZipFile对象
zip_ref = zipfile.ZipFile(zipfilepath, 'r')
# 为了避免命名出现错误,这里手动进行文件名的转码
targetEncode = 'gb2312'
nl = zip_ref.namelist()
# 解压缩以前先新建文件夹, 文件夹名称和压缩包的文件名相同
unzip_dir = os.path.splitext(zipfilepath)[0]
if os.path.exists(unzip_dir):
    print("好像已经解压缩过了哦?")
else:
    os.mkdir(unzip_dir)
    # 对于一些特殊文件夹在这里先过滤掉
    nl = [e for e in nl if '__' not in e]
    # 手动遍历所有文件,然后对文件名进行转码
    for fn in nl:
        correct_fn = fn.encode('cp437').decode(targetEncode)
        correct_fn = os.path.join(unzip_dir, correct_fn)
        with open(correct_fn, 'wb') as output_file:
            with zip_ref.open(fn, 'r') as original_file:
                shutil.copyfileobj(original_file, output_file)

# 然后是读取Excel文件
# 声明filename extract函数
reIdNumberName = re.compile(r"(\d+)_(\d+)_(.*)")

def fileNameExtractor(row,url_header:str,name_header:str):
    # get the query string, like: path=http%3a%2f%2fpubuserqiniu.paperol.cn%2f37299324_1_q4_pDJ5pCCKECzTxAdqQLyuw.docx%3fattname%3d1_4_%25e5%25ae%259e%25e9%25aa%258c%25e6%258a%25a5%25e5%2591%258a.docx&activity=37299324
    urlStr = row[url_header]
    query1 = ups.urlparse(urlStr).query
    # get the path string, like: http://pubuserqiniu.paperol.cn/37299324_1_q4_pDJ5pCCKECzTxAdqQLyuw.docx?attname=1_4_%e5%ae%9e%e9%aa%8c%e6%8a%a5%e5%91%8a.docx
    activityStr = ups.parse_qs(query1)['activity'][0]
    pathStr = ups.parse_qs(query1)['path'][0]
    # parse the path string again to get query string
    query2 = ups.urlparse(pathStr).query
    # now the query string is like attname=1_4_%e5%ae%9e%e9%aa%8c%e6%8a%a5%e5%91%8a.docx
    # parse_qs this string to get the encoded string of file
    attname = ups.parse_qs(query2)['attname'][0]
    snoop.pp(attname)
    mo = reIdNumberName.match(attname)
    if mo is None:
        print(f"{attname} 格式不对")
    conv_filename = f"序号{mo.group(1)}_{row[name_header]}_{mo.group(2)}_{mo.group(3)}"
    # 最后是从url中提取了问卷的activity Id和对应的上传文件文件名, 然后按照Excel中hyperlink公式的格式形成公式
    # rst = '=HYPERLINK("./{}_附件/{}","作业链接")'.format(activityStr, docxFileName)

    return f'=HYPERLINK("./{activityStr}_附件/{conv_filename}","作业链接")'

xl = pd.ExcelFile(xlsfile)
df_xl = xl.parse('Sheet1', dtype=str)
# 目前的问卷自动带有题号,题号后面会紧跟着文件名
reXingming = re.compile('姓名')
reShangchuan = re.compile('上传')
columns = list(df_xl.columns)
name_header = [e for e in columns if reXingming.search(e) is not None]
assert len(name_header)>0, "未发现姓名字样"
assert len(name_header)==1, "多重项包含姓名字样"
name_header = name_header[0]
upload_headers = [e for e in columns if reShangchuan.search(e) is not None]
link_headers = [e.replace('上传','链接:') for e in upload_headers]
assert len(upload_headers)>0, "未发现上传字样"
reUploadNumber = re.compile(r'(\d+)[、,\s]')
for i,u in enumerate(upload_headers):
    df_xl[link_headers[i]] = df_xl.apply(lambda r:fileNameExtractor(r,u,name_header),axis=1)

# '序号','提交答卷时间','所用时间', 0:3
# 6:到结束, 然后drop掉原来的url列
df_rst = df_xl[columns[0:3]+columns[6:]+link_headers]
df_rst = df_rst.drop(columns=upload_headers)

# 设置输出文件的名称
output_xls = os.path.splitext(xlsfile)[0] + "_conv.xlsx"
# 建立excel writer对象
writer = pd.ExcelWriter(output_xls, engine='xlsxwriter')
# 输出
df_rst.to_excel(writer, 'Sheet1')
# 保存输出excel文件
writer.save()


