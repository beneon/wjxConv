import os
import shutil
import zipfile
import pandas as pd
import urllib.parse as ups
# 问卷数据下载后是xls文件, 附件压缩包是zip文件, 两者和代码放在同一个文件夹中
# 列出当前文件夹下面所有的文件
filelist = os.listdir()
# 下面分别是获取xls文件和zip文件的文件名
[xlsfile] = [e for e in filelist if os.path.splitext(e)[1]=='.xls']
[zipfilepath] = [e for e in filelist if os.path.splitext(e)[1]=='.zip']

# 读取xls文件, 3步: 1. 建立ExcelFile, 2. parse文件, 传入参数是工作表的名字, 默认就是'Sheet1', parse后就形成dataframe
# 3. 从里面选择包含附件url的一列. column名字就是文件上传题目的题干. 在本示例中就是"4、请上传作业文档"
xl = pd.ExcelFile(xlsfile)
df_xl = xl.parse('Sheet1')
submitTime = df_xl.columns[1]
questionaireData = df_xl.columns[6:-1]
fileDataHeader = df_xl.columns[-1]
print("""
提取的问卷信息包括:{}, 上传文件的题目题干为:{}
""".format(", ".join(questionaireData),fileDataHeader))
input("确认无误请按回车")
# 从url中提取文件名的操作有些繁琐, 写成一段函数在后面引用
def fileNameExtractor(urlStr):
    # get the query string, like: path=http%3a%2f%2fpubuserqiniu.paperol.cn%2f37299324_1_q4_pDJ5pCCKECzTxAdqQLyuw.docx%3fattname%3d1_4_%25e5%25ae%259e%25e9%25aa%258c%25e6%258a%25a5%25e5%2591%258a.docx&activity=37299324
    query1 = ups.urlparse(urlStr).query
    # get the path string, like: http://pubuserqiniu.paperol.cn/37299324_1_q4_pDJ5pCCKECzTxAdqQLyuw.docx?attname=1_4_%e5%ae%9e%e9%aa%8c%e6%8a%a5%e5%91%8a.docx
    activityStr = ups.parse_qs(query1)['activity'][0]
    pathStr = ups.parse_qs(query1)['path'][0]
    # parse the path string again to get query string
    query2 = ups.urlparse(pathStr).query
    # now the query string is like attname=1_4_%e5%ae%9e%e9%aa%8c%e6%8a%a5%e5%91%8a.docx
    # parse_qs this string to get the encoded string of file
    docxFileName = ups.parse_qs(query2)['attname'][0]
    # 最后是从url中提取了问卷的activity Id和对应的上传文件文件名, 然后按照Excel中hyperlink公式的格式形成公式
    rst = '=HYPERLINK("./{}_附件/{}","作业链接")'.format(activityStr,docxFileName)
    
    return rst

# 将刚才提取信息的函数应用在df_xl的'4、请上传作业文档'列

docxLinkSeries = df_xl[fileHeader].apply(fileNameExtractor)
# 输出excel文件只保留关键的几个列,包括姓名,班级,学号,提交答案时间,以及刚才提取出来的附件文件链接列
dfOut = pd.DataFrame()
dfOut['timePosted'] = df_xl[submitTime]
for e in questionaireData:
    dfOut[e] = df_xl[e]
dfOut['文件链接'] = docxLinkSeries
# 设置输出文件的名称
output_xls = os.path.splitext(xlsfile)[0]+"_conv.xlsx"
# 建立excel writer对象
writer = pd.ExcelWriter(output_xls,engine = 'xlsxwriter')
# 输出
dfOut.to_excel(writer,'Sheet1')
# 保存输出excel文件
writer.save()

# -------
# 接下来是利用python将附件zip文件解压缩. 这个步骤虽然可以手动完成, 但是为了易用性也做进本脚本里面了.


# 首先使用zipfile库读取附件文件, 建立ZipFile对象
zip_ref = zipfile.ZipFile(zipfilepath,'r')
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
        correct_fn = os.path.join(unzip_dir,correct_fn)
        with open(correct_fn,'wb') as output_file:
            with zip_ref.open(fn,'r') as original_file:
                shutil.copyfileobj(original_file,output_file)
