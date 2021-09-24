import os
import re
# 问卷数据下载后是xls文件, 附件压缩包是zip文件, 两者和代码放在同一个文件夹中
# 转换工作文件夹，列出当前文件夹下面所有的文件
cwd = r"""
D:\SynologyDrive\护理学院工作\教学\spss软件实验\2021备课\exp07\hw
"""
cwd = cwd.strip()
os.chdir(cwd)
filelist = os.listdir()
# 下面分别是获取xls文件和zip文件的文件名
xlsfiles = [e for e in filelist if (os.path.splitext(e)[1] == '.xlsx') & ('_conv' not in os.path.splitext(e)[0])]
zipfiles = [e for e in filelist if os.path.splitext(e)[1] == '.zip']
assert len(xlsfiles)==1, f'more than 1 xlsx file identified in path {cwd}'
xlsfile=xlsfiles[0]
#根据文件名找到问卷标题，目前形式是 108786128_2_2020-2021(2)spss软件应用训练实验1_182_182
reXlsTitle = re.compile(r'\d*_\d_(.*)_\d*_\d*')
mo = reXlsTitle.match(os.path.splitext(xlsfile)[0])
if mo:
    xlstitle = mo.group(1)
else:
    xlstitle = None

if len(zipfiles)!=1:
    print(f'{len(zipfiles)} zip file found in path {cwd}, check the folder before continue')
    zipfilepath = None
else:
    zipfilepath = zipfiles[0]