# 基于在线问卷的随堂作业收集系统

## 2024-07-24

**本代码库早就已经停止更新了，最近为了用问卷星数据制作实验报告，重新写了一次代码，现在的代码库在[wjx_conv_v2](https://github.com/beneon/wjx_conv_v2)**

佛山科学技术学院口腔医学院口腔系制作，[相关研究](http://med.wanfangdata.com.cn/Paper/Detail/PeriodicalPaper_zhkqyxyjzz202005008)

## 2020-02-25更新

最近发现问卷星的附件文件名格式变了，现在是序号{id}_姓名_题号_文件名的形式。这次根据新的形式更新了一下代码，现在代码放在conv.py里面

不过说实话，既然文件的序号和df的序号是一一对应的，其实直接在两者之间建立关联可能会更好一点吧。

## 2020-07-02更新

利用序号在zip解压缩文件夹中查找对应的文件，避免链接出错

## 2020-11-28准备软件著作权申报

TODO:
1. 将repo转为private(已经提交了软著申请，所以现在又改回来了）
2. 利用selenium登录问卷星账号（实用性不强，算了）
3. 实现下载附件，下载数据文件（同上）

## 2021-05-20 添加了问卷内容导出为多个docx文件的功能

现在的运行顺序是先运行conv.py解压缩zip文件，建立_conv.xlsx表格
然后是运行answer_sheet2word_doc.py，将各个问题和回答形成docx文档，并且在_conv.xlsx中做链接，同时增加计分项（word文档分和上传文件分），并且做一个总分计算公式
这里还考虑到延迟交作业的问题，所以设置了一个截止时间（开始提交以后的7天），如果超过截止时间设置一个分数乘0.8的惩罚
此外，定义问卷数据和zip所在文档的代码单独移动到了path_definition.py里面
生成word文档的代码放到了generate_report_with_data.py里面

顺带一提，xlrd现在不支持读取xlsx文件了，[我今天才发现](https://stackoverflow.com/questions/65254535/xlrd-biffh-xlrderror-excel-xlsx-file-not-supported) ，就很无语。现在这个repo的代码已经把读取excel的库从xlrd替换成了openpyxl，但是之前有太多涉及读取excel的代码都是用的xlrd这个库……我现在终于对代码维护这个事情有了一点点实质性的体验了……
