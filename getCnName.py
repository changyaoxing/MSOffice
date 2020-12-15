from docx import Document
import xlrd
import re
import xlwt
if __name__ == '__main__':
    #Excel文件路径
    readFileName="D:/2020学术年会/2020年研究生学术年会投稿情况汇总表 - 副本.xls";
    writeFileName="D:/2020学术年会/paperName2.xls";
    data = xlrd.open_workbook(readFileName);
    workbook = xlwt.Workbook();
    worksheet = workbook.add_sheet('My Sheet2')
    #选择Excel中的第一张表
    table = data.sheets()[0];
    #获得有效行数
    nrows=table.nrows;
    i=1
    while(i<nrows):
        #获得单元格内容
        paperName=table.cell_value(i,7);
        worksheet.write(i, 0, paperName)
        paperName=re.sub("^.*\+.*\+.*\+.*\+", "", paperName, count=0, flags=0);
        paperName = re.sub("^.*-.*-.*-.*-", "", paperName, count=0, flags=0);
        paperName = re.sub("^.*_.*_.*_.*_", "", paperName, count=0, flags=0);
        worksheet.write(i, 1, paperName)
        workbook.save(writeFileName)
        i+=1;