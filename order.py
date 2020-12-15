import xlrd
import os
from docx import Document
import xlwt
import re
if __name__ == '__main__':
    orderExcelPath="D:/2020学术年会/order.xls";
    fromExcelPaperNamePath = "D:/2020学术年会/paperNameFromExcel.xls";
    paperNameExcelPath = "D:/2020学术年会/paperName.xls";

    paperNameExcel = xlrd.open_workbook(paperNameExcelPath);
    fromExcelPaperName = xlrd.open_workbook(fromExcelPaperNamePath);
    # 选择Excel中的第一张表
    paperNameTable = paperNameExcel.sheets()[0];
    fromExcelPaperNameTable = fromExcelPaperName.sheets()[0];
    # 获得有效行数

    paperNameDic={};
    paperAuthorDic={};
    fromExcelPaperNameList=[];

    nrows1 = paperNameTable.nrows;
    i = 0
    while (i < nrows1):
        # 获得单元格内容
        nameCN=paperNameTable.cell_value(i, 1);
        nameENG=paperNameTable.cell_value(i, 2);
        author=paperNameTable.cell_value(i, 3);
        paperNameDic[nameCN]=nameENG;
        paperAuthorDic[nameCN]=author;
        i+=1;

    nrows2 = fromExcelPaperNameTable.nrows;
    i = 0
    while (i < nrows2):
        # 获得单元格内容
        name = fromExcelPaperNameTable.cell_value(i, 0);
        fromExcelPaperNameList.append(name);
        i += 1;

    authorDic={};
    resultDic={};
    tempDic={};
    for x in fromExcelPaperNameList:
        authorDic[x]="";
        resultDic[x]="";
        tempDic[x]="";
        for y,z in paperNameDic.items():
            if x in y:
                authorDic[x]=paperAuthorDic[y];
                resultDic[x]=z;
                tempDic[x]=y;
                break;

    workbook = xlwt.Workbook();
    worksheet = workbook.add_sheet('My Sheet')

    nrows2 = fromExcelPaperNameTable.nrows;
    i = 0
    while (i < nrows2):
        # 获得单元格内容
        nameCN = fromExcelPaperNameTable.cell_value(i, 0);
        authorCN = fromExcelPaperNameTable.cell_value(i, 1);
        KeyCN = fromExcelPaperNameTable.cell_value(i, 2);
        AbCN = fromExcelPaperNameTable.cell_value(i, 3);
        KeyENG = fromExcelPaperNameTable.cell_value(i, 4);
        AbENG = fromExcelPaperNameTable.cell_value(i, 5);

        nameENG=resultDic[nameCN];
        authorENG=authorDic[nameCN];
        temp=tempDic[nameCN];


        worksheet.write(i, 0, nameCN);
        worksheet.write(i, 1, temp);
        worksheet.write(i, 2, authorCN);
        worksheet.write(i, 3, nameENG);
        worksheet.write(i, 4, authorENG);
        worksheet.write(i, 5, KeyCN);
        worksheet.write(i, 6, AbCN);
        worksheet.write(i, 7, KeyENG);
        worksheet.write(i, 8, AbENG);

        workbook.save(orderExcelPath);
        i += 1;

