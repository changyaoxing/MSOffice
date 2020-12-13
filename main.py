from docx import Document
import xlrd
if __name__ == '__main__':
    filename="D:/2020学术年会/2020年研究生学术年会投稿情况汇总表 - 副本.xls";
    wordPath="D:/2020学术年会/中南大学计算机学院第二届研究生学术年会论文摘要集.docx";
    data = xlrd.open_workbook(filename);
    document = Document(wordPath);
    table = data.sheets()[0];
    nrows=table.nrows;
    i=1
    while(i<nrows):
        paperName=table.cell_value(i,7);
        paperAuthor = table.cell_value(i, 8);
        paperKeyWordsCN = table.cell_value(i, 9);
        paperAbstractCN = table.cell_value(i, 10);
        paperKeyWordsENG = table.cell_value(i, 11);
        paperAbstractENG = table.cell_value(i, 12);
        if(paperName!=None):
            document.add_paragraph(paperName,style='nameCN');
        if (paperAuthor != None):
            document.add_paragraph(paperAuthor,style='authorCN');
        if (paperAbstractCN != None):
            document.add_paragraph("摘  要  ", style='abAndKeyCN').add_run(paperAbstractCN);
        if (paperKeyWordsCN != None):
            document.add_paragraph("关键字  ", style='abAndKeyCN').add_run(paperKeyWordsCN);
        if (paperAbstractENG != None):
            try:
                document.add_paragraph("Abstract  ", style='abAndKeyENG').add_run(paperAbstractENG);
            except ValueError as e:
                print(paperName);
                print(paperAbstractENG);
        if (paperKeyWordsENG != None):
            document.add_paragraph("Keywords  ", style='abAndKeyENG').add_run(paperKeyWordsENG);
        document.save(wordPath);# 保存文档
        i+=1;
