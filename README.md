# MSOffice
背景：
Excel中有论文的中文名，作者中文名，中英文摘要和关键字；
文件夹中有所有的论文Word文件，文件名包涵论文名，作者，作者信息；
目的：
通过这两种文件获得所有论文的摘要集。

nameSearch：
扫描文件夹中的全部word文件获得对应的论文英文名和作者英文名，并写入Excel文件；

order：
根据Excel论文的中文名匹配nameSearch获得论文中文名获得论文英文名和作者英文名；
将论文的中文名、英文名、中英文作者名、中英文摘要和关键字写入新的Excel文件；

write2Word：
根据order获得的新Excel文件写出所有论文的摘要集；

