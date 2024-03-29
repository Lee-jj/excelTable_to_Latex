# excelTable_to_Latex
help transform the excel table to a type of Latex table

命令行执行：``python excelTable_to_Latex.py``，选择*template.xlsx*三个sheet其中之一，点击**选择**按钮来选择待输出指标，点击**生成Latex代码**按钮来生成Latex格式的表格文本；

========================================================

v1.0
1. Data in the same format as' text.xsl 'is allowed to be input, that is, sheet is named with the name of data set, each row is the name of algorithm, and each column is the name of evaluation index
2. The comparison algorithm can be inserted at will
3. The evaluation indexes of the current version can only be exactly 7, which will be improved in subsequent versions
4, generate table form in Latex, whether the table needs to import package is not sure, the author is checking.
5. The original intention of this project is to realize the rapid conversion of excel tables to Latex tables when the results of different algorithms are compared

[Updated at 2023.04.19]

========================================================

v1.1
1. Added additional variance data tabulation based on v1.0
2. The program structure is standardized
3. Disadvantages: Still only the mean and its corresponding variance of exactly seven indicators can be processed
4. Latex table does not require import packages

[Updated at 2023.04.20]

=========================================================

v1.2
1. Joined tkinter library and designed a simple GUI interface.
2. Try the pyinstaller library and package it as a main.exe file. Test the file can run but open slowly, the file is large.(exe file is too large to upload.)

[Upadated at 2023.04.22]

=========================================================

v1.3
1. 修改表格样式

[Updated at 2024.01.24]

=========================================================

v2.0
1. 引入下拉复选框，读取文档中的指标，并选择输出指标，参考：[Python tkinter自定义多选下拉列表框(带滚动条、全选)](https://blog.csdn.net/darren922/article/details/132985878)

[Updated at 2024.01.25]

=========================================================

v2.1
1. 想要一次输出多个数据库的结果，但是失败了
2. **选择**按钮有bug，还未解决

[Updated at 2024.01.26]
