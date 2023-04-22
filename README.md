# excelTable_to_Latex
help transform the excel table to a type of Latex table

========================================================

v1.0
1. Data in the same format as' text.xsl 'is allowed to be input, that is, sheet is named with the name of data set, each row is the name of algorithm, and each column is the name of evaluation index
2. The comparison algorithm can be inserted at will
3. The evaluation indexes of the current version can only be exactly 7, which will be improved in subsequent versions
4, generate table form in Latex, whether the table needs to import package is not sure, the author is checking.
5. The original intention of this project is to realize the rapid conversion of excel tables to Latex tables when the results of different algorithms are compared
Updated at 2023.04.19

========================================================

v1.1
1. Added additional variance data tabulation based on v1.0
2. The program structure is standardized
3. Disadvantages: Still only the mean and its corresponding variance of exactly seven indicators can be processed
4. Latex table does not require import packages
Updated at 2023.04.20

=========================================================

v1.2
1. Joined tkinter library and designed a simple GUI interface.
2. Try the pyinstaller library and package it as a main.exe file. Test the file can run but open slowly, the file is large.
Upadated at 2023.04.22
