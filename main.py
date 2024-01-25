import pandas as pd
import os
import tkinter as tk
from tkinter import *
import tkinter.ttk
from tkinter import filedialog
from ComBoPicker import Combopicker# 导入自定义下拉多选框


def processTable(datasetName, excelPath, savePath, indicesList):
    datasetName = datasetName
    savePath = savePath
    data = pd.read_excel(excelPath, sheet_name=datasetName, header=1) # 0 or 1

    delList = []
    hasStd = False
    for colName in list(data.columns):
        if '_std' in colName:
            hasStd = True
            if colName[:-4] not in indicesList:
                delList.append(colName)
        else:
            if colName not in indicesList:
                delList.append(colName)

    data = data.drop(delList, axis=1)
    # print(data.drop(delList, axis=1))

    # 记录每个指标最大值与次大值
    METICS_LEN = len(data.columns) // 2 if hasStd else len(data.columns)   # 指标数
    best = [0] * METICS_LEN
    secondBest = [0] * METICS_LEN
    for idx, matric in enumerate(list(data.columns)[:len(best)]):
        temp = []
        for value in list(data[matric]):
            if pd.isnull(value) or type(value) == type('\\'):
                pass
            else:
                temp.append(value)
        temp.sort(reverse=True)
        best[idx] = temp[0] if len(temp) > 0 else 0
        secondBest[idx] = temp[1] if len(temp) > 1 else 0
    
    # 写Latex
    f = open(savePath, 'w')
    f.writelines('\\begin{table*}[t]\n')
    f.writelines('\\caption{Clustering performance on ' + datasetName + '}\n')
    f.writelines('\\label{clusteringResultOn' + datasetName.upper() + '}\n')
    f.writelines('\\centering\n\\resizebox{\\textwidth}{!}\n{\n')
    f.writelines('\\begin{tabular}{c' + 'c' * METICS_LEN + '}\n\\toprule[2pt]\n')
    f.writelines('Dataset&\\multicolumn{' + str(METICS_LEN) + '}{c}{\\textbf{' + datasetName.upper() + '}}\\\\\n')

    # 指标
    f.writelines('\\midrule[1pt]\nMetrics ')
    for metics in list(data.columns)[:len(best)]:
        f.writelines('&' + metics + ' ')
    f.writelines('\\\\\n\\midrule\n')

    # 数据内容
    for algName in list(data.index):
        # 对比算法名称
        f.writelines(algName + ' ')
        for idx, value in enumerate(list(data.loc[algName])[:len(best)]):
            # 处理标准差
            try:
                std = list(data.loc[algName])[idx+len(best)]
                std = 0 if pd.isnull(std) else std
            except:
                std = 0
            # 处理均值
            if pd.isnull(value):
                f.writelines('&' + ' ')
            elif type(value) == type('\\'):
                f.writelines('&' + value)
            else:
                if value == best[idx]:
                    f.writelines('&\\textbf{' + '{:.3f}'.format(value) + '$\\pm$' + '{:.3f}'.format(std) + '} ')
                elif value == secondBest[idx]:
                    f.writelines('&\\underline{' + '{:.3f}'.format(value) + '$\\pm$' + '{:.3f}'.format(std) + '} ')
                else:
                    f.writelines('&' + '{:.3f}'.format(value) + '$\\pm$' + '{:.3f}'.format(std) + ' ')
        f.writelines('\\\\\n')

    # 结尾内容
    f.writelines('\\toprule[2pt]\n\\end{tabular}\n}\n\\end{table*}\n')
    f.close()


def getIndices(excelPath, datasetName):
    data = pd.read_excel(excelPath, sheet_name=datasetName, header=1) # 0 or 1
    # print(data)
    ansList = []
    for colName in list(data.columns):
        if 'Unnamed' in colName or '_std' in colName:
            pass
        else:
            ansList.append(colName)
    return ansList


def mkdir(path):
    folder = os.path.exists(path)
    if not folder:
        os.makedirs(path)


def main():
    root = tk.Tk() # 创建窗口
    root.title("Table Change")
    root.geometry("1000x750")
    # 提示文本
    lbl = Label(root, text="请选择excel文件", font=('宋体', 15)).place(x=10, y=10, anchor='nw')
    lblDataset = Label(root, text="请选择数据集", font=('宋体', 15)).place(x=10, y=55, anchor='nw')
    lblLatex = Label(root, text="LaTex 表格代码", font=('宋体', 15)).place(x=10, y=100, anchor='nw')

    # 文件路径提示文本
    lblFile = Label(root, bg='white', width=50, font=('宋体', 15))
    lblFile.place(x=180, y=10, anchor='nw')

    # sheet选择框
    combo = tkinter.ttk.Combobox(root)
    combo.place(x=180, y=55, anchor='nw')

    # 指标选择框
    indices = Frame(root)
    indices.pack(expand=False, fill="both", padx=10, pady=10)
    indices.place(x=380, y=50, anchor='nw')
    Label(indices, text='请选择需要展示的指标', font=('宋体', 15)).pack(side='left')

    # 创建输出文本框
    outputTxt = Text(root, width=139, height=46)
    outputTxt.place(x=10, y=135, anchor='nw')


    # 按钮选择文件
    def clicked():
        f_path = filedialog.askopenfilename()
        lblFile.configure(text=f_path)
        try:
            sheetNames = pd.ExcelFile(f_path).sheet_names
            combo['values'] = sheetNames
        except:
            combo['values'] = ()

    btn = Button(root, text="查询", font=('宋体', 15), command=clicked).place(x=700, y=10, anchor='nw')


    def clickedChoice():
        global COMBOPICKER
        indicesList = getIndices(excelPath=lblFile.cget("text"), datasetName=combo.get())
        COMBOPICKER = Combopicker(indices, values=indicesList)
        COMBOPICKER.pack(padx=10, pady=10, anchor='nw')
    btnChoice = Button(root, text="选择", font=('宋体', 15), command=clickedChoice).place(x=750, y=50, anchor='nw')


    # 确定按钮
    def clickedEnter():
        try:
            outputTxt.delete('1.0', 'end')
            datasetName = combo.get()
            excelPath = lblFile.cget("text")
            indicesList = COMBOPICKER.current_value.split(',')
            savePath = './textContent/' + datasetName + '.txt'
            mkdir('./textContent')
            processTable(datasetName, excelPath, savePath, indicesList)
            f = open(savePath, 'r', encoding='utf-8') # 输出内容到文本框中
            for lines in f:
                outputTxt.insert('insert', lines)
            f.close()
        except:
            pass

    btnEnter = Button(root, text="生成Latex代码", font=('宋体', 15), command=clickedEnter).place(x=830, y=50, anchor='nw')



    root.mainloop() # 进入主循环
    print('Done.')

if __name__ == "__main__":
    main()
