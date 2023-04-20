import pandas as pd

METICS_LEN = 7    # 指标数

def processTable(datasetName, excelPath, savePath):
    datasetName = datasetName
    savePath = savePath
    data = pd.read_excel(excelPath, sheet_name=datasetName,header=1)

    # 删除未命名的列
    delList = []
    for colName in list(data.columns):
        if 'Unnamed' in colName:
            delList.append(colName)
    data = data.drop(delList, axis=1)
    # print(data.drop(delList, axis=1))

    # 记录每个指标最大值与次大值
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
    f.writelines('\\caption{Clustering performance on ' + datasetName + '. (The best result is in bold, and the second-best result is underlined.)}\n')
    f.writelines('\\label{clusteringResultOn' + datasetName.upper() + '}\n')
    f.writelines('\\centering\n\\scalebox{0.7}\n{\n')
    f.writelines('\\begin{tabular}{ccccccccc}\n\\toprule\n')
    f.writelines('Dataset&\\multicolumn{7}{c}{\\textbf{' + datasetName.upper() + '}}\\\\\n')

    # 指标
    f.writelines('\\midrule\nMetrics ')
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
    f.writelines('\\bottomrule\n\\end{tabular}\n}\n\\end{table*}\n')
    f.close()



def main():
    datasetName = 'dataset2'   # excel的sheet名称
    excelPath = 'test.xlsx'
    savePath = datasetName + '.txt'
    processTable(datasetName, excelPath, savePath)
    print('Done.')

if __name__ == "__main__":
    main()