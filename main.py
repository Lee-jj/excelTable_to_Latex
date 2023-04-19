import pandas as pd

datasetName = 'dataset2'
savePath = datasetName + '.txt'
data = pd.read_excel('test.xlsx', sheet_name=datasetName,header=1)

# 记录每个指标最大值与次大值
best = [0] * len(list(data.columns))
secondBest = [0] * len(list(data.columns))
for idx, matric in enumerate(list(data.columns)):
    temp = []
    for value in list(data[matric]):
        if pd.isnull(value) or type(value) == type('\\'):
            pass
        else:
            temp.append(value)
    temp.sort(reverse=True)
    best[idx] = temp[0] if len(temp) > 0 else 0
    secondBest[idx] = temp[1] if len(temp) > 1 else 0
    

f = open(savePath, 'w')
f.writelines('\\begin{table*}[t]\n')
f.writelines('\\caption{Clustering performance on ' + datasetName + '. (The best result is in bold, and the second-best result is underlined.)}\n')
f.writelines('\\label{clusteringResultOn' + datasetName.upper() + '}\n')
f.writelines('\\centering\n\\scalebox{0.7}\n{\n')
f.writelines('\\begin{tabular}{ccccccccc}\n\\toprule\n')
f.writelines('Dataset&\\multicolumn{7}{c}{\\textbf{' + datasetName.upper() + '}}\\\\\n')

# 指标
f.writelines('\\midrule\nMetrics ')
for metics in list(data.columns):
    f.writelines('&' + metics + ' ')
f.writelines('\\\\\n\\midrule\n')

# 数据内容
for algName in list(data.index):
    # 对比算法名称
    f.writelines(algName + ' ')
    for idx, value in enumerate(list(data.loc[algName])):
        if pd.isnull(value):
            f.writelines('&' + ' ')
        elif type(value) == type('\\'):
            f.writelines('&' + value)
        else:
            if value == best[idx]:
                f.writelines('&\\textbf{' + '{:.3f}'.format(value) + '$\\pm$0.000}' + ' ')
            elif value == secondBest[idx]:
                f.writelines('&\\underline{' + '{:.3f}'.format(value) + '$\\pm$0.000}' + ' ')
            else:
                f.writelines('&' + '{:.3f}'.format(value) + '$\\pm$0.000' + ' ')
    f.writelines('\\\\\n')

f.writelines('\\bottomrule\n\\end{tabular}\n}\n\\end{table*}\n')
f.close()

print('Done.')