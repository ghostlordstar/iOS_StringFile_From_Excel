

import time
import sys
import os
from openpyxl import load_workbook

def logo():
    print("""
        
  ______ __   __ _____  ______  _        _______  ____      _____  _    _               
 |  ____|\ \ / // ____||  ____|| |      |__   __|/ __ \    / ____|| |  (_)              
 | |__    \ V /| |     | |__   | |         | |  | |  | |  | (___  | |_  _  _ __    __ _ 
 |  __|    > < | |     |  __|  | |         | |  | |  | |   \___ \ | __|| || '_ \  / _` |
 | |____  / . \| |____ | |____ | |____     | |  | |__| |   ____) || |_ | || | | || (_| |
 |______|/_/ \_\\\_____||______||______|    |_|   \____/   |_____/  \__||_||_| |_| \__, |
                                                                                   __/ |
                                                                                  |___/ 
        
    """)

def __help():
    print("此脚本用来帮助开发者从excel生成国际化文件（iOS、Android）")
    print("首先需要安装`openpyxl`库，可以使用`pip3 install openpyxl`命令安装")
    print("-------------参数--------------\n")
    print("参数 `-h` 或 `--help`    查看帮助.")
    print("参数 `-ep` 或 `--excelPath`         excel文件的目录，默认为`~/Desktop/international_test.excel`")
    print("参数 `-sp` 或 `--stringFilePath`    导出国际化文件的目录，默认为`~/Desktop/InternationalStrings/`")
    print("参数 `-kc` 或 `--keyColumn`         指定key在excel中的列，默认为`B`")
    print("参数 `-vc` 或 `--valueColumn`       指定value在excel中的列，默认为`[C]`，可以传入数组(以`,`分割)，可以根据传入数组生成不同的国际化文件")
    print("参数 `-is` 或 `--ignoreSheets`      指定忽略的sheet，可以传入数组，(以`,`分割)")
    print("-------------------------------\n")


# 打开excel
def loadExcel(excelFilePath):
    wb = load_workbook(filename=excelFilePath)
    return wb

# 过滤不需要的sheet
def filterSheetNames(filtration,sheetNames):
    newNames = sheetNames
    for filter in filtration:
        if filter in newNames:
            newNames.remove(filter)
    return newNames

# 读取sheet并处理所有文案， 并将处理后的文案保存到`allStringDict`， 将所有key保存到`keys`
def processSheetStringList(sheetName, wb, valueColumn, needWriteKeys):
    tmpSheet = wb[sheetName] # 取出指定的sheet
    tmpSheetMaxRow = tmpSheet.max_row
    valueName = tmpSheet["%s1"%valueColumn].value
    for row in range(tmpSheetMaxRow):
        keyIndex = "B%s"%(row + 2)  # 从第二行开始读取
        valueIndex = "%s%s"%(valueColumn,row + 2)
        key = "%s"%tmpSheet[keyIndex].value
        value = "%s"%tmpSheet[valueIndex].value
        if key != "None" and len(key) > 0 and value != None and len(value) > 0:   # key和value都不为空时写入文件
            if needWriteKeys == True:
                if (key in keys) == False:
                    keys.append(key)    # 添加新key
                else:
                    keys.remove(key)    # 删掉原来的key
                    keys.append(key)    # 将key添加到尾部
            allStringDict[key] = value
    return valueName

# 将处理好的文件写入国际化文件
def writeAllStringToIntenationalFile(file):
    for key in keys:
        value = allStringDict[key]
        if len(value) > 0:
            file.write('\"%s\" = \"%s\";\n' %(key, value))

# 创建文件并添加文件头
def create_iOS_InitializeStringFile(path):
    if path != None and len(path) > 0:
        if os.path.exists(path) == False:
            os.makedirs(path)
        tmpStringFile = open('%s/Localizable.strings'%path, 'w')

    if tmpStringFile != None:
        tmpStringFile.write("""
/* 
  Localizable.strings
  YourProjectName

  Created by walker on %s.
  Copyright © %s YourProjectName. All rights reserved.
*/


"""%(time.strftime("%Y/%m/%d %H:%M", time.localtime()), time.strftime("%Y", time.localtime())))
    return tmpStringFile

# 写入国际化文件
def writeInternationalStringToFile(filePath):
    file = create_iOS_InitializeStringFile(filePath)
    writeAllStringToIntenationalFile(file)
    file.close()
    print(' 🤖 [%s/Localizable.strings]写入完成！\n' % filePath)

# 转换excel中指定的value为国际化文件
def convertExcelToString(valueColumn):
    valueName = ""
    for sheetName in needProcessSheetNames: # 遍历所有sheet，并将所有的key-value缓存起来
        tmpValueName = processSheetStringList(sheetName, excel, valueColumn, True)
        if tmpValueName != None and len(tmpValueName) > 0 and len(valueName) <= 0:
            valueName = tmpValueName
    writeInternationalStringToFile("%s/%s" % (outPath, valueName))

# main 函数
if __name__ == "__main__":
    # 显示logo
    logo()
    # 输出传入参数
    print('传入参数为：%s\n'%sys.argv)
    # 初始化变量
    keys = []   # key存的数组，用来保存所有的key
    allStringDict = {}  # key value存的字典
    excelPath = '/Users/apple/Desktop/international_test.xlsx'
    outPath = 'InternationalStrings'
    keyColumn = 'B'
    valueColumns = ['C'] # 默认转换的国际化列名称
    ignoreSheets = ['what\'s new','Backend']

    # 分割并处理传入的参数
    for tmpArg in sys.argv:
        if '-h' in tmpArg or '--help' in tmpArg:
            __help()
            exit(0)
        elif '-ep=' in tmpArg:
            excelPath = tmpArg.replace('-ep=', '')
        elif '--excelPath=' in tmpArg:
            excelPath = tmpArg.replace('--excelPath=', '')
        elif '-sp=' in tmpArg:
            outPath = tmpArg.replace('-sp=', '')
        elif '--stringFilePath=' in tmpArg:
            outPath = tmpArg.replace('--stringFilePath=', '')
        elif '-kc=' in tmpArg:
            keyColumn = tmpArg.replace('-kc=', '').upper()
        elif '--keyColumn=' in tmpArg:
            keyColumn = tmpArg.replace('--keyColumn=', '').upper()
        elif '-vc=' in tmpArg:
            valueColumns = tmpArg.replace('-vc=', '').upper().split(',')
        elif '--valueColumn=' in tmpArg:
            valueColumns = tmpArg.replace('--valueColumn=', '').upper().split(',')
        elif '-is=' in tmpArg:
            ignoreSheets = tmpArg.replace('-is=', '').upper().split(',')
        elif '--ignoreSheets=' in tmpArg:
            ignoreSheets = tmpArg.replace('--ignoreSheets=', '').upper().split(',')

    #print(excelPath, outPath, keyColumn, valueColumns, ignoreSheets)

    # 读取excel文件
    excel = loadExcel(excelPath)

    # 过滤不需要的sheet
    needProcessSheetNames = filterSheetNames(ignoreSheets, excel.sheetnames)

    # 遍历转换所有国际化文案
    for vc in valueColumns:
        convertExcelToString(vc)    # 逐列转换string文件
    print('✅ 处理完成！')