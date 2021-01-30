


import time
from openpyxl import load_workbook

def logo():
    print("""
        
  ______ __   __ _____  ______  _        _______  ____      _____  _    _               
 |  ____|\ \ / // ____||  ____|| |      |__   __|/ __ \    / ____|| |  (_)              
 | |__    \ V /| |     | |__   | |         | |  | |  | |  | (___  | |_  _  _ __    __ _ 
 |  __|    > < | |     |  __|  | |         | |  | |  | |   \___ \ | __|| || '_ \  / _` |
 | |____  / . \| |____ | |____ | |____     | |  | |__| |   ____) || |_ | || | | || (_| |
 |______|/_/ \_\\_____||______||______|    |_|   \____/   |_____/  \__||_||_| |_| \__, |
                                                                                   __/ |
                                                                                  |___/ 
        
    """)


# 打开excel
def loadExcel(excelFilePath):
    wb = load_workbook(filename=excelFilePath)
    return wb


# 过滤不需要的sheet
def filterSheetNames(filtration,sheetNames):
    newNames = sheetNames
    for filter in filtration:
        newNames.remove(filter)
    return newNames


# 读取sheet并写入文件
def writeStringFileFromSheet(sheetName, wb):
    tmpSheet = wb[sheetName] # 取出指定的sheet
    tmpSheetMaxRow = tmpSheet.max_row
    for row in range(tmpSheetMaxRow):
        keyIndex = "B%s"%(row + 2)  # 从第二行开始读取
        valueIndex = "C%s"%(row + 2)
        key = tmpSheet[keyIndex].value
        value = tmpSheet[valueIndex].value
        if key != None and len(key) > 0 and value != None and len(value) > 0:   # key和value都不为空时写入文件
            stringFile.write('\"%s\" = \"%s\";\n' %(key, value))


# 把指定sheet文案写入国际化文件
def writeStringFromSheets(sheets, wb):
    for sheet in sheets:
        writeStringFileFromSheet(sheet, wb)


# 创建文件并添加文件头
def openAndInitializeStringFile(path):

    if path != None and len(path) > 0:
        tmpStringFile = open(path, 'w')
    else:
        tmpStringFile = open('Localizable.strings', 'w')

    tmpStringFile.write("""
/* 
  Localizable.strings
  Playhouse

  Created by walker on %s.
  Copyright © 2021 LFG. All rights reserved.
*/
    """%time.strftime("%Y/%m/%d %H:%M", time.localtime()))
    return tmpStringFile

if __name__ == "__main__":

    logo()

    # 读取excel文件
    excel = loadExcel('/Users/walker/Desktop/lfg_lfgInternational.xlsx')

    # 过滤不需要的sheet
    needProcessSheetNames = filterSheetNames(['what\'s new','Backend'], excel.sheetnames)

    # 创建国际化文件
    stringFile = openAndInitializeStringFile(None)

    # 将指定sheets写入到StringFile
    writeStringFromSheets(needProcessSheetNames, excel)

    # TODO: 去重

    # 关闭文件
    stringFile.close()
