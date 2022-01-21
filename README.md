# iOS_StringFile_From_Excel
从excel表格中生成iOS的国际化文案

## 功能:
- [x] 支持本地Excel文件（xlsx、xls）
- [x] 支持过滤sheet 
- [x] 支持指定某些列进行国际化文件转换
- [x] 支持key去重，最新添加的value覆盖旧值
- [ ] 支持生成Android国际化文件

## 截图展示：
![Excel 示例](https://github.com/ghostlordstar/iOS_StringFile_From_Excel/blob/main/shot/excel.png?raw=true)
![en Localizable.strings 示例](https://github.com/ghostlordstar/iOS_StringFile_From_Excel/blob/main/shot/en.png?raw=true)
![zh-hans Localizable.strings 示例](https://github.com/ghostlordstar/iOS_StringFile_From_Excel/blob/main/shot/zh-hans.png?raw=true)

## 准备工作：
1. python3环境
2. 安装`openpyxl`插件(可用命令`pip3 install openpyxl`安装)

## 用法:
```commandline
python3 CreateStringFile.py -ep=Example/intenational_test.xlsx -sp=Example/StringFiles/ -vc=C,D 
```
更多参数用法：
```commandline
参数 `-h` 或 `--help`               查看帮助
参数 `-ep` 或 `--excelPath`         excel文件的目录，默认为`~/Desktop/international_test.excel`
参数 `-sp` 或 `--stringFilePath`    导出国际化文件的目录，默认为`~/Desktop/InternationalStrings/`
参数 `-kc` 或 `--keyColumn`         指定key在excel中的列，默认为`B`
参数 `-vc` 或 `--valueColumn`       指定value在excel中的列，默认为`[C]`，可以传入数组(以`,`分割)，可以根据传入数组生成不同的国际化文件
参数 `-is` 或 `--ignoreSheets`      指定忽略的sheet，可以传入数组，(以`,`分割)
```
