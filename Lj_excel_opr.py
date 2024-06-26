# workb  = xlrd.open_workbook(FileName)    #打开xlsx文件
# worksheets = workb.sheet_names() #抓取所有sheet页的名称
#
#
#
#
#
#

import xlrd,xlwt
import sys,os

outputFile = ""
inputFile = ""


def Read_excel(FileName):
    workb   = xlrd.open_workbook(FileName)    #打开xlsx文件
    table   = workb.sheets()[0]                #打开第一张表
    nrows   = table.nrows                     #获取表的行数
    worksheets = workb.sheet_names()
    
    for i in range(len(worksheets)):
        print("Page %d is %s" % (i, worksheets[i]))
    
    # 循环逐行输出
    for i in range(nrows):
        if (i == 0):
            continue
        print(table.row_values(i)[:4]) #取前5列值
        print(table.row_values(i)[0])
    print(table.row_values(0)[0])


def write_content(sheet, x, y, content):
    sheet.write(x, y, content)
    
writeSheetList = {'Title', '联系方式'}


def write_excel(FileName):
    workbk = xlwt.Workbook(FileName)
    #writeSheet = None
    #for i in range(len(writeSheetList)):
    writeSheet = workbk.add_sheet('title', cell_overwrite_ok=True)
    write_content(writeSheet, 0, 0 , "联系人")
    write_content(writeSheet, 0, 1 , "电话")
    
    write_content(writeSheet, 1, 0 , "张三")
    write_content(writeSheet, 1, 1 , "13648089999")
    
    workbk.save(FileName)
    
def main(argv):
    print(argv[0])
    print(argv[1])
    Read_excel("./data/input.xlsx")
    write_excel("./data/output.xlsx")
    pass
    	
if __name__ == "__main__":
    main(sys.argv)	#main(sys.argv[1:]) means point??? skip app name
    pass