import xdrlib ,sys
import xlrd
# import os
class ExcelToTxt:
    def __init__(self, path):
        self._path = path
        self._data = None
        self._sheet_names = None
        

        
        
        
    def openExcelFile(self):
        try:
            self._data = xlrd.open_workbook(self._path)
            return self._data
        except :
            print('error open excel file ...')
            return False

    def getAllSheetNames(self):
        return self._data.sheet_names()

    def excel_table_byindex(self,colnameindex=0,by_index=0):
        table = self._data.sheets()[by_index]
        nrows = table.nrows #行数
        ncols = table.ncols #列数
        if nrows < 1: return []
        print('nrows: ', nrows)
        print('ncols: ', ncols)
        list = []
        for rownum in range(0,nrows):
             row = table.row_values(rownum)
             list.append(','.join(row))
        return list

    #根据名称获取Excel表格中的数据   参数:file：Excel文件路径     colnameindex：表头列名所在行的所以  ，by_name：Sheet1名称
    def excel_table_byname(self,by_name=u'Sheet1'):
        table = self._data.sheet_by_name(by_name)
        nrows = table.nrows #行数 
        if nrows < 1: return []
        list =[]
        for rownum in range(0,nrows):
             row = table.row_values(rownum)
             s= ','.join(str(x) for x in row)
             list.append(s)
        return list



