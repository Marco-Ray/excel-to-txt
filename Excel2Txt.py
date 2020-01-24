import os
import glob
import sys
import xdrlib ,sys
import xlrd


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



def rwfun(filename):
    filename = filename.replace('\\','/')
    print(filename)
    newfile = path_output + '\\' + os.path.splitext(os.path.basename(filename))[0] + '.txt'
    print(newfile)
    f = open(newfile,'w', encoding='utf-8')
    et = ExcelToTxt(filename)
    
    if not et.openExcelFile(): 
        print('error')
        return
    sn = et.getAllSheetNames()
    for  s in sn:
        ls = et.excel_table_byname(s)
        data = '\n'.join(ls)
        data += '\n'
        # print(data)
        f.write(str(data))
        f.flush()
        print('='*100)
    f.close()

def run(text):    
    for filename in Filelist:
        rwfun(filename)

def get_all_files(path):
    Filelist = []
    for home, dirs, files in os.walk(path):
        for filename in files:
            Filelist.append(os.path.join(home,filename)) # 文件名列表，包含完整路径

            # # 文件名列表，只包含文件名
            # Filelist.append(filename)
    return Filelist
        
if __name__ == '__main__':
    path_input = os.path.dirname(sys.argv[0]) + r'\input'
    path_output = os.path.dirname(sys.argv[0]) + r'\output'
    Filelist = get_all_files(path_input)
    run(Filelist)
    #run(path)
    print('over ... ')