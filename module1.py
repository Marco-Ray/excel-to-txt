from ExcelToTxt import ExcelToTxt
import os

def rwfun(filename):
    filename = filename.replace('\\','/')
    print(filename)
    newfile = os.path.splitext(os.path.basename(filename))[0]+'.txt'
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
    for filename in glob.glob(text):
        rwfun(filename)
        
if __name__ == '__main__':
    run(r'E:\FeigeDownload\sql\*.xls')
    run(r'E:\FeigeDownload\sql\*.xlsx')

    print('over ... ')
