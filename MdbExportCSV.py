import pyodbc
import csv
from glob import glob

def procMdb(mdb:str, tablename: str, output: str = '.'):
    '''
        mdb 要處理的access 
        tablename 要匯出的tablename
        output 輸出路徑
    '''
    DRV = '{Microsoft Access Driver (*.mdb)}'
    PWD = 'pw'
    
    try:
        con = pyodbc.connect('DRIVER={};DBQ={};PWD={}'.format(DRV, mdb, PWD))
        cur = con.cursor()
        SQL = 'SELECT * FROM {};'.format(tablename)
        rows = cur.execute(SQL).fetchall()

        fieldnames = [des[0] for des in cur.description]
        fieldnames = ['序號' if x == 'ID' else x for x in fieldnames]

        cur.close()
        con.close()
    except pyodbc.Error as ex:
        print(ex)

    try:
        csv_filename = output +'_' +tablename + '.csv'
        f = open( csv_filename , 'w', encoding='big5', newline='')
        w = csv.writer(f)
        w.writerow(fieldnames)
        w.writerows(rows)
        f.close
        print("匯出CSV {}".format( csv_filename ))
    except:
        print('建立csv發生錯誤 ')

if __name__ == '__main__':
    tableList = ['000資料庫版本',
                 '001識別資料',
                 '002維修場成本',
                 '003路線里程',
                 '004車站成本',
                 '005維修場車站歸屬關係',
                 '007車站路線通行費',
                 '008駕駛員薪資',
                 '010駕駛員路線時數',
                 '014車輛成本',
                 '016車輛維修場成本',
                 '017車輛里程時數',
                 '019公司成本',
                 '18項成本統計分析',
                 '18項車站路線成本',
                 '18項路線成本']

    ProcFile = glob('*.mdb')
    if ProcFile:
        for mdbfilename in ProcFile:
            filename = './'+ mdbfilename.split('.')[0]
            print( "{} 檔案處理中....\n".format(mdbfilename))

            for tablename in tableList:
                procMdb(mdbfilename, tablename,output=filename)

        input("請按任何一鍵繼續...")         


