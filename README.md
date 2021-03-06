# MdbExportCSV
路線成本別計算軟體傳送檔案匯出CSV

由於一般人不會用 mdb 檔案，每次必須一個table逐步轉成 csv 檔案，耗時費力，尤其現在人手不夠沒有寫自動化功能，時間都不夠用。
直接下載使用: https://github.com/eric2100/MdbExportCSV/releases/download/1.0/MdbExportCSV.exe
# 編譯環境
Python 3.10.1 
需要 package: pyodbc、csv、glob 

# 編譯方法
```
pyinstaller -F MdbExportCSV.py -i ./MdbExportCSV.ico
```

# 使用方法
將 MdbExportCSV.exe 跟 路線成本別計算軟體4.0產生的mdb放在同一個資料夾，點兩下執行，該軟體會自動找 *.mdb 批次轉換出 csv。

```
T10908公司統編.mdb 檔案處理中....

匯出CSV ./T10908公司統編_000資料庫版本.csv
匯出CSV ./T10908公司統編_001識別資料.csv
匯出CSV ./T10908公司統編_002維修場成本.csv
匯出CSV ./T10908公司統編_003路線里程.csv
匯出CSV ./T10908公司統編_004車站成本.csv
匯出CSV ./T10908公司統編_005維修場車站歸屬關係.csv
匯出CSV ./T10908公司統編_007車站路線通行費.csv
匯出CSV ./T10908公司統編_008駕駛員薪資.csv
匯出CSV ./T10908公司統編_010駕駛員路線時數.csv
匯出CSV ./T10908公司統編_014車輛成本.csv
匯出CSV ./T10908公司統編_016車輛維修場成本.csv
匯出CSV ./T10908公司統編_017車輛里程時數.csv
匯出CSV ./T10908公司統編_019公司成本.csv
匯出CSV ./T10908公司統編_18項成本統計分析.csv
匯出CSV ./T10908公司統編_18項車站路線成本.csv
匯出CSV ./T10908公司統編_18項路線成本.csv
請按任何一鍵繼續...
```
