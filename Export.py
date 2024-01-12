import cx_Oracle
import csv
import sys
import os
import datetime
from datetime import datetime
import pandas as pd

import warnings

sVersion = "ExportQuery v. 2024-01-12.003"

if len(sys.argv) == 4:
    path = sys.argv[1]
    month = sys.argv[2]
    year  = sys.argv[3]
else:
    sm = datetime.today().strftime("%m")
    month = f"{int(sm)-1:02}"
    year  = datetime.today().strftime("%Y")
    
if month == '00':  #run in January -> use last December
    month = '12'
    year = str(int(year) - 1)
    
path = f"z:/Save/TSS/Reporting/MBR/{year}/{month}/DATA-IN/" 

sqlPath = 'v:/GL/TSS/MBR/sql/'

print (f"ExportQry:: {sVersion}" )
print("Month: ", month) 
print("Year: ", year)
print("Path: ", path)
print("SQL-Path: ", sqlPath)



if not os.path.exists(path):
    os.makedirs(path)

#sys.exit()

def ExportDataFromSQLQuery(dbuser, dbpwd, dbconnection, sqlpath, sqlfilename, company, exportpath, filename, month, year):
    print(company, filename)
    connection = cx_Oracle.connect(user=dbuser, password=dbpwd, dsn=dbconnection)
    cursor = connection.cursor()
    # load sql
    sqlfile = open(sqlpath + sqlfilename,'r')
    sql = sqlfile.read()
    sqlfile.close()
    # replace parameter in sql
    sql=sql.replace(':month', month)
    sql=sql.replace(':year', year)

    fName = exportpath + filename + '-' + company 

    df_ora = pd.read_sql(sql, con=connection)
    df_ora.to_csv(fName + '.csv', sep=';', decimal = '.', encoding="utf-8")
    df_ora.to_excel(fName + ".xlsx", index=False)
    '''
    #execute query and export data
    cursor.execute(sql)

    with open(fName, 'w', newline='') as export_file:
        writer = csv.writer(export_file, quoting=csv.QUOTE_ALL, delimiter=';')
        col_names = []
        for i in range(0, len(cursor.description)):
            col_names.append(cursor.description[i][0])
        writer.writerow(col_names)    
        for row in cursor:
           for v in row:
            if 
            writer.writerow(row)
    '''
    connection.close


warnings.filterwarnings('ignore')

try:
    cx_Oracle.init_oracle_client(lib_dir=r"c:\Program Files\ORACLE\Client")


    ExportDataFromSQLQuery("PPS_DI", "nretni", "srv-pps01.di-netz.de/XE.di-netz.de", sqlPath, "angebote.sql", "di", path,"angebote", month, year)
                    
    ExportDataFromSQLQuery("PPS_DI", "nretni", "srv-pps01.di-netz.de/XE.di-netz.de", sqlPath, "won.sql", "di", path,"won", month, year)
                        
    ExportDataFromSQLQuery("PPS_DI", "nretni", "srv-pps01.di-netz.de/XE.di-netz.de", sqlPath, "lost.sql", "di", path, "lost", month, year)                    

    ExportDataFromSQLQuery("PPS_TTE", "nretni", "srv-pps01.di-netz.de/XE.di-netz.de", sqlPath, "angebote.sql", "tte", path,"angebote", month, year)
                    
    ExportDataFromSQLQuery("PPS_TTE", "nretni", "srv-pps01.di-netz.de/XE.di-netz.de", sqlPath, "won.sql", "tte", path,"won", month, year)
                        
    ExportDataFromSQLQuery("PPS_TTE", "nretni", "srv-pps01.di-netz.de/XE.di-netz.de", sqlPath, "lost.sql", "tte", path, "lost", month, year)   

    #ExportDataFromSQLQuery("PPS_DI", "nretni", "srv-pps01.di-netz.de/XE.di-netz.de", ".\\sql\\", "rechnungen.sql",
    #                    "di", path, "rechnungen", month, year)

    #ExportDataFromSQLQuery("PPS_TTE", "nretni", "srv-pps01.di-netz.de/XE.di-netz.de", ".\\sql\\", "rechnungen.sql",
    #                    "tte", path, "rechnungen", month, year)                    

    #ExportDataFromSQLQuery("PPS_DI", "nretni", "srv-pps01.di-netz.de/XE.di-netz.de", ".\\sql\\", "ka_netto.sql",
    #                    "di", path, "ka_netto", month, year)

    #ExportDataFromSQLQuery("PPS_TTE", "nretni", "srv-pps01.di-netz.de/XE.di-netz.de", ".\\sql\\", "ka_netto.sql",
    #                    "tte", path, "ka_netto", month, year)                      


    print("***  Export SQL  -done-  ***")

    exit(0)

except Exception as e:
    print(f"******** ERROR in ExportQry  *********")
    print(e)          

    exit(-1)      
