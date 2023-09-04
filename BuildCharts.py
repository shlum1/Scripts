#import collections
#import collections.abc

#from pptx import Presentation
#from pptx.enum.shapes import MSO_SHAPE
#from pptx.dml.color import RGBColor
#from pptx.util import Inches, Pt
#from pptx.enum.dml import MSO_THEME_COLOR

#from pptx.chart.data import CategoryChartData
#from pptx.enum.chart import XL_CHART_TYPE
#from pptx.chart.data import ChartData
#from pptx.util import Inches
#import numpy as np 

from datetime import datetime, timedelta

#import plotly.graph_objects as go
import plotly.express as px
import pandas as pd
import locale
from sys import argv


sVersion = "BuildCharts v. 2023-03-21.005"

gYear = '2023'
sCompanies = ['DI', 'TTE']
sDataDir="xx/Data-IN"
showChart = False





def BuildOrderHistory():
   '''
   import plotly.graph_objects as go
   x = [1, 2, 3, 4]

   fig = go.Figure()
   fig.add_trace(go.Bar(x=x, y=[1, 4, 9, 16]))
   fig.add_trace(go.Bar(x=x, y=[6, -8, -4.5, 8]))
   fig.add_trace(go.Bar(x=x, y=[-15, -3, 4.5, -8]))
   fig.add_trace(go.Bar(x=x, y=[-1, 3, -3, -4]))

   fig.update_layout(barmode='relative', title_text='Relative Barmode')
   fig.show()
   '''


   pass


#******************************************************************************************************************
#*
#*
#*
#******************************************************************************************************************
def BuildInvoiceSunburst(company:str, sMonth: str):
   try:
      minRev = 500
      addDays = 5 
      
      if not sMonth.isnumeric():
         sMonth = '02'   # for debugging only

      if sMonth == '12':
         year = int(gYear) + 1
         month = 1
      else:
        year = int(gYear)
        month = int(sMonth)+1


      chkDate = datetime(year, month, 1, 0, 0, 0)
      
      due = chkDate - timedelta(days=addDays)
      sDue=due.strftime('%Y-%m-%d 00:00:00')

      locale.setlocale(locale.LC_ALL, 'German')

      print(f'  --> BuildInvoiceSunburst({company}, {sDue}) ...')



      #colors=["red", "green", "blue", "goldenrod", "magenta"]
      colorMap={'Issued':'#1074e6', 'Overdue':'#033da8', '1st Reminder': '#e35263', '2nd Reminder': '#c92450', '3rd Reminder': '#a1032d', '(?)':'white'}

      fName = f'{sDataDir}/{company}_{year}_{month:02}_Rechnungen.csv'
      print(fName)
      df=pd.read_csv(fName, sep=';', decimal=",", encoding = 'latin1')
      df.fillna(0, inplace=True)


      sQry ='(STATUS not in ["Storniert", "Ungültig", "Storniert", "Storniert Teilbezahlt", "Bezahlt"]) & (RECHNUNGSWERT > 0.0) & (KUNDENNUMMER not in [410000, 420000])'   # 
      df.query(sQry, inplace = True) 
      
      


   except Exception as e:
      print (f" Exception bei BuildInvoiceSunburst {e}") 

   #print(df.to_string())
   
   df['FAELLIGKEITSDATUM'] = pd.to_datetime(df['FAELLIGKEITSDATUM'],  format="%d.%m.%Y")
   df['STAT']='Overdue'  # the later remains
   df.loc[(df['FAELLIGKEITSDATUM'] >= sDue), 'STAT'] = 'Issued'
   df.loc[df['STATUS'] == 'Gemahnt', 'STAT'] = '1st Reminder'
   df.loc[df['STATUS'] == '2. Mahnung', 'STAT'] = '2nd Reminder'
   df.loc[df['STATUS'] == '3. Mahnung', 'STAT'] = '3rd Reminder'

   df['F_OPEN'] = df['OFFENE_POSTEN_EUR'].apply(lambda x: locale.format_string('%.0f', x, True))
   df['KUNDE'] = df['KUNDE'].apply(lambda s: s[:20])
   
   total = df['OFFENE_POSTEN_EUR'].sum()
   overdue = total - df[df['STAT'] == 'Issued'].OFFENE_POSTEN_EUR.sum()

   total = locale.format_string('%.0f', total, True)
   overdue = locale.format_string('%.0f', overdue, True)

   dx=df[df['OFFENE_POSTEN_EUR'] > minRev] 
   
   sTitle = f'{company} ({month}/{gYear})   {df.STATUS.count()} invoices, total: €{total}, overdue: €{overdue}'

   fig=px.sunburst(data_frame = dx, path = ["STAT", "KUNDE", "F_OPEN"], maxdepth = -1, values = 'OFFENE_POSTEN_EUR', color = 'STAT', color_discrete_map = colorMap, width=800, height=800, title=sTitle )
   if showChart:
      fig.show()

   dx.to_csv(f'BuildCharts-{company}.csv', index=None, sep=';', mode='a')
   df.to_excel(f'BuildCharts-{company}.xlsx', index=False)
   fig.write_image(f"Invoices-{company}.svg")
   fig.write_image(f"Invoices-{company}.png")



###################################################################################################################################################################################
###################################################################################################################################################################################
if __name__ == '__main__':

   print('#### ', sVersion)
    
   if len(argv) >= 3:
        path = argv[1]
        month = argv[2]
        showChart = False
        try:
            gYear = argv[3]
        except:
            pass  #use pre-defined year
   else:
        month = 'xx'
        path = 'Test/xx'
        print("Usage: python won&lost.py <path> <month>")
        showChart = True
        #exit(1)


   try:
      sDataDir = f'{path}/data-in'
      print(sDataDir, gYear, month)



      for c in sCompanies:
         BuildInvoiceSunburst(c, month)


      print('... BuildCharts -done- \n\n')
      exit(0)

   except Exception as e:
      print("******** ERROR in BuildCharts *********")
      print(e)          

      exit(-1)      

