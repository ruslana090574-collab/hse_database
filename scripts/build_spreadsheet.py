from pathlib import Path
import json
from datetime import date,datetime
from decimal import Decimal
from openpyxl import Workbook,load_workbook
from openpyxl.styles import Font,PatternFill,Alignment
from openpyxl.chart import BarChart,Reference
from openpyxl.worksheet.table import Table,TableStyleInfo
from model import ROOT,YEAR,LABELS

def main():
 data=json.loads((ROOT/'results/queries_results.json').read_text())
 records=[dict(zip(data['09_Расчёт_стоимости_полисов']['columns'],row)) for row in data['09_Расчёт_стоимости_полисов']['rows']]
 wb=Workbook();params=wb.active;params.title='Параметры'
 params.append(['Параметр','Значение']);params.append(['Начало периода',date(YEAR,1,1)]);params.append(['Конец периода',date(YEAR,12,31)])
 params.append(['Источник данных','HSE_INSURANCE / 09_Расчёт_стоимости_полисов'])
 params.append(['Обновление','Зарегистрируйте ODB как HSE_INSURANCE, затем Данные → Обновить диапазон.'])
 params.append(['Данные','Учебные, вымышленные. Курсы не являются актуальными котировками ЦБ.'])
 params['B2'].number_format=params['B3'].number_format='DD.MM.YYYY'
 cols=['PolicyID','IssuedOn','FullName','CountryName','PolicyCoverage','CurrencyLabel','Days','DailyServicesRUB','PolicyFactor','PremiumRUB','CoverageRUB','AgentID','AgentName','ContractNo']
 ws=wb.create_sheet('Полисы');ws.append([LABELS.get(x,x) for x in cols])
 numeric={'PolicyID','PolicyCoverage','Days','DailyServicesRUB','PolicyFactor','PremiumRUB','CoverageRUB','AgentID'}
 for record in records:
  row=[]
  for col in cols:
   value=record[col]
   if col=='IssuedOn':value=date.fromisoformat(value)
   elif col in numeric:value=float(value)
   row.append(value)
  ws.append(row)
 for row in ws.iter_rows(min_row=2):
  row[1].number_format='DD.MM.YYYY'
  for i in (7,8,9,10):row[i].number_format='#,##0.00'
 ws.auto_filter.ref=ws.dimensions;ws.freeze_panes='C2'
 tab=Table(displayName='PolicyData',ref=ws.dimensions);tab.tableStyleInfo=TableStyleInfo(name='TableStyleMedium2',showRowStripes=True);ws.add_table(tab)
 service=wb.create_sheet('Услуги_полисов');sr=data['15_Услуги_полисов_исторические']
 service.append([LABELS.get(c,c) for c in sr['columns']]);service.extend([]) if False else None
 for row in sr['rows']:service.append(row)
 agents=wb.create_sheet('Агенты');agents.append(['Код агента','ФИО агента','Число полисов','Выручка, руб.','Средний полис, руб.','Комиссия, %','Комиссия, руб.','Лидер'])
 q=data['10_Агенты_количество_за_период'];arows=q['rows']
 for i,row in enumerate(arows,2):
  agents.append([int(row[0]),row[1],f'=COUNTIFS(Полисы!$L$2:$L$14,A{i},Полисы!$B$2:$B$14,">="&Параметры!$B$2,Полисы!$B$2:$B$14,"<="&Параметры!$B$3)',f'=SUMIFS(Полисы!$J$2:$J$14,Полисы!$L$2:$L$14,A{i},Полисы!$B$2:$B$14,">="&Параметры!$B$2,Полисы!$B$2:$B$14,"<="&Параметры!$B$3)',f'=IF(C{i}=0,"Нет продаж",D{i}/C{i})',8+int(row[0]),f'=ROUND(D{i}*F{i}/100,2)',f'=IF(AND(C{i}>0,C{i}=MAX($C$2:$C$5)),"Лидер","")'])
  for cell in [f'D{i}',f'E{i}',f'G{i}']:agents[cell].number_format='#,##0.00'
 for label,col in [('Количество полисов',3),('Выручка за период, руб.',4)]:
  chart=BarChart();chart.title=label;chart.y_axis.title=label;chart.add_data(Reference(agents,min_col=col,min_row=1,max_row=5),titles_from_data=True);chart.set_categories(Reference(agents,min_col=2,min_row=2,max_row=5));chart.height=8;chart.width=20;agents.add_chart(chart,'A8' if col==3 else 'A24')
 summary=wb.create_sheet('Итоги');summary.append(['Показатель','Значение']);summary.append(['Полисов за период','=SUM(Агенты!C2:C5)']);summary.append(['Выручка за период, руб.','=SUM(Агенты!D2:D5)']);summary.append(['Средняя стоимость, руб.','=IF(B2=0,"Нет полисов",B3/B2)']);summary.append(['Комиссия агентам, руб.','=SUM(Агенты!G2:G5)']);summary.append(['Полисов всего',len(records)]);summary.append(['Полисов прошлого года',5])
 for row in (3,4,5):summary[f'B{row}'].number_format='#,##0.00'
 for sheet in wb:
  sheet.freeze_panes=sheet.freeze_panes or 'A2'
  sheet.sheet_view.showGridLines=False
  for cell in sheet[1]:cell.font=Font(name='Liberation Sans',bold=True,color='FFFFFF',size=11);cell.fill=PatternFill('solid',fgColor='17365D');cell.alignment=Alignment(wrap_text=True,vertical='center')
  sheet.row_dimensions[1].height=32
  for row in sheet.iter_rows(min_row=2):
   for cell in row:cell.font=Font(name='Liberation Sans',size=11);cell.alignment=Alignment(vertical='top')
  for column in sheet.columns:
   letter=column[0].column_letter
   width=max(len(str(cell.value or '')) if not str(cell.value).startswith('=') else 14 for cell in list(column)[:15])+2
   sheet.column_dimensions[letter].width=min(max(width,14),46)
 wb.save(ROOT/'Документы'/'Анализ_полисов.xlsx')
 print('Workbook with all policy/service data, date parameters, formulas and 2 charts saved.')
if __name__=='__main__':main()
