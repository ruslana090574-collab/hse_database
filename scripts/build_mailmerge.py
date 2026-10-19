"""Native Writer database fields plus actual LibreOffice mail merge jobs."""
import json
from pathlib import Path
from uno_common import *
from model import ROOT,DB_NAME,YEAR

TEMPLATES=[
('Полис_шаблон','19_Полисы_для_слияния_Writer','СТРАХОВОЙ ПОЛИС — УЧЕБНЫЙ ОБРАЗЕЦ',[
 ('Номер полиса','PolicyID'),('Дата оформления','IssuedOn'),('Застрахованное лицо','FullName'),('Дата рождения','BirthDate'),('Паспорт','PassportNo'),('Пол (M/F)','Gender'),('Страна поездки','CountryName'),('Начало страхования','StartsOn'),('Окончание страхования','EndsOn'),('Продолжительность, дней','Days'),('Страховое покрытие','PolicyCoverage'),('Валюта покрытия','CurrencyLabel'),('Стоимость услуг в день, руб.','DailyServicesRUB'),('Коэффициент страны','PolicyFactor'),('СТРАХОВАЯ ПРЕМИЯ, РУБ.','PremiumRUB'),('Курс на дату оформления','PolicyRate'),('Страховой агент','AgentName'),('Договор агента','ContractNo'),('Включённые услуги и тарифы','ServicesList')]),
('Договор_шаблон','20_Договоры_для_слияния_Writer','АГЕНТСКИЙ ДОГОВОР — УЧЕБНЫЙ ОБРАЗЕЦ',[
 ('Номер договора','ContractNo'),('Дата договора','ContractDate'),('Страховой агент','AgentName'),('Паспорт агента','AgentPassport'),('Дата выдачи паспорта','IssueDate'),('Кем выдан','Issuer'),('Начало договора','ContractStart'),('Окончание договора','ContractEnd'),('Вознаграждение, % от премии','CommissionPct'),('Банк','BankName'),('БИК','BankBIK'),('Счёт','BankAccount')])]

def main():
 ctx,desktop=connect();db=desktop.loadComponentFromURL(url(ROOT/DB_NAME),'_blank',0,(P('Hidden',True),))
 context=ctx.ServiceManager.createInstanceWithContext('com.sun.star.sdb.DatabaseContext',ctx)
 if context.hasByName('HSE_INSURANCE'):context.revokeObject('HSE_INSURANCE')
 context.registerObject('HSE_INSURANCE',db.DataSource)
 output={}
 for stem,query,title,fields in TEMPLATES:
  doc=desktop.loadComponentFromURL('private:factory/swriter','_blank',0,(P('Hidden',True),))
  page=doc.StyleFamilies.getByName('PageStyles').getByName('Standard');page.Width=21000;page.Height=29700;page.LeftMargin=1800;page.RightMargin=1800
  text=doc.Text;cursor=text.createTextCursor();cursor.CharFontName='Liberation Serif';cursor.CharHeight=12
  text.insertString(cursor,title+'\nСистема учёта клиентов страховой компании / ЛР-1\n\n',False)
  for label,column in fields:
   text.insertString(cursor,label+': ',False)
   master=doc.createInstance('com.sun.star.text.FieldMaster.Database')
   master.DataBaseName='HSE_INSURANCE';master.DataTableName=query;master.DataColumnName=column;master.DataCommandType=1
   field=doc.createInstance('com.sun.star.text.TextField.Database');field.attachTextFieldMaster(master)
   if column in ['IssuedOn','StartsOn','EndsOn','BirthDate','ContractDate','ContractStart','ContractEnd','IssueDate','DailyServicesRUB','PolicyFactor','PremiumRUB','PolicyRate','CommissionPct']:
    locale=__import__('uno').createUnoStruct('com.sun.star.lang.Locale');locale.Language='ru';locale.Country='RU'
    code='DD.MM.YYYY' if column in ['IssuedOn','StartsOn','EndsOn','BirthDate','ContractDate','ContractStart','ContractEnd','IssueDate'] else '0,00'
    key=doc.NumberFormats.queryKey(code,locale,True)
    if key<0:key=doc.NumberFormats.addNew(code,locale)
    field.DataBaseFormat=False;field.NumberFormat=key
   text.insertTextContent(cursor,field,False);text.insertString(cursor,'\n',False)
  text.insertString(cursor,'\nВсе персональные и финансовые сведения вымышлены. Документ создан только для лабораторной работы.\n',False)
  if stem.startswith('Полис'):
   text.insertString(cursor,'Период считается как разность дат окончания и начала (без включения конечной даты). Перечень выбранных услуг и их договорные тарифы указан выше. Для реального страхования требуется утверждённый страховщиком бланк.\n',False)
  else:
   text.insertString(cursor,'Предмет учебного договора: оформление страховых полисов клиентам. Вознаграждение исчисляется от страховой премии. Реквизиты приведены только для проверки заполнения и непригодны для платежей.\n',False)
  path=ROOT/'Документы'/f'{stem}.odt'
  doc.storeAsURL(url(path),(P('FilterName','writer8'),));doc.close(True)
  folder=ROOT/'Документы'/('Полисы_слияние' if stem.startswith('Полис') else 'Договоры_слияние');folder.mkdir(exist_ok=True)
  if any(folder.glob('*.odt')):raise RuntimeError('Output folder is not empty; preserve existing generated documents before repeating mail merge')
  merger=ctx.ServiceManager.createInstanceWithContext('com.sun.star.text.MailMerge',ctx)
  merger.DocumentURL=url(path);merger.DataSourceName='HSE_INSURANCE';merger.CommandType=1;merger.Command=query
  merger.OutputType=2;merger.OutputURL=url(folder);merger.SaveAsSingleFile=False;merger.FileNamePrefix='Полис_' if stem.startswith('Полис') else 'Договор_';merger.SaveFilter='writer8'
  merger.execute(())
  made=sorted(p.name for p in folder.glob('*.odt'));output[stem]=made
  print('Merged',stem,len(made),'documents',flush=True)
 (ROOT/'results/mailmerge.json').write_text(json.dumps(output,ensure_ascii=False,indent=2))
 db.store()
if __name__=='__main__':main()
