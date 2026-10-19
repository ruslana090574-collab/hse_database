"""Create an editable grouped Base Report Builder report (not a screenshot)."""
import uno,json
from uno_common import *
from model import ROOT,DB_NAME,QUERIES,YEAR
from com.sun.star.awt import Point,Size

def add(rep,section,text,x,y,width=3000,field=False,bold=False):
 service='com.sun.star.report.FormattedField' if field else 'com.sun.star.report.FixedText'
 item=rep.createInstance(service)
 if field:
  item.DataField=text;item.FormatsSupplier=rep
  locale=uno.createUnoStruct('com.sun.star.lang.Locale');locale.Language='ru';locale.Country='RU'
  code='DD.MM.YYYY' if 'IssuedOn' in text else ('@' if any(name in text for name in ['PassportNo','FullName','CountryName','CurrencyLabel']) else ('0,00' if any(name in text for name in ['Premium','RUB']) else '0'))
  key=rep.NumberFormats.queryKey(code,locale,True)
  if key<0:key=rep.NumberFormats.addNew(code,locale)
  item.FormatKey=key
 else:item.Label=text
 item.CharHeight=9
 if bold:item.CharWeight=150
 item.setPosition(Point(x,y));item.setSize(Size(width,650));section.add(item)
 return item

def counter(container,name,formula,initial=None):
 if initial is None:initial='rpt:1' if name.endswith('Count') else 'rpt:[PremiumRUB]'
 item=container.createFunction();item.Name=name;item.Formula=formula;item.InitialFormula=initial;item.PreEvaluated=True
 container.insertByIndex(container.Count,item)

def main():
 ctx,desktop=connect();db=desktop.loadComponentFromURL(url(ROOT/DB_NAME),'_blank',0,(P('Hidden',True),));cn=db.DataSource.getConnection('','')
 name='01_Продажа_полисов_прошлый_год'
 if db.ReportDocuments.hasByName(name):db.ReportDocuments.removeByName(name)
 fd=db.ReportDocuments.createInstanceWithArguments('com.sun.star.sdb.DocumentDefinition',(P('Name',name),P('ActiveConnection',cn),P('DocumentServiceName','com.sun.star.report.ReportDefinition')))
 db.ReportDocuments.insertByName(name,fd)
 rep=fd.getComponent();rep.Caption='Продажа полисов за прошлый календарный год';rep.Name=name
 rep.CommandType=2;rep.Command=QUERIES['08_Полисы_прошлый_год_для_отчёта'];rep.EscapeProcessing=False;rep.ActiveConnection=cn
 rep.ReportHeaderOn=True;rep.PageHeaderOn=True;rep.ReportFooterOn=True
 rep.ReportHeader.Height=1800
 add(rep,rep.ReportHeader,f'ПРОДАЖА ПОЛИСОВ ЗА {YEAR-1} ГОД',0,0,18500,bold=True)
 add(rep,rep.ReportHeader,'Группировка по клиенту. Итоги: число полисов и страховая премия в рублях.',0,750,18500)
 group=rep.Groups.createGroup();group.Expression='PassportNo';group.HeaderOn=True;group.FooterOn=True
 group.Header.Height=1600;group.Footer.Height=1400;rep.Groups.insertByIndex(0,group)
 add(rep,group.Header,'Клиент:',0,0,2000,bold=True);add(rep,group.Header,'field:[FullName]',2200,0,12000,field=True)
 add(rep,group.Header,'Паспорт:',0,750,2000);add(rep,group.Header,'field:[PassportNo]',2200,750,6000,field=True)
 columns=[('PolicyID','Полис',0,1300),('IssuedOn','Оформлен',1400,2600),('CountryName','Страна',4100,3300),('Days','Дней',7500,1100),('PremiumRUB','Премия, руб.',8700,3000),('PolicyCoverage','Покрытие',11800,3500),('CurrencyLabel','Валюта',15400,2100)]
 rep.PageHeader.Height=850;rep.Detail.Height=750
 for field,title,x,width in columns:
  add(rep,rep.PageHeader,title,x,0,width,bold=True)
  add(rep,rep.Detail,'field:['+field+']',x,0,width,field=True)
 for container,prefix in [(group.Functions,'Client'),(rep.Functions,'Total')]:
  counter(container,prefix+'Count','rpt:['+prefix+'Count] + 1')
  counter(container,prefix+'Premium','rpt:['+prefix+'Premium] + [PremiumRUB]')
 add(rep,group.Footer,'Итого по клиенту, полисов:',0,0,7000,bold=True)
 add(rep,group.Footer,'rpt:[ClientCount]',7200,0,1400,field=True)
 add(rep,group.Footer,'rpt:[ClientPremium]',8700,0,3000,field=True,bold=True)
 rep.ReportFooter.Height=1800
 add(rep,rep.ReportFooter,'ОБЩИЙ ИТОГ, полисов / премия руб.:',0,0,8500,bold=True)
 add(rep,rep.ReportFooter,'rpt:[TotalCount]',7200,750,1400,field=True)
 add(rep,rep.ReportFooter,'rpt:[TotalPremium]',8700,750,3000,field=True,bold=True)
 command=uno.createUnoStruct('com.sun.star.ucb.Command');command.Name='store';command.Handle=-1;command.Argument=None;fd.execute(command,0,None)
 db.store();print('Grouped report definition saved inside Base.',flush=True)
 try:
  engine=ctx.ServiceManager.createInstanceWithContext('com.sun.star.report.ReportEngine',ctx)
  engine.ReportDefinition=rep;engine.ActiveConnection=cn
  generated=engine.createDocumentModel()
  generated.storeAsURL(url(ROOT/'Документы/Продажи_прошлый_год_Base.odt'),(P('FilterName','writer8'),P('Overwrite',True)))
  generated.close(True);print('ReportBuilder generated ODT successfully.',flush=True)
  status='Generated successfully'
 except Exception as error:
  status=getattr(error,'Message',str(type(error)))
  print('Report engine:',status,flush=True)
 (ROOT/'results/report_builder.json').write_text(json.dumps({'report':name,'status':status},ensure_ascii=False,indent=2))
 cn.close();db.store()
if __name__=='__main__':main()
