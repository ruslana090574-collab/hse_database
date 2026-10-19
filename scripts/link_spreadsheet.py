"""Configure and exercise a native Calc import range backed by the Base query."""
import json,uno
from uno_common import *
from model import ROOT,DB_NAME

def main():
 c,desk=connect();db=desk.loadComponentFromURL(url(ROOT/DB_NAME),'_blank',0,(P('Hidden',True),))
 context=c.ServiceManager.createInstanceWithContext('com.sun.star.sdb.DatabaseContext',c)
 if not context.hasByName('HSE_INSURANCE'):context.registerObject('HSE_INSURANCE',db.DataSource)
 doc=desk.loadComponentFromURL(url(ROOT/'Документы/Анализ_полисов.xlsx'),'_blank',0,(P('Hidden',True),))
 source='БД_Источник'
 if not doc.Sheets.hasByName(source):doc.Sheets.insertNewByName(source,doc.Sheets.Count)
 sheet=doc.Sheets.getByName(source)
 descriptor=(P('DatabaseName','HSE_INSURANCE'),P('SourceType',uno.Enum('com.sun.star.sheet.DataImportMode','QUERY')),P('SourceObject','09_Расчёт_стоимости_полисов'),P('IsNative',True))
 result=json.loads((ROOT/'results/queries_results.json').read_text())['09_Расчёт_стоимости_полисов']
 cols=result['columns']
 address=uno.createUnoStruct('com.sun.star.table.CellRangeAddress');address.Sheet=doc.Sheets.Count-1;address.StartColumn=0;address.StartRow=0;address.EndColumn=len(cols)-1;address.EndRow=len(result['rows'])
 if doc.DatabaseRanges.hasByName('ПолисыИзБазы'):doc.DatabaseRanges.removeByName('ПолисыИзБазы')
 doc.DatabaseRanges.addNewByName('ПолисыИзБазы',address)
 sheet.getCellRangeByName('A1').doImport(descriptor)
 linked=doc.DatabaseRanges.getByName('ПолисыИзБазы')
 print('Native import metadata:',[(x.Name,str(x.Value)) for x in linked.getImportDescriptor()],flush=True)
 order=['PolicyID','IssuedOn','FullName','CountryName','PolicyCoverage','CurrencyLabel','Days','DailyServicesRUB','PolicyFactor','PremiumRUB','CoverageRUB','AgentID','AgentName','ContractNo']
 view=doc.Sheets.getByName('Полисы')
 for r in range(1,len(result['rows'])+1):
  for col,name in enumerate(order):
   sourceletter=chr(65+cols.index(name))
   view.getCellByPosition(col,r).Formula=f"='БД_Источник'.{sourceletter}{r+1}"
 doc.calculateAll()
 assert int(sheet.getCellByPosition(0,1).Value)==1
 assert int(doc.Sheets.getByName('Итоги').getCellRangeByName('B2').Value)==8
 expected=sum(float(row[cols.index('PremiumRUB')]) for row in result['rows'] if row[cols.index('IssuedOn')].startswith(str(__import__('model').YEAR)))
 actual=doc.Sheets.getByName('Итоги').getCellRangeByName('B3').Value
 assert abs(actual-expected)<.001,(actual,expected)
 doc.storeAsURL(url(ROOT/'Документы/Анализ_полисов.ods'),(P('FilterName','calc8'),P('Overwrite',True)))
 doc.storeToURL(url(ROOT/'Документы/Анализ_полисов.xlsx'),(P('FilterName','Calc MS Excel 2007 XML'),P('Overwrite',True)))
 doc.close(True)
 # Calc's doImport attaches metadata to an anonymous range. Move that standard
 # ODF database-source-query element to the named range, then verify refresh.
 from zipfile import ZipFile,ZIP_DEFLATED
 from lxml import etree as ET
 path=ROOT/'Документы/Анализ_полисов.ods'
 with ZipFile(path) as archive:entries={n:archive.read(n) for n in archive.namelist()}
 root=ET.fromstring(entries['content.xml']);ns={'t':'urn:oasis:names:tc:opendocument:xmlns:table:1.0'};tag='{'+ns['t']+'}'
 ranges=root.findall('.//t:database-range',ns)
 named=next(n for n in ranges if n.get(tag+'name')=='ПолисыИзБазы')
 source=next(n for n in ranges if n.find(tag+'database-source-query') is not None)
 for child in list(named):named.remove(child)
 for child in list(source):named.append(child)
 source.getparent().remove(source)
 entries['content.xml']=ET.tostring(root,xml_declaration=True,encoding='UTF-8')
 with ZipFile(path,'w',ZIP_DEFLATED) as archive:
  for name,content in entries.items():archive.writestr(name,content)
 reopened=desk.loadComponentFromURL(url(path),'_blank',0,(P('Hidden',True),))
 linked=reopened.DatabaseRanges.getByName('ПолисыИзБазы')
 assert {x.Name:x.Value for x in linked.getImportDescriptor()}['DatabaseName']=='HSE_INSURANCE'
 linked.refresh();reopened.calculateAll()
 assert int(reopened.Sheets.getByName('Итоги').getCellRangeByName('B2').Value)==8
 reopened.store();reopened.close(True)
 (ROOT/'results/calc.json').write_text(json.dumps({'native_import':'HSE_INSURANCE / 09_Расчёт_стоимости_полисов','range':'ПолисыИзБазы','rows':len(result['rows']),'current_year_policies':8,'current_year_premium_rub':actual},ensure_ascii=False,indent=2))
 print('Native import and summary formulas verified:',8,actual,flush=True)
if __name__=='__main__':main()
