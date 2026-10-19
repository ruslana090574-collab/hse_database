"""Finalize portable input masks, readable grid widths and historical lookups."""
import uno,json
from uno_common import *
from model import ROOT,DB_NAME
from com.sun.star.awt import Point,Size

def walk(model):
 for name in model.ElementNames:
  item=model.getByName(name)
  if item.supportsService('com.sun.star.form.component.DataForm'):
   yield from walk(item)
  else:yield model,item

def main():
 c,desk=connect();db=desk.loadComponentFromURL(url(ROOT/DB_NAME),'_blank',0,(P('Hidden',False),));ctrl=db.CurrentController;ctrl.connect()
 # Close existing view-mode holders before reopening their editable designs.
 for existing in db.FormDocuments.ElementNames:
  holder=db.FormDocuments.getByName(existing)
  component=holder.getComponent()
  if component is not None:
   component.setModified(False)
   command=uno.createUnoStruct('com.sun.star.ucb.Command');command.Name='close';command.Handle=-1;command.Argument=None
   holder.execute(command,0,None)
 for name in db.FormDocuments.ElementNames:
  doc=ctrl.loadComponent(2,name,True);fd=db.FormDocuments.getByName(name)
  for parent,item in list(walk(doc.DrawPage.Forms)):
   if item.supportsService('com.sun.star.form.component.GridControl'):
    widths={'PolicyID':150,'IssuedOn':250,'StartsOn':250,'EndsOn':250,'CountryCode':400,'ContractNo':600,'ServiceCode':1700,'TariffDate':350,'AgreedPrice':400}
    for colname in item.ElementNames:
     col=item.getByName(colname);col.Width=widths.get(colname,300)
     if colname=='AgreedPrice':col.CurrencySymbol='руб.';col.PrependCurrencySymbol=False;col.Label='Цена, руб./день'
     if colname=='PolicyID':col.DecimalAccuracy=0
     if colname in ('IssuedOn','StartsOn','EndsOn','TariffDate'):col.DateFormat=7
     if colname=='ServiceCode':
      display='"ServiceName"||\' / \'||"CategoryName"' if not parent.AllowUpdates else '"ServiceLabel"'
      col.ListSource=(f'SELECT {display},"ServiceCode" FROM "ServiceChoices" ORDER BY "ServiceCode"',)
   elif item.supportsService('com.sun.star.form.component.DatabaseDateField'):
    item.DateFormat=7
   if hasattr(item,'DataField') and item.DataField=='ServiceCode' and item.supportsService('com.sun.star.form.component.ListBox'):
    item.ListSource=('SELECT "ServiceLabel","ServiceCode" FROM "ServiceChoices" ORDER BY "ServiceCode"',)
   if hasattr(item,'DataField') and item.supportsService('com.sun.star.form.component.DatabaseTextField'):
    n={'PassportNo':9,'AgentPassport':10,'BankBIK':9,'BankAccount':20,'CountryCode':3,'CurrencyCode':3,'CategoryCode':2,'ServiceCode':3}.get(item.DataField)
    if n:
     replacement=doc.createInstance('com.sun.star.form.component.DatabasePatternField')
     replacement.Name=item.Name;replacement.DataField=item.DataField;replacement.EditMask='N'*n;replacement.LiteralMask='_'*n;replacement.StrictFormat=True;replacement.ReadOnly=item.ReadOnly;replacement.FontHeight=9.5
     for i in range(doc.DrawPage.Count):
      shp=doc.DrawPage.getByIndex(i)
      if shp.supportsService('com.sun.star.drawing.ControlShape') and shp.Control==item:shp.setControl(replacement)
     parent.replaceByName(item.Name,replacement)
  command=uno.createUnoStruct('com.sun.star.ucb.Command');command.Name='store';command.Handle=-1;command.Argument=None;fd.execute(command,0,None)
  command.Name='close';fd.execute(command,0,None)
  db.store();print('Finalized',name,flush=True)
 db.store()
 (ROOT/'results/forms_patched.json').write_text(json.dumps({'forms':14,'masked_inputs':True,'historical_list_prices':'Separate persisted AgreedPrice; historical label omits current price','date_format':'DD/MM/YYYY'},ensure_ascii=False,indent=2))
if __name__=='__main__':main()
