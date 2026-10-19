"""Create bound Base forms, lookups and master/detail subforms inside the .odb."""
import uno,json
from uno_common import connect,P,url,rows
from model import ROOT,DB_NAME,LABELS,QUERIES
from com.sun.star.awt import Point,Size

LOOKUPS={
'GenderCode':'SELECT "GenderName","GenderCode" FROM "Gender" ORDER BY "GenderCode"',
'PassportNo':'''SELECT "PassportNo"||' | '||"FullName"||' | '||CAST("BirthDate" AS VARCHAR(10)),"PassportNo" FROM "Client" ORDER BY "FullName"''',
'CountryCode':'SELECT "CountryName","CountryCode" FROM "Country" ORDER BY "CountryName"',
'CurrencyCode':'SELECT "AlphaCode"||\' — \'||"CurrencyName","CurrencyCode" FROM "Currency" ORDER BY "AlphaCode"',
'CategoryCode':'SELECT "CategoryName","CategoryCode" FROM "ServiceCategory" ORDER BY "CategoryCode"',
'ServiceCode':'SELECT "ServiceLabel","ServiceCode" FROM ('+QUERIES['16_Услуги_актуальные_для_списка']+') x',
'AgentID':'SELECT "AgentName","AgentID" FROM "Agent" ORDER BY "AgentID"',
'ContractNo':'''SELECT d."ContractNo"||' | '||a."AgentName",d."ContractNo" FROM "AgentContract" d JOIN "Agent" a ON a."AgentID"=d."AgentID" ORDER BY d."ContractNo"'''}
READONLY={'PolicyID','PolicyFactor','PolicyCoverage','PolicyCurrency','PolicyRate','RateDate','TariffDate','AgreedPrice'}


def shape(doc,model,x,y,w,h):
    shp=doc.createInstance('com.sun.star.drawing.ControlShape')
    shp.setSize(Size(w,h));shp.setPosition(Point(x,y));shp.setControl(model)
    doc.DrawPage.add(shp)
    return shp


def label(doc,form,name,text,x,y,w=6500):
    item=doc.createInstance('com.sun.star.form.component.FixedText')
    item.Name=name;item.Label=text;item.FontHeight=9.5
    form.insertByName(name,item);shape(doc,item,x,y,w,500)


def form_model(doc,container,name,table,cn,master=None,detail=None,readonly=False,insertonly=False):
    model=doc.createInstance('com.sun.star.form.component.DataForm')
    model.Name=name;model.Command=table;model.CommandType=0
    model.DataSourceName='';model.ActiveConnection=cn
    model.AllowUpdates=not readonly;model.AllowInserts=not readonly;model.AllowDeletes=not readonly
    model.IgnoreResult=insertonly
    if master:model.MasterFields=(master,);model.DetailFields=(detail,)
    container.insertByName(name,model)
    return model


def grid(doc,parent,name,table,columns,cn,x,y,w,h,master,detail,readonly=False):
    sub=form_model(doc,parent,name,table,cn,master,detail,readonly=readonly)
    grid=doc.createInstance('com.sun.star.form.component.GridControl');grid.Name=name+'_grid'
    grid.FontHeight=9
    sub.insertByName(grid.Name,grid)
    for field,kind,width in columns:
        col=grid.createColumn(kind)
        col.Name=field;col.DataField=field;col.Label=LABELS.get(field,field);col.Width=width
        if kind=='ListBox':
            col.ListSourceType=uno.Enum('com.sun.star.form.ListSourceType','SQL')
            col.ListSource=(LOOKUPS[field],);col.BoundColumn=1
        if field in READONLY or readonly:col.ReadOnly=True
        grid.insertByName(field,col)
    shape(doc,grid,x,y,w,h)
    return sub


def main():
    ctx,desktop=connect()
    db=desktop.loadComponentFromURL(url(ROOT/DB_NAME),'_blank',0,(P('Hidden',True),))
    cn=db.DataSource.getConnection('','')
    # Table order is independent of the collection's UI sort order.
    tables=['Client','Policy','Gender','Currency','ExchangeRate','Country','ServiceCategory','Service','ServicePrice','Agent','AgentContract','PolicyService']
    definitions=[('01_Клиенты_редактирование','Client','edit'),('02_Клиенты_регистрация','Client','insert'),('03_Клиенты_просмотр','Client','view'),('04_Оформление_полиса','Policy','edit')]
    definitions += [(f'{i:02d}_'+LABELS.get(t,t).replace(' ','_'),t,'edit') for i,t in enumerate(tables[2:],5)]
    made=[]
    for name,table,mode in definitions:
        if db.FormDocuments.hasByName(name):
            print('Already exists',name,flush=True);continue
        fd=db.FormDocuments.createInstanceWithArguments('com.sun.star.sdb.DocumentDefinition',(
            P('Name',name),P('ActiveConnection',cn),P('ClassID','8BC6B165-B1B2-4EDD-AA47-DAE2EE689DD6')))
        db.FormDocuments.insertByName(name,fd)
        doc=fd.getComponent()
        print("Opened form",name,flush=True)
        page=doc.StyleFamilies.getByName('PageStyles').getByName('Standard')
        page.Width=29700;page.Height=21000;page.IsLandscape=True
        page.LeftMargin=1200;page.RightMargin=1200;page.TopMargin=1000;page.BottomMargin=1000
        doc.Text.String='ЛР-1  /  СТРАХОВАЯ КОМПАНИЯ\n'+LABELS.get(table,table)+'  —  '+{'edit':'ввод и редактирование','insert':'регистрация новой записи','view':'только просмотр'}[mode]+'\nВсе сведения учебные и вымышленные. Справочники выбираются из раскрывающихся списков.'
        cursor=doc.Text.createTextCursor();cursor.gotoEnd(True);cursor.CharFontName='Liberation Sans';cursor.CharHeight=11
        model=form_model(doc,doc.DrawPage.Forms,'MainForm',table,cn,readonly=mode=='view',insertonly=mode=='insert')
        result=cn.createStatement().executeQuery('SELECT * FROM "'+table+'" WHERE 1=0')
        meta=result.getMetaData()
        fields=[(meta.getColumnName(i),meta.getColumnType(i),meta.getPrecision(i)) for i in range(1,meta.getColumnCount()+1)]
        result.close()
        for i,(field,kind,precision) in enumerate(fields):
            col=i%3;row=i//3;x=1300+col*9000;y=3400+row*1700
            label(doc,model,'label_'+field,LABELS.get(field,field),x,y,8400)
            lookup=field in LOOKUPS and not (field=={'Client':'PassportNo','Service':'ServiceCode','Country':'CountryCode','Currency':'CurrencyCode','ServiceCategory':'CategoryCode','Agent':'AgentID','AgentContract':'ContractNo','Gender':'GenderCode'}.get(table))
            if lookup:
                control=doc.createInstance('com.sun.star.form.component.ListBox')
                control.ListSourceType=uno.Enum('com.sun.star.form.ListSourceType','SQL');control.ListSource=(LOOKUPS[field],)
                control.BoundColumn=1;control.Dropdown=True
            elif kind==91:
                control=doc.createInstance('com.sun.star.form.component.DatabaseDateField');control.Dropdown=True
            elif kind in (2,3):
                control=doc.createInstance('com.sun.star.form.component.DatabaseNumericField');control.DecimalAccuracy=2
            elif kind in (4,5,-5):
                control=doc.createInstance('com.sun.star.form.component.DatabaseNumericField');control.DecimalAccuracy=0
            else:
                control=doc.createInstance('com.sun.star.form.component.DatabaseTextField');control.MaxTextLen=precision
            control.Name=field;control.DataField=field;control.FontHeight=9.5
            ro=mode=='view' or (table=='Policy' and field in READONLY) or (table=='Agent' and field=='AgentID') or (table=='PolicyService' and field in ('TariffDate','AgreedPrice'))
            if ro:control.ReadOnly=True
            model.insertByName(field,control);shape(doc,control,x,y+550,8400,800)
        bottom=3400+((len(fields)+2)//3)*1700+300
        if table=='Client' and mode!='insert':
            label(doc,model,'PoliciesLabel','Полисы выбранного клиента; услуги выбранного полиса показаны ниже',1300,bottom,26000)
            policies=grid(doc,model,'Policies','Policy',[
                ('PolicyID','NumericField',1500),('IssuedOn','DateField',2500),('StartsOn','DateField',2500),('EndsOn','DateField',2500),('CountryCode','ListBox',4000),('ContractNo','ListBox',6000)],cn,1300,bottom+650,26000,3300,'PassportNo','PassportNo',readonly=True)
            grid(doc,policies,'Services','PolicyService',[('ServiceCode','ListBox',17000),('TariffDate','DateField',3500),('AgreedPrice','CurrencyField',4000)],cn,1300,bottom+4300,26000,3000,'PolicyID','PolicyID',readonly=True)
        if table=='Policy':
            label(doc,model,'ServicesLabel','Услуги выбранного полиса: выберите услугу; дата тарифа и цена сохраняются автоматически',1300,bottom,26000)
            grid(doc,model,'Services','PolicyService',[('ServiceCode','ListBox',17500),('TariffDate','DateField',3500),('AgreedPrice','CurrencyField',4000)],cn,1300,bottom+650,26000,4000,'PolicyID','PolicyID')
        command=uno.createUnoStruct('com.sun.star.ucb.Command')
        command.Name='store';command.Handle=-1;command.Argument=None
        fd.execute(command,0,None)
        db.store();made.append(name);print('Created form',name,flush=True)
    cn.close();db.store();db.close(True)
    (ROOT/'results'/'forms.json').write_text(json.dumps(made,ensure_ascii=False,indent=2))
    print('FORMS CREATED',len(made),flush=True)
if __name__=='__main__':main()
