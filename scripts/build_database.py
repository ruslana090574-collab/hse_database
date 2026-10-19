"""Build the actual embedded .odb with constraints, saved queries and Base forms."""
from pathlib import Path
import json,csv
from uno_common import connect,P,url,rows,execute
from model import ROOT,DB_NAME,DDL,DATA,QUERIES,YEAR,save_sources

def main():
    save_sources()
    ctx,desktop=connect()
    sm=ctx.ServiceManager
    target=ROOT/DB_NAME
    if target.exists():
        raise RuntimeError('Refusing to overwrite existing database; move it aside explicitly before rebuilding')
    db=sm.createInstanceWithContext('com.sun.star.sdb.OfficeDatabaseDocument',ctx)
    db.DataSource.URL='sdbc:embedded:firebird'
    db.DataSource.Info=tuple(P(k,v) for k,v in {'ParameterNameSubstitution':True,'EnableOuterJoinEscape':False,'EscapeDateTime':False,'BooleanComparisonMode':2}.items())
    db.storeAsURL(url(target),())
    cn=db.DataSource.getConnection('','')
    for i,sql in enumerate(DDL,1):
        try:execute(cn,sql)
        except Exception:
            print('FAILED DDL',i,sql,flush=True);raise
    print('Schema created',flush=True)
    for i,sql in enumerate(DATA,1):
        try:execute(cn,sql)
        except Exception:
            print('FAILED DATA',i,sql,flush=True);raise
    cn.commit()
    queries=db.DataSource.QueryDefinitions
    for name,sql in QUERIES.items():
        query=queries.createInstance()
        query.Command=sql
        query.EscapeProcessing=not name.startswith(('19_','20_'))
        queries.insertByName(name,query)
    results={}
    for name,sql in QUERIES.items():
        sql=sql.replace(':Дата_с',f"'{YEAR}-01-01'").replace(':Дата_по',f"'{YEAR}-12-31'")
        names,data=rows(cn,sql)
        results[name]={'columns':names,'rows':data}
        with (ROOT/'results'/f'{name}.csv').open('w',encoding='utf-8-sig',newline='') as stream:
            writer=csv.writer(stream,delimiter=';');writer.writerow(names);writer.writerows(data)
        print(name,len(data),'rows',flush=True)
    (ROOT/'results'/'queries_results.json').write_text(json.dumps(results,ensure_ascii=False,indent=2))
    cn.commit();db.store()
    (ROOT/'results'/'build.json').write_text(json.dumps({'year':YEAR,'tables':12,'views':2,'queries':len(QUERIES),'database':DB_NAME},ensure_ascii=False,indent=2))
    print('DATABASE BUILT',target,flush=True)
    cn.close();db.store();db.close(True)
if __name__=='__main__':main()
