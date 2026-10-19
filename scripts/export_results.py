import csv,json
from uno_common import *
from model import ROOT,DB_NAME,YEAR

def main():
 ctx,desktop=connect();db=desktop.loadComponentFromURL(url(ROOT/DB_NAME),'_blank',0,(P('Hidden',True),));cn=db.DataSource.getIsolatedConnection('','')
 results={};overrides={}
 for name in sorted(db.DataSource.QueryDefinitions.ElementNames):
  q=db.DataSource.QueryDefinitions.getByName(name);sql=q.Command;overrides[name]=sql
  executed=sql.replace(':DateStart',f"'{YEAR}-01-01'").replace(':DateEnd',f"'{YEAR}-12-31'")
  names,data=rows(cn,executed);results[name]={'columns':names,'rows':data}
  with (ROOT/'results'/f'{name}.csv').open('w',encoding='utf-8-sig',newline='') as stream:
   w=csv.writer(stream,delimiter=';');w.writerow(names);w.writerows(data)
 (ROOT/'results/queries_results.json').write_text(json.dumps(results,ensure_ascii=False,indent=2))
 (ROOT/'sql/query_overrides.json').write_text(json.dumps(overrides,ensure_ascii=False,indent=2))
 cn.close();db.store()
 print('Exported 20 query results and exact final SQL definitions.')
if __name__=='__main__':main()
