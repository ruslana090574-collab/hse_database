"""Behavioral acceptance checks against the saved embedded Firebird database."""
from pathlib import Path
from decimal import Decimal,ROUND_HALF_UP
from datetime import date
import json
from uno_common import connect,P,url,rows,execute
from model import ROOT,DB_NAME,YEAR,QUERIES,POLICIES,services

def main():
 ctx,desktop=connect()
 db=desktop.loadComponentFromURL(url(ROOT/DB_NAME),'_blank',0,(P('Hidden',True),))
 cn=db.DataSource.getIsolatedConnection('','');cn.setAutoCommit(False)
 checks=[]
 def passed(name):checks.append({'check':name,'status':'PASS'});print('PASS',name,flush=True)
 def scalar(sql):return rows(cn,sql)[1][0][0]
 assert int(scalar('SELECT COUNT(*) FROM "Client"'))==8;passed('8 clients persisted in reopened .odb')
 assert int(scalar('SELECT COUNT(*) FROM "Policy"'))==13;passed('13 policies persisted in reopened .odb')
 assert int(scalar(QUERIES['03_Количество_полисов_текущий_год']))==8;passed('8 current-year policies')
 _,allcounts=rows(cn,QUERIES['05_Все_клиенты_полисы_текущий_год'])
 counts={r[0]:int(r[3]) for r in allcounts}
 assert len(counts)==8 and counts['710000006']==counts['710000007']==counts['710000008']==0
 passed('LEFT JOIN includes both clients with old policies only and clients with no policies')
 _,nopolicies=rows(cn,QUERIES['07_Клиенты_без_полисов_текущий_год'])
 assert {r[0] for r in nopolicies}=={'710000006','710000007','710000008'};passed('NOT EXISTS returns all 3 correct zero-current-year clients')
 sql=QUERIES['12_Агенты_лидеры_за_период'].replace(':DateStart',f"'{YEAR}-01-01'").replace(':DateEnd',f"'{YEAR}-12-31'")
 _,leaders=rows(cn,sql);assert {r[0] for r in leaders}=={'1','2'} and all(int(r[2])==3 for r in leaders)
 passed('Both agents tied for first place are returned; agent with no sales is not a leader')
 cols,actual=rows(cn,'SELECT "PolicyID","PremiumRUB" FROM "PolicySummary" ORDER BY "PolicyID"')
 price={r[0]:Decimal(str(r[3])) for r in services}
 factors={'276':Decimal('1.20'),'250':Decimal('1.30'),'840':Decimal('1.80'),'156':Decimal('1.10'),'036':Decimal('1.50')}
 expected={str(pid):(sum(price[code]+(2 if year==YEAR else 0) for code in chosen)*factors[country]*days).quantize(Decimal('.01'),rounding=ROUND_HALF_UP) for pid,ci,year,month,day,country,agent,days,chosen in POLICIES}
 assert all(Decimal(amount)==expected[pid] for pid,amount in actual);passed('All 13 premium amounts match independent Decimal calculations')
 before=rows(cn,'SELECT "PolicyID","PremiumRUB","CoverageRUB" FROM "PolicySummary" ORDER BY "PolicyID"')[1]
 execute(cn,'UPDATE "Country" SET "Factor"="Factor"+1')
 execute(cn,'UPDATE "ServicePrice" SET "BasePrice"="BasePrice"+100')
 execute(cn,'UPDATE "ExchangeRate" SET "Rate"="Rate"+50')
 assert rows(cn,'SELECT "PolicyID","PremiumRUB","CoverageRUB" FROM "PolicySummary" ORDER BY "PolicyID"')[1]==before
 cn.rollback();passed('Historical premiums and cover values stay unchanged after reference-data changes')
 def rejected(name,statements):
  failed=False
  try:
   for statement in statements:execute(cn,statement)
  except Exception:failed=True
  finally:cn.rollback()
  assert failed,name
  passed(name)
 rejected('Invalid country code rejected',["INSERT INTO \"Country\" VALUES ('A36','Test','036',100,1)"])
 rejected('Lower-case currency code rejected',["INSERT INTO \"Currency\" VALUES ('999','abc','Test')"])
 rejected('Empty currency name rejected',["INSERT INTO \"Currency\" VALUES ('999','ZZZ',' ')"])
 rejected('Duplicate primary key rejected',["INSERT INTO \"Gender\" VALUES ('1','мужской','M')"])
 rejected('Gender code/name/symbol mismatch rejected',["UPDATE \"Gender\" SET \"GenderSymbol\"='F' WHERE \"GenderCode\"='1'"])
 rejected('Negative service price rejected',["UPDATE \"ServicePrice\" SET \"BasePrice\"=-1 WHERE \"ServiceCode\"='001'"])
 rejected('Zero cover rejected',["UPDATE \"Country\" SET \"Coverage\"=0 WHERE \"CountryCode\"='276'"])
 rejected('Cyrillic client name rejected',["UPDATE \"Client\" SET \"FullName\"='ИВАНОВ' WHERE \"PassportNo\"='710000001'"])
 rejected('Passport with non-digit rejected',["UPDATE \"Client\" SET \"PassportNo\"='71000000A' WHERE \"PassportNo\"='710000007'"])
 rejected('Underage client rejected',[f"UPDATE \"Client\" SET \"BirthDate\"='{YEAR-10}-01-01' WHERE \"PassportNo\"='710000007'"])
 rejected('Expired passport rejected',["UPDATE \"Client\" SET \"ExpiryDate\"=CURRENT_DATE WHERE \"PassportNo\"='710000007'"])
 rejected('Future passport issue date rejected',["UPDATE \"Client\" SET \"IssueDate\"=DATEADD(1 DAY TO CURRENT_DATE) WHERE \"PassportNo\"='710000007'"])
 rejected('Orphan foreign key rejected',["UPDATE \"Client\" SET \"GenderCode\"='3' WHERE \"PassportNo\"='710000007'"])
 base='INSERT INTO "Policy" ("PassportNo","IssuedOn","StartsOn","EndsOn","CountryCode","ContractNo") VALUES '
 rejected('Policy shorter than seven days rejected',[base+f"('710000007','{YEAR}-08-01','{YEAR}-08-02','{YEAR}-08-08','276','AG-01')"])
 rejected('Policy start before issue date rejected',[base+f"('710000007','{YEAR}-08-02','{YEAR}-08-01','{YEAR}-08-10','276','AG-01')"])
 rejected('Passport validity covering entire trip enforced',[base+f"('710000007','{YEAR}-08-01','{YEAR+6}-08-02','{YEAR+6}-08-09','276','AG-01')"])
 rejected('Agent contract validity enforced',[base+f"('710000007','{YEAR-2}-08-01','{YEAR-2}-08-02','{YEAR-2}-08-10','276','AG-01')"])
 rejected('Duplicate service within policy rejected',["INSERT INTO \"PolicyService\" (\"PolicyID\",\"ServiceCode\") VALUES (6,'001')"])
 rejected('Deleting referenced client rejected',["DELETE FROM \"Client\" WHERE \"PassportNo\"='710000001'"])
 rejected('Changing date of policy with services rejected',[f"UPDATE \"Policy\" SET \"IssuedOn\"='{YEAR}-01-14' WHERE \"PolicyID\"=6"])
 # Birthday-based validation, separate from current-age validation.
 rejected('18th birthday must precede the policy issue date',[
  f"INSERT INTO \"Client\" VALUES ('999999999','TEST CLIENT','{YEAR-18}-01-01','1','{YEAR-5}-01-01','{YEAR+5}-01-01')",
  base+f"('999999999','{YEAR-1}-08-01','{YEAR-1}-08-02','{YEAR-1}-08-10','276','AG-01')"])
 # Cascading delete of a policy removes only its dependent service selections.
 execute(cn,'DELETE FROM "Policy" WHERE "PolicyID"=6')
 assert scalar('SELECT COUNT(*) FROM "PolicyService" WHERE "PolicyID"=6')=='0'
 cn.rollback();passed('Deleting a policy cascades to its services; rollback restores demo data')
 # Exercise saved query execution with Base's own parser and parameter binding.
 for name in QUERIES:
  rowset=ctx.ServiceManager.createInstanceWithContext('com.sun.star.sdb.RowSet',ctx)
  rowset.ActiveConnection=cn;rowset.CommandType=1;rowset.Command=name
  if ':DateStart' in QUERIES[name]:
   count=QUERIES[name].count(':DateStart')+QUERIES[name].count(':DateEnd')
   for i in range(count):
    dt=__import__('uno').createUnoStruct('com.sun.star.util.Date')
    dt.Year=YEAR;dt.Month=1 if i%2==0 else 12;dt.Day=1 if i%2==0 else 31
    rowset.setDate(i+1,dt)
  rowset.execute();rowset.last();n=rowset.getRow();rowset.dispose()
 passed('All 20 queries run through Base RowSet, including bound date parameters')
 cn.rollback();cn.close();db.close(True)
 (ROOT/'results'/'validation.json').write_text(json.dumps(checks,ensure_ascii=False,indent=2))
 print('ALL',len(checks),'CHECKS PASSED',flush=True)
if __name__=='__main__':main()
