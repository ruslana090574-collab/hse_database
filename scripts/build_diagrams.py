import re,subprocess,html
from model import ROOT,DDL,LABELS

def main():
 out=ROOT/'Документы'
 lines=['digraph DB {','graph [rankdir=LR, bgcolor="white", nodesep=0.4, ranksep=0.7, pad=0.3];','node [shape=plain,fontname="DejaVu Sans",fontsize=10];','edge [fontname="DejaVu Sans",fontsize=9,color="#56718c",arrowsize=0.8];']
 for sql in DDL:
  if not sql.startswith('CREATE TABLE'):continue
  table=re.search(r'CREATE TABLE "([^"]+)"',sql).group(1)
  fields=re.findall(r'^"([^"]+)" ([A-Z]+(?:\([^\n]*?\))?)',sql.split('(',1)[1],re.M)
  rows=''.join('<TR><TD ALIGN="LEFT">'+html.escape(f+' — '+LABELS.get(f,f))+'</TD><TD ALIGN="LEFT">'+html.escape(kind)+'</TD></TR>' for f,kind in fields)
  lines.append('"'+table+'" [label=<<TABLE BORDER="0" CELLBORDER="1" CELLSPACING="0" CELLPADDING="5"><TR><TD COLSPAN="2" BGCOLOR="#17365d"><FONT COLOR="white"><B>'+html.escape(table+' / '+LABELS.get(table,table))+'</B></FONT></TD></TR>'+rows+'</TABLE>>];')
  for columns,parent,parentcols in re.findall(r'FOREIGN KEY \((.*?)\) REFERENCES "([^"]+)" \((.*?)\)',sql):
   lines.append('"'+parent+'" -> "'+table+'" [arrowhead=crow,arrowtail=tee,dir=both,label="1 : N"];')
 lines.append('}')
 dot=out/'Схема_БД.dot';dot.write_text('\n'.join(lines))
 for ext in ['svg','png','pdf']:subprocess.run(['dot','-T'+ext,str(dot),'-o',str(out/('Схема_БД.'+ext))],check=True)
 alternative='''digraph Passports { graph [rankdir=LR,bgcolor="white",pad=.3]; node [shape=box,style="rounded,filled",fillcolor="#eef3f8",fontname="DejaVu Sans",fontsize=11]; edge [fontname="DejaVu Sans",fontsize=10];
ClientPerson [label="Клиент\nClientID (PK)\nФИО, дата рождения, код пола"];
TravelPassport [label="Заграничный паспорт\nPassportNo (PK)\nClientID (FK)\nдата выдачи, действителен до"];
PolicyV2 [label="Полис\nPolicyID (PK)\nPassportNo (FK)\nContractNo (FK)\nдаты, страна, зафиксированные условия"];
AgentPerson [label="Агент\nAgentID (PK)\nФИО, статус"];
AgentPassport [label="Паспорт агента\nAgentPassportID (PK)\nAgentID (FK)\nномер, выдан, срок действия"];
ContractV2 [label="Договор агента\nContractNo (PK)\nAgentPassportID (FK)\nсроки, комиссия, банковские реквизиты"];
ClientPerson -> TravelPassport [label="1 : N"]; TravelPassport -> PolicyV2 [label="1 : N"]; AgentPerson -> AgentPassport [label="1 : N"]; AgentPassport -> ContractV2 [label="1 : N"]; ContractV2 -> PolicyV2 [label="1 : N"]; }'''
 dot=out/'Несколько_паспортов.dot';dot.write_text(alternative)
 for ext in ['svg','png','pdf']:subprocess.run(['dot','-T'+ext,str(dot),'-o',str(out/('Несколько_паспортов.'+ext))],check=True)
 print('Created relational and multiple-passport diagrams in SVG/PNG/PDF.')
if __name__=='__main__':main()
