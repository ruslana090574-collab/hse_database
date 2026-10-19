"""Capture actual LibreOffice windows; no synthetic UI images are generated."""
import os,time,json
os.environ.setdefault('DISPLAY',':99')
from PIL import ImageGrab
from uno_common import *
from model import ROOT,DB_NAME

def show(doc):
 controller=doc.CurrentController if hasattr(doc,'CurrentController') else doc
 w=controller.Frame.ContainerWindow;w.setVisible(True);w.setPosSize(0,0,1440,1000,15)
 try:w.toFront()
 except Exception:pass
 time.sleep(.8)

def shot(name):
 ImageGrab.grab().save(str(ROOT/'screenshots'/name));print('Captured',name,flush=True)

def main():
 c,desk=connect();db=desk.loadComponentFromURL(url(ROOT/DB_NAME),'_blank',0,(P('Hidden',False),));ctrl=db.CurrentController;ctrl.connect()
 show(db);shot('01_База_данных.png')
 f=ctrl.loadComponent(2,'03_Клиенты_просмотр',False);show(f)
 assert f.DrawPage.Forms.getByName('MainForm').getByName('FullName').Text=='IVANOV IVAN IVANOVICH'
 shot('02_Клиент_полисы_услуги.png')
 p=ctrl.loadComponent(2,'04_Оформление_полиса',False);show(p);shot('05_Оформление_полиса.png')
 q=ctrl.loadComponent(1,'05_Все_клиенты_полисы_текущий_год',False);show(q);shot('06_Запрос_клиенты_ноль.png')
 t=ctrl.loadComponent(0,'Client',True);show(t);shot('04_Конструктор_таблицы.png')
 report=desk.loadComponentFromURL(url(ROOT/'Документы/Продажи_прошлый_год_Base.odt'),'_blank',0,(P('Hidden',False),));show(report)
 report.CurrentController.ViewSettings.ZoomValue=85;time.sleep(.5);shot('07_Отчёт_с_итогами.png')
 calc=desk.loadComponentFromURL(url(ROOT/'Документы/Анализ_полисов.ods'),'_blank',0,(P('Hidden',False),));calc.CurrentController.setActiveSheet(calc.Sheets.getByName('Агенты'));show(calc);shot('08_Calc_анализ.png')
 policy=desk.loadComponentFromURL(url(ROOT/'Документы/Полисы_слияние/Полис_0.odt'),'_blank',0,(P('Hidden',False),));show(policy);shot('09_Writer_полис.png')
 (ROOT/'results/ui_smoke.json').write_text(json.dumps({'actual_gui_capture':True,'client_form_loaded':True,'client_rows':f.DrawPage.Forms.getByName('MainForm').RowCount,'screenshots':8},ensure_ascii=False,indent=2))
if __name__=='__main__':main()
