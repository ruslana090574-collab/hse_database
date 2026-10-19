"""Helpers for LibreOffice's bundled/system Python with the UNO module."""
import uno,time,json
from pathlib import Path
from com.sun.star.beans import PropertyValue

def P(name,value):
    prop=PropertyValue();prop.Name=name;prop.Value=value;return prop

def connect(port=2002):
    local=uno.getComponentContext()
    resolver=local.ServiceManager.createInstanceWithContext('com.sun.star.bridge.UnoUrlResolver',local)
    for _ in range(300):
        try:
            ctx=resolver.resolve(f'uno:socket,host=127.0.0.1,port={port};urp;StarOffice.ComponentContext')
            return ctx,ctx.ServiceManager.createInstanceWithContext('com.sun.star.frame.Desktop',ctx)
        except Exception:time.sleep(.1)
    raise RuntimeError(f'Start LibreOffice with --accept=socket,host=127.0.0.1,port={port};urp;StarOffice.ComponentContext')

def url(path):return uno.systemPathToFileUrl(str(Path(path).resolve()))

def rows(connection,sql):
    result=connection.createStatement().executeQuery(sql)
    meta=result.getMetaData();names=[meta.getColumnLabel(i) for i in range(1,meta.getColumnCount()+1)]
    data=[]
    while result.next():
        row=[]
        for i in range(1,len(names)+1):
            value=result.getString(i).rstrip()
            if result.wasNull():value=None
            row.append(value)
        data.append(row)
    result.close()
    return names,data

def execute(connection,sql):
    statement=connection.createStatement()
    try:statement.executeUpdate(sql)
    finally:statement.close()
