import re
try:
    import xml.etree.cElementTree as ET
except:
    import xml.etree.ElementTree as ET
import glob

import os
import time

# import xlwt
import xlsxwriter

import argparse

from multiprocessing import Process,Pool,Lock

XMLPATH='report'

DEFAULT_STYLE={
        'font_size': 12,  
        'bold': False,  
        'font_color': 'black',  
        'align': 'left',  
        'valign':'vcenter',
        'font_name':'Courier New',
        'top': 2,  
        'left': 2,  
        'right': 2,  
        'bottom': 2  
}
TITLE=[
    (u'No.',8),
    ('IP',22),
    (u'Port',10),
    (u'Service',18),
    (u'Status',15),
]

def get_xml(filepath=XMLPATH):
    try:
        return map(lambda x:os.path.join(filepath,x),glob.glob1(filepath,'*.xml'))
    except:
        return []

def get_style(default=DEFAULT_STYLE,**kw):
    return  default.update(**kw)

def parseNmap(filename):
    try:
        tree=ET.parse(filename)
        root=tree.getroot()
    except Exception as e:

        print (e)
        return {}
    data_lst=[]
    for host in root.iter('host'):
        if host.find('status').get('state') == 'down':
            continue
        address=host.find('address').get('addr',None)
        # print address
        if not address:
            continue
        ports=[]
        for port in host.iter('port'):
            state=port.find('state').get('state','')
            port_num= port.get('portid',None)
            serv=port.find('service')
    
            serv= serv.get('name','') if serv is not None else ""
            # print serv
            ports.append([port_num,serv,state])
        data_lst.append({address:ports})
    # return {address:ports}
    return data_lst


def reportEXCEL(filename,datalst,title=TITLE,style=DEFAULT_STYLE,**kwargs):
    if not datalst:
        return ''
    if  os.path.exists(filename):
        print (u"%s The file already exists" % filename)
        path,name=os.path.split(filename)
        filename=os.path.splitext(name)[0]
        filename=filename+str(time.strftime("%Y%m%d%H%M%S",time.localtime()))+'.xlsx'
        filename=os.path.join(path,filename)
        print ('Data will save as new file named :%s ' % filename)

    book=xlsxwriter.Workbook(filename)
    title_style= style if not kwargs.get('title',None) else kwargs.get('title')

    row_hight=[20,16] if not kwargs.get('row_set',None) else kwargs.get('row_set')  
    sheet_name= 'sheet' if not kwargs.get('sheet_name',None) else kwargs.get('sheet_name')
    sheet=book.add_worksheet(sheet_name)

    row_hight=row_hight+(2000-len(row_hight))*[row_hight[-1]]
    for row , h in enumerate(row_hight):
        sheet.set_row(row,h)
    col_width=map(lambda x:x[1],title)
    for col , w in enumerate(col_width):
        sheet.set_column(col,col,w)
    title_style = book.add_format(title_style)
    for index,t in enumerate(title):

        sheet.write(0,index,t[0],title_style)

    row=1
    col=0
    style=book.add_format(style)
    index2=0
    for index,item in enumerate(datalst):
        # print item
        for ip,ports in item.items():
            port_num=len(ports)
            if not ports:
                continue
            index2=index2+1
            for  i,data in enumerate(ports):
                sheet.write(row,2,data[0],style)
                sheet.write(row,3,data[1],style)
                sheet.write(row,4,data[2],style)
                row = row + 1
            if row-port_num+1 != row:
                sheet.merge_range('B'+str(row-port_num+1)+':B'+str(row),ip,style)
                sheet.merge_range('A'+str(row-port_num+1)+':A'+str(row),index2,style)
            else:
                # print index2
                sheet.write(row-1,0,index2,style)
                sheet.write(row-1,1,ip,style)
    print  ('Reprot result of xml parser to file: %s' % filename)
    book.close()

def main(XMLPATH,REPORTFILENAME):

    data_lst=[]
    for xml in get_xml(XMLPATH):

        data=parseNmap(xml)
        if data:
            data_lst.extend(data)
            # print data


    reportEXCEL(REPORTFILENAME,data_lst)


if __name__ == '__main__':
    import sys
    if len(sys.argv)<3:
        print ('[!] Usage: parserXML.py XMLPATH [reportfilename]')
        print ('[!] Demo: parserXML.py  xmldir  result.xlsx')
    else:

        XMLPATH=sys.argv[1]
        REPORTFILENAME = sys.argv[2]
        print ('[-] set parser XML file dir: %s' % XMLPATH)
        print ('[-] set report Excel file name: %s' % REPORTFILENAME)

        if not os.path.exists(XMLPATH):
                print ("[!] '%s' path does not exists!" % XMLPATH)
                exit(1)
        main(XMLPATH,REPORTFILENAME)
    pass
