from openpyxl import Workbook
import os, datetime
from bs4 import BeautifulSoup

def readFile(filename):
    mdlist = []
    sub_mdlist = []
    if not os.path.exists(filename): 
        print("Cannot find .xml file!")
        os._exit(0)
        return
    with open(filename,'r') as f:
        data = f.read()
    Bs_data = BeautifulSoup(data, "xml")

    for _node in Bs_data.find("patient").findChildren(recursive=True):
        if _node.get('value') is not None:
            sub_mdlist.append(_node.parent.name.upper()+'_'+_node.name.upper())
    mdlist.append(sub_mdlist)

    for _patient in Bs_data.find_all("patient"):
        sub_mdlist = []
        for _node in _patient.findChildren(recursive=True):
            if _node.get('value') is not None:
                sub_mdlist.append(_node.get('value'))
        mdlist.append(sub_mdlist)
    return mdlist

def to_Excel(mdlist):
    wb = Workbook()
    ws = wb.active
    for i,row in enumerate(mdlist):
        for j,value in enumerate(row):
            ws.cell(row=i+1, column=j+1).value = value
    newfilename = os.path.abspath("./output/data_xml_to_excel.xlsx")
    wb.save(newfilename)
    print("Process completed")
    return

result = readFile("output/patient1.xml")
if result:
    to_Excel(result)