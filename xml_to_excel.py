from openpyxl import Workbook
import os, datetime
from bs4 import BeautifulSoup
import numpy as np

_infinite = np.inf

def readFile(filename):
    if not os.path.exists(filename): 
        print("Cannot find .xml file!")
        os._exit(0)
        return
    with open(filename,'r') as f:
        data = f.read()
    Bs_data = BeautifulSoup(data, "xml")

    bn_details = Bs_data.find_all('node')
    benh_nhan = Bs_data.find_all('Benh_Nhan')
    
    mdlist = []
    temp = []
    for _node in bn_details:
        temp.append(_node.get('id'))
    mdlist.append(temp)

    for _benh_nhan in benh_nhan:
        temp = []
        for _node in bn_details:
            node_text = check_data(_class=_node.get('class') , _type=_node.get('type') , _id=_node.get('id') , _data=_node.get_text())
            temp.append(node_text)
        mdlist.append(temp)
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

def check_data(_class=None, _type=None, _id=None, _data=None):
    if _data == "NODATA":
        return None
    elif _class == "number":
        if ("int" in _type) is True:
            if len(_data) <= int(_type[4:]):
                return int(_data)
            else:
                return None
        elif ("float" in _type) is True:
            stt_float = _type.find("float_")
            stt_decimal = _type.find("_decimal_")
            if len(_data) <= int(_type[6:stt_decimal]):
                return round(float(_data), int(_type[stt_decimal+9:]))
    elif _class == "code":
        tmp_length = 0
        if _type[7:] == "n":
            tmp_length = _infinite
        else:
            tmp_length = int(_type[7:])
        if len(_data) <= tmp_length:
            return _data
        else:
            return None
    elif _class == "time":
        if _type == "year_month_day":
            tmp_year = int(_data[0:4])
            tmp_month = int(_data[4:6])
            tmp_day = int(_data[6:8])
            print(tmp_year, tmp_month, tmp_day)
            return datetime.datetime(tmp_year, tmp_month, tmp_day)
        elif _type == "year_month_day_hour_minute":
            tmp_year = int(_data[0:4])
            tmp_month = int(_data[4:6])
            tmp_day = int(_data[6:8])
            tmp_hour = int(_data[8:10])
            tmp_minute = int(_data[10:12])
            return datetime.datetime(tmp_year, tmp_month, tmp_day, tmp_hour, tmp_minute)
        elif _type == "year":
            return int(_data)
        elif _type == "month":
            return int(_data)
    elif _class == "detail":
        if _type.find("selection") != -1:
            if _id == "GIOI_TINH":
                if _data == "1":
                    return "Nam"
                elif _data == "2":
                    return "Nữ"
                elif  _data == "3":
                    return "Chưa xác định"
            elif _id == "MA_LYDO_VVIEN":
                if _data == "1":
                    return "Đúng tuyến"
                elif _data == "2":
                    return "Cấp cứu"
                elif  _data == "3":
                    return "Trái tuyến"
                elif  _data == "4":
                    return "Thông tuyến"
            elif _id == "KET_QUA_DTRI":
                if _data == "1":
                    return "Khỏi"
                elif _data == "2":
                    return "Đỡ"
                elif  _data == "3":
                    return "Không thay đổi"
                elif  _data == "4":
                    return "Nặng hơn"
                elif  _data == "5":
                    return "Tử vong"
            elif _id == "TINH_TRANG_RV":
                if _data == "1":
                    return "Ra viện"
                elif _data == "2":
                    return "Chuyển viện"
                elif  _data == "3":
                    return "Trốn viện"
                elif  _data == "4":
                    return "Xin ra viện"
            elif _id == "MA_LOAI_KCB":
                if _data == "1":
                    return "Khám bệnh"
                elif _data == "2":
                    return "Điều trị ngoại trú"
                elif  _data == "3":
                    return "Điều trị nội trú"
            elif _id == "MA_KHUVUC":
                if _data == "1":
                    return "K1"
                elif _data == "2":
                    return "K2"
                elif  _data == "3":
                    return "K3"
    return None

result = readFile("output/patient_names.xml")
if result:
    to_Excel(result)