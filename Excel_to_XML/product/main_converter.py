from openpyxl import load_workbook
from yattag import Doc, indent

wb = load_workbook("data_simplified.xlsx")
ws = wb.worksheets[0]

# Create Yattag doc, tag and text objects
doc, tag, text = Doc().tagtext()

xml_header = '<?xml version="1.0" encoding="UTF-8"?>'
xml_schema = '<xs:schema xmlns:xs="http://www.w3.org/2001/XMLSchema"></xs:schema>'

doc.asis(xml_header)
doc.asis(xml_schema)

with tag('Các Bệnh Nhân'):
    # Use ws.max_row for all rows
    for col in ws.iter_cols(min_col=2, max_col=3, min_row=2, max_row=41):
        col = [cell.value for cell in col]
        count = 0
        with tag("Bệnh Nhân"):
            with tag("MA_LK"):
                text(col[count])
                count += 1
            with tag("STT"):
                text(col[count])
                count += 1
            with tag("MA_BN"):
                text(col[count])
                count += 1
            with tag("HO_TEN"):
                text(col[count])
                count += 1
            with tag("NGAY_SINH"):
                text(col[count])
                count += 1
            with tag("GIOI_TINH"):
                text(col[count])
                count += 1
            with tag("DIA_CHI"):
                text(col[count])  
                count += 1
            with tag("MA_THE"):
                text(col[count]) 
                count += 1
            with tag("MA_DKBD"):
                text(col[count])  
                count += 1 
            with tag("GT_THE_TU"):
                text(col[count])
                count += 1
            with tag("GT_THE_DEN"):
                text(col[count])
                count += 1
            with tag("MIEN_CUNG_CT"):
                text(col[count])
                count += 1
            with tag("TEN_BENH"):
                text(col[count])
                count += 1
            with tag("MA_BENH"):
                text(col[count])
                count += 1
            with tag("MA_BENHKHAC"):
                text(col[count])
                count += 1
            with tag("MA_LYDO_VVIEN"):
                text(col[count])
                count += 1
            with tag("MA_NOI_CHUYEN"):
                text(col[count])
                count += 1
            with tag("MA_TAI_NAN"):
                text(col[count])
                count += 1
            with tag("NGAY_VAO"):
                text(col[count])
                count += 1
            with tag("NGAY_RA"):
                text(col[count])
                count += 1
            with tag("SO_NGAY_DTRI"):
                text(col[count])
                count += 1
            with tag("KET_QUA_DTRI"):
                text(col[count])
                count += 1
            with tag("TINH_TRANG_RV"):
                text(col[count])
                count += 1
            with tag("NGAY_TTOAN"):
                text(col[count])
                count += 1
            with tag("T_THUOC"):
                text(col[count])
                count += 1
            with tag("T_VTYT"):
                text(col[count])
                count += 1
            with tag("T_TONGCHI"):
                text(col[count])
                count += 1
            with tag("T_BNTT"):
                text(col[count])
                count += 1
            with tag("T_BNCCT"):
                text(col[count])
                count += 1
            with tag("T_BHTT"):
                text(col[count])
                count += 1
            with tag("T_NGUONKHAC"):
                text(col[count])
                count += 1
            with tag("T_NGOAIDS"):
                text(col[count])
                count += 1
            with tag("NAM_QT"):
                text(col[count])
                count += 1
            with tag("THANG_QT"):
                text(col[count])
                count += 1
            with tag("MA_LOAI_KCB"):
                text(col[count])
                count += 1
            with tag("MA_KHOA"):
                text(col[count])
                count += 1
            with tag("MA_CSKCB"):
                text(col[count])
                count += 1
            with tag("MA_KHUVUC"):
                text(col[count])
                count += 1
            with tag("MA_PTTT_QT"):
                text(col[count])
                count += 1
            with tag("CAN_NANG"):
                text(col[count])
                count += 1

result = indent(
    doc.getvalue(),
    #indentation = '    ',
    indent_text = True
)

with open("patient_names.xml", "w") as f:
    f.write(result)