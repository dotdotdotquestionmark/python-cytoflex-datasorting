import win32com.client as win32
import sys
# import openpyxl
# import pandas


def tempconnection():
    excelpath = r"C:\Users\syw20\Desktop\Excel Test\practice python excel.xlsm"
    excel = win32.Dispatch('Excel.Application')
    excel.Visible = True
    global wb, sheet, source
    wb = excel.Workbooks.Open(excelpath)
    sheet = wb.Worksheets(1)
    source = wb.Worksheets('rawdata')


def connection():
    # gathering command line arguements for address and dilutions
    excel = win32.Dispatch('Excel.Application')
    excel.Visible = True
    global bl_dilution, pos_dilution, neg_dilution, wb, sheet, source
    excelpath = sys.argv[1]
    bl_dilution = sys.argv[2]
    pos_dilution = sys.argv[3]
    neg_dilution = sys.argv[4]
    # setting workbook and worksheet
    wb = excel.Workbooks.Open(excelpath)
    sheet = wb.Worksheets(1)
    source = wb.Worksheets('rawdata')

    sheet.Cells(1, 1).Value = "connection successful"
    datagrab = source.Cells(4, 1).Value
    sheet.Cells(12, 1).Value = datagrab


def sample_scanner():
    print("running")
    sheet.Cells(15, 1).Value = "work"
    # define possible names
    baselinename = ["bl", "baseline", "lnw"]
    positivename = ["pos", "+", "positive"]
    row, column = 4, 0
    samples = []
    while True:

        sample = source.Cells(row, column).Value
        if sample is None:
            print("end scan")
            break
        elif any(bl in sample for bl in baselinename):
            baselinerow = row
            print(baselinerow)
            row += 1
        elif any(pos in sample for pos in positivename):
            samples.append(sample.rsplit(' ', 1)[0])
            row += 1
        else:
            row += 1

    print(samples)


def main():
    # tempconnection()
    # sample_scanner()
    connection()


if __name__ == "__main__":
    main()
