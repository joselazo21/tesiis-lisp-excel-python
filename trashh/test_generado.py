#!/usr/bin/env python3
import openpyxl

# DATA
mitabla_data = [["A", 1], ["B", 2]]

def generar():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "TestExcel"

    # Hoja1
    for r,d in enumerate(mitabla_data,1):
        for c,v in enumerate(d,1):
            ws.cell(r,c,v)

    return wb

def main():
    wb = generar()
    wb.save('TestExcel.xlsx');print('OK: TestExcel.xlsx')

if __name__=='__main__':main()
