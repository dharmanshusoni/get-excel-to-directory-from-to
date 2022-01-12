from sys import path
import xlrd
import os
from pathlib import Path
import glob
import shutil

loc = ("D:\Live Projects\prodcom\ToolbankDataExport.xlsx")
newpath = Path("D:/Live Projects New/") 

wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(0)
sheet.cell_value(0, 0)

for i in range(sheet.nrows):
    try:
        if i > 0:
            if(sheet.cell_value(i, 51) and sheet.cell_value(i, 52)) and sheet.cell_value(i, 52).strip() != 42:
                src_dir = Path(sheet.cell_value(i, 51))
                dst_dir = Path(sheet.cell_value(i, 52))
                if (not os.path.exists(dst_dir)) and (os.path.isfile(src_dir)):
                    os.makedirs(dst_dir)
                    for jpgfile in glob.iglob(os.path.join(src_dir)):
                        shutil.copy(jpgfile, dst_dir)
    except :
        print ("An error occurred")
        print (sheet.cell_value(i, 51))
        print (sheet.cell_value(i, 52))
