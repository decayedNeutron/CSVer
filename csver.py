import win32com.client
import os
import sys
from pywintypes import com_error

FILE_1 = sys.argv[1]

PARENT_DIR = os.path.abspath(os.curdir)



fil = FILE_1.split('\\')[-1].split(".")[0]



XL_EXT = r'.xlsx'

PATH_XL = FILE_1
print("from : " + PATH_XL)

CSV_EXT = r'.csv'
PATH_CSV = os.path.join(PARENT_DIR,fil + '_CSVs\\')
print("to : " + PATH_CSV)

mode = 0o666
NEW_DIR = os.path.join(PARENT_DIR, fil + r'_CSVs')

excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = False

try:
    wb = excel.Workbooks.Open(PATH_XL)
    print('Making CSVs')
    if not os.path.isdir(NEW_DIR):
        os.mkdir(NEW_DIR)

    for i in range(wb.WorkSheets.Count):
        wb.WorkSheets(i+1).Select()
        name = wb.ActiveSheet.Name
        wb.ActiveSheet.SaveAs(PATH_CSV+name+CSV_EXT,6)
        print(PATH_CSV+name+CSV_EXT)

except Exception as e:
    print('Failed to create CSVs')
    print(e)
else:
    print('CSVs succesfully saved')
finally:
    wb.Close()
    excel.Quit()
