from studentMarksExcelAutomation.exceldata import process_wb
from win32com.client import Dispatch

# Calling the function
filename1 = "dataResultsMain.xlsx"
process_wb(filename1)

# Automatically opening manipulated Excel file
xl = Dispatch("Excel.Application")
xl.visible = True

wb = xl.Workbooks.Open(r'C:\Users\User\PycharmProjects\HelloWorld\studentMarksExcelAutomation\newdataResultsMain.xlsx')