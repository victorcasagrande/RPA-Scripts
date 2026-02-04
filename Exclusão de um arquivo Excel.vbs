‘Exclusão de um arquivo Excel

Set obj = createobject(“Excel.Application”)   ‘Creating an Excel Object
obj.visible=True                                    ‘Making an Excel Object visible
Set obj1 = obj.Workbooks.open(“C:\newexcelfile.xls”)    ‘Opening an Excel file
Set obj2=obj1.Worksheets(“Sheet1”)    ‘Referring Sheet1 of excel file
obj2.Rows(“4:4”).Delete           ‘Deleting 4th row from Sheet1
obj1.Save()                                   ‘Saving the file with the changes
obj1.Close                                             ‘Closing a Workbook
obj.Quit                                                  ‘Exit from Excel Application
Set obj1=Nothing                                 ‘Releasing Workbook object
Set obj2 = Nothing                               ‘Releasing Worksheet object