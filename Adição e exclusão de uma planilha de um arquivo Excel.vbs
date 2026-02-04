‘Adição e exclusão de uma planilha de um arquivo Excel

Set obj = createobject(“Excel.Application”)   ‘Creating an Excel Object
obj.visible=True                                    ‘Making an Excel Object visible
Set obj1 = obj.Workbooks.open(“C:\newexcelfile.xls”)    ‘Opening an Excel file
Set obj2=obj1.sheets.Add  ‘Adding a new sheet in the excel file
obj2.name=”Sheet1”     ‘Assigning a name to the sheet created above
Set obj3= obj1.Sheets(“Sheet1”)  ‘Accessing Sheet1
obj3.Delete       ‘Deleting a sheet from an excel file
obj1.Close                                             ‘Closing a Workbook
obj.Quit                                                  ‘Exit from Excel Application
Set obj1=Nothing                                 ‘Releasing Workbook object
Set obj2 = Nothing                               ‘Releasing Worksheet object
Set obj3 = Nothing                              ‘Releasing Worksheet object
Set obj=Nothing                                   ‘Releasing Excel object