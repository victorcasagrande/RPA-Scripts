‘Copiar e colar dados de um arquivo Excel para outro arquivo Excel
Set obj = createobject(“Excel.Application”)   ‘Creating an Excel Object
obj.visible=True                                    ‘Making an Excel Object visible
Set obj1 = obj.Workbooks.open(“C:\newexcelfile.xls”)    ‘Opening an Excel file1
Set obj2 = obj.Workbooks.open(“C:\newexcelfile1.xls”)    ‘Opening an Excel file2
obj1.Worksheets(“Sheet1”).usedrange.copy  ‘Copying from an Excel File1
obj2.Worksheets(“Sheet1”).usedrange.pastespecial  ‘Pasting in Excel File2
obj1.Save                                              ‘ Saving Workbook1
obj2.Save                                              ‘Saving Workbook2
obj1.Close                                             ‘Closing a Workbook
obj.Quit                                                 ‘Exit from Excel Application
Set obj1=Nothing                                ‘Releasing Workbook1 object
Set obj2 = Nothing                              ‘Releasing Workbook2 object
Set obj=Nothing                                  ‘Releasing Excel object