‘Lendo / abrindo um arquivo do Excel usando o objeto do Excel

Set obj = createobject(“Excel.Application”)   ‘Creating an Excel Object
obj.visible=True                                    ‘Making an Excel Object visible
Set obj1 = obj.Workbooks.open(“C:\newexcelfile.xls”)    ‘Opening an Excel file
Set obj2=obj1.Worksheets(“Sheet1”)    ‘Referring Sheet1 of excel file
Msgbox obj2.Cells(2,2).Value  ‘Value from the specified cell will be read and shown
obj1.Close                                             ‘Closing a Workbook
obj.Quit                                                  ‘Exit from Excel Application
Set obj1=Nothing                                 ‘Releasing Workbook object
Set obj2 = Nothing                               ‘Releasing Worksheet object
Set obj=Nothing                                   ‘Releasing Excel object