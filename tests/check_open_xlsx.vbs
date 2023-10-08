'Check open XLSX file
On Error Resume Next 'This line is required
Dim excel
Set excel = CreateObject("Excel.Application")
'oExcel.DisplayAlerts = False
Set book = excel.Workbooks.Open(Wscript.Arguments.Item(0), False, True)
book.Close
excel.Quit

WScript.Quit Err.Number
