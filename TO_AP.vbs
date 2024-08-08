Option Explicit

Dim xlApp, xlBook, path
path = "C:\Users\emeas\Desktop\MacroWB.xlsm"

Set xlApp = CreateObject("Excel.Application")
xlApp.Visible = False
Set xlBook = xlApp.Workbooks.Open(path, 0, True)

xlApp.Run "'" & path & "'!AP_email.AP_email"

xlBook.Close
xlApp.Quit

Set xlApp = Nothing
Set xlBook = Nothing

WScript.Quit
