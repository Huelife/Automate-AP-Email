Option Explicit

Dim xlApp, xlBook, path
path = "C:\Users\emeas\Desktop\AP_email.xlsm"

Set xlApp = CreateObject("Excel.Application")
xlApp.Visible = False
Set xlBook = xlApp.Workbooks.Open(path, 0, True)

xlApp.Run "'" & path & "'!AP_emails.Compile_emails"

xlBook.Close
xlApp.Quit

Set xlApp = Nothing
Set xlBook = Nothing

WScript.Quit
