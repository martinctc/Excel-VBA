Sub SaveWorksheetsAsXLSX_And_Encrypt()

'This macro allows you to save all Worksheets in a Workbook as separate XLSX files
'The Password argument allows you to encrypt the Workbooks

Dim WS As Excel.Worksheet
Dim SaveToDirectory As String
Dim CurrentWorkbook As String
Dim CurrentFormat As Long
CurrentWorkbook = ThisWorkbook.FullName
CurrentFormat = ThisWorkbook.FileFormat

' Specify the Directory that you would like your XLSX files saved in
' The file names of the Workbooks is a function of the names of the Worksheets
SaveToDirectory = "\folderA\subfolderA\"
For Each WS In ThisWorkbook.Worksheets
    Sheets(WS.Name).Copy
    ActiveWorkbook.SaveAs Filename:=SaveToDirectory & "Data Analysis 2019" & "- " & WS.Name & ".xlsx", FileFormat:=51, _
    Password:="D1FF1ULTP455W0RD"
    ActiveWorkbook.Close savechanges:=False
    ThisWorkbook.Activate
Next
Application.DisplayAlerts = False
ThisWorkbook.SaveAs Filename:=CurrentWorkbook, FileFormat:=CurrentFormat
Application.DisplayAlerts = True
' Temporarily turn alerts off to prevent the user being prompted
'  about overwriting the original file.
End Sub