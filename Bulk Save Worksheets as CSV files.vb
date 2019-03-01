Sub SaveWorksheetsAsCsv()

' This macro takes all the Worksheets within your active Workbook and saves it as separate files.
' Change parameters below to set export directory.

Dim WS As Excel.Worksheet
Dim SaveToDirectory As String
Dim CurrentWorkbook As String
Dim CurrentFormat As Long

CurrentWorkbook = ThisWorkbook.FullName
CurrentFormat = ThisWorkbook.FileFormat

' Update directory to save files
' Use backslashes (as opposed to Python or R)
SaveToDirectory = "C:\test\"

For Each WS In ThisWorkbook.Worksheets
    Sheets(WS.Name).Copy
    ActiveWorkbook.SaveAs Filename:=SaveToDirectory & ThisWorkbook.Name & "-" & WS.Name & ".csv", FileFormat:=xlCSV
    ActiveWorkbook.Close savechanges:=False
    ThisWorkbook.Activate
Next

Application.DisplayAlerts = False
ThisWorkbook.SaveAs Filename:=CurrentWorkbook, FileFormat:=CurrentFormat
Application.DisplayAlerts = True
' Temporarily turn alerts off to prevent the user being prompted
'  about overwriting the original file.

End Sub