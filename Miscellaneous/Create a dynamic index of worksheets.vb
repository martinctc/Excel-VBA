Private Sub Dynamic_Index()
'INSTRUCTIONS TO CREATE A DYNAMIC INDEX OF WORKSHEETS
'You have to first insert a new sheet called 'Index'.
'Right click on the sheet to 'View Code'
'Run code as normal
'Adapted from ExtendOffice.com
'https://www.extendoffice.com/documents/excel/2653-excel-dynamic-list-of-worksheet-names.html
   
    Dim xSheet As Worksheet
    Dim xRow As Integer
    Dim calcState As Long
    Dim scrUpdateState As Long
    Application.ScreenUpdating = False
    xRow = 1
    With Me
        .Columns(1).ClearContents
        .Cells(1, 1) = "INDEX"
        .Cells(1, 1).Name = "Index"
    End With
    For Each xSheet In Application.Worksheets
        If xSheet.Name <> Me.Name Then
            xRow = xRow + 1
            With xSheet
                .Range("A1").Name = "Start_" & xSheet.Index
                .Hyperlinks.Add anchor:=.Range("A1"), Address:="", _
                SubAddress:="Index", TextToDisplay:="Back to Index"
            End With
            Me.Hyperlinks.Add anchor:=Me.Cells(xRow, 1), Address:="", _
            SubAddress:="Start_" & xSheet.Index, TextToDisplay:=xSheet.Name
        End If
    Next
    Application.ScreenUpdating = True
End Sub
