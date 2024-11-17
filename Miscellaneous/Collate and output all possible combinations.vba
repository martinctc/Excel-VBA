Sub Collate_and_output_all_combos()
'You will have a two-column data table (with headers) and as many rows as you want.
'It doesn't matter whether your data consists of text or number - no "calculation" is run directly on the values.
'You want to "multiply out" your data to get all possible combinations.
'Ordering: Column A for the variable to repeat multiple times (e.g. Alice, Alice, Alice, Bob, Bob, Bob)
'Ordering: Column B for the variable to display in sequence (e.g. 15, 20, 30, 15, 20, 30)
'Leave first row blank
Dim wb As Workbook
Dim ws As Worksheet
Dim k, p, i As Integer
Set wb = ActiveWorkbook
Set ws = wb.ActiveSheet
Application.ScreenUpdating = False
'The value 'k' is the total count of values in column A.
'The value 'p' is the total count of values in column B.
'The value 'i' shows the total number of combinations of your two variables - simple multiplification.
k = ws.Application.CountA(Range("A:A"))
p = ws.Application.CountA(Range("B:B"))
i = k * p
'Prints these values on the header row of your first three columns.
Range("A1").Value = k
Range("B1").Value = p
Range("C1").Value = i
'The commented column immediately below is an alternative method using formulas instead of VBA code - just ignore.
'Range("A1").Formula = "=COUNTA(A2:A9999)"
'Range("B1").Formula = "=COUNTA(B2:B9999)"
'Range("A1").Copy
'Range("A1").PasteSpecial (xlPasteValues)
'Range("B1").Copy
'Range("B1").PasteSpecial (xlPasteValues)
'Range("C1").Formula = "=A1*B1"
'i = Range("C1")

'The output would appear in Columns D and E.
'Please ensure you save your work first!
'Column references may be changed to suit your needs.
Range("D:D").ClearContents
Range("E:E").ClearContents
Range("D1").Value = "Col1"
Range("E1").Value = "Col2"
Range("D2").Formula = "=INDIRECT(""A""&IF(MOD(ROW(A1),$B$1)=0,QUOTIENT(ROW(A1),$B$1)+1,QUOTIENT(ROW(A1),$B$1)+2))"
Range("D2").Select
Range("D2").Copy
Range("D2").Resize(i, 1).PasteSpecial (xlPasteAll)
Range("E2").Formula = "=IF(MOD(ROW(B1),$B$1)=0,INDIRECT(""B""&$B$1+1),INDIRECT(""B""&MOD(ROW(B1),$B$1)+1))"
Range("E2").Select
Range("E2").Copy
Range("E2").Resize(i, 1).PasteSpecial (xlPasteAll)
Range("A1").Select
Application.ScreenUpdating = True
MsgBox "All done mate."
End Sub