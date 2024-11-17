Sub comma_delimited()
'Updated - Results now shows on debugging window instead
Dim rng As Range
Dim InputRng As Range, OutRng As Range
Set InputRng = Application.Selection
Set InputRng = Application.InputBox("Range :", "Source", Default:=InputRng.Address, Type:=8)
'Set OutRng = Application.InputBox("Out put to (single cell):", "Output cell", Type:=8)
'Set OutRng = ActiveSheet.Range("E2")
outStr = ""
For Each rng In InputRng
    If outStr = "" Then
        outStr = rng.Value
    Else
        outStr = outStr & ", " & rng.Value
    End If
Next
'OutRng.Value = outStr
Debug.Print outStr
End Sub