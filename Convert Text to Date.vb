Sub DateValueConverter()

'Note: Automatically takes selection as Input and replaces with Output

Dim rng, cel As Range
Set rng = Application.Selection
For Each cel In rng
    cel.Value = DateValue(cel.Value)
Next
End Sub