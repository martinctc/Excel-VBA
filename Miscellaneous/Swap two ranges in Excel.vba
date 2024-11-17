Sub SwapTwoRanges()
' This macro swaps the values from two selected ranges in Excel, without transferring the formatting.  
' An InputBox will pop up twice to prompt you to enter the ranges.

Dim Rng1 As Range, Rng2 As Range
Dim arr1 As Variant, arr2 As Variant

xTitleId = "Range Swapper"
Set Rng1 = Application.Selection
Set Rng1 = Application.InputBox("Range1:", xTitleId, Rng1.Address, Type:=8)
Set Rng2 = Application.InputBox("Range2:", xTitleId, Type:=8)

Application.ScreenUpdating = False
arr1 = Rng1.Value
arr2 = Rng2.Value
Rng1.Value = arr2
Rng2.Value = arr1
Application.ScreenUpdating = True
End Sub