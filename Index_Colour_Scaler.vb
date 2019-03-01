Sub Index_Scaler()
Dim rng As Range

' Applies colour conditional formatting to indicate over and under indices

Dim lowtier, midtier, hightier As Long
' Change values below to set "tier" paramters
lowtier = 50
midtier = 100
hightier = 150

Set rng = Selection
    rng.FormatConditions.AddColorScale ColorScaleType:=3
    rng.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    
rng.FormatConditions(1).ColorScaleCriteria(1).Type = xlConditionValueNumber
rng.FormatConditions(1).ColorScaleCriteria(2).Type = xlConditionValueNumber
rng.FormatConditions(1).ColorScaleCriteria(3).Type = xlConditionValueNumber
rng.FormatConditions(1).ColorScaleCriteria(1).Value = lowtier
rng.FormatConditions(1).ColorScaleCriteria(2).Value = midtier
rng.FormatConditions(1).ColorScaleCriteria(3).Value = hightier
    
    With rng.FormatConditions(1).ColorScaleCriteria(1).FormatColor
        .Color = 7039480
        .TintAndShade = 0
    End With
    
    
    With rng.FormatConditions(1).ColorScaleCriteria(2).FormatColor
        .Color = 8711167
        .TintAndShade = 0
    End With
    
    With rng.FormatConditions(1).ColorScaleCriteria(3).FormatColor
        .Color = 8109667
        .TintAndShade = 0
    End With
End Sub
