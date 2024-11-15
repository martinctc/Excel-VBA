' Returns the hyperlink from a text string that is hyperlinked
Function GetHyperlink(rng As Range) As String
    If rng.Hyperlinks.Count > 0 Then
        GetHyperlink = rng.Hyperlinks(1).Address
    Else
        GetHyperlink = ""
    End If
End Function
