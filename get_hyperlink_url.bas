Function ExtractFullHL(cell As Range) As String
    On Error Resume Next
    Dim fullAddress As String
    If cell.Hyperlinks.Count > 0 Then
        fullAddress = cell.Hyperlinks(1).Address
        If cell.Hyperlinks(1).SubAddress <> "" Then
            fullAddress = fullAddress & "#" & cell.Hyperlinks(1).SubAddress
        End If
    End If
    ExtractFullHL = fullAddress
    On Error GoTo 0
End Function
