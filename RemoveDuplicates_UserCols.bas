Sub RemDuplicate_UserCols()

' Get user input "MyUserInput" of # columns to work on,
' loop through column 1:MyUserInput, and remove duplicates & sort ascending

Dim i As Integer

MyUserInput = InputBox("Enter # Columns")
CheckGoAhead = InputBox("Are you sure you want to remove duplicates? Cannot undo (1: Yes, 2: No)")

If CheckGoAhead = 1 Then
    For i = 1 To CInt(MyUserInput)
        Application.CutCopyMode = False
        ActiveSheet.Range(Cells(1, i), Cells(65536, i)).RemoveDuplicates Columns:=1, Header:=xlNo
        ActiveSheet.Sort.SortFields.Clear
        ActiveSheet.Sort.SortFields.Add Key:=Range(Cells(1, i), Cells(65536, i)), _
            SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        With ActiveSheet.Sort
            .SetRange Range(Cells(1, i), Cells(65536, i))
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
    Next i

    MsgBox "Remove Duplicates ran on " & MyUserInput & " columns"

Else
    MsgBox "Macro not run"
End If

End Sub
