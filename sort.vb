Sub Sort()
'
' Sort Macro
'

'
    ActiveCell.Range("A1:F10").Select
    ActiveWorkbook.Worksheets("Business Trip Budget").ListObjects("Data").Sort. _
        SortFields.Clear
    ActiveWorkbook.Worksheets("Business Trip Budget").ListObjects("Data").Sort. _
        SortFields.Add2 Key:=ActiveCell.Offset(0, 4).Range("A1:A10"), SortOn:= _
        xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Business Trip Budget").ListObjects("Data").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub
