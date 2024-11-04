Sub OrdenarData()
'
    ThisWorkbook.Sheets("CONTROLEUTP").Select
    Range("A1:G1").Select
    Range("G1").Activate
    Selection.AutoFilter
    Range("E5").Select
    ActiveWorkbook.Worksheets("CONTROLEUTP").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("CONTROLEUTP").AutoFilter.Sort.SortFields.Add Key:= _
        Range("F1"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("CONTROLEUTP").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Selection.AutoFilter
End Sub

