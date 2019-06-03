Attribute VB_Name = "Module1"
Sub OrganizeTable()
'
' OrganizeTable Macro
'
' Keyboard Shortcut: Ctrl+o
'
    Range("B2:G27").Select
    ActiveSheet.Sort.SortFields.Clear
    ActiveSheet.Sort.SortFields.Add Key:= _
        Range("B2"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveSheet.Sort
        .SetRange Range("B2:G27")
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub

