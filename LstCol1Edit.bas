Attribute VB_Name = "Module6"
'Sub FinaAccount1()
'    Dim EndColumn As Integer
'    EndColumn =
'
'End Sub
'Sub ColumnL()
'    Dim TotalLength As Long
'    TotalLength = 1048576 ' Max Number of columns in Excel
'    Dim ColumnLengthValue As Long
'    ColumnLengthValue = TotalLength
'        For i = TotalLength To 1
'            If IsEmpty(Cells(i, 2).Value) = True Then
'               ColumnLengthValue = ColumnLengthValue - 1
'               Cells(i, 2).Value = 1
'            Else
'                Exit For
'            End If
'        Next i
'    Range("C3").Value = ColumnLengthValue
'End Sub
Sub Range_Find_Method()
'Finds the last non-blank cell on a sheet/range.

Dim lRow As Long
Dim lCol As Long
Dim rng As Range
    Set rng = Application.InputBox("Select Account Number Column", "Obtain Range Object", Type:=8) ' Prompts user to select the cell VBA will scroll through

    lRow = Cells.Find(What:="*", _
                    After:=Range("M1000"), _
                    LookAt:=xlPart, _
                    LookIn:=xlFormulas, _
                    SearchOrder:=xlByRows, _
                    SearchDirection:=xlPrevious, _
                    MatchCase:=False).Row

    MsgBox "Last Row: " & lRow

End Sub
'Sub RangeSelectionPrompt()
'    Dim rng As Range
'    Set rng = Application.InputBox("Select a range", "Obtain Range Object", Type:=8)
'
'    MsgBox "The cells selected were " & rng.Address
'End Sub
