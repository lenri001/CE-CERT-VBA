Attribute VB_Name = "Module1"
Sub Test()
Attribute Test.VB_ProcData.VB_Invoke_Func = " \n14"
'
'testing code by shaan
' Macro1 Macro
'

'
    Range("A1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
    ActiveChart.SetSourceData Source:=Range("Sheet1!$A$1:$CS$30")
    ActiveChart.Legend.Select
    Selection.Delete
    ActiveSheet.ChartObjects("Chart 1").Activate
End Sub
