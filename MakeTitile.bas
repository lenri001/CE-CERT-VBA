Attribute VB_Name = "Module3"
Sub MakeTitle()
Dim titttle As String ' Declares Variable as a String that will become a title
titttle = Sheets("Combined Charts").Range("E2").Value & Chr(13) & "Account # " & Sheets("Combined Charts").Range("A2").Value & Chr(13) & "1/1/2017 - 12/31/2018"
'The string has active values
Worksheets(1).ChartObjects(1).Activate
ActiveChart.HasTitle = True
ActiveChart.ChartTitle.Text = titttle 'Inserts String into title
ActiveWorkbook.Save
End Sub

