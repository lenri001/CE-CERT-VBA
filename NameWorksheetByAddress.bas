Attribute VB_Name = "Module15"
Sub NameWorksheetByAdresss()
    Dim titttle As String ' Declares Variable as a String that will become a title
    titttle = Range("D2").Value & " Account # " & Range("F2").Value
    ActiveSheet.Name = Range("D2").Value 'titttle
    ActiveWorkbook.Save
End Sub
