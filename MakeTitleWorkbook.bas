Attribute VB_Name = "Module7"
Sub MakeWorkbookTitle()
'
' MakeWorkbookTitle Macro
'

'
    Range("H1").Select ' Selects Loaction for the title of Max (kW)
    ActiveCell.FormulaR1C1 = "Max (kW)" ' Wrties the title for Max (kW)
    Range("H2").Select ' Selects Loaction for the value of Max (kW)
    ActiveCell.FormulaR1C1 = "=MAX(RC[-1]:R[70079]C[-1])" ' Writes the formula for Max (kW)
    Dim Deletefile As String
    Deletefile = ActiveWorkbook.FullName ' We wamt to delete the original file since a new one will be created
    Dim titttle As String ' Declares Variable as a String that will become a title
    titttle = "Max KW(" & Sheets("Combined Charts").Range("H2").Value & ") Address (" & Sheets("Combined Charts").Range("E2").Value & ") Account # (" & Sheets("Combined Charts").Range("A2").Value & ")" & ".xlsb"
    'The string has active values grabs the Cell Values that will be part of the title
    ActiveWorkbook.SaveAs Filename:=titttle, FileFormat:=50 ' Gets the string and creates the title (Also saves the workbook), fileformat = 50 saves it as a .xlsb file
    'Kill Deletefile ' This deltes the old file
    'ActiveWorkbook.Close 'Closes the workbook when we are done
End Sub

