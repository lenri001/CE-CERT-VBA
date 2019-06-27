Attribute VB_Name = "Module12"
Sub Blue()
'
' Blue Macro
'

'
    Dim intCount As Integer ' Declares a variable
    For intCount = 1 To 96 ' This will iterate the command 96 times
    ActiveChart.FullSeriesCollection(intCount).Select ' This defines which series you want to color
    With Selection.Format.Fill ' Fills in the line
        .Visible = msoTrue
        .ForeColor.RGB = RGB(0, 176, 240) ' color code for light blue
        .Transparency = 0
        .Solid
    End With
    With Selection.Format.Line ' Fillins in the border of the line
        .Visible = msoTrue
        .ForeColor.RGB = RGB(0, 176, 240)
        .Transparency = 0
    End With
    Next intCount ' Will go to the next series to color
    ActiveChart.ChartArea.Select
End Sub
