Attribute VB_Name = "Module4"
Sub HourToWatt()
'
' HourToWatt Macro
'
' Keyboard Shortcut: Ctrl+h
'
    Range("G1").Select
    Selection.ClearContents
    Range("G1").Select
    ActiveCell.FormulaR1C1 = "kW"
    Range("G2").Select
    ActiveCell.FormulaR1C1 = "=RC[-4]*4"
    Range("G2").Select
    Selection.AutoFill Destination:=Range("G2:G70081"), Type:=xlFillDefault
    Range("G2:G70081").Select
    ActiveWindow.SmallScroll Down:=3
    ActiveWindow.ScrollRow = 69955
    ActiveWindow.ScrollRow = 69861
    ActiveWindow.ScrollRow = 69767
    ActiveWindow.ScrollRow = 69673
    ActiveWindow.ScrollRow = 69579
    ActiveWindow.ScrollRow = 69108
    ActiveWindow.ScrollRow = 68920
    ActiveWindow.ScrollRow = 68637
    ActiveWindow.ScrollRow = 67696
    ActiveWindow.ScrollRow = 65719
    ActiveWindow.ScrollRow = 61011
    ActiveWindow.ScrollRow = 59410
    ActiveWindow.ScrollRow = 58469
    ActiveWindow.ScrollRow = 52631
    ActiveWindow.ScrollRow = 51690
    ActiveWindow.ScrollRow = 49524
    ActiveWindow.ScrollRow = 45005
    ActiveWindow.ScrollRow = 42463
    ActiveWindow.ScrollRow = 41239
    ActiveWindow.ScrollRow = 34178
    ActiveWindow.ScrollRow = 33236
    ActiveWindow.ScrollRow = 31541
    ActiveWindow.ScrollRow = 27869
    ActiveWindow.ScrollRow = 27022
    ActiveWindow.ScrollRow = 21467
    ActiveWindow.ScrollRow = 20902
    ActiveWindow.ScrollRow = 17136
    ActiveWindow.ScrollRow = 16665
    ActiveWindow.ScrollRow = 16100
    ActiveWindow.ScrollRow = 15065
    ActiveWindow.ScrollRow = 14406
    ActiveWindow.ScrollRow = 13935
    ActiveWindow.ScrollRow = 10546
    ActiveWindow.ScrollRow = 10075
    ActiveWindow.ScrollRow = 8474
    ActiveWindow.ScrollRow = 6309
    ActiveWindow.ScrollRow = 5555
    ActiveWindow.ScrollRow = 4614
    ActiveWindow.ScrollRow = 1695
    ActiveWindow.ScrollRow = 660
    ActiveWindow.ScrollRow = 1
    ActiveSheet.ChartObjects("Chart 1").Activate
    ActiveChart.Axes(xlValue).AxisTitle.Select
    ActiveChart.Axes(xlValue, xlPrimary).AxisTitle.Text = "kW"
    Selection.Format.TextFrame2.TextRange.Characters.Text = "kW"
    With Selection.Format.TextFrame2.TextRange.Characters(1, 2).ParagraphFormat
        .TextDirection = msoTextDirectionLeftToRight
        .Alignment = msoAlignCenter
    End With
    With Selection.Format.TextFrame2.TextRange.Characters(1, 2).Font
        .BaselineOffset = 0
        .Bold = msoFalse
        .NameComplexScript = "+mn-cs"
        .NameFarEast = "+mn-ea"
        .Fill.Visible = msoTrue
        .Fill.ForeColor.RGB = RGB(89, 89, 89)
        .Fill.Transparency = 0
        .Fill.Solid
        .Size = 10
        .Italic = msoFalse
        .Kerning = 12
        .Name = "+mn-lt"
        .UnderlineStyle = msoNoUnderline
        .Strike = msoNoStrike
    End With
    ActiveChart.ChartArea.Select
    ActiveChart.SeriesCollection.NewSeries
    ActiveChart.FullSeriesCollection(2).Name = "='Combined Charts'!$G$1"
    ActiveChart.FullSeriesCollection(2).Values = "='Combined Charts'!$G$2:$G$70081"
    ActiveChart.FullSeriesCollection(1).Delete
    ActiveChart.FullSeriesCollection(1).XValues = _
        "='Combined Charts'!$B$2:$B$70081"
    ActiveWindow.SmallScroll Down:=-42
    ActiveWindow.ScrollRow = 70044
    ActiveWindow.ScrollRow = 69855
    ActiveWindow.ScrollRow = 69761
    ActiveWindow.ScrollRow = 65148
    ActiveWindow.ScrollRow = 63642
    ActiveWindow.ScrollRow = 59405
    ActiveWindow.ScrollRow = 57899
    ActiveWindow.ScrollRow = 51026
    ActiveWindow.ScrollRow = 48108
    ActiveWindow.ScrollRow = 43118
    ActiveWindow.ScrollRow = 40200
    ActiveWindow.ScrollRow = 32951
    ActiveWindow.ScrollRow = 30691
    ActiveWindow.ScrollRow = 23819
    ActiveWindow.ScrollRow = 16758
    ActiveWindow.ScrollRow = 12992
    ActiveWindow.ScrollRow = 11109
    ActiveWindow.ScrollRow = 7438
    ActiveWindow.ScrollRow = 6214
    ActiveWindow.ScrollRow = 3296
    ActiveWindow.ScrollRow = 2825
    ActiveWindow.ScrollRow = 1130
    ActiveWindow.ScrollRow = 754
    ActiveWindow.ScrollRow = 1
    ActiveChart.FullSeriesCollection(1).Select
    With Selection.Format.Fill
        .Visible = msoTrue
        .ForeColor.RGB = RGB(0, 176, 240)
        .Transparency = 0
        .Solid
    End With
    With Selection.Format.Line
        .Visible = msoTrue
        .ForeColor.RGB = RGB(0, 176, 240)
        .Transparency = 0
    End With
    Dim titttle As String ' Declares Variable as a String that will become a title
titttle = Sheets("Combined Charts").Range("E2").Value & Chr(13) & "Account # " & Sheets("Combined Charts").Range("A2").Value & Chr(13) & "1/1/2017 - 12/31/2018"
'The string has active values
Worksheets(1).ChartObjects(1).Activate
ActiveChart.HasTitle = True
ActiveChart.ChartTitle.Text = titttle 'Inserts String into title
ActiveWorkbook.Save
End Sub

