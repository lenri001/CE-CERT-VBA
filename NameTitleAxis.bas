Attribute VB_Name = "Module1"
Sub NameTitleAxis()
Attribute NameTitleAxis.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Axis Macro
'

'
    
    'x-axis name
    ActiveChart.Axes(xlCategory, xlPrimary).HasTitle = True 'Creates the x-axis
    ActiveChart.Axes(xlCategory, xlPrimary).AxisTitle.Characters.Text = "15 Minute Interval Data" 'Names the x-axis
    'y-axis name
    ActiveChart.Axes(xlValue, xlPrimary).HasTitle = True 'Creates the y-axis
    ActiveChart.Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = "Demand (kW)" 'Names the y-axis
    'ActiveChart.SetElement (msoElementPrimaryValueAxisTitleAdjacentToAxis)
    'Selection.Caption = "Demand (kW)"
End Sub
