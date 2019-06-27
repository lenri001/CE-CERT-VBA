Attribute VB_Name = "Module13"
Function LOADFACTOR(onpeak_kWh, midpeak_kWh, offpeak_kWh, onpeak_kW, midpeak_kW, offpeak_kW, days)
avgload = (onpeak_kWh + midpeak_kWh + offpeak_kWh) / (days * 24#)
maxload = Max(onpeak_kW, midpeak_kW, offpeak_kW)
LOADFACTOR = avgload / maxload
End Function
Function PRIORITY(looadfactor, setval1, setval2, setval3, setval4)
        If looadfactor < setval1 Then
            PRIORITY = "High Priority 1"
        ElseIf looadfactor < setval2 Then
            PRIORITY = "High Priority 2"
        ElseIf looadfactor < setval3 Then
            PRIORITY = "High Priority 3"
        ElseIf looadfactor < setval4 Then
            PRIORITY = "Medium Priority"
        Else
            PRIORITY = "Low Priority"
        End If
End Function
