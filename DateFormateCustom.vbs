' ScriptName : DateFormateCustom
' Creator : Sungman Han
' Creation Date : 2019.06.21
' Descrition : Date Formate Custom Function(Input Date Formate m/d/y h:m:s)

Function DateFormateCustom(vTempInputDate,vTempStandard)

  Dim vTempMonth,vTempDay,vTempYear,vTempCount,vTempSplit

  vTempCount = 0

  If InStr(vTempDateString,"/") <> 0 Then
    vTempSplit = "/"
  ElseIf InStr(vTempDateString,"-") <> 0 Then
    vTempSplit = "_"
  ElseIf InStr(vTempDateString,".") <> 0 Then
    vTempSplit = "."
  End If

  For Each vTemp In Split(vTempInputDate,vTempSplit)

    vTempCount = vTempCount + 1

    If vTempCount = 1 Then
      ' Month
      vTempMonth = vTemp
    ElseIf vTempCount = 2 Then
      ' Day
      vTempDay = vTemp
    Else
      ' Year
      vTempYear = vTemp
    End If

  Next

  DateCustom = vTempYear&vTempStandard&vTempMonth&vTempStandard&vTempDay

End Function
