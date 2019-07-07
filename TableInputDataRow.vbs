' ScriptName : TableInputDataRow
' Creator : Sungman Han
' Creation Date : 2019.05.21
' Descrition : When you insert data into the SAP table, you increment the row and insert a value.

'------------------------------------------------
' Function List
'------------------------------------------------
Function MultiInput(vTempListString,vTempID)

  vCount = 0

  For Each vTempString In Split(vTempListString,",")

    Dim tempTwo

    tempTwo = vTempID

    Call VerticalScrolling(tempTwo, vCount)

    Set table = session.findById(tempTwo)
    Set row = table.getcell(1, 1)
    row.text = vTempString
    vCount = vCount + 1

  Next

  session.findById("wnd[1]/tbar[0]/btn[8]").press

End Function

Public Sub VerticalScrolling(vID, vIx)
  Session.findById(vID).verticalScrollbar.Position = vIx
End Sub
