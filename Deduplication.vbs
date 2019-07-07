' ScriptName : Deduplication
' Creator : Sungman Han
' Creation Date : 2019.06.15
' Descrition : Deduplication
' Ex)
' input data : vTempString  = "a,a,a,a,a,c,c,c,b,b,b,b,e,f,g,t,i,j,j,k,l", vTempStandard = ","
' output data : "a,c,b,e,f,g,t,j,k,l"

Function Deduplication(vTempString,vTempStandard)

  vTempStringOne = ""
  vCopyString = ""
  vCopyString = vTempString

  For Each x In Split(vTempString,vTempStandard)

    If InStr(vCopyString,x) <> 0 Then

      vTempStringOne = vTempStringOne&","&x
      vCopyString = Replace(vCopyString,x,"")

    End If

  Next

  Deduplication = Mid(vTempStringOne,2,Len(vTempStringOne))

End Function
