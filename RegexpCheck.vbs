' ScriptName : RegexpCheck
' Creator : Sungman Han
' Creation Date : 2019.05.28
' Descrition : Regexp Check, Input tmp(check data), pat(pattern)

Function RegexpCheck(tmp,pat)

  Dim vReturnValue

  vReturnValue = ""

  With CreateObject("Vbscript.regexp")
    .Global = True
    .ignorecase = True
    .Pattern = pat

    If .test(tmp) Then
      vReturnValue = .Replace(tmp, "")
    Else
      vReturnValue = tmp
    End If

  End With

  If vReturnValue <> "" Then RegexpCheck = vReturnValue

End Function
