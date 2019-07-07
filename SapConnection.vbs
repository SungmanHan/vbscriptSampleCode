' ScriptName : SapConnection
' Creator : Sungman Han
' Creation Date : 2019.05.28
' Descrition : SAP Connection

'------------------------------------------------
' Variable Setting
'------------------------------------------------
Dim vIP,vLanguage,vUserName,vPassWord,vApplication,vConnection,vSession

vIP = WScript.Arguments.Item(0)
vLanguage = WScript.Arguments.Item(1)
vUserName = WScript.Arguments.Item(2)
vPw = WScript.Arguments.Item(3)

'------------------------------------------------
' Sap Used Object
'------------------------------------------------

Do
  Set SapGuiAuto = GetObject("SAPGUI")
  Set vApplication = SapGuiAuto.GetScriptingEngine
  WScript.Sleep 1000
Loop While vApplication Is Nothing

Do
  Set vConnection = vApplication.OpenConnectionByConnectionString(vIP, True)
  WScript.Sleep 1000
Loop While vConnection Is Nothing

Do
  Set vSession = vConnection.Children(0)
  WScript.Sleep 1000
Loop While vSession Is Nothing

If IsObject(WScript) Then
  WScript.ConnectObject vSession, "on"
  WScript.ConnectObject vApplication, "on"
End If

'------------------------------------------------
' Sap Login
'------------------------------------------------
vSession.findById("wnd[0]").maximizea
vSession.findById("wnd[0]/usr/txtRSYST-LANGU").text = vLanguage
vSession.findById("wnd[0]/usr/txtRSYST-BNAME").text = vUserName
vSession.findById("wnd[0]/usr/pwdRSYST-BCODE").text = vPassWord
vSession.findById("wnd[0]").sendVKey 0
