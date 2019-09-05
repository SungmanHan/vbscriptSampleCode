Dim vFSO,vYear,vMonth,vDay,vFolderName,vTargetFolderPath

vTargetFolderPath = "C:\Users\powergen\Downloads"
vFolderPath = "C:\"

vYear = Year(Date)
vMonth = Month(Date)
vDay = Day(Date)
vFolderName = vYear&"_"&vMonth&"_"&vDay
vFolderPath = vFolderPath&"\"&vFolderName

Set vFSO = CreateObject("Scripting.FileSystemObject")

If vFSO.FolderExists(vFolderPath) Then
Else
	Msgbox(vFolderPath)
	vFSO.CreateFolder vFolderPath
End If

vFSO.GetFolder(vTargetFolderPath).Copy vFolderPath

Set vFSO = Nothing