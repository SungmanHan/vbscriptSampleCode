On Error Resume Next
' ScriptName : ReservationExcelEdit
' Creator : Sungman Han
' Creation Date : 2019.05.30
' Descrition : 엑셀 파일 = 1) 컬럼삭제 2) Row Check 3) Set 납품업체코드

'--------------------------------------------------------------------------------
' 변수 선언
'--------------------------------------------------------------------------------
Dim vExcelPath,objexcel,objWorkbook,objSheet,vMaxRow,vSupplier,vTempColumn

vExcelPath = WScript.Arguments.Item(0)
vSupplier = WScript.Arguments.Item(1)

If Not IsObject(objexcel) Then
Set objexcel = GetObject(,"Excel.Application")
End If

If Not IsObject(objWorkbook) Then
Set objWorkbook = objexcel.Workbooks(vExcelPath)
End If

If Not IsObject(objSheet) Then
Set objSheet = objWorkbook.Worksheets("Sheet1")
End If

'--------------------------------------------------------------------------------
' 컬럼삭제
'--------------------------------------------------------------------------------
objSheet.Range("C:C").EntireColumn.Delete

'--------------------------------------------------------------------------------
' Max Row 확인
'--------------------------------------------------------------------------------
vTempColumn = FindColume("플랜트",objSheet)
vMaxRow = CheckLow(2,vTempColumn,objSheet)

'--------------------------------------------------------------------------------
' 납품처 Max Row만큼 값 넣기
'--------------------------------------------------------------------------------
vTempColumn = FindColume("납품처",objSheet)
Call setValue(objSheet,vTempColumn,vMaxRow,vSupplier)

'--------------------------------------------------------------------------------
' 생산버전 -> Status명으로 변경, 값 빈값으로 변경
'--------------------------------------------------------------------------------
vTempColumn = FindColume("생산버전",objSheet)
objSheet.Cells(1,vTempColumn).Value = "STATUS"

vMaxRow = CheckLow(2,FindColume("플랜트",objSheet),objSheet)
Call setValue(objSheet,vTempColumn,vMaxRow,"")

'--------------------------------------------------------------------------------
' Max Row 확인
'--------------------------------------------------------------------------------
vMaxRow = CheckLow(2,FindColume("플랜트",objSheet),objSheet)

'--------------------------------------------------------------------------------
' 저장 및 객체 사용 종료
'--------------------------------------------------------------------------------
objWorkbook.Save

Set objSheet = Nothing
Set objWorkbook = Nothing
Set objExcel = Nothing

WScript.StdOut.WriteLine(vMaxRow)

Err.Source 6
WScript.StdOut.WriteLine("Fail")
Err.Clear

'--------------------------------------------------------------------------------
' Function List
'--------------------------------------------------------------------------------
Function FindColume(vTempSearch,vTempSheet)

  Set FoundCell = vTempSheet.Range("A:Z").Find(vTempSearch,,,1)

  If Not FoundCell Is Nothing Then

    FindColume = FoundCell.Column

  Else

    FindColume = 0

  End If

  Set FoundCell = Nothing

End Function

Function CheckLow(vRow,vCol,vTempSheet)
  Dim vTempCount2,vGetValu2

  vTempCount2 = 0
  vGetValu2 = ""

  Do

    vGetValu2 = vTempSheet.Cells(vRow,vCol).Value

    If vGetValu2 <> "" Then
      vTempCount2 = vTempCount2 + 1
    End If

    vRow = vRow + 1

  Loop While vGetValu2 <> ""

  CheckLow = vTempCount2

End Function

Function setValue(vTempSheet,vTempColun,vTempMaxRow,vTempSupplier)

  Dim vTempRow,vTempCol
  vTempCol = vTempColun
  vTempRow = 2

  Do

    vTempSheet.Cells(vTempRow,vTempCol).Value = vTempSupplier

    vTempRow = vTempRow + 1

    vTempMaxRow = vTempMaxRow - 1

  Loop While vTempMaxRow <> 0

End Function
