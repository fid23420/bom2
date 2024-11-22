Sub MatchAndOrganizeWithUtilityGroupAndBulkLogic()
    Dim wsCurrent As Worksheet
    Dim wsSource As Worksheet
    Dim wsSheet1 As Worksheet
    Dim filterValue1 As String, filterValue2 As String
    Dim lastRowSource As Long, lastRowSheet1 As Long
    Dim destinationRow As Long
    Dim materialColumn As Integer, widthColumn As Integer
    Dim lengthColumn As Integer, elevationColumn As Integer
    Dim nameColumn As Integer, qtyColumn As Integer, utilityGroupColumn As Integer, groupColumn As Integer
    Dim i As Long, j As Long
    Dim materialOrName As String
    Dim lengthOrQty As Variant
    Dim uniqueKey As Variant
    Dim keyParts() As String
    Dim matchFound As Boolean
    Dim currentWidth As String
    Dim isWaterGroup As Boolean

    ' 현재 실행 중인 시트와 데이터 시트 설정
    Set wsCurrent = ThisWorkbook.ActiveSheet
    Set wsSource = ThisWorkbook.Sheets("붙여넣기")
    Set wsSheet1 = ThisWorkbook.Sheets("Sheet1")

    ' 현재 시트의 A2와 B2 값 가져오기
    filterValue1 = Trim(UCase(wsCurrent.Range("A2").Value))
    filterValue2 = Trim(UCase(wsCurrent.Range("B2").Value))

    ' '붙여넣기' 시트 및 Sheet1의 마지막 데이터 행 찾기
    lastRowSource = wsSource.Cells(wsSource.Rows.Count, "A").End(xlUp).Row
    lastRowSheet1 = wsSheet1.Cells(wsSheet1.Rows.Count, "A").End(xlUp).Row

    ' 각 열의 위치 찾기
    materialColumn = 0
    widthColumn = 0
    lengthColumn = 0
    elevationColumn = 0
    nameColumn = 0
    qtyColumn = 0
    utilityGroupColumn = 0
    groupColumn = 0

    For i = 1 To wsSource.Cells(1, wsSource.Columns.Count).End(xlToLeft).Column
        Select Case Trim(UCase(wsSource.Cells(1, i).Value))
            Case "MATERIAL"
                materialColumn = i
            Case "WIDTH/DIAMETER"
                widthColumn = i
            Case "LENGTH"
                lengthColumn = i
            Case "ELEVATION"
                elevationColumn = i
            Case "NAME"
                nameColumn = i
            Case "QTY"
                qtyColumn = i
            Case "UTILITY GROUP"
                utilityGroupColumn = i
            Case "GROUP"
                groupColumn = i
        End Select
    Next i

    ' 필요한 열이 없는 경우 메시지 출력 후 종료
    If materialColumn = 0 Or widthColumn = 0 Or lengthColumn = 0 Or elevationColumn = 0 Or nameColumn = 0 Or qtyColumn = 0 Or utilityGroupColumn = 0 Or groupColumn = 0 Then
        MsgBox "필수 열(MATERIAL, WIDTH/DIAMETER, LENGTH, ELEVATION, NAME, QTY, UTILITY GROUP, GROUP)이 누락되었습니다!", vbExclamation
        Exit Sub
    End If

    ' Dictionary 생성
    Dim sumDict As Object
    Set sumDict = CreateObject("Scripting.Dictionary")

    ' 조건에 맞는 데이터를 Dictionary에 추가
    For i = 2 To lastRowSource
        ' 기존 A열 조건에서 Group 열로 변경
        If InStr(1, Trim(UCase(wsSource.Cells(i, groupColumn).Value)), filterValue1, vbTextCompare) > 0 And _
           InStr(1, Trim(UCase(wsSource.Cells(i, groupColumn).Value)), filterValue2, vbTextCompare) > 0 Then

            ' Material 또는 Name 구분
            If Trim(wsSource.Cells(i, materialColumn).Value) <> "" Then
                materialOrName = wsSource.Cells(i, materialColumn).Value

                ' "CHANNEL BOLT" 항목 제외
                If UCase(materialOrName) = "CHANNEL BOLT" Then
                    GoTo NextIteration ' 현재 행 제외
                End If

                ' Length 값 처리 (Material에만 적용)
                If IsNumeric(wsSource.Cells(i, lengthColumn).Value) Then
                    lengthOrQty = Application.Round(wsSource.Cells(i, lengthColumn).Value / 1000, 1)
                Else
                    lengthOrQty = 0 ' Length가 없으면 기본값 0
                End If

                ' Utility Group 확인
                isWaterGroup = InStr(1, Trim(UCase(wsSource.Cells(i, utilityGroupColumn).Value)), "/5D-WATER", vbTextCompare) > 0

                ' Unique Key 생성 (Material/Name + Width + Elevation + Utility Group)
                uniqueKey = Trim(UCase(materialOrName)) & "|" & Trim(UCase(wsSource.Cells(i, widthColumn).Value)) & "|" & Trim(UCase(wsSource.Cells(i, elevationColumn).Value)) & "|" & IIf(isWaterGroup, "/5D-WATER", "OTHER")
            ElseIf Trim(UCase(wsSource.Cells(i, nameColumn).Value)) <> "" And _
                     InStr(1, Trim(UCase(wsSource.Cells(i, nameColumn).Value)), "[BULK-", vbTextCompare) > 0 Then
                materialOrName = wsSource.Cells(i, nameColumn).Value
                lengthOrQty = Val(wsSource.Cells(i, qtyColumn).Value)

                ' Unique Key 생성 (Name 항목은 Elevation만 기준으로 합산)
                uniqueKey = Trim(UCase(materialOrName)) & "||" & Trim(UCase(wsSource.Cells(i, elevationColumn).Value))
            Else
                GoTo NextIteration
            End If

            ' Dictionary에 추가 및 수량 합산
            If sumDict.exists(uniqueKey) Then
                sumDict(uniqueKey) = sumDict(uniqueKey) + Val(lengthOrQty)
            Else
                ' 초기 데이터 저장
                sumDict.Add uniqueKey, Val(lengthOrQty)
            End If
        End If
NextIteration:
    Next i

    ' [BULK-] Material 항목 처리 및 F열 값 반환
    Dim materialValue As String
    Dim matched As Boolean
    For i = 15 To wsCurrent.Cells(wsCurrent.Rows.Count, "D").End(xlUp).Row
        materialValue = Trim(UCase(wsCurrent.Cells(i, "D").Value))
        matched = False

        ' [BULK-] 단어가 포함된 항목만 처리
        If InStr(1, materialValue, "[BULK-", vbTextCompare) > 0 Then
            For j = 2 To lastRowSheet1
                ' Sheet1의 A열과 Material 매칭
                If Trim(UCase(wsSheet1.Cells(j, "A").Value)) = materialValue Then
                    ' 매칭된 F열 값 반환
                    wsCurrent.Cells(i, "E").Value = wsSheet1.Cells(j, "F").Value
                    matched = True
                    Exit For
                End If
            Next j

            ' 매칭되지 않을 경우 원래 값 유지
            If Not matched Then
                wsCurrent.Cells(i, "E").Value = "No Match"
            End If
        End If
    Next i

    ' Dictionary의 데이터를 현재 시트에 기록
    destinationRow = 15
    For Each uniqueKey In sumDict.Keys
        keyParts = Split(uniqueKey, "|")
        wsCurrent.Cells(destinationRow, "D").Value = keyParts(0) ' Material/Name
        wsCurrent.Cells(destinationRow, "E").Value = IIf(UBound(keyParts) > 1, keyParts(1), "") ' Width/Diameter
        wsCurrent.Cells(destinationRow, "G").Value = keyParts(UBound(keyParts) - 1) ' Elevation
        wsCurrent.Cells(destinationRow, "F").Value = sumDict(uniqueKey) ' 합산된 수량

        ' '/5D-WATER' 조건에 따라 음영 처리
        If keyParts(UBound(keyParts)) = "/5D-WATER" Then
            wsCurrent.Rows(destinationRow).Interior.Color = RGB(173, 216, 230) ' 파란색 음영
        End If

        destinationRow = destinationRow + 1
    Next uniqueKey

    MsgBox "rev.07 코드가 업데이트되었으며, BULK 항목 매칭이 완료되었습니다!", vbInformation
End Sub
