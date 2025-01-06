Sub FilterAndCopyData()
    Dim sourceWorkbook As Workbook
    Dim targetWorkbook As Workbook
    Dim sourceSheet As Worksheet
    Dim targetSheet As Worksheet
    Dim sourceFilePath As String
    Dim targetFilePath As String
    Dim folderPath As String
    Dim keyword As String
    Dim fileName As String
    Dim cell As Range
    Dim lastRow As Long
    Dim copyRange As Range
    Dim targetRow As Long
    Dim checkDate As Date
    Dim currentRow As Long
    Dim prevDate As Date

    ' 설정
    folderPath = "C:\Users\sk22.id\Downloads\관리대장\"
    targetFilePath = "C:\Users\sk22.id\Downloads\관리대장\3붙.xlsx"
    keyword = "보고서" ' 원하는 키워드를 여기에 직접 입력

    ' 폴더에서 키워드 포함 파일 찾기
    fileName = Dir(folderPath & "*" & keyword & "*.xlsx")

    If fileName = "" Then
        MsgBox "지정된 키워드에 해당하는 파일이 없습니다.", vbExclamation
        Exit Sub
    End If

    ' 소스 워크북 열기
    sourceFilePath = folderPath & fileName
    Set sourceWorkbook = Workbooks.Open(sourceFilePath)
    Set sourceSheet = sourceWorkbook.Sheets("관리대장")

    ' 대상 워크북 열기
    If Dir(targetFilePath) = "" Then
        Set targetWorkbook = Workbooks.Add
        targetWorkbook.SaveAs targetFilePath
    Else
        Set targetWorkbook = Nothing
        On Error Resume Next
        Set targetWorkbook = Workbooks(targetFilePath)
        On Error GoTo 0
        If targetWorkbook Is Nothing Then
            Set targetWorkbook = Workbooks.Open(targetFilePath, ReadOnly:=False)
        End If
    End If

    ' "P4" 시트 확인 및 설정
    On Error Resume Next
    Set targetSheet = targetWorkbook.Sheets("P4")
    On Error GoTo 0

    If targetSheet Is Nothing Then
        MsgBox "P4 시트가 없습니다. 작업을 중단합니다.", vbExclamation
        sourceWorkbook.Close False
        Exit Sub
    End If

    ' P4 시트 초기화
    targetSheet.Cells.Clear

    ' 필터 기준 설정
    checkDate = Date - 7

    ' 마지막 행 가져오기
    lastRow = sourceSheet.Cells(sourceSheet.Rows.Count, "B").End(xlUp).Row

    ' 데이터 필터링 및 복사
    targetRow = 1 ' P4 시트 처음부터 출력
    prevDate = 0 ' 이전 날짜 초기화

    For currentRow = 2 To lastRow ' 헤더 제외
        With sourceSheet
            If IsDate(.Cells(currentRow, "B")) Then
                If .Cells(currentRow, "B").Value >= checkDate Then
                    If InStr(.Cells(currentRow, "H"), "5D") = 0 And InStr(.Cells(currentRow, "H"), "출도") = 0 Then
                        ' 날짜가 변경되었으면 공백 행 추가
                        If .Cells(currentRow, "B").Value <> prevDate Then
                            targetSheet.Cells(targetRow, 1).EntireRow.Insert
                            targetRow = targetRow + 1
                        End If

                        ' 복사 범위 설정 (B~V 열)
                        targetSheet.Cells(targetRow, 1).Resize(, 21).Value = .Cells(currentRow, "B").Resize(, 21).Value
                        targetRow = targetRow + 1

                        ' 이전 날짜 갱신
                        prevDate = .Cells(currentRow, "B").Value
                    End If
                End If
            End If
        End With
    Next currentRow

    ' 워크북 저장 (열린 상태 유지)
    targetWorkbook.Save
    sourceWorkbook.Close False

    MsgBox "데이터 복사가 완료되었습니다.", vbInformation
End Sub
