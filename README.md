Sub UpdateQuantity()
    Dim wsActive As Worksheet
    Dim wsTarget As Worksheet
    Dim lastRow As Long
    Dim lastCol As Long
    Dim i As Long, j As Long
    Dim headerRow As Long
    Dim code As String
    Dim length As Double
    Dim elevation As String
    Dim targetRow As Long
    Dim colToUpdate As String
    Dim codeCol As Long, lengthCol As Long, elevationCol As Long
    
    ' 활성화 시트와 타겟 시트 설정
    Set wsActive = ActiveSheet
    Set wsTarget = Worksheets("섬시트")
    
    ' 활성화 시트의 마지막 행 및 마지막 열 찾기
    lastRow = wsActive.Cells(wsActive.Rows.Count, "A").End(xlUp).Row
    lastCol = wsActive.Cells(1, wsActive.Columns.Count).End(xlToLeft).Column
    
    ' 헤더 행 찾기 (일반적으로 1행이라고 가정)
    headerRow = 1
    
    ' 헤더 열 위치 찾기
    For j = 1 To lastCol
        Select Case LCase(wsActive.Cells(headerRow, j).Value)
            Case "code"
                codeCol = j
            Case "length"
                lengthCol = j
            Case "elevation"
                elevationCol = j
        End Select
    Next j
    
    ' 활성화 시트의 데이터를 순회
    For i = 2 To lastRow
        code = wsActive.Cells(i, codeCol).Value
        On Error Resume Next
        length = wsActive.Cells(i, lengthCol).Value
        On Error GoTo 0
        elevation = wsActive.Cells(i, elevationCol).Value
        
        ' 타겟 시트에서 일치하는 code의 행 찾기
        targetRow = 0
        On Error Resume Next
        targetRow = Application.WorksheetFunction.Match(code, wsTarget.Columns("F"), 0)
        On Error GoTo 0
        
        ' 일치하는 행이 있을 경우
        If targetRow > 0 Then
            Select Case elevation
                Case "---"
                    colToUpdate = "H"
                Case "고소작업 10% 할증"
                    colToUpdate = "J"
                Case "고소작업 20% 할증"
                    colToUpdate = "K"
                Case "고소10%+유해10%"
                    colToUpdate = "L"
