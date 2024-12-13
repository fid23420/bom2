Sub UpdateQuantity()
    Dim wsActive As Worksheet
    Dim wsTarget As Worksheet
    Dim lastRow As Long
    Dim code As String
    Dim length As Double
    Dim elevation As String
    Dim targetRow As Long
    Dim colToUpdate As String
    
    ' 활성화 시트와 타겟 시트 설정
    Set wsActive = ActiveSheet
    Set wsTarget = Worksheets("섬시트")
    
    ' 활성화 시트의 마지막 행 찾기
    lastRow = wsActive.Cells(wsActive.Rows.Count, "A").End(xlUp).Row
    Debug.Print "Last Row: " & lastRow
    
    ' 활성화 시트의 데이터를 순회
    For i = 2 To lastRow
        code = wsActive.Cells(i, 1).Value
        On Error Resume Next
        length = wsActive.Cells(i, 2).Value
        On Error GoTo 0
        elevation = wsActive.Cells(i, 3).Value
        
        Debug.Print "Processing Row: " & i & ", Code: " & code & ", Length: " & length & ", Elevation: " & elevation
        
        ' 타겟 시트에서 일치하는 code의 행 찾기
        targetRow = 0
        On Error Resume Next
        targetRow = Application.WorksheetFunction.Match(code, wsTarget.Columns("F"), 0)
        On Error GoTo 0
        
        Debug.Print "Target Row for Code " & code & ": " & targetRow
        
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
                Case "고소20%+유해10%"
                    colToUpdate = "M"
                Case Else
                    colToUpdate = ""
            End Select
            
            Debug.Print "Column to Update for Elevation " & elevation & ": " & colToUpdate
            
            If colToUpdate <> "" Then
                ' 기존 값과 새 값을 합산하여 입력
                wsTarget.Cells(targetRow, colToUpdate).Value = wsTarget.Cells(targetRow, colToUpdate).Value + length
                Debug.Print "Updated " & colToUpdate & targetRow & " with Length: " & length
            End If
        Else
            Debug.Print "No matching row found for Code: " & code
        End If
    Next i
    
    Debug.Print "Macro Execution Completed"
End Sub
