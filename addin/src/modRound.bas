Option Explicit

Public Function RoundColumn(ByVal ws As Worksheet, ByVal first As Long, ByVal column As String, ByVal visibleOnly As Boolean, ByVal overwriteFormulas As Boolean) As Long
    Dim c As Long, r As Long, finish As Long, cell As Range, v As Variant
    Dim block(1 To 2048, 1 To 1) As Variant, count As Long, start As Long, eligible As Boolean
    EnsureWritable ws
    c = ColumnNumber(ws, column): finish = LastRow(ws, c)
    ProgressTick "Zaokraglanie", 0, finish - first + 1, True
    For r = first To finish
        ProgressTick "Zaokraglanie", r - first + 1, finish - first + 1
        If Included(ws, r, visibleOnly) Then
            Set cell = ws.Cells(r, c)
            If cell.MergeCells Then Fail "Scalona komorka w wierszu " & r & "."
            If cell.HasArray Then Fail "Formula tablicowa w wierszu " & r & "."
            If IsError(cell.Value2) Then Fail "Blad w wierszu " & r & "."
        End If
    Next r
    ProgressTick "Zaokraglanie: zapis", 0, finish - first + 1, True
    For r = first To finish
        eligible = False
        If Included(ws, r, visibleOnly) Then
            Set cell = ws.Cells(r, c): v = cell.Value2
            If Not cell.HasFormula Or overwriteFormulas Then
                If Not IsEmpty(v) And VarType(v) <> vbBoolean Then
                    If IsNumeric(v) And Len(CStr(v)) > 0 Then eligible = True
                End If
            End If
        End If
        If eligible Then
            If count = 0 Then start = r
            count = count + 1
            block(count, 1) = Application.WorksheetFunction.Round(CDbl(v), 0)
            RoundColumn = RoundColumn + 1
        End If
        If (Not eligible Or count = 2048 Or r = finish) And count > 0 Then
            WriteRoundedBlock ws, c, start, count, block
            count = 0
        End If
        ProgressTick "Zaokraglanie: zapis", r - first + 1, finish - first + 1
    Next r
End Function

Private Sub WriteRoundedBlock(ByVal ws As Worksheet, ByVal col As Long, ByVal first As Long, ByVal count As Long, ByRef block() As Variant)
    Dim values() As Variant, i As Long
    ReDim values(1 To count, 1 To 1)
    For i = 1 To count
        values(i, 1) = block(i, 1)
    Next i
    ws.Cells(first, col).Resize(count, 1).Value2 = values
End Sub
