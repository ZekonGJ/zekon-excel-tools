Option Explicit

Public Function SplitRows(ByVal ws As Worksheet, ByVal first As Long, ByVal quantityColumn As String, ByVal sequenceColumn As String, ByVal visibleOnly As Boolean) As Long
    Dim qc As Long, sc As Long, finish As Long, r As Long, n As Long, j As Long
    Dim quantityValues As Variant
    Dim v As Variant, extra As Double, quantities() As Long, bottom As Range, sequence() As Variant, rowHeight As Double, rowHidden As Boolean
    EnsureWritable ws
    qc = ColumnNumber(ws, quantityColumn): sc = ColumnNumber(ws, sequenceColumn)
    If qc = sc Then Fail "Ilosc i numer sztuki musza byc w roznych kolumnach."
    If ws.ListObjects.Count > 0 Then Fail "Rozbijanie: najpierw zamien tabele Excela na zwykly zakres (Projekt tabeli > Konwertuj na zakres)."
    finish = LastRow(ws, qc)
    If finish < first Then Fail "Brak danych do rozbicia."
    ReDim quantities(first To finish)
    quantityValues = ReadColumnValues(ws, first, finish, qc)
    ProgressTick "Sprawdzanie ilosci", 0, finish - first + 1, True
    For r = first To finish
        ProgressTick "Sprawdzanie ilosci", r - first + 1, finish - first + 1
        If Included(ws, r, visibleOnly) Then
            v = quantityValues(r - first + 1, 1)
            If IsError(v) Then Fail "Blad ilosci w wierszu " & r
            If Len(CStr(v)) > 0 Then
                quantities(r) = WholeNumber(CStr(v), "Ilosc, wiersz " & r, ws.Rows.Count)
                If IsNull(ws.Rows(r).MergeCells) Then Fail "Scalone komorki w wierszu " & r
                If ws.Rows(r).MergeCells Then Fail "Scalone komorki w wierszu " & r
                extra = extra + quantities(r) - 1
            End If
        End If
    Next r
    Set bottom = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False, SearchFormat:=False)
    If Not bottom Is Nothing Then
        If bottom.Row + extra > ws.Rows.Count Then Fail "Za malo wierszy w arkuszu."
    End If
    If extra > 100000 Then Fail "Jedna operacja moze dodac najwyzej 100 000 wierszy. Podziel dane na mniejsze partie."
    ProgressTick "Rozbijanie: wiersze zrodlowe", 0, finish - first + 1, True
    For r = finish To first Step -1
        n = quantities(r)
        If n > 0 Then
            rowHeight = ws.Rows(r).RowHeight: rowHidden = ws.Rows(r).Hidden
            If n > 1 Then
                ws.Rows(r + 1).Resize(n - 1).Insert Shift:=xlDown
                ' Repeat the source row over one destination block. Excel adjusts formulas.
                ws.Rows(r).Copy Destination:=ws.Rows(r + 1).Resize(n - 1)
                ws.Rows(r + 1).Resize(n - 1).Hidden = rowHidden
                ws.Rows(r + 1).Resize(n - 1).RowHeight = rowHeight
            End If
            ReDim sequence(1 To n, 1 To 1)
            For j = 1 To n
                sequence(j, 1) = j
            Next j
            ws.Cells(r, qc).Resize(n, 1).Value2 = 1
            ws.Cells(r, sc).Resize(n, 1).Value2 = sequence
            SplitRows = SplitRows + n
        End If
        ProgressTick "Rozbijanie: wiersze zrodlowe", finish - r + 1, finish - first + 1
    Next r
    Application.CutCopyMode = False
End Function
