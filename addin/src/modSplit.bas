Option Explicit

Public Function SplitRows(ByVal ws As Worksheet, ByVal first As Long, ByVal quantityColumn As String, ByVal sequenceColumn As String, ByVal visibleOnly As Boolean) As Long
    Dim qc As Long, sc As Long, finish As Long, r As Long, n As Long, j As Long
    Dim quantityValues As Variant
    Dim v As Variant, extra As Double, quantities() As Long, bottom As Range, sequence() As Variant, rowHeight As Double, rowHidden As Boolean
    Dim filterState As Variant, filterTop As Long, filterBottom As Long, filterLeft As Long, filterWidth As Long
    Dim restoreNeeded As Boolean, e As Long, message As String, restoreError As String
    On Error GoTo Bad
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
    ' quantities is the immutable selection captured while the original filter is active.
    If ws.AutoFilterMode Then
        filterState = CaptureSplitFilter(ws, filterTop, filterBottom, filterLeft, filterWidth)
        If first <= filterTop And finish >= filterTop Then Fail "Poczatek danych obejmuje naglowek filtra."
        restoreNeeded = True
        If ws.FilterMode Then ws.ShowAllData
        ws.AutoFilterMode = False
    End If
    ProgressTick "Rozbijanie: wiersze zrodlowe", 0, finish - first + 1, True
    For r = finish To first Step -1
        n = quantities(r)
        If n > 0 Then
            rowHeight = ws.Rows(r).RowHeight: rowHidden = ws.Rows(r).Hidden
            If n > 1 Then
                ws.Rows(r + 1).Resize(n - 1).Insert Shift:=xlDown
                ' Track successful inserts, including copies of the last filtered row.
                If restoreNeeded Then
                    If r < filterTop Then
                        filterTop = filterTop + n - 1: filterBottom = filterBottom + n - 1
                    ElseIf r <= filterBottom Then
                        filterBottom = filterBottom + n - 1
                    End If
                End If
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
    If restoreNeeded Then
        RestoreSplitFilter ws, filterState, filterTop, filterBottom, filterLeft, filterWidth
        restoreNeeded = False
    End If
    Exit Function
Bad:
    e = Err.Number: message = Err.Description
    On Error Resume Next
    Application.CutCopyMode = False
    If restoreNeeded Then
        Err.Clear
        RestoreSplitFilter ws, filterState, filterTop, filterBottom, filterLeft, filterWidth
        If Err.Number <> 0 Then restoreError = Err.Description
    End If
    On Error GoTo 0
    If Len(restoreError) > 0 Then message = message & vbCrLf & "Nie odtworzono filtra: " & restoreError
    Err.Raise e, "SplitRows", message
End Function

Private Function CaptureSplitFilter(ByVal ws As Worksheet, ByRef top As Long, ByRef bottom As Long, ByRef left As Long, ByRef width As Long) As Variant
    Dim rg As Range, item As Filter, state() As Variant, i As Long, op As Long
    Set rg = ws.AutoFilter.Range
    top = rg.Row: bottom = top + rg.Rows.Count - 1
    left = rg.Column: width = rg.Columns.Count
    ReDim state(1 To width, 1 To 6)
    For i = 1 To width
        Set item = ws.AutoFilter.Filters(i)
        state(i, 1) = item.On
        If item.On Then
            op = item.Operator: state(i, 2) = op
            If op = xlFilterIcon Then Fail "Rozbijanie: filtr ikon wymaga wczesniejszej zamiany na filtr wartosci. Dane nie zostaly zmienione."
            ' Criteria1 or Criteria2 can be absent, including grouped date filters.
            On Error Resume Next
            Err.Clear: state(i, 3) = item.Criteria1: state(i, 5) = (Err.Number = 0)
            Err.Clear: state(i, 4) = item.Criteria2: state(i, 6) = (Err.Number = 0)
            Err.Clear
            On Error GoTo 0
            If Not state(i, 5) And Not state(i, 6) Then Fail "Nie mozna odczytac kryteriow filtra. Dane nie zostaly zmienione."
        End If
    Next i
    CaptureSplitFilter = state
End Function

Private Sub RestoreSplitFilter(ByVal ws As Worksheet, ByVal state As Variant, ByVal top As Long, ByVal bottom As Long, ByVal left As Long, ByVal width As Long)
    Dim rg As Range, i As Long
    ws.AutoFilterMode = False
    Set rg = ws.Range(ws.Cells(top, left), ws.Cells(bottom, left + width - 1))
    rg.AutoFilter
    For i = 1 To width
        If state(i, 1) Then
            If state(i, 5) And state(i, 6) Then
                rg.AutoFilter Field:=i, Criteria1:=state(i, 3), Operator:=state(i, 2), Criteria2:=state(i, 4)
            ElseIf state(i, 5) Then
                If state(i, 2) = 0 Then
                    rg.AutoFilter Field:=i, Criteria1:=state(i, 3)
                Else
                    rg.AutoFilter Field:=i, Criteria1:=state(i, 3), Operator:=state(i, 2)
                End If
            Else
                rg.AutoFilter Field:=i, Operator:=state(i, 2), Criteria2:=state(i, 4)
            End If
        End If
    Next i
End Sub
