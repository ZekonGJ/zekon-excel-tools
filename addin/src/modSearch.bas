Option Explicit

Public Function SearchPositions(ByVal ws As Worksheet, ByVal first As Long, ByVal p As Variant, ByVal visibleOnly As Boolean) As Long
    Dim bt As Worksheet, sc As Long, bc As Long, br As Long, sr As Long, tr As Long
    Dim source As Object, target As Object, raw As Object, missing As Object
    Dim sourceValues As Variant, targetValues As Variant
    Dim r As Long, k As String, lastCol As Long, criteria As Variant, rg As Range, c As Range
    Dim report As Worksheet, row As Long, nm As String, formula As String, idx As Long, marker As Range
    Dim headerCell As Range, firstCol As Long, filterField As Long, key As Variant
    Dim filterValues() As String, filterIndex As Long, blockStart As Long
    EnsureWritable ws
    sc = ColumnNumber(ws, CStr(p(1)))
    On Error Resume Next
    Set bt = ws.Parent.Worksheets(CStr(p(2)))
    On Error GoTo 0
    If bt Is Nothing Then Fail "Nie znaleziono arkusza: " & CStr(p(2))
    If bt Is ws Then Fail "Wybierz rozne arkusze: lista pozycji i zestawienie."
    EnsureWritable bt
    If ws.Parent.ProtectStructure Then Fail "Struktura skoroszytu jest chroniona."
    If bt.ListObjects.Count > 0 Then Fail "Wyszukiwanie wymaga zwyklego zakresu w arkuszu docelowym, nie tabeli Excela."
    bc = ColumnNumber(bt, CStr(p(3)))
    br = ResolveFirstRow(bt, CStr(p(3)), CStr(p(4)), False)
    Set rg = SearchFilterRange(bt, br, bc)
    sr = LastRow(ws, sc): tr = LastRow(bt, bc)
    If sr < first Or tr < br Then Fail "Brak danych w jednym z arkuszy."
    sourceValues = ReadColumnValues(ws, first, sr, sc)
    targetValues = ReadColumnValues(bt, br, tr, bc)
    Set source = CreateObject("Scripting.Dictionary"): source.CompareMode = vbTextCompare
    Set target = CreateObject("Scripting.Dictionary"): target.CompareMode = vbTextCompare
    Set raw = CreateObject("Scripting.Dictionary"): raw.CompareMode = vbTextCompare
    Set missing = CreateObject("Scripting.Dictionary"): missing.CompareMode = vbTextCompare
    ProgressTick "Wyszukiwanie: lista", 0, sr - first + 1, True
    For r = first To sr
        ProgressTick "Wyszukiwanie: lista", r - first + 1, sr - first + 1
        If Included(ws, r, visibleOnly) Then
            k = PositionKey(sourceValues(r - first + 1, 1))
            If Len(k) > 0 Then source(k) = True
        End If
    Next r
    If source.Count = 0 Then Fail "Lista pozycji jest pusta."
    ProgressTick "Wyszukiwanie: zestawienie", 0, tr - br + 1, True
    For r = br To tr
        ProgressTick "Wyszukiwanie: zestawienie", r - br + 1, tr - br + 1
        k = PositionKey(targetValues(r - br + 1, 1))
        If Len(k) > 0 Then
            target(k) = True
            If source.Exists(k) Then raw(CStr(targetValues(r - br + 1, 1))) = True
        End If
    Next r
    If raw.Count > 10000 Then Fail "Za duzo kryteriow filtra (ponad 10 000). Podziel liste."
    ProgressTick "Wyszukiwanie: lista", 0, sr - first + 1, True
    For r = first To sr
        ProgressTick "Wyszukiwanie: lista", r - first + 1, sr - first + 1
        If Included(ws, r, visibleOnly) Then
            k = PositionKey(sourceValues(r - first + 1, 1))
            If Len(k) > 0 Then
                If Not target.Exists(k) Then missing(CStr(r)) = k
            End If
        End If
    Next r
    Set rg = SearchFilterRange(bt, br, bc)
    filterField = bc - rg.Column + 1
    If bt.Visible <> xlSheetVisible Then Fail "Arkusz zestawienia musi byc widoczny."
    bt.Parent.Activate
    bt.Activate
    If bt.AutoFilterMode Then bt.AutoFilterMode = False
    If raw.Count = 1 Then
        For Each key In raw.Keys
            k = Replace(Replace(Replace(CStr(key), "~", "~~"), "*", "~*"), "?", "~?")
        Next key
        rg.AutoFilter Field:=filterField, Criteria1:="=" & k
    ElseIf raw.Count > 1 Then
        ReDim filterValues(0 To raw.Count - 1)
        filterIndex = 0
        For Each key In raw.Keys
            filterValues(filterIndex) = CStr(key)
            filterIndex = filterIndex + 1
        Next key
        criteria = filterValues
        rg.AutoFilter Field:=filterField, Criteria1:=criteria, Operator:=xlFilterValues
    Else
        k = "__ZEKON_NO_MATCH__"
        Do While target.Exists(k)
            k = k & "_"
        Loop
        rg.AutoFilter Field:=filterField, Criteria1:="=" & k
    End If
    ' Wlasne reguly CF oznaczaja stan z chwili wykonania i zachowuja wypelnienia.
    Set report = ws.Parent.Worksheets.Add(After:=ws.Parent.Worksheets(ws.Parent.Worksheets.Count))
    report.Cells(1, 1).Value2 = "Brakujaca pozycja"
    report.Cells(1, 2).Value2 = "Wiersz zrodlowy"
    report.Cells(1, 4).Value2 = ws.Name
    report.Cells(2, 4).Value2 = "Raport z " & Format$(Now, "yyyy-mm-dd hh:nn:ss")
    report.Columns(1).NumberFormat = "@": row = 2
    ProgressTick "Wyszukiwanie: lista", 0, sr - first + 1, True
    For r = first To sr
        ProgressTick "Wyszukiwanie: lista", r - first + 1, sr - first + 1
        If missing.Exists(CStr(r)) Then
            report.Cells(row, 1).Value2 = missing(CStr(r))
            report.Cells(row, 2).Value2 = r: row = row + 1
        End If
    Next r
    If row = 2 Then report.Cells(2, 1).Value2 = "Brak brakujacych pozycji"
    nm = OwnMissingName(ws, sc)
    Set marker = ws.Columns(sc)
    For idx = marker.FormatConditions.Count To 1 Step -1
        If marker.FormatConditions(idx).Type = xlExpression Then
            If InStr(1, marker.FormatConditions(idx).Formula1, nm, vbTextCompare) > 0 Then marker.FormatConditions(idx).Delete
        End If
    Next idx
    ' Porownanie tekstow nie zalezy od jezyka funkcji ani separatora listy.
    formula = "=""" & nm & """=""" & nm & """"
    blockStart = 0
    ProgressTick "Wyszukiwanie: lista", 0, sr - first + 1, True
    For r = first To sr
        ProgressTick "Wyszukiwanie: lista", r - first + 1, sr - first + 1
        If missing.Exists(CStr(r)) Then
            If blockStart = 0 Then blockStart = r
        ElseIf blockStart > 0 Then
            MarkMissingBlock ws.Range(ws.Cells(blockStart, sc), ws.Cells(r - 1, sc)), formula
            blockStart = 0
        End If
    Next r
    If blockStart > 0 Then MarkMissingBlock ws.Range(ws.Cells(blockStart, sc), ws.Cells(sr, sc)), formula
    report.Rows(1).Font.Bold = True: report.Columns("A:D").AutoFit
    bt.Activate
    SearchPositions = missing.Count
End Function

Private Sub MarkMissingBlock(ByVal cells As Range, ByVal formula As String)
    With cells.FormatConditions.Add(Type:=xlExpression, Formula1:=formula)
        .Interior.Color = RGB(255, 199, 206)
        .SetFirstPriority
        .StopIfTrue = False
    End With
End Sub

' Preflight may read a read-only source when changes will target a separate copy.
Public Sub ValidateSearch(ByVal ws As Worksheet, ByVal first As Long, ByVal p As Variant, ByVal visibleOnly As Boolean, Optional ByVal copyPlanned As Boolean = False)
    Dim bt As Worksheet, rg As Range, sc As Long, bc As Long, br As Long, r As Long
    Dim found As Boolean
    EnsureSearchAccess ws, copyPlanned
    sc = ColumnNumber(ws, CStr(p(1)))
    On Error Resume Next
    Set bt = ws.Parent.Worksheets(CStr(p(2)))
    On Error GoTo 0
    If bt Is Nothing Then Fail "Nie znaleziono arkusza: " & CStr(p(2))
    If bt Is ws Then Fail "Wybierz rozne arkusze: lista pozycji i zestawienie."
    EnsureSearchAccess bt, copyPlanned
    If ws.Parent.ProtectStructure Then Fail "Struktura skoroszytu jest chroniona."
    If bt.ListObjects.Count > 0 Then Fail "Wyszukiwanie wymaga zwyklego zakresu, nie tabeli Excela."
    If bt.Visible <> xlSheetVisible Then Fail "Arkusz zestawienia musi byc widoczny."
    bc = ColumnNumber(bt, CStr(p(3)))
    br = ResolveFirstRow(bt, CStr(p(3)), CStr(p(4)), False)
    Set rg = SearchFilterRange(bt, br, bc)
    p(4) = CStr(br)
    For r = first To LastRow(ws, sc)
        ProgressTick "Sprawdzanie listy", r - first + 1, 0
        If Included(ws, r, visibleOnly) Then
            If Len(PositionKey(ws.Cells(r, sc).Value2)) > 0 Then found = True: Exit For
        End If
    Next r
    If Not found Then Fail "Lista pozycji jest pusta. Sprawdz arkusz listy, kolumne i opcje widocznych wierszy."
End Sub

Private Sub EnsureSearchAccess(ByVal ws As Worksheet, ByVal copyPlanned As Boolean)
    If copyPlanned Then
        ' Sheet protection survives SaveCopyAs; do not remove or ignore it.
        If ws.ProtectContents Then Fail "Arkusz jest chroniony: " & ws.Name
    Else
        EnsureWritable ws
    End If
End Sub

Private Function SearchFilterRange(ByVal bt As Worksheet, ByVal firstData As Long, ByVal positionCol As Long) As Range
    Dim existing As Range, headerCell As Range, leftCol As Long, rightCol As Long, finish As Long, headerRow As Long
    If bt.AutoFilterMode Then Set existing = bt.AutoFilter.Range
    If firstData < 2 Then
        If Not existing Is Nothing Then
            Fail "Pierwszy wiersz zestawienia oznacza pierwszy wiersz DANYCH. Dla istniejacego filtra w arkuszu " & bt.Name & " wpisz " & (existing.Row + 1) & "."
        End If
        Fail "Pierwszy wiersz zestawienia oznacza pierwszy wiersz DANYCH i musi byc co najmniej 2."
    End If
    finish = LastRow(bt, positionCol)
    If finish < firstData Then Fail "Brak danych w kolumnie pozycji zestawienia."
    If Not existing Is Nothing Then
        If existing.Row = firstData - 1 Then
            leftCol = existing.Column
            rightCol = existing.Column + existing.Columns.Count - 1
            If positionCol < leftCol Then leftCol = positionCol
            If positionCol > rightCol Then rightCol = positionCol
            Set SearchFilterRange = bt.Range(bt.Cells(firstData - 1, leftCol), bt.Cells(finish, rightCol))
            Exit Function
        End If
    End If
    headerRow = firstData - 1
    Do While headerRow > 1
        If Application.CountA(bt.Rows(headerRow)) > 0 Then Exit Do
        headerRow = headerRow - 1
    Loop
    Set headerCell = bt.Rows(headerRow).Find(What:="*", After:=bt.Cells(headerRow, 1), LookIn:=xlFormulas, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, LookAt:=xlPart, MatchCase:=False, SearchFormat:=False)
    If headerCell Is Nothing Then Fail "Brak naglowkow bezposrednio nad danymi oraz zgodnego istniejacego filtra. Sprawdz pierwszy wiersz zestawienia."
    rightCol = headerCell.Column
    Set headerCell = bt.Rows(headerRow).Find(What:="*", After:=bt.Cells(headerRow, bt.Columns.Count), LookIn:=xlFormulas, SearchOrder:=xlByColumns, SearchDirection:=xlNext, LookAt:=xlPart, MatchCase:=False, SearchFormat:=False)
    leftCol = headerCell.Column
    If positionCol < leftCol Then leftCol = positionCol
    If positionCol > rightCol Then rightCol = positionCol
    Set SearchFilterRange = bt.Range(bt.Cells(headerRow, leftCol), bt.Cells(finish, rightCol))
End Function
