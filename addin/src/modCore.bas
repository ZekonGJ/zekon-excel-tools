Option Explicit

Public Const APP_ID As String = "ZekonTools"
Public Const APP_VERSION As String = "1.0.0-rc2"
Public LastCopyPath As String
Private progressForm As Object
Private progressActive As Boolean, cancelRequested As Boolean
Private previousCalculation As XlCalculation
Private previousStatus As Variant, previousStatusVisible As Boolean
Private lastTick As Single, workStarted As Single

Public Sub ZekonPanel()
    frmZekon.Show
End Sub

Public Sub Auto_Open()
    InstallMenu
End Sub

Public Sub Auto_Close()
    On Error Resume Next
    Application.CommandBars("Worksheet Menu Bar").Controls("ZEKON").Delete
    On Error GoTo 0
End Sub

Public Sub InstallMenu()
    Dim p As Object, b As Object
    Auto_Close
    Set p = Application.CommandBars("Worksheet Menu Bar").Controls.Add(Type:=10, Temporary:=True)
    p.Caption = "ZEKON"
    Set b = p.Controls.Add(Type:=1, Temporary:=True)
    b.Caption = "Otworz Narzedzia ZEKON"
    b.OnAction = "'" & Replace(ThisWorkbook.FullName, "'", "''") & "'!modCore.ZekonPanel"
End Sub

Public Sub Fail(ByVal message As String)
    Err.Raise vbObjectError + 701, "Narzedzia ZEKON", message
End Sub

Public Function WholeNumber(ByVal value As String, ByVal title As String, ByVal maxValue As Long) As Long
    Dim n As Double
    If Not IsNumeric(value) Then Fail title & ": wpisz liczbe calkowita."
    n = CDbl(value)
    If n < 1 Or n > maxValue Or n <> Fix(n) Then Fail title & ": niepoprawna wartosc."
    WholeNumber = CLng(n)
End Function

Public Function ColumnNumber(ByVal ws As Worksheet, ByVal value As String) As Long
    Dim i As Long, s As String, n As Long
    s = UCase$(Trim$(value))
    If Len(s) = 0 Or Len(s) > 3 Then Fail "Niepoprawna kolumna: " & value
    For i = 1 To Len(s)
        If Mid$(s, i, 1) < "A" Or Mid$(s, i, 1) > "Z" Then Fail "Podaj litere kolumny: " & value
        n = n * 26 + Asc(Mid$(s, i, 1)) - 64
    Next i
    If n > ws.Columns.Count Then Fail "Kolumna poza arkuszem: " & value
    ColumnNumber = n
End Function

Public Function LastRow(ByVal ws As Worksheet, ByVal col As Long) As Long
    LastRow = ws.Cells(ws.Rows.Count, col).End(xlUp).Row
End Function

Public Function Included(ByVal ws As Worksheet, ByVal row As Long, ByVal visibleOnly As Boolean) As Boolean
    Included = True
    If visibleOnly Then Included = Not ws.Rows(row).Hidden
End Function

Public Function PositionKey(ByVal value As Variant) As String
    If IsError(value) Then Fail "Blad Excela w kolumnie pozycji. Popraw dane zrodlowe."
    PositionKey = Trim$(Replace(CStr(value), ChrW(160), " "))
End Function

Public Function CopyWorkbook(ByVal source As Workbook) As Workbook
    Dim folder As String, path As String, ext As String, security As Long, i As Long
    Dim e As Long, message As String
    Select Case source.FileFormat
        Case 51: ext = ".xlsx"
        Case 52: ext = ".xlsm"
        Case 50: ext = ".xlsb"
        Case 56, -4143: ext = ".xls"
        Case Else: Fail "Kopia tego formatu nie jest obslugiwana. Zapisz dane jako XLSX/XLSM/XLSB/XLS."
    End Select
    folder = Environ$("LOCALAPPDATA") & "\ZekonTools"
    If Dir$(folder, vbDirectory) = "" Then MkDir folder
    folder = folder & "\Kopie"
    If Dir$(folder, vbDirectory) = "" Then MkDir folder
    Do
        i = i + 1
        path = folder & "\ZEKON_" & Format$(Now, "yyyymmdd_hhnnss") & "_" & i & ext
    Loop While Len(Dir$(path)) > 0
    ProgressTick "Zapisywanie kopii skoroszytu", 0, 0, True
    source.SaveCopyAs path
    LastCopyPath = path
    security = Application.AutomationSecurity
    On Error GoTo Bad
    Application.AutomationSecurity = 3
    ProgressTick "Otwieranie kopii skoroszytu", 0, 0, True
    Set CopyWorkbook = Application.Workbooks.Open(Filename:=path, UpdateLinks:=0, ReadOnly:=False)
    Application.AutomationSecurity = security
    Exit Function
Bad:
    e = Err.Number: message = Err.Description
    Application.AutomationSecurity = security
    Err.Raise e, "CopyWorkbook", message
End Function

Public Sub EnsureWritable(ByVal ws As Worksheet)
    If ws.ProtectContents Then Fail "Arkusz jest chroniony."
    If ws.Parent.ReadOnly Then Fail "Skoroszyt jest tylko do odczytu. Wybierz prace na kopii."
End Sub

Public Function OwnMissingName(ByVal ws As Worksheet, ByVal col As Long) As String
    OwnMissingName = "_ZEKON_missing_" & ws.CodeName & "_" & CStr(col)
End Function

' AUTO is evaluated afresh for the selected worksheet and column on every run.
Public Function ResolveFirstRow(ByVal ws As Worksheet, ByVal column As String, ByVal setting As String, Optional ByVal numericData As Boolean = False) As Long
    Dim col As Long, lo As ListObject, rg As Range, hit As Range, firstCell As Range
    Dim n As Long, firstAddress As String, headerRow As Long, candidate As Long
    col = ColumnNumber(ws, column)
    If Len(Trim$(setting)) > 0 And UCase$(Trim$(setting)) <> "AUTO" Then
        ResolveFirstRow = WholeNumber(setting, "Pierwszy wiersz danych", ws.Rows.Count)
        Exit Function
    End If
    ' More than one table in the column is ambiguous: never silently choose one.
    For Each lo In ws.ListObjects
        If col >= lo.Range.Column And col < lo.Range.Column + lo.Range.Columns.Count Then
            If candidate <> 0 Then Fail "Kilka tabel w kolumnie " & column & ". Wskaz pierwszy wiersz zamiast AUTO."
            If lo.DataBodyRange Is Nothing Then Fail "Tabela w kolumnie " & column & " jest pusta."
            candidate = lo.DataBodyRange.Row
        End If
    Next lo
    If candidate > 0 Then ResolveFirstRow = candidate: Exit Function
    If ws.AutoFilterMode Then
        Set rg = ws.AutoFilter.Range
        If col >= rg.Column And col < rg.Column + rg.Columns.Count Then
            ResolveFirstRow = rg.Row + 1
            Exit Function
        End If
    End If
    Set rg = ws.Columns(col)
    Set hit = rg.Find(What:="*", After:=ws.Cells(ws.Rows.Count, col), LookIn:=xlValues, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False)
    If hit Is Nothing Then Fail "Kolumna " & column & " w arkuszu " & ws.Name & " jest pusta."
    Set firstCell = hit
    firstAddress = hit.Address
    ' Inspect up to 50 nonempty cells, including headers after blank title rows.
    For n = 1 To 50
        If IsDataHeader(hit.Value2, numericData) Then
            headerRow = hit.Row
            Exit For
        End If
        Set hit = rg.Find(What:="*", After:=hit, LookIn:=xlValues, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False)
        If hit Is Nothing Then Exit For
        If hit.Address = firstAddress Then Exit For
    Next n
    If headerRow > 0 Then
        Set hit = rg.Find(What:="*", After:=ws.Cells(headerRow, col), LookIn:=xlValues, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False)
        If Not hit Is Nothing Then
            If hit.Row > headerRow Then
                If IsDataValue(hit, numericData) Then ResolveFirstRow = hit.Row: Exit Function
            End If
        End If
    Else
        If IsDataValue(firstCell, numericData) Then ResolveFirstRow = firstCell.Row: Exit Function
    End If
    Fail "Nie mozna jednoznacznie rozpoznac poczatku danych: " & ws.Name & " / " & column & ". Wpisz numer wiersza zamiast AUTO lub dodaj naglowek/filtr."
End Function

Private Function IsDataValue(ByVal cell As Range, ByVal numericData As Boolean) As Boolean
    Dim v As Variant, t As String
    v = cell.Value2
    If IsError(v) Or IsEmpty(v) Then Exit Function
    If cell.MergeCells Or VarType(v) = vbBoolean Then Exit Function
    t = Trim$(CStr(v))
    If Len(t) = 0 Then Exit Function
    If cell.HasFormula Then
        If InStr(1, cell.Formula, "SUBTOTAL(", vbTextCompare) > 0 Then Exit Function
        If InStr(1, cell.Formula, "SUM(", vbTextCompare) > 0 Then Exit Function
        If InStr(1, cell.Formula, "AGGREGATE(", vbTextCompare) > 0 Then Exit Function
    End If
    If numericData Then
        IsDataValue = IsNumeric(v)
    Else
        ' Headerless position lists: numeric IDs or codes containing digits.
        IsDataValue = (t Like "*[0-9]*")
    End If
End Function

Private Function IsDataHeader(ByVal v As Variant, ByVal numericData As Boolean) As Boolean
    Dim t As String
    If IsError(v) Or IsEmpty(v) Then Exit Function
    t = UCase$(Trim$(Replace(Replace(CStr(v), ChrW(160), " "), vbLf, " ")))
    t = Replace(t, ChrW(211), "O"): t = Replace(t, ChrW(321), "L")
    t = Replace(t, ChrW(346), "S"): t = Replace(t, ChrW(262), "C")
    t = Replace(t, ChrW(280), "E"): t = Replace(t, ChrW(260), "A")
    If numericData Then
        Select Case t
            Case "PCS.", "PCS", "QTY", "QUANTITY", "ILOSC", "SZT", "SZT.", "ILOSC SZTUK", "WEIGHT/PCS", "WEIGHT TOT.", "WEIGHT", "WAGA", "MASA", "MASA [KG]", "WAGA [KG]", "LENGTH", "LENGHT", "DLUGOSC", "WIDTH", "HEIGHT", "SURFACE/PCS", "SURFACE TOT.", "WARTOSC", "VALUE"
                IsDataHeader = True
        End Select
    Else
        Select Case t
            Case "ASSEMBLY PART", "LEADING PART", "POZYCJA", "POZYCJE", "NR POZYCJI", "NUMER POZYCJI", "NUMERY POZYCJI", "POSITION", "POSITIONS", "POS", "POS.", "PART NUMBER", "MARK", "ELEMENT", "NR ELEMENTU"
                IsDataHeader = True
        End Select
    End If
End Function

Public Sub BeginWork(ByVal panel As Object)
    Set progressForm = panel
    previousCalculation = Application.Calculation
    previousStatus = Application.StatusBar
    previousStatusVisible = Application.DisplayStatusBar
    progressActive = True: cancelRequested = False
    workStarted = Timer: lastTick = -1
    Application.DisplayStatusBar = True
    Application.Calculation = xlCalculationManual
    ProgressTick "Sprawdzanie ustawien", 0, 0, True
End Sub

Public Sub EndWork()
    If Not progressActive Then Exit Sub
    ' Restore even when calculation or UI cleanup raises an error.
    On Error Resume Next
    cancelRequested = False
    If Not progressForm Is Nothing Then
        progressForm.lblStatus.Caption = "Konczenie pracy / przeliczanie..."
        progressForm.Repaint
    End If
    Application.Calculation = previousCalculation
    Application.StatusBar = previousStatus
    Application.DisplayStatusBar = previousStatusVisible
    Set progressForm = Nothing
    progressActive = False
    On Error GoTo 0
End Sub

Public Sub RequestWorkCancel()
    cancelRequested = True
End Sub

Public Sub ProgressTick(ByVal phase As String, ByVal done As Long, ByVal total As Long, Optional ByVal force As Boolean = False)
    Dim nowTime As Single, elapsed As Long, ratio As Double, message As String
    If Not progressActive Then Exit Sub
    If cancelRequested Then Fail "Przerwano na zyczenie uzytkownika. Wynik moze byc czesciowy."
    If Not force Then
        If done Mod 64 <> 0 And done <> total Then Exit Sub
    End If
    nowTime = Timer
    If Not force And lastTick >= 0 And nowTime >= lastTick Then
        If nowTime - lastTick < 0.2 Then Exit Sub
    End If
    lastTick = nowTime
    elapsed = CLng(nowTime - workStarted)
    If elapsed < 0 Then elapsed = elapsed + 86400
    message = phase
    If total > 0 Then
        ratio = done / CDbl(total)
        If ratio > 1 Then ratio = 1
        message = message & " " & Format$(ratio, "0%") & " (" & done & "/" & total & ")"
    Else
        message = message & " - prosze czekac"
    End If
    message = message & " | " & elapsed & " s"
    Application.StatusBar = "ZEKON: " & message
    If Not progressForm Is Nothing Then
        progressForm.lblStatus.Caption = message
        progressForm.progressFill.Width = 568 * ratio
        progressForm.Repaint
    End If
    DoEvents
    If cancelRequested Then Fail "Przerwano na zyczenie uzytkownika. Wynik moze byc czesciowy."
End Sub

Public Function ReadColumnValues(ByVal ws As Worksheet, ByVal first As Long, ByVal finish As Long, ByVal col As Long) As Variant
    Dim singleValue(1 To 1, 1 To 1) As Variant
    If finish = first Then
        singleValue(1, 1) = ws.Cells(first, col).Value2
        ReadColumnValues = singleValue
    Else
        ReadColumnValues = ws.Range(ws.Cells(first, col), ws.Cells(finish, col)).Value2
    End If
End Function
