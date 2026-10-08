Option Explicit

' Read-only adapter: no assignments to the source, filters or Application.Calculation.
Public Function ProductionFields() As Variant
    ProductionFields = Array("Order", "Scope", "Position", "Lfd", "Profile", "Quantity", "UnitWeight", "TotalWeight", _
        "Length", "Width", "Height", "UnitSurface", "TotalSurface", "Drawing", "Revision", "Coating")
End Function

Public Function ProductionDictionary() As Object
    Set ProductionDictionary = CreateObject("Scripting.Dictionary")
    ProductionDictionary.CompareMode = vbBinaryCompare
End Function

Private Function CleanHeading(ByVal value As Variant) As String
    Dim s As String, i As Long, ch As String
    If IsError(value) Or IsEmpty(value) Then Exit Function
    s = UCase$(CStr(value))
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        Select Case ch
            Case " ", vbTab, vbCr, vbLf, ".", "_", "-", ChrW(160)
            Case Else: CleanHeading = CleanHeading & ch
        End Select
    Next i
End Function

Public Function ProductionHeaderField(ByVal value As Variant) As String
    Select Case CleanHeading(value)
        Case "ASSEMBLYPART": ProductionHeaderField = "Position"
        Case "LFD", "LFDNR": ProductionHeaderField = "Lfd"
        Case "LEADINGPART": ProductionHeaderField = "Profile"
        Case "PCS", "PCS/STK": ProductionHeaderField = "Quantity"
        Case "WEIGHT/PCS", "WEIGHTPERPCS": ProductionHeaderField = "UnitWeight"
        Case "WEIGHTTOT", "WEIGHTTOTAL": ProductionHeaderField = "TotalWeight"
        Case "ZEKONNUMER", "ZEKONNUMBER": ProductionHeaderField = "Scope"
        Case "LENGTH", "LENGHT": ProductionHeaderField = "Length"
        Case "WIDTH": ProductionHeaderField = "Width"
        Case "HEIGHT": ProductionHeaderField = "Height"
        Case "SURFACE/PCS", "SURFACEPERPCS": ProductionHeaderField = "UnitSurface"
        Case "SURFACETOT", "SURFACETOTAL": ProductionHeaderField = "TotalSurface"
        Case "DWGNR", "DRAWINGNO": ProductionHeaderField = "Drawing"
        Case "REVISION", "REV": ProductionHeaderField = "Revision"
        Case "CORRPROT/SYS", "CORRPROTSYS": ProductionHeaderField = "Coating"
    End Select
End Function

Public Function ProductionHeaders(ByVal ws As Worksheet, ByRef headerRow As Long) As Object
    Dim grid As Variant, row As Long, col As Long, lastCol As Long, candidate As Object, result As Object, field As String, required As Variant
    lastCol = ws.UsedRange.Column + ws.UsedRange.Columns.Count - 1
    grid = ws.Range(ws.Cells(1, 1), ws.Cells(50, lastCol)).Value2
    headerRow = 0
    For row = 1 To 50
        Set candidate = ProductionDictionary()
        For col = 1 To lastCol
            field = ProductionHeaderField(grid(row, col))
            If Len(field) > 0 Then
                If candidate.Exists(field) Then
                    candidate(field) = -1
                Else
                    candidate.Add field, col
                End If
            End If
        Next col
        If candidate.Exists("Position") And candidate.Exists("Quantity") Then
            If headerRow > 0 Then Fail "Kilka wierszy nagłówków Lieferliste. Wskaż jednoznaczny arkusz."
            headerRow = row: Set result = candidate
        End If
    Next row
    If headerRow = 0 Then Fail "Nie znaleziono ASSEMBLY PART i PCS. w pierwszych 50 wierszach."
    For Each required In result.Keys
        If result(required) < 1 Then Fail "Niejednoznaczny nagłówek: " & CStr(required) & ", wiersz " & headerRow
    Next required
    For Each required In Array("Position", "Lfd", "Profile", "Quantity", "UnitWeight", "TotalWeight")
        If Not result.Exists(required) Then Fail "Brak nagłówka Lieferliste: " & CStr(required)
    Next required
    Set ProductionHeaders = result
End Function

Public Function ProductionText(ByVal cell As Range) As String
    Dim value As Variant, fmt As String, i As Long, allZeros As Boolean
    value = cell.Value2
    If IsError(value) Then Fail "Błąd identyfikatora: " & cell.Parent.Name & "!" & cell.Address(False, False)
    If IsEmpty(value) Then Exit Function
    If VarType(value) = vbString Then
        ProductionText = Trim$(Replace(CStr(value), ChrW(160), " "))
    ElseIf IsNumeric(value) And VarType(value) <> vbBoolean Then
        If CDbl(value) <> Fix(CDbl(value)) Then Fail "Niecałkowity identyfikator: " & cell.Address(False, False)
        If Abs(CDbl(value)) >= 10 ^ 15 Then Fail "Identyfikator przekracza precyzję Excela. Zapisz jako tekst: " & cell.Address(False, False)
        fmt = CStr(cell.NumberFormat): allZeros = Len(fmt) > 0
        For i = 1 To Len(fmt)
            If Mid$(fmt, i, 1) <> "0" Then allZeros = False
        Next i
        If allZeros Then
            ProductionText = Format$(CDbl(value), fmt)
        ElseIf fmt = "General" Or fmt = "@" Then
            ProductionText = Format$(CDbl(value), "0")
        Else
            Fail "Niejednoznaczny format identyfikatora. Użyj tekstu: " & cell.Address(False, False)
        End If
    Else
        Fail "Niepoprawny typ identyfikatora: " & cell.Address(False, False)
    End If
    For i = 1 To Len(ProductionText)
        If AscW(Mid$(ProductionText, i, 1)) >= 0 And AscW(Mid$(ProductionText, i, 1)) < 32 Then Fail "Znak sterujący w identyfikatorze: " & cell.Address(False, False)
    Next i
End Function

Private Function FieldText(ByVal ws As Worksheet, ByVal row As Long, ByVal headers As Object, ByVal field As String) As String
    Dim value As Variant
    If Not headers.Exists(field) Then Exit Function
    value = ws.Cells(row, headers(field)).Value2
    If IsError(value) Then Fail "Błąd danych " & field & ", wiersz " & row
    If Not IsEmpty(value) Then FieldText = Trim$(Replace(CStr(value), ChrW(160), " "))
End Function

Private Function FieldNumber(ByVal ws As Worksheet, ByVal row As Long, ByVal headers As Object, ByVal field As String, ByVal required As Boolean) As Variant
    Dim value As Variant
    FieldNumber = Empty
    If headers.Exists(field) Then value = ws.Cells(row, headers(field)).Value2
    If IsError(value) Then Fail "Błąd Excela w polu " & field & ", wiersz " & row
    If IsEmpty(value) Or Len(Trim$(CStr(value))) = 0 Then
        If required Then Fail "Brak danych " & field & ", wiersz " & row
        Exit Function
    End If
    If VarType(value) = vbString Or VarType(value) = vbBoolean Then Fail "Pole " & field & " nie jest liczbą Excela, wiersz " & row
    If Not IsNumeric(value) Then Fail "Niepoprawna liczba " & field & ", wiersz " & row
    If CDbl(value) < 0 Then Fail "Ujemna wartość " & field & ", wiersz " & row
    FieldNumber = CDbl(value)
End Function

Public Function ProductionKey(ByVal order As String, ByVal scope As String, ByVal position As String, ByVal lfd As String) As String
    ProductionKey = Len(order) & ":" & order & Len(scope) & ":" & scope & Len(position) & ":" & position & Len(lfd) & ":" & lfd
End Function

Public Function ProductionSymbol(ByVal position As String, ByVal lfd As String) As String
    ProductionSymbol = position
    If Len(lfd) > 0 Then ProductionSymbol = position & "\" & lfd
End Function

Private Sub AddIssue(ByVal record As Object, ByVal issue As String)
    If Len(record("Control")) > 0 Then record("Control") = record("Control") & "; "
    record("Control") = record("Control") & issue
End Sub

Public Function ReadProductionSource(ByVal ws As Worksheet, Optional ByVal orderName As String = "", Optional ByVal missingScope As String = "") As Object
    Dim headers As Object, records As Object, record As Object, scopes As Object, result As Object
    Dim headerRow As Long, row As Long, last As Long, field As Variant, key As String, title As String, value As Variant
    Dim totalQuantity As Double, totalWeight As Double
    Set headers = ProductionHeaders(ws, headerRow)
    If Len(orderName) = 0 Then
        For Each field In Array("D1", "D2")
            If IsError(ws.Range(CStr(field)).Value2) Then Fail "Błąd nazwy zamówienia w " & CStr(field)
            title = Trim$(CStr(ws.Range(CStr(field)).Value2))
            If Len(title) > 0 Then
                If Len(orderName) > 0 And orderName <> title Then Fail "Niejednoznaczna nazwa zamówienia w D1/D2. Wskaż nazwę."
                orderName = title
            End If
        Next field
    End If
    If Len(Trim$(orderName)) = 0 Then Fail "Brak nazwy zamówienia. Wskaż nazwę."
    Set records = ProductionDictionary(): Set scopes = ProductionDictionary()
    last = ws.Cells(ws.Rows.Count, CLng(headers("Position"))).End(xlUp).Row
    For row = headerRow + 1 To last
        If Not ws.Rows(row).Hidden Then
            If Len(ProductionText(ws.Cells(row, CLng(headers("Position"))))) > 0 Then
                Set record = ProductionDictionary()
                record.Add "Order", orderName: record.Add "Scope", ""
                If headers.Exists("Scope") Then record("Scope") = ProductionText(ws.Cells(row, headers("Scope")))
                If Len(record("Scope")) = 0 Then record("Scope") = missingScope
                If Len(record("Scope")) = 0 Then Fail "Brak Zekon numer w wierszu " & row & ". Wskaż zakres dla pustych pól."
                record.Add "Position", ProductionText(ws.Cells(row, CLng(headers("Position"))))
                record.Add "Lfd", ProductionText(ws.Cells(row, CLng(headers("Lfd"))))
                record.Add "Profile", FieldText(ws, row, headers, "Profile")
                If Len(record("Profile")) = 0 Then Fail "Brak profilu w wierszu " & row
                For Each field In Array("Quantity", "UnitWeight", "TotalWeight")
                    value = FieldNumber(ws, row, headers, CStr(field), True): record.Add CStr(field), value
                Next field
                If record("Quantity") < 1 Or record("Quantity") <> Fix(record("Quantity")) Then Fail "PCS. musi być dodatnią liczbą całkowitą, wiersz " & row
                If Abs(record("Quantity") * record("UnitWeight") - record("TotalWeight")) > 0.0100001 Then Fail "Niezgodna masa PCS. × WEIGHT/PCS i WEIGHT TOT., wiersz " & row
                For Each field In Array("Length", "Width", "Height", "UnitSurface", "TotalSurface")
                    value = FieldNumber(ws, row, headers, CStr(field), False): record.Add CStr(field), value
                Next field
                For Each field In Array("Drawing", "Revision", "Coating")
                    record.Add CStr(field), FieldText(ws, row, headers, CStr(field))
                Next field
                record.Add "Control", "": record.Add "SourceRow", row
                If Len(record("Lfd")) > 0 And record("Quantity") <> 1 Then AddIssue record, "Lfd wypełnione przy PCS. innym niż 1 — sprawdź"
                If IsEmpty(record("UnitSurface")) Then
                    If headers.Exists("TotalSurface") Then
                        If ws.Cells(row, headers("TotalSurface")).HasFormula Then
                            If Not IsEmpty(record("TotalSurface")) Then
                                If record("TotalSurface") = 0 Then record("TotalSurface") = Empty
                            End If
                        End If
                    End If
                    AddIssue record, "Brak powierzchni jednostkowej"
                End If
                If IsEmpty(record("TotalSurface")) Then AddIssue record, "Brak powierzchni łącznej"
                If Not IsEmpty(record("UnitSurface")) And Not IsEmpty(record("TotalSurface")) Then
                    If Abs(record("UnitSurface") * record("Quantity") - record("TotalSurface")) > 0.0100001 Then AddIssue record, "Niezgodna powierzchnia jednostkowa i łączna"
                End If
                key = ProductionKey(record("Order"), record("Scope"), record("Position"), record("Lfd"))
                If records.Exists(key) Then Fail "Powtórzony klucz: " & ProductionSymbol(record("Position"), record("Lfd")) & "; wiersze " & records(key)("SourceRow") & " i " & row
                records.Add key, record
                If Not scopes.Exists(record("Scope")) Then scopes.Add record("Scope"), 0
                scopes(record("Scope")) = scopes(record("Scope")) + 1
                totalQuantity = totalQuantity + record("Quantity"): totalWeight = totalWeight + record("TotalWeight")
            ElseIf Len(FieldText(ws, row, headers, "Profile")) > 0 Then
                Fail "Brak ASSEMBLY PART przy danych w wierszu " & row
            End If
        End If
    Next row
    If records.Count = 0 Then Fail "Brak widocznych rekordów Lieferliste."
    Set result = ProductionDictionary()
    result.Add "Records", records: result.Add "Scopes", scopes: result.Add "Order", orderName
    result.Add "Quantity", totalQuantity: result.Add "Weight", totalWeight: result.Add "HeaderRow", headerRow
    Set ReadProductionSource = result
End Function

Public Function ProductionFilePart(ByVal value As String) As String
    Dim i As Long, ch As String, s As String, stem As String
    For i = 1 To Len(value)
        ch = Mid$(value, i, 1)
        If InStr(1, "<>:""/\|?*", ch, vbBinaryCompare) > 0 Or (AscW(ch) >= 0 And AscW(ch) < 32) Then ch = "_"
        s = s & ch
    Next i
    s = Trim$(s)
    Do While Len(s) > 0
        If Right$(s, 1) <> "." And Right$(s, 1) <> " " Then Exit Do
        s = Left$(s, Len(s) - 1)
    Loop
    If Len(s) = 0 Then Fail "Nazwa pliku jest pusta po oczyszczeniu."
    stem = UCase$(Split(s, ".")(0))
    Select Case stem
        Case "CON", "PRN", "AUX", "NUL", "COM1", "COM2", "COM3", "COM4", "COM5", "COM6", "COM7", "COM8", "COM9", _
             "LPT1", "LPT2", "LPT3", "LPT4", "LPT5", "LPT6", "LPT7", "LPT8", "LPT9": s = "_" & s
    End Select
    ProductionFilePart = s
End Function

Public Function ProductionFileNames(ByVal imported As Object) As Object
    Dim scope As Variant, operation As Variant, name As String, names As Object, owners As Object
    Set names = ProductionDictionary(): Set owners = ProductionDictionary(): owners.CompareMode = vbTextCompare
    For Each scope In imported("Scopes").Keys
        For Each operation In Array("składanie", "spawanie", "malowanie", "załadunki")
            name = CStr(operation) & " - " & ProductionFilePart(imported("Order")) & "_" & ProductionFilePart(CStr(scope)) & ".xlsx"
            If owners.Exists(name) Then Fail "Kolizja nazw plików po oczyszczeniu: " & name & ". Zakresy " & owners(name) & " oraz " & scope
            owners.Add name, scope: names.Add CStr(operation) & "|" & CStr(scope), name
        Next operation
    Next scope
    Set ProductionFileNames = names
End Function
