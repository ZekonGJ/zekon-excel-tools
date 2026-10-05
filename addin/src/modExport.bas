Option Explicit

Public Function ExportRows(ByVal ws As Worksheet, ByVal first As Long, ByVal p As Variant, ByVal visibleOnly As Boolean, ByVal painting As Boolean) As Long
    Dim cols(1 To 6) As Long, i As Long, r As Long, finish As Long, n As Long, outRow As Long
    Dim data() As Variant, v As Variant, result As Workbook, dest As Worksheet, headings As Variant, sourceData As Variant, leftCol As Long, rightCol As Long
    For i = 1 To 6
        cols(i) = ColumnNumber(ws, CStr(p(i + 1)))
    Next i
    outRow = WholeNumber(CStr(p(8)), "Pierwszy wiersz eksportu", ws.Rows.Count)
    If outRow < 2 Then Fail "Eksport musi zaczynac sie od wiersza 2 lub dalszego."
    If cols(4) = cols(6) Then Fail "Dlugosc i numer sztuki wskazuja te sama kolumne. Wybierz poprawne kolumny."
    finish = LastRow(ws, cols(2))
    If finish < first Then Fail "Brak pozycji do eksportu."
    leftCol = cols(1): rightCol = cols(1)
    For i = 2 To 6
        If cols(i) < leftCol Then leftCol = cols(i)
        If cols(i) > rightCol Then rightCol = cols(i)
    Next i
    sourceData = ws.Range(ws.Cells(first, leftCol), ws.Cells(finish, rightCol)).Value2
    ProgressTick "Eksport", 0, finish - first + 1, True
    For r = first To finish
        ProgressTick "Eksport", r - first + 1, finish - first + 1
        If Included(ws, r, visibleOnly) Then
            If Len(PositionKey(sourceData(r - first + 1, cols(2) - leftCol + 1))) > 0 Then
                For i = 1 To 6
                    If IsError(sourceData(r - first + 1, cols(i) - leftCol + 1)) Then Fail "Blad Excela w wierszu " & r
                Next i
                v = sourceData(r - first + 1, cols(5) - leftCol + 1)
                If Len(CStr(v)) > 0 Then
                    If Not IsNumeric(v) Then Fail "Waga nie jest liczba, wiersz " & r
                End If
                n = n + 1
            End If
        End If
    Next r
    If n = 0 Then Fail "Brak pozycji do eksportu."
    If n + outRow - 1 > ws.Rows.Count Then Fail "Za duzo wierszy eksportu."
    ReDim data(1 To n, 1 To 7): n = 0
    ProgressTick "Eksport", 0, finish - first + 1, True
    For r = first To finish
        ProgressTick "Eksport", r - first + 1, finish - first + 1
        If Included(ws, r, visibleOnly) Then
            If Len(PositionKey(sourceData(r - first + 1, cols(2) - leftCol + 1))) > 0 Then
                n = n + 1: data(n, 1) = CStr(p(1))
                For i = 1 To 6
                    data(n, i + 1) = sourceData(r - first + 1, cols(i) - leftCol + 1)
                Next i
                data(n, 2) = CStr(sourceData(r - first + 1, cols(1) - leftCol + 1))
                data(n, 3) = CStr(sourceData(r - first + 1, cols(2) - leftCol + 1))
                v = sourceData(r - first + 1, cols(5) - leftCol + 1)
                If Len(CStr(v)) > 0 Then
                    data(n, 6) = Application.WorksheetFunction.Round(CDbl(v), 0)
                Else
                    data(n, 6) = Empty
                End If
            End If
        End If
    Next r
    Set result = Application.Workbooks.Add(xlWBATWorksheet)
    Set dest = result.Worksheets(1)
    If painting Then
        dest.Name = "Malowanie"
        headings = Array("Auf.Name", "Auftr.", "Pos.", "Profil", "Lange", "Gewicht", "Lfn nr")
    Else
        dest.Name = "Ocynk"
        headings = Array("Auf. Name", "Auftr.", "Pos.", "Profil", "Lange", "Gewicht", "Lfn nr")
    End If
    For i = 0 To 6
        dest.Cells(1, i + 1).Value2 = headings(i)
    Next i
    dest.Columns("A:C").NumberFormat = "@"
    dest.Cells(outRow, 1).Resize(n, 7).Value2 = data
    dest.Rows(1).Font.Bold = True
    dest.Columns("A:G").AutoFit
    ExportRows = n
End Function
