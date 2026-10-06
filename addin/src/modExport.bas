Option Explicit

Public Function ExportRows(ByVal ws As Worksheet, ByVal first As Long, ByVal p As Variant, ByVal visibleOnly As Boolean, ByVal painting As Boolean, Optional ByVal outputPath As String = "") As Long
    Dim cols(1 To 6) As Long, i As Long, r As Long, finish As Long, n As Long, outRow As Long
    Dim data() As Variant, v As Variant, result As Workbook, dest As Worksheet, headings As Variant, sourceData As Variant, leftCol As Long, rightCol As Long
    If Not painting Then ValidateZincTarget outputPath
    For i = 1 To 6
        cols(i) = ColumnNumber(ws, CStr(p(i + 1)))
    Next i
    If painting Then
        outRow = WholeNumber(CStr(p(8)), "Pierwszy wiersz eksportu", ws.Rows.Count)
    Else
        outRow = 2
    End If
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
    If Not painting Then CheckZincCapacity n
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
    If painting Then
        Set result = Application.Workbooks.Add(xlWBATWorksheet)
        Set dest = result.Worksheets(1)
        dest.Name = "Malowanie"
        headings = Array("Auf.Name", "Auftr.", "Pos.", "Profil", "Lange", "Gewicht", "Lfn nr")
        For i = 0 To 6
            dest.Cells(1, i + 1).Value2 = headings(i)
        Next i
        dest.Columns("A:C").NumberFormat = "@"
        dest.Cells(outRow, 1).Resize(n, 7).Value2 = data
        dest.Rows(1).Font.Bold = True
        dest.Columns("A:G").AutoFit
    Else
        ProgressTick "Tworzenie wzor.xls", 0, 0, True
        Set result = NewZincWorkbook()
        Set dest = result.Worksheets(1)
        ' Temporary Text prevents numeric IDs/formula-like names being reinterpreted.
        ' Restoring General keeps stored strings (including leading zeros) intact.
        dest.Cells(2, 1).Resize(n, 4).NumberFormat = "@"
        dest.Cells(2, 1).Resize(n, 7).Value2 = data
        dest.Cells(2, 1).Resize(n, 8).NumberFormat = "General"
        dest.Rows("2:" & CStr(n + 1)).RowHeight = dest.Rows(1).RowHeight
        ValidateZincTarget outputPath
        ProgressTick "Zapis wzor.xls (Excel 97-2003)", 0, 0, True
        result.CheckCompatibility = False
        result.SaveAs Filename:=outputPath, FileFormat:=xlExcel8, CreateBackup:=False, AddToMru:=False
        result.Worksheets(1).Activate
    End If
    ExportRows = n
End Function

Public Sub CheckZincCapacity(ByVal count As Double)
    If count > 65535 Then Fail "Format wzorca XLS miesci maksymalnie 65535 pozycji i naglowek. Ogranicz eksport filtrem; nie zapisano niepelnego wyniku."
End Sub

Public Function ChooseZincExportPath(ByVal source As Workbook) As String
    Dim picker As Object, folder As String
    Set picker = Application.FileDialog(4)
    picker.Title = "Wybierz folder zapisu pliku wzor.xls"
    picker.AllowMultiSelect = False
    If Len(source.Path) > 0 And InStr(source.Path, "://") = 0 Then picker.InitialFileName = source.Path & Application.PathSeparator
    If picker.Show <> -1 Then Exit Function
    folder = picker.SelectedItems(1)
    If Right$(folder, 1) <> Application.PathSeparator Then folder = folder & Application.PathSeparator
    ChooseZincExportPath = folder & "wzor.xls"
    ValidateZincTarget ChooseZincExportPath
End Function

Public Sub ValidateZincTarget(ByVal path As String)
    Dim openBook As Workbook, separator As Long, folder As String
    separator = InStrRev(path, Application.PathSeparator)
    If separator = 0 Then Fail "Wybierz folder zapisu pliku wzor.xls."
    If StrComp(Mid$(path, separator + 1), "wzor.xls", vbBinaryCompare) <> 0 Then Fail "Plik eksportu ocynku musi nazywac sie wzor.xls."
    folder = Left$(path, separator)
    If (GetAttr(folder) And vbDirectory) = 0 Then Fail "Nieprawidlowy folder eksportu."
    If Len(Dir$(path, vbNormal Or vbReadOnly Or vbHidden Or vbSystem Or vbDirectory)) > 0 Then Fail "Plik wzor.xls juz istnieje w tym folderze. Wybierz inny folder albo przenies poprzedni wynik. Nie nadpisano pliku."
    For Each openBook In Application.Workbooks
        If StrComp(openBook.Name, "wzor.xls", vbTextCompare) = 0 Then Fail "Zamknij poprzedni plik wzor.xls przed kolejnym eksportem."
    Next openBook
End Sub

Public Sub ValidateZincExportSize(ByVal ws As Worksheet, ByVal first As Long, ByVal p As Variant, ByVal visibleOnly As Boolean, ByVal splitFirst As Boolean)
    Dim r As Long, finish As Long, col As Long, qtyCol As Long, count As Double, qty As Variant
    Dim positions As Variant, quantities As Variant, position As Variant
    col = ColumnNumber(ws, CStr(p(3)))
    If splitFirst Then qtyCol = ColumnNumber(ws, CStr(p(9)))
    finish = LastRow(ws, col)
    If finish < first Then Fail "Brak pozycji do eksportu."
    positions = ws.Range(ws.Cells(first, col), ws.Cells(finish, col)).Value2
    If splitFirst Then quantities = ws.Range(ws.Cells(first, qtyCol), ws.Cells(finish, qtyCol)).Value2
    ProgressTick "Sprawdzanie limitu XLS", 0, finish - first + 1, True
    For r = first To finish
        ProgressTick "Sprawdzanie limitu XLS", r - first + 1, finish - first + 1
        If Included(ws, r, visibleOnly) Then
            If finish = first Then position = positions Else position = positions(r - first + 1, 1)
            If Len(PositionKey(position)) > 0 Then
                If splitFirst Then
                    If finish = first Then qty = quantities Else qty = quantities(r - first + 1, 1)
                    If IsError(qty) Then Fail "Blad ilosci w wierszu " & r
                    If Not IsNumeric(qty) Then Fail "Ilosc nie jest liczba, wiersz " & r
                    If CDbl(qty) < 1 Or CDbl(qty) <> Fix(CDbl(qty)) Then Fail "Niepoprawna ilosc, wiersz " & r
                    count = count + CDbl(qty)
                Else
                    count = count + 1
                End If
                CheckZincCapacity count
            End If
        End If
    Next r
End Sub

Private Function NewZincWorkbook() As Workbook
    Dim folder As String, path As String, index As Long, number As Long, message As String
    folder = Environ$("TEMP") & "\ZekonZinc_" & Format$(Now, "yyyymmdd_hhnnss")
    Do While Len(Dir$(folder & "_" & CStr(index), vbDirectory)) > 0
        index = index + 1
    Loop
    folder = folder & "_" & CStr(index)
    On Error GoTo Bad
    MkDir folder
    path = folder & "\zinc-template.xls"
    WriteZincTemplate path
    Set NewZincWorkbook = Application.Workbooks.Add(Template:=path)
    Kill path
    RmDir folder
    Exit Function
Bad:
    number = Err.Number: message = Err.Description
    On Error Resume Next
    If Len(path) > 0 Then Kill path
    RmDir folder
    On Error GoTo 0
    Err.Raise number, "NewZincWorkbook", message
End Function
