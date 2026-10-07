Option Explicit

Private Sub AssertTrue(ByVal condition As Boolean, ByVal message As String)
    If Not condition Then Fail "TEST: " & message
End Sub

Public Function ZekonSelfTest() As String
    Dim wb As Workbook, ws As Worksheet, bt As Worksheet, result As Workbook
    Dim n As Long, p(1 To 10) As Variant, errCode As Long, oldAlerts As Boolean
    Dim e As Long, message As String, expectedFilter As String
    On Error GoTo Bad
    oldAlerts = Application.DisplayAlerts
    Set wb = Workbooks.Add(xlWBATWorksheet): Set ws = wb.Worksheets(1)
    TestAutoStart wb
    TestSplitWithFilter wb
    TestWorkLifecycle
    TestZincExports
    TestWorkingCopies
    ws.Name = "Test"
    ws.Range("F2").Value2 = 2.5: ws.Range("F3").Formula = "=2.5"
    ws.Range("F4").Value2 = -2.5: ws.Rows(4).Hidden = True
    n = RoundColumn(ws, 2, "F", True, False)
    AssertTrue n = 1, "Zaokraglanie: liczba zmian"
    AssertTrue ws.Range("F2").Value2 = 3, "Zaokraglanie polowek"
    AssertTrue ws.Range("F3").HasFormula, "Ochrona formul"
    AssertTrue ws.Range("F4").Value2 = -2.5, "Pominiecie ukrytych"
    ws.Rows(4).Hidden = False
    ws.Range("A2").Value2 = "P1": ws.Range("E2").Value2 = 3
    ws.Range("H2").Formula = "=E2*10"
    ws.Rows(2).RowHeight = 27
    ws.Range("A3").Value2 = "P2": ws.Range("E3").Value2 = 2
    ws.Rows(3).Hidden = True
    n = SplitRows(ws, 2, "E", "C", True)
    AssertTrue n = 3, "Rozbijanie: liczba sztuk"
    AssertTrue ws.Range("C4").Value2 = 3, "Numer ostatniej sztuki"
    AssertTrue ws.Range("H3").Formula = "=E3*10" And ws.Range("H4").Formula = "=E4*10", "Blokowe kopiowanie formul wzglednych"
    AssertTrue ws.Rows(3).RowHeight = 27 And ws.Rows(4).RowHeight = 27, "Wysokosc kopiowanych wierszy"
    AssertTrue ws.Range("E2").Value2 = 1 And ws.Range("E4").Value2 = 1, "Ilosc po rozbiciu"
    AssertTrue ws.Range("E5").Value2 = 2 And ws.Rows(5).Hidden, "Ukryty wiersz bez rozbicia"
    ws.Range("E2").Value2 = 0
    On Error Resume Next
    n = SplitRows(ws, 2, "E", "C", True)
    errCode = Err.Number: Err.Clear
    On Error GoTo Bad
    AssertTrue errCode <> 0, "Odrzucenie zerowej ilosci"
    ws.Range("E2").Value2 = 1
    ws.Range("D2:D4").NumberFormat = "@": ws.Range("D2:D4").Value2 = "001"
    ws.Range("B2:B4").Value2 = "K1": ws.Range("M2:M4").Value2 = "HEA100"
    ws.Range("H2:H4").Value2 = 1000: ws.Range("F2:F4").Value2 = 2.5
    p(1) = "Test": p(2) = "B": p(3) = "D": p(4) = "M"
    p(5) = "H": p(6) = "F": p(7) = "C": p(8) = "3"
    n = ExportRows(ws, 2, p, True, True)
    Set result = ActiveWorkbook
    AssertTrue n = 3, "Eksport widocznych"
    AssertTrue result.Worksheets(1).Range("C3").Value2 = "001", "Zera wiodace"
    AssertTrue result.Worksheets(1).Range("F3").Value2 = 3, "Waga eksportu"
    AssertTrue IsEmpty(result.Worksheets(1).Range("A2").Value2), "Wiersz 2 malowania"
    result.Close SaveChanges:=False: Set result = Nothing
    Set bt = wb.Worksheets.Add: bt.Name = "bt-list"
    bt.Range("B7").Value2 = "Pozycja": bt.Range("B8").Value2 = "P1"
    bt.Range("B9").Value2 = "P3": ws.Rows(5).Hidden = False
    p(1) = "A": p(2) = "bt-list": p(3) = "B": p(4) = "8"
    n = SearchPositions(ws, 2, p, True)
    AssertTrue n = 1, "Wykrywanie brakow"
    AssertTrue Not bt.Rows(8).Hidden And bt.Rows(9).Hidden, "Filtr pozycji"
    AssertTrue ws.Cells(5, 1).FormatConditions.Count > 0, "Oznaczenie braku"
    bt.Range("B10").Value2 = "P2"
    n = SearchPositions(ws, 2, p, True)
    AssertTrue n = 0, "Kilka dopasowan bez brakow"
    AssertTrue Not bt.Rows(8).Hidden And bt.Rows(9).Hidden And Not bt.Rows(10).Hidden, "Filtr wielu pozycji"
    AssertTrue ws.Cells(5, 1).FormatConditions.Count = 0, "Usuniecie nieaktualnego oznaczenia"
    bt.Range("B8").Value2 = "X1": bt.Range("B10").Value2 = "X2"
    n = SearchPositions(ws, 2, p, True)
    AssertTrue n = 4, "Brak dopasowan"
    AssertTrue bt.Rows(8).Hidden And bt.Rows(9).Hidden And bt.Rows(10).Hidden, "Pusty wynik filtra"
    ' Reproduce LieferlisteZMSCG: headers in 6, blank filter row 7, data from 8.
    bt.AutoFilterMode = False
    bt.Range("B7").ClearContents
    bt.Range("B6").Value2 = "ASSEMBLY PART"
    bt.Range("B8").Value2 = "P1": bt.Range("B10").Value2 = "P2"
    bt.Range("A7:BD10").AutoFilter
    expectedFilter = bt.AutoFilter.Range.Address(True, True, xlA1)
    AssertTrue bt.AutoFilter.Range.Row = 7 And bt.AutoFilter.Range.Rows.Count = 4, "Wiersze filtra przed wyszukiwaniem"
    p(4) = "1": errCode = 0
    On Error Resume Next
    ValidateSearch ws, 2, p, True
    errCode = Err.Number: Err.Clear
    On Error GoTo Bad
    AssertTrue errCode <> 0, "Preflight odrzuca pierwszy wiersz 1"
    p(4) = "8"
    ValidateSearch ws, 2, p, True
    n = SearchPositions(ws, 2, p, True)
    AssertTrue n = 0, "Pusty wiersz naglowka istniejacego filtra"
    AssertTrue bt.AutoFilter.Range.Address(True, True, xlA1) = expectedFilter, "Zachowany zakres filtra: oczekiwany " & expectedFilter & ", otrzymany " & bt.AutoFilter.Range.Address(True, True, xlA1)
    AssertTrue Not bt.Rows(8).Hidden And bt.Rows(9).Hidden And Not bt.Rows(10).Hidden, "Filtr ukladu Lieferliste"
    wb.Close SaveChanges:=False: Set wb = Nothing
    TestPanelVersion
    Application.DisplayAlerts = oldAlerts
    ZekonSelfTest = "PASS: rounding, formulas, hidden rows, split, zero quantity, export, leading zeros, filter, missing report, blank filter header, preflight, automatic starts, block copy, batch rounding, cancellation, calculation restore, form initialization, filtered split, filtered split all rows, filtered last row, filter criteria restore, permanent version label, zinc template, zinc xls roundtrip, zinc split export, zinc no overwrite, zinc row limit, readonly search copy, shared private copy, copy source unchanged, copy protection preserved."
    Exit Function
Bad:
    e = Err.Number: message = Err.Description
    EndWork
    On Error Resume Next
    If Not result Is Nothing Then result.Close SaveChanges:=False
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    Unload frmZekon
    Application.DisplayAlerts = oldAlerts
    On Error GoTo 0
    Err.Raise e, "ZekonSelfTest", message
End Function

Private Sub TestWorkingCopies()
    Dim folder As String, suffix As Long, oldCopy As String
    Dim number As Long, message As String
    oldCopy = LastCopyPath
    On Error GoTo Bad
    folder = Environ$("TEMP") & "\ZekonCopyTest_" & Format$(Now, "yyyymmdd_hhnnss")
    Do While Len(Dir$(folder & "_" & CStr(suffix), vbDirectory)) > 0
        suffix = suffix + 1
    Loop
    folder = folder & "_" & CStr(suffix)
    MkDir folder
    TestOneWorkingCopy folder & "\readonly.xlsx", False
    TestOneWorkingCopy folder & "\shared.xlsx", True
    RmDir folder
    LastCopyPath = oldCopy
    Exit Sub
Bad:
    number = Err.Number: message = Err.Description
    LastCopyPath = oldCopy
    Err.Raise number, "TestWorkingCopies", message & " | Folder testu: " & folder
End Sub

Private Sub TestOneWorkingCopy(ByVal sourcePath As String, ByVal legacyShared As Boolean)
    Dim source As Workbook, working As Workbook, ws As Worksheet, bt As Worksheet
    Dim p(1 To 10) As Variant, first As Long, n As Long, caught As Long
    Dim before As String, copyPath As String, oldSecurity As Long
    Dim oldEvents As Boolean, oldAlerts As Boolean, number As Long, message As String
    oldSecurity = Application.AutomationSecurity
    oldEvents = Application.EnableEvents: oldAlerts = Application.DisplayAlerts
    On Error GoTo Bad
    Set source = Workbooks.Add(xlWBATWorksheet)
    Set ws = source.Worksheets(1): ws.Name = "List"
    ws.Range("A1").Value2 = "Position"
    ws.Range("A2").Value2 = "P1": ws.Range("A3").Value2 = "PX"
    ws.Range("F1").Value2 = "Weight": ws.Range("F2").Value2 = 2.5
    ws.Range("G2").Formula = "=F2*2": ws.Rows(2).RowHeight = 27
    Set bt = source.Worksheets.Add(After:=ws): bt.Name = "Target"
    bt.Range("C1").Value2 = "Position": bt.Range("C2").Value2 = "P1": bt.Range("C3").Value2 = "P2"
    If legacyShared Then
        source.SaveAs Filename:=sourcePath, FileFormat:=xlOpenXMLWorkbook, AccessMode:=xlShared
        AssertTrue source.MultiUserEditing, "Utworzenie wspoldzielonego skoroszytu testowego"
    Else
        source.SaveAs Filename:=sourcePath, FileFormat:=xlOpenXMLWorkbook, ReadOnlyRecommended:=True
    End If
    source.Close SaveChanges:=False: Set source = Nothing
    Set source = Workbooks.Open(Filename:=sourcePath, UpdateLinks:=0, ReadOnly:=Not legacyShared, IgnoreReadOnlyRecommended:=True, Notify:=False)
    Set ws = source.Worksheets("List"): Set bt = source.Worksheets("Target")
    AssertTrue source.ReadOnly = (Not legacyShared), "Stan odczytu oryginalu testowego"
    AssertTrue source.MultiUserEditing = legacyShared, "Stan wspoldzielenia oryginalu testowego"
    before = ReadTestFile(sourcePath)
    p(1) = "A": p(2) = "Target": p(3) = "C": p(4) = "AUTO"
    first = ResolveFirstRow(ws, "A", "AUTO")
    If Not legacyShared Then
        On Error Resume Next
        ValidateSearch ws, first, p, True
        caught = Err.Number: Err.Clear
        On Error GoTo Bad
        AssertTrue caught <> 0, "Oryginal tylko do odczytu odrzucony bez pracy na kopii"
    End If
    ' Same two validation phases as the panel: source preflight, then copy access.
    ValidateSearch ws, first, p, True, True
    Set working = CopyWorkbook(source): copyPath = working.FullName
    AssertTrue StrComp(copyPath, sourcePath, vbTextCompare) <> 0, "Kopia w innym pliku"
    AssertTrue Not working.ReadOnly And Not working.MultiUserEditing, "Kopia zapisywalna i niewspoldzielona"
    AssertTrue Application.AutomationSecurity = oldSecurity, "Przywrocony tryb zabezpieczen otwierania kopii"
    AssertTrue Application.EnableEvents = oldEvents And Application.DisplayAlerts = oldAlerts, "Przywrocone zdarzenia i alerty kopii"
    Set ws = working.Worksheets("List")
    ValidateSearch ws, first, p, True
    n = SearchPositions(ws, first, p, True)
    AssertTrue n = 1 And working.Worksheets.Count = 3, "Wyszukiwanie i raport w kopii"
    AssertTrue working.Worksheets("Target").Rows(3).Hidden, "Filtr zastosowany w kopii"
    AssertTrue ws.Range("A3").FormatConditions.Count > 0, "Oznaczenie brakow w kopii"
    n = RoundColumn(ws, 2, "F", False, False)
    AssertTrue n = 1 And ws.Range("F2").Value2 = 3, "Pozostale operacje modyfikuja kopie"
    AssertTrue ws.Range("G2").Formula = "=F2*2" And ws.Rows(2).RowHeight = 27, "Formuly i format w kopii"
    working.Save
    AssertTrue source.ReadOnly = (Not legacyShared) And source.MultiUserEditing = legacyShared, "Tryb oryginalu bez zmian"
    AssertTrue source.Worksheets.Count = 2 And Not source.Worksheets("Target").Rows(3).Hidden, "Oryginal bez raportu i filtra"
    AssertTrue source.Worksheets("List").Range("F2").Value2 = 2.5, "Dane oryginalu bez zmian"
    AssertTrue source.Worksheets("List").Range("A3").FormatConditions.Count = 0, "Format oryginalu bez zmian"
    AssertTrue ReadTestFile(sourcePath) = before, "Plik oryginalu nie zostal zapisany"
    ws.Protect Password:="test"
    caught = 0
    On Error Resume Next
    ValidateSearch ws, first, p, True, True
    caught = Err.Number: Err.Clear
    On Error GoTo Bad
    AssertTrue caught <> 0 And ws.ProtectContents, "Plan kopii nie omija ochrony arkusza"
    ws.Unprotect Password:="test"
    working.Protect Password:="test", Structure:=True
    caught = 0
    On Error Resume Next
    ValidateSearch ws, first, p, True, True
    caught = Err.Number: Err.Clear
    On Error GoTo Bad
    AssertTrue caught <> 0 And working.ProtectStructure, "Plan kopii nie omija ochrony struktury"
    working.Unprotect Password:="test"
    working.Close SaveChanges:=False: Set working = Nothing
    source.Close SaveChanges:=False: Set source = Nothing
    Kill copyPath: Kill sourcePath
    Exit Sub
Bad:
    number = Err.Number: message = Err.Description
    On Error Resume Next
    If Not working Is Nothing Then working.Close SaveChanges:=False
    If Not source Is Nothing Then source.Close SaveChanges:=False
    Application.AutomationSecurity = oldSecurity
    Application.EnableEvents = oldEvents: Application.DisplayAlerts = oldAlerts
    On Error GoTo 0
    Err.Raise number, "TestOneWorkingCopy", message
End Sub

Private Function ReadTestFile(ByVal path As String) As String
    Dim f As Integer, number As Long, message As String
    Dim data As String
    On Error GoTo Bad
    f = FreeFile
    Open path For Binary Access Read Shared As #f
    data = Space$(LOF(f))
    Get #f, , data
    Close #f
    ReadTestFile = data
    Exit Function
Bad:
    number = Err.Number: message = Err.Description
    On Error Resume Next
    Close #f
    On Error GoTo 0
    Err.Raise number, "ReadTestFile", message
End Function

Private Sub TestZincExports()
    Dim source As Workbook, ws As Worksheet, reference As Workbook, result As Workbook
    Dim folder As String, target As String, p(1 To 10) As Variant, n As Long, index As Long
    Dim number As Long, message As String, caught As Long, sizeBefore As Long
    On Error GoTo Bad
    folder = Environ$("TEMP") & "\ZekonExportTest_" & Format$(Now, "yyyymmdd_hhnnss")
    Do While Len(Dir$(folder & "_" & CStr(index), vbDirectory)) > 0
        index = index + 1
    Loop
    folder = folder & "_" & CStr(index)
    MkDir folder: MkDir folder & "\direct": MkDir folder & "\split"
    WriteZincTemplate folder & "\reference.xls"
    Set reference = Workbooks.Open(Filename:=folder & "\reference.xls", UpdateLinks:=0, ReadOnly:=True)
    AssertTrue reference.FileFormat = xlExcel8, "Wzorzec BIFF8 / Excel 97-2003"
    AssertTrue Application.CountA(reference.Worksheets(1).UsedRange) = 8, "Wzorzec bez danych klienta"
    Set source = Workbooks.Add(xlWBATWorksheet): Set ws = source.Worksheets(1)
    ws.Range("A1").Value2 = "Kontrakt": ws.Range("B1").Value2 = "Pozycja"
    ws.Range("C1").Value2 = "Profil": ws.Range("D1").Value2 = "Dlugosc"
    ws.Range("E1").Value2 = "Waga": ws.Range("F1").Value2 = "Sztuka"
    ws.Range("G1").Value2 = "Ilosc": ws.Range("H1").Value2 = "Grupa"
    ws.Range("A2:B3").NumberFormat = "@"
    ws.Range("A2:A3").Value2 = "00123": ws.Range("B2:B3").Value2 = "0007"
    ws.Range("C2:C3").Value2 = "HEA100": ws.Range("D2:D3").Value2 = 1200.5
    ws.Range("E2:E3").Value2 = 2.5: ws.Range("F2:F3").Value2 = 7
    ws.Range("G2").Value2 = 2: ws.Range("G3").Value2 = 4
    ws.Range("H2").Value2 = "YES": ws.Range("H3").Value2 = "NO"
    ws.Range("A1:H3").AutoFilter Field:=8, Criteria1:="YES"
    p(1) = "=2+2": p(2) = "A": p(3) = "B": p(4) = "C"
    p(5) = "D": p(6) = "E": p(7) = "F": p(8) = "99": p(9) = "G"
    target = folder & "\direct\wzor.xls"
    n = ExportRows(ws, 2, p, True, False, target)
    Set result = ActiveWorkbook
    AssertTrue n = 1, "Ocynk: tylko widoczne pozycje"
    result.Close SaveChanges:=False: Set result = Nothing
    Set result = Workbooks.Open(Filename:=target, UpdateLinks:=0, ReadOnly:=True)
    AssertZincLayout result, reference, 1
    AssertTrue result.Worksheets(1).Range("G2").Value2 = 7, "Ocynk: zachowany numer sztuki"
    result.Close SaveChanges:=False: Set result = Nothing
    sizeBefore = FileLen(target)
    On Error Resume Next
    n = ExportRows(ws, 2, p, True, False, target)
    caught = Err.Number: Err.Clear
    On Error GoTo Bad
    AssertTrue caught <> 0 And FileLen(target) = sizeBefore, "Ocynk: brak nadpisania istniejacego pliku"
    ws.Range("G2").Value2 = 65536: caught = 0
    On Error Resume Next
    ValidateZincExportSize ws, 2, p, True, True
    caught = Err.Number: Err.Clear
    On Error GoTo Bad
    AssertTrue caught <> 0 And ws.Range("G2").Value2 = 65536, "Ocynk: limit XLS przed rozbiciem"
    ws.Range("G2").Value2 = 2
    ValidateZincExportSize ws, 2, p, True, True
    n = SplitRows(ws, 2, "G", "F", True)
    ws.Calculate
    target = folder & "\split\wzor.xls"
    n = ExportRows(ws, 2, p, True, False, target)
    Set result = ActiveWorkbook
    AssertTrue n = 2, "Sztuki + ocynk: liczba po rozbiciu pod filtrem"
    result.Close SaveChanges:=False: Set result = Nothing
    Set result = Workbooks.Open(Filename:=target, UpdateLinks:=0, ReadOnly:=True)
    AssertZincLayout result, reference, 2
    AssertTrue result.Worksheets(1).Range("G2").Value2 = 1 And result.Worksheets(1).Range("G3").Value2 = 2, "Sztuki + ocynk: numeracja w XLS"
    result.Close SaveChanges:=False: Set result = Nothing
    reference.Close SaveChanges:=False: Set reference = Nothing
    source.Close SaveChanges:=False: Set source = Nothing
    Kill folder & "\direct\wzor.xls": Kill folder & "\split\wzor.xls": Kill folder & "\reference.xls"
    RmDir folder & "\direct": RmDir folder & "\split": RmDir folder
    Exit Sub
Bad:
    number = Err.Number: message = Err.Description
    On Error Resume Next
    If Not result Is Nothing Then result.Close SaveChanges:=False
    If Not reference Is Nothing Then reference.Close SaveChanges:=False
    If Not source Is Nothing Then source.Close SaveChanges:=False
    On Error GoTo 0
    Err.Raise number, "TestZincExports", message & " | Pliki diagnostyczne: " & folder
End Sub

Private Sub AssertZincLayout(ByVal result As Workbook, ByVal reference As Workbook, ByVal count As Long)
    Dim ws As Worksheet, expected As Worksheet, i As Long, c As Long
    Dim headings As Variant
    headings = Array("Auf. Name", "Auftr.", "Pos.", "Profil", "Lange", "Gewicht", "Lfn nr.", "Zekon Unterlieferanten")
    AssertTrue result.Name = "wzor.xls" And result.FileFormat = xlExcel8, "Nazwa wzor.xls i rzeczywisty format XLS"
    AssertTrue result.Worksheets.Count = 3, "Trzy arkusze wzorca"
    AssertTrue Not result.HasVBProject, "Eksport bez makr"
    For i = 1 To 3
        Set ws = result.Worksheets(i): Set expected = reference.Worksheets(i)
        AssertTrue ws.Name = "Arkusz" & CStr(i), "Nazwy arkuszy wzorca"
        AssertTrue ws.StandardHeight = expected.StandardHeight, "Domyslna wysokosc wierszy"
        For c = 1 To 10
            AssertTrue Abs(ws.Columns(c).ColumnWidth - expected.Columns(c).ColumnWidth) < 0.01, "Szerokosc kolumny " & c
        Next c
        AssertTrue Abs(ws.PageSetup.LeftMargin - expected.PageSetup.LeftMargin) < 0.01, "Lewy margines wzorca"
        AssertTrue Abs(ws.PageSetup.RightMargin - expected.PageSetup.RightMargin) < 0.01, "Prawy margines wzorca"
        If i > 1 Then AssertTrue Application.CountA(ws.UsedRange) = 0, "Puste arkusze pomocnicze"
    Next i
    Set ws = result.Worksheets(1): Set expected = reference.Worksheets(1)
    For c = 1 To 8
        AssertTrue ws.Cells(1, c).Value2 = headings(c - 1), "Naglowek " & c
        AssertTrue ws.Cells(1, c).NumberFormat = "General" And ws.Cells(2, c).NumberFormat = "General", "Format Ogolny " & c
        AssertTrue ws.Cells(1, c).Font.Name = "Calibri" And ws.Cells(2, c).Font.Name = "Calibri", "Czcionka Calibri " & c
        AssertTrue ws.Cells(1, c).Font.Size = 11 And ws.Cells(2, c).Font.Size = 11, "Rozmiar czcionki " & c
        AssertTrue Not ws.Cells(1, c).Font.Bold And Not ws.Cells(2, c).Font.Bold, "Brak pogrubienia " & c
        AssertTrue ws.Cells(1, c).Interior.Pattern = expected.Cells(1, c).Interior.Pattern, "Tlo naglowka " & c
        AssertTrue ws.Cells(1, c).HorizontalAlignment = expected.Cells(1, c).HorizontalAlignment, "Wyrownanie naglowka " & c
    Next c
    AssertTrue ws.Rows(1).RowHeight = 14.5 And ws.Rows(2).RowHeight = 14.5, "Wysokosc 14.5 pkt"
    AssertTrue ws.Range("A2").Value2 = "=2+2" And Not ws.Range("A2").HasFormula, "Nazwa kontraktu pozostaje tekstem"
    AssertTrue ws.Range("B2").Value2 = "00123" And VarType(ws.Range("B2").Value2) = vbString, "Tekstowy kontrakt i zera wiodace po zapisie XLS"
    AssertTrue ws.Range("C2").Value2 = "0007" And VarType(ws.Range("C2").Value2) = vbString, "Tekstowa pozycja i zera wiodace po zapisie XLS"
    AssertTrue ws.Range("E2").Value2 = 1200.5 And ws.Range("F2").Value2 = 3, "Dlugosc i zaokraglona waga"
    AssertTrue Application.CountA(ws.Columns(8)) = 1, "Kolumna podwykonawcy bez dopisywania danych"
    AssertTrue LastRow(ws, 3) = count + 1, "Dane od drugiego wiersza bez pozostalosci wzorca"
End Sub

Private Sub TestAutoStart(ByVal wb As Workbook)
    Dim ws As Worksheet, lo As ListObject, n As Long, e As Long
    Dim alerts As Boolean
    Set ws = wb.Worksheets.Add
    ws.Range("F6").Value2 = "WEIGHT/PCS": ws.Range("F8").Value2 = 2.5
    AssertTrue ResolveFirstRow(ws, "F", "AUTO", True) = 8, "Auto: naglowek i pusty wiersz"
    n = RoundColumn(ws, ResolveFirstRow(ws, "F", "AUTO", True), "F", False, False)
    AssertTrue ws.Range("F8").Value2 = 3, "Zaokraglanie od wykrytego poczatku"
    ws.Cells.Clear
    ws.Range("A1").Value2 = "P001": ws.Range("A2").Value2 = "P002"
    AssertTrue ResolveFirstRow(ws, "A", "AUTO") = 1, "Auto: lista bez naglowka"
    ws.Cells.Clear
    ws.Range("B4").Value2 = "Pozycja": ws.Range("B5").Value2 = "P1"
    AssertTrue ResolveFirstRow(ws, "B", "AUTO") = 5, "Auto: naglowek w innym wierszu"
    ws.Cells.Clear
    ws.Range("B12").Value2 = "P1": ws.Range("B13").Value2 = "P2"
    ws.Range("A11:D13").AutoFilter
    ws.Rows(12).Hidden = True
    AssertTrue ResolveFirstRow(ws, "B", "AUTO") = 12, "Auto: filtr, pusty naglowek, ukryty pierwszy wiersz"
    ws.AutoFilterMode = False: ws.Rows(12).Hidden = False: ws.Cells.Clear
    ws.Range("B3").Value2 = "Pozycja": ws.Range("B4").Value2 = "P1"
    Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range("B3:B4"), , xlYes)
    AssertTrue ResolveFirstRow(ws, "B", "AUTO") = 4, "Auto: tabela Excela"
    lo.Unlist: ws.Cells.Clear
    ws.Range("A1").Value2 = "Nieznany tytul": ws.Range("A2").Value2 = 2.5
    On Error Resume Next
    n = ResolveFirstRow(ws, "A", "AUTO", True)
    e = Err.Number: Err.Clear
    On Error GoTo 0
    AssertTrue e <> 0, "Auto: niejednoznaczny tytul wymaga wskazania"
    AssertTrue ResolveFirstRow(ws, "A", "2", True) = 2, "Reczne nadpisanie AUTO"
    ws.Cells.Clear
    ws.Range("F2:F2052").Value2 = 2.5
    n = RoundColumn(ws, 2, "F", False, False)
    AssertTrue n = 2051 And ws.Range("F2052").Value2 = 3, "Zapis zaokraglen ponad granica porcji 2048"
    alerts = Application.DisplayAlerts
    Application.DisplayAlerts = False
    ws.Delete
    Application.DisplayAlerts = alerts
End Sub

Private Sub TestWorkLifecycle()
    Dim previous As XlCalculation, e As Long, panel As Object
    previous = Application.Calculation
    BeginWork panel
    AssertTrue Application.Calculation = xlCalculationManual, "Wstrzymanie przeliczania"
    RequestWorkCancel
    On Error Resume Next
    ProgressTick "Test przerwania", 1, 1, True
    e = Err.Number: Err.Clear
    On Error GoTo 0
    EndWork
    AssertTrue e <> 0, "Zadanie przerwania jest obslugiwane"
    AssertTrue Application.Calculation = previous, "Przywrocenie przeliczania po przerwaniu"
End Sub

Private Sub TestSplitWithFilter(ByVal wb As Workbook)
    Dim ws As Worksheet, n As Long, scenario As Long, e As Long
    Set ws = wb.Worksheets.Add
    For scenario = 1 To 4
        ws.AutoFilterMode = False: ws.Cells.Clear: ws.Rows("1:20").Hidden = False
        ws.Range("A1").Value2 = "Group": ws.Range("B1").Value2 = "Qty"
        ws.Range("C1").Value2 = "Piece": ws.Range("D1").Value2 = "Formula"
        ws.Range("A2").Value2 = "YES": ws.Range("B2").Value2 = 2
        ws.Range("A3").Value2 = "NO": ws.Range("B3").Value2 = 4
        ws.Range("A4").Value2 = "YES": ws.Range("B4").Value2 = 3
        ws.Range("D2:D4").FormulaR1C1 = "=RC[-2]*10"
        ws.Range("A1:D4").AutoFilter Field:=1, Criteria1:=Array("YES"), Operator:=xlFilterValues
        If scenario <> 4 Then ws.Range("A1:D4").AutoFilter Field:=2, Criteria1:=">=2", Operator:=xlAnd, Criteria2:="<=4"
        If scenario = 3 Then
            ws.Range("B2").Value2 = 0
            On Error Resume Next
            n = SplitRows(ws, 2, "B", "C", True)
            e = Err.Number: Err.Clear
            On Error GoTo 0
            AssertTrue e <> 0, "Filtrowane rozbijanie: walidacja przed zmianami"
            AssertTrue ws.AutoFilter.Range.Rows.Count = 4 And ws.Rows(3).Hidden, "Filtr zachowany po bledzie walidacji"
        Else
            n = SplitRows(ws, 2, "B", "C", scenario <> 2)
            If scenario <> 2 Then
                AssertTrue n = 5, "Rozbijanie tylko widocznych pod filtrem"
                AssertTrue ws.Range("A4").Value2 = "NO" And ws.Range("B4").Value2 = 4, "Odfiltrowana pozycja bez zmian"
                AssertTrue ws.Range("C7").Value2 = 3 And ws.Range("D7").Formula = "=B7*10", "Ostatni filtrowany wiersz i formula"
                AssertTrue ws.AutoFilter.Range.Address = "$A$1:$D$7", "Powiekszony zakres filtra"
            Else
                AssertTrue n = 9 And ws.Range("C7").Value2 = 4, "Rozbijanie wszystkich przy aktywnym filtrze"
                AssertTrue ws.AutoFilter.Range.Address = "$A$1:$D$10", "Zakres filtra po rozbiciu wszystkich"
            End If
            If scenario = 4 Then
                AssertTrue Not ws.Rows(2).Hidden And ws.Rows(4).Hidden And Not ws.Rows(7).Hidden, "Widocznosc po rozbiciu nieciaglych pozycji"
                AssertTrue ws.AutoFilter.Filters(1).On, "Przywrocony filtr pozycji"
            Else
                AssertTrue ws.AutoFilter.Filters(1).On And ws.AutoFilter.Filters(2).On, "Przywrocone oba filtry"
                AssertTrue ws.AutoFilter.Filters(2).Criteria1 = ">=2" And ws.AutoFilter.Filters(2).Criteria2 = "<=4", "Przywrocone kryteria AND"
                AssertTrue ws.Rows(2).Hidden, "Filtr ilosci ponownie oceniony po zmianie na 1"
            End If
        End If
    Next scenario
End Sub

Private Sub TestPanelVersion()
    Dim panel As frmZekon, actual As String, expected As String
    Dim e As Long, message As String
    On Error GoTo Bad
    Set panel = New frmZekon
    Load panel
    expected = "Wersja: " & APP_VERSION
    actual = CStr(panel.Controls("version").Caption)
    AssertTrue StrComp(actual, expected, vbBinaryCompare) = 0, "Wersja panelu: oczekiwano [" & expected & "]; odczytano [" & Replace(Replace(actual, vbCr, "<CR>"), vbLf, "<LF>") & "]"
    AssertTrue panel.Caption = "Narzedzia ZEKON | " & APP_VERSION, "Wersja paska tytulu"
    Unload panel
    Set panel = Nothing
    Exit Sub
Bad:
    e = Err.Number: message = Err.Description
    On Error Resume Next
    If Not panel Is Nothing Then Unload panel
    Set panel = Nothing
    On Error GoTo 0
    Err.Raise e, "TestPanelVersion", message
End Sub
