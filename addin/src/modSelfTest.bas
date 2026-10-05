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
    TestWorkLifecycle
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
    Load frmZekon
    Unload frmZekon
    Application.DisplayAlerts = oldAlerts
    ZekonSelfTest = "PASS: rounding, formulas, hidden rows, split, zero quantity, export, leading zeros, filter, missing report, blank filter header, preflight, automatic starts, block copy, batch rounding, cancellation, calculation restore, form initialization."
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
