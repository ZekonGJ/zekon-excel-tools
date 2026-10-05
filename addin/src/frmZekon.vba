Option Explicit

Private mode As Long
Private busy As Boolean

Private Sub UserForm_Initialize()
    Dim wb As Workbook
    Me.Caption = "Narzedzia ZEKON | " & APP_VERSION
    Me.Controls("version").Caption = "Wersja dodatku" & vbLf & APP_VERSION
    imgLogo.Picture = LoadPicture(ThisWorkbook.Path & "\logo.bmp")
    For Each wb In Application.Workbooks
        If Not wb.IsAddin And wb.Name <> ThisWorkbook.Name Then cboBook.AddItem wb.Name
    Next wb
    If cboBook.ListCount > 0 Then
        cboBook.ListIndex = 0
        On Error Resume Next
        cboBook.Value = Application.ActiveWorkbook.Name
        cboSheet.Value = Application.ActiveSheet.Name
        On Error GoTo 0
    End If
    txtProfile.Text = "Domyslny"
    chkCopy.Value = True
    SelectMode 1
End Sub

Private Sub cboBook_Change()
    Dim ws As Worksheet, wb As Workbook
    cboSheet.Clear
    If cboBook.ListIndex < 0 Then Exit Sub
    Set wb = Application.Workbooks(CStr(cboBook.Value))
    For Each ws In wb.Worksheets
        cboSheet.AddItem ws.Name
    Next ws
    If cboSheet.ListCount > 0 Then cboSheet.ListIndex = 0
End Sub

Private Sub SelectMode(ByVal selected As Long)
    Dim labels As Variant, defaults As Variant, i As Long, title As String, detail As String
    If busy Then Exit Sub
    mode = selected
    chkFormulas.Visible = (mode = 1)
    chkCopy.Enabled = (mode <> 4 And mode <> 5)
    labels = Array(): defaults = Array()
    Select Case mode
        Case 1
            title = "Zaokraglanie"
            detail = "Zaokraglij wartosci do liczb calkowitych. Formuly domyslnie pozostaja bez zmian."
            labels = Array("Kolumna do zaokraglenia")
            defaults = Array("F")
        Case 2
            title = "Wyszukiwanie pozycji"
            detail = "Filtruj zestawienie wedlug listy. Braki otrzymaja czerwone oznaczenie i osobny raport."
            labels = Array("Kolumna listy pozycji", "Arkusz zestawienia", "Kolumna pozycji w zestawieniu", "Poczatek zestawienia (AUTO lub wiersz)")
            defaults = Array("A", "bt-list", "B", "AUTO")
        Case 3
            title = "Rozbijanie na sztuki"
            detail = "Powiel wiersze wedlug ilosci, ustaw ilosc = 1 i nadaj numery sztuk od 1."
            labels = Array("Kolumna ilosci", "Kolumna numeru sztuki")
            defaults = Array("E", "C")
        Case 4, 5, 6, 7
            If mode = 4 Or mode = 6 Then title = "Eksport: malowanie" Else title = "Eksport: ocynk"
            detail = "Ustaw mapowanie kolumn. Dlugosc wymaga wskazania. Eksport pomija puste pozycje."
            labels = Array("Nazwa kontraktu", "Kolumna nr kontraktu", "Kolumna pozycji", "Kolumna profilu", "Kolumna dlugosci", "Kolumna wagi", "Kolumna numeru sztuki", "Pierwszy wiersz eksportu")
            If mode = 4 Or mode = 6 Then
                defaults = Array("", "B", "D", "M", "", "F", "C", "3")
            Else
                defaults = Array("", "B", "D", "M", "", "F", "C", "2")
            End If
            If mode >= 6 Then
                title = "Sekwencja: sztuki + " & IIf(mode = 6, "malowanie", "ocynk")
                detail = "Najpierw rozbij ilosci; potem eksportuj z tego samego arkusza. Numer sztuki: kolumna z pola 7."
                labels = Array("Nazwa kontraktu", "Kolumna nr kontraktu", "Kolumna pozycji", "Kolumna profilu", "Kolumna dlugosci", "Kolumna wagi", "Kolumna numeru sztuki", "Pierwszy wiersz eksportu", "Kolumna ilosci do rozbicia")
                If mode = 6 Then defaults = Array("", "B", "D", "M", "", "F", "C", "3", "E")
                If mode = 7 Then defaults = Array("", "B", "D", "M", "", "F", "C", "2", "E")
            End If
    End Select
    lblTitle.Caption = title
    lblDetail.Caption = detail
    For i = 1 To 10
        Me.Controls("lblP" & i).Visible = (i <= UBound(labels) + 1)
        Me.Controls("txtP" & i).Visible = (i <= UBound(labels) + 1)
        Me.Controls("txtP" & i).Text = ""
        If i <= UBound(labels) + 1 Then
            Me.Controls("lblP" & i).Caption = labels(i - 1)
            Me.Controls("txtP" & i).Text = defaults(i - 1)
        End If
    Next i
    txtFirst.Text = "AUTO"
    chkVisible.Value = True
    chkFormulas.Value = False
    LoadProfile
    For i = 1 To 7
        Me.Controls("nav" & i).BackColor = RGB(30, 42, 56)
        If i = mode Then Me.Controls("nav" & i).BackColor = RGB(189, 32, 45)
    Next i
    lblStatus.Caption = "Wybierz dane, sprawdz ustawienia i uruchom operacje."
End Sub

Private Function SectionName() As String
    Dim s As String
    s = Trim$(txtProfile.Text)
    If Len(s) = 0 Then Fail "Wpisz nazwe profilu ustawien."
    If Len(s) > 60 Or InStr(s, "\") > 0 Then Fail "Niepoprawna nazwa profilu."
    SectionName = "Profil_" & s & "_" & mode
End Function

Private Sub LoadProfile()
    Dim section As String, i As Long
    section = SectionName()
    For i = 1 To 10
        Me.Controls("txtP" & i).Text = GetSetting(APP_ID, section, "p" & i, Me.Controls("txtP" & i).Text)
    Next i
    txtFirst.Text = GetSetting(APP_ID, section, "first_auto", "AUTO")
    If mode = 2 Then txtP4.Text = GetSetting(APP_ID, section, "target_first_auto", "AUTO")
    If mode >= 4 Then
        txtP8.Text = IIf(mode = 4 Or mode = 6, "3", "2")
        txtP8.Visible = False: lblP8.Visible = False
    End If
    chkVisible.Value = (GetSetting(APP_ID, section, "visible", "1") = "1")
End Sub

Private Sub cmdLoad_Click()
    On Error GoTo Bad
    SelectMode mode
    lblStatus.Caption = "Wczytano ustawienia: " & txtProfile.Text
    Exit Sub
Bad:
    MsgBox Err.Description, vbExclamation, "ZEKON"
End Sub

Private Sub cmdSave_Click()
    Dim section As String, i As Long
    On Error GoTo Bad
    section = SectionName()
    For i = 1 To 10
        SaveSetting APP_ID, section, "p" & i, Me.Controls("txtP" & i).Text
    Next i
    SaveSetting APP_ID, section, "first_auto", txtFirst.Text
    If mode = 2 Then SaveSetting APP_ID, section, "target_first_auto", txtP4.Text
    SaveSetting APP_ID, section, "visible", IIf(chkVisible.Value, "1", "0")
    lblStatus.Caption = "Zapisano profil dla tej funkcji: " & txtProfile.Text
    Exit Sub
Bad:
    MsgBox Err.Description, vbExclamation, "ZEKON"
End Sub

Private Sub cmdRun_Click()
    Dim wb As Workbook, ws As Worksheet, working As Workbook, sheetName As String
    Dim first As Long, i As Long, n As Long, p(1 To 10) As Variant, dummy As Long
    Dim oldEvents As Boolean, oldScreen As Boolean, altered As Boolean, mutating As Boolean
    Dim errorText As String, painting As Boolean, resultText As String
    On Error GoTo Bad
    If busy Then Exit Sub
    LastCopyPath = ""
    Set wb = Application.Workbooks(CStr(cboBook.Value))
    sheetName = CStr(cboSheet.Value): Set ws = wb.Worksheets(sheetName)
    For i = 1 To 10
        p(i) = Trim$(Me.Controls("txtP" & i).Text)
    Next i
    busy = True
    SetWorkingControls True
    oldEvents = Application.EnableEvents: oldScreen = Application.ScreenUpdating
    altered = True
    Application.EnableEvents = False: Application.ScreenUpdating = False
    BeginWork Me
    Select Case mode
        Case 1, 3: first = ResolveFirstRow(ws, CStr(p(1)), txtFirst.Text, True)
        Case 2: first = ResolveFirstRow(ws, CStr(p(1)), txtFirst.Text, False)
        Case Else: first = ResolveFirstRow(ws, CStr(p(3)), txtFirst.Text, False)
    End Select
    If mode >= 4 Then
        p(8) = IIf(mode = 4 Or mode = 6, "3", "2")
        For i = 2 To 7
            dummy = ColumnNumber(ws, CStr(p(i)))
        Next i
        If ColumnNumber(ws, CStr(p(5))) = ColumnNumber(ws, CStr(p(7))) Then Fail "Dlugosc i numer sztuki wskazuja te sama kolumne."
        dummy = WholeNumber(CStr(p(8)), "Pierwszy wiersz eksportu", ws.Rows.Count)
        If dummy < 2 Then Fail "Pierwszy wiersz eksportu musi byc co najmniej 2."
        If mode >= 6 Then
            dummy = ColumnNumber(ws, CStr(p(9)))
            If dummy = ColumnNumber(ws, CStr(p(7))) Then Fail "Ilosc i numer sztuki musza byc w roznych kolumnach."
        End If
    End If
    If mode = 2 Then ValidateSearch ws, first, p, chkVisible.Value
    mutating = (mode <= 3 Or mode >= 6)
    If mutating And Not chkCopy.Value Then
        If MsgBox("Zmienisz dane w oryginale: " & wb.Name & " / " & ws.Name & vbCrLf & "Operacji VBA nie mozna cofnac przez Ctrl+Z. Kontynuowac?", vbYesNo + vbExclamation, "ZEKON") <> vbYes Then
            EndWork
            Application.EnableEvents = oldEvents: Application.ScreenUpdating = oldScreen
            altered = False: busy = False: SetWorkingControls False
            Exit Sub
        End If
    End If
    If mutating And chkCopy.Value Then
        Set working = CopyWorkbook(wb)
        Set ws = working.Worksheets(sheetName)
    End If
    Select Case mode
        Case 1: n = RoundColumn(ws, first, CStr(p(1)), chkVisible.Value, chkFormulas.Value)
        Case 2: n = SearchPositions(ws, first, p, chkVisible.Value)
        Case 3: n = SplitRows(ws, first, CStr(p(1)), CStr(p(2)), chkVisible.Value)
        Case 4, 5: n = ExportRows(ws, first, p, chkVisible.Value, mode = 4)
        Case 6, 7
            n = SplitRows(ws, first, CStr(p(9)), CStr(p(7)), chkVisible.Value)
            ProgressTick "Przeliczanie danych do eksportu", 0, 0, True
            ws.Calculate
            n = ExportRows(ws, first, p, chkVisible.Value, mode = 6)
    End Select
    If mode = 1 Or mode = 3 Then
        ws.Parent.Activate
        ws.Activate
    End If
    EndWork
    Application.EnableEvents = oldEvents: Application.ScreenUpdating = oldScreen
    altered = False: busy = False: SetWorkingControls False
    resultText = "Gotowe. Poczatek danych zrodlowych: wiersz " & first & ". "
    If mode = 2 Then resultText = resultText & "Brakujacych wystapien: " & n Else resultText = resultText & "Przetworzonych wierszy/komorek: " & n
    resultText = resultText & vbCrLf & "Sprawdz wynik i zapisz go w wybranym miejscu."
    If Len(LastCopyPath) > 0 Then resultText = resultText & vbCrLf & "Kopia robocza: " & LastCopyPath
    lblStatus.Caption = "Gotowe. Wynik jest otwarty w Excelu."
    MsgBox resultText, vbInformation, "ZEKON"
    Unload Me
    Exit Sub
Bad:
    errorText = Err.Description
    If altered Then
        EndWork
        Application.EnableEvents = oldEvents: Application.ScreenUpdating = oldScreen
    End If
    busy = False: SetWorkingControls False
    lblStatus.Caption = "Operacja zatrzymana."
    If Len(LastCopyPath) > 0 Then errorText = errorText & vbCrLf & "Kopia robocza: " & LastCopyPath
    MsgBox errorText & vbCrLf & "W razie bledu w trakcie operacji sprawdz otwarty wynik; zmiany nie sa automatycznie wycofywane.", vbExclamation, "ZEKON"
End Sub

Private Sub cmdClose_Click()
    If busy Then
        RequestWorkCancel
        cmdClose.Enabled = False
        lblStatus.Caption = "Przerywanie po zakonczeniu biezacej czynnosci..."
    Else
        Unload Me
    End If
End Sub

Private Sub nav1_Click()
    SelectMode 1
End Sub
Private Sub nav2_Click()
    SelectMode 2
End Sub
Private Sub nav3_Click()
    SelectMode 3
End Sub
Private Sub nav4_Click()
    SelectMode 4
End Sub
Private Sub nav5_Click()
    SelectMode 5
End Sub
Private Sub nav6_Click()
    SelectMode 6
End Sub
Private Sub nav7_Click()
    SelectMode 7
End Sub

Private Sub SetWorkingControls(ByVal working As Boolean)
    Dim ctl As Object
    For Each ctl In Me.Controls
        Select Case TypeName(ctl)
            Case "CommandButton", "TextBox", "ComboBox", "CheckBox"
                ctl.Enabled = Not working
        End Select
    Next ctl
    cmdClose.Enabled = True
    cmdClose.Caption = IIf(working, "Przerwij", "Zamknij")
    If Not working Then chkCopy.Enabled = (mode <> 4 And mode <> 5)
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If busy Then
        Cancel = True
        RequestWorkCancel
        lblStatus.Caption = "Przerywanie po zakonczeniu biezacej czynnosci..."
    End If
End Sub
