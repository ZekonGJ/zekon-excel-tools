import json
from pathlib import Path
r=Path(__file__).resolve().parents[1]
spec=json.loads((r/'interface.json').read_text())
def literal(v):
    if isinstance(v,bool):return 'True' if v else 'False'
    if isinstance(v,(int,float)):return str(v)
    return ' & vbLf & '.join('"'+s.replace('"','""')+'"' for s in v.split('\n'))
lines=['Attribute VB_Name = "ZekonBuilder"','Option Explicit','', '''Public Sub ZbudujDodatekZekon()
    Dim selected As Variant, root As String, dest As String, output As String
    Dim wb As Workbook, parts As Object, part As Object, form As Object, designer As Object, ctl As Object
    Dim modules As Variant, item As Variant, test As Variant, f As Integer
    Dim oldEvents As Boolean, oldScreen As Boolean, created As Boolean, errorText As String
    Dim stage As String, sourceCode As String, openIndex As Long
    On Error GoTo Bad
    oldEvents = Application.EnableEvents: oldScreen = Application.ScreenUpdating
    If ThisWorkbook.IsAddin Then Err.Raise vbObjectError + 906, , "Zaimportuj instalator do pustego skoroszytu, nie do dodatku."
    stage = "Wybor plikow"
    selected = Application.GetOpenFilename("ZEKON interface (interface.json),interface.json", , "Wskaz interface.json w rozpakowanym folderze ZekonTools")
    If VarType(selected) = vbBoolean Then Exit Sub
    root = Left$(CStr(selected), InStrRev(CStr(selected), "\\"))
    dest = Environ$("LOCALAPPDATA") & "\\ZekonTools"
    If Dir$(dest, vbDirectory) = "" Then MkDir dest
    dest = dest & "\\1.0.0-rc2a"
    If Dir$(dest, vbDirectory) = "" Then MkDir dest
    output = dest & "\\ZekonTools_1.0.0-rc2a.xlam"
    If Len(Dir$(output)) > 0 Then Err.Raise vbObjectError + 902, , "Dodatek juz istnieje. Nie nadpisano: " & output
    Application.EnableEvents = False: Application.ScreenUpdating = False
    ' Each release has a different filename; do not close the user's loaded add-in.
    stage = "Tworzenie skoroszytu i modulow"
    Set wb = Workbooks.Add(xlWBATWorksheet)
    Set parts = wb.VBProject.VBComponents
    modules = Array("modCore", "modRound", "modSplit", "modSearch", "modExport", "modSelfTest")
    For Each item In modules
        Set part = parts.Add(1)
        part.Name = CStr(item)
        sourceCode = ReadUtf8(root & "src\\" & item & ".bas")
        If CStr(item) = "modCore" Then sourceCode = Replace(sourceCode, "1.0.0-rc1a", "1.0.0-rc2a")
        If CStr(item) = "modCore" Then sourceCode = Replace(sourceCode, "1.0.0-rc1b", "1.0.0-rc2a")
        part.CodeModule.AddFromString sourceCode
    Next item
    stage = "Tworzenie panelu"
    Set form = parts.Add(3)
    form.Name = "frmZekon"
    form.Properties("Caption") = "Narzedzia ZEKON"
    form.Properties("Width") = 820
    form.Properties("Height") = 618
    form.Properties("BackColor") = 16118774
    form.Properties("StartUpPosition") = 1
    Set designer = form.Designer
    BuildControls designer
    form.CodeModule.AddFromString ReadUtf8(root & "src\\frmZekon.vba")
    parts(wb.CodeName).CodeModule.AddFromString ReadUtf8(root & "src\\ThisWorkbook.vba")
    stage = "Kopiowanie logo"
    FileCopy root & "assets\\logo.bmp", dest & "\\logo.bmp"
    wb.BuiltinDocumentProperties("Title") = "Narzedzia ZEKON 1.0.0-rc2a"
    wb.IsAddin = True
    stage = "Zapis XLAM"
    wb.SaveAs Filename:=output, FileFormat:=55
    created = True
    stage = "Kontrola zapisanego pliku"
    If Len(Dir$(output)) = 0 Then Err.Raise vbObjectError + 904, , "Brak pliku po SaveAs: " & output
    ' Zapis XLAM moze pozostawic wb jako skoroszyt roboczy, np. Arkusz2.
    ' Testy uruchamiamy z projektu otwartego ponownie z rzeczywistego pliku.
    stage = "Zamkniecie skoroszytu roboczego"
    Set ctl = Nothing
    Set designer = Nothing
    Set form = Nothing
    Set part = Nothing
    Set parts = Nothing
    wb.Close SaveChanges:=False
    Set wb = Nothing
    stage = "Otwarcie zapisanego dodatku"
    Set wb = Application.Workbooks.Open(Filename:=output, UpdateLinks:=0, ReadOnly:=False)
    If StrComp(wb.FullName, output, vbTextCompare) <> 0 Then Err.Raise vbObjectError + 905, , "Otwarty plik ma inna sciezke: " & wb.FullName
    stage = "Uruchomienie testow Excel"
    test = Application.Run("'" & Replace(wb.FullName, "'", "''") & "'!modSelfTest.ZekonSelfTest")
    If Left$(CStr(test), 5) <> "PASS:" Then Err.Raise vbObjectError + 903, , "Test dodatku nie zakonczyl sie PASS."
    stage = "Zapis wyniku testow"
    f = FreeFile
    Open dest & "\\test-result.txt" For Output As #f
    Print #f, Format$(Now, "yyyy-mm-dd hh:nn:ss")
    Print #f, CStr(test)
    Close #f: f = 0
    wb.Save
    wb.Close SaveChanges:=False
    Set wb = Nothing
    Application.EnableEvents = oldEvents: Application.ScreenUpdating = oldScreen
    MsgBox "DODATEK UTWORZONY. Testy funkcjonalne: PASS." & vbCrLf & output & vbCrLf & "Wlacz go przez Plik > Opcje > Dodatki > Dodatki programu Excel > Przejdz > Przegladaj." & vbCrLf & "Wylacz ponownie dostep do modelu obiektowego projektu VBA.", vbInformation, "ZEKON"
    Exit Sub
Bad:
    errorText = "Etap: " & stage & vbCrLf & "Blad " & Err.Number & ": " & Err.Description & vbCrLf & "Plik docelowy: " & output
    On Error Resume Next
    If f > 0 Then Close #f
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    If created Then Name output As dest & "\\FAILED_" & Format$(Now, "yyyymmdd_hhnnss") & ".xlam"
    Application.EnableEvents = oldEvents: Application.ScreenUpdating = oldScreen
    On Error GoTo 0
    MsgBox "Nie zakonczono budowania: " & errorText & vbCrLf & "Nie wlaczaj nieprzetestowanego pliku. Zachowaj ten komunikat diagnostyczny.", vbExclamation, "ZEKON 1.0.0-rc2a"
End Sub

Private Function ReadUtf8(ByVal path As String) As String
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2: stream.Charset = "utf-8"
    stream.Open
    stream.LoadFromFile path
    ReadUtf8 = stream.ReadText
    stream.Close
End Function

Private Sub BuildControls(ByVal designer As Object)
    Dim ctl As Object
''']
for c in spec:
    lines.append(f'    Set ctl = designer.Controls.Add({literal(c["type"])}, {literal(c["name"])}, True)')
    for k,v in c['props'].items():lines.append('    ctl.'+k+' = '+literal(v))
    if 'Image' not in c['type']:
        lines.append('    ctl.Font.Name = "Segoe UI"')
        lines.append('    ctl.Font.Size = '+str(c['fontSize']))
        if c['bold']:lines.append('    ctl.Font.Bold = True')
lines.append('End Sub')
(r/'ZbudujDodatek.bas').write_text('\n'.join(lines),encoding='ascii')
print('VBA builder generated')
