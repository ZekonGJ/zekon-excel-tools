Option Explicit

' The four workbooks are promoted by moving their containing directory together.
' A sibling journal survives a failure between backup and promotion.
Private Function PackageFS() As Object
    Set PackageFS = CreateObject("Scripting.FileSystemObject")
End Function

Private Sub SafePackagePath(ByVal folder As String)
    If Len(folder) = 0 Or Len(folder) > 220 Then Fail "Niepoprawna lub zbyt długa ścieżka pakietu."
    If Right$(folder, 1) = "\" Or InStr(folder, vbCr) > 0 Or InStr(folder, vbLf) > 0 Or InStr(folder, "..") > 0 Then Fail "Niepoprawna ścieżka pakietu."
End Sub

Public Function BeginProductionStage(ByVal currentFolder As String) As String
    Dim fs As Object, path As String, i As Long
    SafePackagePath currentFolder: Set fs = PackageFS()
    If fs.FileExists(currentFolder & ".pending") Then Fail "Pakiet ma niedokończony zapis: " & currentFolder & ".pending"
    If fs.FileExists(currentFolder & ".lock") Then Fail "Pakiet jest aktualizowany lub pozostała blokada po przerwaniu: " & currentFolder & ".lock"
    For i = 1 To 1000
        path = currentFolder & ".stage-" & Format$(Now, "yyyymmdd-hhnnss") & "-" & i
        If Not fs.FolderExists(path) Then Exit For
    Next i
    If i > 1000 Then Fail "Nie można utworzyć unikalnego folderu roboczego."
    With fs.CreateTextFile(currentFolder & ".lock", False, True)
        .WriteLine path: .Close
    End With
    On Error GoTo Bad
    fs.CreateFolder path: BeginProductionStage = path: Exit Function
Bad:
    On Error Resume Next: fs.DeleteFile currentFolder & ".lock", True: On Error GoTo 0
    Fail "Nie udało się utworzyć folderu roboczego pakietu."
End Function

Public Sub CommitProductionStage(ByVal currentFolder As String, ByVal stageFolder As String, ByVal expectedFiles As Variant, ByVal updateId As String)
    Dim fs As Object, item As Variant, backup As String, wb As Workbook, seen As Object
    Dim number As Long, message As String, movedOld As Boolean, promoted As Boolean, hadCurrent As Boolean, lockText As String
    SafePackagePath currentFolder: SafePackagePath stageFolder
    If updateId <> ProductionFilePart(updateId) Or Len(updateId) > 60 Then Fail "Niepoprawny identyfikator aktualizacji."
    Set fs = PackageFS(): Set seen = ProductionDictionary(): seen.CompareMode = vbTextCompare
    If Left$(stageFolder, Len(currentFolder) + 7) <> currentFolder & ".stage-" Then Fail "Folder roboczy nie należy do tego pakietu."
    If Not fs.FileExists(currentFolder & ".lock") Then Fail "Brak blokady własnego zapisu."
    With fs.OpenTextFile(currentFolder & ".lock", 1, False, -1)
        lockText = .ReadLine: .Close
    End With
    If lockText <> stageFolder Then Fail "Blokada należy do innej operacji."
    If fs.FileExists(currentFolder & ".pending") Then Fail "Niedokończona poprzednia aktualizacja."
    If UBound(expectedFiles) - LBound(expectedFiles) + 1 <> 4 Then Fail "Pakiet musi zawierać cztery pliki."
    For Each item In expectedFiles
        If ProductionFilePart(CStr(item)) <> CStr(item) Then Fail "Niepoprawna nazwa pliku pakietu."
        If seen.Exists(CStr(item)) Then Fail "Powtórzona nazwa pliku pakietu."
        seen.Add CStr(item), True
        If Not fs.FileExists(stageFolder & "\" & CStr(item)) Then Fail "Brak gotowego pliku: " & CStr(item)
        If fs.GetFile(stageFolder & "\" & CStr(item)).Size = 0 Then Fail "Pusty plik pakietu: " & CStr(item)
    Next item
    For Each wb In Application.Workbooks
        If StrComp(wb.Path, currentFolder, vbTextCompare) = 0 Or StrComp(wb.Path, stageFolder, vbTextCompare) = 0 Then Fail "Zamknij skoroszyty pakietu przed zatwierdzeniem: " & wb.Name
    Next wb
    hadCurrent = fs.FolderExists(currentFolder): backup = currentFolder & ".backup-" & updateId
    If fs.FolderExists(backup) Then Fail "Kopia zapasowa tego identyfikatora już istnieje."
    With fs.CreateTextFile(currentFolder & ".pending", False, True)
        .WriteLine "ZEKON-PACKAGE-1": .WriteLine currentFolder: .WriteLine stageFolder: .WriteLine backup
        .WriteLine IIf(hadCurrent, "UPDATE", "CREATE"): .Close
    End With
    On Error GoTo Bad
    If hadCurrent Then fs.MoveFolder currentFolder, backup: movedOld = True
    fs.MoveFolder stageFolder, currentFolder: promoted = True
    fs.DeleteFile currentFolder & ".pending", True: fs.DeleteFile currentFolder & ".lock", True
    Exit Sub
Bad:
    number = Err.Number: message = Err.Description
    On Error Resume Next
    If movedOld And Not promoted Then
        If Not fs.FolderExists(currentFolder) Then fs.MoveFolder backup, currentFolder
    End If
    On Error GoTo 0
    Err.Raise number, "CommitProductionStage", message & vbCrLf & "Zapis nie został potwierdzony. Dziennik: " & currentFolder & ".pending"
End Sub

Public Sub AbandonProductionStage(ByVal currentFolder As String, ByVal stageFolder As String)
    Dim fs As Object, lockText As String
    Set fs = PackageFS()
    If fs.FileExists(currentFolder & ".pending") Then Exit Sub
    If Not fs.FileExists(currentFolder & ".lock") Then Exit Sub
    With fs.OpenTextFile(currentFolder & ".lock", 1, False, -1): lockText = .ReadLine: .Close: End With
    If lockText = stageFolder Then fs.DeleteFile currentFolder & ".lock", True
    ' Staged files are retained for diagnosis; originals are never deleted.
End Sub

Public Function ProductionRecoveryJournal(ByVal currentFolder As String) As Object
    Dim fs As Object, stream As Object, state As Object, signature As String
    SafePackagePath currentFolder: Set fs = PackageFS(): Set state = ProductionDictionary()
    If Not fs.FileExists(currentFolder & ".pending") Then Fail "Brak dziennika przerwanego zapisu."
    Set stream = fs.OpenTextFile(currentFolder & ".pending", 1, False, -1)
    signature = stream.ReadLine
    state.Add "Target", stream.ReadLine: state.Add "Stage", stream.ReadLine
    state.Add "Backup", stream.ReadLine: state.Add "Mode", stream.ReadLine: stream.Close
    If signature <> "ZEKON-PACKAGE-1" Or state("Target") <> currentFolder Then Fail "Nieprawidłowy dziennik pakietu."
    If Left$(state("Stage"), Len(currentFolder) + 7) <> currentFolder & ".stage-" Or Left$(state("Backup"), Len(currentFolder) + 8) <> currentFolder & ".backup-" Then Fail "Nieprawidłowe ścieżki w dzienniku."
    SafePackagePath state("Stage"): SafePackagePath state("Backup")
    If state("Mode") <> "CREATE" And state("Mode") <> "UPDATE" Then Fail "Nieprawidłowy tryb dziennika."
    Set ProductionRecoveryJournal = state
End Function

Public Function ProductionFingerprint(ByVal folder As String, ByVal files As Variant) As String
    Dim fs As Object, file As Variant, entry As Object
    Set fs = PackageFS()
    For Each file In files
        Set entry = fs.GetFile(folder & "\" & CStr(file))
        ProductionFingerprint = ProductionFingerprint & CStr(file) & "|" & entry.Size & "|" & CDbl(entry.DateLastModified) & vbLf
    Next file
End Function
