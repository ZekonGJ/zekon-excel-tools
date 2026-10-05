# Windows PowerShell 5.1, desktop Microsoft Excel. No policy or registry changes.
$ErrorActionPreference = 'Stop'
$root = $PSScriptRoot
$excel = $null
$book = $null
$success = $false
$candidateCreated = $false
$destination = Join-Path $env:LOCALAPPDATA 'ZekonTools\1.0.0-rc2b'
$output = Join-Path $destination 'ZekonTools_1.0.0-rc2b.xlam'
try {
    if (Get-Process EXCEL -ErrorAction SilentlyContinue) {
        throw 'Zamknij wszystkie okna Excela i uruchom instalator ponownie.'
    }
    if (Test-Path $output) {
        throw "Plik juz istnieje: $output. Nie nadpisano go. Przy ponownym budowaniu przenies poprzedni plik w inne miejsce."
    }
    New-Item -ItemType Directory -Path $destination -Force | Out-Null
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $book = $excel.Workbooks.Add()
    try { $components = $book.VBProject.VBComponents; $null = $components.Count }
    catch { throw 'Excel blokuje dostep do projektu VBA. Wykonaj krok 3 instrukcji i sproboj ponownie. Ustawien zabezpieczen nie zmieniono.' }
    foreach ($file in Get-ChildItem (Join-Path $root 'src\*.bas')) {
        $component = $components.Add(1)
        $component.Name = $file.BaseName
        $component.CodeModule.AddFromString([IO.File]::ReadAllText($file.FullName, [Text.Encoding]::UTF8))
    }
    $form = $components.Add(3)
    $form.Name = 'frmZekon'
    $form.Properties.Item('Caption').Value = 'Narzedzia ZEKON'
    $form.Properties.Item('Width').Value = 820
    $form.Properties.Item('Height').Value = 618
    $form.Properties.Item('BackColor').Value = 16118774
    $form.Properties.Item('StartUpPosition').Value = 1
    $designer = $form.Designer
    $spec = Get-Content (Join-Path $root 'interface.json') -Raw -Encoding UTF8 | ConvertFrom-Json
    foreach ($item in $spec) {
        $control = $designer.Controls.Add($item.type, $item.name, $true)
        foreach ($property in $item.props.PSObject.Properties) {
            $control.($property.Name) = $property.Value
        }
        if ($item.type -ne 'Forms.Image.1') {
            $control.Font.Name = 'Segoe UI'
            $control.Font.Size = $item.fontSize
            if ($item.bold) { $control.Font.Bold = $true }
        }
    }
    $form.CodeModule.AddFromString([IO.File]::ReadAllText((Join-Path $root 'src\frmZekon.vba'), [Text.Encoding]::UTF8))
    $components.Item($book.CodeName).CodeModule.AddFromString([IO.File]::ReadAllText((Join-Path $root 'src\ThisWorkbook.vba'), [Text.Encoding]::UTF8))
    Copy-Item (Join-Path $root 'assets\logo.bmp') (Join-Path $destination 'logo.bmp')
    $book.BuiltinDocumentProperties.Item('Title').Value = 'Narzedzia ZEKON 1.0.0-rc2b'
    $book.BuiltinDocumentProperties.Item('Comments').Value = '1.0.0-rc2b; modular VBA; Windows Excel'
    $book.IsAddin = $true
    $book.SaveAs($output, 55)
    $candidateCreated = $true
    if (-not (Test-Path -LiteralPath $output)) { throw "Brak pliku po SaveAs: $output" }
    $control = $null
    $designer = $null
    $form = $null
    $component = $null
    $components = $null
    $book.Close($false)
    $book = $null
    $book = $excel.Workbooks.Open($output, 0, $false)
    $macroPath = ([string]$book.FullName).Replace("'", "''")
    $test = $excel.Run("'$macroPath'!modSelfTest.ZekonSelfTest")
    if (-not ([string]$test).StartsWith('PASS:')) { throw "Test nie potwierdzil poprawnego wyniku: $test" }
    [IO.File]::WriteAllText((Join-Path $destination 'test-result.txt'), ([DateTime]::Now.ToString('s') + "`r`n" + $test), [Text.Encoding]::UTF8)
    $book.Save()
    $book.Close($false)
    $book = $null
    $success = $true
    Write-Host ''
    Write-Host 'DODATEK UTWORZONY. Testy funkcjonalne: PASS.' -ForegroundColor Green
    Write-Host $output
    Write-Host 'Teraz wlacz dodatek w Excelu zgodnie z krokiem 6 instrukcji.'
    Write-Host 'Wylacz ponownie dostep do modelu obiektowego projektu VBA.'
}
catch {
    Write-Host ('BLAD: ' + $_.Exception.Message) -ForegroundColor Red
    Write-Host 'Nie zakonczono budowania dodatku. Zapisz lub zrob zrzut tego komunikatu.'
}
finally {
    if ($book) { try { $book.Close($false) } catch {} }
    if ($excel) {
        try { $excel.Quit() } catch {}
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    if (-not $success -and $candidateCreated -and (Test-Path $output)) {
        # Keep a failed candidate separate; never register it as a working add-in.
        Move-Item $output (Join-Path $destination ('FAILED_' + [DateTime]::Now.ToString('yyyyMMddHHmmss') + '.xlam'))
    }
}
if (-not $success) { exit 1 }
