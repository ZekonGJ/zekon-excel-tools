"""Embed the data-free XLS reference in VBA; the installer still carries two files."""
from pathlib import Path
import hashlib
import re

root = Path(__file__).resolve().parents[1]
data = (root / 'assets/zinc-template.xls').read_bytes()
chunks = [data[i:i + 120].hex().upper() for i in range(0, len(data), 120)]
lines = ['Option Explicit', '', "' Generated from assets/zinc-template.xls, with sample data removed.",
         f"' SHA256: {hashlib.sha256(data).hexdigest()}", '',
         'Public Sub WriteZincTemplate(ByVal path As String)',
         '    Dim f As Integer, number As Long, message As String',
         '    If Len(Dir$(path)) > 0 Then Fail "Plik wzorca juz istnieje: " & path',
         '    On Error GoTo Bad', '    f = FreeFile', '    Open path For Binary Access Write As #f']
for i in range(0, len(chunks), 80):
    lines.append(f'    ZincPart{i // 80} f')
lines += ['    Close #f', '    Exit Sub', 'Bad:',
          '    number = Err.Number: message = Err.Description', '    On Error Resume Next',
          '    Close #f', '    On Error GoTo 0', '    Err.Raise number, "WriteZincTemplate", message',
          'End Sub', '', 'Private Sub ZincHex(ByVal f As Integer, ByVal value As String)',
          '    Dim b() As Byte, i As Long', '    ReDim b(0 To Len(value) \\ 2 - 1)',
          '    For i = 0 To UBound(b)', '        b(i) = CByte("&H" & Mid$(value, i * 2 + 1, 2))',
          '    Next i', '    Put #f, , b', 'End Sub']
for i in range(0, len(chunks), 80):
    lines += ['', f'Private Sub ZincPart{i // 80}(ByVal f As Integer)']
    lines += [f'    ZincHex f, "{chunk}"' for chunk in chunks[i:i + 80]]
    lines += ['End Sub']
text = '\n'.join(lines) + '\n'
assert bytes.fromhex(''.join(re.findall(r'ZincHex f, "([0-9A-F]+)"', text))) == data
(root / 'src/modZincTemplate.bas').write_text(text, encoding='ascii')
print('PASS: embedded data-free XLS template matches asset, byte for byte')
