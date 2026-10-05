"""Create standalone RC2 builder B05 from the verified RC2 repair generator output."""
from pathlib import Path
import re, zipfile, hashlib
root=Path(__file__).resolve().parents[1]
# Regenerate the embedded VBA from maintained sources before applying builder-only fixes.
import subprocess,sys
subprocess.run([sys.executable,str(root/'addin/tools/generate_builder.py')],check=True)
subprocess.run([sys.executable,str(root/'addin/tools/generate_repair.py')],check=True)
s=(root/'ZbudujDodatek_NAPRAWA.bas').read_text()
s=s.replace('ZekonInstallerRC2','ZekonBuilderB05').replace('AktualizujZekonRC2','ZEKON_BUDUJ_RC2B_B05')
a=s.index('    stage = "Wybor plikow"')
b=s.index('    dest = Environ$',a)
s=s[:a]+'''    MsgBox "Generator B05 - dodatek 1.0.0-rc2b" & vbCrLf & "Kod i logo sa zawarte w tym pliku. Zostanie utworzony osobny folder proby.", vbInformation, "ZEKON RC2 / B05"
'''+s[b:]
needle='    output = dest & "\\ZekonTools_1.0.0-rc2b.xlam"'
s=s.replace(needle,'''    root = dest & "\\B05_" & Format$(Now, "yyyymmdd_hhnnss")
    dest = root
    openIndex = 0
    Do While Len(Dir$(dest, vbDirectory)) > 0
        openIndex = openIndex + 1
        dest = root & "_" & CStr(openIndex)
    Loop
    MkDir dest
'''+needle)
s=s.replace('    FileCopy root & "assets\\logo.bmp", dest & "\\logo.bmp"','    WriteEmbeddedLogo dest & "\\logo.bmp"')
s=s.replace(' & vbCrLf & "Wlacz go przez Plik > Opcje > Dodatki > Dodatki programu Excel > Przejdz > Przegladaj."',' & vbCrLf & "Przeslij plik XLAM oraz test-result.txt z otwartego folderu."')
s=s.replace('    Exit Sub\nBad:', '    Shell "explorer.exe """ & dest & """", vbNormalFocus\n    Exit Sub\nBad:',1)
s=s.replace('    Application.EnableEvents = True\n    Application.ScreenUpdating = True\n','',1)
s=s.replace('"ZEKON 1.0.0-rc2b"','"ZEKON RC2 / B05"')
logo=(root/'addin/assets/logo.bmp').read_bytes()
chunks=[logo[i:i+120].hex().upper() for i in range(0,len(logo),120)]
extra=['Private Sub WriteEmbeddedLogo(ByVal path As String)','    Dim f As Integer, n As Long, msg As String','    On Error GoTo Failed','    f = FreeFile','    Open path For Binary Access Write As #f']
for i in range(0,len(chunks),80):extra.append(f'    LogoPart{i//80} f')
extra += ['    Close #f','    Exit Sub','Failed:','    n = Err.Number: msg = Err.Description','    On Error Resume Next','    Close #f','    On Error GoTo 0','    Err.Raise n, "WriteEmbeddedLogo", msg','End Sub','Private Sub WriteHex(ByVal f As Integer, ByVal value As String)','    Dim b() As Byte, i As Long','    ReDim b(0 To Len(value) \\ 2 - 1)','    For i = 0 To UBound(b)','        b(i) = CByte("&H" & Mid$(value, i * 2 + 1, 2))','    Next i','    Put #f, , b','End Sub']
for i in range(0,len(chunks),80):
 extra += [f'Private Sub LogoPart{i//80}(ByVal f As Integer)']+[f'    WriteHex f, "{h}"' for h in chunks[i:i+80]]+['End Sub']
s+='\n'+'\n'.join(extra)+'\n'
assert bytes.fromhex(''.join(re.findall(r'    WriteHex f, "([0-9A-F]+)"',s)))==logo
assert 'GetOpenFilename' not in s
assert 'Public Sub ZEKON_BUDUJ_RC2B_B05()' in s
assert max(map(len,s.splitlines()))<1024
# Preserve every embedded production module, byte for byte at the line level.
for p in (root/'addin/src').iterdir():
 collected=[]
 for name,body in re.findall(r'Private Function (Source_\w+)\(\) As String\n(.*?)\nEnd Function',s,re.S):
  if name.startswith('Source_'+p.stem+'_'):
   collected += [x.replace('""','"') for x in re.findall(r'^    s = s & "(.*)" & vbCrLf$',body,re.M)]
 assert collected==p.read_text().splitlines(),p.name
out=root/'support/ZEKON_BUDUJ_RC2B_B05.bas'
out.write_bytes(s.replace('\n','\r\n').encode('ascii'))
with zipfile.ZipFile(root/'support/ZEKON_BUDUJ_RC2B_B05.zip','w',zipfile.ZIP_DEFLATED) as z:z.write(out,out.name)
print('PASS: embedded VBA unchanged; embedded logo matches; standalone B05; bounded procedures and line lengths.')
print(hashlib.sha256(out.read_bytes()).hexdigest())
