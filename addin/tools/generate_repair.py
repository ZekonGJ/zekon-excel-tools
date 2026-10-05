"""Build a repair installer carrying current sources, independent of an older unpacked src folder."""
from pathlib import Path
import re
r=Path(__file__).resolve().parents[1]
s=(r/'ZbudujDodatek.bas').read_text().replace('Attribute VB_Name = "ZekonBuilder"','Attribute VB_Name = "ZekonInstallerRC2"').replace('Public Sub ZbudujDodatekZekon()','Public Sub AktualizujZekonRC2()')
s=s.replace('    On Error GoTo Bad\n', '    Application.EnableEvents = True\n    Application.ScreenUpdating = True\n    On Error GoTo Bad\n',1)
s=s.replace('        sourceCode = ReadUtf8(root & "src\\" & item & ".bas")','        sourceCode = EmbeddedSource(CStr(item))')
s=s.replace('ReadUtf8(root & "src\\frmZekon.vba")','EmbeddedSource("frmZekon")').replace('ReadUtf8(root & "src\\ThisWorkbook.vba")','EmbeddedSource("ThisWorkbook")')
extra=['Private Function EmbeddedSource(ByVal moduleName As String) As String','    Select Case moduleName']
chunks=[]
for p in sorted((r/'src').iterdir()):
 lines=p.read_text().splitlines()
 names=[]
 for start in range(0,len(lines),40):
  name='Source_'+p.stem+'_'+str(start//40);names.append(name)
  chunks += ['Private Function '+name+'() As String','    Dim s As String']
  chunks += ['    s = s & "'+line.replace('"','""')+'" & vbCrLf' for line in lines[start:start+40]]
  chunks += ['    '+name+' = s','End Function','']
 extra.append('        Case "'+p.stem+'"')
 extra += ['            EmbeddedSource = EmbeddedSource & '+name+'()' for name in names]
extra += ['        Case Else: Err.Raise vbObjectError + 908, , "Nieznany modul instalatora"','    End Select','End Function','']
s+='\n'+'\n'.join(extra+chunks)
assert max(map(len,s.splitlines()))<1024
# Verify every embedded line against the actual module sources.
for p in (r/'src').iterdir():
 collected=[]
 for name,body in re.findall(r'Private Function (Source_\w+)\(\) As String\n(.*?)\nEnd Function',s,re.S):
  if name.startswith('Source_'+p.stem+'_'):
   collected += [x.replace('""','"') for x in re.findall(r'^    s = s & "(.*)" & vbCrLf$',body,re.M)]
 assert collected==p.read_text().splitlines(),p.name
(r.parent/'ZbudujDodatek_NAPRAWA.bas').write_bytes(s.replace('\n','\r\n').encode('ascii'))
print('PASS: installer embeds every current module exactly; bounded source chunks and VBA line lengths.')
