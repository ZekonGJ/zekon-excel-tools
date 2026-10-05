"""Validate published, native-tested packages independently of pending VBA source edits."""
import hashlib,json,zipfile,io,shutil,argparse
from pathlib import Path
r=Path(__file__).resolve().parents[1]
def verify():
    offer=json.loads((r/'updates/latest.json').read_text())
    import re
    v=offer['Version']
    if offer['SchemaVersion']!=1 or offer['ReleaseNumber']<1 or not re.fullmatch(r'\d+\.\d+\.\d+(?:-[A-Za-z0-9.]+)?',v):raise ValueError('Invalid feed')
    p=r/'updates'/f'ZekonTools_{v}.zekonupdate'
    data=p.read_bytes()
    assert hashlib.sha256(data).hexdigest().upper()==offer['Sha256']
    with zipfile.ZipFile(io.BytesIO(data)) as z:
        assert sorted(z.namelist())==['ZekonTools.xlam','logo.bmp','manifest.json']
        m=json.loads(z.read('manifest.json'))
        assert m['Version']==v and m['ReleaseNumber']==offer['ReleaseNumber']
        assert set(m['Sha256'])=={'ZekonTools.xlam','logo.bmp'}
        for name,h in m['Sha256'].items():assert hashlib.sha256(z.read(name)).hexdigest().upper()==h
        with zipfile.ZipFile(io.BytesIO(z.read('ZekonTools.xlam'))) as x:
            assert 'xl/vbaProject.bin' in x.namelist()
            assert v in x.read('docProps/core.xml').decode()
    proof=json.loads((r/'updates'/f'proof_{v}.json').read_text())
    assert proof['XlamSha256']==m['Sha256']['ZekonTools.xlam'] and 'PASS:' in proof['NativeExcelReport']
    print('PASS: published feed, package hashes, version and recorded native Excel report:',v)
    return p
if __name__=='__main__':
    ap=argparse.ArgumentParser();ap.add_argument('--export-dir',type=Path);a=ap.parse_args()
    p=verify()
    if a.export_dir:a.export_dir.mkdir(parents=True,exist_ok=True);shutil.copyfile(p,a.export_dir/p.name)
