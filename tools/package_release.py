"""Package a reviewed native-Excel build; never synthesize a working XLAM from text."""
import argparse, hashlib, json, re, zipfile
from pathlib import Path
import xml.etree.ElementTree as ET
ROOT = Path(__file__).resolve().parents[1]
def package(xlam, report, version, number, output):
    if not re.fullmatch(r'\d+\.\d+\.\d+(?:-[A-Za-z0-9.]+)?', version) or number < 1:
        raise ValueError('Invalid release version/number')
    source = (ROOT/'addin/src/modCore.bas').read_text()
    if f'APP_VERSION As String = "{version}"' not in source:
        raise ValueError('Version does not match VBA source')
    text = report.read_text(encoding='utf-8-sig')
    required = ('PASS:', 'block copy', 'batch rounding', 'cancellation', 'calculation restore', 'filtered split', 'filter criteria restore', 'permanent version label')
    if not all(marker in text for marker in required):
        raise ValueError('Missing full native rc2+ Excel PASS report; do not release an untested add-in')
    with zipfile.ZipFile(xlam) as book:
        if 'xl/vbaProject.bin' not in book.namelist():
            raise ValueError('Missing VBA project')
        if 'application/vnd.ms-excel.addin.macroEnabled.main+xml' not in book.read('[Content_Types].xml').decode():
            raise ValueError('File is not an XLAM')
        properties = ET.fromstring(book.read('docProps/core.xml'))
        if not any(version in (node.text or '') for node in properties):
            raise ValueError('XLAM properties do not identify the expected version')
    files={'ZekonTools.xlam':xlam.read_bytes(), 'logo.bmp':(ROOT/'addin/assets/logo.bmp').read_bytes()}
    manifest={'ReleaseNumber':number, 'Version':version, 'AddinFile':'ZekonTools.xlam',
              'Sha256':{name:hashlib.sha256(data).hexdigest().upper() for name,data in files.items()}}
    output.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(output,'w',zipfile.ZIP_DEFLATED) as out:
        out.writestr('manifest.json',json.dumps(manifest))
        for name,data in files.items(): out.writestr(name,data)
    return manifest
if __name__=='__main__':
    ap=argparse.ArgumentParser()
    ap.add_argument('--xlam',type=Path,default=ROOT/'release-input/ZekonTools.xlam')
    ap.add_argument('--report',type=Path,default=ROOT/'release-input/test-result.txt')
    ap.add_argument('--version',default='1.0.0-rc2b')
    ap.add_argument('--release-number',type=int,default=4)
    ap.add_argument('--output',type=Path,default=ROOT/'installer/payload.zip')
    a=ap.parse_args()
    try: print(json.dumps(package(a.xlam,a.report,a.version,a.release_number,a.output),indent=2))
    except (OSError, ValueError, zipfile.BadZipFile, KeyError) as e: ap.exit(1, f'RELEASE BLOCKED: {e}\n')
