import importlib.util, tempfile, unittest, zipfile, re
from pathlib import Path
spec=importlib.util.spec_from_file_location('pack',Path(__file__).resolve().parents[1]/'tools/package_release.py')
module=importlib.util.module_from_spec(spec);spec.loader.exec_module(module)
VERSION=re.search(r'APP_VERSION As String = "([^"]+)"',(module.ROOT/'addin/src/modCore.bas').read_text()).group(1)
REPORT='PASS: block copy, batch rounding, cancellation, calculation restore, filtered split, filter criteria restore, permanent version label, zinc template, zinc xls roundtrip, zinc split export, zinc no overwrite, zinc row limit, readonly search copy, shared private copy, copy source unchanged, copy protection preserved'
class ReleaseGate(unittest.TestCase):
    def test_old_pass_report_rejected(self):
        with tempfile.TemporaryDirectory() as d:
            p=Path(d); (p/'report').write_text('PASS: rounding, filter')
            with self.assertRaisesRegex(ValueError, 'Missing full native'): module.package(p/'missing',p/'report',VERSION,6,p/'out.zip')
    def test_report_without_new_export_tests_rejected(self):
        with tempfile.TemporaryDirectory() as d:
            p=Path(d); (p/'report').write_text(REPORT.split(', zinc template')[0])
            with self.assertRaisesRegex(ValueError, 'Missing full native'): module.package(p/'missing',p/'report',VERSION,6,p/'out.zip')
    def test_report_without_copy_tests_rejected(self):
        with tempfile.TemporaryDirectory() as d:
            p=Path(d); (p/'report').write_text(REPORT.split(', readonly search copy')[0])
            with self.assertRaisesRegex(ValueError, 'Missing full native'): module.package(p/'missing',p/'report',VERSION,7,p/'out.zip')
    def test_version_mismatch_rejected(self):
        with self.assertRaises(ValueError): module.package(Path('missing'),Path('missing'),'9.9.9',2,Path('out'))
    def test_non_xlam_rejected(self):
        with tempfile.TemporaryDirectory() as d:
            p=Path(d);(p/'report').write_text(REPORT)
            with zipfile.ZipFile(p/'bad.xlam','w') as z:z.writestr('data','not VBA')
            with self.assertRaisesRegex(ValueError, 'Missing VBA project'):module.package(p/'bad.xlam',p/'report',VERSION,6,p/'out.zip')
if __name__=='__main__':unittest.main()
