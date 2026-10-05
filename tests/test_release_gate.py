import importlib.util, tempfile, unittest, zipfile
from pathlib import Path
spec=importlib.util.spec_from_file_location('pack',Path(__file__).resolve().parents[1]/'tools/package_release.py')
module=importlib.util.module_from_spec(spec);spec.loader.exec_module(module)
class ReleaseGate(unittest.TestCase):
    def test_old_pass_report_rejected(self):
        with tempfile.TemporaryDirectory() as d:
            p=Path(d); (p/'report').write_text('PASS: rounding, filter')
            with self.assertRaises(ValueError): module.package(p/'missing',p/'report','1.0.0-rc2',2,p/'out.zip')
    def test_version_mismatch_rejected(self):
        with self.assertRaises(ValueError): module.package(Path('missing'),Path('missing'),'9.9.9',2,Path('out'))
    def test_non_xlam_rejected(self):
        with tempfile.TemporaryDirectory() as d:
            p=Path(d);(p/'report').write_text('PASS: block copy, batch rounding, cancellation, calculation restore')
            with zipfile.ZipFile(p/'bad.xlam','w') as z:z.writestr('data','not VBA')
            with self.assertRaises(ValueError):module.package(p/'bad.xlam',p/'report','1.0.0-rc2',2,p/'out.zip')
if __name__=='__main__':unittest.main()
