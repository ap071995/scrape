"""Exercise the Windows launcher with isolated, non-GUI application fixtures."""

import json
import os
from pathlib import Path
import shutil
import subprocess
import tempfile
import time
import unittest
import venv


WORKSPACE = Path(__file__).resolve().parent
LAUNCHER = WORKSPACE / "Run_Scraper.bat"
SUPPORT_MODULES = (
    "business_email", "company_public_data", "company_search_dom", "google_maps_dom",
    "email_domain_check", "maps_office_scope", "benchmark", "run_support", "enrichment_session",
    "linkedin_links", "scraper_ui", "link_import", "worksheet_editor"
)


def remove_fixture(path):
    """Refuse recursive cleanup outside this test's workspace."""
    resolved = Path(path).resolve()
    if resolved == WORKSPACE or not resolved.is_relative_to(WORKSPACE):
        raise AssertionError(f"Unsafe launcher fixture cleanup: {resolved}")
    shutil.rmtree(resolved)


@unittest.skipUnless(os.name == "nt", "The launcher requires Windows cmd.exe")
class LauncherTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.root = Path(tempfile.mkdtemp(prefix="launcher-tests-", dir=WORKSPACE))
        cls.addClassCleanup(remove_fixture, cls.root)
        # These disposable runners reference the installed Python. No packages or
        # existing user environments are installed, modified, or imported.
        cls.runtime = cls.root / "fixture-runtime"
        venv.EnvBuilder(with_pip=False).create(cls.runtime)
        cls.cmd = str(Path(os.environ["SystemRoot"]) / "System32" / "cmd.exe")
        cls.system_path = str(Path(os.environ["SystemRoot"]) / "System32")

    def setUp(self):
        self.case = Path(tempfile.mkdtemp(prefix="case-", dir=self.root))
        self.addCleanup(remove_fixture, self.case)
        self.app = self.case / "Application with spaces ! & (release)"
        self.app.mkdir()
        self.caller = self.case / "different working directory"
        self.caller.mkdir()
        shutil.copy2(LAUNCHER, self.app / LAUNCHER.name)
        (self.app / "Scraper.py").write_text(
            "import json, os, sys\n"
            "from pathlib import Path\n"
            "APP_VERSION = 'fixture'\n"
            "if __name__ == '__main__':\n"
            "    Path('launch.json').write_text(json.dumps({\n"
            "        'cwd': os.getcwd(), 'file': str(Path(__file__).resolve()),\n"
            "        'python': sys.executable, 'argv': sys.argv[1:]}))\n"
            "    raise SystemExit(int(os.environ.get('FIXTURE_EXIT_CODE', '0')))\n",
            encoding="utf-8",
        )
        for module in SUPPORT_MODULES:
            (self.app / (module + ".py")).write_text("# Fixture module.\n", encoding="utf-8")
        self.env = os.environ.copy()
        self.env.update({"PATH": self.system_path, "PYTHONNOUSERSITE": "1"})
        for name in ("PYTHONPATH", "PYTHONHOME", "VIRTUAL_ENV", "FIXTURE_EXIT_CODE"):
            self.env.pop(name, None)

    def install_runtime(self, name=".venv"):
        destination = self.app / name
        shutil.copytree(self.runtime, destination)
        return destination / "Scripts" / "python.exe"

    def prepare_windowless(self):
        shutil.copy2(WORKSPACE / 'Run_Scraper.vbs', self.app / 'Run_Scraper.vbs')
        self.logs_root = self.case / 'Local App Data with spaces & symbols'
        self.logs_root.mkdir()
        self.env['LOCALAPPDATA'] = str(self.logs_root)

    def run_windowless(self, check=False):
        script_host = str(Path(os.environ['SystemRoot']) / 'System32' / 'cscript.exe')
        return subprocess.run([script_host, '//nologo', str(self.app / 'Run_Scraper.vbs')]
                              + (['--check'] if check else []),
                              cwd=self.caller, env=self.env, stdin=subprocess.DEVNULL,
                              stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                              text=True, timeout=20)

    def test_windowless_check_creates_log_without_starting_app(self):
        self.prepare_windowless()
        self.install_runtime()
        result = self.run_windowless(check=True)
        self.assertEqual(result.returncode, 0, result.stdout)
        self.assertEqual(result.stdout, '')
        self.assertFalse((self.app / 'launch.json').exists())
        logs = list(self.logs_root.rglob('*.log'))
        self.assertEqual(len(logs), 1)
        self.assertIn('vfixture is ready', logs[0].read_text())

    def test_windowless_launch_keeps_correct_directory_and_interpreter(self):
        self.prepare_windowless()
        python = self.install_runtime()
        result = self.run_windowless()
        self.assertEqual(result.returncode, 0, result.stdout)
        self.assertEqual(result.stdout, '')
        launch = json.loads((self.app / 'launch.json').read_text())
        self.assertEqual(Path(launch['cwd']), self.app)
        self.assertEqual(Path(launch['python']), python)
        self.assertEqual(launch['argv'], [])

    def test_windowless_failed_check_preserves_diagnostic_and_exit_code(self):
        self.prepare_windowless()
        self.install_runtime()
        (self.app / 'google_maps_dom.py').write_text('import absent_fixture_dependency\n')
        result = self.run_windowless(check=True)
        self.assertNotEqual(result.returncode, 0)
        self.assertEqual(result.stdout, '')
        self.assertFalse((self.app / 'launch.json').exists())
        logs = list(self.logs_root.rglob('*.log'))
        self.assertEqual(len(logs), 1)
        self.assertIn('absent_fixture_dependency', logs[0].read_text())

    def test_background_failure_returns_without_pause(self):
        self.install_runtime()
        result = subprocess.run(
            f'"{self.cmd}" /d /v:off /s /c ""{self.app / LAUNCHER.name}" --background"',
            cwd=self.caller, env={**self.env, 'FIXTURE_EXIT_CODE': '23'},
            stdin=subprocess.DEVNULL, stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
            text=True, timeout=20)
        self.assertEqual(result.returncode, 23, result.stdout)
        self.assertNotIn('Press any key', result.stdout)

    def test_old_batch_shortcut_delegates_to_windowless_launcher(self):
        self.prepare_windowless()
        self.install_runtime()
        result = self.run_launcher(check=False)
        self.assertEqual(result.returncode, 0, result.stdout)
        deadline = time.monotonic() + 10
        while not (self.app / 'launch.json').exists() and time.monotonic() < deadline:
            time.sleep(0.05)
        self.assertTrue((self.app / 'launch.json').exists())
        logs = list(self.logs_root.rglob('*.log'))
        self.assertEqual(len(logs), 1)

    def run_launcher(self, check=True, **environment):
        arguments = " --check" if check else ""
        result = subprocess.run(
            f'"{self.cmd}" /d /v:off /s /c ""{self.app / LAUNCHER.name}"{arguments}"',
            cwd=self.caller,
            env={**self.env, **environment},
            stdin=subprocess.DEVNULL,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            timeout=20,
        )
        if check:
            self.assertNotIn("Press any key", result.stdout, result.stdout)
        return result

    def assert_failed(self, result, message):
        self.assertNotEqual(result.returncode, 0, result.stdout)
        self.assertIn(message, result.stdout)
        self.assertFalse((self.app / "launch.json").exists())

    def test_check_in_special_character_path_does_not_launch(self):
        self.install_runtime()
        result = self.run_launcher()
        self.assertEqual(result.returncode, 0, result.stdout)
        self.assertIn("vfixture is ready", result.stdout)
        self.assertFalse((self.app / "launch.json").exists())

    def test_main_receives_application_path_and_working_directory(self):
        python = self.install_runtime()
        result = self.run_launcher(check=False)
        self.assertEqual(result.returncode, 0, result.stdout)
        launch = json.loads((self.app / "launch.json").read_text())
        self.assertEqual(Path(launch["cwd"]), self.app)
        self.assertEqual(Path(launch["file"]), self.app / "Scraper.py")
        self.assertEqual(Path(launch["python"]), python)
        self.assertEqual(launch["argv"], [])

    def test_main_failure_preserves_exit_code(self):
        self.install_runtime()
        result = self.run_launcher(check=False, FIXTURE_EXIT_CODE="23")
        self.assertEqual(result.returncode, 23, result.stdout)
        self.assertIn("application stopped with an error", result.stdout)

    def test_missing_application_is_reported(self):
        (self.app / "Scraper.py").unlink()
        self.assert_failed(self.run_launcher(), "Scraper.py is missing")

    def test_missing_support_module_is_reported(self):
        self.install_runtime()
        for module in SUPPORT_MODULES:
            with self.subTest(module=module):
                path = self.app / (module + ".py")
                path.unlink()
                result = self.run_launcher()
                self.assert_failed(result, "requirements check failed")
                self.assertIn(module, result.stdout)
                path.write_text("# Restored fixture module.\n", encoding="utf-8")

    def test_missing_third_party_package_is_reported(self):
        self.install_runtime()
        (self.app / "google_maps_dom.py").write_text("import absent_fixture_dependency\n")
        result = self.run_launcher()
        self.assert_failed(result, "requirements check failed")
        self.assertIn("absent_fixture_dependency", result.stdout)
        self.assertIn("-m pip install", result.stdout)

    def test_no_python_is_reported_without_waiting(self):
        self.assert_failed(self.run_launcher(), "Python 3 was not found")

    def test_inherited_exit_code_cannot_hide_failure(self):
        self.assert_failed(self.run_launcher(SCRAPER_EXIT_CODE="0"), "Python 3 was not found")

    def test_negative_requirement_exit_code_prevents_launch(self):
        self.install_runtime()
        (self.app / "google_maps_dom.py").write_text("raise SystemExit(-1)\n")
        self.assert_failed(self.run_launcher(), "requirements check failed")

    def test_broken_local_environment_stops_before_launch(self):
        python = self.install_runtime()
        (python.parent.parent / "pyvenv.cfg").write_text(
            f"home = {self.case / 'missing-python-installation'}\n"
            "include-system-site-packages = false\n"
        )
        self.assert_failed(self.run_launcher(), "Python interpreter could not start")

    def test_dot_venv_preferred_to_other_environment_names(self):
        python = self.install_runtime()
        for name in ("env", "venv"):
            folder = self.app / name / "Scripts"
            folder.mkdir(parents=True)
            (folder / "python.exe").write_bytes(b"not an executable")
        result = self.run_launcher(check=False)
        self.assertEqual(result.returncode, 0, result.stdout)
        launch = json.loads((self.app / "launch.json").read_text())
        self.assertEqual(Path(launch["python"]), python)

    def test_alternative_local_environment_names(self):
        for name in ("env", "venv"):
            with self.subTest(name=name):
                python = self.install_runtime(name)
                result = self.run_launcher()
                self.assertEqual(result.returncode, 0, result.stdout)
                remove_fixture(python.parent.parent)

    def test_python_on_path_fallback(self):
        python = self.install_runtime("path-runtime")
        result = self.run_launcher(PATH=str(python.parent) + os.pathsep + self.system_path)
        self.assertEqual(result.returncode, 0, result.stdout)
        self.assertIn("vfixture is ready", result.stdout)

    def test_py_launcher_preferred_to_python_on_path(self):
        py_launcher = shutil.which("py.exe")
        if py_launcher is None:
            self.skipTest("Python launcher is not installed")
        broken_python = self.case / "path-python"
        broken_python.mkdir()
        (broken_python / "python.exe").write_bytes(b"not an executable")
        result = self.run_launcher(PATH=os.pathsep.join(
            (str(broken_python), str(Path(py_launcher).parent), self.system_path)
        ))
        self.assertEqual(result.returncode, 0, result.stdout)
        self.assertIn("vfixture is ready", result.stdout)


if __name__ == "__main__":
    unittest.main()
