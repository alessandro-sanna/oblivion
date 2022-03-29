import os.path
import unittest
from OblivionSource.FileExecutionPhase import OfficeSandbox
import psutil


class MyTestCase(unittest.TestCase):
    def __init__(self):
        super().__init__()
        self.TestClass = OfficeSandbox()
        self.app = None

    def test_build_strings(self):
        pass

    def test_get_instrumented_macros(self):
        pass

    def test_open_program(self):
        app = self.TestClass.__open_program()
        self.assertIsNotNone(app)
        processes = [y.name() for y in [psutil.Process(x) for x in psutil.pids()]
                     if y.status() == "running" and "word" in y.name().lower()]
        self.assertGreater(len(processes), 0)
        self.app = app

    def test_open_file(self):
        doc = self.TestClass.__open_file(self.TestClass.program, self.app, self.TestClass.running_file)
        file_basename = os.path.basename(self.TestClass.running_file)
        if "word" == self.TestClass.program:
            active_file = self.app.ActiveDocument.Name
        elif "excel" == self.TestClass.program:
            active_file = self.app.ActiveWorkbook.Name
        else:
            active_file = None
        doc_name = doc.Name

        self.assertEqual(file_basename, active_file)
        self.assertEqual(active_file, doc_name)

    def test_replace_macros(self, office_file):
        pass

    @staticmethod
    def test_empty_macros(office_file):
        pass

    def test_make_macros(self, office_file):
        pass

    def test_write_macros(self, office_file):
        pass

    @staticmethod
    def test_close_file(office_file):
        pass

    @staticmethod
    def test_close_program(app):
        pass


if __name__ == '__main__':
    unittest.main()
