import os.path
import unittest

import pywintypes

from OblivionSource.FileExecutionPhase import OfficeSandbox
import psutil


class MyTestCase(unittest.TestCase):
    def __init__(self):
        super().__init__()

        self.app = None
        self.program = "word"
        self.office_file = None

        self.TestClass = OfficeSandbox()

    def test_open_program(self):
        app = self.TestClass.__open_program()
        self.assertIsNotNone(app)
        processes = [y.name() for y in [psutil.Process(x) for x in psutil.pids()]
                     if y.status() == "running" and "word" in y.name().lower()]
        self.assertGreater(len(processes), 0)
        self.app = app

    def test_open_file(self):
        doc = self.TestClass.__open_file(self.program, self.app, self.TestClass.running_file)
        file_basename = os.path.basename(self.TestClass.running_file)
        try:
            if "word" == self.program:
                active_file = self.app.ActiveDocument.Name
            elif "excel" == self.program:
                active_file = self.app.ActiveWorkbook.Name
            else:
                active_file = None
        except pywintypes.com_error as exc:
            self.assertNotEqual(exc.excepinfo[-2], 37016)
            # questo codice di errore significa che il documento non è stato aperto

        doc_name = doc.Name
        self.assertEqual(file_basename, active_file)
        self.assertEqual(active_file, doc_name)
        self.office_file = doc

    def test_empty_macros(self):
        self.TestClass.__empty_macros(self.office_file)
        for macro in self.office_file.VBProject.VBComponents:
            self.assertEqual(macro.CodeModule.CountOfLines, 0)

    def test_make_macros(self):
        self.TestClass.__get_instrumented_macros()
        doc = self.TestClass.__make_macros(self.office_file)
        macro_dict = self.TestClass.macro_dict
        for macro in doc.VBProject.VBComponents:
            name = macro.Name
            try:
                ext = self.TestClass.macro_extension_code_dict_rev[macro.Type]
            except KeyError:
                ext = "cls"
            key = name + '.' + ext
            self.assertIn(key, macro_dict.keys())

    def test_write_macros(self):
        self.TestClass.__get_instrumented_macros()
        doc = self.TestClass.__write_macros(self.office_file)
        macro_dict = self.TestClass.macro_dict
        for macro in doc.VBProject.VBComponents:
            name = macro.Name
            try:
                ext = self.TestClass.macro_extension_code_dict_rev[macro.Type]
            except KeyError:
                ext = "cls"
            key = name + '.' + ext
            target_code = macro_dict[key]
            this_code = macro.CodeModule.Lines(1, macro.CodeModule.CountOfLines)
            self.assertEqual(this_code, target_code)


if __name__ == '__main__':
    unittest.main()
