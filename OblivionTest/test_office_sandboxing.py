import os.path
import unittest
import pywintypes
import psutil
from OblivionSource.FileExecutionPhase import OfficeSandbox


class OfficeSandboxTesting(unittest.TestCase):
    def setUp(self) -> None:
        self.app = None
        self.program = "word"
        self.office_file = None
        self.running_file = os.path.abspath(os.path.join("OblivionTest", "test_files", "documento_prova.docm"))
        instrumented_code_path = os.path.abspath(os.path.join("OblivionTest", "test_files", "macro_dict_prova.pkl"))
        log_file = os.path.abspath(os.path.join("OblivionTest", "test_files", "log_file_prova.txt"))
        program, main_class, main_module = "word", "OpusApp", "document"
        auto_open, auto_close, no_clean_slate_flag = True, False, True
        self.TestClass = OfficeSandbox(self.running_file, instrumented_code_path, log_file,
                                       program, main_class, main_module,
                                       auto_open, auto_close, no_clean_slate_flag,
                                       log_flag=False)


class SingleMethods(OfficeSandboxTesting):
    def setUp(self) -> None:
        super().setUp()
        self.app = self.TestClass._OfficeSandbox__open_program()
        self.app.Visible = 0
        self.office_file = self.TestClass._OfficeSandbox__open_file(self.program, self.app, self.running_file)

    def tearDown(self) -> None:
        self.office_file.Close(0)
        self.app.Quit()

    def test_open_program(self):
        # In questo test __open_program non viene chiamato, perchè lo è nel setUp dei metodi
        processes = [y.name() for y in [psutil.Process(x) for x in psutil.pids()]
                     if y.status() == "running" and "word" in y.name().lower()]
        self.assertGreater(len(processes), 0)

    def test_open_file(self):
        # In questo test __open_file non viene chiamato, perchè lo è nel setUp dei metodi
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

        doc_name = self.office_file.Name
        self.assertEqual(file_basename, active_file)
        self.assertEqual(active_file, doc_name)

    def test_empty_macros(self):
        self.TestClass._OfficeSandbox__empty_macros(self.office_file)
        for macro in self.office_file.VBProject.VBComponents:
            self.assertEqual(macro.CodeModule.CountOfLines, 0)

    def test_make_macros(self):
        doc = self.TestClass._OfficeSandbox__make_macros(self.office_file)
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
        doc = self.TestClass._OfficeSandbox__write_macros(self.office_file)
        macro_dict = self.TestClass.macro_dict
        for macro in doc.VBProject.VBComponents:
            name = macro.Name
            if name == "vhook":
                continue
            try:
                ext = self.TestClass.macro_extension_code_dict_rev[macro.Type]
            except KeyError:
                ext = "cls"
            key = name + '.' + ext
            target_code = macro_dict[key]
            this_code = macro.CodeModule.Lines(1, macro.CodeModule.CountOfLines)
            self.assertEqual(this_code, target_code)


class GlobalRun(OfficeSandboxTesting):
    def test_run(self):
        self.TestClass.run()


if __name__ == '__main__':
    unittest.main()
