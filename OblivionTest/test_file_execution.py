import unittest
import os
from ..OblivionSource.FileExecutionPhase import FileExecution, FileExecutionException


class MyTestCase(unittest.TestCase):
    def __init__(self):
        super().__init__()
        self.TestClass = FileExecution
    
    def test_init(self):
        pass

    def test_run(self):
        pass

    def test_build_command(self):
        pass

    def test_launch(self, command):
        pass

    def test_get_auto_exec(self) -> (bool, bool):
        with open(os.path.join("test_files", "sample_macro_code.txt"), "r") as fb:
            macro_sample = fb.read()
        self.assertEqual((True, True), self.TestClass.__get_auto_exec(macro_sample))
        pass


if __name__ == '__main__':
    unittest.main()
