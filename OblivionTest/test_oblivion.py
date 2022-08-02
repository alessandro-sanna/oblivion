import time
import unittest
import sys
import subprocess
import os


class MyTestCase(unittest.TestCase):
    def setUp(self) -> None:
        self.script = "oblivion.py"
        os.chdir(r"..")

    @staticmethod
    def run_command(script, command):
        args = command.split()
        subprocess.check_output([sys.executable, script] + args)

    def test_timeout(self):
        timeout = 10
        command = f"-f OblivionTest/test_files/auto_both_test.docm -o OblivionTest/test_out -t {timeout}"
        tic = time.time()
        self.run_command(self.script, command)
        toc = time.time()
        self.assertLessEqual(toc - tic, timeout + 1)

    def test_run(self):
        command = f"-f OblivionTest/test_files/auto_both_test.docm -o OblivionTest/test_out"
        self.run_command(self.script, command)
        self.assertTrue(os.path.exists(r"OblivionTest/test_out/auto_both_test_output.docm.txt"))


if __name__ == '__main__':
    unittest.main()
