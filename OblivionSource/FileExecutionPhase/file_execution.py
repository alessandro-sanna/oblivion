import sys
import os
import shutil
import time
from itertools import product
import subprocess
from copy import deepcopy


class FileExecutionException(Exception):
    pass


class FileExecution:
    def __init__(self, running_file, output_file, output_file_in_sandbox, 
                 log_file, log_file_in_sandbox, instrumented_code_path,
                 sandbox_name, sandbox_exe, ext_info, no_clean_slate_flag):
        self.auto_open, self.auto_close = self.__get_auto_exec(instrumented_code_path)
        if not self.auto_open and not self.auto_close:
            raise FileExecutionException("Code cannot run itself.")

        self.running_file = running_file
        self.output_file = output_file
        self.output_file_in_sandbox = output_file_in_sandbox

        self.log_file = log_file
        self.log_file_in_sandbox = log_file_in_sandbox
        
        self.instrumented_code_path = instrumented_code_path

        self.ext_info = ext_info
        self.sandbox_name = sandbox_name
        self.sandbox_exe = sandbox_exe

        self.no_clean_slate_flag = no_clean_slate_flag

    def run(self):
        file_name = deepcopy(self.output_file_in_sandbox)

        command = self.__build_command()
        self.__launch(command)

        if not os.path.exists(file_name):
            raise FileExecutionException("File execution produced no output.")
        else:
            while not self.__is_file_available(file_name):
                time.sleep(1)

            shutil.copy2(file_name, self.output_file)


    def __build_command(self):
        script_name = os.path.join("OblivionSource", "FileExecutionPhase", "office_sandbox.py")
        flags = [str(int(f)) for f in (self.auto_open, self.auto_close, self.no_clean_slate_flag)]
        script_args = [self.running_file, self.instrumented_code_path,
                       self.ext_info["program"], self.ext_info["main_class"], self.ext_info["main_module"],
                       ] + flags

        command = [self.sandbox_exe, f"/box:{self.sandbox_name}", "/wait",
                   sys.executable.replace("python.exe", "pythonw.exe"),
                   script_name] + script_args + [self.log_file]
        
        return command
    
    def __launch(self, command):
        try:
            subprocess.check_call(command)
        except subprocess.CalledProcessError as exc:
            self.__print_crash(exc, exc.returncode)

    def __print_crash(self, exc=None, return_code=0):
        reason = f"File execution crashed, exit code = {return_code},"
        try:
            shutil.copy2(self.log_file_in_sandbox, self.log_file)
        except FileNotFoundError:
            reason += " with no log"
        else:
            with open(self.log_file, "r") as lf:
                output = [x for x in lf.readlines() if x][-1]
            reason += f": {output}." + f"Detailed log at {os.path.relpath(self.log_file)}"
        finally:
            message = f"{reason}."
            if exc is not None:
                message += f" Caught exception {exc}"
            raise FileExecutionException(message)

    @staticmethod
    def __get_auto_exec(instrumented_code_path) -> (bool, bool):
        with open(instrumented_code_path, "r") as icf:
            code = icf.read().lower()
            prefixes = ["auto", "document", "workbook"]
            joints = ["", "_"]
            suffixes = ["open", "close"]
            flags = list(False for _ in suffixes)
            
            keywords = [''.join(x) for x in product(prefixes, joints, suffixes)]
            
            for index in range(len(suffixes)):
                suffix = suffixes[index]
                check_list = [x for x in keywords if x.endswith(suffix)]
                for kw in check_list:
                    if kw in code:
                        flags[index] = True
                        break
        
        return flags

    @staticmethod
    def __is_file_available(file_path) -> bool:
        casted = str(file_path)

        try:
            if os.path.exists(casted):
                os.rename(casted, casted)
            else:
                return False
        except OSError:
            return False
        else:
            return True
