import multiprocessing
import sys
import os
import shutil
from itertools import product


class FileExecutionException(Exception):
    pass


class FileExecution:
    def __init__(self, running_file, output_file, output_file_in_sandbox, instrumented_code_path,
                 sandbox_name, sandbox_exe, no_office_mode_flag, no_clean_slate_flag):
        self.auto_open, self.auto_close = self.__get_auto_exec(instrumented_code)
        if not self.auto_open and not self.auto_close:
            raise FileExecutionException("Code cannot run itself.")

        self.running_file = running_file
        self.output_file = output_file
        self.output_file_in_sandbox = output_file_in_sandbox
        self.instrumented_code_path = instrumented_code_path

        self.sandbox_name = sandbox_name
        self.sandbox_exe = sandbox_exe

        self.no_office_mode_flag = no_office_mode_flag
        self.no_clean_slate_flag = no_clean_slate_flag

    def run(self):
        command = self.__build_command()
        self.__launch(commmand)
        self.__retrieve_output()

    def __build_command(self):
        if not self.no_office_mode_flag:
            script_name = "office_sandbox.pyw"
            script_args = [self.running_file, self.instrumented_code, 
                           self.auto_open, self.auto_close, self.no_clean_slate_flag]
        else:
            script_name = "wscript_sandbox.pyw"
            script_args = [self.instrumented_code]

        command = [self.sandbox_exe, f"/box:{self.sandbox_name}",
                   sys.executable.replace("python.exe", "pythonw.exe"), 
                   script_name] + script_args
        
        return command
    
    def __launch(self, command):
        pass
    
    def __retrieve_output(self):
        if os.path.exists(self.output_file_in_sandbox):
            shutil.copy2(self.output_file_in_sandbox, self.output_file)
        else:
            raise FileExecutionException("File execution produced no output")
    
    @staticmethod
    def __get_auto_exec(instrumented_code_path) -> (bool, bool):
        with open(instrumented_code_path, "r") as icf:
            code = icf.read().lower()
            prefixes = ["auto", "document", "workbook"]
            joints = ["", "_"]
            suffixes = ["open", "close"]
            flags = (False for _ in suffixes)
            
            keywords = [''.join(x) for x in product(prefixes, joints, suffixes)]
            
            for index in range(len(suffixes)):
                suffix = suffixes[index]
                check_list = [x for x in keywords if x.endswith(suffix)]
                for kw in check_list:
                    if kw in code:
                        flags[index] = True
                        break
        
        return flags
