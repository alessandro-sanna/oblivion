import os
import json
import time
from numpy import around
from types import SimpleNamespace
import multiprocessing
import threading
from easyprocess import EasyProcess
# Oblivion phases
from OblivionSource.MacroExtractionPhase import MacroExtraction, MacroExtractionException
from OblivionSource.MacroInstrumentationPhase import MacroInstrumentation, MacroInstrumentationException
from OblivionSource.FileExecutionPhase import FileExecution, FileExecutionException
from OblivionSource.PostProcessingPhase import PostProcessing, PostProcessingException
from OblivionPlugins.InteractionManager import InteractionManager


class OblivionCoreException(Exception):
    pass


class OblivionCore:
    def __init__(self, arg_namespace):
        self.args = arg_namespace

        self.current_original_file = None
        self.current_modified_file = None
        self.current_extension = None
        self.current_output_file = None
        self.current_report_file = None

        self.current_sandbox_original = None
        self.current_sandbox_modified = None
        self.current_sandbox_output = None
        self.current_sandbox_report = None

        with open(os.path.join("OblivionResources", "config", "configuration.json"), "r") as cj:
            self.config = json.load(cj, object_hook=lambda d: SimpleNamespace(**d))

        with open(os.path.join("OblivionResources", "config", "extensions.json"), "r") as cj:
            self.ext_info = json.load(cj)
            self.extensions = list(self.ext_info.keys())

        self.original_macro_path = os.path.join("OblivionResources", "files", "original_macro.txt")
        self.original_macro_data_path = os.path.join("OblivionResources", "data", "original_macro_data.json")

        self.instrumented_macro_data = os.path.join("OblivionResources", "files", "instrumented_macro.txt")
        self.instrumented_macro_data_path = os.path.join("OblivionResources", "data", "instrumented_macro_data.json")

        self.exclusion_path = None
        self.interaction_manager_enabled = True

    def execute(self, single=False):
        if single:
            self.__execute_on_file()
        else:
            self.__execute_on_folder()

    def __execute_on_folder(self):
        tf = self.args.target_folder
        file_list = [os.path.join(tf, x) for x in os.listdir(tf) if x.split('.')[-1] in self.extensions]

        for fp in file_list:
            self.args.target_file = fp
            self.__execute_on_file()

    def __execute_on_file(self):
        pool = multiprocessing.Pool(processes=1)
        async_obj = pool.apply_async(self.run)

        self.current_original_file = self.args.target_file
        self.current_extension = self.args.target_file.split(".")[-1]
        self.current_sandbox_original = self.__path_to_sandbox(self.args.target_file)

        self.current_modified_file, \
            self.current_output_file, self.current_report_file = self.__path_to_output(self.args.target_file)

        self.current_sandbox_modified,\
            self.current_sandbox_output, self.current_sandbox_report = \
            (self.__path_to_sandbox(x) for x in self.__path_to_output(self.args.target_file))

        try:
            async_obj.get(timeout=self.args.time_limit)
        except multiprocessing.TimeoutError:
            raise OblivionCoreException("[-] Timeout!")  # handle

    def run(self):
        print(f"[?] Current sample: {os.path.basename(self.current_original_file)}")
        starting_time = time.time()
        try:
            # Preliminary
            self.__clean_sandbox()
            # Phases
            self.__macro_extraction()
            self.__macro_instrumentation()
            # Dynamic Analysis
            int_thread, enable_event = self.__interaction_manager()
            self.__file_extraction()
            enable_event.clear()
            int_thread.join()
            self.__post_processing()
        except MacroExtractionException as exc:
            raise OblivionCoreException(f"[-] Macro extraction failed: {exc}")  # handle
        except MacroInstrumentationException as exc:
            raise OblivionCoreException(f"[-] Macro instrumentation failed: {exc}")  # handle
        except FileExecutionException as exc:
            raise OblivionCoreException(f"[-] File execution failed: {exc}")  # handle
            # repeating process must go here
        except PostProcessingException as exc:
            raise OblivionCoreException(f"[-] Report generation failed: {exc}")  # handle
        else:
            pass  # handle
        finally:
            ending_time = time.time()
            self.__clean_sandbox()  # handle
            print(f"[?] Analysis time: {around(ending_time - starting_time, decimals=2)}")
            pass

    def __interaction_manager(self):
        enable_event = threading.Event()
        enable_event.set()
        phase = InteractionManager(self.current_original_file, self.ext_info[self.current_extension], enable_event)
        int_thread = threading.Thread(target=phase.run, daemon=True)
        int_thread.start()

        return int_thread, enable_event

    def __macro_extraction(self):
        phase = MacroExtraction(self.current_original_file, self.extensions, self.original_macro_data_path)
        macro_data = phase.run()

        with open(self.original_macro_path, "w") as foObj:
            for macro_name, macro_code in macro_data.items():
                foObj.write(f"{macro_name}\n{macro_code}\n\n")

        print("[+] Macro successfully extracted")

    def __macro_instrumentation(self):
        phase = MacroInstrumentation(self.original_macro_data_path, self.exclusion_path,
                                     self.instrumented_macro_data_path)
        phase.run()
        print("[+] Macro successfully instrumented")

    def __file_extraction(self):
        log_path = os.path.join("OblivionResources", "logs", "sbx_out_err.log")
        log_path_sbx = self.__path_to_sandbox(log_path)

        phase = FileExecution(self.current_original_file, self.current_output_file, self.current_sandbox_output,
                              log_path, log_path_sbx, self.instrumented_macro_data_path, self.config.Sandbox_name,
                              self.config.Sandboxie_path, self.ext_info[self.current_extension],
                              self.args.no_clean_slate)
        phase.run()
        print("[+] File successfully executed")

    def __post_processing(self):
        program = self.ext_info[self.current_extension]["program"]
        phase = PostProcessing(self.current_original_file, self.current_output_file, self.original_macro_path,
                               self.current_extension, self.config.PowerDecode_path, self.config.Sandboxie_path,
                               self.config.Sandbox_name, program, self.current_report_file)
        phase.run()
        print("[+] Report successfully generated")

    @staticmethod
    def __path_to_output(path):
        ext = path.split('.')[-1]
        modified_path = path.replace('.' + ext, f"_obl3_modified.{ext}")
        # output_path = path.replace('.' + ext, f"_{ext}_output.txt")
        output_path = path + ".txt"
        report_path = path.replace('.' + ext, f"_{ext}_report.txt")
        return modified_path, output_path, report_path

    def __path_to_sandbox(self, path):
        old_root = os.path.join("C:\\", "Users", os.getenv("username"))
        new_root = os.path.join("C:\\", "Sandbox", os.getenv("username"), self.config.Sandbox_name, "user", "current")
        sandbox_path = path.replace(old_root, new_root)
        return sandbox_path

    def __clean_sandbox(self):
        EasyProcess([self.config.Sandboxie_path, "/terminate_all"]).call().wait()
        EasyProcess([self.config.Sandboxie_path, "delete_sandbox_silent"]).call().wait()
