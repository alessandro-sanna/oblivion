import os
import json
import time
from numpy import around
from types import SimpleNamespace
import multiprocessing
import threading
from easyprocess import EasyProcess
import queue
# Oblivion phases
from OblivionSource.MacroExtractionPhase import MacroExtraction, MacroExtractionException
from OblivionSource.MacroInstrumentationPhase import MacroInstrumentation, MacroInstrumentationException
from OblivionSource.FileExecutionPhase import FileExecution, FileExecutionException, InteractionManager
from OblivionSource.PostProcessingPhase import PostProcessing, PostProcessingException


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

        self.exclusion_lines = None
        self.starting_time = -1
        self.last_phase_ended_at = -1
        self.current_attempts = 0

    def execute(self, single=False):
        if single:
            self.__execute_on_file()
        else:
            self.__execute_on_folder()

    def __execute_on_folder(self):
        tf = self.args.target_directory
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

        print("---")
        try:
            async_obj.get(timeout=self.args.time_limit)
        except multiprocessing.TimeoutError:
            print("[-] Timeout!")  # handle
        except OblivionCoreException as exc:
            print(exc)
        finally:
            del pool  # I have no clue why this is explicitly needed but ok

    def run(self):
        print(f"[?] Current sample: {os.path.basename(self.current_original_file)}")
        self.starting_time = time.time()
        self.last_phase_ended_at = self.starting_time
        self.current_attempts = 0
        while True:
            try:
                # Preliminary
                self.__clean_sandbox()
                # Phases
                self.__macro_extraction()
                self.__macro_instrumentation()
                # Dynamic Analysis calls file executions and needed daemons
                self.__dynamic_analysis()
                self.__post_processing()
            except MacroExtractionException as exc:
                raise OblivionCoreException(f"[-] Macro extraction failed: {exc}")  # handle
            except MacroInstrumentationException as exc:
                raise OblivionCoreException(f"[-] Macro instrumentation failed: {exc}")  # handle
            except FileExecutionException as exc:
                feasible = self.__can_it_run_again(exc)
                if not feasible:
                    raise OblivionCoreException(f"[-] File execution failed: {exc}")  # handle
                else:
                    self.__fix_instrumentation()
                    continue
            except PostProcessingException as exc:
                raise OblivionCoreException(f"[-] Report generation failed: {exc}")  # handle
            else:
                break  # handle
            finally:
                self.__clean_sandbox()  # handle
                print(f"[?] Analysis time: {self.__toc(total=True)}")

    def __interaction_manager(self):
        enable_event = threading.Event()
        enable_event.set()
        exception_queue = queue.Queue()
        phase = InteractionManager(self.current_original_file, self.ext_info[self.current_extension],
                                   enable_event, exception_queue)
        int_thread = threading.Thread(target=phase.run, daemon=True)
        int_thread.start()

        return int_thread, enable_event, exception_queue

    def __macro_extraction(self):
        phase = MacroExtraction(self.current_original_file, self.extensions, self.original_macro_data_path)
        macro_data = phase.run()

        with open(self.original_macro_path, "w") as foObj:
            for macro_name, macro_code in macro_data.items():
                foObj.write(f"{macro_name}\n{macro_code}\n\n")

        print(f"[+] Macro successfully extracted in {self.__toc()}")

    def __macro_instrumentation(self):
        phase = MacroInstrumentation(self.original_macro_data_path, self.exclusion_lines,
                                     self.instrumented_macro_data_path)
        phase.run()
        print(f"[+] Macro successfully instrumented in {self.__toc()}")

    def __dynamic_analysis(self):
        int_thread, enable_event, exception_queue = self.__interaction_manager()

        try:
            self.__file_extraction()
        except FileExecutionException as main_exc:
            try:
                child_exc = exception_queue.get(block=False)
            except queue.Empty:
                raise FileExecutionException(main_exc)
            else:
                raise FileExecutionException(child_exc)
        else:
            print(f"[+] File successfully executed in {self.__toc()}")
        finally:
            enable_event.clear()
            int_thread.join()

    def __file_extraction(self):
        log_path = os.path.join("OblivionResources", "logs", "sbx_out_err.log")
        log_path_sbx = self.__path_to_sandbox(log_path)

        phase = FileExecution(self.current_original_file, self.current_output_file, self.current_sandbox_output,
                              log_path, log_path_sbx, self.instrumented_macro_data_path, self.config.Sandbox_name,
                              self.config.Sandboxie_path, self.ext_info[self.current_extension],
                              self.args.no_clean_slate)
        phase.run()

    def __post_processing(self):
        program = self.ext_info[self.current_extension]["program"]
        phase = PostProcessing(self.current_original_file, self.current_output_file, self.original_macro_path,
                               self.current_extension, self.config.PowerDecode_path, self.config.Sandboxie_path,
                               self.config.Sandbox_name, program, self.current_report_file)
        phase.run()
        print(f"[+] Report successfully generated in {self.__toc()}")

    @staticmethod
    def __path_to_output(path):
        ext = path.split('.')[-1]
        modified_path = path.replace('.' + ext, f"_obl3_modified.{ext}")
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

    def __toc(self, total=False):
        toc = time.time()
        tic = self.last_phase_ended_at if not total else self.starting_time
        time_amount = around(toc - tic, decimals=3)
        self.last_phase_ended_at = toc
        return time_amount

    def __can_it_run_again(self, exc):
        self.current_attempts += 1

        if self.current_attempts >= self.args.max_retries:
            return False

        if "vba error" not in str(exc).lower():
            return False

        try:
            with open(self.current_output_file, "r") as foOutput:
                self.exclusion_lines = [line for line in foOutput.readlines()
                                        if line.startswith("Exception")
                                        and "=" in line]
                return len(self.exclusion_lines) > 0
                # If it starts with "Exception" it's an error, if it has no '='
                # it was not caused by vhook therefore we cannot intervene
        except FileNotFoundError:
            return False

    def __fix_instrumentation(self):
        raise NotImplementedError("Instrumentation fix will be completed in the future")
