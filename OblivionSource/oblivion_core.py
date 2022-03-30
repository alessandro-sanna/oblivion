import os
import json
from types import SimpleNamespace
import multiprocessing
# Oblivion phases
from OblivionSource.MacroExtractionPhase import MacroExtraction, MacroExtractionException
from OblivionSource.MacroInstrumentationPhase import MacroInstrumentation, MacroInstrumentationException
from OblivionSource.FileExecutionPhase import FileExecution, FileExecutionException
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
        async_obj = pool.apply_async(self.__run)

        self.current_original_file = self.args.target_file
        self.current_sandbox_original = self.__path_to_sandbox(self.args.target_file)

        self.current_modified_file, \
            self.current_output_file, self.current_report_file = self.__path_to_output(self.args.target_file)

        self.current_sandbox_modified,\
            self.current_sandbox_output, self.current_sandbox_report = \
            (self.__path_to_sandbox(x) for x in self.__path_to_output(self.args.target_file))

        try:
            async_obj.get(timeout=self.args.time_limit)
        except OblivionCoreException:
            pass  # handle
        except multiprocessing.TimeoutError:
            pass  # handle

    def __run(self):
        try:
            self.__macro_extraction()
            self.__macro_instrumentation()
            self.__file_extraction()
            self.__post_processing()
        except MacroExtractionException as exc:
            pass  # handle
        except MacroInstrumentationException as exc:
            pass  # handle
        except FileExecutionException as exc:
            pass  # handle
        except PostProcessingException as exc:
            pass  # handle
        else:
            pass  # handle
        finally:
            pass  # handle

    def __macro_extraction(self):
        phase = MacroExtraction(self.current_original_file, self.extensions.keys())
        macro_code = phase.run()
        print("[+] Macro successfully extracted")

        with open(self.original_macro_path, "w") as omf:
            omf.write(macro_code)

    def __macro_instrumentation(self):
        pass

    def __file_extraction(self):
        log_path = os.path.join("OblivionResources", "logs", "sbx_out_err.log")
        log_path_sbx = self.__path_to_sandbox(log_path)
        phase = FileExecution()

    def __post_processing(self):
        pass

    @staticmethod
    def __path_to_output(path):
        ext = path.split('.')[-1]
        modified_path = path.replace(ext, f"_obl3_modified.{ext}")
        output_path = path.replace(ext, f"_{ext}_output.txt")
        report_path = path.replace(ext, f"_{ext}_report.txt")
        return modified_path, output_path, report_path

    @staticmethod
    def __path_to_sandbox(path):
        # implement, should be a replace
        return path

