import argparse
import shutil
import pathlib
import sys
from os.path import join, dirname
from pywinauto.findwindows import find_elements
from easyprocess import EasyProcess
from oletools import olevba

import pywinauto
import psutil
import time
import os
import re
import json
from types import SimpleNamespace


class PreProcessingException(Exception):
    pass


class PreProcessing:
    def __init__(self, args_namespace, config_namespace, ext_info):
        # self.args_dict = args_dict
        # self.execution_type = execution_type

        # self.folder_path = ""  # folder in which the office files are
        # folder in which the preprocessed files have to be saved
        # self.output_folder_path = ""
        self.folder_path = args_namespace.folder_path
        self.output_folder_path = args_namespace.output_path
        self.execution_type = args_namespace.mode
        self.copy_flag = args_namespace.copy
        self.enumerate_flag = args_namespace.enumerate
        self.ext_info = ext_info
        self.interaction_path = ""
        self.corrupted_path = ""
        self.pwd_protected_path = ""
        self.empty_path = ""
        self.file_name = ""
        self.file_ext = ""
        self.executables_folder_path = ""
        self.not_executables_folder_path = ""
        self.no_execution_folder_path = ""
        self.pwd_protected_folder_path = ""
        self.sandboxie_path = config_namespace.Sandboxie_path
        self.sandbox_name = config_namespace.Sandbox_name
        self.is_err = 0
        # self.__parse_args()

    def execute(self):
        if self.execution_type == "Static":
            self.__static_analysis()
        elif self.execution_type == "Dynamic":
            self.__dynamic_analysis()

    def __static_analysis(self):
        file_list = os.listdir(self.folder_path)
        total = len(file_list)
        corr_ind = 0
        for index, file_str in enumerate(file_list):
            self.file_path = join(self.folder_path, file_str)
            if self.__is_office_file(self.file_path):
                self.file_name = file_str
                self.__set_file_ext()
                analysis_out = self.__olevba_analysis(self.file_path)
                if analysis_out == "Password protected":
                    self.__manage_pwd_protected()
                elif analysis_out == "Error":
                    self.__manage_corrupted()
                else:  # if the olevba output is correctly generated
                    if self.__is_file_with_interaction(analysis_out):
                        self.__manage_interaction()
                    else:
                        macro_list = self.__get_macros(analysis_out)
                        if len(macro_list) == 0:
                            self.__manage_empty()
                        elif len(macro_list) > 0:
                            error = \
                                self.__check_errors_in_macro_name(macro_list)
                            if error is True:  # Errors in macro name
                                self.__manage_corrupted()
                                corr_ind += 1
                            else:
                                self.__manage_correct(macro_list)
            print(f"\r{index + 1}/{total}, {corr_ind} corrupted", end='')

    def __dynamic_analysis(self):
        self.__manage_folder_creation()
        self.is_err = False
        file_list = [f for f in os.listdir(self.folder_path) if
                     self.__is_office_file(join(self.folder_path, f)) and
                     '$' not in f]
        try:
            total = len(file_list)
            for i_file, file_str in enumerate(file_list):
                print(f"\r{i_file}/{total}", end="")

                self.file_name = file_str
                self.file_path = join(self.folder_path, self.file_name)
                self.__set_file_ext()
                pid = self.__open_doc()
                if pid is None:
                    self.__manage_correct_case()
                    continue
                app = pywinauto.Application(backend="win32")
                app.connect(process=pid)
                start_time = time.time()
                while 1:
                    if time.time() - start_time > 60:
                        if self.__is_opening_window():
                            self.__manage_opening_window()
                            break
                        else:  # file already opened but not executing
                            self.__clean_sandbox()
                            self.__move_file(self.file_path,
                                             join(self.not_executables_folder_path,
                                                  self.file_name))
                            break
                    else:
                        win_list = self.__get_main_window()
                        if len(win_list) >= 1:  # found an office main window
                            time.sleep(3)
                            dll_err_win = self.__is_dll_error()
                            if dll_err_win is not None:
                                self.__manage_dll_error(dll_err_win, app)
                            elif self.__check_error() is True or \
                                    self.is_err is True:
                                self.__manage_error_case()
                                break
                            elif self.__check_pwd():
                                self.__manage_pwd_case()
                                break
                            else:
                                self.__close_office(app, win_list[0].handle)
                                start_inn_time = time.time()
                                while 1:
                                    if time.time() - start_inn_time > 10:
                                        self.__close_office(app,
                                                            win_list[0].handle)
                                    if time.time() - start_inn_time > 20:
                                        try:
                                            self.__kill_office_by_pid(pid)
                                        except psutil.NoSuchProcess:
                                            self.__manage_error_case()
                                            break
                                    w_list = self.__get_main_window()
                                    if len(w_list) == 0:
                                        self.__manage_correct_case()
                                        break
                                    else:
                                        if self.__check_error():
                                            self.__manage_error_case()
                                            break
                                break
                        else:  # To get the error on last opening or WinSock error
                            flag = False
                            while flag is False:
                                err_win = find_elements(class_name=u"#32770")
                                if len(err_win) > 0 and len(win_list) == 0:
                                    if self.__is_outlook_error(err_win):
                                        self.__manage_error_case()
                                        flag = True
                                        break
                                    else:
                                        self.__manage_error_on_last_opening(
                                            app, err_win)
                                else:
                                    break
                            if flag is True:
                                break
        except Exception as ex:
            raise PreProcessingException(ex)

    def __set_file_ext(self):
        self.file_ext = self.file_name.split('.')[-1].lower()
        program_name = self.ext_info[self.file_ext]["program"].capitalize()
        self.interaction_path = join(self.output_folder_path, program_name, "Interaction")
        self.corrupted_path = join(self.output_folder_path, program_name, "Corrupted")
        self.pwd_protected_path = join(self.output_folder_path, program_name, "Password_Protected")
        self.analyzable_path = join(self.output_folder_path, program_name, "Analyzable")
        self.empty_path = join(self.analyzable_path, "Empty")

    def __is_opening_window(self):
        if len(find_elements(class_name=u"Sandbox:DefaultBox:MsoSplash")) > 0:
            return True
        else:
            return False

    def __manage_opening_window(self):
        self.__clean_sandbox()
        self.__move_file(self.file_path, join(self.no_execution_folder_path,
                                              self.file_name))

    def __manage_error_case(self):
        self.__clean_sandbox()
        self.__move_file(self.file_path, join(self.not_executables_folder_path,
                                              self.file_name))
        self.is_err = False

    def __manage_pwd_case(self):
        self.__clean_sandbox()
        self.__move_file(self.file_path, join(self.pwd_protected_folder_path,
                                              self.file_name))

    def __manage_correct_case(self):
        self.__clean_sandbox()
        self.__move_file(self.file_path, join(self.executables_folder_path,
                                              self.file_name))

    def __manage_folder_creation(self):
        self.executables_folder_path = join(self.output_folder_path,
                                            "Executables")
        self.__create_dir(self.executables_folder_path)
        self.not_executables_folder_path = join(self.output_folder_path,
                                                "No_Executables")
        self.__create_dir(self.not_executables_folder_path)
        self.no_execution_folder_path = join(self.output_folder_path,
                                             "No_Execution")
        self.__create_dir(self.no_execution_folder_path)
        self.pwd_protected_folder_path = join(self.output_folder_path,
                                              "Password_Protected")
        self.__create_dir(self.pwd_protected_folder_path)

    def __manage_correct(self, macro_list):
        macro_type_list = self.__get_macros_type(macro_list)
        macro_folder_name = '-'.join(macro_type_list)
        macro_folder_path = join(self.analyzable_path, macro_folder_name)
        self.__create_dir(macro_folder_path)
        self.__move_file(self.file_path, join(macro_folder_path, self.file_name))

    def __manage_wrong(self, wrong_path):
        self.__create_dir(wrong_path)
        self.__move_file(self.file_path, join(wrong_path, self.file_name))

    def __manage_empty(self):
        self.__manage_wrong(self.empty_path)

    def __manage_pwd_protected(self):
        self.__manage_wrong(self.pwd_protected_path)

    def __manage_interaction(self):
        self.__manage_wrong(self.interaction_path)

    def __manage_corrupted(self):
        self.__manage_wrong(self.corrupted_path)

    def __manage_dll_error(self, err_win, app):
        self.__click_coords(app, err_win.handle, (430, 140))

    """
    def __parse_args(self):
        self.__set_folder_path()
        self.__set_output_folder_path()
        

    def __set_folder_path(self):
        if "-ps" in self.args_dict:
            if isdir(self.args_dict["-ps"][0]):
                self.folder_path = self.args_dict["-ps"][0]
            else:
                raise PreProcessingException(u"the given directory path is "
                                             u"not a folder or doesn't exist")
        elif "-pd" in self.args_dict:
            if isdir(self.args_dict["-pd"][0]):
                self.folder_path = self.args_dict["-pd"][0]
            else:
                raise PreProcessingException(u"the given directory path is "
                                             u"not a folder or doesn't exist")
        else:
            raise PreProcessingException(u"no directory detected in the "
                                         u"arguments")

    def __set_output_folder_path(self):
        if "-ps" in self.args_dict:
            if isdir(self.args_dict["-ps"][1]):
                self.output_folder_path = self.args_dict["-ps"][1]
            else:
                raise PreProcessingException(u"the given directory path is "
                                             u"not a folder or doesn't exist")
        elif "-pd" in self.args_dict:
            if isdir(self.args_dict["-pd"][1]):
                self.output_folder_path = self.args_dict["-pd"][1]
            else:
                raise PreProcessingException(u"the given directory path is "
                                             u"not a folder or doesn't exist")
        else:
            raise PreProcessingException(u"no output path detected in the "
                                         u"arguments")
    """

    def __check_errors_in_macro_name(self, macros_list):
        if False in map(self.__extension_error, macros_list):
            return True
        else:
            return False

    def __get_macros_type(self, macros_list):
        return_list = list(set(map(self.__cleaner, macros_list)))
        return_list.sort()
        return return_list

    def __open_doc(self):
        try:
            cmd = [self.sandboxie_path, self.file_path]
            p = EasyProcess(cmd).call().wait()
            start_time = time.time()
            while 1:
                if time.time() - start_time < 15:
                    for pr in psutil.process_iter():
                        # program_name =
                        process_name = self.ext_info[self.file_ext]["process_name"]
                        if process_name in pr.name():
                            return pr.pid

                        code_to_delete_maybe = """
                        if self.file_ext == "doc" or self.file_ext == "docm":
                            if pr.name() == "WINWORD.EXE":
                                pid = pr.pid
                                return pid
                        elif self.file_ext == "xls" or self.file_ext == "xlsm":
                            if pr.name() == "EXCEL.EXE":
                                pid = pr.pid
                                return pid
                        """

                else:
                    if p.return_code == 0 and len(p.stderr.strip()) == 0:
                        return None
                    raise PreProcessingException("File not found: {}"
                                                 "".format(self.file_path))
        except WindowsError as ex:
            return
        except psutil.NoSuchProcess as ex:
            return

    def __clean_sandbox(self):
        EasyProcess([self.sandboxie_path, "/terminate_all"]).call().wait()
        EasyProcess([self.sandboxie_path,
                     "delete_sandbox_silent"]).call().wait()

    def __get_main_window(self):
        return self.__get_win_list(self.sandbox_name, self.ext_info[self.file_ext]["main_class"])

    def __check_error(self):
        window = find_elements(class_name=u"#32770")
        if len(window) == 1:
            return True
        window = find_elements(
            title=u"Microsoft Visual Basic, Applications"
                  u" Edition")
        if len(window) == 1:
            return True
        window = find_elements(title=u"Microsoft Visual Basic")
        if len(window) == 1:
            return True
        window = find_elements(title=u"[#] Microsoft Visual Basic [#]")
        if len(window) == 1:
            return True
        window = find_elements(class_name=u"Sandbox:DefaultBox:NUIDialog")
        if len(window) == 1:
            return True
        window = find_elements(title_re=u".*Microsoft Visual Basic.*")
        if len(window) == 1:
            return True
        window = find_elements(
            class_name=u"Sandbox:DefaultBox:bosa_sdm_msword")
        if len(window) == 1:
            return True
        return False

    def __check_pwd(self):
        win_list = self.__get_win_list(self.sandbox_name, self.ext_info[self.file_ext]["bosa_class"])
        return len(win_list) == 1 and u"Password" in win_list[0].name

    def __is_outlook_error(self, err_win):
        win = err_win[0]
        if win.name == u"Microsoft Outlook":
            for child in win.iter_children():
                if u"Static" in child.class_name and \
                        u"The profile name is not valid. Enter a " \
                        u"valid profile name." in child.rich_text:
                    return True
        return False

    def __manage_error_on_last_opening(self, app, err_win):
        if err_win[0].name == u"Microsoft Excel" or \
                self.__win_exists(app, err_win[0].handle, "Static",
                                  u".*stato avviato correttamente l'ultima "
                                  u"volta.*"):
            self.__click_coords(app, err_win[0].handle, (450, 180))
            self.__click_coords(app, err_win[0].handle, (470, 15))
        elif err_win[0].name == u"[#] Microsoft Word [#]":
            self.__click_coords(app, err_win[0].handle, (562, 15))
            self.__click_coords(app, err_win[0].handle, (383, 15))
        elif err_win[0].name == u"Sandboxie RpcSs" or \
                self.__win_exists(app, err_win[0].handle, u"Static",
                                  u".*Could not initialize WinSock.*"):
            self.__click_coords(app, err_win[0].handle, (220, 135))
        elif err_win[0].name == u'[#] Internet Explorer 11 [#]':
            self.__click_coords(app, err_win[0].handle, (420, 345))
        elif err_win[0].name == u'Sandboxie' or \
                self.__win_exists(app, err_win[0].handle, u"Static",
                                  u".*'Gestione licenze di Sandboxie.*"):
            self.__click_coords(app, err_win[0].handle, (500, 350))
        elif (err_win[0].name is not None and
              u'[#] C:\\Users\\Oblivion\\AppData\\'
              u'Local\\Temp' in err_win[0].name) or \
                self.__win_exists(app, err_win[0].handle,
                                  u"Sandbox:DefaultBox:DirectUIHWND"):
            self.__click_coords(app, err_win[0].handle, (520, 110))
            self.is_err = True

    @staticmethod
    def __is_dll_error():
        err_win = find_elements(class_name=u"#32770")
        if len(err_win) > 0 and err_win[0].name == u"Windows - Inizializzaz" \
                                                   u"ione DLL non riuscita":
            return err_win[0]
        else:
            return None

    @staticmethod
    def __click_coords(app, handle, coords):
        err_win = app.window(handle=handle).wrapper_object()
        try:
            err_win.click_input(coords=coords)
        except WindowsError:
            return

    @staticmethod
    def __close_office(app, handle):
        # can this be replaced with OfficeManager.close_office?
        try:
            rect = app.window(handle=handle).wrapper_object().client_rect()
            app.window(handle=handle).wrapper_object().click_input(
                coords=(rect.right - 10, 15))
        except pywinauto.controls.hwndwrapper.InvalidWindowHandle:
            return
        except RuntimeError:
            return

    @staticmethod
    def __olevba_analysis(f_path):
        olevba_path = join(dirname(olevba.__file__), "olevba.py")
        cmd = [sys.executable, olevba_path, f_path]
        p = EasyProcess(cmd).call(timeout=120).wait()
        output = p.stdout
        error = p.stderr
        if len(error) != 0 and \
                not error.startswith(u'pydev debugger: process'):
            if "ERROR    Problems with encryption in main" in error:
                return "Password protected"
            else:
                return "Error"
        else:
            return output

    @staticmethod
    def __create_dir(dir_path):
        if os.path.exists(dir_path):
            return
        else:
            os.makedirs(dir_path)

    def __move_file(self, old_path, new_path):
        try:
            if self.enumerate_flag:
                fp = os.path.join(os.path.dirname(new_path), "amount.txt")
                pathlib.Path(fp).touch(exist_ok=True)
                with open(fp, "r+") as frwCount:
                    amount = frwCount.read()
                    if amount:
                        amount = int(amount)
                    else:
                        amount = 0
                    frwCount.seek(0)
                    frwCount.write(f"{amount + 1}")
            else:
                if self.copy_flag:
                    shutil.copy2(old_path, new_path)
                else:
                    os.rename(old_path, new_path)
        except WindowsError:
            pass

    @staticmethod
    def __is_file_with_interaction(olevba_output):
        olevba_output = olevba_output.lower()
        if ".showwindow" in olevba_output or "msgbox" in olevba_output:
            return True
        else:
            return False

    @staticmethod
    def __get_macros(olevba_output):
        if re.findall(r"^No VBA macros found.", olevba_output, re.MULTILINE):
            return []
        else:
            return re.findall(r"(?<=VBA MACRO )(.*?)(?= )", olevba_output,
                              re.MULTILINE)

    @staticmethod
    def __extension_error(string):
        if '.' in string:
            return True
        else:
            return False

    @staticmethod
    def __cleaner(string):
        return string.split('.')[-1]

    @staticmethod
    def __kill_office_by_pid(pid):
        psutil.Process(pid).terminate()

    @staticmethod
    def __win_exists(app, window_handle, class_name, title=None):
        if title is not None:
            try:
                app.window(handle=window_handle). \
                    child_window(class_name=class_name, title_re=title). \
                    wrapper_object()
                return True
            except:
                return False
        else:
            try:
                app.window(handle=window_handle). \
                    child_window(class_name=class_name).wrapper_object()
                return True
            except:
                return False

    @staticmethod
    def __get_win_list(sandbox, x_class):
        class_name = "Sandbox:" + sandbox + ":" + x_class
        win_list = pywinauto.findwindows.find_elements(class_name=class_name)
        return win_list

    def __is_office_file(self, file_name):
        return file_name.split('.')[-1].lower() in self.ext_info.keys()

    def __path_to_sandbox(self, path):
        old_root = os.path.join("C:\\", "Users", os.getenv("username"))
        new_root = os.path.join("C:\\", "Sandbox", os.getenv("username"), self.sandbox_name, "user", "current")
        sandbox_path = path.replace(old_root, new_root)
        return sandbox_path

class DataGet:
    def __init__(self, config_path, ext_info_path):
        self.config_path = config_path
        self.ext_info_path = ext_info_path

    def get_args(self):
        args = self.parse_args()
        self.validate_args(args)
        return args

    @staticmethod
    def parse_args():
        parser = argparse.ArgumentParser(prog="preprocess.py", prefix_chars="-")
        parser.add_argument("-d", "--folder_path", nargs='?', type=str, required=True,
                            help="choose directory to analyze")
        parser.add_argument("-o", "--output_path", nargs='?', type=str, required=True,
                            help="choose where to save pre-processing results")
        parser.add_argument("-m", "--mode", nargs='?', required=True,
                            choices={"Static", "Dynamic"},
                            help="choose pre-processing mode")

        parser.add_argument("-c", "--copy", required=False,
                            action="store_true",
                            help="choose if copying or moving")
        parser.add_argument("-e", "--enumerate", required=False,
                            action="store_true",
                            help="overwrite standard behaviour, instead just count the files")
        return parser.parse_args(sys.argv[1:])

    @staticmethod
    def validate_args(args):
        if not os.path.isdir(args.folder_path) or not os.path.isdir(args.output_path):
            raise argparse.ArgumentError(None, "Invalid arguments! -d and -o must be valid directories")

    def get_config(self):
        with open(self.config_path, "r") as cj:
            return json.load(cj, object_hook=lambda d: SimpleNamespace(**d))

    def get_ext_info(self):
        with open(self.ext_info_path, "r") as cj:
            return json.load(cj)


if __name__ == '__main__':
    os.chdir(os.path.dirname(os.path.dirname(os.getcwd())))
    config_path = os.path.join("OblivionResources", "config", "configuration.json")
    ext_info_path = os.path.join("OblivionResources", "config", "extensions.json")
    get_obj = DataGet(config_path, ext_info_path)

    args = get_obj.get_args()
    config = get_obj.get_config()
    ext_info = get_obj.get_ext_info()

    prep_obj = PreProcessing(args, config, ext_info)
    prep_obj.execute()
