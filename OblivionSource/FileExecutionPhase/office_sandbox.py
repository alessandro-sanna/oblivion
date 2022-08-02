import sys
import win32com.client
import os
import json
import pywintypes


class OfficeSandboxException(Exception):
    pass


class OfficeSandbox:
    def __init__(self, running_file, instrumented_code_path, log_file,
                 program, main_class, main_module,
                 auto_open, auto_close, no_clean_slate_flag, log_flag=True):
        self.running_file = running_file
        self.extension = running_file.split('.')[-1]
        self.instrumented_code_path = instrumented_code_path

        if log_flag:
            self.log_file = log_file
            log_fp = open(self.log_file, "w")
            sys.stdout = log_fp
            sys.stderr = log_fp

        self.program = program
        self.main_class = main_class
        self.main_module = main_module

        self.auto_open, self.auto_close = auto_open, auto_close
        self.no_clean_slate_flag = no_clean_slate_flag

        self.macro_extension_code_dict = {'bas': 1, 'cls': 2, 'frm': 3}
        self.macro_extension_code_dict_rev = {v: k for k, v in self.macro_extension_code_dict.items()}
        # To-be-set-variables
        self.macro_dict = None
        self.app_name = None
        self.vhook_module_path = None
        self.clean_file_path = None
        self.output_file_path = None
        # Preliminary phase
        self.__build_strings()
        self.__get_instrumented_macros()

    def __del__(self):
        sys.stdout.close()
        sys.stderr.close()

    def run(self):
        # Macro insertion
        app = self.__open_program()
        file_to_modify = self.__open_file(self.program, app, self.running_file, add_reference=True)
        self.__replace_macros(file_to_modify)
        file_to_modify.SaveAs(self.output_file_path)
        self.__close_file(file_to_modify)
        self.__close_program(app)
        # Proper execution
        app = self.__open_program(security_level=1)
        app.Visible = True
        file_to_run = self.__open_file(self.program, app, self.output_file_path, add_reference=False)
        # should i wait for auto_open == True?
        self.__close_file(file_to_run)
        # should i wait for auto_close == True?
        self.__close_program(app)

    def __build_strings(self):
        script_name = "class_" + self.program.lower() + ".vba"
        self.vhook_module_path = os.path.join("OblivionResources", "files", "vba_snip", script_name)
        # self.main_macro_name = "This" + self.main_module.capitalize() + ".cls"
        self.app_name = self.program.capitalize() + ".Application"
        clean_name = "clean." + self.extension
        self.clean_file_path = os.path.join("OblivionResources", "files", "clean_office", clean_name)
        base_folder = os.path.dirname(self.running_file)
        base_name = ''.join(os.path.basename(self.running_file).split(".")[:-1]) + \
                    "_output." + self.extension
        self.output_file_path = os.path.join(base_folder, base_name)

    def __get_instrumented_macros(self):
        with open(self.instrumented_code_path, "r") as icf:
            macro_dict = json.load(icf)

        with open(self.vhook_module_path, "r") as vhf:
            instrumentation_module = vhf.read()
            to_replace = "<insert path here from oblivion>"
            path_file = os.path.abspath(self.running_file)
            instrumentation_module = instrumentation_module.replace(to_replace, path_file)
            macro_dict.update({"vhook.bas": instrumentation_module})

        self.macro_dict = macro_dict

    def __open_program(self, security_level=3):
        app = win32com.client.DispatchEx(self.app_name)
        app.Application.AutomationSecurity = security_level
        app.Visible = False
        app.DisplayAlerts = False
        return app

    @staticmethod
    def __open_file(program, app, file_name, add_reference=True):
        if program == "word":
            return_file = app.Documents.Open(file_name)
        elif program == "excel":
            return_file = app.Workbooks.Open(file_name)
        else:
            raise OfficeSandboxException(f"Program {program} is not supported.")

        if add_reference:
            try:
                return_file.VBProject.References.AddFromGUID("{420B2830-E718-11CF-893D-00A0C9054228}", 1, 0)
            except pywintypes.com_error as exc:
                exc_code = exc.excepinfo[-2]
                if exc_code != 1032813:
                    raise OfficeSandboxException(f"scrrun.dll reference could not be added to file, error {exc_code}")
        return return_file

    def __replace_macros(self, office_file):
        if self.no_clean_slate_flag:
            office_file = self.__empty_macros(office_file)
        else:
            office_file = self.__make_macros(office_file)
        self.__write_macros(office_file)

    @staticmethod
    def __empty_macros(office_file):
        for macro in office_file.VBProject.VBComponents:
            macro.CodeModule.DeleteLines(1, macro.CodeModule.CountOfLines)
        return office_file

    def __make_macros(self, office_file):
        for name in self.macro_dict.keys():
            name, ext = name.split('.')
            try:
                macro_type = self.macro_extension_code_dict[ext]
                macro = office_file.VBProject.VBComponents.Add(macro_type)
                macro.Name = name
            except pywintypes.com_error:
                # You probably tried to recreate a Type 100 macro: fallback
                try:
                    office_file.VBProject.VBComponents.Remove(macro)
                except NameError:
                    raise OfficeSandboxException(f"Macro {name} creation failed.")
        return office_file

    def __write_macros(self, office_file):
        for macro in office_file.VBProject.VBComponents:
            name = macro.Name
            try:
                ext = self.macro_extension_code_dict_rev[macro.Type]
            except KeyError:
                ext = "cls"
            key = f"{name}.{ext}"
            new_code = self.macro_dict[key]
            macro.CodeModule.AddFromString(new_code)
        return office_file

    @staticmethod
    def __close_file(office_file):
        office_file.Close(0)

    @staticmethod
    def __close_program(app):
        app.Quit()


if __name__ == '__main__':
    running_file = sys.argv[1]
    instrumented_code_path = sys.argv[2]
    program, main_class, main_module = (x for x in sys.argv[3: 6])
    auto_open, auto_close, no_clean_slate_flag = (bool(int(x)) for x in sys.argv[6: 9])
    log_file = sys.argv[9]
    office_sbx_obj = OfficeSandbox(running_file, instrumented_code_path, log_file,
                                   program, main_class, main_module,
                                   auto_open, auto_close, no_clean_slate_flag)
    office_sbx_obj.run()
