import glob
import json
import os


class ReportProcessing:
    def __init__(self, report_folder, jsons_folder=None):
        self.report_path_list = glob.glob(os.path.join(report_folder, "**", "*"), recursive=True)
        self.jsons_folder = jsons_folder
        self.keys = [
            'Macro Oblivion Report',
            'Executable Files',
            'Other File Traces',
            'Domain Traces',
            'CreateObject Actions',
            'Shell Actions',
            'Cmd Actions',
            'Deobfuscated Powershell',
            'Environment Variables',
            'External Calls',
            'Exceptions',
            'System File Writes',
            'Auto Exec Methods',
            'Suspicious calls',
            'Interactions',
            'Variable Values',
            'Dynamic Call Graph'
        ]
        self.current_field = None

    def run(self):
        total = len(self.report_path_list)
        for index, report in enumerate(self.report_path_list):
            dict_data = self.__report_to_dict(report)
            dict_data = self.__process_data(dict_data)

            if self.jsons_folder is not None:
                self.__dict_to_json(os.path.basename(report), dict_data)

            print(f"\r{index + 1}/{total}", end='')
            yield dict_data

    def __process_data(self, dict_data):
        for field_name in dict_data.keys():
            processed_data = self.__manage(field_name)
            dict_data[field_name] = processed_data
        return dict_data

    def __dict_to_json(self, report_name, dict_data):
        output_file = os.path.join(self.jsons_folder, report_name.replace("txt", "json"))
        with open(output_file, "w") as foJson:
            json.dump(dict_data, foJson)

    @staticmethod
    def __report_to_dict(report_path):
        output_dict = dict()
        with open(report_path, "r") as foReport:
            report_text = foReport.read()
            tmp_list = report_text.split("###")[1:]
            for index in range(0, len(tmp_list), 2):
                k = tmp_list[index].strip()
                v = tmp_list[index + 1].strip().splitlines()
                output_dict.update({k: v})
        return output_dict

    def __manage(self, field_name):
        self.current_field = field_name
        if self.__is_empty():
            return None

        if field_name == 'Macro Oblivion Report':
            return self.__manage_header()
        if field_name == 'Executable Files':
            return self.__manage_executable_files()
        if field_name == 'Other File Traces':
            return self.__manage_other_file_traces()
        if field_name == 'Domain Traces':
            return self.__manage_domain_traces()
        if field_name == 'CreateObject Actions':
            return self.__manage_createobject_actions()
        if field_name == 'Shell Actions':
            return self.__manage_shell_actions()
        if field_name == 'Cmd Actions':
            return self.__manage_cmd_actions()
        if field_name == 'Deobfuscated Powershell':
            return self.__manage_deobfuscated_powershell()
        if field_name == 'Environment Variables':
            return self.__manage_environment_variables()
        if field_name == 'External Calls':
            return self.__manage_external_calls()
        if field_name == 'Exceptions':
            return self.__manage_exceptions()
        if field_name == 'System File Writes':
            return self.__manage_system_file_writes()
        if field_name == 'Auto Exec Methods':
            return self.__manage_auto_exec_methods()
        if field_name == 'Suspicious calls':
            return self.__manage_suspicious_calls()
        if field_name == 'Interactions':
            return self.__manage_interactions()
        if field_name == 'Variable Values':
            return self.__manage_variable_values()
        if field_name == 'Dynamic Call Graph':
            return self.__manage_dynamic_call_graph()

        return self.__manage_field()

    def __is_empty(self):
        return self.current_field[0].lower().startswith("nothing found")

    def __manage_header(self):
        self.__manage_field()

    def __manage_executable_files(self):
        self.__manage_field()

    def __manage_other_file_traces(self):
        self.__manage_field()

    def __manage_domain_traces(self):
        self.__manage_field()

    def __manage_createobject_actions(self):
        self.__manage_field()

    def __manage_shell_actions(self):
        self.__manage_field()

    def __manage_cmd_actions(self):
        self.__manage_field()

    def __manage_deobfuscated_powershell(self):
        self.__manage_field()

    def __manage_environment_variables(self):
        self.__manage_field()

    def __manage_external_calls(self):
        self.__manage_field()

    def __manage_exceptions(self):
        self.__manage_field()

    def __manage_system_file_writes(self):
        self.__manage_field()

    def __manage_auto_exec_methods(self):
        self.__manage_field()

    def __manage_suspicious_calls(self):
        self.__manage_field()

    def __manage_interactions(self):
        self.__manage_field()

    def __manage_variable_values(self):
        self.__manage_field()

    def __manage_dynamic_call_graph(self):
        self.__manage_field()

    def __manage_field(self):
        return False
