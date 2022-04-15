# Macro-Oblivion
# Created by Alessandro Medda and Davide Maiorca
# University of Cagliari, Department of Electrical and Electronic Engineering
# V. 0.1 Beta - December 2017
# This is an extended version of the VHook Tool released by ESET

from os.path import basename, join, exists, splitdrive
from oletools.olevba import VBA_Parser
from easyprocess import EasyProcess
from collections import OrderedDict
from psutil import process_iter
from urllib.parse import urlparse
import queue as Queue
import time
import sys
import re
import os
import numpy as np


class PostProcessingException(Exception):
    pass


class PostProcessing(object):
    def __init__(self, file_path, output_path, original_macro_path, file_ext,
                 powerdecode_path, sandboxie_path, sandboxie_name, program):
        self.variables = OrderedDict()  # variables
        self.nodes = OrderedDict()  # method calls
        self.exe_set = set()  # executables set
        self.url_set = set()  # url set
        self.create_objs = set()  # CreateObject arguments
        self.other_files = set()  # Other files found
        self.shell_commands = set()  # Shell arguments
        self.cmd_commands = set()  # Cmd commands
        self.exceptions = set()  # Exception set
        self.environs = set()  # Get Environment variables
        self.ext_calls = set()  # Get external calls
        self.file_writes = set()  # Get external file writes
        self.susp_calls_dict = {}  # Get suspicius calls from macro
        self.exec_calls_dict = {}  # get auto exec function from macro
        self.interaction_lines = list()  # get interaction info from macro
        self.ioc_dict = {}  # get executable file names from macro
        self.file_ext = file_ext
        self.file_path = file_path
        self.program = program
        self.original_macro_path = original_macro_path
        self.macro_text = self.__get_macro(original_macro_path)
        self.output_name = basename(output_path)
        self.complete_path = output_path
        self.__deobf_list = list()
        self.__sandboxie_path = sandboxie_path
        self.__sandboxie_name = sandboxie_name
        self.__powerdecode_path = powerdecode_path
        self.__temp_ps_path = join(os.getcwd(), "data", "temp_powershell.ps1")
        self.__out_path = join(os.getcwd(), "data", "out.txt")
        self.__sand_out_path = self.__get_sandbox_path(self.__out_path)
        self.__err_path = join(os.getcwd(), "data", "err.txt")
        self.__sand_err_path = self.__get_sandbox_path(self.__err_path)
        self.__get_calls_from_macro()
        self.__parse_variables()
        self.__parse_methods()
        self.__get_executables()
        self.__get_other_files()
        self.__get_domains()
        self.__get_create_obj()
        self.__get_shell()
        self.__get_cmd()
        self.__get_environs()
        self.__get_exception()
        self.__get_ext_calls()
        self.__get_sys_file_write()
        self.__get_deobfuscation()
        self.__get_interactions()

    def save_report(self, output_path):
        self.out = open(output_path, "w")  # destination
        self.out.write("### Macro Oblivion Report ###\n\n\n")
        self.__print_info()
        self.__print_executables()
        self.__print_other_files()
        self.__print_urls()
        self.__print_create_objects()
        self.__print_shell()
        self.__print_cmd()
        self.__print_deobfuscation()
        self.__print_environs()
        self.__print_ext_calls()
        self.__print_exception()
        self.__print_sys_file_write()
        self.__print_macro()
        self.__print_interactions()
        self.__print_variables()
        self.__print_methods()
        self.out.close()

    def is_powershell_present(self):
        if len(self.shell_commands) > 0:
            return True
        elif len(self.cmd_commands) > 0:
            return True
        else:
            return False

    def __get_deobfuscation(self):
        if self.__powerdecode_path is None:
            return ""
        else:
            ps_list = self.__get_powershell_list()
            for ps_code in ps_list:
                ps_code = self.__clean_powershell(ps_code)
                self.__write_file(self.__temp_ps_path, ps_code)
                plainscript = self.__deobfuscate_ps()
                if plainscript != "":
                    if plainscript.strip() != ps_code.strip():
                        self.__deobf_list.append(plainscript)
                    else:
                        self.__deobf_list.append("PowerShell script already "
                                                 "clear")
                else:
                    self.__deobf_list.append("Impossible to deobfuscate the "
                                             "PowerShell code")
                self.__clean_sandbox()
                os.remove(self.__temp_ps_path)
                self.__deobf_list = list(self.__filter_commands(self.__deobf_list))

    def __clean_sandbox(self):
        EasyProcess([self.__sandboxie_path, "/terminate_all"]).call().wait()
        EasyProcess([self.__sandboxie_path,
                     "delete_sandbox_silent"]).call().wait()

    def __get_powershell_list(self):
        ps_list = list()
        for shell in self.shell_commands:
            if shell not in ps_list and ("powershell " in shell.lower() or
                                         "powershell.exe " in shell.lower()):
                ps_list.append(shell)
        for cmd in self.cmd_commands:
            if cmd not in ps_list and ("powershell " in cmd.lower() or
                                       "powershell.exe " in cmd.lower()):
                ps_list.append(cmd)
        return ps_list

    def __deobfuscate_ps(self):
        powerdecode_name = \
            self.__powerdecode_path.split("\\")[-1].split(".")[0]
        cmd = "\"" + self.__sandboxie_path + "\"" + " /hide_window " \
              "powershell.exe -ExecutionPolicy Bypass -WindowStyle Hidden " \
              "-Command \"& {Import-Module -Name '" + self.__powerdecode_path \
              + "'; & " + powerdecode_name + " -InputFile '" + \
              self.__temp_ps_path + "' -OutputFileName " + "'" + \
              self.__out_path + "' 2> '" + self.__err_path + "'}\""
        p = EasyProcess(cmd).call().wait()
        while True:
            if not self.__is_powershell_running():
                break
        if p.return_code == 0:
            if exists(self.__sand_out_path):
                if not self.__is_error_found():
                    out_text = self.__read_file(self.__sand_out_path)
                    plainscript = self.__get_plainscript(out_text)
                    if plainscript != "":
                        return plainscript
                    else:
                        return ""
                else:
                    return ""
            else:
                raise PostProcessingException("PowerDecode did not generate any output, report failed")# exit(-1)
        else:
            exit(-1)

    def __clean_powershell(self, ps_str):
        ps_str = ps_str.replace("  ", " ").strip()
        pre_ps_str = ps_str[:ps_str.lower().index("powershell")]
        if '"' in pre_ps_str and ps_str.endswith('"'):
            ps_str = ps_str[ps_str.find('"'): ps_str.rfind('"')]
        if ps_str == "powershell.exe":
            return ""
        if "powershell . " in ps_str.lower():
            ps_str = self.__clean_substr(ps_str, "powershell . ")
        elif "powershell.exe " in ps_str.lower():
            ps_str = self.__clean_substr(ps_str, "powershell.exe ")
        elif "powershell " in ps_str.lower():
            ps_str = self.__clean_substr(ps_str, "powershell ")
        if " -command " in ps_str.lower():
            return self.__clean_substr(ps_str, " -command ")
        elif " -c " in ps_str.lower():
            return self.__clean_substr(ps_str, " -c ")
        elif " if (" in ps_str.lower():
            return ps_str[ps_str.lower().index(" if ("):].strip()
        else:
            ps_str = self.__remove_ps_options(ps_str)
        return ps_str.strip()

    def __remove_ps_options(self, ps_str):
        flag_list = ["psconsolefile", "executionpolicy", "windowstyle",
                     "version", "inputformat", "outputformat",
                     "encodedcommand", "configurationname", "file",
                     "exec", "noexit", "noprofile", "nologo", "sta", "mta",
                     "noninteractive", "nop", "ec", "w"]
        par_list = ["bypass", "allsigned", "default", "2.0", "3.0",
                    "normal", "minimized", "maximized", "hidden"]
        opt_list = ps_str.split(" -")
        for el in opt_list:
            el = el.replace("-", "")
            if any(el.lower().startswith(flag) for flag in flag_list):
                el_list = el.split(" ")
                el_list = self.__clean_list(el_list)
                if len(el_list) == 1 and el_list[0].lower() in flag_list:
                    idx = ps_str.index(el_list[0])
                    ps_str = ps_str[: idx - 1] + ps_str[idx + len(el_list[0]):]
                elif len(el_list) > 1:
                    if el_list[0].lower() in flag_list and \
                            el_list[1].lower() in par_list:
                        idx = ps_str.index(el_list[0])
                        ps_str = \
                            ps_str[: idx - 1] + ps_str[idx + len(el_list[0]):]
                        idx = ps_str.index(el_list[1])
                        ps_str = ps_str[: idx] + ps_str[idx + len(el_list[1]):]
                    elif el_list[0].lower() in flag_list and \
                            el_list[1].lower() not in par_list:
                        idx = ps_str.index(el_list[0])
                        ps_str = \
                            ps_str[: idx - 1] + ps_str[idx + len(el_list[0]):]
        return ps_str

    def __get_sandbox_path(self, path):
        out_drive = splitdrive(path)[0]
        cwd_drive = splitdrive(os.getcwd())[0]
        username = os.getenv("username").replace(" ", "_")
        if out_drive == cwd_drive:
            file_path = path.replace(join("C:\\", "Users",
                                          os.getenv("username")), "")
            sandbox_path = \
                join("C:\\", "Sandbox", username, self.__sandboxie_name,
                     "user", "current") + file_path
            return sandbox_path
        else:
            relative_path = splitdrive(path)[1]
            sandbox_path = join("C:\\", "Sandbox", username,
                                self.__sandboxie_name, "drive",
                                out_drive.replace(":", "")) + relative_path
            return sandbox_path

    def __is_error_found(self):
        txt = self.__read_file(self.__sand_err_path)
        if txt == "":
            return False
        else:
            return True

    def __parse_variables(self):
        """This method traces the evolution of each variable of the macro."""
        sys_functions = ['mid', "createobject", "left", "right", "environ",
                         "shell"]
        total_data = []
        queue = Queue.LifoQueue()
        with open(self.complete_path) as target_file:
            lines = target_file.readlines()
            for i in range(0, len(lines)):
                line = lines[i]
                line = line.replace("\\Oblivion\\", "\\USER\\")
                try:
                    if line.split()[0].replace(" ", "").lower() \
                            in sys_functions:
                        continue
                except:
                    pass
                if "**I" in line or "**R" in line or "Exception:" in line:
                    continue
                if ":" in line:
                    try:
                        if "**" in lines[i+1]:
                            continue
                        else:
                            pass
                    except:
                        pass
                if self.__is_assignment(line) is True:
                    if line.split("=")[0].replace(" ","").lower() \
                            in sys_functions:
                        continue
                    else:
                        queue.put(line.split("=")[0])
                        continue

                if queue.empty() is False:
                    total_data.append(queue.get())
                    total_data.append(line)

                if len(total_data) < 2:
                    continue
                else:
                    if total_data[1] == "Vero\n":
                        total_data[1] = "True\n"
                    if total_data[1] == "Falso\n":
                        total_data[1] = "False\n"
                    try:
                        self.variables[total_data[0]].append(total_data[1])
                        total_data = []
                    except:
                        self.variables[total_data[0]] = [total_data[1]]
                        total_data = []

    def __print_variables(self):
        """Internals for variable tracing."""

        self.out.write("### Variable Values ###\n\n")
        for key in self.variables.keys():
            self.out.write("# " + str(key) + "\n")
            count = 1
            old = None
            if len(self.variables[key]) == 1:
                self.out.write(self.variables[key][0][:-1] + str("\n"))
            else:
                for value in self.variables[key][:-1]:
                    if str(value) != str(old):
                        if count == 1:
                            if old is not None:
                                self.out.write(old[:-1] + str("\n"))
                            count = 1
                        else:
                            self.out.write(old[:-1] + " (x" + str(count) + ")"
                                           + str("\n"))
                            count = 1
                        old = value
                    else:
                        count += 1
                else:
                    if str(self.variables[key][-2]) == \
                            str(self.variables[key][-1]):
                        self.out.write(self.variables[key][-1][:-1] + " (x" +
                                       str(count + 1) + ")" + str("\n"))
                    else:
                        if count > 1:
                            self.out.write(self.variables[key][-2][:-1] +
                                           " (x" + str(count) + ")"
                                           + str("\n"))
                        else:
                            self.out.write(self.variables[key][-2][:-1] +
                                           str("\n"))
                        self.out.write(self.variables[key][-1][:-1] +
                                       str("\n"))

    def __print_nodes(self, node_lists):
        for node_list in node_lists:
            for i in range(0, len(node_list)):
                self.out.write(node_list[i])
                if i == len(node_list)-1:
                    self.out.write("\n")
                else:
                    self.out.write("-->")

    def __print_methods(self):
        self.out.write("\n### Dynamic Call Graph ###\n\n")
        if len(self.nodes.keys()) > 0:
            node_lists = sorted(self.__paths(next(iter(self.nodes.keys()))))
            set_node = set()
            for a in node_lists:
                set_node.add(tuple(a))
            set_node = sorted(set_node, key=lambda t: len(t))
            self.__print_nodes(set_node)
        else:
            self.out.write("Nothing Found"+"\n\n")

    def __paths(self, v):
        """Generate the maximal cycle-free paths in graph starting at v.
        graph must be a mapping from vertices to collections of
        neighbouring vertices.
        """
        path = [v]  # path traversed so far
        seen = {v}  # set of vertices in path

        def search():
            dead_end = True
            for neighbour in self.nodes[path[-1]]:
                if neighbour not in seen:
                    dead_end = False
                    seen.add(neighbour)
                    path.append(neighbour)
                    for p in search():
                        yield p
                    path.pop()
                    seen.remove(neighbour)
            if dead_end:
                yield list(path)

        for p in search():
            yield p

    def __parse_methods(self):
        callers = []
        with open(self.complete_path) as target_file:
            for line in target_file.readlines():
                line = line.replace("\\Oblivion\\","\\USER\\")
                if "Invoked Method" in line:
                    func_name = line.split("** ")[-1][:-1]
                    if func_name in self.nodes.keys():
                        callers.append(func_name)
                        self.nodes[callers[-2]].append(func_name)
                    elif not callers:
                        callers.append(func_name)
                        self.nodes[func_name] = []
                    else:
                        callers.append(func_name)
                        self.nodes[callers[-1]] = []
                        self.nodes[callers[-2]].append(func_name)
                if "Return to Caller" in line:
                    if len(callers) > 1:
                        callers.remove(callers[-1])

    def __print_info(self):
        fname = self.output_name.split("_")[0]
        self.out.write("Date and Time: " +
                       str(time.strftime("%Y-%m-%d %H:%M"))+"\n")
        self.out.write("Hash 256: " + fname + "\n")

        program_name = self.program.capitalize()
        file_type = "File Type: " + program_name + "\n\n"
        self.out.write(file_type)

    def __get_executables(self):
        """Check for executables"""
        with open(self.complete_path) as target_file:
            for line in target_file.readlines():
                line = line.replace("\\Oblivion\\","\\USER\\")
                line = line.replace("^", "")
                line = line.replace("\n", "")
                if '.exe' in line.lower():
                    for executable in re.findall(r'[a-z|A-Z|0-9|\u4e00-\u9fff|\u3040-\u309f|\.|\_|\-\%|\\|\:|//]+.exe',
                                                 line.lower()):
                        self.exe_set.add(executable)

        if len(self.ioc_dict.keys()) > 0:
            for key, value in self.ioc_dict.items():
                self.exe_set.add(key)

    def __get_other_files(self):
        """Check for other dropped files"""
        with open(self.complete_path) as target_file:
            for line in target_file.readlines():
                line = line.replace("\\Oblivion\\","\\USER\\")
                line = line.replace("\n", "")
                if '.xls' in line.lower() or \
                    '.doc' in line.lower() or \
                    '.csv' in line.lower() or \
                    '.xlsx' in line.lower() or \
                    '.docx' in line.lower():
                    for file_type in re.findall(r'[a-z|A-Z|0-9|\u4e00-\u9fff|\u3040-\u309f|\.|\_|\-\%|\\|\:|//]+.\.csv',
                                                line.lower()):
                        self.other_files.add(file_type)
                    for file_type in re.findall(r'[a-z|A-Z|0-9|\u4e00-\u9fff|\u3040-\u309f|\.|\_|\-\%|\\|\:|//]+.\.xls',
                                                line.lower()):
                        self.other_files.add(file_type)
                    for file_type in re.findall(r'[a-z|A-Z|0-9|\u4e00-\u9fff|\u3040-\u309f|\.|\_|\-\%|\\|\:|//]+.\.doc',
                                                line.lower()):
                        self.other_files.add(file_type)
                    for file_type in re.findall(r'[a-z|A-Z|0-9|\u4e00-\u9fff|\u3040-\u309f|\.|\_|\-\%|\\|\:|//]+.\.xlsx',
                                                line.lower()):
                        self.other_files.add(file_type)
                    for file_type in re.findall(r'[a-z|A-Z|0-9|\u4e00-\u9fff|\u3040-\u309f|\.|\_|\-\%|\\|\:|//]+.\.docx',
                                                line.lower()):
                        self.other_files.add(file_type)

    def __print_executables(self):
        self.out.write("### Executable Files ### \n\n")
        if len(self.exe_set) == 0:
            self.out.write("Nothing Found"+"\n\n")
        else:
            for el in self.exe_set:
                self.out.write(str(el)+"\n")
            self.out.write("\n")

    def __print_other_files(self):
        self.out.write("### Other File Traces ### \n\n")
        if len(self.other_files) == 0:
            self.out.write("Nothing Found"+"\n\n")
        else:
            for el in self.other_files:
                self.out.write(str(el)+"\n")
            self.out.write("\n")

    def __get_domains(self):
        """Check for contacted Domains"""
        if os.path.getsize(self.complete_path) < 10000000:
            with open(self.complete_path) as target_file:
                for line in target_file.readlines():
                    line = line.replace("\\Oblivion\\","\\USER\\")
                    line = line.replace("\n", "")
                    GRUBER_URLINTEXT_PAT = \
                        re.compile(r'(?i)\b((?:https?://|www\d{0,3}[.]|[a-z0-9.\-]+[.][a-z]{2,4}/)(?:[^\s()<>]|\(([^\s()<>]+|(\([^\s()<>]+\)))*\))+(?:\(([^\s()<>]+|(\([^\s()<>]+\)))*\)|[^\s`!()\[\]{};:\'".,<>?\xab\xbb\u201c\u201d\u2018\u2019]))')
                    candidates = re.findall(GRUBER_URLINTEXT_PAT, line)
                    for c in candidates:
                        try:
                            url = urlparse(c[0])[1].replace("\n", "")
                            for candidate_url in self.url_set:
                                if url.find(candidate_url) != -1:
                                    self.url_set.remove(candidate_url)
                                    break
                            if url != "" and "." in url:
                                self.url_set.add(url)
                        except ValueError:
                            pass
        else:
            urls = re.findall(
                r"\b(?:(?:https?|ftp|file)://|www\.|ftp\.)[-A-Z0-"
                r"9+&@#/%=~_|$?!:,.]*[A-Z0-9+&@#/%=~_|$]",
                self.macro_text)
            for url in urls:
                if "/" in url:
                    url = url.split("/")[-1]
                elif "\\" in url:
                    url = url.split("\\")[-1]

                if url not in self.url_set:
                    self.url_set.add(url)

            with open(self.complete_path) as out_fd:
                urls = re.findall(r"\b(?:(?:https?|ftp|file)://|www\.|ftp\.)"
                                  r"[-A-Z0-9+&@#/%=~_|$?!:,.]*[A-Z0-9+&@#/%=~"
                                  r"_|$]", out_fd.read())
            for url in urls:
                if "/" in url:
                    url = url.split("/")[-1]
                elif "\\" in url:
                    url = url.split("\\")[-1]

                if url not in self.url_set:
                    self.url_set.add(url)

    def __print_urls(self):
        self.out.write("### Domain Traces ### \n\n")
        if len(self.url_set) == 0:
            self.out.write("Nothing Found"+"\n\n")
        else:
            for el in self.url_set:
                self.out.write(str(el)+"\n")
            self.out.write("\n")

    def __get_create_obj(self):
        with open(self.complete_path) as target_file:
            for line in target_file.readlines():
                if "createobject" in line.lower() and "(" not in line.lower():
                    self.create_objs.add(line.split()[-1])

    def __print_create_objects(self):
        self.out.write("### CreateObject Actions ### \n\n")
        if len(self.create_objs) == 0:
            self.out.write("Nothing Found"+"\n\n")
        else:
            for el in self.create_objs:
                self.out.write(str(el)+"\n")
            self.out.write("\n")

    def __get_shell(self):
        with open(self.complete_path) as target_file:
            for line in target_file.readlines():
                line = line.replace("\\Oblivion\\", "\\USER\\")
                line = line.replace("\n", "")
                if line.lower().startswith("shell"):
                    data = line.split()
                    if len(data) > 1:
                        if data[1] == "":
                            pass
                        else:
                            self.shell_commands.add(" ".join(data[1:]))
        self.shell_commands = self.__filter_commands(self.shell_commands)

    @staticmethod
    def __filter_commands(shell_commands):
        filtered = []
        if len(shell_commands):
            shell_commands = list(shell_commands)
            shell_commands.sort()
            for index in range(len(shell_commands) - 1):
                if not shell_commands[index + 1].startswith(shell_commands[index]):
                    filtered += [shell_commands[index]]
            filtered += [shell_commands[-1]]
        return set(filtered)

    def __get_cmd(self):
        with open(self.complete_path) as target_file:
            for line in target_file.readlines():
                line = line.replace("\\Oblivion\\", "\\USER\\")
                line = line.replace("\n", "")
                if "cmd.exe" in line.lower():
                    if " : " in line:
                        continue
                    if line.lower().startswith("shell"):
                        data = line.split()
                        self.cmd_commands.add(" ".join(data[1:]))
                    else:
                        self.cmd_commands.add(line)
            sorted_cmd = sorted(list(self.cmd_commands), key=lambda x: len(x))
            candidates_to_remove = []
            for i in range(0, len(sorted_cmd)-1):
                delta = \
                    sum(1 for a, b in zip(sorted_cmd[i], sorted_cmd[i+1])
                        if a != b) + \
                    abs(len(sorted_cmd[i]) - len(sorted_cmd[i+1]))
                if delta < 4:
                    candidates_to_remove.append(sorted_cmd[i])
            for c in candidates_to_remove:
                self.cmd_commands.remove(c)

    @staticmethod
    def __get_sandboxed_filename(file_name, sandbox):
        sbx_filename = os.path.join(
            file_name.replace(os.path.basename(file_name), "").replace("C:\\Users\\" +
            os.getenv("username"), "C:\\Sandbox\\" +
            os.getenv("username") + "\\" + sandbox + "\\user\\current"),
            os.path.basename(file_name))
        return sbx_filename

    def __extract_from_sandbox(self, file_name):
        sbx_file_name = self.__get_sandboxed_filename(file_name, self.__sandboxie_name)
        if os.path.exists(sbx_file_name):
            os.rename(sbx_file_name, file_name)

    def __get_interactions(self):
        interaction_manager_report = os.path.join(os.getcwd(), "data", "interaction_result.log")
        if os.path.exists(interaction_manager_report):
            with open(interaction_manager_report, "rb") as fb:
                rep = fb.read().decode('utf-8').splitlines()
            clicks = 0
            rep = list(np.unique(np.array(rep)))
            for x in rep:
                if "[*] Window forcefully closed." in x or "[x] Clicked on" in x:
                    clicks += 1
            self.interaction_lines = rep
            self.interaction_lines += ["Clicks: " + str(clicks)]
            os.remove(interaction_manager_report)
        else:
            self.interaction_lines = ["Nothing Found\n"]
            return

    def __get_interaction_API(self):
        with open(self.complete_path, "r") as f:
            lines = f.read().splitlines()
        msgbox_calls = 0
        inputbox_calls = 0
        for line in lines:
            if "messagebox" in line.lower():
                msgbox_calls += 1
            if "inputbox" in line.lower():
                inputbox_calls += 1
        self.interaction_API = {"MessageBox": msgbox_calls, "InputBox": inputbox_calls}

    def __print_interactions(self):
        self.out.write("### Interactions ### \n\n")
        for line in self.interaction_lines:
            self.out.write(line + "\n")

        self.__get_interaction_API()
        self.__print_interaction_API()

        self.out.write("\n")

    def __print_interaction_API(self):
        for method, calls in self.interaction_API.items():
            log_string = "Called legacy API " + method + " " + str(calls) + " times.\n"
            self.out.write(log_string)

    def __get_exception(self):
        with open(self.complete_path) as target_file:
            for line in target_file.readlines():
                line = line.replace("\\Oblivion\\","\\USER\\")
                line = line.replace("\n", "")
                if line.lower().startswith("exception:"):
                    data = line.split(" ")
                    if data[1] == "":
                        pass
                    else:
                        self.exceptions.add(" ".join(data[1:]))

    def __get_environs(self):
        with open(self.complete_path) as target_file:
            for line in target_file.readlines():
                line = line.replace("\\Oblivion\\","\\USER\\")
                line = line.replace("\n", "")
                if line.lower().startswith("environ"):
                    data = line.split(" ")
                    self.environs.add(" ".join(data[1:]))

    def __get_ext_calls(self):
        with open(self.complete_path) as target_file:
            for line in target_file.readlines():
                line = line.replace("\\Oblivion\\","\\USER\\")
                line = line.replace("\n", "")
                if "External Call" in line:
                    data = line.split("** ")
                    self.ext_calls.add(data[-1])

    def __get_sys_file_write(self):
        with open(self.complete_path) as target_file:
            for line in target_file.readlines():
                line = line.replace("\\Oblivion\\","\\USER\\")
                line = line.replace("\n", "")
                if "xlstart" in line.lower():
                    data = line.split("** ")
                    self.file_writes.add(data[-1])

    def __print_shell(self):
        self.out.write("### Shell Actions ### \n\n")
        if len(self.shell_commands) == 0:
            self.out.write("Nothing Found"+"\n\n")
        else:
            for el in set(self.shell_commands):
                el = el.replace("^", "")
                self.out.write(str(el)+"\n")
            self.out.write("\n")

    def __print_cmd(self):
        self.out.write("### Cmd Actions ### \n\n")
        if len(self.cmd_commands) == 0:
            self.out.write("Nothing Found"+"\n\n")
        else:
            for el in set(self.cmd_commands):
                el = el.replace("^", "")
                self.out.write(str(el)+"\n")
            self.out.write("\n")

    def __print_deobfuscation(self):
        self.out.write("### Deobfuscated Powershell ### \n\n")
        if len(self.__deobf_list) == 0:
            self.out.write("Nothing Found" + "\n\n")
        else:
            for el in self.__deobf_list:
                self.out.write(str(el) + "\n")
            self.out.write("\n")

    def __print_exception(self):
        self.out.write("### Exceptions ### \n\n")
        if len(self.exceptions) == 0:
            self.out.write("Nothing Found"+"\n\n")
        else:
            for el in self.exceptions:
                el = el.replace("^","")
                self.out.write(str(el)+"\n")
            self.out.write("\n")

    def __print_environs(self):
        self.out.write("### Environment Variables ### \n\n")
        if len(self.environs) == 0:
            self.out.write("Nothing Found"+"\n\n")
        else:
            for el in self.environs:
                el = el.replace("^","")
                self.out.write(str(el)+"\n")
            self.out.write("\n")

    def __print_ext_calls(self):
        self.out.write("### External Calls ### \n\n")
        if len(self.ext_calls) == 0:
            self.out.write("Nothing Found"+"\n\n")
        else:
            for el in self.ext_calls:
                el = el.replace("^","")
                self.out.write(str(el)+"\n")
            self.out.write("\n")

    def __print_sys_file_write(self):
        self.out.write("### System File Writes ### \n\n")
        if len(self.file_writes) == 0:
            self.out.write("Nothing Found"+"\n\n")
        else:
            for el in self.file_writes:
                el = el.replace("^", "")
                self.out.write(str(el)+"\n")
            self.out.write("\n")

    def __get_calls_from_macro(self):
        self.susp_calls_dict = {}
        self.exec_calls_dict = {}
        self.ioc_dict = {}
        macro_analysis_results = VBA_Parser(self.file_path).analyze_macros()
        for kw_type, keyword, description in macro_analysis_results:
            if kw_type == "AutoExec":
                self.exec_calls_dict[keyword] = description
            elif kw_type == "Suspicious":
                self.susp_calls_dict[keyword] = description
            elif kw_type == "IOC":
                self.ioc_dict[keyword] = description

    def __print_macro(self):
        self.out.write("### Auto Exec Methods ### \n\n")
        if len(self.exec_calls_dict.keys()) == 0:
            self.out.write("Nothing Found" + "\n\n")
        else:
            for key, value in self.exec_calls_dict.items():
                self.out.write("{} -> {}\n".format(key, value))
            self.out.write("\n")

        self.out.write("### Suspicious calls ### \n\n")
        if len(self.susp_calls_dict.keys()) == 0:
            self.out.write("Nothing Found" + "\n\n")
        else:
            for key, value in self.susp_calls_dict.items():
                if "(" in value:
                    value = value[:value.index("(")]
                self.out.write("{} -> {}\n".format(key, value))
            self.out.write("\n")
        self.out.write("\n")

    @staticmethod
    def __write_file(path, text):
        with open(path, "w") as fd:
            return fd.write(text)

    @staticmethod
    def __is_powershell_running():
        for p in process_iter():
            try:
                if "powershell.exe" == p.name():
                    return True
            except ValueError:
                continue
        return False

    @staticmethod
    def __read_file(file_path):
        with open(file_path, "rb") as fd:
            bytecode = fd.read()
            while True:
                try:
                    decoded = bytecode.decode("utf-8")
                except UnicodeDecodeError as ex:
                    s = str(ex).find("position ") + 9
                    e = str(ex).find(":", s)
                    number = int(str(ex)[s: e])
                    try:
                        bytecode = bytecode[:number] + bytecode[number + 1:]
                    except IndexError:
                        bytecode = bytecode[:number]
                else:
                    break
        return decoded

    @staticmethod
    def __get_plainscript(txt):
        txt_list = txt.replace("\x00", "").replace("\xff", "").\
            replace("\xfe", "").splitlines()
        for line in txt_list:
            if "Plainscript" in line:
                script = txt_list[txt_list.index(line) + 3]
                last_line = script
                idx = 1
                while True:
                    script_next_line = txt_list[
                        txt_list.index(last_line) + idx]
                    last_line = script_next_line
                    if script_next_line.strip() == "":
                        break
                    else:
                        script = script + script_next_line
                        idx += 1
                return script
            if "Active Malware Hosting URLs:" in line:
                return ""
        return ""

    @staticmethod
    def __clean_substr(text_str, substring):
        text_str = \
            text_str[text_str.lower().index(substring) + len(
                substring):].strip()
        if text_str[0] == '"' and text_str[-1] == '"':
            text_str = text_str[1: -1]
        return text_str

    @staticmethod
    def __get_macro_path():
        return sys.argv[3]

    @staticmethod
    def __clean_list(el_list):
        try:
            for v in range(0, el_list.count("")):
                el_list.remove("")
            return el_list
        except ValueError:
            return el_list

    @staticmethod
    def __get_complete_path(report_name):
        folder_path = os.path.join(os.path.abspath("."), sys.argv[2])
        complete_path = os.path.join(folder_path, report_name)
        return complete_path

    @staticmethod
    def __is_assignment(line):
        data = line.split(" ")
        if (len(data) == 3) and data[2] == "\n":
            return True
        else:
            return False

    @staticmethod
    def __get_macro(original_macro_path):
        with open(original_macro_path, 'rb') as fd:
            original_macro = fd.read().decode("utf-8")
        return original_macro

    def post_processing(self, report_file_path):
        self.save_report(report_file_path)
        return self.is_powershell_present()


if __name__ == "__main__":
    file_path = sys.argv[1]
    output_file_path = sys.argv[2]
    original_macro_path = sys.argv[3]
    file_ext = sys.argv[4]
    report_file_path = sys.argv[5]
    powerdecode_path = sys.argv[6]
    sandboxie_path = sys.argv[7]
    sandbox_name = sys.argv[8]

    post_processing = \
        PostProcessing(file_path, output_file_path, original_macro_path, file_ext,
                       powerdecode_path, sandboxie_path, sandbox_name)

    post_processing.save_report(report_file_path)
    if post_processing.is_powershell_present():
        print(u"Found PowerShell code.")
    else:
        print(u"No PowerShell found.")
    exit(0)
