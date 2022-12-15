import json
import os
import re
import itertools


class MacroInstrumentationException(Exception):
    pass


class MacroInstrumentation:
    IMPORTANT_FUNCTIONS_LIST = ["CallByName"]
    EXTERNAL_FUNCTION_DECLARATION_REGEXP = re.compile(
        r'Declare *(?:PtrSafe)? *(?:Sub|Function) *(.*?) *Lib *"[^"]+" *'
        r'(?:Alias *"([^"]+)")?')
    EXTERNAL_FUNCTION_REGEXP = None
    EXTERNAL_FUNCTION_REGEXP_2 = None
    BEGIN_FUNCTION_REGEXP = re.compile(r"\s*(?:function|sub) (.*?)\(",
                                       re.IGNORECASE)
    END_FUNCTION_REGEXP = re.compile(r"^\s*end\s*(?:function|sub)",
                                     re.IGNORECASE)
    METHOD_CALL_REGEXP = re.compile(
        r"^\s*([a-z_\.0-9]+\.[a-z_0-9]+)\s*\((.*)\)", re.IGNORECASE)
    METHOD_CALL_REGEXP_2 = re.compile(r"^\s*([a-z_\.0-9]+\.[a-z_0-9]+) +(.*)",
                                      re.IGNORECASE)
    IMPORTANT_FUNCTION_REGEXP = re.compile(
        r"^(Set\s+)?([a-z0-9]+\s*=)?\s*(" + "|".join(
            IMPORTANT_FUNCTIONS_LIST) + r")\s*\((.*)\)", re.IGNORECASE)
    IMPORTANT_FUNCTION_REGEXP_2 = re.compile(
        r"^(Set\s+)?([a-z0-9]+\s*=)?\s*(" + "|".join(
            IMPORTANT_FUNCTIONS_LIST) + ")(.*)", re.IGNORECASE)

    is_auto_open_function = False
    current_function_name = ""
    declared_function_original_names = {}

    lines = []
    i = 0
    counter = 0
    line = ""

    output = []
    constant_values = []

    disable = False
    is_init_added = False

    def __init__(self, original_macro_path, exclusions_path, instrumented_macro_data_path):
        with open(original_macro_path, "r") as omf:
            self.__original_macros = json.load(omf)
        self.__instrumented_macros = {k: None for k in self.__original_macros.keys()}
        self.__exclusions_path = exclusions_path
        # self.__get_lines_to_exclude()

        self.lines_to_exclude = None
        self.lines_to_exclude_params = None
        self.instrumented_macro_data_path = instrumented_macro_data_path

    def run(self):
        for macro_name in self.__instrumented_macros.keys():
            self.__instrumented_macros[macro_name] = \
                self.__instrument(self.__original_macros[macro_name])
        with open(self.instrumented_macro_data_path, "w") as jsf:
            json.dump(self.__instrumented_macros, jsf, indent=4)

    def __instrument(self, macro):
        self.output = list()
        self.__prepare_external_function_calls(macro.replace(" _\n", " "))
        macro_lines = [line.strip() for line in macro.splitlines()]
        self.lines = macro_lines
        self.__sanitize()
        self.__dispatch()
        return "\n".join(self.output)

    """
        Fix inline code
    """
    def __sanitize(self):
        new_lines = []
        self.lines = list(map(self.__delete_comments_spaces_static, self.lines))
        self.lines = [line for line in self.lines if line != ""]
        for l in self.lines:
            line_without_str = re.sub(r'\".*?\"', '', l)
            if l[:13] == "Attribute VB_":
                continue
            elif ':' in line_without_str:
                if 'Sub ' in line_without_str or \
                        'Function ' in line_without_str or \
                        'Call ' in line_without_str:
                    if self.__is_char_before_flag(l) is False:
                        l = l.replace(':=', '!![COLONEQUAL]!!')
                        suspicious_lines = l.split(':')
                        suspicious_lines = [sus.replace('!![COLONEQUAL]!!', ':=')
                                            for sus in suspicious_lines]
                        new_lines += suspicious_lines
                    else:
                        new_lines.append(l)
                elif "With " in line_without_str:
                    new_lines += self.__manage_colon_in_line(l)
                    if 'End With' not in l:
                        new_lines.append('End With')
                elif "If " in line_without_str and \
                        "Then " in line_without_str:
                    suspicious_lines = list()

                    starting_line = l[:l.find(" Then")] + " Then \n"
                    starting_split = starting_line.split(":")
                    suspicious_lines += starting_split

                    if "Else " in line_without_str:
                        temp_line = l[l.find(" Then") + 6: l.rfind(" Else ")].split(":")
                    else:
                        temp_line = l[l.find(" Then") + 6:].split(":")

                    if type(temp_line) is str:
                        suspicious_lines.append(temp_line)
                    else:
                        suspicious_lines += temp_line

                    if "Else" in line_without_str:
                        suspicious_lines.append("Else")
                        temp_line = l[l.rfind(" Else ") + 6:].split(":")
                        if type(temp_line) is str:
                            suspicious_lines.append(temp_line)
                        else:
                            suspicious_lines += temp_line
                    else:
                        pass

                    new_lines += suspicious_lines + ["End If "]
                else:
                    suspicious_lines = self.__manage_colon_in_line(l)
                    if type(suspicious_lines) is str:
                        new_lines.append(suspicious_lines)
                    else:
                        new_lines += suspicious_lines
            elif len(l) <= 1:
                continue
            else:
                new_lines.append(l)
        self.lines = new_lines

    def __is_instr_exception(self):
        exception_list = ["strconv(", "cstr(", "lines(", "cbyte(", "cint(",
                          "csng(", "cstr(", "cdate(", "clng(", "sin(", "sgn("]
        if any(el in self.line.lower() for el in exception_list):
            return True
        else:
            return False

    def __get_lines_to_exclude(self):
        if os.path.exists(self.__exclusions_path) and \
                os.path.getsize(self.__exclusions_path) != 0:
            with open(self.__exclusions_path, "r") as f:
                lines = f.read().splitlines()
            self.lines_to_exclude = [l for l in lines if "#P#" not in l]
            self.lines_to_exclude_params = [l.replace(" #P#", '') for l in
                                            lines if "#P#" in l]
        else:
            self.lines_to_exclude = None
            self.lines_to_exclude_params = None

    def __manage_colon_in_line(self, line):
        """
        Crucial cases:
            := used in condiditions to se parameters of an object
            : inside quote marks
            : at the end of the line as last char
        """
        if line.count(':') == 1 and line[-1] == ':':
            return line

        num_quote = 0
        last_id = 0
        return_list = []
        for id in range(0, len(line)):
            char = line[id]
            if char == '\"':
                num_quote += 1
            if char == ':':
                if id != len(line) - 1 and line[id + 1] == '=':
                    continue
                else:
                    try:
                        rest = num_quote % 2
                    except ZeroDivisionError:
                        rest = 0
                    if rest == 0:
                        return_list.append(line[last_id: id])
                        last_id = id + 1
            if id == len(line) - 1:
                return_list.append(line[last_id:])
        return return_list

    def __get_index(self, word, line):
        try:
            return line.index(word)
        except ValueError:
            return -1

    def __is_char_before_flag(self, line):
        flag_index = self.__get_index(':', line)
        sub_index = self.__get_index('Sub', line)
        function_index = self.__get_index('Function', line)
        call_index = self.__get_index('Call', line)
        if sub_index != -1 and flag_index < sub_index:
            return True
        elif function_index != -1 and flag_index < function_index:
            return True
        elif call_index != -1 and flag_index < call_index:
            return True
        else:
            return False

    def __delete_comments_spaces_static(self, line):
        line_mod = re.sub(r'\".*?\"', '', line)
        if '\'' in line_mod:
            if line[-1] == '_':
                self.lines.remove(self.lines[self.lines.index(line) + 1])
            return self.__remove_static_flag(line[:line.rfind('\'')].strip())
        else:
            return self.__remove_static_flag(line.strip())

    def __remove_static_flag(self, line):
        line_mod = re.sub(r'\".*?\"', '', line)
        if 'Static Function' in line_mod:
            return line.replace('Static Function', 'Public Function')
        elif 'Static Sub' in line_mod:
            return line.replace('Static Sub', 'Public Sub')
        elif 'Static' in line_mod and 'Function' not in line_mod and \
                'Sub' not in line_mod:
            return line.replace('Static', 'Dim')
        else:
            return line

    def __remove_private_flag(self, line):
        line_mod = re.sub(r'\".*?\"', '', line)
        if 'Private Function' in line_mod:
            return line.replace('Private Function', 'Public Function')
        elif 'Private Sub' in line_mod:
            return line.replace('Private Sub', 'Public Sub')
        elif 'Private ' in line_mod and 'Function' not in line_mod and \
                'Sub' not in line_mod:
            return line.replace('Private ', 'Public ')
        else:
            return line

    def __add_content_to_output(self, content):
        self.output.append(content)

    def __add_line_to_output(self, i):
        self.__add_content_to_output(self.__get_line(i))

    def __add_current_line_to_output(self):
        self.__add_content_to_output(self.__get_current_line())

    """
        See: https://msdn.microsoft.com/en-us/library/ba9sxbw4.aspx
        We simply skip those kind of lines
    """

    def __is_long_line(self):
        counter = 0
        while True:
            if self.__get_line(self.current_line + counter)[-2:] == " _":
                counter += 1
            else:
                return counter

    """
        Those functions starts automatically when a document is opened
        We need to initialize our logger there
    """

    def __is_autostart_function(self):
        auto_open, auto_close = self.__get_auto_exec(self.current_function_name.lower())
        return auto_open or auto_close

    @staticmethod
    def __get_auto_exec(code_ref) -> (bool, bool):
        if os.path.isfile(code_ref):
            with open(code_ref, "r") as icf:
                code = icf.read().lower()
        else:  # if here, "code_ref" is the function name: no need to process
            code = code_ref
            prefixes = ["auto", "document", "workbook"]
            joints = ["", "_"]
            suffixes = ["open", "close"]
            flags = list(False for _ in suffixes)

            keywords = [''.join(x) for x in itertools.product(prefixes, joints, suffixes)]

            for index in range(len(suffixes)):
                suffix = suffixes[index]
                check_list = [x for x in keywords if x.endswith(suffix)]
                for kw in check_list:
                    if kw in code:
                        flags[index] = True
                        break

        return flags
    """
        If current line is function declaration, get its name
        Its then used for checking if we return something from this function
    """

    def __is_begin_function_line(self):
        line_without_str = re.sub(r'\".*?\"', '', self.__get_current_line())
        matched = re.search(self.BEGIN_FUNCTION_REGEXP, line_without_str)
        if matched and not "declare" in line_without_str.lower():
            self.current_function_name = matched.group(1).strip()
            return True

        return False

    """
        Check if its function end
        We add there exception handler
    """

    def __is_end_function_line(self):
        line_without_str = re.sub(r'\".*?\"', '', self.__get_current_line())
        if re.search(self.END_FUNCTION_REGEXP, line_without_str):
            return True
        elif 'end sub' in line_without_str.lower() and ':' in line_without_str:
            return True
        return False

    """
        For non-object return types, you assign the value to the name of function
        So we can check if this function return something and add our logger here
    """

    def __is_return_string_from_function_line(self):
        if self.current_function_name != "":
            if re.search("^ *" + re.escape(self.current_function_name) + " *=",
                         self.__get_current_line(), re.IGNORECASE):
                return True
        return False

    """
        Found all `Class.Method params` and Class.Method (params)` calls
        We dont support params passed by name, like name:=value
        We need to check if its not method assign like variable = Class.Method
    """

    def __is_method_call_line(self):
        line_without_str = re.sub(r'\".*?\"', '', self.__get_current_line())
        matched = re.search(self.METHOD_CALL_REGEXP, line_without_str)
        if matched:
            matched = re.search(self.METHOD_CALL_REGEXP,
                                self.__get_current_line())
        else:
            matched = re.search(self.METHOD_CALL_REGEXP_2, line_without_str)
            if matched:
                matched = re.search(self.METHOD_CALL_REGEXP_2,
                                    self.__get_current_line())
            else:
                return False

        if matched:
            method_name = matched.group(1).strip()
            params = matched.group(2).strip()

            if len(params) > 0:
                # We dont support params passed by name
                # And Class.Method = 1
                if "=" in self.__get_current_line():
                    return False

                return [method_name, params]

            return [method_name]
        return False

    """
        Check if its call to previously defined external library
    """

    def __is_external_function_call_line(self):
        # Do we have any external declarations
        if self.EXTERNAL_FUNCTION_REGEXP == None:
            return False

        # Skip if its declaration, not usage
        if re.search(self.EXTERNAL_FUNCTION_DECLARATION_REGEXP,
                     self.__get_current_line()):
            return False

        matched = re.search(self.EXTERNAL_FUNCTION_REGEXP,
                            self.__get_current_line())

        if not matched:
            matched = re.search(self.EXTERNAL_FUNCTION_REGEXP_2,
                                self.__get_current_line())

        if matched:
            name = matched.group(1).strip()
            rest = matched.group(2).strip()
            if name in self.declared_function_original_names:
                rest = re.sub(r"ByVal", "", rest, flags=re.IGNORECASE)
                return [name, rest]

        return False

    def __is_enum_line(self):
        if " Enum " in self.__get_current_line():
            return True

    def __is_comment(self):
        line_without_str = re.sub(r'\".*?\"', '', self.__get_current_line())
        if line_without_str.startswith("Rem") or \
                line_without_str.startswith("'"):
            return True

    def __is_end_enum_line(self):
        if " End Enum " in self.__get_current_line():
            return True

    """
        We hook some important function like CallByName which cannot be hooked using another techniques
    """

    def __is_important_function_call_line(self):
        matched = re.search(self.IMPORTANT_FUNCTION_REGEXP,
                            self.__get_current_line())

        if not matched:
            matched = re.search(self.IMPORTANT_FUNCTION_REGEXP_2,
                                self.__get_current_line())

        if matched:
            name = matched.group(3).strip()
            rest = matched.group(4).strip()

            return [name, rest]

        return False

    def __is_dim_line(self):
        if self.__get_current_line().startswith("Dim"):
            return True

    """
        Find all external library declarations like:
        Private Declare Function GetDesktopWindow Lib "user32" () As Long
    """

    def __prepare_external_function_calls(self, content):
        declared_function_list = []

        for f in re.findall(self.EXTERNAL_FUNCTION_DECLARATION_REGEXP,
                            content):
            declared_function_list.append(re.escape(f[0].strip()))
            if f[1] != "":
                self.declared_function_original_names[f[0].strip()] = f[
                    1].strip()
            else:
                self.declared_function_original_names[f[0].strip()] = f[
                    0].strip()

        if len(declared_function_list) > 0:
            self.EXTERNAL_FUNCTION_REGEXP = re.compile(
                "({})\s*\((.*)\)".format("|".join(declared_function_list)))
            self.EXTERNAL_FUNCTION_REGEXP_2 = re.compile(
                "({})\s*(.*)".format("|".join(declared_function_list)))

    """
        Get single line by ids number
    """

    def __get_line(self, i):
        if i < self.counter:
            return self.lines[i]
        return ""

    """
        Get current line
    """

    def __get_current_line(self):
        if self.current_line < self.counter:
            return self.lines[self.current_line]
        return ""

    """
        Set current line, so we can then use get_current_line
    """

    def __set_current_line(self):
        self.line = self.__get_line(self.current_line)

    """
        Some function have special aliases for null support
    """

    def __replace_function_aliases(self):
        line = self.lines[self.current_line]
        line = re.sub(r"(VBA\.CreateObject)", "CreateObject", line,
                      flags=re.IGNORECASE)
        line = re.sub(r"Left\$", "Left", line, flags=re.IGNORECASE)
        line = re.sub(r"Right\$", "Right", line, flags=re.IGNORECASE)
        line = re.sub(r"Mid\$", "Mid", line, flags=re.IGNORECASE)
        line = re.sub(r"Environ\$", "Environ", line, flags=re.IGNORECASE)
        self.lines[self.current_line] = line

    """
        Main program loop
    """

    def __dispatch(self):
        self.current_line = 0
        self.counter = len(self.lines)
        is_enum = False
        jump_line = False
        while self.current_line < self.counter:
            self.__set_current_line()
            self.__replace_function_aliases()
            # Define the line type and handle various possibilities
            if (self.lines_to_exclude is not None and
                self.line in self.lines_to_exclude) or \
                    self.__is_instr_exception():
                self.__add_current_line_to_output()
                jump_line = True
            elif self.__is_long_line() > 0:
                self.__parse_long_line(self.__is_long_line())
                continue
            elif self.__is_comment():
                self.__add_content_to_output(self.__get_current_line())
            elif self.__is_begin_function_line() is True:
                self.__parse_function_begin_line(self.__is_autostart_function())
            elif self.__is_return_string_from_function_line() is True:
                self.__parse_string_return()
            elif self.__is_method_call_line() != False:
                self.__parse_method_call_line(self.__is_method_call_line())
            elif self.__is_external_function_call_line() != False:
                self.__parse_external_function_call(
                    self.__is_external_function_call_line())
            elif self.__is_important_function_call_line() != False:
                self.__parse_important_function_call(
                    self.__is_important_function_call_line())
            elif self.__is_end_function_line():
                self.__parse_end_function_line()
            elif self.__is_enum_line():
                is_enum = True
                self.__add_content_to_output(self.__get_current_line())
            elif self.__is_dim_line():
                self.__add_content_to_output(self.__get_current_line())
            elif is_enum is True and self.__is_end_enum_line():
                is_enum = False
            else:
                self.__add_current_line_to_output()
            current_line = self.__get_current_line()
            row = self.__get_current_line().split()
            if is_enum is not True and jump_line is False:
                self.__handle_function_details(current_line, row)
            elif jump_line is True:
                jump_line = False
                self.current_line += 1
            else:
                self.current_line += 1

    def __handle_function_details(self, current_line, line_data):
        if "\"" in current_line:
            str_line = current_line[0:current_line.find("\"")]
        elif " '" in current_line:
            str_line = current_line[0:current_line.find("'")]
        else:
            str_line = current_line
        if "=" in str_line and "Function" not in str_line and \
                "Sub" not in str_line and "If" not in str_line \
                and line_data[0] != self.current_function_name:
            if ">=" not in str_line or "<=" not in str_line:
                for element in line_data:
                    if element == "=":
                        equal_index = line_data.index("=")
                        equal_string = " ".join(line_data[:equal_index])
                        if "'" != line_data[0] and "'" != line_data[0][
                            0] and self.disable == False:
                            if "Const " in equal_string:
                                # Append the constant value to the list
                                self.constant_values.append(line_data[-1])
                            elif "If" in equal_string:
                                pass
                            elif "ElseIf" in equal_string:
                                self.__elseif_handler(equal_index, line_data)
                            elif "Set" in equal_string:
                                self.__set_handler(current_line, equal_index,
                                                  line_data)
                            elif "For" in equal_string:
                                self.__for_handler()
                            elif "Loop Until" in equal_string:
                                self.__loop_until_handler(equal_index,
                                                         line_data)
                            elif "With" == equal_string:
                                pass
                            elif "ChDir" in equal_string:
                                self.__chdrid_handler(equal_index, line_data)
                            elif "Loop While" in equal_string:
                                pass
                            elif "Shell" == equal_string:
                                self.__shell_handler(line_data)
                            elif "Print" == equal_string:
                                pass
                            elif ".InsertLines" == equal_string:
                                pass
                            elif "If" in current_line:
                                pass
                            else:  # It's an assignment without vb keywords
                                self.__handle_standard_assignment(equal_index,
                                                                 line_data,
                                                                 current_line)
                    else:
                        pass
        elif "For" in current_line[0:3]:
            self.disable = True
        elif "Next" in current_line[0:4]:
            self.disable = False
        else:
            pass
        self.current_line += 1

    def __handle_standard_assignment(self, equal_index, line_data,
                                    current_line):
        result = ' '.join(map(str, line_data[(equal_index + 1):]))
        rep_str = ' '.join(map(str, line_data[:equal_index])).replace("\"",
                                                                      "\"\"")
        if "Array" in current_line:
            self.__add_content_to_output(
                "oblivion_log(" + "\"" + rep_str + " = " + "\"" + ")")
            self.__add_content_to_output("oblivion_log(" + rep_str + ")")

        elif "NaN" in current_line or "Dim" in current_line:  # Array log is not supported yet
            pass
        elif result == self.current_function_name:
            pass
        else:
            if rep_str == self.current_function_name:
                self.__add_content_to_output("oblivion_log(" + result + ")")
            else:
                self.__add_content_to_output(
                    "oblivion_log(" + "\"" + rep_str + " = " + "\"" + ")")
                self.__add_content_to_output("oblivion_log(" + result + ")")

    def __shell_handler(self, line_data):
        result = ' '.join(map(str, line_data[:]))
        self.__add_content_to_output("oblivion_log(" + result + ")")

    def __chdrid_handler(self, equal_index, line_data):
        result = ' '.join(map(str, line_data[equal_index - 1:equal_index]))
        self.__add_content_to_output("oblivion_log(" + result + ")")

    def __loop_until_handler(self, equal_index, line_data):
        result = ' '.join(map(str, line_data[2:equal_index]))
        self.__add_content_to_output("oblivion_log(" + result + ")")

    def __for_handler(self):
        self.disable = False

    def __set_handler(self, current_line, equal_index, line_data):
        if "Array" in current_line or "NaN" in current_line:  # Array log is not supported yet
            pass

        self.__add_content_to_output("oblivion_log(" + "\"" + str(
            line_data[equal_index - 1]) + " = " + "\"" + ")")
        self.__add_content_to_output("oblivion_log(" + "\"" + str(
            " ".join(line_data[equal_index + 1:])).replace("\"",
                                                           "\'") + "\"" + ")")

    def __elseif_handler(self, equal_index, line_data):
        if "(" in line_data[1][0]:
            temp = line_data
            temp[1] = temp[1][1:]
            result = ' '.join(map(str, temp[1:equal_index]))
            self.__add_content_to_output("oblivion_log(" + result + ")")
        else:
            result = ' '.join(map(str, line_data[1:equal_index]))
            self.__add_content_to_output("oblivion_log(" + result + ")")

    def __parse_end_function_line(self):
        if self.__is_autostart_function():
            self.__add_content_to_output('oblivion_exception_handler:')
            self.__add_content_to_output(
                'oblivion_log ("Exception: " & Err.Description)')
            self.__add_content_to_output('On Error Resume Next')
            self.__add_content_to_output("oblivion_log_object.Close")
        self.__add_current_line_to_output()

    def __parse_important_function_call(self, is_important_function_call_line):
        """Verify the goodness of the function called log_call_to_function"""
        if "))(0" not in self.__get_current_line():
            self.__add_content_to_output('log_call_to_function "{}", {}'.format(
                is_important_function_call_line[0],
                is_important_function_call_line[1]))
        self.__add_current_line_to_output()
        self.__add_content_to_output(
            "oblivion_log(" + "\"" + "**Return to Caller** " + "\"" + ")")

    def __parse_external_function_call(self, is_external_function_call_line):
        """Verify the goodness of the function called log_call_to_function"""
        if "))(0" not in self.__get_current_line():
            pass
        self.__add_current_line_to_output()
        self.__add_content_to_output(
            "oblivion_log(" + "\"" + "**Return to Caller** " + "\"" + ")")

    def __parse_method_call_line(self, is_method_call_line):
        if len(is_method_call_line) == 1:
            self.__add_content_to_output(
                'log_call_to_method "{}"'.format(is_method_call_line[0]))
        else:
            sanity_check = " ".join(
                [is_method_call_line[0], is_method_call_line[1]])
            check = False
            if sanity_check.count("(") == sanity_check.count(")"):
                if sanity_check.find("(") < sanity_check.find(")"):
                    check = True
            if check is True:
                self.__add_content_to_output(
                    "log_call_to_method \"{}\", {}".format(
                        is_method_call_line[0], is_method_call_line[1]))
        self.__add_current_line_to_output()
        if "." in is_method_call_line[
            0]:  # this should be a system method call and therefore should not be called
            self.__add_content_to_output(
                "oblivion_log(" + "\"" + "**External Call** " + str(
                    is_method_call_line[0]) + " \"" + ")")
            pass
        else:
            self.__add_content_to_output(
                "oblivion_log(" + "\"" + "**Return to Caller** " + "\"" + ")")

    def __parse_string_return(self):
        self.__add_current_line_to_output()
        self.__add_content_to_output(
            'log_return_from_string_function "{}", {}'.format(
                self.current_function_name, self.current_function_name))
        self.__add_content_to_output(
            "oblivion_log(" + "\"" + "**Return to Caller** " + "\"" + ")")

    def __parse_function_begin_line(self, is_autostart_function):
        if is_autostart_function:
            self.__parse_autostart_function()
        else:
            self.current_lines_auto_open_function = False
            self.__add_current_line_to_output()
            self.__add_content_to_output(
                "oblivion_log(" + "\"" + "**Invoked Method** " + self.current_function_name + "\"" + ")")
            if self.lines_to_exclude_params is None:
                params_list = self.__get_params_list()
                if len(params_list) > 0:
                    self.__add_content_to_output(
                        "oblivion_log(" + "\"" + "**Invoked Method Parameters** " + "\"" + ")")
                    for param in params_list:
                        self.__add_content_to_output(
                            "oblivion_log(" + "\"" + param + " = \"" + ")")
                        self.__add_content_to_output("oblivion_log(" + param + ")")
                    self.__add_content_to_output(
                        "oblivion_log(" + "\"" + "**End Invoked Method Parameters** " + "\"" + ")")
            elif self.lines_to_exclude_params is not None and \
                    self.line not in self.lines_to_exclude_params:
                params_list = self.__get_params_list()
                if len(params_list) > 0:
                    self.__add_content_to_output(
                        "oblivion_log(" + "\"" + "**Invoked Method Parameters** " + "\"" + ")")
                    for param in params_list:
                        self.__add_content_to_output(
                            "oblivion_log(" + "\"" + param + " = \"" + ")")
                        self.__add_content_to_output("oblivion_log(" + param + ")")
                    self.__add_content_to_output(
                        "oblivion_log(" + "\"" + "**End Invoked Method Parameters** " + "\"" + ")")

    def __get_params_list(self):
        as_list = [" AS ", " As ", " aS ", " as "]
        params_list = []
        line_without_brackets_content = self.__get_str_without_brackets_content(
            self.line)
        as_el = [el for el in as_list if el in line_without_brackets_content]
        if len(as_el) > 0 and \
                line_without_brackets_content.index(
                    "(") < line_without_brackets_content.index(as_el[0]):
            as_el = as_el[0]
            line = self.line[:self.line.rfind(as_el)]
            if len(line) > 0:
                line = line[line.index("(") + 1: line.rfind(")")]
            else:
                line = self.line[
                       self.line.index("(") + 1: self.line.rfind(")")]
        else:
            line = self.line[self.line.index("(") + 1: self.line.rfind(")")]

        if len(line) > 0:
            line = line.split(",")

            for param in line:
                param_list = param.split(" ")
                params_list_lower = param.lower().split(" ")
                if "as" in params_list_lower:
                    as_index = params_list_lower.index("as")
                    params_list.append(param_list[as_index - 1])
                else:
                    params_list.append(param_list[-1])
            return params_list
        else:
            return []

    def __get_str_without_brackets_content(self, string):
        return_str = ""
        is_open = False
        is_close = True
        for id in range(0, len(string)):
            char = string[id]
            if char == '(':
                if is_close is True and is_open is False:  # first (
                    return_str += char
                    is_open = True
                    is_close = False
            elif char == ')':
                if is_open is True and is_close is False:
                    return_str += char
                    is_open = False
                    is_close = True
            else:
                if is_open is False and is_close is True:
                    return_str += char
        return return_str

    def __handle_inline_function(self):
        """Handle functions in one line only"""
        line = self.lines[self.current_line]
        data = line.split(":")
        invoked_method = "oblivion_log(" + "\"" + "**Invoked Method** " + self.current_function_name + "\"" + ")"
        command = data[
                      0] + ": " + "On Error GoTo oblivion_exception_handler: " + "oblivion_init: " + invoked_method + ": "
        for i in range(1, len(data)):
            if i == len(data) - 1:
                command = command + 'oblivion_exception_handler: '
                command = command + 'oblivion_log ("Exception: " & Err.Description)' + ": "
                command = command + 'On Error Resume Next' + ": "
                command = command + data[i]
            else:
                command = command + data[i] + ": "
        self.__add_content_to_output(command)

    def __parse_autostart_function(self):
        self.current_lines_auto_open_function = True
        if "sub" in self.lines[self.current_line].lower() and "end sub" in \
                self.lines[self.current_line].lower():
            self.__handle_inline_function()
        else:
            self.__add_current_line_to_output()
            self.__add_content_to_output(
                "On Error GoTo oblivion_exception_handler:")
            self.__add_content_to_output("oblivion_init")
            self.__add_content_to_output(
                "oblivion_log(" + "\"" + "**Invoked Method** " + self.current_function_name + "\"" + ")")

    def __parse_long_line(self, is_long_line):
        for ii in range(self.current_line,
                        self.current_line + is_long_line + 1):
            self.__add_line_to_output(ii)
        self.current_line += is_long_line + 1
