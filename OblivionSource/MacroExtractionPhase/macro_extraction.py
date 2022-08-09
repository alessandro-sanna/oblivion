import oletools.olevba
from oletools.olevba import VBA_Parser
import json
import os


class MacroExtractionException(Exception):
    pass


class MacroExtraction:
    def __init__(self, file_path, extensions, original_macro_data_path):
        if file_path.split('.')[-1] in extensions:
            self.__office_file_path = file_path
        else:
            raise MacroExtractionException("The given file is not a recognized office file")
        self.__macro_data = dict()
        self.__original_macro_data_path = original_macro_data_path

    def run(self):
        self.__extract_macro()
        return self.__macro_data

    def __extract_macro(self):
        try:
            vb_parser = VBA_Parser(self.__office_file_path)
        except oletools.olevba.FileOpenError:
            raise MacroExtractionException(f"File looks like a Office document but isn't")

        there_is_code = False
        try:
            vb_parser.extract_all_macros()
        except IndexError:
            raise MacroExtractionException(u"Macro(s) found, but they were impossible to extract.")

        if vb_parser.detect_vba_macros():
            for (filename, stream_path, vba_filename, vba_code) in \
                    vb_parser.extract_macros():
                new_macro = self.__parse_macro(vba_filename, vba_code)
                if new_macro[vba_filename].strip():
                    there_is_code = True

                self.__macro_data.update(new_macro)
            vb_parser.close()

            if not there_is_code:
                raise MacroExtractionException(u"All macros are empty")

            with open(self.__original_macro_data_path, "w") as fpJson:
                json.dump(self.__macro_data, fpJson, indent=4)

        else:
            vb_parser.close()
            raise MacroExtractionException(u"No VBA macro found.")

    @staticmethod
    def __parse_macro(vba_filename, vba_code):
        code = "\n".join([line for line in vba_code.splitlines()
                          if "Attribute VB_" not in line]).strip() + "\n"
        macro_data = {vba_filename: code}
        return macro_data



"""    
try:
    self.__macro_list.append("VBA MACRO " + vba_filename + "~~\n")
    self.__macro_list.append('-*' * 30 + '\n\n')
    code = "\n".join([line for line in vba_code.splitlines()
                     if "Attribute VB_" not in line]).strip() + "\n"
    self.__macro_list.append(code)
    self.__macro_list.append('-' * 60 + '\n\n')
except TypeError:
    raise MacroExtractionException("Macro could not be parsed.")
"""
