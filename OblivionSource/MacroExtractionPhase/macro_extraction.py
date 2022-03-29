from oletools.olevba import VBA_Parser
import pickle
import os


class MacroExtractionException(Exception):
    pass


class MacroExtractor:
    def __init__(self, file_path, extensions):
        if file_path.split('.')[-1] in extensions:
            self.__office_file_path = file_path
        else:
            raise MacroExtractionException("The given file is not a recognized office file")
        self.__macro_list = list()
        self.__extract_macro()

    def run(self):
        return "".join(self.__macro_list)

    def __extract_macro(self):
        vb_parser = VBA_Parser(self.__office_file_path)

        try:
            vb_parser.extract_all_macros()
        except IndexError:
            raise MacroExtractionException(u"Macro(s) found, but they were impossible to extract.")

        if vb_parser.detect_vba_macros():
            for (filename, stream_path, vba_filename, vba_code) in \
                    vb_parser.extract_macros():
                self.__parse_macro(vba_filename, vba_code)
            vb_parser.close()
            return
        else:
            vb_parser.close()
            raise MacroExtractionException(u"No VBA macro found.")

    @staticmethod
    def __parse_macro(vba_filename, vba_code):
        code = "\n".join([line for line in vba_code.splitlines()
                          if "Attribute VB_" not in line]).strip() + "\n"
        macro_data = [vba_filename, code]

        with open(os.path.join("OblivionResources", "data", "original_macro_data.pkl", "wb")) as pkl:
            pickle.dump(macro_data, pkl)

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
