import argparse
import sys
import os
from OblivionSource import OblivionCore, PreProcessing


class OblivionException(Exception):
    pass


class Oblivion:
    def __init__(self):
        self.args = self.__parse_args()
        self.__validate_args()

    @staticmethod
    def __parse_args():
        parser = argparse.ArgumentParser(prog="oblivion.py", prefix_chars="-")
        exclusive_path_mode = parser.add_mutually_exclusive_group(required=True)

        exclusive_path_mode.add_argument("-f", "--file", nargs='?', type=str, dest="target_file",
                                         help="path of a single to-be-analyzed file, cannot be used with -d")
        exclusive_path_mode.add_argument("-d", "--directory", nargs='?', type=str, dest="target_directory",
                                         help="path of a directory of to-be-analyzed files, cannot be used with -f")
        parser.add_argument("-o", "--output", nargs='?', type=str, dest="output_directory", required=True,
                            help="path of the directory where Oblivion will save the report file")
        parser.add_argument("-t", "--time_limit", nargs='?', type=float, default=99999.0,
                            help="maximum time per single analysis")
        parser.add_argument("-p", "--preprocessing", nargs='?', choices={"static", "dynamic"},
                            help="if set, activate pre-processing mode")
        parser.add_argument("-mdb", "--use_mongo_db", action="store_true",
                            help="if set, save reports in a mongo database")
        parser.add_argument("-ncs", "--no_clean_slate", action="store_true",
                            help="if set, inject instrumentation in file as it is")
        parser.add_argument("-im", "--use_interaction_manager", action="store_true",
                            help="if set, run the interaction manager plug-in")
        parser.add_argument("-noi", "--no_office_instance", action="store_true",
                            help="if set, try to execute only VBS code without Office")

        if len(sys.argv) == 1:
            print("No argument supplied\n")
            parser.print_help()
            exit(0)

        return parser.parse_args(sys.argv[1:])

    def __validate_args(self):
        if self.args.target_file is not None and not os.path.isfile(self.args.target_file):
            raise OblivionException("Error! -f must be a file path.")
        if self.args.target_directory is not None and not os.path.isdir(self.args.target_directory):
            raise OblivionException("Error! -d must be a folder path.")
        if self.args.output_directory is not None and not os.path.isfile(self.args.output_directory):
            raise OblivionException("Error! -o must be a folder path.")

    def is_file_run(self):
        return self.args.target_file is not None

    def is_dir_run(self):
        return self.args.target_directory is not None

    def is_preprocessing_run(self):
        return self.args.preprocessing is not None


if __name__ == '__main__':
    oblivion_obj = Oblivion()
    normal_execution = oblivion_obj.is_file_run() ^ oblivion_obj.is_dir_run()

    if normal_execution:
        core_obj = OblivionCore(oblivion_obj.args)
        core_obj.execute(single=oblivion_obj.is_file_run())

    elif oblivion_obj.is_preprocessing_run():
        prep_obj = PreProcessing(oblivion_obj.args, oblivion_obj.args.preprocessing.capitalize())
        prep_obj.execute()

    else:
        raise ValueError("Options not recognized, please use -h for usage.")
