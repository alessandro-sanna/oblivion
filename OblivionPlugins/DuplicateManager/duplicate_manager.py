import os
import sys
import glob
import argparse
import tempfile
import hashlib
import pickle
import json
from collections import OrderedDict
from OblivionSource.MacroExtractionPhase import MacroExtraction, MacroExtractionException


class DuplicateManagerException(Exception):
    pass


class DuplicateManager:
    def __init__(self, target_folder, extensions, output_file):
        self.file_list = glob.glob(os.path.join(target_folder, "**", "*"), recursive=True)
        self.extensions = extensions
        self.output_file = output_file
        self.output_dict = dict()
        self.hashes_dict = dict()

    def run(self):
        self.__get_hashes()

        for file_name, sha256_hash in self.hashes_dict.items():
            if sha256_hash not in self.output_dict.keys():
                self.output_dict.update({sha256_hash: [file_name]})
            else:
                self.output_dict[sha256_hash] += [file_name]

        with open(self.output_file, "w") as foJson:
            json.dump(self.output_dict, foJson, indent=4)

    def __get_hashes(self):
        for file_name in self.file_list:
            try:
                macro_code = self.__standardize(self.__get_macro_dict(file_name))
            except DuplicateManagerException as exc:
                print(exc)
                continue

            with tempfile.TemporaryFile() as foTmp:
                pickle.dump(macro_code, foTmp)
                sha256_hash = self.__get_sha256(foTmp)

            self.hashes_dict.update({file_name: sha256_hash})
            print(f"[+] {file_name} ok: got {sha256_hash}")

    def __get_macro_dict(self, file):
        try:
            macro_dict = MacroExtraction(file, self.extensions).run()
        except MacroExtractionException as exc:
            raise DuplicateManagerException(f"[!] extraction from {file} failed: {exc}")

        return macro_dict

    @staticmethod
    def __standardize(macro_dict):
        macro_dict = {k.lower(): v.lower() for (k, v) in macro_dict.items()}
        return OrderedDict(sorted(macro_dict.items(), key=lambda t: t[0]))

    @staticmethod
    def __get_sha256(file_object, buf_size=65536):
        sha256 = hashlib.sha256()
        file_object.seek(0, 0)

        while data := file_object.read(buf_size):
            sha256.update(data)

        return sha256.hexdigest()


def parse_args():
    parser = argparse.ArgumentParser(prog="dupmacro.py", prefix_chars="-")
    parser.add_argument("-d", "--target_directory", nargs='?', type=str, required=True,
                        help="path of the directory where the files to be filtered are stored")
    parser.add_argument("-e", "--extensions", nargs='?', type=str, default="doc,docm,xls,xlsm",
                        help="comma-separated list of extensions to be considered")
    parser.add_argument("-o", "--output_file", nargs='?', type=str, required=True,
                        help="json file path where to save results")
    return parser.parse_args(sys.argv[1:])


if __name__ == '__main__':
    args = parse_args()
    duplicate_obj = DuplicateManager(args.target_directory, args.extensions.split(','), args.output_file)
    duplicate_obj.run()
