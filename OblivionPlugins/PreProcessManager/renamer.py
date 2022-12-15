import sys
import os
import glob


if __name__ == '__main__':
    magic = b"\xD0\xCF\x11\xE0"
    rf = sys.argv[1]
    file_list = glob.glob(os.path.join(rf, "**"), recursive=True)
    total = len(file_list)
    not_conv = 0

    for index, file in enumerate(file_list):
        if file.endswith("x"):
            with open(file, "rb") as fp:
                this_magic = fp.read(4)
            if this_magic == magic:
                os.rename(file, file[:-1])
            else:
                not_conv += 1
                # print(file, this_magic)
        print(f"\r{index+1} / {total}; {not_conv} nc", end='')
