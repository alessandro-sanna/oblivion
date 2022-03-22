import sys


class OfficeSandboxException(Exception):
    pass


class OfficeSandbox:
    def __init__(self, running_file, instrumented_code_path, log_file,
                 program, main_class, main_module,
                 auto_open, auto_close, no_clean_slate_flag):
        self.running_file = running_file
        self.extension = running_file.split('.')[-1]
        self.instrumented_code_path = instrumented_code_path

        self.log_file = log_file
        log_fp = open(self.log_file, "w")
        sys.stdout = log_fp
        sys.stderr = log_fp

        self.program = program
        self.main_class = main_class
        self.main_module = main_module

        self.auto_open, self.auto_close = auto_open, auto_close
        self.no_clean_slate_flag = no_clean_slate_flag

        
    def run(self):
        pass

    def __open_file(self, file_name):
        pass


if __name__ == '__main__':
    running_file = sys.argv[1]
    instrumented_code_path = sys.argv[2]
    program, main_class, main_module = (x for x in sys.argv[3: 6])
    auto_open, auto_close, no_clean_slate_flag = (bool(int(x)) for x in sys.argv[6: 9])
    log_file = sys.argv[9]
    office_sbx_obj = OfficeSandbox(running_file, instrumented_code_path, log_file,
                                   program, main_class, main_module,
                                   auto_open, auto_close, no_clean_slate_flag)
    