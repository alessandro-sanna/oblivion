import sys
import time
import os
import pyautogui
import psutil
import pywinauto
from copy import deepcopy
import win32gui


class InteractionManager:
    def __init__(self, target_file, program_info, enable_event):
        self.target_file = target_file
        self.program_info = program_info  # the ext_info dict
        self.current_process = None
        self.current_app = None

        self.log_folder = os.path.join("OblivionResources", "logs", "InteractionPlugin")
        self.log_file = os.path.join(self.log_folder, "interactions.txt")
        if os.path.exists(self.log_file):
            os.remove(self.log_file)
        self.log_fp = open(self.log_file, "w")

        self.enable_event = enable_event
        self.names_to_ignore = ["sandboxie", "outlook"]
        self.names_to_alert = ["visual basic"]

    def run(self):
        self.__set_office()
        window_list = self.__get_new_windows([])

        while self.enable_event.is_set():
            try:
                window_list += self.__get_new_windows(exclusion_list=window_list)
                window = window_list.pop(0)
                self.__manage_window(window)

            except IndexError:
                time.sleep(0.5)

        self.log_fp.close()

    def __set_office(self):
        queue = psutil.process_iter()
        while True:
            try:
                proc = next(queue)
            except StopIteration:
                time.sleep(0.5)
            else:
                if self.program_info["process_name"] in proc.name().upper():
                    self.current_process = proc
                    self.current_app = pywinauto.Application().connect(process=int(proc.pid))
                    break

    def __get_new_windows(self, exclusion_list):
        ppid = int(self.current_process.pid)
        windows_list = \
            pywinauto.findwindows.find_elements(class_name="#32770") + \
            pywinauto.findwindows.find_elements(parent=ppid)
        windows_list = [w for w in windows_list if w not in exclusion_list]
        return windows_list

    def __manage_window(self, window):
        win32gui.SetForegroundWindow(window.handle)
        time.sleep(0.1)
        self.__log(window)

        if not self.__preliminary_close(window):
            elements = deepcopy(window.children())
            for element in elements:
                self.__interact(element)
                if not self.__is_enabled(window):
                    break

    def __interact(self, elem):
        if elem.class_name == "Edit":
            self.__textbox_strategy(elem)
            return

        if elem.class_name == "Button":
            self.__button_strategy(elem)
            return
        # raise InteractionMgrException("I dont know this window type")

    def __button_strategy(self, elem):
        rect = self.current_app.window(handle=elem.handle).wrapper_object().client_rect()
        self.current_app.window(handle=elem.handle).wrapper_object().click_input(coords=(rect.right - 1, rect.bottom - 1))
        self.log_fp.write(f"[x] Clicked on Button {elem.name}\n")

    def __textbox_strategy(self, elem):
        textbox = self.current_app.window(handle=elem.handle).wrapper_object()
        textbox.set_edit_text("OBLIVION")  # gibberish, to replace
        self.log_fp.write("[x] Textbox edited\n")

    def __preliminary_close(self, window):
        name = window.name.lower()
        if any([n for n in self.names_to_ignore if n in name]):
            self.__close_window(window)
            return True
        if any([n for n in self.names_to_ignore if n in name]):
            self.__close_window(window)
            # should I close Office here?
            return True

        return False

    def __close_window(self, window):
        if self.__is_enabled(window):
            self.current_app.window(handle=window.handle).close()

    def __log(self, window):
        screen = pyautogui.screenshot()
        w_name = window.name if window.name != "" else "unnamed"
        file_name = os.path.basename(self.target_file)
        scr_path = os.path.join(self.log_folder, f"{file_name}+{w_name}+{window.handle}.png")
        screen.save(scr_path)
        self.log_fp.write(f"[x] Found window {w_name} with handle: {window.handle} and elements: {window.children()}\n")
        self.log_fp.flush()

    @staticmethod
    def __is_enabled(window):
        return pywinauto.findwindows.find_element(handle=window.handle).enabled


if __name__ == '__main__':
    target_file = r"C:\Users\diee\PycharmProjects\OblivionRef\OblivionTest\test_files\auto_both_test.docm"
    ext_info = {
    "main_class": "OpusApp",
    "bosa_class": "bosa_sdm_msword",
    "program": "word",
    "main_module": "document",
    "process_name": "WINWORD.EXE"
    }
    phase = InteractionManager(target_file, ext_info, True)
    phase.run()