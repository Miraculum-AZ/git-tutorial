import pyautogui
import time
# --Covering GUI and all GUI-related tasks
# GUI creation
import tkinter as tk
from tkinter import filedialog as fd
from tkinter import ttk
import customtkinter as ctk
from Settings import *
from threading import Timer, Thread

class App(ctk.CTk):
    def __init__(self, size):
        # window setup
        super().__init__(fg_color="#c1c9c3")
        ctk.set_appearance_mode("dark")
        # setting main window attributes
        self.title("Keep the screen running")
        try:
            self.iconbitmap("null.ico")
        except:
            print("No logo yet :(")
        # spawning our GUI in the middle of the screen
        self.geometry(
            f"{size[0]}x{size[1]}+{round(self.winfo_screenwidth() / 2 - 150)}+{round(self.winfo_screenheight() / 2 - 150)}"
        )
        # limiting its size
        self.resizable(False, False)
        # --Create layout
        self.btn_1 = Generic_Button(self, text_var="Keep alive")
        self.btn_1.configure(command=lambda: self.thread_for_unlock())
        self.btn_1.pack(fill="both", padx=(15, 15), pady=(3, 0))

        self.btn_2 = Generic_Button(self, text_var="S.t.o.o.o.p.")
        self.btn_2.configure(command=lambda: self.stop())
        self.btn_2.pack(fill="both", padx=(15, 15), pady=(3, 0))
        # main variable
        self.stay_onguard = True
        # run the GUI
        self.mainloop()

    def keep_alive(self):
        while True:
            if self.stay_onguard == True:
                pyautogui.press('volumedown')
                time.sleep(300)
                pyautogui.press('volumeup')
                time.sleep(300)
            else: 
                break
        
    def stop(self):
        self.stay_onguard = False

    def thread_for_unlock(self):
        self.stay_onguard = True
        self.t_fin = Thread(target=self.keep_alive)
        self.t_fin.start()
        

class Generic_Button(ctk.CTkButton):
    def __init__(self, parent, text_var):
        super().__init__(
            parent,
            text=text_var,
            corner_radius=5,
            border_width=1,
            border_spacing=5,
            fg_color=ENTRY_NORMAL,
            hover_color=BUTTON_HOVER,
            border_color=BUTTON_BORDER,
            height=40,
            font=("", 13, "bold"),
            anchor="center",
        )


App((200,95))