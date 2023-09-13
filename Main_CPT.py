#--Covering GUI and all GUI-related tasks
# GUI creation
import tkinter as tk
from tkinter import filedialog as fd 
from tkinter import ttk
import customtkinter as ctk
# flexible dates
from datetime import date, timedelta, datetime, timezone
import calendar
# validation
import re
# custom colors
try:
    from Settings import *
except: pass
# custom SAP library
try:
    from ok_sap_script import *
except: pass
# calculations and database
import pandas as pd
import numpy as np
# memory file (csv manipulations)
import csv
# outlook
import win32com.client as win32
# get current user
import os
#--To extract network drives:
import wmi

#--Covering BI and SAP part
import subprocess
import xlwings as xw
from xlwings.constants import FileFormat
from threading import Timer, Thread
import shutil # for file manipulation
#from ok_sap_script import *

#--Handling
import time 
import logging

class App(ctk.CTk):
    def __init__(self, size):
        # window setup
        super().__init__(fg_color="#c1c9c3")
        ctk.set_appearance_mode("dark")
        # setting global variables
        global version
        version = "0.9"
        self.curr_user = os.getlogin()

        # main variables
        self.default_path_file_dialogue = f"C:/Users/{self.curr_user}/Desktop/Automations/Freight_Accruals"
        # setting main window attributes
        self.title("Select GUI")
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
        #--Populate main variables
        self.populate_vars()
        # --Create layout
        self.create_layout()
        #--Events
        self.bind('<Control-i>', lambda event: self.binding_instructions())
        self.toplevel_instructions = None
        self.toplevel_monthly_emails = None
        self.toplevel_misalignments = None
        # run the GUI
        self.mainloop()

    def binding_instructions(self):
        if self.toplevel_instructions is None or not self.toplevel_instructions.winfo_exists():
            self.toplevel_instructions = Top_Level_Window((490,280), "Instructions")
        else:
            self.toplevel_instructions.focus()

    def populate_vars(self):
        #--Overall
        self.curr_user = os.getlogin()
        self.environment = "ECC (QE3)"
        #--Price determination
        self.receiver = "oleksandr.komarov@zoetis.com"
        self.extract_path = f"C:/Users/{self.curr_user}/Desktop/Automations/CPT_TP/Price_determination/Output"
        self.extract_name = f"PD-{datetime.today().strftime('%d-%m-%Y')}.xlsx"
        self.t_code_sap = "se16"
        self.table_sap = "mbew"
        self.variant_sap = "MBEW_PD"
        self.variant_sap_check = "/MBEW_NB"
        #--Monthly emails

        #--TP misalignment

#--LAYOUT block----------------------
    def create_layout(self):
        self.create_elements()
        self.pack_elements()
        self.configure_elements()
             
    def create_elements(self):
        # --Creating the elements
            # frame
        self.frame_final = ctk.CTkFrame(self, border_width=4, border_color="#275936", fg_color=LABEL_FG)
            # label
        self.label_info = Generic_Label(self.frame_final, name="Select required control")
            # separators
        self.separator1 = ttk.Separator(self.frame_final, orient="horizontal")
        self.separator2 = ttk.Separator(self.frame_final, orient="horizontal")
        self.separator3 = ttk.Separator(self.frame_final, orient="horizontal")
        self.separator4 = ttk.Separator(self.frame_final, orient="horizontal")
            # buttons
        self.btn_pd = Generic_Button(self.frame_final, text_var="Request Price Determination")
        self.btn_me = Generic_Button(self.frame_final, text_var="Open monthly emails control")
        self.btn_tm = Generic_Button(self.frame_final, text_var="Open TP misalignment control")
            # rights
        self.rights = ctk.CTkLabel(self.frame_final, text="Author:\n© 2023, Oleksandr Komarov\noleksandr.komarov@zoetis.com", justify='left', font=('', 9), anchor='sw', text_color='#000000')

    def pack_elements(self):
        #--Putting elements on the frame
            # pack label
        self.label_info.pack(fill="both", padx=3, pady=(3, 0))
        self.separator1.pack(fill="x", padx=20, pady=(0, 3))
            # pack buttons
        self.btn_pd.pack(fill="both", padx=(15, 15), pady=(3, 0))
        self.separator2.pack(fill="x", padx=20, pady=(3, 3))
        self.btn_me.pack(fill="both", padx=(15, 15), pady=(3, 0))
        self.separator3.pack(fill="x", padx=20, pady=(3, 3))
        self.btn_tm.pack(fill="both", padx=(15, 15), pady=(3, 0))    
        self.separator4.pack(fill="x", padx=20, pady=(3, 10))
            # pack the frame
        self.frame_final.pack(fill="both", padx=3, pady=3)
            # pack rights
        self.rights.pack(fill='x', padx=20, pady=(0,10))

    def configure_elements(self):
        self.btn_pd.configure(command=lambda:self.request_price_determination())
        self.btn_me.configure(command=lambda:self.open_monthly_emails())
        self.btn_tm.configure(command=lambda:self.open_tp_misalignment())
#--End LAYOUT block------------------

#--Button commands App---------------
    def request_price_determination(self):
        #ToDo
        # Generate an email request:

        # below is the code to be incorporated into orchestrator
        # ensure there are no open sessions
        sap_close()
        # open new session
        open_sap = sap_open()
        time.sleep(5)
        if open_sap == True:
            session = sap_logon(environment=self.environment, client=1)
            sap_code(tcode=self.t_code_sap, session=session)
            # enter mbew
            session.findById("wnd[0]/usr/ctxtDATABROWSE-TABLENAME").text = self.table_sap
            session.findById("wnd[0]").sendVKey(0)
            session.findById("wnd[0]/usr/ctxtI1-LOW").text = "10000000"
            session.findById("wnd[0]/usr/ctxtI1-HIGH").text = "19999999"
            session.findById("wnd[0]/usr/txtMAX_SEL").text = "10000000"
            session.findById("wnd[0]").sendVKey(8)

            # select layout
            sap_layout(session=session, layout=self.variant_sap)
            # check if there is no output
            mbew_body = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell")
            rows = mbew_body.rowCount
            if rows !=0:
                print("There is some data. Export")
                # extract
                sap_extract(session=session, extr_path=self.extract_path, extr_name=self.extract_name)
                close_excel()
                # send by email
                outlook = win32.Dispatch('outlook.application')
                mail = outlook.CreateItem(0)
                mail.To = self.receiver
                mail.Subject = f"Automatic email for price determination control"
                mail.Body = f"Please consult the attachment"
                mail.Attachments.Add(f"{self.extract_path}/{self.extract_name}")
                #mail.Display(True)
                mail.Send()
                sap_close()
            else:
                print("There is no data. End here")
                outlook = win32.Dispatch('outlook.application')
                mail = outlook.CreateItem(0)
                mail.To = self.receiver
                mail.Subject = f"Automatic email for price determination control"
                mail.Body = f"There are no materials to be corrected"
                #mail.Display(True)
                mail.Send()
                sap_close()

    def open_monthly_emails(self):
        if self.toplevel_monthly_emails is None or not self.toplevel_monthly_emails.winfo_exists():
            self.toplevel_monthly_emails = Top_Level_Window((200,300), "Monthly emails")
        else:
            self.toplevel_monthly_emails.focus()

    def open_tp_misalignment(self):
        if self.toplevel_misalignments is None or not self.toplevel_misalignments.winfo_exists():
            self.toplevel_misalignments = Top_Level_Window((300,180), "Misalignments")
        else:
            self.toplevel_misalignments.focus()

#--End Button commands App-----------

class Top_Level_Window(ctk.CTkToplevel):
    def __init__(self, size, title, resizable = False):
        super().__init__(fg_color="#c1c9c3", takefocus=True)
        # assign to self
        self.curr_user = os.getlogin()
        self.email = 'oleksandr.komarov@zoetis.com'
        self.size = size
        self.title_app = title
        # apply
        self.geometry(f"{self.size[0]}x{self.size[1]}")
        self.title(self.title_app)
        ctk.set_appearance_mode("dark")
        self.iconbitmap(default='null.ico')
        self.resizable(resizable, resizable)

        if self.title_app == "Instructions":
            self.create_layout_instructions()
        elif self.title_app == "Monthly emails":
            self.create_layout_monthly_emails()
        elif self.title_app == "Misalignments":
            self.create_layout_misalignments()
        else:
            print("Unknown window")

    def create_layout_instructions(self):
        # create a frame
        self.frame_instructions = ctk.CTkScrollableFrame(self, width=200, height=400, border_width=4, 
                                                         border_color="#275936", fg_color=LABEL_FG, scrollbar_button_color=BUTTON_BORDER, scrollbar_button_hover_color="#275936")
        
        ctk.CTkLabel(self.frame_instructions, 
                     text="Below are the instructions to this automation", 
                     font=('', 15, 'underline'), 
                     justify="center",
                     fg_color="#1f472b").pack(fill="both", pady=0)
        self.separator = ttk.Separator(self.frame_instructions, orient='horizontal')
        self.separator.pack(fill='x', pady=1)
        ctk.CTkLabel(self.frame_instructions, text='Make sure the correct folder structure is in place.', font=('', 13, 'bold'), justify="left").pack(fill="both")
        ctk.CTkLabel(self.frame_instructions, text='Consider the following example: Desktop\Automations\Freight_Accruals.', font=('', 13), justify="left").pack(fill="both")
        ctk.CTkLabel(self.frame_instructions, text='Put your model files under this location (2941, 2946, 2116).', font=('', 13), justify="left").pack(fill="both")
        ctk.CTkLabel(self.frame_instructions, text='Put the received mail (.msg) under the same location.', font=('', 13), justify="left").pack(fill="both")
        self.separator = ttk.Separator(self.frame_instructions, orient='horizontal')
        self.separator.pack(fill='x', pady=1)        
        ctk.CTkLabel(self.frame_instructions, text='Make sure the TRAX accruals file is saved as Excel workbook .xlsx', font=('', 13)).pack(fill="both")
        ctk.CTkLabel(self.frame_instructions, text='If it saved as Strict Open XML Spreadsheet .xlsx, please resave it', font=('', 13)).pack(fill="both")
        ctk.CTkLabel(self.frame_instructions, text='This is very important for the automation to work', font=('', 13)).pack(fill="both")
        ctk.CTkLabel(self.frame_instructions, text='It is assumed that the format of this file does not change', font=('', 13)).pack(fill="both")
        ctk.CTkLabel(self.frame_instructions, text='It is supposed to have "BE" sheet and the columns are in the same order', font=('', 13)).pack(fill="both")
        ctk.CTkLabel(self.frame_instructions, text='They need to be named consistently, for the columns are defined by names', font=('', 13)).pack(fill="both")
        self.separator = ttk.Separator(self.frame_instructions, orient='horizontal')
        self.separator.pack(fill='x', pady=1) 
        ctk.CTkLabel(self.frame_instructions, text='Select the files by clicking the buttons', font=('', 13, 'bold')).pack(fill="both")
        ctk.CTkLabel(self.frame_instructions, text='You can select for which CoCd the reports have to be generated', font=('', 13, 'bold')).pack(fill="both")
        ctk.CTkLabel(self.frame_instructions, text='After that click on "Populate accruals file(s)" button', font=('', 13)).pack(fill="both")
        ctk.CTkLabel(self.frame_instructions, text='The files will be created under \Freight_Accruals\Output folder', font=('', 13)).pack(fill="both")
        ctk.CTkLabel(self.frame_instructions, text='After that you are free to rename/change/move them', font=('', 13)).pack(fill="both")
        self.separator = ttk.Separator(self.frame_instructions, orient='horizontal')
        self.separator.pack(fill='x', pady=1) 
        ctk.CTkLabel(self.frame_instructions, text='If the execution fails, please contact oleksandr.komarov@zoetis.com', font=('', 13)).pack(fill="both")
        ctk.CTkLabel(self.frame_instructions, text='Thank you for reading this', font=('', 13)).pack(fill="both")
        ctk.CTkLabel(self.frame_instructions, text='I hope this automation will make your life easier :)', font=('', 13)).pack(fill="both")

        self.frame_instructions.pack(fill="both", padx=3, pady=2)

    def create_layout_monthly_emails(self):
        # default path for filedialogue
        self.default_path_file_dialogue_emails = f"C:/Users/{self.curr_user}/Desktop/Automations/CPT_TP/Monthly_emails"
        # define path to email file
        self.final_path_to_emails = None
        self.check_path_to_emails = f"C:/Users/{self.curr_user}/Desktop/Automations/CPT_TP/Monthly_emails/Contacts_emails_ME.xlsx"
        # for flash and non-flash files
        self.final_path_to_flash = None
        self.final_path_to_non_flash = None
        # create a frame
        self.frame_emails_remote = Generic_Frame(self)
        # create one label
        label_remote = Generic_Label(self.frame_emails_remote, "Execute control remotely").pack(fill="both", padx=3, pady=(3, 0))
        # add a separator
        Generic_Separator(self.frame_emails_remote).pack(fill="x", padx=20, pady=(0, 3))
        # create one button
        self.btn_run_remotely = Generic_Button(self.frame_emails_remote, "Send request")
        self.btn_run_remotely.configure(command=lambda:self.monthly_emails_run_remotely())
        self.btn_run_remotely.configure(fg_color="#124e78")
        self.btn_run_remotely.configure(hover_color="#56789a")
        self.btn_run_remotely.pack(fill="both", padx=3, pady=(3, 3))
        # pack the frame        
        self.frame_emails_remote.pack(fill="both", padx=3, pady=2)

        # create a frame
        self.frame_emails_local = Generic_Frame(self)
        # add one label
        label_local = Generic_Label(self.frame_emails_local, "Execute control locally").pack(fill="both", padx=3, pady=(3, 0))
        # add a separator
        Generic_Separator(self.frame_emails_local).pack(fill="x", padx=20, pady=(0, 3))
        # add four buttons (3 file dialogues, 1 execute button)
            #1
        self.btn_select_contacts = Generic_Button(self.frame_emails_local, "Select contacts file ✗")
        self.btn_select_contacts.configure(command=lambda:self.select_contacts_file())
        self.btn_select_contacts.pack(fill="both", padx=3, pady=(3, 0))
        self.confirm_path_to_emails = os.path.exists(self.check_path_to_emails)
        if self.confirm_path_to_emails == True:
            self.final_path_to_emails = self.check_path_to_emails
            self.btn_select_contacts.configure(text="Select contacts file ✓")
            #2
        self.btn_select_flash = Generic_Button(self.frame_emails_local, "Select flash file ✗")
        self.btn_select_flash.configure(command=lambda:self.select_flash_file())
        self.btn_select_flash.pack(fill="both", padx=3, pady=(3, 0))
            #3
        self.btn_select_non_flash = Generic_Button(self.frame_emails_local, "Select non-flash file ✗")
        self.btn_select_non_flash.configure(command=lambda:self.select_non_flash_file())
        self.btn_select_non_flash.pack(fill="both", padx=3, pady=(3, 0))
        # add a separator
        Generic_Separator(self.frame_emails_local).pack(fill="x", padx=20, pady=(3, 3))
            #4
        self.btn_run_locally = Generic_Button(self.frame_emails_local, "Run locally")
        self.btn_run_locally.configure(command=lambda:self.monthly_emails_run_locally())
        self.btn_run_locally.configure(fg_color="#124e78")
        self.btn_run_locally.configure(hover_color="#56789a")
        self.btn_run_locally.pack(fill="both", padx=3, pady=(3, 3))

        # pack the frame        
        self.frame_emails_local.pack(fill="both", padx=3, pady=2)

    def create_layout_misalignments(self):
        #--Define paths
            # default path for filedialogue
        self.default_path_file_dialogue_misalignments = f"C:/Users/{self.curr_user}/Desktop/Automations/CPT_TP/Misalignments"
            # define path to previous file
        self.final_path_to_prev_misalignments = None
            # define path to current file
        self.final_path_to_curr_misalignments = None
        #--Create layout
            # create a frame
        self.frame_misalignments_main = Generic_Frame(self)
            # populate elements
                # add a label
        Generic_Label(self.frame_misalignments_main, "Do not use for SOX").pack(fill="both", padx=3, pady=(3, 0))
                # add a separator
        Generic_Separator(self.frame_misalignments_main).pack(fill="x", padx=20, pady=(0, 3))  
                # add three buttons
                    #1
        self.btn_select_prev_file = Generic_Button(self.frame_misalignments_main, "Select previous file ✗")
        self.btn_select_prev_file.configure(command=lambda:self.select_prev_file())
        self.btn_select_prev_file.pack(fill="both", padx=3, pady=(3, 0))
                    #2
        self.btn_select_curr_file = Generic_Button(self.frame_misalignments_main, "Select current file ✗")
        self.btn_select_curr_file.configure(command=lambda:self.select_curr_file())
        self.btn_select_curr_file.pack(fill="both", padx=3, pady=(3, 0))
        # add a separator
        Generic_Separator(self.frame_misalignments_main).pack(fill="x", padx=20, pady=(3, 3))
                    #3
        self.btn_run_misalignments = Generic_Button(self.frame_misalignments_main, "Run control")
        self.btn_run_misalignments.configure(command=lambda:self.misalignments_run_locally())
        self.btn_run_misalignments.configure(fg_color="#124e78")
        self.btn_run_misalignments.configure(hover_color="#56789a")
        self.btn_run_misalignments.pack(fill="both", padx=3, pady=(3, 3)) 
            # pack the frame        
        self.frame_misalignments_main.pack(fill="both", padx=3, pady=2)

    def select_prev_file(self):
        self.final_path_to_prev_misalignments = fd.askopenfilename(
            initialdir=self.default_path_file_dialogue_misalignments,
            title="Choose previous file",
            parent=self,
            filetypes=(("xlsx files", "*.xlsx"),("xlsm files", "*.xlsb"),("xlsm files", "*.xlsm"),),)
        if self.final_path_to_prev_misalignments == "" or len(self.final_path_to_prev_misalignments) < 1:
            self.final_path_to_prev_misalignments = None
            self.btn_select_prev_file.configure(text="Select previous file ✗")
            print("No path was provided")
        else:
            self.btn_select_prev_file.configure(text="Select previous file ✓")
            self.btn_select_prev_file.configure(state="normal")
            print(self.final_path_to_prev_misalignments)

    def select_curr_file(self):
        self.final_path_to_curr_misalignments = fd.askopenfilename(
            initialdir=self.default_path_file_dialogue_misalignments,
            title="Choose previous file",
            parent=self,
            filetypes=(("xlsx files", "*.xlsx"),("xlsm files", "*.xlsb"),("xlsm files", "*.xlsm"),),)
        if self.final_path_to_curr_misalignments == "" or len(self.final_path_to_curr_misalignments) < 1:
            self.final_path_to_curr_misalignments = None
            self.btn_select_curr_file.configure(text="Select current file ✗")
            print("No path was provided")
        else:
            self.btn_select_curr_file.configure(text="Select current file ✓")
            self.btn_select_curr_file.configure(state="normal")
            print(self.final_path_to_curr_misalignments)

    def misalignments_run_locally(self):
        if (self.btn_select_prev_file.cget("text") == "Select previous file ✓" 
            and self.btn_select_curr_file.cget("text") == "Select current file ✓"):
            print(self.btn_select_prev_file, self.btn_select_curr_file)
            # run the control
            self.c = wmi.WMI ()
            self.drives = []
            for drive in self.c.Win32_LogicalDisk ():
                # prints all the drives details including name, type and size
                #print(drive)
                self.drives.append(drive.Caption[0])
            print (self.drives[1])#, drive.VolumeName, DRIVE_TYPES[drive.DriveType])
            #--Time stamp:
            self.timestr = time.strftime("%Y_%m_%d-%H_%M_%S")
        #--SAP dump -> slightly reworked:
            self.full_file_path = f"C:/Users/{self.curr_user}/Desktop/Automations/CPT_TP/Misalignments/Output/{self.timestr}"
            self.full_file_path_SD = f"{self.drives[1]}:\FINBEL\BEP\Pricing\Python_Controls\TP_Misalignment/{self.timestr}"
            os.mkdir(self.full_file_path)

        #--BLOCK_FOR_SD----------------------------------------------------------------------------
            try:
                os.mkdir(self.full_file_path_SD)
                self.full_file_SD = f"{self.full_file_path_SD}/Full_File_For_{self.timestr}.xlsx"
                self.final_file_SD = f"{self.full_file_path_SD}/Misaligned_File_For_{self.timestr}.xlsx"
            except: 
                #--Use logging here:
                try:
                    self.alternative_path = f"A:\BEP\Pricing\Python_Controls\TP_Misalignment/{self.timestr}"
                    os.mkdir(self.alternative_path)
                    self.full_file_SD = f"{self.alternative_path}/Full_File_For_{self.timestr}.xlsx"
                    self.final_file_SD = f"{self.alternative_path}/Misaligned_File_For_{self.timestr}.xlsx"
                except:
                    None
        #------------------------------------------------------------------------------------------

            self.full_file = f"{self.full_file_path}/Full_File_For_{self.timestr}.xlsx"
            print(self.full_file)
        #--File with the misalignments:
            self.final_file = f"{self.full_file_path}/Misaligned_File_For_{self.timestr}.xlsx"
            print(self.final_file)
        #--------------------

        # upload previous_file to pandas:
            df_prev = pd.read_excel(self.final_path_to_prev_misalignments)
        # upload current SAP dump to pandas:
            df = pd.read_excel(self.final_path_to_curr_misalignments)
        # perform operations in pandas
            df = df[df["Unnamed: 2"].notna()] # drop NaN
            df.columns = df.iloc[0] # promote headers
            df = df[df['Material'] != "Material"] # remove all rows that mention "Material"
            df = df.drop(df.columns[[0, 3, 4]],axis = 1) # drop useless columns
        # export this file as full file
            df.to_excel(self.full_file, index = False)
            try:
                df.to_excel(self.full_file_SD, index = False)
            except:
                try:
                    df.to_excel(self.alternative_path, index = False)
                except:
                    None
        # filter all aligned items:
            df = df[df['Comment for USD'] != "OK"]
        # add column "Rounding" with calculated logic to track aligned items with rounding:
            df['Rounding?'] = np.where( ( (df['Sales Price @ Plan – USD']  - df['Purchase Price @ Plan – USD'] <=0.02)
            & (df['Sales Price @ Plan – USD'] - df['Purchase Price @ Plan – USD'] >=-0.02)
            & (df['Sales Price @ Plan – USD'] - df['Standard Price @ Plan – USD']<=0.02)
            & (df['Sales Price @ Plan – USD'] - df['Standard Price @ Plan – USD']>=-0.02)
        ), "Rounding", "Check")
        # filter roundings out:
            df = df[df['Rounding?'] != "Rounding"]
        # add concatenation column:
            df['Concatenate'] = df['Material'].astype(str) + df['Supplying Plant'] + df['Receiving Plant']
        # perform vlookup based on the "Concanetation" column:
            df = pd.merge(df,df_prev[['Concatenate','Comments']],on='Concatenate', how='left')
        # save final file locally and on the shared drive:
            df.to_excel(self.final_file, index = False)
        #--Save final file on the drive:
            try:
                df.to_excel(self.final_file_SD, index = False)
            except:
                try:
                    df.to_excel(self.alternative_path, index = False)
                except:
                    None
        else:
            print("Data is missing")   

    def select_contacts_file(self):
        self.final_path_to_emails = fd.askopenfilename(
            initialdir=self.default_path_file_dialogue_emails,
            title="Choose contacts file",
            parent=self,
            filetypes=(("xlsx files", "*.xlsx"),("xlsm files", "*.xlsb"),("xlsm files", "*.xlsm"),),)
        # self.focus() because we need to lift it after file dialogue -> mitigated by adding parent above
        if self.final_path_to_emails == "" or len(self.final_path_to_emails) < 1:
            self.final_path_to_emails = None
            self.btn_select_contacts.configure(text="Select contacts file ✗")
            # self.check_all_ticks_final()
            print("No path was provided")
        else:
            self.btn_select_contacts.configure(text="Select contacts file ✓")
            self.btn_select_contacts.configure(state="normal")
            print(self.final_path_to_emails)
        
    def select_flash_file(self):
        self.final_path_to_flash = fd.askopenfilename(
            initialdir=self.default_path_file_dialogue_emails,
            title="Choose flash file",
            parent=self,
            filetypes=(("xlsx files", "*.xlsx"),("xlsm files", "*.xlsb"),("xlsm files", "*.xlsm"),),)
        if self.final_path_to_flash == "" or len(self.final_path_to_flash) < 1:
            self.final_path_to_flash = None
            self.btn_select_flash.configure(text="Select flash file ✗")
            # self.check_all_ticks_final()
            print("No path was provided")
        else:
            self.btn_select_flash.configure(text="Select flash file ✓")
            self.btn_select_flash.configure(state="normal")
            print(self.final_path_to_flash)
        
    def select_non_flash_file(self):
        self.final_path_to_non_flash = fd.askopenfilename(
            initialdir=self.default_path_file_dialogue_emails,
            title="Choose non flash file",
            parent=self,
            filetypes=(("xlsx files", "*.xlsx"),("xlsm files", "*.xlsb"),("xlsm files", "*.xlsm"),),)
        if self.final_path_to_non_flash == "" or len(self.final_path_to_non_flash) < 1:
            self.final_path_to_non_flash = None
            self.btn_select_non_flash.configure(text="Select non flash file ✗")
            # self.check_all_ticks_final()
            print("No path was provided")
        else:
            self.btn_select_non_flash.configure(text="Select non flash file ✓")
            self.btn_select_non_flash.configure(state="normal")
            print(self.final_path_to_non_flash)
        
    def monthly_emails_run_remotely(self):
        # create outlook request
        '''
        Here is the code that you'll need to pass into orchestrator:
        # This function is for testing TP montly emails in SAP
        # Test SAP part

        from ok_sap_script import *
        import pandas as pd
        import numpy as np
        from datetime import datetime
        import time
        import os
        import win32com.client as win32

        curr_user = os.getlogin()
        pi01_path = f"pi01-{datetime.today().strftime('%d-%m-%Y')}.xlsx"
        ziv1_path = f"ziv1-{datetime.today().strftime('%d-%m-%Y')}.xlsx"

        def sap_extract_me():   
            first_day = datetime.today().replace(day=1).strftime("%d.%m.%Y")
            environment = "ECC (QE3)"
            t_code_sap = "ZP2M_OTBD_SP_FL_NONF"
            variant_flash = "FLASH_ME"
            variant_non_flash = "NON_FLASH_ME" 
            

            # ensure there are no open sessions
            close_sap = sap_close()
            # open new session
            open_sap = sap_open()
            time.sleep(5)
            if open_sap == True:
                session = sap_logon(environment=environment, client=1)
                sap_code(tcode=t_code_sap, session=session)
                try:
                    sap_variant(session=session, var_to_use=variant_flash)
                except: 
                    sap_variant(session=session, var_to_use=variant_flash, version=2) # in our case, V2 is the most probable scenario

                session.findById("wnd[0]/usr/ctxtS_DATBI-LOW").text = first_day
                # run the report
                sap_run(session=session)
                sap_extract(session=session, extr_path="C:\\Users\\KOMAROVO\\Desktop\\Automations\\CPT_TP\\Monthly_emails\\Output\\Files_extracted\\", extr_name=pi01_path)
                close_sap = sap_close()
            # Do the same but for non-flash
                # open new session
            open_sap = sap_open()
            time.sleep(5)
            if open_sap == True:
                session = sap_logon(environment=environment, client=1)
                sap_code(tcode=t_code_sap, session=session)
                try:
                    sap_variant(session=session, var_to_use=variant_non_flash, version=1)
                except: 
                    sap_variant(session=session, var_to_use=variant_non_flash, version=2) # in our case, V2 is the most probable scenario

                session.findById("wnd[0]/usr/ctxtS_DATB1-LOW").text = first_day # this one differs
                # Run the report
                sap_run(session=session)
                sap_extract(session=session, extr_path="C:\\Users\\KOMAROVO\\Desktop\\Automations\\CPT_TP\\Monthly_emails\\Output\\Files_extracted", extr_name=ziv1_path)
                close_sap = sap_close()
                close_excel()

        def excel_rework_me():
            # flash
            #close_excel()
            global list_LEs
            list_LEs = []
            # grab with pandas, loop through, populate separete files in Files_spit, attach to email later
            df_pi01 = pd.read_excel(f"C:/Users/{curr_user}/Desktop/Automations/CPT_TP/Monthly_emails/Output/Files_extracted/{pi01_path}")
            df_email_contacts = pd.read_excel(f"C:/Users/{curr_user}/Desktop/Automations/CPT_TP/Monthly_emails/Contacts_emails_ME.xlsx", sheet_name="flash_copy")
            for index, row in df_email_contacts.iterrows():
                df_pi01_split = df_pi01[df_pi01['Sales Organization'] == row['LE']]
                if df_pi01_split.empty: # if no record for particular market exist
                    pass
                else:
                    save_to = f"C:/Users/{curr_user}/Desktop/Automations/CPT_TP/Monthly_emails/Output/Files_split/{row['LE']}-{pi01_path}"
                    df_pi01_split.to_excel(save_to, index=False)
                    list_LEs.append(row['LE'])
                    send_emails_me_flash(save_to, row['LE'])
            # non flash
            # do the same for ziv1 file
            df_ziv1 = pd.read_excel(f"C:/Users/{curr_user}/Desktop/Automations/CPT_TP/Monthly_emails/Output/Files_extracted/{ziv1_path}")
            df_email_contacts = pd.read_excel(f"C:/Users/{curr_user}/Desktop/Automations/CPT_TP/Monthly_emails/Contacts_emails_ME.xlsx", sheet_name="non_flash_copy")
            for index, row in df_email_contacts.iterrows():
                df_pi01_split = df_ziv1[df_ziv1['Customer Number'] == row['Customer LE']]
                if df_pi01_split.empty: # if no record for particular market exist
                    pass
                else:
                    save_to = f"C:/Users/{curr_user}/Desktop/Automations/CPT_TP/Monthly_emails/Output/Files_split/{row['LE']}-{ziv1_path}"
                    df_pi01_split.to_excel(save_to, index=False)
                    list_LEs.append(row['LE'])
                    send_emails_me_non_flash(save_to, row['LE'])
            save_last_x_emails()
            
        def save_last_x_emails():
            number_emails = len(list_LEs)
            # save files
            outlook = win32.Dispatch('outlook.application')
            mapi = outlook.GetNameSpace("MAPI")
            # 5 is for Sent items folder, 6 is for Inbox folder, more here: https://learn.microsoft.com/en-us/office/vba/api/outlook.oldefaultfolders
            folder = mapi.GetDefaultFolder(5)#.folders("example_folder") # default folder is defined via mapi (Messaging Application Programming Interface)
            items= folder.Items
            items.Sort("[ReceivedTime]", Descending=True)
            msgs = items.GetFirst()
            msgs.SaveAs(f"C:\\Users\\{curr_user}\\Desktop\\Automations\\CPT_TP\\Monthly_emails\\Output\\Emails\\{msgs.Subject}.msg")
            for _ in range(number_emails - 1):
                msgs = items.GetNext()
                msgs.SaveAs(f"C:\\Users\\{curr_user}\\Desktop\\Automations\\CPT_TP\\Monthly_emails\\Output\\Emails\\{msgs.Subject}.msg")
            print(msgs)


        def send_emails_me_flash(save_path, cocd):
            df_email_contacts = pd.read_excel(f"C:/Users/{curr_user}/Desktop/Automations/CPT_TP/Monthly_emails/Contacts_emails_ME.xlsx", sheet_name="flash_copy")
            for index, row in df_email_contacts.iterrows():
                if row['LE'] == cocd:
                    #--Send file by email:
                    outlook = win32.Dispatch('outlook.application')
                    mail = outlook.CreateItem(0)
                    mail.To = row['Emails2']
                    mail.Subject = f"Automatic email for {row['LE']}"
                    mail.Body = f"Please find attached a file with TP extract for your Cocd {row['LE']}"
                    mail.Attachments.Add(save_path)
                    #mail.Display(True)
                    #mail.SaveAs(f"C:\\Users\\{curr_user}\\Desktop\\Automations\\CPT_TP\\Emails\\Email to {row['LE']}.msg") # this one will save the mails in the unsaved state. And if put below mail.Send() will crush, for the item is already sent
                    mail.Send()
                    # time.sleep(3) # give enough time for the mail to appear in the mailbox
                    # mapi = outlook.GetNameSpace("MAPI")
                    # # 5 is for Sent items folder, 6 is for Inbox folder, more here: https://learn.microsoft.com/en-us/office/vba/api/outlook.oldefaultfolders
                    # folder = mapi.GetDefaultFolder(5)#.folders("example_folder") # default folder is defined via mapi (Messaging Application Programming Interface)
                    # items= folder.Items
                    # items.Sort("[ReceivedTime]", Descending=False)
                    # msgs = items.GetLast()
                    # msgs.SaveAs(f"C:\\Users\\{curr_user}\\Desktop\\Automations\\CPT_TP\\Monthly_emails\\Output\\Emails\\Email to {row['LE']}.msg")
                    # print(msgs)
                else: pass

        def send_emails_me_non_flash(save_path, cocd):
            # same for the non flash
            df_email_contacts = pd.read_excel(f"C:/Users/{curr_user}/Desktop/Automations/CPT_TP/Monthly_emails/Contacts_emails_ME.xlsx", sheet_name="non_flash_copy")
            for index, row in df_email_contacts.iterrows():
                if row['LE'] == cocd:
                    #--Send file by email:
                    outlook = win32.Dispatch('outlook.application')
                    mail = outlook.CreateItem(0)
                    mail.To = row['Emails2']
                    mail.Subject = f"Automatic email for {row['LE']}"
                    mail.Body = f"Please find attached a file with TP extract for your Cocd {row['LE']}"
                    mail.Attachments.Add(save_path)
                    #mail.Display(True)
                    #mail.SaveAs(f"C:\\Users\\{curr_user}\\Desktop\\Automations\\CPT_TP\\Emails\\Email to {row['LE']}.msg") # this one will save the mails in the unsaved state. And if put below mail.Send() will crush, for the item is already sent
                    mail.Send()
                    # time.sleep(3) # give enough time for the mail to appear in the mailbox
                    # mapi = outlook.GetNameSpace("MAPI")
                    # # 5 is for Sent items folder, 6 is for Inbox folder, more here: https://learn.microsoft.com/en-us/office/vba/api/outlook.oldefaultfolders
                    # folder = mapi.GetDefaultFolder(5)#.folders("example_folder") # default folder is defined via mapi (Messaging Application Programming Interface)
                    # items= folder.Items
                    # items.Sort("[ReceivedTime]", Descending=False)
                    # msgs = items.GetLast()
                    # msgs.SaveAs(f"C:\\Users\\{curr_user}\\Desktop\\Automations\\CPT_TP\\Monthly_emails\\Output\\Emails\\Email to {row['LE']}.msg")
                    # print(msgs) 
                else: pass

        def send_emails_me_original(save_path):
            df_email_contacts = pd.read_excel(f"C:/Users/{curr_user}/Desktop/Automations/CPT_TP/Monthly_emails/Contacts_emails_ME.xlsx", sheet_name="flash_copy")
            for index, row in df_email_contacts.iterrows():
                print(row['LE'], row['Emails2'])
                #--Send file by email:
                outlook = win32.Dispatch('outlook.application')
                mail = outlook.CreateItem(0)
                mail.To = row['Emails2']
                mail.Subject = f"Automatic email for {row['LE']}"
                mail.Body = f"Please find attached a file with TP extract for your Cocd {row['LE']}"
                mail.Attachments.Add(save_path)
                #mail.Display(True)
                #mail.SaveAs(f"C:\\Users\\{curr_user}\\Desktop\\Automations\\CPT_TP\\Emails\\Email to {row['LE']}.msg") # this one will save the mails in the unsaved state. And if put below mail.Send() will crush, for the item is already sent
                mail.Send()
                time.sleep(2) # give enough time for the mail to appear in the mailbox
                mapi = outlook.GetNameSpace("MAPI")
                # 5 is for Sent items folder, 6 is for Inbox folder, more here: https://learn.microsoft.com/en-us/office/vba/api/outlook.oldefaultfolders
                folder = mapi.GetDefaultFolder(5)#.folders("example_folder") # default folder is defined via mapi (Messaging Application Programming Interface)
                items= folder.Items
                items.Sort("[ReceivedTime]", Descending=False)
                msgs = items.GetLast()
                msgs.SaveAs(f"C:\\Users\\{curr_user}\\Desktop\\Automations\\CPT_TP\\Monthly_emails\\Output\\Emails\\Email to {row['LE']}.msg")
                print(msgs)
            # same for the non flash
            df_email_contacts = pd.read_excel(f"C:/Users/{curr_user}/Desktop/Automations/CPT_TP/Monthly_emails/Contacts_emails_ME.xlsx", sheet_name="non_flash_copy")
            for index, row in df_email_contacts.iterrows():
                print(row['LE'], row['Emails2'])
                #--Send file by email:
                outlook = win32.Dispatch('outlook.application')
                mail = outlook.CreateItem(0)
                mail.To = row['Emails2']
                mail.Subject = f"Automatic email for {row['LE']}"
                mail.Body = f"Please find attached a file with TP extract for your Cocd {row['LE']}"
                #mail.Attachments.Add(entry_path.get())
                #mail.Display(True)
                #mail.SaveAs(f"C:\\Users\\{curr_user}\\Desktop\\Automations\\CPT_TP\\Emails\\Email to {row['LE']}.msg") # this one will save the mails in the unsaved state. And if put below mail.Send() will crush, for the item is already sent
                mail.Send()
                time.sleep(2) # give enough time for the mail to appear in the mailbox
                mapi = outlook.GetNameSpace("MAPI")
                # 5 is for Sent items folder, 6 is for Inbox folder, more here: https://learn.microsoft.com/en-us/office/vba/api/outlook.oldefaultfolders
                folder = mapi.GetDefaultFolder(5)#.folders("example_folder") # default folder is defined via mapi (Messaging Application Programming Interface)
                items= folder.Items
                items.Sort("[ReceivedTime]", Descending=False)
                msgs = items.GetLast()
                msgs.SaveAs(f"C:\\Users\\{curr_user}\\Desktop\\Automations\\CPT_TP\\Monthly_emails\\Output\\Emails\\Email to {row['LE']}.msg")
                print(msgs)
                
        sap_extract_me()
        excel_rework_me()
        '''
        pass

    def monthly_emails_run_locally(self):
        if (self.btn_select_contacts.cget("text") == "Select contacts file ✓" 
            and self.btn_select_flash.cget("text") == "Select flash file ✓" 
            and self.btn_select_non_flash.cget("text") == "Select non flash file ✓"):
            print(self.final_path_to_emails, self.final_path_to_flash, self.final_path_to_non_flash)
            print("Ready to rumble!")
            # paste only excel related part below
            self.excel_rework_me()
        else:
            print("Data is missing")

    def excel_rework_me(self):
        self.pi01_name = f"pi01-{datetime.today().strftime('%d-%m-%Y')}.xlsx"
        self.ziv1_name = f"ziv1-{datetime.today().strftime('%d-%m-%Y')}.xlsx"
        self.directory = f"C:/Users/{self.curr_user}/Desktop/Automations/CPT_TP/Monthly_emails/Output/Files_split/{datetime.today().strftime('%d-%m-%Y')}"
        if not os.path.exists(self.directory):
            os.makedirs(self.directory)
        # flash
        #close_excel()
        self.list_LEs = []
        # grab with pandas, loop through, populate separete files in Files_spit, attach to email later
        df_pi01 = pd.read_excel(self.final_path_to_flash)
        df_email_contacts = pd.read_excel(self.final_path_to_emails, sheet_name="flash_copy")
        for index, row in df_email_contacts.iterrows():
            df_pi01_split = df_pi01[df_pi01['Sales Organization'] == row['LE']]
            if df_pi01_split.empty: # if no record for particular market exist
                pass
            else:
                save_to = f"{self.directory}/{row['LE']}-{self.pi01_name}"
                df_pi01_split.to_excel(save_to, index=False)
                self.list_LEs.append(row['LE'])
                self.send_emails_me_flash(save_to, row['LE'])
        # non flash
        # do the same for ziv1 file
        df_ziv1 = pd.read_excel(self.final_path_to_non_flash)
        df_email_contacts = pd.read_excel(self.final_path_to_emails, sheet_name="non_flash_copy")
        for index, row in df_email_contacts.iterrows():
            df_pi01_split = df_ziv1[df_ziv1['Customer Number'] == row['Customer LE']]
            if df_pi01_split.empty: # if no record for particular market exist
                pass
            else:
                save_to = f"{self.directory}/{row['LE']}-{self.ziv1_name}"
                df_pi01_split.to_excel(save_to, index=False)
                self.list_LEs.append(row['LE'])
                self.send_emails_me_non_flash(save_to, row['LE'])
        self.save_last_x_emails()
    
    def save_last_x_emails(self):
        number_emails = len(self.list_LEs)
        self.directory_emails = f"C:/Users/{self.curr_user}/Desktop/Automations/CPT_TP/Monthly_emails/Output/Emails/{datetime.today().strftime('%d-%m-%Y')}"
        if not os.path.exists(self.directory_emails):
            os.makedirs(self.directory_emails)
        # save files
        outlook = win32.Dispatch('outlook.application')
        mapi = outlook.GetNameSpace("MAPI")
        # 5 is for Sent items folder, 6 is for Inbox folder, more here: https://learn.microsoft.com/en-us/office/vba/api/outlook.oldefaultfolders
        folder = mapi.GetDefaultFolder(5)#.folders("example_folder") # default folder is defined via mapi (Messaging Application Programming Interface)
        items= folder.Items
        items.Sort("[ReceivedTime]", Descending=True)
        msgs = items.GetFirst()
        msgs.SaveAs(f"C:\\Users\\{self.curr_user}\\Desktop\\Automations\\CPT_TP\\Monthly_emails\\Output\\Emails\\{datetime.today().strftime('%d-%m-%Y')}\\{msgs.Subject}.msg")
        for _ in range(number_emails - 1):
            msgs = items.GetNext()
            msgs.SaveAs(f"C:\\Users\\{self.curr_user}\\Desktop\\Automations\\CPT_TP\\Monthly_emails\\Output\\Emails\\{datetime.today().strftime('%d-%m-%Y')}\\{msgs.Subject}.msg")
        print(msgs)

    def send_emails_me_flash(self, save_path, cocd):
        df_email_contacts = pd.read_excel(self.final_path_to_emails, sheet_name="flash_copy")
        for index, row in df_email_contacts.iterrows():
            if row['LE'] == cocd:
                #--Send file by email:
                outlook = win32.Dispatch('outlook.application')
                mail = outlook.CreateItem(0)
                mail.To = row['Emails2']
                mail.Subject = f"Automatic email for {row['LE']}"
                mail.Body = f"Please find attached a file with TP extract for your Cocd {row['LE']}"
                mail.Attachments.Add(save_path)
                #mail.Display(True)
                #mail.SaveAs(f"C:\\Users\\{curr_user}\\Desktop\\Automations\\CPT_TP\\Emails\\Email to {row['LE']}.msg") # this one will save the mails in the unsaved state. And if put below mail.Send() will crush, for the item is already sent
                mail.Send()
                # time.sleep(3) # give enough time for the mail to appear in the mailbox
                # mapi = outlook.GetNameSpace("MAPI")
                # # 5 is for Sent items folder, 6 is for Inbox folder, more here: https://learn.microsoft.com/en-us/office/vba/api/outlook.oldefaultfolders
                # folder = mapi.GetDefaultFolder(5)#.folders("example_folder") # default folder is defined via mapi (Messaging Application Programming Interface)
                # items= folder.Items
                # items.Sort("[ReceivedTime]", Descending=False)
                # msgs = items.GetLast()
                # msgs.SaveAs(f"C:\\Users\\{curr_user}\\Desktop\\Automations\\CPT_TP\\Monthly_emails\\Output\\Emails\\Email to {row['LE']}.msg")
                # print(msgs)
            else: pass

    def send_emails_me_non_flash(self, save_path, cocd):
        # same for the non flash
        df_email_contacts = pd.read_excel(self.final_path_to_emails, sheet_name="non_flash_copy")
        for index, row in df_email_contacts.iterrows():
            if row['LE'] == cocd:
                #--Send file by email:
                outlook = win32.Dispatch('outlook.application')
                mail = outlook.CreateItem(0)
                mail.To = row['Emails2']
                mail.Subject = f"Automatic email for {row['LE']}"
                mail.Body = f"Please find attached a file with TP extract for your Cocd {row['LE']}"
                mail.Attachments.Add(save_path)
                #mail.Display(True)
                #mail.SaveAs(f"C:\\Users\\{curr_user}\\Desktop\\Automations\\CPT_TP\\Emails\\Email to {row['LE']}.msg") # this one will save the mails in the unsaved state. And if put below mail.Send() will crush, for the item is already sent
                mail.Send()
            else: pass

#--Generic classes--------------------
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

class Generic_Checkbox(ctk.CTkCheckBox):
    def __init__(self, parent, text_var, variable_var):
        super().__init__(
            parent,
            text=text_var,
            variable=variable_var,
            onvalue="on",
            offvalue="off",
            font=("", 13, "bold"),
            hover_color=BUTTON_HOVER,
            fg_color=BUTTON_HOVER,
        )

class Generic_Label(ctk.CTkLabel):
    def __init__(self, parent, name):
        super().__init__(
            parent,
            text=name,
            font=("", 14, "bold"),
            fg_color=LABEL_FG,
            text_color=TEXT_NORMAL,
        )

class Generic_Entry(ctk.CTkEntry):
    def __init__(self, parent, radius):
        super().__init__(
            parent,
            justify="center",
            fg_color=ENTRY_NORMAL,
            corner_radius=radius,
            font=("", 14, "bold"),
            takefocus=False)

class Generic_Frame(ctk.CTkFrame):
    def __init__(self, parent):
        super().__init__(
            parent, 
            border_width=4, 
            border_color=LABEL_FG, 
            fg_color=LABEL_FG)

class Generic_Separator(ttk.Separator):
    def __init__(self, parent):
        super().__init__(
            parent, 
            orient="horizontal")
#--End generic classes----------------

if __name__ == "__main__":
    app = App(size=(240, 235))
