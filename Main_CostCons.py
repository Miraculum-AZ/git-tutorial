# importing modules
import shutil  # for file manipulation
import os
import win32com.client as win32
from ok_sap_script import *
import time
from datetime import date, timedelta, datetime, timezone
import pyautogui
import pandas as pd
import numpy as np
from functools import wraps
import logging # custom logging created
from pathlib import Path # working with / // or \


class MyLogger:
    # Define custom log levels
    Custom1 = 15 # number are for levels having different number behind (10 for Debug and 50 for Error, I suppose)
    Custom2 = 45

    def __init__(self, log_file_path, log_level=logging.INFO):
        '''Define path, define minimal log level, call setup function'''
        self.receiver_log = "oleksandr.komarov@zoetis.com" # my name in case of log failure
        self.log_file_path = log_file_path
        self.log_level = log_level
        self.logger = self._setup_logger()

    def _setup_logger(self): # _ means that this method is to be used inside this class only (even though nothing prevants you from using it anyway)
        logger = logging.getLogger(self.log_file_path)
        logger.setLevel(self.log_level)
        self._add_file_handler(logger) # how to handle log .txt file
        self._add_console_handler(logger) # how logs are displayed in the console

        return logger

    def _add_file_handler(self, logger):
        file_formatter = self._get_file_formatter()
        with open(self.log_file_path, mode='w') as log_file:
            file_handler = logging.FileHandler(self.log_file_path, mode='w')
            file_handler.setFormatter(file_formatter)
            logger.addHandler(file_handler)

    def _add_console_handler(self, logger):
        console_formatter = self._get_console_formatter()
        console_handler = logging.StreamHandler()
        console_handler.setFormatter(console_formatter)
        logger.addHandler(console_handler)

    def _get_file_formatter(self):
        if self.log_level == logging.INFO:
            return logging.Formatter('%(asctime)s - %(levelname)s - %(message)s', datefmt='%Y-%m-%d %H:%M:%S')
        else:
            return logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s: error occured at line %(lineno)d', datefmt='%Y-%m-%d %H:%M:%S')

    def _get_console_formatter(self):
        if self.log_level <= logging.INFO:
            return logging.Formatter('%(levelname)s - %(message)s')
        elif self.log_level <= logging.ERROR:
            return logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s: error occured at line %(lineno)d', datefmt='%Y-%m-%d %H:%M:%S')
        else:
            if self.log_level == MyLogger.Custom1:
                return logging.Formatter('%(levelname)s - %(message)s')
            elif self.log_level == MyLogger.Custom2:
                return logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s', datefmt='%Y-%m-%d %H:%M:%S')
            else:
                return logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')

    def log(self, message, level=logging.INFO):
        '''Main function to write a log message'''
        self.logger.log(level, message)
        if level >= logging.ERROR:
            self.reason = message
            self._send_email_failed()

    def _send_email_failed(self):
        # send this failure by email
        self.outlook = win32.Dispatch('outlook.application')
        self.mail = self.outlook.CreateItem(0)
        self.mail.To = self.receiver_log
        self.mail.Subject = f"Control failure :("
        self.mail.Body = f"This is to inform you that one of the controls failed.\nReason {self.reason}\n.Please consult attached log file to know a reason."
        self.mail.Attachments.Add(f"{self.log_file_path}")
        #self.mail.Display(True)
        self.mail.Send()

class MainControl():
    '''Set variables and functions that are copied from control to control.
    By doing this we opt for more standartisation'''
    def __init__(self):
        # main variables
        self.curr_user = os.getlogin() # what user runs this code
        self.date_stamp = datetime.today().strftime("%d-%m-%Y") # datestamp to differentiate between runs
        self.receiver = "oleksandr.komarov@zoetis.com" # who will receive an email with control
        # paths
        self.generic_path = Path(f"C:/Users/{self.curr_user}/Desktop/Automations/CPT_TP/") # path to the folder, where controls are stored
        # SAP variables global
        self.environment = "ECC Production" # which environment to use
        # handling dates in SAP
        self.period = date.today().strftime("%#m") # use # to remove a leading 0 -> result 7 and not 07, 8 and not 08
        self.period_0 = date.today().strftime("%m") # print with leading 0 -> 07, 08, etc.
        self.year = datetime.today().year # get current year
        print("Another foolish comment")
            # Additional dates
        self.last_month_last_day = (datetime.today().replace(day=1) - timedelta(days=1)).strftime("%d.%m.%Y")
        self.last_month_first_day = (datetime.today().replace(day=1) - timedelta(days=1)).replace(day=1).strftime("%d.%m.%Y")
        self.last_month =  (datetime.today().replace(day=1) - timedelta(days=1)).strftime("%#m")

    def __repr__(self) -> str: # how our control is seen for the users, who print it (more official)
        print("Some test here")
        return (f"Control on {self.curr_user} machine") # an instance of this class will have this text if printed
    
    def __str__(self) -> str: # how our control is seen for the users, who print it (more user-friendly)
        return (f"Control by name {self.__class__.__name__} on {self.curr_user} machine") # an instance of this class will have this text if printed
    
    @staticmethod
    def sap_decorator(sap_function):
        '''Closing SAP sessions before and after running a script, entering tcode, performing logging and error handling'''
        @wraps(sap_function)
        def sap_wrapper(self, t_code_sap, *args, **kwargs):
            sap_close() # ensure there are no open session
            self.main_logger.log(level=logging.INFO, message="SAP is terminated prior function execution")
            open_sap = sap_open() # open new session
            self.main_logger.log(level=logging.INFO, message="New SAP session is opened")
            time.sleep(5) # make sure SAP opens up
            if open_sap == True:
                self.main_logger.log(level=logging.INFO, message=f"Trying to enter {self.environment} environment")
                self.session = sap_logon(environment=self.environment, client=1)
                self.main_logger.log(level=logging.INFO, message=f"Trying to execute {t_code_sap} code in SAP")
                sap_code(tcode=t_code_sap, session=self.session)
                sap_function(self) # this is the main fuction to be decorated
            else:
                self.main_logger.log(level=logging.ERROR, message="SAP did not open")
            sap_close() # close SAP
            self.main_logger.log(level=logging.INFO, message="SAP is terminated after function execution")
        return sap_wrapper
    
    @staticmethod
    def timer_decorator(func_to_time):
        '''This one is used to track the execution time of any particular function'''
        import time
        @wraps(func_to_time)
        def time_wrapper(self, *args, **kwargs):
            t1 = time.time()
            func_to_time(self, *args, **kwargs)
            t2 = time.time() - t1
            self.main_logger.log(level=logging.INFO, message='{} run in: {} seconds'.format(func_to_time.__name__, round(t2,2)))
        return time_wrapper

    @staticmethod
    def excel_decorator(excel_function):
        '''Opening Excel, applying some standard parameters for runtime optimisation, etc.'''
        def excel_wrapper(self, *args, **kwargs):
            try:
                # default Excel runtime optimisation
                self.excel = win32.Dispatch("Excel.Application")
                self.excel.AskToUpdateLinks = False
                self.excel.DisplayAlerts = False
                self.excel.Visible = True
                self.excel.ScreenUpdating = False
                excel_function(self) # our main Excel function to run
                # no closing decorators as it differs from control to control
            except Exception as e:
                # return to a normal Excel, then close it
                self.excel.ScreenUpdating = True
                self.excel.Application.Calculation = -4105  # to set xlCalculationAutomatic
                close_excel() # close excel
                self.main_logger.log(level=logging.ERROR, message=f'Error while using Excel or Excel-related functions, namely, {e}')
        return excel_wrapper
    
    def copy_file(self, file_from: str, file_to: str):
        '''Simple function to copy one file to a given location'''
        try:
            shutil.copy(file_from, file_to)
            self.main_logger.log(level=logging.INFO, message=f"File {file_from} is copied as {file_to}")
        except PermissionError as e:
            self.main_logger.log(level=logging.ERROR, message=f"Could not create a new file, the file might be open")
            os.system("taskkill /f /im  excel.exe")

    def take_screenshot(self, name:str, path:str):
        '''A simple function to take screenshots using pyautogui\n 
        It is not used in this class and purposed for inheritance\n
        Args:
            name (str): a file name, spare .png. 
            path (str): fr path, meaning / being used as separators.\n
        Example:
            take_screenshot(name=faglb03_scr_for_30-08-2023, path=fr'C:/Users/KOMAROVO/Desktop/Automations/CPT_TP/GIT/Screenshots/')\n
        Returns:
            None
            '''
        try: # try taking screenshot
            screenshot = pyautogui.screenshot() 
            screenshot.save(str(Path(f'{path}/{name}.png')))
            self.main_logger.log(level=logging.INFO, message=f"Screenshot is taken with a name of {name}.png under location {path}")
        except Exception as e: 
            self.main_logger.log(level=logging.ERROR, message=f"Failed to take screenshot {name} for a reason {e}")

    def add_screenshots(self, worksheet, list_screenshots, picture_path):
        '''This function will add screenshot or screenshots into Excel worksheet'''
        try:
            # remove old screenshots
            self.pictures = worksheet.Pictures()
            for pic in self.pictures:
                pic.Delete()
            self.main_logger.log(level=logging.INFO, message=f"Old screenshot removed from {worksheet}")
            # add screenshots
            self.left, self.top, self.width, self.height = 0, 30, 792, 534
            for screenshot in list_screenshots:
                self.picture_filename = screenshot
                self.full_picture_path = str(Path(f"{picture_path}/{self.picture_filename}.png"))
                # insert a new screenshot with given parameters
                self.picture = worksheet.Shapes.AddPicture(self.full_picture_path, LinkToFile=False, SaveWithDocument=True, Left=self.left, Top=self.top, Width=self.width, Height=self.height)
                self.top += self.height # the following one will be put below the current one
                self.main_logger.log(level=logging.INFO, message=f"Screenshot with a name {self.picture_filename}.png is added")
        except Exception as e:
            self.main_logger.log(level=logging.ERROR, message=f"No screenshots are found or {e}")

    def send_email(self, subject:str = "Automatic email", body:str="Please consult the attachment"):
        '''Send email'''
        self.subject = subject
        self.body = body
        # send by email
        self.outlook = win32.Dispatch('outlook.application')
        self.mail = self.outlook.CreateItem(0)
        self.mail.To = self.receiver
        self.mail.Subject = f"{self.subject} for {self.date_stamp}"
        self.mail.Body = self.body
        try:
            self.mail.Attachments.Add(f"{self.new_file_path}")
        except:
            self.main_logger.log(level=logging.ERROR, message=f"No attachment found under location {self.new_file_path}")
        #self.mail.Display(True)
        self.mail.Send()
        self.main_logger.log(level=logging.INFO, message=f"Email sent to {self.receiver}")

class Cost_consistency(MainControl):
    def __init__(self):
        super().__init__() # inherit from the MainControl
        self.name = self.__class__.__name__ # name of our class -> name of the folder, where this control is stored
        self.generic_path = str(Path(f"{self.generic_path}/{self.name}")) # up to the folder, where this control is stored
        self.output_path = str(Path(f"{self.generic_path}/Output"))
        self.screenshots_path = str(Path(f"{self.generic_path}/Screenshots"))
        self.workings_path = str(Path(f"{self.generic_path}/Workings"))

        # Logging
        self.log_file_path = f"{self.generic_path}/Logging/{self.name}_log_for_{self.date_stamp}.txt"
        self.main_logger = MyLogger(self.log_file_path)
        self.main_logger.log(level=logging.INFO, message="Class instantiated")

        # SAP variables ZM2D_28
        self.t_code_sap = "ZM2D_28"
        self.variant_sap = "COST_CONS_AUTO"
        self.layout_sap = "CC_AUTO"

        # SAP variables mb51
        self.t_code_sap_mb51 = "MB51"
        self.variant_sap_mb51 = "COST_CONS_AUTO"
        self.layout_sap_mb51 = "/DPM"

        # extract paths zm2d_28
        self.extract_path_zm2d_28 = self.output_path
        self.extract_name_zm2d_28 = f"zm2d_28_extract_for_{self.date_stamp}.xlsx"

        # extract paths mb51
        self.extract_path_mb51 = self.output_path
        self.extract_name_mb51 = f"mb51_extract_for_{self.date_stamp}.xlsx"

        # Paths to files
            # model file
        self.model_file_path = str(Path(f"{self.generic_path}/{self.name}_model_file.xlsx"))
        self.new_file_path = str(Path(f"{self.output_path}/{self.name}_for_{self.date_stamp}.xlsx"))
            # zm2d_28 file path
        self.zm2d_28_file_path = str(Path(f"{self.output_path}/{self.extract_name_zm2d_28}"))
            # mb51 file path
        self.mb51_file_path = str(Path(f"{self.output_path}/{self.extract_name_mb51}"))
            # historical cost file
        self.final_path_to_historical_cost = f"S:/Robotics_COE_Prod/RPA/MM60/Outputs/{self.year}/{self.period_0}/"
        self.historical_cost_file = f"Historical cost AP{self.period_0} LE2941.xlsm"
        self.final_joined_path_to_historical_cost = str(Path(f"{self.final_path_to_historical_cost}{self.historical_cost_file}"))
                # same but on my local drive
        self.new_file_path_to_historical_cost = str(Path(f"{self.output_path}/{self.historical_cost_file}"))

    @MainControl.timer_decorator
    def __call__(self, *args, **kwargs):
        self.run_logic_sap_zm2d_28(self.t_code_sap) # extract zm2d_28 from SAP
        self.run_logic_excel() # paste hictorical cost and zm2d_28 in Excel, add screenshots
        self.run_logic_excel_part_2() # final part of working with mb51 excel and main file 
        close_excel()
        x = 1+1

    @MainControl.timer_decorator
    @MainControl.sap_decorator
    def run_logic_sap_zm2d_28(self):
        # self.session is defined inside sap_decorator
            try:
                sap_variant(session=self.session, var_to_use=self.variant_sap)
            except: 
                sap_variant(session=self.session, var_to_use=self.variant_sap, version=2) # in our case, V2 is the most probable scenario
            # adjust the dates
            self.session.findById("wnd[0]/usr/txtSP$00001-LOW").text = self.last_month
            self.session.findById("wnd[0]/usr/txtSP$00002-LOW").text = self.year
            self.session.findById("wnd[0]/usr/ctxt%LAYOUT").text = self.layout_sap
            sap_run(session=self.session) # run report
             # take screenshots of first and last pages
            #pyautogui.hotkey('ctrl', 'end') # I almist feel this one is cheating
            self.session.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell").pressToolbarContextButton("&MB_VIEW")
            self.session.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell").selectContextMenuItem("&PRINT_BACK_PREVIEW")
            self.session.findById("wnd[0]/usr/lbl[12,6]").setFocus() # just in case
            self.take_screenshot(name=f"scr1_for_{self.extract_name_zm2d_28[:-5]}", path=self.screenshots_path) # first page
            sap_screen_nagivation(session=self.session, action="down")
            time.sleep(3)
            self.take_screenshot(name=f"scr2_for_{self.extract_name_zm2d_28[:-5]}", path=self.screenshots_path) # last page
            sap_extract(session=self.session, extr_path=self.extract_path_zm2d_28, extr_name=self.extract_name_zm2d_28)
            close_excel() # close excel
        
    @MainControl.timer_decorator
    @MainControl.excel_decorator
    def run_logic_excel(self):
        self.copy_file(file_from=self.model_file_path, file_to=self.new_file_path) # copy model file
    # working with historical cost file
        self.copy_file(file_from=self.final_joined_path_to_historical_cost, file_to=self.new_file_path_to_historical_cost) # copy historical cost file
        # open created Excel file and perform the following manipulations
            # open new file and disable calculations
        self.new_file = self.excel.Workbooks.Open(self.new_file_path) # main file
        self.new_file_hc = self.excel.Workbooks.Open(self.new_file_path_to_historical_cost)
        self.excel.Application.Calculation = (-4135)  # to set xlCalculationManual # Workbook needs to be opened
            # other steps 
        self.new_file_hc_ws = self.new_file_hc.Sheets("MM60 Report") # open historical cost file
        self.last_row_historical_file_ws = self.new_file_hc_ws.Cells(1, 1).End(-4121).Row -1 # find last row of this file # do not include totals
        print(self.last_row_historical_file_ws)

        # open main file and paste the values from mm60 to hc sheet
        self.new_file_hcost_ws = self.new_file.Worksheets("HC")
        self.new_file_hc_ws.Range(f"A2:G{self.last_row_historical_file_ws}").SpecialCells(12).Copy(Destination=self.new_file_hcost_ws.Range(f"B2")) 
        # add concatenate formula =IF(ISBLANK(C2),"",B2&"-"&C2)
        self.new_file_hcost_ws.Range(f"A2:A{self.last_row_historical_file_ws}").Formula = f'=IF(ISBLANK(C2),"",B2&"-"&C2)'
        self.new_file_hc.Close() # close RPA file with hc 
        self.new_file.Save() # save the main file
    # with that we can forget hc file and never think of it anymore
    # working with zm2d_28 extract from SAP
        try:
            # work with the main file
            self.new_file_zm2d_28_ws = self.new_file.Worksheets("ZM2D_28") # navigate to the respective ws
            self.last_row_new_file_zm2d_28_ws = self.new_file_zm2d_28_ws.Cells(1, 1).End(-4121).Row # last row of the current data set
            self.new_file_zm2d_28_ws.Range(f"A2:AE{self.last_row_new_file_zm2d_28_ws}").ClearContents() # prepare the space
            # work with zm2d_28 extract
            self.zm2d_28_file = self.excel.Workbooks.Open(self.zm2d_28_file_path) # open zm2d_28 sap extract
            self.zm2d_28_ws = self.zm2d_28_file.Sheets(1) # open first and only ws it has
            self.last_row_zm2d_28_ws = self.zm2d_28_ws.Cells(1, 1).End(-4121).Row -1 # we do not need to include the totals
            self.zm2d_28_ws.Range(f"A2:AA{self.last_row_zm2d_28_ws}").SpecialCells(12).Copy(Destination=self.new_file_zm2d_28_ws.Range(f"D2")) # copy info
            # add formulas
            self.new_file_zm2d_28_ws.Range(f"A2:A{self.last_row_zm2d_28_ws}").Formula = f'=IF(D2="","",E2&"-"&J2)'
            self.new_file_zm2d_28_ws.Range(f"B2:B{self.last_row_zm2d_28_ws}").Formula = f'=IF(D2="","",IF(W2=AA2,"NO","YES"))'
            self.new_file_zm2d_28_ws.Range(f"C2:C{self.last_row_zm2d_28_ws}").Formula = f'=VLOOKUP(A2,costs,5,FALSE)/VLOOKUP(A2,costs,8,FALSE)'
            self.zm2d_28_file.Close() # close zm2d_28 extract
            # apply filter
            self.new_file_zm2d_28_ws.Range("A1:AD1").AutoFilter(Field=2, Criteria1="YES") 
            # paste screenshots of zm2d_28 first and last page
            try:
                self.new_file_screenshots_ws = self.new_file.Sheets("Screenshots")
                self.pictures = self.new_file_screenshots_ws.Pictures()
                for pic in self.pictures:
                    pic.Delete() # remove all existing screenshots
                    #1
                #C:/Users/{self.curr_user}/Desktop/Automations/CPT_TP/Cost_consistency/Screenshots/scr2_for_{self.extract_name_zm2d_28[:-5]}.png'
                self.picture_filename = f'scr1_for_{self.extract_name_zm2d_28[:-5]}.png'
                self.picture_path = f'{self.screenshots_path}/{self.picture_filename}'
                # insert a new screenshot with given parameters
                self.left, self.top, self.width, self.height = 0, 15, 950, 640
                self.picture = self.new_file_screenshots_ws.Shapes.AddPicture(self.picture_path, LinkToFile=False, SaveWithDocument=True, Left=self.left, Top=self.top, Width=self.width, Height=self.height)
                    #2
                self.picture_filename = f'scr2_for_{self.extract_name_zm2d_28[:-5]}.png'
                self.picture_path = f'{self.screenshots_path}/{self.picture_filename}'
                # insert a new screenshot with given parameters
                self.left, self.top, self.width, self.height = 950, 15, 950, 640
                self.picture = self.new_file_screenshots_ws.Shapes.AddPicture(self.picture_path, LinkToFile=False, SaveWithDocument=True, Left=self.left, Top=self.top, Width=self.width, Height=self.height)
            except Exception as e:
                self.main_logger.log(level=logging.ERROR, message=f"No screenshot or {e}")
            self.new_file.Save() # save main Excel file
        except Exception as e:
            self.main_logger.log(level=logging.ERROR, message=f"Something happened while copying ZM2D_28, namely {e}")
            self.new_file.Save()
        '''# copy to and from the clipboard -> nice but probably not the best    
            # materials
        # self.new_file_zm2d_28_ws.Range(f"J2:J{self.last_row_zm2d_28_ws}").SpecialCells(12).Copy()
        # win32clipboard.OpenClipboard()
        # self.copied_materials = win32clipboard.GetClipboardData(win32clipboard.CF_UNICODETEXT)
        # win32clipboard.CloseClipboard()
        #     # plants
        # self.new_file_zm2d_28_ws.Range(f"E2:E{self.last_row_zm2d_28_ws}").SpecialCells(12).Copy()
        # win32clipboard.OpenClipboard()
        # self.copied_plants = win32clipboard.GetClipboardData(win32clipboard.CF_UNICODETEXT)
        # win32clipboard.CloseClipboard()'''
    # Here execute second SAP function
        self.run_logic_sap_mb51(self.t_code_sap_mb51) # extract mb51 file from SAP

    @MainControl.timer_decorator
    @MainControl.excel_decorator
    def run_logic_excel_part_2(self):    
    # reopen main file
        self.new_file = self.excel.Workbooks.Open(self.new_file_path) # main file
        self.excel.Application.Calculation = (-4135)  # to set xlCalculationManual # Workbook needs to be opened
        self.new_file_mb51_ws = self.new_file.Worksheets("MB51") # navigate to the MB51 ws
        self.last_row_new_file_mb51_ws = self.new_file_mb51_ws.Cells(1, 1).End(-4121).Row # last row of the current data set
        self.new_file_mb51_ws.Range(f"A2:AA{self.last_row_new_file_mb51_ws}").ClearContents() # prepare the space
    # work with mb51 extract
        self.mb51_file = self.excel.Workbooks.Open(self.mb51_file_path) # open mb51 extract
        self.mb51_ws = self.mb51_file.Sheets(1) # navigate to the extrac
        self.last_row_mb51_ws = self.mb51_ws.Cells(1, 1).End(-4121).Row -1 # we do not need to include the totals
        self.mb51_ws.Range(f"A2:R{self.last_row_mb51_ws}").SpecialCells(12).Copy(Destination=self.new_file_mb51_ws.Range(f"I2")) # copy info
    # add formulas
        self.new_file_mb51_ws.Range(f"A2:A{self.last_row_mb51_ws}").Formula = f'=J2&"-"&K2'
        self.new_file_mb51_ws.Range(f"B2:B{self.last_row_mb51_ws}").Formula = f'=W2/T2'
        self.new_file_mb51_ws.Range(f"C2:C{self.last_row_mb51_ws}").Formula = f'=IF(ISNA(VLOOKUP(A2,costupdates,2,FALSE)),"NO",VLOOKUP(A2,costupdates,2,FALSE))'
        self.new_file_mb51_ws.Range(f"D2:D{self.last_row_mb51_ws}").Formula = f'=IF(C2="NO","-",VLOOKUP(A2,costupdates,3,FALSE))'
        self.new_file_mb51_ws.Range(f"E2:E{self.last_row_mb51_ws}").Formula = f'=IFERROR(D2-B2,"No change")'
        self.new_file_mb51_ws.Range(f"F2:F{self.last_row_mb51_ws}").Formula = f'=IFERROR(E2*T2, "No impact")'
        self.new_file_mb51_ws.Range(f"G2:G{self.last_row_mb51_ws}").Formula = f'=IF(D2="-","NO",IF(ROUND(B2,2)=ROUND(D2,2),"NO","YES"))'
        self.new_file_mb51_ws.Range(f"H2:H{self.last_row_mb51_ws}").Formula = f'=IF(AND(N2="101",O2="GR stock in transit"), "No entry required for movement type 101 GR stock in transit", "")'
        self.new_file_mb51_ws.Range(f"AA2:AA{self.last_row_mb51_ws}").Formula = f"=VLOOKUP(J2,'plant owner'!A:B,2,FALSE)"
    # add filter
        self.new_file_mb51_ws.Range("A1:AA1").AutoFilter(Field=7, Criteria1="YES") 
    # paste screenshots
        try:
            self.new_file_screenshots_ws = self.new_file.Sheets("Screenshots")
            self.pictures = self.new_file_screenshots_ws.Pictures()
                #1
            #C:/Users/{self.curr_user}/Desktop/Automations/CPT_TP/Cost_consistency/Screenshots/scr2_for_{self.extract_name_zm2d_28[:-5]}.png'
            self.picture_filename = f'scr1_for_mb51_{self.date_stamp}.png'
            self.picture_path = f'{self.screenshots_path}/{self.picture_filename}'
            # insert a new screenshot with given parameters
            self.left, self.top, self.width, self.height = 0, 640, 950, 640
            self.picture = self.new_file_screenshots_ws.Shapes.AddPicture(self.picture_path, LinkToFile=False, SaveWithDocument=True, Left=self.left, Top=self.top, Width=self.width, Height=self.height)
                #2
            self.picture_filename = f'scr2_for_mb51_{self.date_stamp}.png'
            self.picture_path = f'{self.screenshots_path}/{self.picture_filename}'
            # insert a new screenshot with given parameters
            self.left, self.top, self.width, self.height = 950, 640, 950, 640
            self.picture = self.new_file_screenshots_ws.Shapes.AddPicture(self.picture_path, LinkToFile=False, SaveWithDocument=True, Left=self.left, Top=self.top, Width=self.width, Height=self.height)
        except Exception as e:
            self.main_logger.log(level=logging.ERROR, message=f"No screenshot or {e}")
        self.new_file.Save() # save main Excel file
        # back to default
        self.excel.ScreenUpdating = True
        self.excel.Application.Calculation = -4105  # to set xlCalculationAutomatic
        self.new_file.Close() # close the main file
        self.mb51_file.Close() # close mb51 file
        self.send_email() # send new file by email
    
    @MainControl.timer_decorator
    @MainControl.sap_decorator
    def run_logic_sap_mb51(self):
        # self.session is defined inside sap_decorator
        try:
            sap_variant(session=self.session, var_to_use=self.variant_sap_mb51)
        except: 
            sap_variant(session=self.session, var_to_use=self.variant_sap_mb51, version=2) # in our case, V2 is the most probable scenario
        # adjust parameters
        self.session.findById("wnd[0]/usr/btn%_MATNR_%_APP_%-VALU_PUSH").press()
        self.new_file_zm2d_28_ws.Range(f"J2:J{self.last_row_zm2d_28_ws}").SpecialCells(12).Copy()
        self.session.findById("wnd[1]/tbar[0]/btn[24]").press()
        self.session.findById("wnd[1]/tbar[0]/btn[8]").press()
        self.session.findById("wnd[0]/usr/btn%_WERKS_%_APP_%-VALU_PUSH").press()
        self.new_file_zm2d_28_ws.Range(f"E2:E{self.last_row_zm2d_28_ws}").SpecialCells(12).Copy()
        self.session.findById("wnd[1]/tbar[0]/btn[24]").press()
        self.session.findById("wnd[1]/tbar[0]/btn[8]").press()
        self.session.findById("wnd[0]/usr/ctxtBUDAT-LOW").text = self.last_month_first_day
        self.session.findById("wnd[0]/usr/ctxtBUDAT-HIGH").text = self.last_month_last_day
        self.session.findById("wnd[0]/usr/ctxtALV_DEF").text = self.layout_sap_mb51
        # run the report
        sap_run(session=self.session)
        sap_screen_nagivation(session=self.session, action="up")
        #self.session.findById("wnd[0]/usr/lbl[1,1]").setFocus() # ensure report is loaded
            #scr1
        try: # first page
            screenshot_first_page_mb51 = pyautogui.screenshot()
            screenshot_first_page_mb51.save(fr'C:/Users/{self.curr_user}/Desktop/Automations/CPT_TP/Cost_consistency/Screenshots/scr1_for_mb51_{self.date_stamp}.png')
        except Exception as e: print(e)
        sap_screen_nagivation(session=self.session, action="down")
        time.sleep(2)
            #scr2
        try: # last page
            screenshot_last_page_mb51 = pyautogui.screenshot()
            screenshot_last_page_mb51.save(fr'C:/Users/{self.curr_user}/Desktop/Automations/CPT_TP/Cost_consistency/Screenshots/scr2_for_mb51_{self.date_stamp}.png')
        except Exception as e: print(e)
        self.session.findById("wnd[0]/tbar[1]/btn[48]").press() # change the layout
        sap_extract(session=self.session, extr_path=self.extract_path_mb51, extr_name=self.extract_name_mb51) #extract the report
        close_excel()

if __name__ == "__main__": 
    control = Cost_consistency()
    control()
