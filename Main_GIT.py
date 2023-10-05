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

    def __repr__(self) -> str: # how our control is seen for the users, who print it (more official)
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
        self.body = body # define email body
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

class GIT(MainControl):
    def __init__(self):
        super().__init__() # inherit from the MainControl
        self.name = self.__class__.__name__ # name of our class -> name of the folder, where this control is stored
        self.generic_path = str(Path(f"{self.generic_path}/{self.name}")) # up to the folder, where this control is stored
        self.output_path = str(Path(f"{self.generic_path}/Output"))
        self.screenshots_path = str(Path(f"{self.generic_path}/Screenshots"))
        self.workings_path = str(Path(f"{self.generic_path}/Workings"))

        # Logging
        self.log_file_path = f"{self.generic_path}/{self.name}_log_for_{self.date_stamp}.txt"
        self.main_logger = MyLogger(self.log_file_path)
        self.main_logger.log(level=logging.INFO, message="Class instantiated")

        # SAP variables S_PL0_xxx
        self.t_code_sap = "S_PL0_86000030"
        self.variant_sap = "GIT_AUTO"

        # SAP variables spool
        self.t_code_sap_spool = "sm37"
        self.job_name = "*"
        self.sap_person_name = "*"
        self.abap_prog_name = "ZR2RI_SIT_OUTBOUND"
            # path extracr
        #self.extract_path_spool = fr"C:\\Users\\{self.curr_user}\\Desktop\\Automations\\CPT_TP\\GIT\\Output\\"
        self.extract_path_spool = self.output_path #convert to str to avoid WindowsPath issue
        self.extract_name_spool = f"spool_{self.name}_for_{self.date_stamp}.txt"

        # SAP variables faglb03
        self.t_code_sap_faglb03 = "faglb03"
        self.variant_sap_faglb03 = "GIT_AUTO"

        # Paths to files
            # model file
        self.model_file_path = str(Path(f"{self.generic_path}/{self.name}_model_file.xlsb"))
        self.new_file_path = str(Path(f"{self.output_path}/{self.name}_for_{self.date_stamp}.xlsb"))
            # spool file path
        self.spool_file_path = str(Path(f"{self.output_path}/{self.extract_name_spool}"))
            # S_PL0_86000030 file path
        self.spl0_file_path = str(Path(f"{self.output_path}/SP_report_for_{self.date_stamp}"))
            # historical cost file
        self.final_path_to_historical_cost = f"S:/Robotics_COE_Prod/RPA/MM60/Outputs/{self.year}/{self.period_0}/"
        self.historical_cost_file = f"Historical cost AP{self.period_0} LE2941.xlsm"
        self.final_joined_path_to_historical_cost = str(Path(f"{self.final_path_to_historical_cost}{self.historical_cost_file}"))
            # alternative path to historical cost file
        self.final_path_to_historical_cost_alt = f"S:/Robotics_COE_Prod/RPA/MM60/Outputs/{self.year}/{str(int(self.period_0) - 1).zfill(2)}/" # ensure there are two digits
        self.final_joined_path_to_historical_cost_alt = str(Path(f"{self.final_path_to_historical_cost_alt}{self.historical_cost_file}"))
                # same but on my local drive
        self.new_file_path_to_historical_cost = str(Path(f"{self.output_path}/{self.historical_cost_file}"))
            # screenshots
        self.spl_scr = f"spl_scr_for_{self.date_stamp}"
        self.faglb03_scr = f"faglb03_scr_for_{self.date_stamp}"
        self.spool_scr_1 = f"spool_1_scr_for_{self.date_stamp}"
        self.spool_scr_2 = f"spool_2_scr_for_{self.date_stamp}"
            # list them
        self.list_screenshots = [self.spl_scr, self.faglb03_scr, self.spool_scr_1]

	# alternative way to instantiate the class (alternative constructor)
    @classmethod
    def from_string(cls, string):
        one, two, three = string.split("-")
        return cls()

    @MainControl.timer_decorator
    def __call__(self, *args, **kwargs):
        #self.run_logic_sap_spool(self.t_code_sap_spool) # this function will extract spool from SAP
        #self.run_logic_sap_SP(self.t_code_sap) # this function will perform S_PL0_86000030 part, extracting the file before saving it with a timestamp
        self.sap_faglb03(self.t_code_sap_faglb03)
        self.run_logic_HC() # extract historical cost file from PRA folder 
        self.create_final_file()
        x = 1+1

    @MainControl.timer_decorator
    def create_final_file(self):
        #--Reworking spool file: 
        self.main_logger.log(level=logging.INFO, message=f"Start parsing spool .txt file under {self.spool_file_path}")
        # looks like a very nasty .txt file, which requires specific approach
        self.df_sap = pd.read_csv(self.spool_file_path, 
                            on_bad_lines='skip', sep="\t", encoding="ANSI",skiprows=13, skipinitialspace = True)
        self.df_sap = self.df_sap.loc[:, ~self.df_sap.columns.str.contains('^Unnamed')] # drop all unnamed columns -> ~ stands for bool
        self.df_sap = self.df_sap[self.df_sap['Plant'] != "Plant"] # drop all rows that have "Plant" in their names (those are repetitions of headers)
        self.df_sap.dropna(subset=['Plant'], inplace=True)
        # convert respective columns to numeric
        self.df_sap['Quantity'] = self.df_sap['Quantity'].str.replace(',', '').astype(float)
        self.df_sap['Amount in LC'] = self.df_sap['Amount in LC'].str.replace(',', '').astype(float)
        self.df_sap['Net Order Value in PO Curr.'] = self.df_sap['Net Order Value in PO Curr.'].str.replace(',', '').astype(float)
        self.df_sap['PO Quantity'] = self.df_sap['PO Quantity'].str.replace(',', '').astype(float)

        # export to excel
        self.git_tab_path = f"C:/Users/{self.curr_user}/Desktop/Automations/CPT_TP/GIT/Output/GIT_tab_for_{self.date_stamp}.xlsx"
        self.df_sap.to_excel(self.git_tab_path, index=False)
        self.main_logger.log(level=logging.INFO, message=f"End parsing spool .txt file. Save under {self.git_tab_path}")

        #--Data Reworked APxx-xx tab:
        self.df_sap_le = self.df_sap.copy() # creating a copy of the file with only 4 LEs 
        self.df_sap_le = self.df_sap_le[(self.df_sap_le['Company Code'] == '2941') | (self.df_sap_le['Company Code'] == '2942') | 
                            (self.df_sap_le['Company Code'] == '2946') | (self.df_sap_le['Company Code'] == '2951')]
        # reset index:
        self.df_sap_le = self.df_sap_le.reset_index(drop=True)
        # add index column:
        self.df_sap_le['INDEX'] = self.df_sap_le.index + 2
        # adding columns:
        # concatenate:
        self.df_sap_le.insert(loc=0, column='Concatenate', value=self.df_sap_le['Material Number'] + "-" + self.df_sap_le['Plant'])
        # other:
        self.df_sap_le['BUoM historic'] = "=VLOOKUP(A" + self.df_sap_le['INDEX'].astype(str) + ",MM60!A:I,7,FALSE)"
        self.df_sap_le['Std price per BUoM'] = "=VLOOKUP(A" + self.df_sap_le['INDEX'].astype(str) + ",MM60!A:I,9,FALSE)"
        self.df_sap_le['Value at historical cost'] = "=IF(AJ" + self.df_sap_le['INDEX'].astype(str) + "=AC" + self.df_sap_le['INDEX'].astype(str) + ",0,D" + self.df_sap_le['INDEX'].astype(str) + "*AO" + + self.df_sap_le['INDEX'].astype(str) + ")"
        self.df_sap_le['Test UOM'] = "=AN" + self.df_sap_le['INDEX'].astype(str) + "=E" + self.df_sap_le['INDEX'].astype(str)
        self.df_sap_le['Diff $'] = "=IF(AJ" + self.df_sap_le['INDEX'].astype(str) + "=AC" + self.df_sap_le['INDEX'].astype(str) + ",0,AP" + self.df_sap_le['INDEX'].astype(str) + "-F" + self.df_sap_le['INDEX'].astype(str) + ")"
        # drop index colums:
        self.df_sap_le = self.df_sap_le.drop('INDEX', axis=1)
        # export to excel:
        self.reworked_tab_path = f"C:/Users/{self.curr_user}/Desktop/Automations/CPT_TP/GIT/Output/Reworked_tab_for_{self.date_stamp}.xlsx"
        self.df_sap_le.to_excel(self.reworked_tab_path, index=False) 
    # working with RPA mm60 file
        try:
            self.df_mm60 = pd.read_excel(self.new_file_path_to_historical_cost, 
                        sheet_name="MM60 Report", converters={'Price':float})
            pd.options.display.float_format = '{:20,.2f}'.format # handling scientific notation
            # add columns:
            self.df_mm60.insert(loc=0, column='Concatenate', value=self.df_mm60['Material'] + "-" + self.df_mm60['Plant'])
            self.df_mm60['Std price per unit'] = self.df_mm60['Price'] / self.df_mm60['Price unit']
            self.mm60_tab_path = f"C:/Users/{self.curr_user}/Desktop/Automations/CPT_TP/GIT/Output/mm60_tab_for_{self.date_stamp}.xlsx"
            self.df_mm60.to_excel(self.mm60_tab_path, index=False)
        except Exception as e:
            self.main_logger.log(level=logging.ERROR, message=f"Something went wring with converting the historical file, namely {e}")

            '''# Working with Roel file
            # df_roel = pd.read_excel(roel_file,
            #             converters={'BusA': str})
            # # remove all unwanted columns:
            # df_roel = df_roel[['Plnt', 'Material', 'BusA', 'Standard price', 'Crcy', 'BUn', 'per']]
            # # add useful columns:
            # df_roel.insert(loc=0, column='Concatenate', value=df_roel['Material'] + "-" + df_roel['Plnt'])
            # df_roel['Std price per unit'] = df_roel['Standard price'] / df_roel['per']
            # # drop last row:
            # df_roel.drop(df_roel.tail(1).index,inplace=True) # drop last (n) rows
            # mm60_tab_path = f"C:/Users/{curr_user}/Desktop/GIT_BOT/Workings/mm60_tab.xlsx"
            # df_roel.to_excel(mm60_tab_path, index=False)'''
# activating Excel:
        self.excel = win32.Dispatch("Excel.Application")
        self.excel.AskToUpdateLinks = False
        self.excel.DisplayAlerts = False
        self.excel.Visible = True 
        self.excel.ScreenUpdating = False
#--Active sheet on temporary extract:
        self.git_tab_wb = self.excel.Workbooks.Open(self.git_tab_path)
        self.git_tab_ws = self.git_tab_wb.Sheets("Sheet1") 
        # new file
        self.output_wb = self.excel.Workbooks.Open(self.new_file_path) # open the new file to paste all the data there
        self.output_ws = self.output_wb.Sheets("GIT") # start from GIT sheet
        self.output_ws.Range("A:AL").Delete() # clean the space
        self.git_tab_ws.Range('A1').CurrentRegion.Copy(Destination = self.output_ws.Range('A1')) # paste new info
        self.output_ws.Cells.Replace("=", "=") # replace = with = to make formulas work
        self.git_tab_wb.Close() # close git_tab excel
# working with mm60 tab
        self.mm60_tab_wb = self.excel.Workbooks.Open(self.mm60_tab_path) # which file to open
        self.mm60_tab_ws = self.mm60_tab_wb.Sheets("Sheet1") # which sheet to open
        self.output_ws = self.output_wb.Sheets("MM60") # open sheet to paste the data
        self.output_ws.Range("A:I").Delete() # clean the space
        self.mm60_tab_ws.Range('A1').CurrentRegion.Copy(Destination = self.output_ws.Range('A1')) # paste the data
        self.output_ws.Cells.Replace("=", "=") # this one is for formulas to work
        self.mm60_tab_wb.Close() # close wb
# working with spl tab S_PL0_86000030
        self.spl_tab_wb = self.excel.Workbooks.Open(self.spl0_file_path) # which file to open
        self.spl_tab_ws = self.spl_tab_wb.Sheets("Sheet1")  # which sheet to open
        self.output_ws = self.output_wb.Sheets("S_PL0_86000030") # open sheet to paste the data
        self.output_ws.Range("A2:AN100000").Delete() # clean the space, spare 1st row (headers)
        self.spl_tab_ws.Range('A3:AN100000').Copy(Destination = self.output_ws.Range('A2')) # paste the data (from A3 because of the headers)
        self.spl_tab_wb.Close() # close wb
# working with reworked tab
        self.reworked_tab_wb = self.excel.Workbooks.Open(self.reworked_tab_path) # which file to open
        self.reworked_tab_ws = self.reworked_tab_wb.Sheets("Sheet1")  # which sheet to open
        self.output_ws = self.output_wb.Sheets("Reworked") # open sheet to paste the data
        self.output_ws.Range("A2:AR100000").Delete() # clean the space
        self.reworked_tab_ws.Range('A2:AR100000').Copy(Destination = self.output_ws.Range('A2')) # paste the new data
        self.output_ws.Cells.Replace("=", "=") # this one is for formulas to work
        self.reworked_tab_wb.Close() # close reworked wb, for it's no longer needed
# pasting screenhosts
        self.screenshots_ws = self.output_wb.Sheets("Screenshots")
        self.add_screenshots(worksheet=self.screenshots_ws, list_screenshots=self.list_screenshots, picture_path=self.screenshots_path)
# final steps
        self.output_wb.RefreshAll() # refesh all the pivots we have in the workbook
        self.excel.ScreenUpdating = True
        self.output_wb.Close(True) # True is for saving the file
        self.main_logger.log(level=logging.INFO, message=f"End of final control, all files are populated")
        close_excel()
        self.send_email(subject="Automated report for GIT") 
        
    @MainControl.timer_decorator
    def run_logic_HC(self): # extract historical cost from the shared drive (RPA hc)
        # copy model file
        self.copy_file(self.model_file_path, self.new_file_path)
        # copying hc file
        try:
            self.copy_file(self.final_joined_path_to_historical_cost, self.new_file_path_to_historical_cost)
        except:
            self.copy_file(self.final_joined_path_to_historical_cost_alt, self.new_file_path_to_historical_cost)

    @MainControl.timer_decorator
    @MainControl.sap_decorator
    def run_logic_sap_spool(self):
        try:
            sap_enter_spool(session=self.session, job_name=self.job_name, sap_user_name=self.sap_person_name, abap_prog_name=self.abap_prog_name)
            self.main_logger.log(level=logging.INFO, message=f"Spool has been entered. Job name {self.job_name} is found")
            time.sleep(5)
            self.take_screenshot(name=self.spool_scr_1, path=self.screenshots_path)
            sap_extract_txt(session=self.session, extr_path=self.extract_path_spool, extr_name=self.extract_name_spool)
            self.main_logger.log(level=logging.INFO, message=f"Spool {self.extract_name_spool} is extracted and stored under {self.extract_path_spool}.")
            self.main_logger.log(level=logging.INFO, message=f"Spool is executed. Proceed with closing decorator")
        except Exception as e:
            self.main_logger.log(level=logging.ERROR, message=f"There is an issue with spool.\nIssue: {e}")
            
    @MainControl.timer_decorator		
    @MainControl.sap_decorator
    def run_logic_sap_SP(self):
        try:
            try:
                sap_variant(session=self.session, var_to_use=self.variant_sap)
            except: 
                sap_variant(session=self.session, var_to_use=self.variant_sap, version=2) # in our case, V2 is the most probable scenario
            self.main_logger.log(level=logging.INFO, message=f"Applied {self.variant_sap} for {self.t_code_sap}")
			# adjust parameters
            self.session.findById("wnd[0]/usr/ctxtPAR_01").text = self.period
            self.session.findById("wnd[0]/usr/ctxtPAR_02").text = self.period
            self.session.findById("wnd[0]/usr/ctxtPAR_06").text = self.year
            self.main_logger.log(level=logging.INFO, message=f"Applied following parameters {self.period}, {self.period} and {self.year} for {self.t_code_sap}")
            # run the report
            sap_run(session=self.session)
            self.main_logger.log(level=logging.INFO, message=f"Executed {self.t_code_sap}. Output is generated")
            # see if whether it was saved before
            try:
                self.session.findById("wnd[1]/usr/sub:SAPLKEC1:0110/radCEC01-CHOICE[1,0]").select()
                self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
            except: pass
            # extract the data
            self.session.findById("wnd[0]/usr/lbl[1,1]").setFocus()
            self.take_screenshot(name=self.spl_scr, path=self.screenshots_path)
            self.session.findById("wnd[0]").sendVKey(2)
            self.session.findById("wnd[0]/usr/lbl[1,11]").setFocus()
            self.session.findById("wnd[0]").sendVKey(2)
            self.session.findById("wnd[0]/tbar[1]/btn[47]").press()
            self.session.findById("wnd[1]/usr/btnD2000_PUSH_01").press()
            self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
            self.session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[0,0]").select()
            self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
            self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
            
            self.main_logger.log(level=logging.INFO, message=f"Trying to locate SP extract")
            self.save_extract_sap() # run the logic to save Excel file generated
            self.main_logger.log(level=logging.INFO, message=f"End of SP control. Proceed with closing decorator")
        except Exception as e:
            self.main_logger.log(level=logging.ERROR, message=f"An issue occured while performing SP part of the control.\nIssue: {e}")
    
    def save_extract_sap(self):
        self.target_workbook_name = "Worksheet in Basis (1)"
        workbook = self.get_workbook_by_name()

        if workbook: # if True is returned
            self.main_logger.log(level=logging.INFO, message=f"Workbook {workbook.Name} is found")
            # save this workbook
            workbook.SaveAs(Filename=f"C:\\Users\\{self.curr_user}\\Desktop\\Automations\\CPT_TP\\GIT\\Output\\SP_report_for_{self.date_stamp}")
            # close excel
            self.excel.DisplayAlerts = True
            close_excel()
            # close sap
            sap_close()
        else: 
            self.main_logger.log(level=logging.ERROR, message=f"No workbook named '{self.target_workbook_name}' is found")

    def get_workbook_by_name(self):
        try:
            self.excel = win32.GetActiveObject("Excel.Application")
            self.workbooks = self.excel.Workbooks
            self.excel.DisplayAlerts = False

            for workbook in self.workbooks:
                print(workbook.Name)
                if workbook.Name == self.target_workbook_name: return workbook
            return None
        except Exception as e:
            self.main_logger.log(level=logging.ERROR, message=f"Excel not found, details: {e}")
            return None

    @MainControl.timer_decorator
    @MainControl.sap_decorator
    def sap_faglb03(self):
        # try to apply variant
        self.main_logger.log(level=logging.INFO, message=f"Applying {self.variant_sap_faglb03} variant in SAP")
        try:
            sap_variant(session=self.session, var_to_use=self.variant_sap_faglb03)
        except: 
            sap_variant(session=self.session, var_to_use=self.variant_sap_faglb03, version=2) # in our case, V2 is the most probable scenario
        # adjust parameters
        self.session.findById("wnd[0]/usr/txtRYEAR").text = self.year
        # run the report
        sap_run(session=self.session)
        self.main_logger.log(level=logging.INFO, message="Report is executed")
        # take balance
        balance_value = self.session.findById("wnd[0]/usr/cntlFDBL_BALANCE_CONTAINER/shellcont/shell").GetCellValue(f"{self.period}", "BALANCE")
        self.take_screenshot(name=self.faglb03_scr, path=self.screenshots_path) # take screenshot of faglb03
        self.session.findById("wnd[0]/usr/cntlFDBL_BALANCE_CONTAINER/shellcont/shell").setCurrentCell(f"{self.period}","BALANCE")
        self.main_logger.log(level=logging.INFO, message=f"End of the main part of control. Proceed to closing decorator")

if __name__ == "__main__": 
    control = GIT()
    control()
    #print(control)
	#control = Control.from_string("hi-there-hi") # if we need to instantiate a class with same arguments but in different way
	#control()
    #control.faglb03_scr("faglb03")