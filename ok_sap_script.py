from datetime import date, timedelta, datetime, timezone

def sap_open():
    import subprocess
    import time
    subprocess.Popen(['C:/Program Files (x86)/SAP/FrontEnd/SAPgui/saplogon.exe'])
    #time.sleep(5)
    return True

def sap_close():
    import subprocess
    subprocess.call(["TASKKILL", "/F", "/IM", "saplogon.exe"], shell=True)
    return True

def sap_logon(environment:str, client=1):
    import win32com.client
    SapGuiAuto = win32com.client.GetObject("SAPGUI")
    application = SapGuiAuto.GetScriptingEngine
    connection = application.Openconnection(environment, True)
    session = connection.Children(0)
    try:
        session.findById("wnd[0]").maximize()
        session.findById(f"wnd[0]/usr/tblSAPMSYSTTC_IUSRACL/btnIUSRACL-BNAME[{client},0]").setFocus()
        session.findById(f"wnd[0]/usr/tblSAPMSYSTTC_IUSRACL/btnIUSRACL-BNAME[{client},0]").press()
        return session
    except: return session
    
def sap_code(tcode:str, session) -> bool:
    """Enter SAP transaction \n
    tcode - required argument in str format \n
    session - return value of sap_logon function
    """
    import win32com.client
    try:
        session.findById("wnd[0]/tbar[0]/okcd").text = tcode
        session.findById("wnd[0]").sendVKey(0)
        return True
    except: return False

def sap_layout(session, layout:str):
    """Select layout after the report is executed, e.g. apply some filters, etc. \n
    session - returned value of sap_logon function \n
    layout - layout to be selected \n
    """
    session.findById("wnd[0]/tbar[1]/btn[33]").press()
        #--Add "/" for development -> user-specific:
    variant = session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell")
    rows = variant.rowCount
    for _ in range(rows): # loop through the rows until correct variant is found: # for some reason I had (rows-1)
        layout_variant = variant.getCellValue(_, "VARIANT")
        if layout_variant == layout:
            variant.currentCellRow = _
            variant.clickCurrentCell()
            break
    
def sap_variant(session, var_to_use: str, created_by: str ="*", version=1):
    """Select variant. Suitable for faglb03, fagll03, etc. \n
    session - returned value of sap_logon function \n
    var_to_use - variant to apply
    created_by - user, who created the variant, defaults to "*" \n
    version - how to select version, defaults to 1, choose 2 if 1 fails
    """
    import win32com.client # by default
    valid = [1,2]
    if version not in valid: 
        raise ValueError("version must be 1 or 2") # limit the user from choosing invalid variants
    
    session.findById("wnd[0]/tbar[1]/btn[17]").press() # click variant button
    match version:
        case 1:
            session.findById("wnd[1]/usr/txtV-LOW").text = var_to_use # variant to select
            session.findById("wnd[1]/usr/txtENAME-LOW").text = created_by # who created the variant
            session.findById("wnd[1]/tbar[0]/btn[8]").press() 
            return version, True
        
        case 2:
            try:
                session.findById("wnd[1]/usr/txtENAME-LOW").text = created_by # who created the variant
                session.findById("wnd[1]/tbar[0]/btn[8]").press() 
                variant = session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell")
                rows = variant.rowCount
                for _ in range(rows): # loop through the rows until correct variant is found: # for some reason I had (rows-1)
                    layout_variant = variant.getCellValue(_, "VARIANT")
                    if layout_variant == var_to_use:
                        session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").SelectedRows = _
                        session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").doubleClickCurrentCell()
                        break
                return version, True
            except: 
                variant = session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell")
                rows = variant.rowCount
                for _ in range(rows): # loop through the rows until correct variant is found: # for some reason I had (rows-1)
                    layout_variant = variant.getCellValue(_, "VARIANT")
                    if layout_variant == var_to_use:
                        session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").SelectedRows = _
                        session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").doubleClickCurrentCell()
                        break
                return version, False    
            
def sap_run(session) -> None:
    """Run any report \n
    session - returned value of sap_logon function 
    """
    import win32com.client # by default
    try: session.findById("wnd[0]").sendVKey(8)
    except: session.findById("wnd[1]/tbar[0]/btn[8]").press()

def sap_extract(session, extr_path: str, extr_name: str) -> str:
    """Run any report \n
    This function will download sap extract in .xlsx format or replace it if already exists \n
    session - returned value of sap_logon function \n
    extr_path - where to extract (has to be entered in a \ way) \n
    extr_name - how to call the extract
    """
    import win32com.client # by default
    try:
        session.findById("wnd[0]/mbar/menu[0]/menu[3]/menu[1]").select()
    except: pass
    try:
        session.findById("wnd[0]/mbar/menu[0]/menu[10]/menu[3]/menu[1]").select()
    except: pass
    try: # this one is for when you need to extract via Export button
        session.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell").pressToolbarContextButton("&MB_EXPORT")
        session.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell").selectContextMenuItem("&XXL")
    except: pass
    try:
        session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[1]").select()
    except: pass
    session.findById("wnd[1]/usr/ctxtDY_PATH").text = extr_path
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = extr_name
    session.findById("wnd[1]/tbar[0]/btn[0]").press() # try to generate the report
    # generate or overwrite 
    msg = session.FindById("wnd[0]/sbar").Text
    if "already" in msg and "exists" in msg:
        session.findById("wnd[1]/tbar[0]/btn[11]").press() # replace existing report
    return msg

def sap_extract_txt (session, extr_path: str, extr_name: str) -> str:
    """Run any report \n
    This function will download sap extract in .txt format or replace it if already exists \n
    It has some proplems associated with Windows native dialogues \n
    So only possible in .txt \n
    session - returned value of sap_logon function \n
    extr_path - where to extract (has to be entered in a \ way) \n
    extr_name - how to call the extract
    """
    import win32com.client # by default
    try:
        session.findById("wnd[0]/tbar[1]/btn[45]").press() # depending on where this button is located
    except:
        session.findById("wnd[0]/tbar[1]/btn[48]").press() # spool button location
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select()
    session.findById("wnd[1]").sendVKey(0)
    session.findById("wnd[1]/usr/ctxtDY_PATH").text = extr_path
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = extr_name
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[1]/tbar[0]/btn[0]").press() # try to generate the report
    # generate or overwrite 
    msg = session.FindById("wnd[0]/sbar").Text
    if "already" in msg and "exists" in msg:
        session.findById("wnd[1]/tbar[0]/btn[11]").press() # replace existing report
    return msg

def sap_enter_spool (session, job_name: str ="*", sap_user_name: str ="*", from_spool =date.today().strftime("01.%m.%Y"), to_spool =date.today().strftime("%d.%m.%Y"), abap_prog_name: str="*", number_hits: str="100000"):
    '''
    Most variables have default values\n
    Works if you only have 1 single spool output
    '''
    session.findById("wnd[0]/usr/txtBTCH2170-JOBNAME").text = job_name
    session.findById("wnd[0]/usr/txtBTCH2170-USERNAME").text = sap_user_name
    session.findById("wnd[0]/usr/ctxtBTCH2170-FROM_DATE").text = from_spool
    session.findById("wnd[0]/usr/ctxtBTCH2170-TO_DATE").text = to_spool
    session.findById("wnd[0]/usr/txtBTCH2170-ABAPNAME").text = abap_prog_name
    sap_run(session=session)
    session.findById("wnd[0]/usr/lbl[37,13]").setFocus()
    session.findById("wnd[0]").sendVKey(2)
    session.findById("wnd[0]/usr/lbl[14,3]").setFocus()
    session.findById("wnd[0]").sendVKey(2)
    session.findById("wnd[0]/tbar[1]/btn[46]").press()
    session.findById("wnd[1]/usr/txtDIS_TO").text = number_hits
    session.findById("wnd[1]/usr/txtDIS_TO").setFocus()
    session.findById("wnd[1]/usr/txtDIS_TO").caretPosition = 7
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[2]/tbar[0]/btn[0]").press()

def close_excel():
    import subprocess
    subprocess.call(["TASKKILL", "/F", "/IM", "excel.exe"], shell=True)
    return True

def sap_screen_nagivation(session, action: str):
    """Run any report \n
    This function will perform a shortcut \n
    session - returned value of sap_logon function \n
    action - shortcut to be used in SAP
    This will NOT work with table view, we need list output -> but I will try to amend
    """
    try: # try to convert to the list output
        session.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell").pressToolbarContextButton("&MB_VIEW")
        session.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell").selectContextMenuItem("&PRINT_BACK_PREVIEW")
    except: pass
    match action:
        case "down":
            session.findById("wnd[0]").sendVKey(83) # go to the end of a page
        case "up":
            session.findById("wnd[0]").sendVKey(80) # go to the top of a page
        case _:
            return False
