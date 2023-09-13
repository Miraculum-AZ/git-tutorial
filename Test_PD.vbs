If Not IsObject(application) Then
   Set SapGuiAuto  = GetObject("SAPGUI")
   Set application = SapGuiAuto.GetScriptingEngine
End If
If Not IsObject(connection) Then
   Set connection = application.Children(0)
End If
If Not IsObject(session) Then
   Set session    = connection.Children(0)
End If
If IsObject(WScript) Then
   WScript.ConnectObject session,     "on"
   WScript.ConnectObject application, "on"
End If
session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "se16"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtDATABROWSE-TABLENAME").text = "mbew"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtI1-LOW").text = "10000000"
session.findById("wnd[0]/usr/ctxtI1-HIGH").text = "19999999"
session.findById("wnd[0]/usr/txtMAX_SEL").text = "10000000 "
session.findById("wnd[0]/usr/txtMAX_SEL").setFocus
session.findById("wnd[0]/usr/txtMAX_SEL").caretPosition = 10
session.findById("wnd[0]").sendVKey 8
session.findById("wnd[0]/mbar/menu[3]/menu[0]/menu[1]").select
session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").setCurrentCell 1,"TEXT"
session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").selectedRows = "1"
session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").clickCurrentCell
session.findById("wnd[0]/tbar[1]/btn[43]").press
session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\Users\KOMAROVO\Desktop\Automations\CPT_TP"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "MD_export.XLSX"
session.findById("wnd[1]/usr/ctxtDY_PATH").setFocus
session.findById("wnd[1]/usr/ctxtDY_PATH").caretPosition = 44
session.findById("wnd[1]/tbar[0]/btn[0]").press
