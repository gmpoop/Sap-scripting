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
session.findById("wnd[0]").sendVKey 4
session.findById("wnd[1]").sendVKey 12
session.findById("wnd[0]/usr/ctxtRMMG1-MATNR").text = "12391237182e"
session.findById("wnd[0]/usr/ctxtRMMG1-MATNR").caretPosition = 12
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/cmbRMMG1-MTART").key = "FERT"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/cmbRMMG1-MTART").key = "FGTR"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/cmbRMMG1-MBRSH").key = "M"
session.findById("wnd[0]/usr/cmbRMMG1-MTART").key = "LGUT"
session.findById("wnd[0]/usr/cmbRMMG1-MTART").setFocus
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/cmbRMMG1-MTART").key = "MCRO"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/cmbRMMG1-MTART").key = "VERP"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtRMMG1-MATNR").text = "12391237182"
session.findById("wnd[0]/usr/ctxtRMMG1-MATNR").setFocus
session.findById("wnd[0]/usr/ctxtRMMG1-MATNR").caretPosition = 11
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtRMMG1-MATNR").setFocus
session.findById("wnd[0]/usr/ctxtRMMG1-MATNR").caretPosition = 11
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtRMMG1-MATNR").text = "1239123"
session.findById("wnd[0]/usr/ctxtRMMG1-MATNR").setFocus
session.findById("wnd[0]/usr/ctxtRMMG1-MATNR").caretPosition = 7
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(0).selected = true
session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW/txtMSICHTAUSW-DYTXT[0,0]").caretPosition = 12
session.findById("wnd[1]/tbar[0]/btn[0]").press
