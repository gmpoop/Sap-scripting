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
session.findById("wnd[0]/tbar[0]/okcd").text = "CR01"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]").sendVKey 4
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/lbl[1,5]").setFocus
session.findById("wnd[1]/usr/lbl[1,5]").caretPosition = 4
session.findById("wnd[1]").sendVKey 2
session.findById("wnd[0]/usr/txtT001W-NAME1").setFocus
session.findById("wnd[0]/usr/txtT001W-NAME1").caretPosition = 0
session.findById("wnd[0]").sendVKey 0
