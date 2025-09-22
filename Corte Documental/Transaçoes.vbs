print = "Z5M1"

Sub Cogi(rowId)



Conexao_SAP("COGI")	
	
    ' Conexão SAP e execução do script
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
	
'session.findById("wnd[0]").maximize
session.findById("wnd[0]/usr/ctxtS_WERKS-LOW").text = rowId
session.findById("wnd[0]/usr/ctxtS_WERKS-LOW").caretPosition = 3
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]").sendVKey 8	

Call COGI_0()

Session.findById("wnd[0]/usr/lbl[4,6]").SetFocus
Session.findById("wnd[0]/tbar[0]/btn[86]").press

Session.findById("wnd[1]/usr/ctxtPRI_PARAMS-PDEST").Text = print
Session.findById("wnd[1]/usr/subSUBSCREEN:SAPLSPRI:0600/ctxtPRI_PARAMS-PAART").Text = ""
Session.findById("wnd[1]/usr/subSUBSCREEN:SAPLSPRI:0600/cmbPRIPAR_DYN-PRIMM").SetFocus
Session.findById("wnd[1]/usr/subSUBSCREEN:SAPLSPRI:0600/cmbPRIPAR_DYN-PRIMM").Key = "X"
Session.findById("wnd[2]/tbar[0]/btn[0]").press
Session.findById("wnd[1]/tbar[0]/btn[13]").press
Session.findById("wnd[0]/tbar[0]/btn[86]").press
Session.findById("wnd[1]/usr/ctxtPRI_PARAMS-PDEST").Text = "ZPDF"
Session.findById("wnd[1]/usr/subSUBSCREEN:SAPLSPRI:0600/ctxtPRI_PARAMS-PAART").Text = ""
Session.findById("wnd[1]/tbar[0]/btn[13]").press
Session.findById("wnd[2]/tbar[0]/btn[0]").press
Session.findById("wnd[0]/tbar[0]/btn[3]").press
Session.findById("wnd[0]/tbar[0]/btn[3]").press

    
End Sub

Sub COGI_0()

    On Error Resume Next

    ' Tenta focar no elemento da tela SAP
	session.findById("wnd[0]/usr/lbl[4,6]").setFocus
	session.findById("wnd[0]/usr/lbl[4,6]").caretPosition = 4

    If Err.Number <> 0 Then
        MsgBox "Nenhuma Entrada Encontrada Para Seleção! Salve um print para confecção do book.", vbInformation, "CORTE DOCUMENTAL - COGI"
        Err.Clear
        Exit Sub
    End If

End Sub




