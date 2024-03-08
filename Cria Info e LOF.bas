Attribute VB_Name = "Módulo1"
Sub ME11_RegInfo()
On Error GoTo TrataErro

    erro = "D"
    Set SapGuiAuto = GetObject("SAPGUI")  'Get the SAP GUI Scripting object
    Set SAPApp = SapGuiAuto.GetScriptingEngine 'Get the currently running SAP GUI
    Set SAPCon = SAPApp.Children(0) 'Get the first system that is currently connected
    Set Session = SAPCon.Children(0) '
    linha = 2
    Do While Range("A" & linha).Text <> ""
            If Range(erro & linha).Text = "" Then
                 Session.findById("wnd[0]").maximize
                 Session.findById("wnd[0]/tbar[0]/okcd").Text = "/nme11"
                 Session.findById("wnd[0]").sendVKey 0
                 Session.findById("wnd[0]/usr/ctxtEINA-LIFNR").Text = Range("b" & linha).Text 'FORNECEDOR
                 Session.findById("wnd[0]/usr/ctxtEINA-MATNR").Text = Range("A" & linha).Text 'MATERIAL
                 Session.findById("wnd[0]/usr/ctxtEINE-EKORG").Text = "ocal"
                 Session.findById("wnd[0]/usr/ctxtEINE-WERKS").Text = Range("c" & linha).Text 'CENTRO
                 Session.findById("wnd[0]/usr/radRM06I-NORMB").Select
                 Session.findById("wnd[0]").sendVKey 0
                 Session.findById("wnd[0]").sendVKey 0
                 Session.findById("wnd[0]").sendVKey 0
                 Session.findById("wnd[0]/usr/txtEINE-APLFZ").Text = "2"
                 Session.findById("wnd[0]/usr/ctxtEINE-EKGRP").Text = "800"
                 Session.findById("wnd[0]/usr/txtEINE-NORBM").Text = "1"
                 Session.findById("wnd[0]/usr/txtEINE-NETPR").Text = "1"
                 Session.findById("wnd[0]/usr/ctxtEINE-MWSKZ").Text = "p3"
                 Session.findById("wnd[0]/usr/chkEINE-UEBTK").Selected = True
                 Session.findById("wnd[0]/usr/ctxtEINE-INCO1").Text = "SFR"
                 Session.findById("wnd[0]/usr/txtEINE-INCO2").Text = "SFR"
                
                 Session.findById("wnd[0]").sendVKey 0
                 Session.findById("wnd[0]/tbar[0]/btn[11]").press
                 Range(erro & linha).FormulaR1C1 = Now
                 Range(erro & linha).Select
            End If
            linha = linha + 1
    Loop
Exit Sub
TrataErro:

errMessage = Session.findById("wnd[0]/sbar").Text
Range(erro & linha).FormulaR1C1 = errMessage

Call ME11_RegInfo
End Sub

Sub Me01_LOF()
On Error GoTo TrataErro

    erro = "E"
    Set SapGuiAuto = GetObject("SAPGUI")  'Get the SAP GUI Scripting object
    Set SAPApp = SapGuiAuto.GetScriptingEngine 'Get the currently running SAP GUI
    Set SAPCon = SAPApp.Children(0) 'Get the first system that is currently connected
    Set Session = SAPCon.Children(0) '

linha = 2
Do While Range("A" & linha).Text <> ""
    If Range(erro & linha).Text = "" Then
        Session.findById("wnd[0]").maximize
        Session.findById("wnd[0]/tbar[0]/okcd").Text = "/nme01"
        Session.findById("wnd[0]").sendVKey 0
        Session.findById("wnd[0]/usr/ctxtEORD-MATNR").Text = Range("A" & linha).Text 'MATERIAL
        Session.findById("wnd[0]/usr/ctxtEORD-WERKS").Text = Range("c" & linha).Text 'CENTRO
        Session.findById("wnd[0]").sendVKey 0
        Session.findById("wnd[0]/usr/tblSAPLMEORTC_0205/ctxtEORD-VDATU[0,0]").Text = "01012000"
        Session.findById("wnd[0]/usr/tblSAPLMEORTC_0205/ctxtEORD-BDATU[1,0]").Text = "31129999"
        Session.findById("wnd[0]/usr/tblSAPLMEORTC_0205/ctxtEORD-LIFNR[2,0]").Text = Range("B" & linha).Text 'FORNECEDOR
        Session.findById("wnd[0]/usr/tblSAPLMEORTC_0205/ctxtEORD-EKORG[3,0]").Text = "ocal"

        Session.findById("wnd[0]").sendVKey 0
        Session.findById("wnd[0]/tbar[0]/btn[11]").press
        Range(erro & linha).FormulaR1C1 = Now
        Range(erro & linha).Select
     End If
    linha = linha + 1
Loop

Exit Sub
TrataErro:

errMessage = Session.findById("wnd[0]/sbar").Text
Range(erro & linha).FormulaR1C1 = errMessage

Call Me01_LOF
End Sub


