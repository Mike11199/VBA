Attribute VB_Name = "BD_Export_Check_Register_FCHN"
Public BDPass As String
Public BDUserName As String

Sub GL1130_RunFCHN_V1_LOG_IN_BOX()

On Error GoTo ErrSap


Dim ORF_WB As Workbook
Set ORF_WB = ThisWorkbook

SapConnectionString = ORF_WB.Sheets("Macro Input").Range("SAP_Connection")

BD_LOG_ON.Show

Dim StartTime As Double
Dim MinutesElapsed As String
'Remember time when macro starts
StartTime = Timer


'Create the GuiApplication object
Set SAP_Application = CreateObject("Sapgui.ScriptingCtrl.1")


'Open a connection in synchronous mode
Set Connection = SAP_Application.OpenConnection(SapConnectionString, True)
Set SAP_Session = Connection.Children(0)


'Allows for maximizing or minimizing the SAP Window
SessionHWND = SAP_Session.FindById("wnd[0]").Handle
ActivateWindow (SessionHWND)
SAP_Session.FindById("wnd[0]").Maximize



'============Log in to SAP===================================================
BDUserName = BD_LOG_ON.BDUserBox.Value
BDPass = BD_LOG_ON.BDPasswordBox.Value

Sleep 1
SAP_Session.FindById("wnd[0]/usr/txtRSYST-BNAME").Text = BDUserName
Sleep 1
SAP_Session.FindById("wnd[0]/usr/pwdRSYST-BCODE").Text = BDPass
Sleep 1
SAP_Session.FindById("wnd[0]").sendVKey 0
Sleep 1

Unload BD_LOG_ON
BDUserName = ""
BDPass = ""

'============================================================================



SAP_Session.FindById("wnd[0]/tbar[0]/okcd").Text = "/nFCHN "
Application.Wait (Now + TimeValue("00:00:01") / 1000)
SAP_Session.FindById("wnd[0]").sendVKey 0
Application.Wait (Now + TimeValue("00:00:01") / 1000)



Dim GLAccount As Long, FiscalYear As Long, ReconMonth As Variant, ReconMonth_Num As Long, FCHN_From As Variant, FCHN_To As Variant

GLAccount = ORF_WB.Sheets("Macro Input").Range("GL_Account")
FiscalYear = ORF_WB.Sheets("Macro Input").Range("Fiscal_Year")
ReconMonth = ORF_WB.Sheets("Macro Input").Range("Recon_Month")
ReconMonth_Num = ORF_WB.Sheets("Macro Input").Range("ReconMonth_Num")
FCHN_From = ORF_WB.Sheets("Macro Input").Range("FCHN_From")
FCHN_To = ORF_WB.Sheets("Macro Input").Range("FCHN_To")




SAP_Session.FindById("wnd[0]/usr/tabsTABSTRIP_CHK/tabpUCOMM1/ssub%_SUBSCREEN_CHK:RFCHKN10:0001/radPAR_EPOS").Select
SAP_Session.FindById("wnd[0]/usr/ctxtSEL_ZBUK-LOW").Text = "1000"
SAP_Session.FindById("wnd[0]/usr/ctxtSEL_HBKI-LOW").Text = "SCO"
SAP_Session.FindById("wnd[0]/usr/ctxtSEL_HKTI-LOW").Text = "ORF"
SAP_Session.FindById("wnd[0]/usr/tabsTABSTRIP_CHK/tabpUCOMM1/ssub%_SUBSCREEN_CHK:RFCHKN10:0001/radPAR_EPOS").SetFocus
SAP_Session.FindById("wnd[0]/usr/tabsTABSTRIP_CHK/tabpUCOMM2").Select
SAP_Session.FindById("wnd[0]/usr/tabsTABSTRIP_CHK/tabpUCOMM2/ssub%_SUBSCREEN_CHK:RFCHKN10:0002/ctxtSEL_ZALD-LOW").Text = FCHN_From
SAP_Session.FindById("wnd[0]/usr/tabsTABSTRIP_CHK/tabpUCOMM2/ssub%_SUBSCREEN_CHK:RFCHKN10:0002/ctxtSEL_ZALD-HIGH").Text = FCHN_To
SAP_Session.FindById("wnd[0]/tbar[1]/btn[8]").press

Sleep 5000
 
 
 '=====================CHANGE LAYOUT  TO FCHN WITH SAVED LAYOUT======================================================
 
SAP_Session.FindById("wnd[0]/tbar[1]/btn[33]").press
SAP_Session.FindById("wnd[1]/tbar[0]/btn[71]").press
SAP_Session.FindById("wnd[2]/usr/chkSCAN_STRING-START").Selected = False
SAP_Session.FindById("wnd[2]/usr/txtRSYSF-STRING").Text = "/FCHN_MACRO"
SAP_Session.FindById("wnd[2]/usr/chkSCAN_STRING-START").SetFocus
SAP_Session.FindById("wnd[2]/tbar[0]/btn[0]").press
SAP_Session.FindById("wnd[3]/usr/lbl[1,2]").SetFocus
SAP_Session.FindById("wnd[3]/usr/lbl[1,2]").CaretPosition = 6
SAP_Session.FindById("wnd[3]").sendVKey 2
SAP_Session.FindById("wnd[1]/tbar[0]/btn[0]").press

'===================================================================================================================
 





 
 '============EXPORT TO EXCEL FILE===========================================================================================
 

SAP_Session.FindById("wnd[0]/tbar[1]/btn[45]").press
SAP_Session.FindById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select
SAP_Session.FindById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").SetFocus
SAP_Session.FindById("wnd[1]/tbar[0]/btn[0]").press
SAP_Session.FindById("wnd[1]/usr/ctxtDY_PATH").Text = "C:\\TEMP"
SAP_Session.FindById("wnd[1]/usr/ctxtDY_FILENAME").Text = "Export3.txt"
SAP_Session.FindById("wnd[1]/usr/ctxtDY_PATH").SetFocus
SAP_Session.FindById("wnd[1]/usr/ctxtDY_PATH").CaretPosition = 4
SAP_Session.FindById("wnd[1]/tbar[0]/btn[11]").press
 
 
 '========================================================================================================================
 
 
Dim FCHN_Export_WB As Workbook
Application.Wait (Now + TimeValue("00:00:01"))

Set FCHN_Export_WB = Workbooks.Open("C:\TEMP\Export3.txt")

FCHN_Export_WB.Sheets(1).Copy After:=ORF_WB.Sheets("Macro Input")

ORF_WB.ActiveSheet.Name = ReconMonth & "_FCHN YTD"


With ActiveWorkbook.ActiveSheet.Tab
    .Color = 192
    .TintAndShade = 0
End With


Sleep 1
FCHN_Export_WB.Close SaveChanges:=False
Sleep 1


ActivateWindow (Application.hwnd)      'Bring Excel window to front
 



Set SAP_Application = Nothing

            
            Continuewiththis = MsgBox(Prompt:="Delete exported .MHTML file in C:\TEMP?" & vbNewLine & vbNewLine & "Selecting 'No' can cause Excel to crash.", Buttons:=vbQuestion + vbYesNo)
            If Continuewiththis = vbNo Then GoTo Skip_Deleting_Exports
            
                                Kill ("C:\TEMP\Export3.txt")
      
Skip_Deleting_Exports:          'Goes here and skips deleting export file if answer is no
  
 


MinutesElapsed = Format((Timer - StartTime) / 86400, "hh:mm:ss")
  
 MsgBox "Ran successfully in " & MinutesElapsed & " minutes." & vbNewLine & vbNewLine & _
"The macro has finished adding the exported file to you current workbook." & vbNewLine & vbNewLine & _
"Please press OK, and then close the alert message that opens after the macro ends."




Exit Sub
ErrSap:
MsgBox "Error.  Please press OK to end the macro."



End Sub



Sub Change_Layout_Manual()


''If error with layout for FCHN or it was deleted, can use this old manual way to change the layout



' '=====================CHANGE LAYOUT MANUALLY TO FCHN======================================================
'
'SAP_Session.FindById("wnd[0]/usr/lbl[55,16]").SetFocus
'SAP_Session.FindById("wnd[0]/usr/lbl[55,16]").CaretPosition = 10
'SAP_Session.FindById("wnd[0]/tbar[1]/btn[32]").press
'SAP_Session.FindById("wnd[1]/usr/btnAPP_FL_ALL").press
'
'
'
'SAP_Session.FindById("wnd[1]").sendVKey 42
'
'
'SAP_Session.FindById("wnd[1]/usr/btnB_SEARCH").press
''SAP_Session.FindById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "check number from to"
'SAP_Session.FindById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "CHECM"
'SAP_Session.FindById("wnd[2]").sendVKey 0
'SAP_Session.FindById("wnd[1]/usr/btnAPP_WL_SING").press
'
'
'
'
'SAP_Session.FindById("wnd[1]/usr/btnB_SEARCH").press
''SAP_Session.FindById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "payment document"
'SAP_Session.FindById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "VBLNR"
'SAP_Session.FindById("wnd[2]").sendVKey 0
'SAP_Session.FindById("wnd[1]/usr/btnAPP_WL_SING").press
'
'SAP_Session.FindById("wnd[1]/usr/btnB_SEARCH").press
''SAP_Session.FindById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "payment date"
'SAP_Session.FindById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "ZALDT"
'SAP_Session.FindById("wnd[2]/tbar[0]/btn[0]").press
'SAP_Session.FindById("wnd[1]/usr/btnAPP_WL_SING").press
'
'SAP_Session.FindById("wnd[1]/usr/btnB_SEARCH").press
''SAP_Session.FindById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "currency"
'SAP_Session.FindById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "WAERS"
'SAP_Session.FindById("wnd[2]/tbar[0]/btn[0]").press
'SAP_Session.FindById("wnd[1]/usr/btnAPP_WL_SING").press
'
'SAP_Session.FindById("wnd[1]/usr/btnB_SEARCH").press
''SAP_Session.FindById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "amount paid"
'SAP_Session.FindById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "RWBTR"
'SAP_Session.FindById("wnd[2]").sendVKey 0
'SAP_Session.FindById("wnd[1]/usr/btnAPP_WL_SING").press
'
'SAP_Session.FindById("wnd[1]/usr/btnB_SEARCH").press
''SAP_Session.FindById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "recipient"
'SAP_Session.FindById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "EMPFENTW"
'SAP_Session.FindById("wnd[2]").sendVKey 0
'SAP_Session.FindById("wnd[1]/usr/btnAPP_WL_SING").press
'
'SAP_Session.FindById("wnd[1]/usr/btnB_SEARCH").press
''SAP_Session.FindById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "date encash"
'SAP_Session.FindById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "BANCDS"
'SAP_Session.FindById("wnd[2]").sendVKey 0
'SAP_Session.FindById("wnd[1]/usr/btnAPP_WL_SING").press
'
'SAP_Session.FindById("wnd[1]/usr/btnB_SEARCH").press
''SAP_Session.FindById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "name of the payee"
'SAP_Session.FindById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "ZNME1"
'SAP_Session.FindById("wnd[2]").sendVKey 0
'SAP_Session.FindById("wnd[1]/usr/btnAPP_WL_SING").press
'
'SAP_Session.FindById("wnd[1]/usr/btnB_SEARCH").press
''SAP_Session.FindById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "vendor"
'SAP_Session.FindById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "LIFNR"
'SAP_Session.FindById("wnd[2]/tbar[0]/btn[0]").press
'SAP_Session.FindById("wnd[1]/usr/btnAPP_WL_SING").press
'
'
''Work on Position layout section
'
'
'
'SAP_Session.FindById("wnd[0]/tbar[1]/btn[32]").press
'SAP_Session.FindById("wnd[1]/usr/btnLINE").press
'SAP_Session.FindById("wnd[1]/usr/btnAPP_WL_ALL").press
'SAP_Session.FindById("wnd[1]/usr/btnAPP_FL_ALL").press
'
'SAP_Session.FindById("wnd[1]/usr/btnB_SEARCH").press
''SAP_Session.FindById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "document number"
'SAP_Session.FindById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "BELNR"
'SAP_Session.FindById("wnd[2]").sendVKey 0
'SAP_Session.FindById("wnd[1]/usr/btnAPP_WL_SING").press
'
'SAP_Session.FindById("wnd[1]/usr/btnB_SEARCH").press
''SAP_Session.FindById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "line item"
'SAP_Session.FindById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "BUZEI"
'SAP_Session.FindById("wnd[2]").sendVKey 0
'SAP_Session.FindById("wnd[1]/usr/btnAPP_WL_SING").press
'
'SAP_Session.FindById("wnd[1]/usr/btnB_SEARCH").press
''SAP_Session.FindById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "posting date"
'SAP_Session.FindById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "BUDAT"
'SAP_Session.FindById("wnd[2]").sendVKey 0
'SAP_Session.FindById("wnd[1]/usr/btnAPP_WL_SING").press
'
'SAP_Session.FindById("wnd[1]/usr/btnB_SEARCH").press
''SAP_Session.FindById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "currency"
'SAP_Session.FindById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "WAERS"
'SAP_Session.FindById("wnd[2]").sendVKey 0
'SAP_Session.FindById("wnd[1]/usr/btnAPP_WL_SING").press
'
'SAP_Session.FindById("wnd[1]/usr/btnB_SEARCH").press
''SAP_Session.FindById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "amount in foreign cur."
'SAP_Session.FindById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "WRSHB"
'SAP_Session.FindById("wnd[2]").sendVKey 0
'SAP_Session.FindById("wnd[1]/usr/btnAPP_WL_SING").press
'
'SAP_Session.FindById("wnd[1]/usr/btnB_SEARCH").press
''SAP_Session.FindById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "disc. amount"
'SAP_Session.FindById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "ABZUG"
'SAP_Session.FindById("wnd[2]").sendVKey 0
'SAP_Session.FindById("wnd[1]/usr/btnAPP_WL_SING").press
'
'SAP_Session.FindById("wnd[1]/usr/btnB_SEARCH").press
''SAP_Session.FindById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "net amount in foreign crcy"
'SAP_Session.FindById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "NETTO"
'SAP_Session.FindById("wnd[2]/tbar[0]/btn[0]").press
'SAP_Session.FindById("wnd[1]/usr/btnAPP_WL_SING").press
'
'SAP_Session.FindById("wnd[1]/usr/btnB_SEARCH").press
''SAP_Session.FindById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "G/L Account Number for"
'SAP_Session.FindById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "UBHKT"
'SAP_Session.FindById("wnd[2]").sendVKey 0
'SAP_Session.FindById("wnd[1]/usr/btnAPP_WL_SING").press
'
'SAP_Session.FindById("wnd[1]/usr/btnB_SEARCH").press
''SAP_Session.FindById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "Assignment"
'SAP_Session.FindById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "ZUONR"
'SAP_Session.FindById("wnd[2]").sendVKey 0
'SAP_Session.FindById("wnd[1]/usr/btnAPP_WL_SING").press
'
'SAP_Session.FindById("wnd[1]/usr/btnB_SEARCH").press
''SAP_Session.FindById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "Text"
'SAP_Session.FindById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "SGTXT"
'SAP_Session.FindById("wnd[2]/tbar[0]/btn[0]").press
'SAP_Session.FindById("wnd[1]/usr/btnAPP_WL_SING").press
'
'SAP_Session.FindById("wnd[1]/usr/btnB_SEARCH").press
''SAP_Session.FindById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "Reference"
'SAP_Session.FindById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "XBLNR"
'SAP_Session.FindById("wnd[2]/tbar[0]/btn[0]").press
'SAP_Session.FindById("wnd[1]/usr/btnAPP_WL_SING").press
'
'SAP_Session.FindById("wnd[1]/usr/btnB_SEARCH").press
''SAP_Session.FindById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "Check Number"
'SAP_Session.FindById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "CHECT"
'SAP_Session.FindById("wnd[2]").sendVKey 0
'SAP_Session.FindById("wnd[1]/usr/btnAPP_WL_SING").press
'
'SAP_Session.FindById("wnd[1]/tbar[0]/btn[0]").press
'
'
' '=====================CHANGE LAYOUT MANUALLY TO ORF RECON ENDS HERE======================================================
 
 
 

End Sub
