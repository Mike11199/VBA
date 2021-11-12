Attribute VB_Name = "BD_Pull_GL_ACTIVITY_Range"
Public BDPass As String
Public BDUserName As String
Public SAP_Application As Object
Public SAP_Session As Object
Public Recon_WB As Object
Public GLBal As Worksheet
Public GL_Export_WB As Workbook
Public GL_Balance_Array As Variant
Public GL_Activity_Array As Variant
Public FiscalYear As Long
Public ReconMonth As Variant
Public ReconMonth_Num As Long
Public shp As Shape
Public h As Single
Public w As Single
Public l As Single
Public R As Single
Public GLCount As Integer
Public LastRow As Variant
Public LastCol As Variant
Public Number_GL_to_Run As Integer
Public screenshotrows As Integer
Public First_GL_Exported As Integer
Public LoopNumber As Integer

Sub Pull_GL_Range()

On Error GoTo ErrSap

Dim Recon_WB As Workbook
Set Recon_WB = ThisWorkbook

BD_LOG_ON.Show

Dim StartTime As Double
Dim MinutesElapsed As String
'Remember time when macro starts
StartTime = Timer


SapConnectionString = Recon_WB.Sheets("Macro Input").Range("SAP_Connection")


Set SAP_Application = CreateObject("Sapgui.ScriptingCtrl.1")

Set Connection = SAP_Application.OpenConnection(SapConnectionString, True)
Set SAP_Session = Connection.Children(0)


SessionHWND = SAP_Session.FindById("wnd[0]").Handle
ActivateWindow (SessionHWND)



SAP_Session.FindById("wnd[0]").Maximize


        
        BDUserName9 = BD_LOG_ON.BDUserBox.Value
        BDPass9 = BD_LOG_ON.BDPasswordBox.Value
        
        Application.Wait (Now + TimeValue("00:00:01") / 1000)
        SAP_Session.FindById("wnd[0]/usr/txtRSYST-BNAME").Text = BDUserName9
        Application.Wait (Now + TimeValue("00:00:01") / 1000)
        SAP_Session.FindById("wnd[0]/usr/pwdRSYST-BCODE").Text = BDPass9
        Application.Wait (Now + TimeValue("00:00:01") / 1000)
        SAP_Session.FindById("wnd[0]").sendVKey 0
        Application.Wait (Now + TimeValue("00:00:01") / 1000)


Unload BD_LOG_ON
BDUserName9 = ""
BDPass9 = ""

'========================================================================================================================

Application.Wait (Now + TimeValue("00:00:01") / 1000)

Application.Wait (Now + TimeValue("00:00:01") / 1000)
SAP_Session.FindById("wnd[0]/tbar[0]/okcd").Text = "/nfagll03 "
Application.Wait (Now + TimeValue("00:00:01") / 1000)
SAP_Session.FindById("wnd[0]").sendVKey 0
Application.Wait (Now + TimeValue("00:00:01") / 1000)


Dim GLAccount As Long, FiscalYear As Long, ReconMonth As Variant, ReconMonth_Num As Long


FiscalYear = Recon_WB.Sheets("Macro Input").Range("Fiscal_Year")
ReconMonth = Recon_WB.Sheets("Macro Input").Range("Recon_Month")
ReconMonth_Num = Recon_WB.Sheets("Macro Input").Range("ReconMonth_Num")


GL_Range_1 = Recon_WB.Sheets("Macro Input").Range("GL_Range_1")
GL_Range_2 = Recon_WB.Sheets("Macro Input").Range("GL_Range_2")




SAP_Session.FindById("wnd[0]/usr/radX_AISEL").Select
SAP_Session.FindById("wnd[0]/usr/ctxtSD_SAKNR-LOW").Text = GL_Range_1
SAP_Session.FindById("wnd[0]/usr/ctxtSD_SAKNR-HIGH").Text = GL_Range_2
SAP_Session.FindById("wnd[0]/usr/radX_AISEL").SetFocus
SAP_Session.FindById("wnd[0]").sendVKey 8



 '=====================CHANGE LAYOUT  TO ORF RECON WITH SAVED LAYOUT======================================================
SAP_Session.FindById("wnd[0]/tbar[1]/btn[33]").press
SAP_Session.FindById("wnd[1]/tbar[0]/btn[71]").press
SAP_Session.FindById("wnd[2]/usr/chkSCAN_STRING-RANGE").Selected = True
SAP_Session.FindById("wnd[2]/usr/chkSCAN_STRING-START").Selected = False
SAP_Session.FindById("wnd[2]/usr/txtRSYSF-STRING").Text = "/ORF_MACRO"

SAP_Session.FindById("wnd[2]/tbar[0]/btn[0]").press
SAP_Session.FindById("wnd[3]/usr/lbl[1,2]").SetFocus

SAP_Session.FindById("wnd[3]").sendVKey 2
SAP_Session.FindById("wnd[1]/tbar[0]/btn[0]").press
 
'=============================================================================================================================



 '======================EXPORT TO EXCEL WORKBOOK=======================================
 
 
 SAP_Session.FindById("wnd[0]/mbar/menu[0]/menu[3]/menu[1]").Select
SAP_Session.FindById("wnd[1]/usr/ctxtDY_PATH").Text = "C:\\TEMP"
SAP_Session.FindById("wnd[1]/usr/ctxtDY_FILENAME").Text = "EXPORT2.MHTML"
SAP_Session.FindById("wnd[1]/tbar[0]/btn[11]").press
Application.Wait (Now + TimeValue("00:00:10"))
  
  
  
 Set GL_Export_WB2 = Workbooks.Open("C:\TEMP\EXPORT2.MHTML")

GL_Export_WB2.Sheets(1).Copy After:=Recon_WB.Sheets("Macro Input")

Recon_WB.ActiveSheet.Name = ReconMonth & "_All GL 1190 Detail"


Application.Wait (Now + TimeValue("00:00:01"))
GL_Export_WB2.Close SaveChanges:=False
Application.Wait (Now + TimeValue("00:00:01"))
 
 

ActivateWindow (Application.hwnd) 'ExcelWBInFront
 
 
Dim GL2 As Worksheet
Set GL2 = Recon_WB.ActiveSheet

With ActiveWorkbook.ActiveSheet.Tab
        .Color = 192
        .TintAndShade = 0
End With

Recon_WB.Sheets("Macro Input").Activate
Recon_WB.Sheets("Macro Input").Range("A1").Select
 

Set SAP_Application = Nothing

Continuewiththis = MsgBox(Prompt:="Delete exported .MHTML file in C:/TEMP? (selecting no can cause Excel to crash)", Buttons:=vbQuestion + vbYesNo)

If Continuewiththis = vbYes Then
    Kill ("C:\TEMP\EXPORT2.MHTML")
End If

  
MinutesElapsed = Format((Timer - StartTime) / 86400, "hh:mm:ss")
 
 
 MsgBox "Ran successfully in " & MinutesElapsed & " minutes." & vbNewLine & vbNewLine & _
"The macro has finished adding the exported files to you current workbook." & vbNewLine & vbNewLine & _
"Please press OK, and then close the file that opens after the macro ends."


Exit Sub
ErrSap:
MsgBox "Error.  Please press OK to end the macro."




End Sub

