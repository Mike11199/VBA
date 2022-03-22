Attribute VB_Name = "Run_F110_Allowance"
Public BDPass As String
Public BDUserName As String

Sub Run_F110_Allowance()

On Error GoTo ErrSap

Application.DisplayAlerts = False
DoEvents


Dim wb As Workbook
Set wb = ThisWorkbook

BD_LOG_ON.Show
SapConnectionString = wb.Sheets("Macro Input").Range("SAP_Connection")

Dim StartTime As Double
Dim MinutesElapsed As String
'Remember time when macro starts
StartTime = Timer


'Early binding with  intellisense
Set SAP_Application = New SAPFEWSELib.GuiApplication




Rem Open a connection in synchronous mode

'Set Connection = SAP_Application.OpenConnection("EP0 - SAP ECC Production", True)    'Line with hardcoded connection name

Set Connection = SAP_Application.OpenConnection(SapConnectionString, True)
Set SAP_Session = Connection.Children(0)


'=======================================================this code is for maximizing or minimizing the window==================
SessionHWND = SAP_Session.FindById("wnd[0]").Handle

ActivateWindow (SessionHWND)
SAP_Session.FindById("wnd[0]").Maximize


'=======================================================================================================================

'If Not SingleSignOnValue = 1 Then
        
        BDUserName = BD_LOG_ON.BDUserBox.Value
        BDPass = BD_LOG_ON.BDPasswordBox.Value
        
        Sleep 1
        SAP_Session.FindById("wnd[0]/usr/txtRSYST-BNAME").Text = BDUserName
        Sleep 1
        SAP_Session.FindById("wnd[0]/usr/pwdRSYST-BCODE").Text = BDPass
        Sleep 1
        SAP_Session.FindById("wnd[0]").sendVKey 0
        Sleep 1

'End If


Unload BD_LOG_ON
BDUserName = ""
BDPass = ""


'========================================================================================================================

Sleep 1






Dim GLAccount As Long, FiscalYear As Long, ReconMonth As Variant, ReconMonth_Num As Long


FiscalYear = wb.Sheets("Macro Input").Range("Fiscal_Year")
ReconMonth = wb.Sheets("Macro Input").Range("Recon_Month")
ReconMonth_Num = wb.Sheets("Macro Input").Range("ReconMonth_Num")
Invoice_Date = wb.Sheets("Macro Input").Range("INVOICE_DATE")
F110_Run_Date = wb.Sheets("Macro Input").Range("F110_Run_Date")

Sleep 1


Template_End_Row = ThisWorkbook.Sheets("Template").Range("TEMPLATE_SUMMARY").row - 1
Data_End_Row = ThisWorkbook.Sheets("Template").Range("B" & Template_End_Row).End(xlUp).row

Dim DataRange As Variant
Dim Name As String

DataRange = ThisWorkbook.Sheets("Template").Range("A8:T" & Data_End_Row)



'Add proposal for CB0P paper allowance roll payments
                
                Sleep 1
                SAP_Session.FindById("wnd[0]/tbar[0]/okcd").Text = "/nF110 "
                Sleep 1
                SAP_Session.FindById("wnd[0]").sendVKey 0
                Sleep 1
                
                SAP_Session.FindById("wnd[0]/usr/ctxtF110V-LAUFD").Text = F110_Run_Date
                SAP_Session.FindById("wnd[0]/usr/ctxtF110V-LAUFI").Text = "CB0P"
                SAP_Session.FindById("wnd[0]/usr/ctxtF110V-LAUFI").SetFocus
                SAP_Session.FindById("wnd[0]/usr/ctxtF110V-LAUFI").CaretPosition = 4
                SAP_Session.FindById("wnd[0]").sendVKey 0
                Sleep 1
                SAP_Session.FindById("wnd[0]/usr/tabsF110_TABSTRIP/tabpPAR").Select

                SAP_Session.FindById("wnd[0]/usr/tabsF110_TABSTRIP/tabpPAR/ssubSUBSCREEN_BODY:SAPF110V:0202/tblSAPF110VCTRL_FKTTAB/txtF110V-BUKLS[0,0]").Text = "1000"
                SAP_Session.FindById("wnd[0]/usr/tabsF110_TABSTRIP/tabpPAR/ssubSUBSCREEN_BODY:SAPF110V:0202/tblSAPF110VCTRL_FKTTAB/ctxtF110V-ZWELS[1,0]").Text = "P"
                SAP_Session.FindById("wnd[0]/usr/tabsF110_TABSTRIP/tabpPAR/ssubSUBSCREEN_BODY:SAPF110V:0202/tblSAPF110VCTRL_FKTTAB/ctxtF110V-NEDAT[2,0]").Text = "12/31/2090"
                SAP_Session.FindById("wnd[0]/usr/tabsF110_TABSTRIP/tabpPAR/ssubSUBSCREEN_BODY:SAPF110V:0202/subSUBSCR_SEL:SAPF110V:7004/ctxtR_LIFNR-LOW").Text = "1"
                SAP_Session.FindById("wnd[0]/usr/tabsF110_TABSTRIP/tabpPAR/ssubSUBSCREEN_BODY:SAPF110V:0202/subSUBSCR_SEL:SAPF110V:7004/ctxtR_LIFNR-HIGH").Text = "9999"
                SAP_Session.FindById("wnd[0]/usr/tabsF110_TABSTRIP/tabpPAR/ssubSUBSCREEN_BODY:SAPF110V:0202/subSUBSCR_SEL:SAPF110V:7004/ctxtR_LIFNR-HIGH").SetFocus
                SAP_Session.FindById("wnd[0]/usr/tabsF110_TABSTRIP/tabpPAR/ssubSUBSCREEN_BODY:SAPF110V:0202/subSUBSCR_SEL:SAPF110V:7004/ctxtR_LIFNR-HIGH").CaretPosition = 4
                SAP_Session.FindById("wnd[0]/usr/tabsF110_TABSTRIP/tabpSEL").Select
                SAP_Session.FindById("wnd[0]/usr/tabsF110_TABSTRIP/tabpSEL/ssubSUBSCREEN_BODY:SAPF110V:0203/sub:SAPF110V:0203/ctxtF110V-TEXT1[0,11]").SetFocus
                SAP_Session.FindById("wnd[0]/usr/tabsF110_TABSTRIP/tabpSEL/ssubSUBSCREEN_BODY:SAPF110V:0203/sub:SAPF110V:0203/ctxtF110V-TEXT1[0,11]").CaretPosition = 0
                SAP_Session.FindById("wnd[0]").sendVKey 4
                SAP_Session.FindById("wnd[1]/usr/lbl[1,4]").SetFocus
                SAP_Session.FindById("wnd[1]/usr/lbl[1,4]").CaretPosition = 12
                SAP_Session.FindById("wnd[1]/tbar[0]/btn[0]").press
                SAP_Session.FindById("wnd[0]/usr/tabsF110_TABSTRIP/tabpSEL/ssubSUBSCREEN_BODY:SAPF110V:0203/sub:SAPF110V:0203/txtF110V-LIST1[1,11]").Text = "ZI"
                SAP_Session.FindById("wnd[0]/usr/tabsF110_TABSTRIP/tabpSEL/ssubSUBSCREEN_BODY:SAPF110V:0203/sub:SAPF110V:0203/txtF110V-LIST1[1,11]").SetFocus
                SAP_Session.FindById("wnd[0]/usr/tabsF110_TABSTRIP/tabpSEL/ssubSUBSCREEN_BODY:SAPF110V:0203/sub:SAPF110V:0203/txtF110V-LIST1[1,11]").CaretPosition = 2
                SAP_Session.FindById("wnd[0]/usr/tabsF110_TABSTRIP/tabpLOG").Select
                SAP_Session.FindById("wnd[0]/usr/tabsF110_TABSTRIP/tabpLOG/ssubSUBSCREEN_BODY:SAPF110V:0204/chkF110V-XTRFA").Selected = True
                SAP_Session.FindById("wnd[0]/usr/tabsF110_TABSTRIP/tabpLOG/ssubSUBSCREEN_BODY:SAPF110V:0204/chkF110V-XTRZE").Selected = True
                SAP_Session.FindById("wnd[0]/usr/tabsF110_TABSTRIP/tabpLOG/ssubSUBSCREEN_BODY:SAPF110V:0204/chkF110V-XTRBL").Selected = True
                
                Dim star As String
                star = CStr("*")
                star = Clean_NonPrintableCharacters(star)
                
                SAP_Session.FindById("wnd[0]/usr/tabsF110_TABSTRIP/tabpLOG/ssubSUBSCREEN_BODY:SAPF110V:0204/sub:SAPF110V:0204/txtF110V-VONKK[0,0]").Text = star
                SAP_Session.FindById("wnd[0]/usr/tabsF110_TABSTRIP/tabpLOG/ssubSUBSCREEN_BODY:SAPF110V:0204/sub:SAPF110V:0204/txtF110V-VONKK[0,0]").CaretPosition = 1
                
                SAP_Session.FindById("wnd[0]/usr/tabsF110_TABSTRIP/tabpPRI").Select
                SAP_Session.FindById("wnd[0]/usr/tabsF110_TABSTRIP/tabpPRI/ssubSUBSCREEN_BODY:SAPF110V:0205/tblSAPF110VCTRL_DRLTAB/txtF110V-LPROG[0,0]").Text = "ZCS_CLAIM_SCHED_FILL"
                SAP_Session.FindById("wnd[0]/usr/tabsF110_TABSTRIP/tabpPRI/ssubSUBSCREEN_BODY:SAPF110V:0205/tblSAPF110VCTRL_DRLTAB/txtF110V-LPROG[0,0]").SetFocus
                SAP_Session.FindById("wnd[0]/usr/tabsF110_TABSTRIP/tabpPRI/ssubSUBSCREEN_BODY:SAPF110V:0205/tblSAPF110VCTRL_DRLTAB/txtF110V-LPROG[0,0]").CaretPosition = 20
                SAP_Session.FindById("wnd[0]/tbar[0]/btn[11]").press
                
                SAP_Session.FindById("wnd[0]/usr/tabsF110_TABSTRIP/tabpSTA").Select
                
                
                
                                'get rid of save data exit editing pop up that sometimes appears when going back to status tab
                'from the parameters tab.  Only does code if pop up wnd 1 appears
                If SAP_Session.ActiveWindow.Name = "wnd[1]" Then
                
                If SAP_Session.FindById("wnd[1]").Text Like "Exit editing*" Then
                       SAP_Session.FindById("wnd[1]/usr/btnSPOP-OPTION1").press
                End If
                
                End If
                '====================================
                
                SAP_Session.FindById("wnd[0]/tbar[1]/btn[13]").press
                
                SAP_Session.FindById("wnd[1]/usr/chkF110V-XSTRF").Selected = True
                SAP_Session.FindById("wnd[1]/usr/chkF110V-XMITL").Selected = True
                SAP_Session.FindById("wnd[1]/usr/chkF110V-XMITL").SetFocus
                SAP_Session.FindById("wnd[1]/tbar[0]/btn[0]").press
                
                
                
                Sleep 100

                '===============Loop to press status to refresh until proposal has been generated====================

                Status_Bar = SAP_Session.FindById("wnd[0]/sbar").Text
                Status_Bar_Done = Right(Status_Bar, 14)
                
                Do While Status_Bar_Done = "been scheduled" Or Status_Bar_Done = ""
                        SAP_Session.FindById("wnd[0]/tbar[1]/btn[14]").press
                        Status_Bar = SAP_Session.FindById("wnd[0]/sbar").Text
                        Status_Bar_Done = Right(Status_Bar, 14)
                        Sleep 100
                Loop
                '===================looop ends here=============================================================

                'press proposal button
                SAP_Session.FindById("wnd[0]/tbar[1]/btn[21]").press
                
                
                
                
                
                
                ActivateWindow (SessionHWND)
                Sleep 2000
                                
                SAP_Session.FindById("wnd[0]/usr/txtF110O-AZAHL").SetFocus
                OutgoingPayment = SAP_Session.FindById("wnd[0]/usr/txtF110O-AZAHL").Text
                
                Call PrintScreen
                
                Sleep 2000
                

        
                
                
ActivateWindow (Application.hwnd)    'Brings the Excel window to the front


                wb.Sheets.Add(After:=Sheets("Macro Input")).Name = "F110_Screenshots_" & wb.Sheets.Count
                
                Dim screenshot_sheet As Worksheet
                Set screenshot_sheet = wb.Sheets("F110_Screenshots_" & wb.Sheets.Count - 1)
                
                screenshot_sheet.Activate
                
                With ActiveWorkbook.ActiveSheet.Tab
                .Color = 192
                .TintAndShade = 0
                End With
                
               
                screenshot_sheet.Range("A1").Select
                Sleep 4000
         '       screenshot_sheet.Range("A1").PasteSpecial
                DoEvents
                ActiveSheet.Paste Destination:=screenshot_sheet.Range("A1")
                Sleep 6000
                Call PictureFormat
                Sleep 2000
                
                screenshot_sheet.Range("N1") = OutgoingPayment
                
                Dim mynamedrange_1 As Range
                Dim mynamedrange_2 As Range
                Dim myRangeName As String
                
                Set mynamedrange_1 = screenshot_sheet.Range("N1")
                myRangeName = "CB_NP_F110_CB0P"
                
                ThisWorkbook.Names.Add Name:=myRangeName, RefersTo:=mynamedrange_1


                ActivateWindow (SessionHWND)
                SAP_Session.FindById("wnd[0]").Maximize

'Add proposal for CB0P EFT allowance roll payments
                
                Sleep 1
                SAP_Session.FindById("wnd[0]/tbar[0]/okcd").Text = "/nF110 "
                Sleep 1
                SAP_Session.FindById("wnd[0]").sendVKey 0
                Sleep 1
                
                SAP_Session.FindById("wnd[0]/usr/ctxtF110V-LAUFD").Text = F110_Run_Date
                SAP_Session.FindById("wnd[0]/usr/ctxtF110V-LAUFI").Text = "CB0T"
                SAP_Session.FindById("wnd[0]/usr/ctxtF110V-LAUFI").SetFocus
                SAP_Session.FindById("wnd[0]/usr/ctxtF110V-LAUFI").CaretPosition = 4
                SAP_Session.FindById("wnd[0]").sendVKey 0
                Sleep 1
                SAP_Session.FindById("wnd[0]/usr/tabsF110_TABSTRIP/tabpPAR").Select

                SAP_Session.FindById("wnd[0]/usr/tabsF110_TABSTRIP/tabpPAR/ssubSUBSCREEN_BODY:SAPF110V:0202/tblSAPF110VCTRL_FKTTAB/txtF110V-BUKLS[0,0]").Text = "1000"
                SAP_Session.FindById("wnd[0]/usr/tabsF110_TABSTRIP/tabpPAR/ssubSUBSCREEN_BODY:SAPF110V:0202/tblSAPF110VCTRL_FKTTAB/ctxtF110V-ZWELS[1,0]").Text = "T"
                SAP_Session.FindById("wnd[0]/usr/tabsF110_TABSTRIP/tabpPAR/ssubSUBSCREEN_BODY:SAPF110V:0202/tblSAPF110VCTRL_FKTTAB/ctxtF110V-NEDAT[2,0]").Text = "12/31/2090"
                SAP_Session.FindById("wnd[0]/usr/tabsF110_TABSTRIP/tabpPAR/ssubSUBSCREEN_BODY:SAPF110V:0202/subSUBSCR_SEL:SAPF110V:7004/ctxtR_LIFNR-LOW").Text = "1"
                SAP_Session.FindById("wnd[0]/usr/tabsF110_TABSTRIP/tabpPAR/ssubSUBSCREEN_BODY:SAPF110V:0202/subSUBSCR_SEL:SAPF110V:7004/ctxtR_LIFNR-HIGH").Text = "9999"
                SAP_Session.FindById("wnd[0]/usr/tabsF110_TABSTRIP/tabpPAR/ssubSUBSCREEN_BODY:SAPF110V:0202/subSUBSCR_SEL:SAPF110V:7004/ctxtR_LIFNR-HIGH").SetFocus
                SAP_Session.FindById("wnd[0]/usr/tabsF110_TABSTRIP/tabpPAR/ssubSUBSCREEN_BODY:SAPF110V:0202/subSUBSCR_SEL:SAPF110V:7004/ctxtR_LIFNR-HIGH").CaretPosition = 4
                SAP_Session.FindById("wnd[0]/usr/tabsF110_TABSTRIP/tabpSEL").Select
                SAP_Session.FindById("wnd[0]/usr/tabsF110_TABSTRIP/tabpSEL/ssubSUBSCREEN_BODY:SAPF110V:0203/sub:SAPF110V:0203/ctxtF110V-TEXT1[0,11]").SetFocus
                SAP_Session.FindById("wnd[0]/usr/tabsF110_TABSTRIP/tabpSEL/ssubSUBSCREEN_BODY:SAPF110V:0203/sub:SAPF110V:0203/ctxtF110V-TEXT1[0,11]").CaretPosition = 0
                SAP_Session.FindById("wnd[0]").sendVKey 4
                SAP_Session.FindById("wnd[1]/usr/lbl[1,4]").SetFocus
                SAP_Session.FindById("wnd[1]/usr/lbl[1,4]").CaretPosition = 12
                SAP_Session.FindById("wnd[1]/tbar[0]/btn[0]").press
                SAP_Session.FindById("wnd[0]/usr/tabsF110_TABSTRIP/tabpSEL/ssubSUBSCREEN_BODY:SAPF110V:0203/sub:SAPF110V:0203/txtF110V-LIST1[1,11]").Text = "ZI"
                SAP_Session.FindById("wnd[0]/usr/tabsF110_TABSTRIP/tabpSEL/ssubSUBSCREEN_BODY:SAPF110V:0203/sub:SAPF110V:0203/txtF110V-LIST1[1,11]").SetFocus
                SAP_Session.FindById("wnd[0]/usr/tabsF110_TABSTRIP/tabpSEL/ssubSUBSCREEN_BODY:SAPF110V:0203/sub:SAPF110V:0203/txtF110V-LIST1[1,11]").CaretPosition = 2
                SAP_Session.FindById("wnd[0]/usr/tabsF110_TABSTRIP/tabpLOG").Select
                SAP_Session.FindById("wnd[0]/usr/tabsF110_TABSTRIP/tabpLOG/ssubSUBSCREEN_BODY:SAPF110V:0204/chkF110V-XTRFA").Selected = True
                SAP_Session.FindById("wnd[0]/usr/tabsF110_TABSTRIP/tabpLOG/ssubSUBSCREEN_BODY:SAPF110V:0204/chkF110V-XTRZE").Selected = True
                SAP_Session.FindById("wnd[0]/usr/tabsF110_TABSTRIP/tabpLOG/ssubSUBSCREEN_BODY:SAPF110V:0204/chkF110V-XTRBL").Selected = True
                
                
                SAP_Session.FindById("wnd[0]/usr/tabsF110_TABSTRIP/tabpLOG/ssubSUBSCREEN_BODY:SAPF110V:0204/sub:SAPF110V:0204/txtF110V-VONKK[0,0]").Text = star
                SAP_Session.FindById("wnd[0]/usr/tabsF110_TABSTRIP/tabpLOG/ssubSUBSCREEN_BODY:SAPF110V:0204/sub:SAPF110V:0204/txtF110V-VONKK[0,0]").CaretPosition = 1
                
                SAP_Session.FindById("wnd[0]/usr/tabsF110_TABSTRIP/tabpPRI").Select
                SAP_Session.FindById("wnd[0]/usr/tabsF110_TABSTRIP/tabpPRI/ssubSUBSCREEN_BODY:SAPF110V:0205/tblSAPF110VCTRL_DRLTAB/txtF110V-LPROG[0,0]").Text = "ZCS_CLAIM_SCHED_FILL"
                SAP_Session.FindById("wnd[0]/usr/tabsF110_TABSTRIP/tabpPRI/ssubSUBSCREEN_BODY:SAPF110V:0205/tblSAPF110VCTRL_DRLTAB/txtF110V-LPROG[0,0]").SetFocus
                SAP_Session.FindById("wnd[0]/usr/tabsF110_TABSTRIP/tabpPRI/ssubSUBSCREEN_BODY:SAPF110V:0205/tblSAPF110VCTRL_DRLTAB/txtF110V-LPROG[0,0]").CaretPosition = 20
                SAP_Session.FindById("wnd[0]/tbar[0]/btn[11]").press
              
                SAP_Session.FindById("wnd[0]/usr/tabsF110_TABSTRIP/tabpPAR").Select
                SAP_Session.FindById("wnd[0]/usr/tabsF110_TABSTRIP/tabpPAR/ssubSUBSCREEN_BODY:SAPF110V:0202/subSUBSCR_SEL:SAPF110V:7004/btn%_R_LIFNR_%_APP_%-VALU_PUSH").press
                
                ActivateWindow (Application.hwnd)    'Brings the Excel window to the front

                
                IRS_WH = ThisWorkbook.Sheets("Template").Range("TOTAL_FED_WH")
                FTB_WH = ThisWorkbook.Sheets("Template").Range("TOTAL_STATE_WH")
                
                
                 'IRS
                answer = MsgBox("Add IRS Vendor GL for EFT Payment Proposal?" & vbNewLine & vbNewLine & "Fed WH Amount:  " & IRS_WH, vbQuestion + vbYesNo + vbDefaultButton2, "Add IRS Vendor?")
                If answer = vbYes Then
                     SAP_Session.FindById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").Text = "3100001085"
                End If
                
                Sleep 1
                
                ActivateWindow (Application.hwnd)    'Brings the Excel window to the front
                
                'FTB
                answer2 = MsgBox("Add FTB Vendor GL for EFT Payment Proposal?" & vbNewLine & vbNewLine & "State WH Amount:  " & FTB_WH, vbQuestion + vbYesNo + vbDefaultButton2, "Add FTB Vendor?")
                If answer2 = vbYes Then
                      SAP_Session.FindById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").Text = "3100000999"
                End If
                
                Sleep 1000
                
                ActivateWindow (SessionHWND)

                SAP_Session.FindById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").SetFocus
                SAP_Session.FindById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").CaretPosition = 10
                SAP_Session.FindById("wnd[1]/tbar[0]/btn[8]").press
 
                SAP_Session.FindById("wnd[0]/tbar[0]/btn[11]").press
                
      
                SAP_Session.FindById("wnd[0]/usr/tabsF110_TABSTRIP/tabpSTA").Select
                SAP_Session.FindById("wnd[1]/usr/btnSPOP-OPTION1").press
                
                
                SAP_Session.FindById("wnd[0]/tbar[1]/btn[13]").press
                SAP_Session.FindById("wnd[1]/usr/chkF110V-XSTRF").Selected = True
                SAP_Session.FindById("wnd[1]/usr/chkF110V-XMITL").Selected = True
                SAP_Session.FindById("wnd[1]/usr/chkF110V-XMITL").SetFocus
                SAP_Session.FindById("wnd[1]/tbar[0]/btn[0]").press
                
                ActivateWindow (SessionHWND)
                
              Sleep 100

                '===============Loop to press status to refresh until proposal has been generated====================

                Status_Bar = SAP_Session.FindById("wnd[0]/sbar").Text
                Status_Bar_Done = Right(Status_Bar, 14)
                
                Do While Status_Bar_Done = "been scheduled" Or Status_Bar_Done = ""
                        SAP_Session.FindById("wnd[0]/tbar[1]/btn[14]").press
                        Status_Bar = SAP_Session.FindById("wnd[0]/sbar").Text
                        Status_Bar_Done = Right(Status_Bar, 14)
                           Sleep 100
                Loop
                '===================looop ends here=============================================================

                'press proposal button
                SAP_Session.FindById("wnd[0]/tbar[1]/btn[21]").press
                
                ActivateWindow (SessionHWND)
                
                
                Sleep 2000
                
                SAP_Session.FindById("wnd[0]/usr/txtF110O-AZAHL").SetFocus
                OutgoingPayment = SAP_Session.FindById("wnd[0]/usr/txtF110O-AZAHL").Text
                
                Call PrintScreen
                
                Sleep 2000
                  
                
                
                screenshot_sheet.Activate
                DoEvents
                screenshot_sheet.Range("A35").Select
                Sleep 5000
                DoEvents
                'screenshot_sheet.Range("A50").PasteSpecial
                ActiveSheet.Paste Destination:=screenshot_sheet.Range("A35")
                
                screenshot_sheet.Range("N35") = OutgoingPayment
                
                                
                Set mynamedrange_2 = screenshot_sheet.Range("N35")
                myRangeName = "CB_NP_F110_CB0T"
                
                ThisWorkbook.Names.Add Name:=myRangeName, RefersTo:=mynamedrange_2
                
                Sleep 6000
                Call PictureFormat
                Sleep 2000


 
 
 
  Set SAP_Application = Nothing
 
   Range("N38").FormulaR1C1 = "=SUM(R[-37]C:R[-3]C)"
    Range("N38").Select
    Selection.Font.Bold = True
    
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlDouble
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThick
    End With
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Columns("N:N").Select
    Range("N10").Activate
    Selection.Style = "Comma"

    Selection.NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
 
 Call Add_Check_Formulas_to_WOs
 
MinutesElapsed = Format((Timer - StartTime) / 86400, "hh:mm:ss")
 
 
 MsgBox "Ran successfully in " & MinutesElapsed & " minutes." & vbNewLine & vbNewLine & _
"The macro has finished adding the exported files to you current workbook." & vbNewLine & vbNewLine & _
"Please press OK, and then close the two alert messages that open after the macro ends."



Exit Sub
ErrSap:
MsgBox "Error.  Please press OK to end the macro." & vbNewLine & vbNewLine & _
"Please check to see if this payment run or proposal was already created."





End Sub

Sub AddChecktoBal()


Dim wb As Workbook
Set wb = ThisWorkbook


FiscalYear = wb.Sheets("Macro Input").Range("Fiscal_Year")
ReconMonth = wb.Sheets("Macro Input").Range("Recon_Month")
ReconMonth_Num = wb.Sheets("Macro Input").Range("ReconMonth_Num")



Dim rng1 As Range
Dim strSearch As String

If ReconMonth_Num < 10 Then
strSearch = "00" & ReconMonth_Num
Else
strSearch = "0" & ReconMonth_Num
End If

Set rng1 = Range("A:A").Find(strSearch, , xlValues, xlWhole)
 
 rng1.Select
 rng1.Resize(, 5).Select
 
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 49407
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
    
  rng1.Offset(, 5).Formula = "=SUM('" & ReconMonth & "_GL 1130 Detail'!F:F)"
  
  rng1.Offset(, 5).Select
  
  With Selection.Font
        .Color = -16776961
        .TintAndShade = 0
    End With
    
    Selection.Style = "Comma"
    
   rng1.Offset(, 6).Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=RC[-1]-RC[-3]"
    Range("A1").Select





End Sub

Sub PrintScreen()
    keybd_event VK_SNAPSHOT, 1, 0, 0
End Sub

Sub PictureFormat()

With ActiveSheet
Set shp = .Shapes(.Shapes.Count)
End With

Crop_Right = ThisWorkbook.Sheets("Macro Input").Range("Crop_Right").Value
Crop_Bottom = ThisWorkbook.Sheets("Macro Input").Range("Crop_Bottom").Value
Scale_Height = ThisWorkbook.Sheets("Macro Input").Range("Scale_Height").Value
Scale_Width = ThisWorkbook.Sheets("Macro Input").Range("Scale_Width").Value



With shp
h = -(635 - shp.Height)
w = -(1225 - shp.Width)
l = -(Crop_Bottom - shp.Height)
r = -(Crop_Right - shp.Width)
' the new size ratio of our WHOLE screenshot pasted (with keeping aspect ratio)
'.Height = 1260
'.Width = 1680
.LockAspectRatio = False
End With

With shp.PictureFormat
.CropRight = r
'.CropLeft = w
'.CropTop = h
.CropBottom = l
End With

With shp.Line 'optional image borders
.Weight = 1
.DashStyle = msoLineSolid
End With

shp.ScaleWidth Scale_Width, msoTrue, msoScaleFromTopLeft
shp.ScaleHeight Scale_Height, msoTrue, msoScaleFromTopLeft


End Sub

Public Function Clean_NonPrintableCharacters(Str As String) As String

    'Removes non-printable characters from a string

    Dim cleanString As String
    Dim i As Integer

    cleanString = Str

    For i = Len(cleanString) To 1 Step -1
        'Debug.Print Asc(Mid(Str, i, 1))

        Select Case Asc(Mid(Str, i, 1))
            Case 1 To 31, Is >= 127
                'Bad stuff
                'https://www.ionos.com/digitalguide/server/know-how/ascii-codes-overview-of-all-characters-on-the-ascii-table/
                cleanString = Left(cleanString, i - 1) & Mid(cleanString, i + 1)

            Case Else
                'Keep

        End Select
    Next i

    Clean_NonPrintableCharacters = cleanString

End Function


Sub Add_Check_Formulas_to_WOs()

' Macro76 Macro
'

'
    Range("N40").Select
    ActiveCell.FormulaR1C1 = "=TEMPLATE_SUMMARY_GROSS_DB"
    Range("N41").Select
    ActiveCell.FormulaR1C1 = "=TEMPLATE_SUMMARY_GROSS_SB"
    Range("N42").Select
    ActiveCell.FormulaR1C1 = "=TEMPLATE_SUMMARY_GROSS_SR"
        Range("N43").Select
    ActiveCell.FormulaR1C1 = "=RECEIVABLE_TOTAL_SUMMARY"
    Range("O40").Select
    ActiveCell.FormulaR1C1 = "DB"
    Range("O41").Select
    ActiveCell.FormulaR1C1 = "SB"
    Range("O42").Select
    ActiveCell.FormulaR1C1 = "SR"
    Range("O43").Select
    ActiveCell.FormulaR1C1 = "Receivable"
    Range("O40:O47").Select
    With Selection.Font
        .Color = -16776961
        .TintAndShade = 0
    End With
    Selection.Font.Bold = True
    
    
    Range("N44").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlDouble
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThick
    End With
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Range("N44").Select
    Selection.Font.Bold = True

    Range("N44").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=SUM(R[-3]C:R[-1]C)"

    Range("O44").Select
    ActiveCell.FormulaR1C1 = "Sum from Template"
    Range("N44").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=R[-6]C-R[-1]C"
    Range("N45").Select
    Columns("N:N").EntireColumn.AutoFit
        Range("O44").Select
    With Selection.Font
        .Color = -16776961
        .TintAndShade = 0
    End With
    With Selection.Font
        .Color = -16776961
        .TintAndShade = 0
    End With
        Selection.Font.Bold = True
    

    Range("N44").Select
    ActiveCell.FormulaR1C1 = "=SUM(R[-4]C:R[-2]C)-R[-1]C"
    Range("N46").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=R[-8]C-R[-2]C"
    Range("N47").Select
    
        Range("O46").Select
    ActiveCell.FormulaR1C1 = "Check Template to F110 Proposals"
    
    End Sub

