Attribute VB_Name = "ClearingDocsMacro"
Public BDPass As String
Public BDUserName As String

Sub Clear_Docs()

On Error GoTo ErrSap

Dim ORF_WB As Workbook
Set ORF_WB = ThisWorkbook

BD_LOG_ON.Show
SapConnectionString = ORF_WB.Sheets("Macro Input").Range("SAP_Connection")
GL_ACCOUNT = ORF_WB.Sheets("Macro Input").Range("B10")


Set SAP_Application = New SAPFEWSELib.GuiApplication




Set Connection = SAP_Application.OpenConnection(SapConnectionString, True)
Set SAP_Session = Connection.Children(0)


'=======================================================this code is for maximizing or minimizing the window==================
SessionHWND = SAP_Session.FindById("wnd[0]").Handle

ActivateWindow (SessionHWND)
SAP_Session.FindById("wnd[0]").Maximize



'=======================================================================================================================

If Not SingleSignOnValue = 1 Then



        
        BDUserName = BD_LOG_ON.BDUserBox.Value
        BDPass = BD_LOG_ON.BDPasswordBox.Value
        
        Application.Wait (Now + TimeValue("00:00:01") / 1000)
        SAP_Session.FindById("wnd[0]/usr/txtRSYST-BNAME").Text = BDUserName
        Application.Wait (Now + TimeValue("00:00:01") / 1000)
        SAP_Session.FindById("wnd[0]/usr/pwdRSYST-BCODE").Text = BDPass
        Application.Wait (Now + TimeValue("00:00:01") / 1000)
        SAP_Session.FindById("wnd[0]").sendVKey 0
        Application.Wait (Now + TimeValue("00:00:01") / 1000)

End If


Unload BD_LOG_ON
BDUserName = ""
BDPass = ""

'========================================================================================================================

Application.Wait (Now + TimeValue("00:00:01") / 1000)

Application.Wait (Now + TimeValue("00:00:01") / 1000)
SAP_Session.FindById("wnd[0]/tbar[0]/okcd").Text = "/nf-03"
Application.Wait (Now + TimeValue("00:00:01") / 1000)
SAP_Session.FindById("wnd[0]").sendVKey 0
Application.Wait (Now + TimeValue("00:00:01") / 1000)


ThisWorkbook.Sheets("Transposed Document List").Activate
lastcolumn = Range("A4").SpecialCells(xlCellTypeLastCell).Column

X = 1

Do While X < lastcolumn + 1


Application.Wait (Now + TimeValue("00:00:03"))
'press process open items after entering GL #
SAP_Session.FindById("wnd[0]/usr/sub:SAPMF05A:0131/radRF05A-XPOS1[2,0]").Select
'SAP_Session.FindById("wnd[0]/usr/ctxtRF05A-AGKON").Text = "3010032000"
SAP_Session.FindById("wnd[0]/usr/ctxtRF05A-AGKON").Text = GL_ACCOUNT
SAP_Session.FindById("wnd[0]/usr/sub:SAPMF05A:0131/radRF05A-XPOS1[2,0]").SetFocus
SAP_Session.FindById("wnd[0]").sendVKey 16


Application.Wait (Now + TimeValue("00:00:01"))
'if blocked by other user as SAP slow then press enter to get past it
errMessage = SAP_Session.FindById("wnd[0]/sbar").Text

errMessagetest = InStr(0, CStr(errMessage), "blocked", vbBinaryCompare)
If errMessagetest <> 0 Then
    SAP_Session.FindById("wnd[0]").sendVKey 0
End If





'enter in doc #s

lastrowofcolumn = Range(Cells(80, X), Cells(80, X)).End(xlUp).Row

Dim myarray As Variant
myarray = Range(Cells(4, X), Cells(lastrowofcolumn, X))
Range(Cells(4, X), Cells(lastrowofcolumn, X)).Select
'
'Cells(4, X).Select
'Cells(X, 4).Select
'
'Cells(4, X).Select
'Cells(X, 4).Select
'
'Cells(4, 7).Select
'Cells(4, 7).Select
'Cells(4, X).Select

upperarray = UBound(myarray, 1)
lowerarray = LBound(myarray, 1)


Z = lowerarray - 1
Do While Z < upperarray

'test = myarray(1, (Z + 1))


SAP_Session.FindById("wnd[0]/usr/sub:SAPMF05A:0731/txtRF05A-SEL01[" & Z & ",0]").Text = myarray((Z + 1), 1)
Z = Z + 1

Loop



        
SAP_Session.FindById("wnd[0]/usr/sub:SAPMF05A:0731/txtRF05A-SEL01[2,0]").SetFocus
SAP_Session.FindById("wnd[0]/usr/sub:SAPMF05A:0731/txtRF05A-SEL01[2,0]").caretPosition = 9
SAP_Session.FindById("wnd[0]").sendVKey 16


errMessage = SAP_Session.FindById("wnd[0]/sbar").Text

If errMessage = "No appropriate line item is contained in this document" Then
    SAP_Session.FindById("wnd[0]").sendVKey 0
End If

errMessage = SAP_Session.FindById("wnd[0]/sbar").Text

If errMessage = "No appropriate line item is contained in this document" Then
    SAP_Session.FindById("wnd[0]").sendVKey 0
End If

errMessage = SAP_Session.FindById("wnd[0]/sbar").Text

If errMessage = "No appropriate line item is contained in this document" Then
    SAP_Session.FindById("wnd[0]").sendVKey 0
End If

errMessage = SAP_Session.FindById("wnd[0]/sbar").Text

If errMessage = "No appropriate line item is contained in this document" Then
    SAP_Session.FindById("wnd[0]").sendVKey 0
End If

errMessage = SAP_Session.FindById("wnd[0]/sbar").Text

If errMessage = "No appropriate line item is contained in this document" Then
    SAP_Session.FindById("wnd[0]").sendVKey 0
End If

errMessage = SAP_Session.FindById("wnd[0]/sbar").Text

If errMessage = "No appropriate line item is contained in this document" Then
    SAP_Session.FindById("wnd[0]").sendVKey 0
End If


errMessage = SAP_Session.FindById("wnd[0]/sbar").Text

If errMessage = "No appropriate line item is contained in this document" Then
    SAP_Session.FindById("wnd[0]").sendVKey 0
End If

errMessage = SAP_Session.FindById("wnd[0]/sbar").Text

If errMessage = "No appropriate line item is contained in this document" Then
    SAP_Session.FindById("wnd[0]").sendVKey 0
End If




errMessage = SAP_Session.FindById("wnd[0]/sbar").Text
If errMessage <> "No open items were found" Then


      
                            amount_entered = SAP_Session.FindById("wnd[0]/usr/tabsTS/tabpMAIN/ssubPAGE:SAPDF05X:6103/txtRF05A-BETRG").Text
                            amount_not_assigned = SAP_Session.FindById("wnd[0]/usr/tabsTS/tabpMAIN/ssubPAGE:SAPDF05X:6103/txtRF05A-DIFFB").Text
                            amount_assigned = SAP_Session.FindById("wnd[0]/usr/tabsTS/tabpMAIN/ssubPAGE:SAPDF05X:6103/txtRF05A-AKTIV").Text
                           

                            
                        
                        'If amount_entered = "0.00" And (amount_not_assigned = "0.00" Or Not IsEmpty(amount_notassigned)) And (amount_assigned = "0.00" Or Not IsEmpty(amount_assigned)) Then
                        If amount_entered = "0.00 " And amount_not_assigned = "0.00 " And amount_assigned = "0.00 " Then
                          Application.Wait (Now + TimeValue("00:00:01"))
                            SAP_Session.FindById("wnd[0]/tbar[0]/btn[11]").press
                            Range(Cells(4, X), Cells(lastrowofcolumn, X)).Style = "Good"
                            clearingdoc = SAP_Session.FindById("wnd[0]/sbar").Text
                            Range(Cells(2, X), Cells(2, X)).Value = clearingdoc
                            Application.Wait (Now + TimeValue("00:00:03"))
                         Else
                          '  Application.Wait (Now + TimeValue("00:00:01"))
                            Range(Cells(4, X), Cells(lastrowofcolumn, X)).Style = "Bad"
                            SAP_Session.FindById("wnd[0]/tbar[0]/btn[12]").press
                            SAP_Session.FindById("wnd[0]/tbar[0]/btn[12]").press
                            SAP_Session.FindById("wnd[1]/usr/btnSPOP-OPTION1").press
                            
                        End If
                        

End If


                       
                        
                           If errMessage = "No open items were found" Then
                                          '  Application.Wait (Now + TimeValue("00:00:01"))
                                            Range(Cells(4, X), Cells(lastrowofcolumn, X)).Style = "Neutral"
                                            SAP_Session.FindById("wnd[0]/tbar[0]/btn[12]").press
                                            SAP_Session.FindById("wnd[0]/tbar[0]/btn[12]").press
                                            SAP_Session.FindById("wnd[1]/usr/btnSPOP-OPTION1").press
                        End If
        
        
        
                         X = X + 1




Loop






 
 
 
 Set SAP_Application = Nothing

 
 
MsgBox "The macro has finished .  Please press OK, and then close the two Excel windows that open after the macro ends."


Exit Sub
ErrSap:
MsgBox "Error.  Please press OK to end the macro.  ."





End Sub


Sub GL1130_PullGL_Activity_Main_V2_SessionAlreadyOpen()

On Error GoTo ErrSap



If Not IsObject(SapApplication) Then
   Set SapGuiAuto = GetObject("SAPGUI")
   Set SapApplication = SapGuiAuto.GetScriptingEngine
End If
If Not IsObject(Connection) Then
   Set Connection = SapApplication.Children(2)
End If
If Not IsObject(session) Then
   Set SAP_Session = Connection.Children(0)
End If
If IsObject(WScript) Then
   WScript.ConnectObject SAP_Session, "on"
   WScript.ConnectObject SapApplication, "on"
End If


Dim ORF_WB As Workbook
Set ORF_WB = ThisWorkbook


'=======================================================this code is for maximizing or minimizing the window==================
SessionHWND = SAP_Session.FindById("wnd[0]").Handle
ActivateWindow (SessionHWND)

'Start of your code
'Your code
'End of your code

'DeActivateWindow (SessionHWND)
'ActivateWindow (Application.hwnd) 'ExcelWBInFront



'=======================================================================================================================

Application.Wait (Now + TimeValue("00:00:01") / 1000)

Application.Wait (Now + TimeValue("00:00:01") / 1000)
SAP_Session.FindById("wnd[0]/tbar[0]/okcd").Text = "/nFAGLB03 "
Application.Wait (Now + TimeValue("00:00:01") / 1000)
SAP_Session.FindById("wnd[0]").sendVKey 0
Application.Wait (Now + TimeValue("00:00:01") / 1000)




Dim GLAccount As Long, FiscalYear As Long, ReconMonth As Variant, ReconMonth_Num As Long

GLAccount = ORF_WB.Sheets("Macro Input").Range("GL_Account")
FiscalYear = ORF_WB.Sheets("Macro Input").Range("Fiscal_Year")
ReconMonth = ORF_WB.Sheets("Macro Input").Range("Recon_Month")
ReconMonth_Num = ORF_WB.Sheets("Macro Input").Range("ReconMonth_Num")
Application.Wait (Now + TimeValue("00:00:01") / 1000)





Application.Wait (Now + TimeValue("00:00:01") / 1000)
SAP_Session.FindById("wnd[0]/usr/ctxtRACCT-LOW").Text = GLAccount
Application.Wait (Now + TimeValue("00:00:01") / 1000)
SAP_Session.FindById("wnd[0]/usr/txtRYEAR").Text = FiscalYear
Application.Wait (Now + TimeValue("00:00:01") / 1000)
SAP_Session.FindById("wnd[0]/tbar[1]/btn[8]").press
Application.Wait (Now + TimeValue("00:00:01"))


SAP_Session.FindById("wnd[0]/usr/cntlFDBL_BALANCE_CONTAINER/shellcont/shell").PressToolbarContextButton "&MB_EXPORT"
Application.Wait (Now + TimeValue("00:00:01") / 1000)
SAP_Session.FindById("wnd[0]/usr/cntlFDBL_BALANCE_CONTAINER/shellcont/shell").SelectContextMenuItem "&XXL"
Application.Wait (Now + TimeValue("00:00:01") / 1000)
SAP_Session.FindById("wnd[1]/usr/ctxtDY_PATH").Text = "C:\\TEMP"
Application.Wait (Now + TimeValue("00:00:01") / 1000)
SAP_Session.FindById("wnd[1]/usr/ctxtDY_FILENAME").Text = "EXPORT.MHTML"
Application.Wait (Now + TimeValue("00:00:01") / 1000)
SAP_Session.FindById("wnd[1]/tbar[0]/btn[11]").press
Application.Wait (Now + TimeValue("00:00:01") / 1000)

'Wait manually for 4 seconds until Excel Export of GL Balances generates
Application.Wait (Now + TimeValue("00:00:04"))


Dim GL_Export_WB As Workbook
Application.Wait (Now + TimeValue("00:00:01"))

Set GL_Export_WB = Workbooks.Open("C:\TEMP\EXPORT.MHTML")

GL_Export_WB.Sheets(1).Copy After:=ORF_WB.Sheets("Macro Input")

ORF_WB.ActiveSheet.Name = ReconMonth & "_GL 1130 Bal"

Application.Wait (Now + TimeValue("00:00:01"))
GL_Export_WB.Close SaveChanges:=False
Application.Wait (Now + TimeValue("00:00:01"))





ORF_WB.Sheets("Macro Input").Activate
 
SAP_Session.FindById("wnd[0]").Maximize

ActivateWindow (SessionHWND)
 
SAP_Session.FindById("wnd[0]/usr/cntlFDBL_BALANCE_CONTAINER/shellcont/shell").SetCurrentCell ReconMonth_Num, "BALANCE"
SAP_Session.FindById("wnd[0]/usr/cntlFDBL_BALANCE_CONTAINER/shellcont/shell").DoubleClickCurrentCell
 
 
 
 
 
 '=====================CHANGE LAYOUT MANUALLY TO ORF RECON======================================================
 SAP_Session.FindById("wnd[0]/tbar[1]/btn[32]").press
SAP_Session.FindById("wnd[1]/usr/btnAPP_FL_ALL").press

SAP_Session.FindById("wnd[1]/usr/btnB_SEARCH").press
SAP_Session.FindById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "Fund"
SAP_Session.FindById("wnd[2]").sendVKey 0
SAP_Session.FindById("wnd[1]/usr/btnAPP_WL_SING").press

SAP_Session.FindById("wnd[1]/usr/btnB_SEARCH").press
SAP_Session.FindById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "G/L"
SAP_Session.FindById("wnd[2]").sendVKey 0
SAP_Session.FindById("wnd[1]/usr/btnAPP_WL_SING").press

SAP_Session.FindById("wnd[1]/usr/btnB_SEARCH").press
SAP_Session.FindById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "Document Number"
SAP_Session.FindById("wnd[2]").sendVKey 0
SAP_Session.FindById("wnd[1]/usr/btnAPP_WL_SING").press

SAP_Session.FindById("wnd[1]/usr/btnB_SEARCH").press
SAP_Session.FindById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "Posting Date"
SAP_Session.FindById("wnd[2]").sendVKey 0
SAP_Session.FindById("wnd[1]/usr/btnAPP_WL_SING").press

SAP_Session.FindById("wnd[1]/usr/btnB_SEARCH").press
SAP_Session.FindById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "Document Date"
SAP_Session.FindById("wnd[2]").sendVKey 0
SAP_Session.FindById("wnd[1]/usr/btnAPP_WL_SING").press

SAP_Session.FindById("wnd[1]/usr/btnB_SEARCH").press
SAP_Session.FindById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "Amount"
SAP_Session.FindById("wnd[2]").sendVKey 0
SAP_Session.FindById("wnd[1]/usr/btnAPP_WL_SING").press

SAP_Session.FindById("wnd[1]/usr/btnB_SEARCH").press
SAP_Session.FindById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "Assignment"
SAP_Session.FindById("wnd[2]").sendVKey 0
SAP_Session.FindById("wnd[1]/usr/btnAPP_WL_SING").press

SAP_Session.FindById("wnd[1]/usr/btnB_SEARCH").press
SAP_Session.FindById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "Document header text"
SAP_Session.FindById("wnd[2]").sendVKey 0
SAP_Session.FindById("wnd[1]/usr/btnAPP_WL_SING").press

SAP_Session.FindById("wnd[1]/usr/btnB_SEARCH").press
SAP_Session.FindById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "Reference"
SAP_Session.FindById("wnd[2]").sendVKey 0
SAP_Session.FindById("wnd[1]/usr/btnAPP_WL_SING").press

SAP_Session.FindById("wnd[1]/usr/btnB_SEARCH").press
SAP_Session.FindById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "text"
SAP_Session.FindById("wnd[2]").sendVKey 0
SAP_Session.FindById("wnd[1]/usr/btnAPP_WL_SING").press

SAP_Session.FindById("wnd[1]/usr/btnB_SEARCH").press
SAP_Session.FindById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "vendor"
SAP_Session.FindById("wnd[2]").sendVKey 0
SAP_Session.FindById("wnd[1]/usr/btnAPP_WL_SING").press

SAP_Session.FindById("wnd[1]/usr/btnB_SEARCH").press
SAP_Session.FindById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "budget period"
SAP_Session.FindById("wnd[2]").sendVKey 0
SAP_Session.FindById("wnd[1]/usr/btnAPP_WL_SING").press

SAP_Session.FindById("wnd[1]/usr/btnB_SEARCH").press
SAP_Session.FindById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "document type"
SAP_Session.FindById("wnd[2]").sendVKey 0
SAP_Session.FindById("wnd[1]/usr/btnAPP_WL_SING").press

SAP_Session.FindById("wnd[1]/usr/btnB_SEARCH").press
SAP_Session.FindById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "payment method"
SAP_Session.FindById("wnd[2]").sendVKey 0
SAP_Session.FindById("wnd[1]/usr/btnAPP_WL_SING").press

SAP_Session.FindById("wnd[1]/usr/btnB_SEARCH").press
SAP_Session.FindById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "clearing date"
SAP_Session.FindById("wnd[2]").sendVKey 0
SAP_Session.FindById("wnd[1]/usr/btnAPP_WL_SING").press

SAP_Session.FindById("wnd[1]/usr/btnB_SEARCH").press
SAP_Session.FindById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "clearing document"
SAP_Session.FindById("wnd[2]").sendVKey 0
SAP_Session.FindById("wnd[1]/usr/btnAPP_WL_SING").press

SAP_Session.FindById("wnd[1]/usr/btnB_SEARCH").press
SAP_Session.FindById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "value date"
SAP_Session.FindById("wnd[2]").sendVKey 0
SAP_Session.FindById("wnd[1]/usr/btnAPP_WL_SING").press

SAP_Session.FindById("wnd[1]/tbar[0]/btn[0]").press

 '=====================CHANGE LAYOUT MANUALLY TO ORF RECON ENDS HERE======================================================
 
 
 SAP_Session.FindById("wnd[0]/mbar/menu[0]/menu[3]/menu[1]").Select
SAP_Session.FindById("wnd[1]/usr/ctxtDY_PATH").Text = "C:\\TEMP"
SAP_Session.FindById("wnd[1]/usr/ctxtDY_FILENAME").Text = "EXPORT2.MHTML"
SAP_Session.FindById("wnd[1]/tbar[0]/btn[11]").press
Application.Wait (Now + TimeValue("00:00:10"))
  
  
  
 Set GL_Export_WB2 = Workbooks.Open("C:\TEMP\EXPORT2.MHTML")

GL_Export_WB2.Sheets(1).Copy After:=ORF_WB.Sheets("Macro Input")

ORF_WB.ActiveSheet.Name = ReconMonth & "_GL 1130 Detail"

Application.Wait (Now + TimeValue("00:00:01"))
GL_Export_WB2.Close SaveChanges:=False
Application.Wait (Now + TimeValue("00:00:01"))
 
 
' DeActivateWindow (SessionHWND)
ActivateWindow (Application.hwnd) 'ExcelWBInFront
 
 

 Dim GL2 As Worksheet
 Set GL2 = ORF_WB.ActiveSheet

GL2.Sort.SortFields.Clear

GL2.Sort.SortFields.Add2 _
Key:=Range("J1:J100000"), SortOn:=xlSortOnValues, Order:=xlAscending, _
DataOption:=xlSortTextAsNumbers

GL2.Sort.SortFields.Add2 _
Key:=Range("I1:I100000"), SortOn:=xlSortOnValues, Order:=xlAscending, _
DataOption:=xlSortTextAsNumbers

With GL2.Sort
        .SetRange Range("A1:Q100000")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
End With
 
 
 
 
 

'MsgBox "Macro Paused.  Press OK to shut down BD.  It will currently be open in a window of the Excel application."


'Shutdown the connection
'Set session = Nothing
'Connection.CloseSession ("ses[0]")
'Set Connection = Nothing

'Wait a bit for the connection to be closed completely
'Application.Wait (Now + TimeValue("00:00:01") / 1000)
'Set SAP_Application = Nothing


'Kill ("C:\TEMP\EXPORT.MHTML")
'Kill ("C:\TEMP\EXPORT2.MHTML")


 
 
 
 
 
 
 
MsgBox "The macro has finished.  Please press OK.  Two subsequent alert windows will pop up indicating Excel cannot open a file.  This is normal and means the file exports were deleted by the macro from the TEMP folder, otherwise they will automatically open once the macro ends.  All file exports should now be moved into your recon workbook."


Exit Sub
ErrSap:
MsgBox "SAP not opened.  Open SAP and go to a random T-code such as FB03.  Make sure SAP Scripting is enabled, and security pop ups for saving files are disabled (see instructions tab)."



End Sub

Sub AddChecktoBal()


Dim ORF_WB As Workbook
Set ORF_WB = ThisWorkbook

GLAccount = ORF_WB.Sheets("Macro Input").Range("GL_Account")
FiscalYear = ORF_WB.Sheets("Macro Input").Range("Fiscal_Year")
ReconMonth = ORF_WB.Sheets("Macro Input").Range("Recon_Month")
ReconMonth_Num = ORF_WB.Sheets("Macro Input").Range("ReconMonth_Num")



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


