Attribute VB_Name = "Run_FV60"
Public BDPass As String
Public BDUserName As String

Sub Run_FV60()

On Error GoTo ErrSap


Dim ORF_WB As Workbook
Set ORF_WB = ThisWorkbook

BD_LOG_ON.Show
SapConnectionString = ORF_WB.Sheets("Macro Input").Range("SAP_Connection")


save_as_completed = MsgBox("Save each document as completed?  This needs to be done eventually so each can be released.", vbYesNo)

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



'    ActivateWindow (Application.hwnd)    'Brings the Excel window to the front
'
'
'ActivateWindow (SessionHWND)


Dim GLAccount As Long, FiscalYear As Long, ReconMonth As Variant, ReconMonth_Num As Long


FiscalYear = ORF_WB.Sheets("Macro Input").Range("Fiscal_Year")
ReconMonth = ORF_WB.Sheets("Macro Input").Range("Recon_Month")
ReconMonth_Num = ORF_WB.Sheets("Macro Input").Range("ReconMonth_Num")
Invoice_Date = ORF_WB.Sheets("Macro Input").Range("INVOICE_DATE")
Sleep 1


Template_End_Row = ThisWorkbook.Sheets("Template").Range("TEMPLATE_SUMMARY").row - 1
Data_End_Row = ThisWorkbook.Sheets("Template").Range("B" & Template_End_Row).End(xlUp).row

Dim DataRange As Variant
Dim Name As String

DataRange = ThisWorkbook.Sheets("Template").Range("A8:T" & Data_End_Row)


For i = 1 To UBound(DataRange)

                If Not IsEmpty(DataRange(i, 2)) Then
                
                Sleep 1
                SAP_Session.FindById("wnd[0]/tbar[0]/okcd").Text = "/nFV60 "
                Sleep 1
                SAP_Session.FindById("wnd[0]").sendVKey 0
                Sleep 1
                
                
                Name = "*" & Trim(CStr(DataRange(i, 4))) & "*" & Trim(CStr(DataRange(i, 5))) & "*"
                Name = CStr(Clean_NonPrintableCharacters(Name))
                
                
                SAP_Session.FindById("wnd[0]").sendVKey 4
                SAP_Session.FindById("wnd[1]/usr/tabsG_SELONETABSTRIP/tabpTAB001/ssubSUBSCR_PRESEL:SAPLSDH4:0220/sub:SAPLSDH4:0220/txtG_SELFLD_TAB-LOW[4,24]").Text = Name
                SAP_Session.FindById("wnd[1]/tbar[0]/btn[0]").press
                SAP_Session.FindById("wnd[1]").sendVKey 0
                
                SAP_Session.FindById("wnd[0]/usr/tabsTS/tabpMAIN/ssubPAGE:SAPLFDCB:0010/ctxtINVFO-BLDAT").Text = Invoice_Date
                
                If i = 1 Then
                        SAP_Session.FindById("wnd[0]/tbar[1]/btn[16]").press
                        SAP_Session.FindById("wnd[0]/usr/tabsTS/tabp1100/ssubS1100:SAPMF05O:1100/cmbRFOPTE-DMTTP").Key = "2"
                        SAP_Session.FindById("wnd[0]/tbar[0]/btn[3]").press
                End If
                
                SAP_Session.FindById("wnd[0]/usr/tabsTS/tabpMAIN/ssubPAGE:SAPLFDCB:0010/cmbINVFO-BLART").SetFocus
                SAP_Session.FindById("wnd[0]/usr/tabsTS/tabpMAIN/ssubPAGE:SAPLFDCB:0010/cmbINVFO-BLART").Key = "ZI"
                
                
                
                
                SAP_Session.FindById("wnd[0]/usr/tabsTS/tabpMAIN/ssubPAGE:SAPLFDCB:0010/txtINVFO-WRBTR").Text = DataRange(i, 6)
                SAP_Session.FindById("wnd[0]/usr/tabsTS/tabpMAIN/ssubPAGE:SAPLFDCB:0010/txtINVFO-XBLNR").Text = DataRange(i, 2) & " " & Left(DataRange(i, 5), 5)
                
                
                SAP_Session.FindById("wnd[0]/usr/subITEMS:SAPLFSKB:0100/tblSAPLFSKBTABLE/txtACGL_ITEM-WRBTR[4,0]").Text = DataRange(i, 6)
                SAP_Session.FindById("wnd[0]/usr/subITEMS:SAPLFSKB:0100/tblSAPLFSKBTABLE/ctxtACGL_ITEM-GEBER[33,0]").SetFocus
                SAP_Session.FindById("wnd[0]/usr/subITEMS:SAPLFSKB:0100/tblSAPLFSKBTABLE/ctxtACGL_ITEM-GEBER[33,0]").Text = "0835300000"
                SAP_Session.FindById("wnd[0]/usr/subITEMS:SAPLFSKB:0100/tblSAPLFSKBTABLE/ctxtACGL_ITEM-BUDGET_PD[34,0]").SetFocus
                SAP_Session.FindById("wnd[0]/usr/subITEMS:SAPLFSKB:0100/tblSAPLFSKBTABLE/ctxtACGL_ITEM-BUDGET_PD[34,0]").Text = "504-000-76"
                
                
                
                If DataRange(i, 3) = "Service Retirement" Then
                        SAP_Session.FindById("wnd[0]/usr/subITEMS:SAPLFSKB:0100/tblSAPLFSKBTABLE/ctxtACGL_ITEM-HKONT[1,0]").Text = "9041623010"
                ElseIf DataRange(i, 3) = "Disability Retirement" Then
                        SAP_Session.FindById("wnd[0]/usr/subITEMS:SAPLFSKB:0100/tblSAPLFSKBTABLE/ctxtACGL_ITEM-HKONT[1,0]").Text = "9041624010"
                ElseIf DataRange(i, 3) = "Survivor Benefits" Then
                        SAP_Session.FindById("wnd[0]/usr/subITEMS:SAPLFSKB:0100/tblSAPLFSKBTABLE/ctxtACGL_ITEM-HKONT[1,0]").Text = "9041625010"
                End If
                
                
                SAP_Session.FindById("wnd[0]/usr/tabsTS/tabpPAYM").Select
                
'                go to payee line and enter in id based on name from WO, e.g  charles schwab for rollover
                SAP_Session.FindById("wnd[0]/usr/tabsTS/tabpPAYM/ssubPAGE:SAPLFDCB:0020/ctxtINVFO-EMPFB").SetFocus
                SAP_Session.FindById("wnd[0]").sendVKey 4
                SAP_Session.FindById("wnd[1]/tbar[0]/btn[71]").press
                SAP_Session.FindById("wnd[2]/usr/txtRSYSF-STRING").Text = Left(DataRange(i, 20), 5)
                SAP_Session.FindById("wnd[2]/usr/txtRSYSF-STRING").CaretPosition = 4
                SAP_Session.FindById("wnd[2]/tbar[0]/btn[0]").press
                SAP_Session.FindById("wnd[3]/usr/lbl[1,2]").SetFocus
                SAP_Session.FindById("wnd[3]/usr/lbl[1,2]").CaretPosition = 6
                SAP_Session.FindById("wnd[3]/tbar[0]/btn[0]").press
                SAP_Session.FindById("wnd[1]/tbar[0]/btn[0]").press
                
'Enter bseline date
                SAP_Session.FindById("wnd[0]/usr/tabsTS/tabpPAYM/ssubPAGE:SAPLFDCB:0020/ctxtINVFO-ZFBDT").Text = Invoice_Date
                
'Enter payment method
                SAP_Session.FindById("wnd[0]/usr/tabsTS/tabpPAYM/ssubPAGE:SAPLFDCB:0020/ctxtINVFO-ZLSCH").SetFocus
                SAP_Session.FindById("wnd[0]/usr/tabsTS/tabpPAYM/ssubPAGE:SAPLFDCB:0020/ctxtINVFO-ZLSCH").CaretPosition = 0
                SAP_Session.FindById("wnd[0]").sendVKey 4
                SAP_Session.FindById("wnd[1]/usr/lbl[4,14]").SetFocus
                SAP_Session.FindById("wnd[1]/usr/lbl[4,14]").CaretPosition = 9
                SAP_Session.FindById("wnd[1]").sendVKey 2
                
   'enter text rollover name under basic data
                SAP_Session.FindById("wnd[0]/usr/subITEMS:SAPLFSKB:0100/tblSAPLFSKBTABLE").Columns.ElementAt(2).Width = 10
                SAP_Session.FindById("wnd[0]/usr/cmbRF05A-BUSCS").SetFocus
                SAP_Session.FindById("wnd[0]/usr/tabsTS/tabpMAIN").Select
                SAP_Session.FindById("wnd[0]/usr/tabsTS/tabpMAIN/ssubPAGE:SAPLFDCB:0010/ctxtINVFO-SGTXT").Text = "Rollover - " & DataRange(i, 20)
                SAP_Session.FindById("wnd[0]/usr/tabsTS/tabpMAIN/ssubPAGE:SAPLFDCB:0010/ctxtINVFO-SGTXT").SetFocus
                SAP_Session.FindById("wnd[0]/usr/tabsTS/tabpMAIN/ssubPAGE:SAPLFDCB:0010/ctxtINVFO-SGTXT").CaretPosition = 33
                
                

                
                
                If save_as_completed = vbYes Then
                    SAP_Session.FindById("wnd[0]").sendVKey 42
                Else
                    'save button, NOT SAVE AS COMPLETED
                    SAP_Session.FindById("wnd[0]/tbar[0]/btn[11]").press
                End If
                
                

                
                
                
                DocumentNo = SAP_Session.FindById("wnd[0]/sbar").Text
                
                
                'these two get past if says gl code is related to tax, check code
                
                If InStr(1, CStr(DocumentNo), "check code", vbBinaryCompare) Then
                
                        SAP_Session.FindById("wnd[0]").sendVKey 0
                        DocumentNo = SAP_Session.FindById("wnd[0]/sbar").Text
                
                End If
                
                
                If InStr(1, CStr(DocumentNo), "check code", vbBinaryCompare) Then
                
                        SAP_Session.FindById("wnd[0]").sendVKey 0
                        DocumentNo = SAP_Session.FindById("wnd[0]/sbar").Text
                
                End If
                
                
                
                numdoc = 7 + i
                ThisWorkbook.Sheets("Template").Range("U" & numdoc) = Mid(DocumentNo, 9, 11)
                
                End If
Next i


 Set SAP_Application = Nothing

 
 
 
 
 
MinutesElapsed = Format((Timer - StartTime) / 86400, "hh:mm:ss")
 
 
 MsgBox "Ran successfully in " & MinutesElapsed & " minutes." & vbNewLine & vbNewLine & _
"The macro has finished adding the exported files to you current workbook." & vbNewLine & vbNewLine & _
"Please press OK, and then close the two alert messages that open after the macro ends."



Exit Sub
ErrSap:
MsgBox "Error.  Please press OK to end the macro."





End Sub

Sub AddChecktoBal()


Dim ORF_WB As Workbook
Set ORF_WB = ThisWorkbook


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


Sub manual_change_layout()


 'Switch to this old manual way if layout deleted somehow or error
 
'
' '=====================CHANGE LAYOUT MANUALLY TO ORF RECON======================================================
'SAP_Session.FindById("wnd[0]/tbar[1]/btn[32]").press
'SAP_Session.FindById("wnd[1]/usr/btnAPP_FL_ALL").press
'
'
'SAP_Session.FindById("wnd[1]").sendVKey 42
'
'
'SAP_Session.FindById("wnd[1]/usr/btnB_SEARCH").press
''SAP_Session.FindById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "Fund"
'SAP_Session.FindById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "GEBER"
'SAP_Session.FindById("wnd[2]").sendVKey 0
'SAP_Session.FindById("wnd[1]/usr/btnAPP_WL_SING").press
'
'SAP_Session.FindById("wnd[1]/usr/btnB_SEARCH").press
''SAP_Session.FindById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "G/L Account"
'SAP_Session.FindById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "HKONT"
'SAP_Session.FindById("wnd[2]").sendVKey 0
'SAP_Session.FindById("wnd[1]/usr/btnAPP_WL_SING").press
'
'SAP_Session.FindById("wnd[1]/usr/btnB_SEARCH").press
''SAP_Session.FindById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "Document Number"
'SAP_Session.FindById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "BELNR"
'SAP_Session.FindById("wnd[2]").sendVKey 0
'SAP_Session.FindById("wnd[1]/usr/btnAPP_WL_SING").press
'
'SAP_Session.FindById("wnd[1]/usr/btnB_SEARCH").press
''SAP_Session.FindById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "Posting Date"
'SAP_Session.FindById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "BUDAT"
'SAP_Session.FindById("wnd[2]").sendVKey 0
'SAP_Session.FindById("wnd[1]/usr/btnAPP_WL_SING").press
'
'SAP_Session.FindById("wnd[1]/usr/btnB_SEARCH").press
''SAP_Session.FindById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "Document Date"
'SAP_Session.FindById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "BLDAT"
'SAP_Session.FindById("wnd[2]").sendVKey 0
'SAP_Session.FindById("wnd[1]/usr/btnAPP_WL_SING").press
'
'SAP_Session.FindById("wnd[1]/usr/btnB_SEARCH").press
''SAP_Session.FindById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "Amount in doc. curr."
'SAP_Session.FindById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "WRSHB"
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
''SAP_Session.FindById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "Document header text"
'SAP_Session.FindById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "U_BKTXT"
'SAP_Session.FindById("wnd[2]").sendVKey 0
'SAP_Session.FindById("wnd[1]/usr/btnAPP_WL_SING").press
'
'
'
'SAP_Session.FindById("wnd[1]/usr/btnB_SEARCH").press
''SAP_Session.FindById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "Reference"
'SAP_Session.FindById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "XBLNR"
'SAP_Session.FindById("wnd[2]").sendVKey 0
'SAP_Session.FindById("wnd[1]/usr/btnAPP_WL_SING").press
'
'
'
'SAP_Session.FindById("wnd[1]/usr/btnB_SEARCH").press
''SAP_Session.FindById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "text"
'SAP_Session.FindById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "SGTXT"
'SAP_Session.FindById("wnd[2]").sendVKey 0
'SAP_Session.FindById("wnd[1]/usr/btnAPP_WL_SING").press
'
'SAP_Session.FindById("wnd[1]/usr/btnB_SEARCH").press
''SAP_Session.FindById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "vendor"
'SAP_Session.FindById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "LIFNR"
'SAP_Session.FindById("wnd[2]").sendVKey 0
'SAP_Session.FindById("wnd[1]/usr/btnAPP_WL_SING").press
'
'SAP_Session.FindById("wnd[1]/usr/btnB_SEARCH").press
''SAP_Session.FindById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "budget period"
'SAP_Session.FindById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "U_BUDGET_PD"
'SAP_Session.FindById("wnd[2]").sendVKey 0
'SAP_Session.FindById("wnd[1]/usr/btnAPP_WL_SING").press
'
'SAP_Session.FindById("wnd[1]/usr/btnB_SEARCH").press
''SAP_Session.FindById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "document type"
'SAP_Session.FindById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "BLART"
'SAP_Session.FindById("wnd[2]").sendVKey 0
'SAP_Session.FindById("wnd[1]/usr/btnAPP_WL_SING").press
'
'SAP_Session.FindById("wnd[1]/usr/btnB_SEARCH").press
''SAP_Session.FindById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "payment method"
'SAP_Session.FindById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "ZLSCH"
'SAP_Session.FindById("wnd[2]").sendVKey 0
'SAP_Session.FindById("wnd[1]/usr/btnAPP_WL_SING").press
'
'SAP_Session.FindById("wnd[1]/usr/btnB_SEARCH").press
''SAP_Session.FindById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "clearing date"
'SAP_Session.FindById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "AUGDT"
'SAP_Session.FindById("wnd[2]").sendVKey 0
'SAP_Session.FindById("wnd[1]/usr/btnAPP_WL_SING").press
'
'SAP_Session.FindById("wnd[1]/usr/btnB_SEARCH").press
''SAP_Session.FindById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "clearing document"
'SAP_Session.FindById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "AUGBL"
'SAP_Session.FindById("wnd[2]").sendVKey 0
'SAP_Session.FindById("wnd[1]/usr/btnAPP_WL_SING").press
'
'SAP_Session.FindById("wnd[1]/usr/btnB_SEARCH").press
''SAP_Session.FindById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "value date"
'SAP_Session.FindById("wnd[2]/usr/txtGD_SEARCHSTR").Text = "VALUT"
'SAP_Session.FindById("wnd[2]").sendVKey 0
'SAP_Session.FindById("wnd[1]/usr/btnAPP_WL_SING").press
'
'SAP_Session.FindById("wnd[1]/tbar[0]/btn[0]").press


 '=====================CHANGE LAYOUT MANUALLY TO ORF RECON ENDS HERE======================================================
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
