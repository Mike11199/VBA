Attribute VB_Name = "ZAPCLAIM_SP01_mANUAL"
Public BDPass As String
Public BDUserName As String

Sub Run_ZAPCLAIM_SPO1_Manual()

On Error GoTo ErrSap

Application.DisplayAlerts = False
DoEvents


Dim wb As Workbook
Set wb = ThisWorkbook


'for np both paper and eft use req issue date and paper issue date not used at all
ZAPCLAIM_Req_Issue_Date = wb.Sheets("Macro Input").Range("ZAPCLAIM_Req_Issue_Date")
ZAPCLAIM_Paper_Issue_Date = wb.Sheets("Macro Input").Range("ZAPCLAIM_Paper_Issue_Date")

answer3 = MsgBox("Are you sure you want the following issue dates for the claims?" & vbNewLine & vbNewLine & _
"Issue date for the Paper Claim:     " & ZAPCLAIM_Paper_Issue_Date & vbNewLine & vbNewLine & _
"Issue date for the EFT Claim:       " & ZAPCLAIM_Req_Issue_Date & vbNewLine & vbNewLine & vbNewLine & _
"If a Paper or EFT claim exists for this month, these will be the dates used. Change these in the Macro Input box on this sheet if these dates are incorrect." & vbNewLine & vbNewLine & _
"Please note that non-periodics use the ZAPCLAIM_Req_Issue_Date for both paper and EFT claims.", vbOKCancel)

If answer3 = vbCancel Then
    MsgBox "Macro cancelled by user"
    End
End If

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







'Generate claim and print claim schedule for paper allowance roll payments
                
                Sleep 1
                SAP_Session.FindById("wnd[0]/tbar[0]/okcd").Text = "/nZAPCLAIM "
                Sleep 1
                SAP_Session.FindById("wnd[0]").sendVKey 0
                Sleep 1
                
                
                SAP_Session.FindById("wnd[0]/usr/ctxtP_CSTYP").Text = "CBPAC"
                SAP_Session.FindById("wnd[0]/usr/ctxtSO_LAUFD-LOW").Text = F110_Run_Date
                
                'paper always has 1st as the issue date
                SAP_Session.FindById("wnd[0]/usr/ctxtP_RQDAT").Text = ZAPCLAIM_Paper_Issue_Date
                SAP_Session.FindById("wnd[0]/usr/chkP_UPD").Selected = True
                SAP_Session.FindById("wnd[0]/usr/txtSO_LAUFI-LOW").Text = "CB0M"
                SAP_Session.FindById("wnd[0]/tbar[1]/btn[8]").press
                
                
                
                Paper_Claim_Number = SAP_Session.FindById("wnd[0]/usr/lbl[15,2]").Text
                wb.Sheets("Macro Input").Range("CS_3") = Paper_Claim_Number
                SAP_Session.FindById("wnd[0]/tbar[0]/btn[15]").press
                
                
                
                SAP_Session.FindById("wnd[0]/usr/radP_PRI").Select
                SAP_Session.FindById("wnd[0]/usr/chkP_FSIND").Selected = True
                SAP_Session.FindById("wnd[0]/usr/ctxtP_CSNBR").Text = Paper_Claim_Number
                SAP_Session.FindById("wnd[0]/usr/ctxtP_FSHPR").Text = "LOCL"
                SAP_Session.FindById("wnd[0]/usr/ctxtP_FSFRM").Text = "ZFI_CLAIM_FACESHEET"
                SAP_Session.FindById("wnd[0]/usr/ctxtP_FSHPR").SetFocus
                SAP_Session.FindById("wnd[0]/usr/ctxtP_FSHPR").CaretPosition = 4
                
                
                 'add this to make sure the Remittance Advice Printer is NOT checked and enter to refresh the claim schedule type
                SAP_Session.FindById("wnd[0]/usr/ctxtP_FSHPR").SetFocus
                SAP_Session.FindById("wnd[0]/usr/ctxtP_FSHPR").CaretPosition = 4
                SAP_Session.FindById("wnd[0]").sendVKey 0
                SAP_Session.FindById("wnd[0]/usr/chkP_RAIND").Selected = False
                
                
                'execute
                SAP_Session.FindById("wnd[0]/tbar[1]/btn[8]").press
                
                Printed = SAP_Session.FindById("wnd[0]/usr/lbl[26,2]").Text
                
                If Printed = "Printed succesfully" Then
                     wb.Sheets("Macro Input").Range("CS_3_PRINT") = "X"
                End If
                
                SAP_Session.FindById("wnd[0]/tbar[0]/btn[15]").press


 
 
  Set SAP_Application = Nothing
 
 
MinutesElapsed = Format((Timer - StartTime) / 86400, "hh:mm:ss")
 
 
 MsgBox "Ran successfully in " & MinutesElapsed & " minutes." & vbNewLine & vbNewLine & _
"The macro has finished running ZAPCLAIM for CB Manual-Rollover Payments." & vbNewLine & vbNewLine & _
"Please press OK."



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
    ActiveCell.FormulaR1C1 = "=TEMPLATE_SUMMARY_NET_DB"
    Range("N41").Select
    ActiveCell.FormulaR1C1 = "=TEMPLATE_SUMMARY_NET_SB"
    Range("N42").Select
    ActiveCell.FormulaR1C1 = "=TEMPLATE_SUMMARY_NET_SR"
    Range("O40").Select
    ActiveCell.FormulaR1C1 = "DB"
    Range("O41").Select
    ActiveCell.FormulaR1C1 = "SB"
    Range("O42").Select
    ActiveCell.FormulaR1C1 = "SR"
    Range("O40:O42").Select
    With Selection.Font
        .Color = -16776961
        .TintAndShade = 0
    End With
    Selection.Font.Bold = True
    Range("N43").Select
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
    Range("N43").Select
    Selection.Font.Bold = True
    Range("O43").Select
    ActiveCell.FormulaR1C1 = "Check"
    Range("N43").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=SUM(R[-3]C:R[-1]C)"
    Range("O43").Select
    Selection.ClearContents
    Range("O44").Select
    ActiveCell.FormulaR1C1 = "Check"
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
    


End Sub





