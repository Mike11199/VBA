Attribute VB_Name = "Import_CB_Allowance"
Public BDPass As String
Public BDUserName As String

Sub Import_Template_Allowance_CB()

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
                SAP_Session.FindById("wnd[0]/tbar[0]/okcd").Text = "/nzfimassinvpost "
                Sleep 1
                SAP_Session.FindById("wnd[0]").sendVKey 0
                Sleep 1
                
                
                
                SAP_Session.FindById("wnd[0]/usr/radP_CBFILE").Select
                SAP_Session.FindById("wnd[0]/usr/ctxtP_IN_S").SetFocus
                SAP_Session.FindById("wnd[0]/usr/ctxtP_IN_S").CaretPosition = 0
                
                
                                ActivateWindow (Application.hwnd)    'Brings the Excel window to the front
                
              
                
                
                
                
                'Skip?
                answer2 = MsgBox("Skip IRS/FTB Posting if running import multiple times?" & vbNewLine & vbNewLine & "This macro should be run from the G: drive, NOT the desktop, from the folder that the CB Work Orders are saved in.", vbQuestion + vbYesNoCancel + vbDefaultButton2, "Skip IRS/FTB Posting?")
                
                
                  ActivateWindow (SessionHWND)
                
                Sleep 500
                
                If answer2 = vbYes Then
                        SAP_Session.FindById("wnd[0]/usr/chkP_OTH").Selected = True
                End If
                
                If answer2 = vbNo Then
                        SAP_Session.FindById("wnd[0]/usr/chkP_OTH").Selected = False
                End If
                
                    If answer2 = vbCancel Then
                        ActivateWindow (Application.hwnd)
                       MsgBox ("Macro cancelled by user.")
                       End
                End If
                
                
                
                
                
                Sleep 300
                
                ActivateWindow (SessionHWND)
                

                
        Dim wb2 As Workbook
        Set wb2 = Workbooks.Add
        ThisWorkbook.Sheets("Template").Copy Before:=wb2.Sheets(1)
        
        
        Dim relativePath As String
        
        'make sure the relativePath is less than 120 characters or else SAP will consider it to be 'too long' and won't accept it.
        relativePath = ThisWorkbook.Path & Application.PathSeparator & "Import"
        wb2.SaveAs Filename:=relativePath
        wb2.Close

              
        SAP_Session.FindById("wnd[0]/usr/ctxtP_IN_S").Text = relativePath & ".xlsx"
        SAP_Session.FindById("wnd[0]/usr/ctxtP_OUT_S").Text = ThisWorkbook.Path & Application.PathSeparator
        SAP_Session.FindById("wnd[0]/usr/ctxtP_ERR_S").Text = ThisWorkbook.Path & Application.PathSeparator
        SAP_Session.FindById("wnd[0]").sendVKey 8
                
        
        'if blocked by other user as SAP slow then press enter to get past it
        errMessage = SAP_Session.FindById("wnd[0]/sbar").Text
        
        Do While errMessage = ""
        
                errMessage = SAP_Session.FindById("wnd[0]/sbar").Text
                Sleep 1000
        
        Loop


Dim test1, test2, test3, test4 As String

test1 = SAP_Session.FindById("wnd[0]/usr/lbl[0,20]").Text
test2 = SAP_Session.FindById("wnd[0]/usr/lbl[0,21]").Text
test3 = SAP_Session.FindById("wnd[0]/usr/lbl[0,22]").Text
test4 = SAP_Session.FindById("wnd[0]/usr/lbl[0,23]").Text


                
            
            ActivateWindow (SessionHWND)
            SAP_Session.FindById("wnd[0]").Maximize
            
            ActivateWindow (SessionHWND)
            ActivateWindow (SessionHWND)
            ActivateWindow (SessionHWND)
            Sleep 2000
            ActivateWindow (SessionHWND)
            ActivateWindow (SessionHWND)
            Sleep 400
            
            
            
            Call PrintScreen
            
            Sleep 3000
                
      
                
                
                ActivateWindow (Application.hwnd)    'Brings the Excel window to the front


                wb.Sheets.Add(After:=Sheets("Macro Input")).Name = "INVPOST_Results_" & wb.Sheets.Count
                
                Dim screenshot_sheet As Worksheet
                Set screenshot_sheet = wb.Sheets("INVPOST_Results_" & wb.Sheets.Count - 1)
                
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
                'Call PictureFormat
                Sleep 2000
                
               screenshot_sheet.Range("A56") = errMessage
               
               screenshot_sheet.Range("A58") = test1
               screenshot_sheet.Range("A59") = test2
               screenshot_sheet.Range("A61") = test3
               screenshot_sheet.Range("A62") = test4




Dim Export1 As Workbook
Sleep 1000



Set Export1 = Workbooks.Open(test2)

Export1.Sheets(1).Copy After:=wb.Sheets("Macro Input")

wb.ActiveSheet.Name = "INVPOST_Posted_1_" & Sheets.Count

Dim postedsheet As Worksheet
Set postedsheet = wb.ActiveSheet

Columns("B:B").Select
With Selection
    .NumberFormat = "General"
    .Value = .Value
End With


With ActiveWorkbook.ActiveSheet.Tab
    .Color = 192
    .TintAndShade = 0
End With



Sleep 1000
Export1.Close SaveChanges:=False
Sleep 1000


'Only add second file showing postings and or errors if it exists.  If stars, it does not exist
'If Not InStr(1, "**", CStr(test4)) Then      ''not working
If Right(CStr(test4), 1) = "x" Or Right(CStr(test4), 1) = "s" Then

        Set Export1 = Workbooks.Open(test4)
        
        Export1.Sheets(1).Copy After:=wb.Sheets("Macro Input")
        
        wb.ActiveSheet.Name = "INVPOST_Error_or_Posted_2_" & Sheets.Count
        
        
        With ActiveWorkbook.ActiveSheet.Tab
            .Color = 192
            .TintAndShade = 0
        End With
        
        Sleep 1000
        Export1.Close SaveChanges:=False
        Sleep 1000

End If




wb.Sheets("Template").Activate

lastrow5 = wb.Sheets("Template").Range("TEMPLATE_SUMMARY").row - 1



    Range("T8:T" & lastrow5).FormulaR1C1 = _
        "=XLOOKUP(RC[-18]," & postedsheet.Name & "!C[-18]," & postedsheet.Name & "!C[-14])"
    
    Range("U8:U" & lastrow5).FormulaR1C1 = "=RC[-15]-RC[-1]"
    

Columns("T:U").Select
    With Selection.Font
        .Color = -16776961
        .TintAndShade = 0
    End With
    Selection.Style = "Comma"
    Selection.NumberFormat = "_(* #,##0.0_);_(* (#,##0.0);_(* ""-""??_);_(@_)"
    Selection.NumberFormat = "_(* #,##0_);_(* (#,##0);_(* ""-""??_);_(@_)"

 
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



