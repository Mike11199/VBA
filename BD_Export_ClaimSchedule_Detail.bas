Attribute VB_Name = "BD_Export_ClaimSchedule_Detail"
Public BDPass As String
Public BDUserName As String


Sub GL1130_ClaimSchedulePaymentINFO_LogInBox()



Dim ORF_WB As Workbook
Set ORF_WB = ThisWorkbook


Check = MsgBox(Prompt:="Did you run Macro #1 first? " & vbNewLine & vbNewLine & _
"This macro exports a Claim Schedule Detail Report, based on the Claim Schedule lines listed on the GL Detail." & vbNewLine & vbNewLine & _
"Without the GL Detail from Macro #1, this macro will error out.", Buttons:=vbQuestion + vbYesNo)

If Check = vbNo Then
MsgBox ("Macro cancelled.")
End
End If





BD_LOG_ON.Show

Dim StartTime As Double
Dim MinutesElapsed As String
'Remember time when macro starts
StartTime = Timer





Set SAP_Application = CreateObject("Sapgui.ScriptingCtrl.1")


Set Connection = SAP_Application.OpenConnection("EP0 - SAP ECC Production", True)
Set SAP_Session = Connection.Children(0)


'=======================================================this code is for maximizing or minimizing the window==================
SessionHWND = SAP_Session.FindById("wnd[0]").Handle

ActivateWindow (SessionHWND)
SAP_Session.FindById("wnd[0]").Maximize


'=======================================================================================================================

        
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

'========================================================================================================================


Dim GLAccount As Long, FiscalYear As Long, ReconMonth As Variant, ReconMonth_Num As Long, FCHN_From As Variant, FCHN_To As Variant

GLAccount = ORF_WB.Sheets("Macro Input").Range("GL_Account")
FiscalYear = ORF_WB.Sheets("Macro Input").Range("Fiscal_Year")
ReconMonth = ORF_WB.Sheets("Macro Input").Range("Recon_Month")
ReconMonth_Num = ORF_WB.Sheets("Macro Input").Range("ReconMonth_Num")
FCHN_From = ORF_WB.Sheets("Macro Input").Range("FCHN_From")
FCHN_To = ORF_WB.Sheets("Macro Input").Range("FCHN_To")


 
Dim GL_DetailSheet As Worksheet
Set GL_DetailSheet = ORF_WB.Sheets(ReconMonth & "_GL 1130 Detail")
GL_DetailSheet.Activate


GL_DetailSheet.Columns("J:J").Copy GL_DetailSheet.Columns("R:R")
GL_DetailSheet.Columns("R:R") = GL_DetailSheet.Columns("R:R").Value



Dim ClaimRange As Range
Dim cel As Range
Dim LastRow As Long
LastRow = GL_DetailSheet.Range("R100000").End(xlUp).row

Set ClaimRange = GL_DetailSheet.Range("R2:R" & LastRow)
ClaimRange.Select



For Each cel In ClaimRange.Cells

        'Removes X from claims that end with X, and set to zero claim adjustments that end with A as we don't want to pull those
        If Right(cel, 1) = "X" Then
            cel = Left(cel, Len(cel) - 1)
        End If
        
        If Right(cel, 1) = "A" Then
            cel = 0
        End If

Next


'Formatting blue cell with red text, bold
 Range("R1").Select
    With Selection.Font
        .Color = -16776961
        .TintAndShade = 0
    End With
    Selection.Font.Bold = True
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorLight2
        .TintAndShade = 0.399975585192419
        .PatternTintAndShade = 0
    End With
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0.399975585192419
        .PatternTintAndShade = 0
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Range("R1").Select
    ActiveCell.FormulaR1C1 = "Claim #s"
    Range("R2").Select
    Columns("R:R").ColumnWidth = 12.86
    Range("R1").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0.599993896298105
        .PatternTintAndShade = 0
    End With
    Range("R13").Select
    
'End formatting




ClaimRange.Copy



'=====================Add Claim #s here==================================================================


SAP_Session.FindById("wnd[0]/tbar[0]/okcd").Text = "/nZCSPAYMENTDISP"
SAP_Session.FindById("wnd[0]").sendVKey 0

SAP_Session.FindById("wnd[0]/usr/ctxtGD-VARIANT").Text = ""
SAP_Session.FindById("wnd[0]/usr/txtGD-MAX_LINES").Text = ""
SAP_Session.FindById("wnd[0]/usr/ctxtGD-VARIANT").Text = "/ORFCLAIM"
SAP_Session.FindById("wnd[0]").sendVKey 0
SAP_Session.FindById("wnd[0]/usr/ctxtGD-VARIANT").Text = "/ORFCLAIM"
SAP_Session.FindById("wnd[0]").sendVKey 0
SAP_Session.FindById("wnd[0]/usr/tblSAPLZUT_SE16N1SELFIELDS_TC/btnPUSH[4,5]").SetFocus
SAP_Session.FindById("wnd[0]/usr/tblSAPLZUT_SE16N1SELFIELDS_TC/btnPUSH[4,5]").press
SAP_Session.FindById("wnd[1]/tbar[0]/btn[24]").press
SAP_Session.FindById("wnd[1]/tbar[0]/btn[8]").press
SAP_Session.FindById("wnd[0]/tbar[1]/btn[8]").press



 
 '============EXPORT TO EXCEL FILE===========================================================================================
 
ActivateWindow (SessionHWND)
SAP_Session.FindById("wnd[0]").Maximize

SAP_Session.FindById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell").PressToolbarContextButton "&MB_EXPORT"
SAP_Session.FindById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell").SelectContextMenuItem "&XXL"
SAP_Session.FindById("wnd[1]/usr/ctxtDY_FILENAME").Text = "EXPORT5.MHTML"
SAP_Session.FindById("wnd[1]/usr/ctxtDY_FILENAME").CaretPosition = 0
SAP_Session.FindById("wnd[1]/tbar[0]/btn[0]").press
 
 

Dim Claims_Export_WB As Workbook
Sleep 1000

Set Claims_Export_WB = Workbooks.Open("C:\TEMP\Export5.MHTML")

Claims_Export_WB.Sheets(1).Copy After:=ORF_WB.Sheets("Macro Input")

ORF_WB.ActiveSheet.Name = ReconMonth & "_Claims Detail"


With ActiveWorkbook.ActiveSheet.Tab
    .Color = 192
    .TintAndShade = 0
End With

Dim Claims_DetailSheet As Worksheet
Set Claims_DetailSheet = ORF_WB.Sheets(ReconMonth & "_Claims Detail")



Sleep 1000
Claims_Export_WB.Close SaveChanges:=False
Sleep 1000
'========================================================================================================================
 
  
 
 
Claims_DetailSheet.Activate

'add formatting to claims detail exported sheet download
 Rows("1:1").RowHeight = 25.5
    
    
    Range("D1").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 5296274
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Range("L1").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 5296274
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Range("F1").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 5296274
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Range("H1").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 5296274
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Range("K17").Select
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With


Range("E1").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 5296274
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With



'Convert columns to numbers for VLOOKUPS
Claims_DetailSheet.Columns("H:H") = Claims_DetailSheet.Columns("H:H").Value
Claims_DetailSheet.Columns("F:F") = Claims_DetailSheet.Columns("F:F").Value
Claims_DetailSheet.Columns("L:L") = Claims_DetailSheet.Columns("L:L").Value


Continuewiththis = MsgBox(Prompt:="Delete exported .MHTML file in C:\TEMP?" & vbNewLine & vbNewLine & "Selecting 'No' can cause Excel to crash.", Buttons:=vbQuestion + vbYesNo)
If Continuewiththis = vbNo Then GoTo Skip_Deleting_Exports

        Kill ("C:\TEMP\EXPORT5.MHTML")

Skip_Deleting_Exports:          'Goes here and skips deleting export files if answer is no
 
 
 


MinutesElapsed = Format((Timer - StartTime) / 86400, "hh:mm:ss")

MsgBox "Ran successfully in " & MinutesElapsed & " minutes." & vbNewLine & vbNewLine & _
"The macro has finished adding the exported file to you current workbook." & vbNewLine & vbNewLine & _
"Please press OK, and then close the alert message that opens after the macro ends."


End Sub



