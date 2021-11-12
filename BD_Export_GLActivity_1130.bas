Attribute VB_Name = "BD_Export_GLActivity_1130"
Public BDPass As String
Public BDUserName As String

Sub GL1130_PullGL_Activity_Main_V1_LOG_IN_BOX()

On Error GoTo ErrSap


Dim ORF_WB As Workbook
Set ORF_WB = ThisWorkbook

BD_LOG_ON.Show
SapConnectionString = ORF_WB.Sheets("Macro Input").Range("SAP_Connection")

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

Sleep 1
SAP_Session.FindById("wnd[0]/tbar[0]/okcd").Text = "/nFAGLB03 "
Sleep 1
SAP_Session.FindById("wnd[0]").sendVKey 0
Sleep 1




Dim GLAccount As Long, FiscalYear As Long, ReconMonth As Variant, ReconMonth_Num As Long

GLAccount = ORF_WB.Sheets("Macro Input").Range("GL_Account")
FiscalYear = ORF_WB.Sheets("Macro Input").Range("Fiscal_Year")
ReconMonth = ORF_WB.Sheets("Macro Input").Range("Recon_Month")
ReconMonth_Num = ORF_WB.Sheets("Macro Input").Range("ReconMonth_Num")
Sleep 1


Sleep 1
SAP_Session.FindById("wnd[0]/usr/ctxtRACCT-LOW").Text = GLAccount
Sleep 1
SAP_Session.FindById("wnd[0]/usr/txtRYEAR").Text = FiscalYear
Sleep 1
SAP_Session.FindById("wnd[0]/tbar[1]/btn[8]").press
Sleep 1


Sleep 2000
Call PrintScreen



SAP_Session.FindById("wnd[0]/usr/cntlFDBL_BALANCE_CONTAINER/shellcont/shell").PressToolbarContextButton "&MB_EXPORT"
Sleep 1
SAP_Session.FindById("wnd[0]/usr/cntlFDBL_BALANCE_CONTAINER/shellcont/shell").SelectContextMenuItem "&XXL"
Sleep 1
SAP_Session.FindById("wnd[1]/usr/ctxtDY_PATH").Text = "C:\\TEMP"
Sleep 1
SAP_Session.FindById("wnd[1]/usr/ctxtDY_FILENAME").Text = "EXPORT.MHTML"
Sleep 1
SAP_Session.FindById("wnd[1]/tbar[0]/btn[11]").press
Sleep 1


'Wait manually for 4 seconds until Excel Export of GL Balances generates
Sleep 4000


Dim GL_Export_WB As Workbook
Sleep 1000

Set GL_Export_WB = Workbooks.Open("C:\TEMP\EXPORT.MHTML")

GL_Export_WB.Sheets(1).Copy After:=ORF_WB.Sheets("Macro Input")

ORF_WB.ActiveSheet.Name = ReconMonth & "_GL 1130 Bal"
With ActiveWorkbook.ActiveSheet.Tab
    .Color = 192
    .TintAndShade = 0
End With


Dim GLBal As Worksheet
Set GLBal = ORF_WB.ActiveSheet


Sleep 1000
GL_Export_WB.Close SaveChanges:=False
Sleep 1000

'paste screenshot of GL balance============================
ActiveSheet.Paste Destination:=GLBal.Range("H3")

Dim shp As Shape
Dim h As Single, w As Single, l As Single, r As Single


Call PictureFormat
'========================================================


ORF_WB.Sheets("Macro Input").Activate


ActivateWindow (SessionHWND)
SAP_Session.FindById("wnd[0]").Maximize

 
SAP_Session.FindById("wnd[0]/usr/cntlFDBL_BALANCE_CONTAINER/shellcont/shell").SetCurrentCell ReconMonth_Num, "BALANCE"
SAP_Session.FindById("wnd[0]/usr/cntlFDBL_BALANCE_CONTAINER/shellcont/shell").DoubleClickCurrentCell
 
 
 
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
 
 
 
 
 'Export the file to format that can be opened by Excel
 
SAP_Session.FindById("wnd[0]/mbar/menu[0]/menu[3]/menu[1]").Select
SAP_Session.FindById("wnd[1]/usr/ctxtDY_PATH").Text = "C:\\TEMP"
SAP_Session.FindById("wnd[1]/usr/ctxtDY_FILENAME").Text = "EXPORT2.MHTML"
SAP_Session.FindById("wnd[1]/tbar[0]/btn[11]").press
Sleep 5000
  
  
  
 Set GL_Export_WB2 = Workbooks.Open("C:\TEMP\EXPORT2.MHTML")

GL_Export_WB2.Sheets(1).Copy After:=ORF_WB.Sheets("Macro Input")

ORF_WB.ActiveSheet.Name = ReconMonth & "_GL 1130 Detail"
With ActiveWorkbook.ActiveSheet.Tab
    .Color = 192
    .TintAndShade = 0
End With

Sleep 1
GL_Export_WB2.Close SaveChanges:=False
Sleep 1000
 
 
' DeActivateWindow (SessionHWND)
ActivateWindow (Application.hwnd)    'Brings the Excel window to the front
SAP_Session.FindById("wnd[0]").Maximize
 

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
 
 
GLBal.Activate
Call AddChecktoBal
    
    
    
Sleep 2000
GLBal.Range("A1").Select
GL2.Activate
Sleep 1



lastrow2 = ActiveSheet.Cells(ActiveSheet.Rows.Count, "F").End(xlUp).row

Dim c As Range
Dim cell As Range
Set c = GL2.Range("A2:A" & lastrow2)


Sleep 1000

For Each cell In c

If IsEmpty(cell.Value) Then
    cell.EntireRow.Delete
End If

Next cell
 
 
 
 
ORF_WB.Sheets("Macro Input").Activate
ORF_WB.Sheets("Macro Input").Range("A1").Select
Sleep 1





Set SAP_Application = Nothing

            Continuewiththis = MsgBox(Prompt:="Delete exported .MHTML file in C:\TEMP?" & vbNewLine & vbNewLine & "Selecting 'No' can cause Excel to crash.", Buttons:=vbQuestion + vbYesNo)
            
            If Continuewiththis = vbNo Then GoTo Skip_Deleting_Exports
            
                                Kill ("C:\TEMP\EXPORT.MHTML")
                                Kill ("C:\TEMP\EXPORT2.MHTML")
            
Skip_Deleting_Exports:         'Goes here and skips deleting export files if answer is no
 

 Set SAP_Application = Nothing

 
 
 
 
 
MinutesElapsed = Format((Timer - StartTime) / 86400, "hh:mm:ss")
 
 
 MsgBox "Ran successfully in " & MinutesElapsed & " minutes." & vbNewLine & vbNewLine & _
"The macro has finished adding the exported files to you current workbook." & vbNewLine & vbNewLine & _
"Please press OK, and then close the two alert messages that open after the macro ends."



Exit Sub
ErrSap:
MsgBox "Error.  Please press OK to end the macro." & vbNewLine & vbNewLine & "Make sure that sheets ReconMonth_GL 1130 Detail and ReconMonth_GL_GL 1130 Detail do not already exist from a previous run, and that input fields on the Macro Input sheet are filled in correctly."





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
