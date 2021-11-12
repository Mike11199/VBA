Attribute VB_Name = "BD_Pull_GL_BALANCE_ALL"
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



Sub Pull_All_GL_Balance()

Application.DisplayAlerts = False

On Error GoTo ErrSap

Set Recon_WB = ThisWorkbook

Dim StartTime As Double
Dim MinutesElapsed As String
'Remember time when macro starts
StartTime = Timer


BD_LOG_ON.Show
SapConnectionString = Recon_WB.Sheets("Macro Input").Range("SAP_Connection")


Dim shp As Shape
Dim h As Single, w As Single, l As Single, R As Single
    
FiscalYear = Recon_WB.Sheets("Macro Input").Range("Fiscal_Year")
ReconMonth = Recon_WB.Sheets("Macro Input").Range("Recon_Month")
ReconMonth_Num = Recon_WB.Sheets("Macro Input").Range("ReconMonth_Num")


GL_Balance_Array = Recon_WB.Sheets("Macro Input").Range("GL_Balance")
GL_Activity_Array = Recon_WB.Sheets("Macro Input").Range("GL_Activity")

For GLCount = 1 To UBound(GL_Balance_Array, 1)
            Debug.Print GL_Balance_Array(GLCount, 1)        'where 1 represents your second dimension (ie column)
Next GLCount



'=======LOG IN IF THE FIRST GL BALANCE IS NOT EMPTY============================================
If Not IsEmpty(GL_Balance_Array(1, 1)) Then


            Set SAP_Application = CreateObject("Sapgui.ScriptingCtrl.1")
            
            Set Connection = SAP_Application.OpenConnection(SapConnectionString, True)
            Set SAP_Session = Connection.Children(0)
            
            
            SessionHWND = SAP_Session.FindById("wnd[0]").Handle
            ActivateWindow (SessionHWND)
            SAP_Session.FindById("wnd[0]").Maximize

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
                
End If
'=======LOG IN ENDS HERE=============================================================
             
             
             
'=======LOOP THROUGH ALL GL ACCOUNTS STARTS=======================================
             
             
For GLCount = 1 To UBound(GL_Balance_Array, 1)
    If Not IsEmpty(GL_Balance_Array(GLCount, 1)) Then

                            Sleep 2
                            SAP_Session.FindById("wnd[0]/tbar[0]/okcd").Text = "/nFAGLB03 "
                            Sleep 2
                            SAP_Session.FindById("wnd[0]").sendVKey 0
                            Sleep 2
                                
                
                            ActivateWindow (SessionHWND)
                            SAP_Session.FindById("wnd[0]").Maximize
                            
                            SAP_Session.FindById("wnd[0]/usr/ctxtRACCT-LOW").Text = GL_Balance_Array(GLCount, 1)
                            SAP_Session.FindById("wnd[0]/usr/txtRYEAR").Text = FiscalYear
                            SAP_Session.FindById("wnd[0]/tbar[1]/btn[8]").press
                            
                            Sleep 2000
                            Call PrintScreen
                                
                                                    '==================Error handling if no GL Activity=========================================================================================================
                                                     WindowName = SAP_Session.ActiveWindow.Text
                                                    
                                                    If WindowName = "Information" Then
                                                                Call No_GL_Balance_Information
                                                                LoopNumber = LoopNumber + 1
                                                                
                                                               LastRow = Find_Last_Row()
                                                               LastCol = Find_Last_Column()

                                                                GoTo NextGL
                                                    End If
                                                    '===================End Error handling================================================================================================================
                                
                
                                SAP_Session.FindById("wnd[0]/usr/cntlFDBL_BALANCE_CONTAINER/shellcont/shell").PressToolbarContextButton "&MB_EXPORT"
                                SAP_Session.FindById("wnd[0]/usr/cntlFDBL_BALANCE_CONTAINER/shellcont/shell").SelectContextMenuItem "&XXL"
                                SAP_Session.FindById("wnd[1]/usr/ctxtDY_PATH").Text = "C:\\TEMP"
                                SAP_Session.FindById("wnd[1]/usr/ctxtDY_FILENAME").Text = "EXPORT.MHTML"
                                SAP_Session.FindById("wnd[1]/tbar[0]/btn[11]").press

                                Sleep 2
                                ActivateWindow (Application.hwnd) 'ExcelWBInFront
                                
                                
                                Set GL_Export_WB = Workbooks.Open("C:\TEMP\EXPORT.MHTML")
                                
                                
                                If First_GL_Exported <> 1 Then
                                    Call RunFirstGL
                                    First_GL_Exported = 1
                                Else
                                    Call RunNextGL
                                End If
                                
                                Kill ("C:\TEMP\EXPORT.MHTML")
                                Call CloseOtherWorkbook
                                
                                LoopNumber = LoopNumber + 1
                          


End If


NextGL:
                
Next GLCount
                
  
Set SAP_Application = Nothing
                
'Continuewiththis = MsgBox(Prompt:="Delete exported .MHTML file in C:/TEMP? (selecting no can cause Excel to crash)", Buttons:=vbQuestion + vbYesNo)
'
'If Continuewiththis = vbYes Then
'    Kill ("C:\TEMP\EXPORT.MHTML")
'End If

                
Rows("1:1").Select
Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
Range("A1:F1").Select
With Selection.Interior
.Pattern = xlSolid
.PatternColorIndex = xlAutomatic
.ThemeColor = xlThemeColorLight1
.TintAndShade = 0
.PatternTintAndShade = 0
End With


Columns("C:C").EntireColumn.AutoFit
Columns("D:D").EntireColumn.AutoFit

                
                
FormatAddBlackSpaceforScreenshots = MsgBox(Prompt:="Do you want to add rows to align screenshots with GL Balances?  This might only work well for users with a certain monitor size or aspect ratio.", Buttons:=vbQuestion + vbYesNo)

If FormatAddBlackSpaceforScreenshots = vbYes Then
    Call AdvancedFormattingAlignScreenshots
End If
                

With ActiveWorkbook.ActiveSheet.Tab
        .Color = 192
        .TintAndShade = 0
End With

MinutesElapsed = Format((Timer - StartTime) / 86400, "hh:mm:ss")
 
 
 MsgBox "Ran successfully in " & MinutesElapsed & " minutes." & vbNewLine & vbNewLine & _
"The macro has finished adding the exported files to you current workbook." & vbNewLine & vbNewLine & _
"Please press OK, and then close the two alert messages that open after the macro ends."

Application.DisplayAlerts = True

Exit Sub
ErrSap:
MsgBox "Error.  Please press OK to end the macro.  Verify that input cells are updated for the current month."
Application.DisplayAlerts = True



End Sub


Sub AddChecktoBal()

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
   

 Columns("A:A").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A1").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0.399975585192419
        .PatternTintAndShade = 0
    End With
    With Selection.Font
        .Color = -16776961
        .TintAndShade = 0
    End With
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "GL Account"
    Range("A1").Select
    Selection.Font.Bold = True
    


End Sub


Sub addglaccountcell()

ActiveCell.FormulaR1C1 = "GL Account"
    With Selection.Font
        .Color = -16776961
        .TintAndShade = 0
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
    Selection.Font.Bold = True
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0.399975585192419
        .PatternTintAndShade = 0
    End With
    
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With





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
R = -(Crop_Right - shp.Width)
' the new size ratio of our WHOLE screenshot pasted (with keeping aspect ratio)
'.Height = 1260
'.Width = 1680
.LockAspectRatio = False
End With

With shp.PictureFormat
.CropRight = R
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


Function RangeExists(R As String) As Boolean
    Dim Test As Range
    On Error Resume Next
    Set Test = ActiveSheet.Range(R)
    RangeExists = Err.Number = 0
End Function
Sub No_GL_Balance_Information()

        SAP_Session.FindById("wnd[1]/tbar[0]/btn[0]").press
        
        If First_GL_Exported <> 1 Then
        
                Sheets.Add After:=ActiveSheet
                Recon_WB.ActiveSheet.Name = ReconMonth & "_All GL Bal"
                Set GLBal = Recon_WB.ActiveSheet
                GLBal.Activate
        
        End If
        
        GLBal.Activate
        

        'paste the screenshot==========================================================================
        ActiveSheet.Paste Destination:=GLBal.Range("H" & 43 + screenshotrows)
        screenshotrows = screenshotrows + 40

        
        GLBal.Range("A" & LastRow + 3 & ":A" & LastRow + 20).Formula2R1C1 = GL_Balance_Array(GLCount, 1)
        
        GLBal.Range("F" & LastRow + 3 & ":F" & LastRow + 20).Formula2R1C1 = "No Balance"
        
        GLBal.Range("A" & LastRow + 1 & ":F" & LastRow + 1).Select
        
        With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .PatternTintAndShade = 0
        End With
        
        LastRow = ActiveSheet.Cells.Find(What:="*", _
        After:=ActiveSheet.Range("A1"), _
        LookAt:=xlPart, _
        LookIn:=xlFormulas, _
        SearchOrder:=xlByRows, _
        SearchDirection:=xlPrevious, _
        MatchCase:=False).Row
        
        LastCol = ActiveSheet.Cells.Find(What:="*", _
        After:=ActiveSheet.Range("A1"), _
        LookAt:=xlPart, _
        LookIn:=xlFormulas, _
        SearchOrder:=xlByColumns, _
        SearchDirection:=xlPrevious, _
        MatchCase:=False).Column


End Sub

Function Find_Last_Row()

Find_Last_Row = ActiveSheet.Cells.Find(What:="*", _
        After:=ActiveSheet.Range("A1"), _
        LookAt:=xlPart, _
        LookIn:=xlFormulas, _
        SearchOrder:=xlByRows, _
        SearchDirection:=xlPrevious, _
        MatchCase:=False).Row
        

End Function

Function Find_Last_Column()

Find_Last_Column = ActiveSheet.Cells.Find(What:="*", _
        After:=ActiveSheet.Range("A1"), _
        LookAt:=xlPart, _
        LookIn:=xlFormulas, _
        SearchOrder:=xlByColumns, _
        SearchDirection:=xlPrevious, _
        MatchCase:=False).Column

End Function

Sub RunNextGL()

Set GL_Export_WB = Workbooks.Open("C:\TEMP\EXPORT.MHTML")

LastRow2 = Find_Last_Row
lastcol2 = Find_Last_Column

GLBal.Activate

'paste the screenshot===========
ActiveSheet.Paste Destination:=GLBal.Range("H" & 43 + screenshotrows)
screenshotrows = screenshotrows + 40
Call PictureFormat

Set GL_Export_Sheet = GL_Export_WB.Sheets(1)

GL_Export_Sheet.Activate
GL_Export_Sheet.Range(Cells(1, 1), Cells(LastRow2, lastcol2)).Copy GLBal.Range("B" & LastRow + 2)


GLBal.Activate
GLBal.Range("A" & LastRow + 3 & ":A" & LastRow + 20).Formula2R1C1 = GL_Balance_Array(GLCount, 1)


GLBal.Range("A" & LastRow + 1 & ":F" & LastRow + 1).Select
With Selection.Interior
.Pattern = xlSolid
.PatternColorIndex = xlAutomatic
.ThemeColor = xlThemeColorLight1
.TintAndShade = 0
.PatternTintAndShade = 0
End With


GLBal.Range("A" & LastRow + 2).Select
Call addglaccountcell

GLBal.Activate

If ReconMonth_Num < 10 Then
strSearch2 = "00" & ReconMonth_Num
Else
strSearch2 = "0" & ReconMonth_Num
End If

Set rng2 = Range("B" & LastRow + 4 & ":B" & LastRow + 20).Find(strSearch2, , xlValues, xlWhole)

rng2.Select
rng2.Resize(, 5).Select

With Selection.Interior
.Pattern = xlSolid
.PatternColorIndex = xlAutomatic
.Color = 49407
.TintAndShade = 0
.PatternTintAndShade = 0
End With


LastRow = Find_Last_Row
LastCol = Find_Last_Column


Sleep 1000
GL_Export_WB.Close SaveChanges:=False
Sleep 1000



End Sub
Sub RunFirstGL()
      
GL_Export_WB.Sheets(1).Copy After:=Recon_WB.Sheets("Macro Input")
Recon_WB.ActiveSheet.Name = ReconMonth & "_All GL Bal"
Set GLBal = Recon_WB.ActiveSheet

Sleep 2
GL_Export_WB.Close SaveChanges:=False
Sleep 2


GLBal.Activate
Call AddChecktoBal


LastRow = Find_Last_Row
LastCol = Find_Last_Column


Recon_WB.ActiveSheet.Range("A2:A" & LastRow).FormulaR1C1 = GL_Balance_Array(GLCount, 1)


GLBal.Range("A1").Select


ActiveSheet.Paste Destination:=GLBal.Range("H3")

Call PictureFormat
'====================================


End Sub

Sub AdvancedFormattingAlignScreenshots()


    LastRow = Find_Last_Row
    LastCol = Find_Last_Column
    
    PasteLocation = (LoopNumber * 40) - 36
    FirstRow = LastRow - 19



    For i = 1 To LoopNumber
    
            GLBal.Range("A" & FirstRow & ":F" & LastRow).Select
            Selection.Cut
            Range("A" & PasteLocation).Select
            ActiveSheet.Paste
            
            PasteLocation = PasteLocation - 40
            
            
            FirstRow = FirstRow - 20
            If FirstRow <= 0 Then
                FirstRow = 1
            End If
            LastRow = LastRow - 20
        

    Next i




End Sub
 Sub CloseOtherWorkbook()
'Update 20141126
Dim xWB As Workbook
Application.ScreenUpdating = False
For Each xWB In Application.Workbooks
    If Not (xWB Is Application.ThisWorkbook) Then
        xWB.Close
    End If
Next
Application.ScreenUpdating = True
End Sub
