Attribute VB_Name = "Add_CalATERS_Sheets"
Sub GrabData_CalATERS_v2()

On Error GoTo Error

Dim ORF_WB As Workbook
Set ORF_WB = ThisWorkbook

Dim StartTime As Double
Dim MinutesElapsed As String
'Remember time when macro starts
StartTime = Timer


Dim strPath As String
strFileName = Environ("USERPROFILE") & "\OneDrive - California State Teachers' Retirement System\Desktop"


ChDrive Left(strFileName, 1)
ChDir strFileName

MsgBox "Select all Excel files in the relevant folder by holding CTRL, or clicking and dragging the mouse.  Then click 'Open' to add and merge files." & vbNewLine & vbNewLine & _
"This should be a folder on your desktop with all the CalATERS files downloaded from SharePoint."
    
        
Call MergeExcelFiles
Call Merge_Sheets
Call Add_Count_Cells



MinutesElapsed = Format((Timer - StartTime) / 86400, "hh:mm:ss")

MsgBox "Ran successfully in " & MinutesElapsed & " minutes." & vbNewLine & vbNewLine & _
"The Macro has finished running. Please click OK. " & vbNewLine & vbNewLine & _
"Check that all data was copied to the 'CalATERS Info' tab for the current month. " & vbNewLine & vbNewLine & _
"If everything looks correct, manually delete all other tabs besides the 'CalATERS Info' master sheet, between the  'CALATERS -->'  and '<-- CALATERS' index tabs."




Exit Sub
Error:  MsgBox "Something went wrong.  Please try again."


End Sub

Sub Merge_Sheets()

'This will loop through each sheet and select specific columns from each sheet.  Each column will be select by name into a new sheet.  If names change even by one letter this can break
'However, the advantage of this is that changes in column order will NOT cause this to break.  If names change, edit the names
'Then another macro can use the new sheet added to grab data with VLOOKUPS

Dim ORF_WB As Workbook
Set ORF_WB = ThisWorkbook

'Sets ReconMonth name as variable from macro input sheet
Dim ReconMonth As Variant
ReconMonth = ORF_WB.Sheets("Macro Input").Range("Recon_Month")


'Creates a master sheet and color codes it
ORF_WB.Sheets.Add(After:=Sheets("CALATERS -->")).Name = ReconMonth & "_CalATERS Info"
Dim mstr As Worksheet
Set mstr = ORF_WB.Sheets(ReconMonth & "_CalATERS Info")

    With mstr.Tab
        .Color = 192
        .TintAndShade = 0
    End With


'Sets up the number of sheet to the right of the master sheet
Dim aftersheetnumber As Long
aftersheetnumber = mstr.Index + 1



'Activates sheet right of master sheet
ORF_WB.Sheets(aftersheetnumber).Activate


Dim Rng As Range
Application.DisplayAlerts = False





'Sets up the first loop.  This will loop through the first sheet after the master sheet to the last '<-- CALATERS" sheet'"
    For J = aftersheetnumber To ORF_WB.Sheets("<-- CALATERS").Index - 1

        
                ORF_WB.Sheets(J).Activate
                
                'If filters remove them
                If ORF_WB.Sheets(J).AutoFilterMode Then
                    ORF_WB.Sheets(J).AutoFilterMode = False
                End If
    
            
                'LastRow is the last in sheet from bottom ctrl shift up
                LastRow = Cells(Rows.Count, 1).End(xlUp).row
                HeaderRow = Cells(LastRow, 1).End(xlUp).row
                LastCol = 40
            
            
                ' If the first time looping, set last row to 1 so will paste to mstr sheet A1
                If J = aftersheetnumber Then
                       lastrow3 = 1
                End If

    
                'IMPORTANT-- IF ROWS NAMES CHANGE, YOU NEED TO CHANGE THESE VALUES.  MACRO WILL BREAK IF EVEN ONE LETTER DOES NOT MATCH
                MyArr = Array("ORF check #", "Amount", "Vendor #", "Vendor Name", "Trip ID", "GER #", "GER Amount")



                                'This is the second 'for loop'.  The first 'for loop' goes through every sheet.  This one goes through every name in the array above.  So for each sheet, it will grab columns matching each name in the array.
                                For i = LBound(MyArr) To UBound(MyArr)
                                
                                                ORF_WB.Sheets(J).Activate
                                                
                                                Set Rng = ActiveSheet.Cells.Find(What:=MyArr(i), _
                                                    LookIn:=xlFormulas, _
                                                    LookAt:=xlWhole, _
                                                    SearchOrder:=xlByRows, _
                                                    SearchDirection:=xlNext, _
                                                    MatchCase:=False)
                                                    
                                                If Not Rng Is Nothing Then
                                                FirstAddress = Rng.Address
                                                ' Do
                                                Rcount = Rcount + 1
                                                calaters_column = Rng.Column
                                                calaters_row = Rng.row
                                                k = i + 1
                                                
                                                ActiveSheet.Range(Cells(calaters_row, calaters_column), Cells(LastRow, calaters_column)).Select
                                                ActiveSheet.Range(Cells(calaters_row, calaters_column), Cells(LastRow, calaters_column)).Copy 'Destination:=mstr.Range(Cells(LastRow3, k))
                                                mstr.Activate
                                                ActiveSheet.Range(Cells(lastrow3, k), Cells(lastrow3, k)).Select
                                                Selection.PasteSpecial xlPasteAll
                                                Selection.PasteSpecial xlPasteValues
                                                End If
                                
                                Next i

                                                  lastrow3 = Cells(Rows.Count, 1).End(xlUp).row
                                                  lastrow3 = lastrow3 + 1
         
    Next J



'This just turns alerts back on and resizes a few columns on the master sheet to be wider
Application.DisplayAlerts = True
Application.CutCopyMode = False
mstr.Activate
Cells.Select
Selection.Columns.AutoFit
Columns("A:A").Value = Columns("A:A").Value
Cells.Select
Selection.RowHeight = 12.75
mstr.Range("A1").Select
Columns("E:E").ColumnWidth = 16
Columns("B:B").ColumnWidth = 8
Columns("A:A").ColumnWidth = 12
Call FormatTopRow



'MsgBox ("The macro has finished running.  Please click OK.")


End Sub

Sub MergeExcelFiles()
    Dim fnameList, fnameCurFile As Variant
    Dim countFiles, countSheets As Integer
    Dim wksCurSheet As Worksheet
    Dim wbkCurBook, wbkSrcBook As Workbook
 
    fnameList = Application.GetOpenFilename(FileFilter:="Microsoft Excel Workbooks (*.xls;*.xlsx;*.xlsm),*.xls;*.xlsx;*.xlsm", Title:="Choose Excel files to merge", MultiSelect:=True)
 
    If (vbBoolean <> VarType(fnameList)) Then
 
        If (UBound(fnameList) > 0) Then
            countFiles = 0
            countSheets = 0
 
            Application.ScreenUpdating = False
            Application.Calculation = xlCalculationManual
 
            Set wbkCurBook = ActiveWorkbook
 
            For Each fnameCurFile In fnameList
                countFiles = countFiles + 1
 
 
   'G:\Disbursements Units\Admin Accounts Payable\ORF checks\nfchn\20-21
 
    'ChDrive "G"

    'ChDir "G:\Disbursements Units\Admin Accounts Payable\ORF checks\nfchn\20-21"
 
 
                Set wbkSrcBook = Workbooks.Open(FileName:=fnameCurFile)
 
                For Each wksCurSheet In wbkSrcBook.Sheets
                
                
                'IMPORTANT.  ONLY SELECTS SHEETS WITH THE NAME 'Work pool' EXACTLY.  MAY BREAK HERE AND NEED TO CHANGE SHEET NAME
                    If wksCurSheet.Name = "Work pool" Then
                    countSheets = countSheets + 1
                    wksCurSheet.Copy After:=wbkCurBook.Sheets("CALATERS -->")
                    wbname = wbkSrcBook.Name
                     wbname2 = Left(wbname, Len(wbname) - 5)
                    ActiveSheet.Name = wbname2
                    
                    'hard code all formulas
                    With ActiveSheet.UsedRange
                     .Value = .Value
                    End With
           
                    End If
                Next
 
                wbkSrcBook.Close SaveChanges:=False
 
            Next
 
            Application.ScreenUpdating = True
            Application.Calculation = xlCalculationAutomatic
 
           ' MsgBox "Processed " & countFiles & " files" & vbCrLf & "Merged " & countSheets & " worksheets", Title:="Merge Excel files"
        End If
 
    Else
                MsgBox "No files selected.  Exiting macro.", Title:="Merge Excel files"
                End
    End If
End Sub

Sub FormatTopRow()
'
' Macro66 Macro
'

'
    Range("A1:G1").Select
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
    Range("A2").Select
    ActiveWindow.SmallScroll Down:=-6
    ActiveWindow.FreezePanes = True
    Range("A1:G1").Select
    Selection.Font.Bold = False
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Font.Bold = True
    Range("A1").Select
End Sub

Sub Add_Count_Cells()
'
' Macro66 Macro
'

LastRow = Cells(Rows.Count, 1).End(xlUp).row
'
    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "Count"
    Range("A1").Select
    With Selection.Font
        .Color = -16776961
        .TintAndShade = 0
    End With
    Selection.Font.Bold = True
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0.599993896298105
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
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlTop
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Range("A2:A" & LastRow).FormulaR1C1 = "=COUNTIF(R2C7:RC7,RC[6])"

Application.CutCopyMode = False
End Sub
