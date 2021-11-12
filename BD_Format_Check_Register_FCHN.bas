Attribute VB_Name = "BD_Format_Check_Register_FCHN"

Sub GL1130_FCHN_Format()

On Error GoTo Error

Dim ORF_WB As Workbook
Set ORF_WB = ThisWorkbook

Dim StartTime As Double
Dim MinutesElapsed As String
'Remember time when macro starts
StartTime = Timer


Run_FCHN = MsgBox(Prompt:="Do you want to format the FCHN tab? " & vbNewLine & vbNewLine & _
"Make sure that the FCHN tab is present and labeled by Macro #2.1 with the Recon Month Prefix." & vbNewLine & vbNewLine & _
"The macro will select the FCHN tab based on the Recon Month Prefix cell on the 'Macro Input' sheet, appended with the string '_FCHN YTD'.  This must match the name of the sheet.", Buttons:=vbQuestion + vbYesNo)

If Run_FCHN = vbNo Then
MsgBox ("Macro cancelled.")
End
End If




GLAccount = ORF_WB.Sheets("Macro Input").Range("GL_Account")
FiscalYear = ORF_WB.Sheets("Macro Input").Range("Fiscal_Year")
ReconMonth = ORF_WB.Sheets("Macro Input").Range("Recon_Month")
ReconMonth_Num = ORF_WB.Sheets("Macro Input").Range("ReconMonth_Num")

Application.ScreenUpdating = True
Application.DisplayAlerts = True


Dim FCHN As Worksheet
Set FCHN = ORF_WB.Sheets(ReconMonth & "_FCHN YTD")

FCHN.Activate
 
With FCHN

    .Rows("1:10").Delete Shift:=xlUp
    .Rows("1:1").Insert Shift:=xlDown
    .Rows("1:1").Insert Shift:=xlDown
    
    
End With


     '=============delete any random blank columns=============================================

    X = 3
    
    y = 3

    Do While y < 40
    
                row_value_1 = FCHN.Cells(3, X).Value
                row_value_2 = FCHN.Cells(4, X).Value
                
                        If row_value_1 = "" And row_value_2 = "" Then
                                    Columns(X).EntireColumn.Delete
                                    X = X - 1 'if row is deleted have to go back or will skip one
                        End If
            
            
                X = X + 1  'increment by one column
                y = y + 1 'independent counter

    Loop



  Columns("A:B").Select
    Selection.Delete Shift:=xlToLeft



    '=================================================================================
    With FCHN
 
    
    .Rows("3:3").Columns.AutoFit
    .Rows("4:4").Columns.AutoFit
    .Columns("F:F").EntireColumn.AutoFit
    .Columns("F:F").ColumnWidth = 10.57
    .Columns("T:T").EntireColumn.AutoFit
    .Columns("Y:Y").EntireColumn.AutoFit
    
    
    .Columns("S:S").ColumnWidth = 23.14
    .Columns("S:S").ColumnWidth = 12.29
    .Columns("R:R").EntireColumn.AutoFit
    .Columns("B:B").ColumnWidth = 5
    .Columns("C:C").ColumnWidth = 11.71
 
 End With
 

 
 With FCHN

 
        .Range("B3").FormulaR1C1 = "Itm"
        .Range("C3").FormulaR1C1 = "Pstng Date"
        .Range("D3").FormulaR1C1 = "Crcy"
        .Range("F3").FormulaR1C1 = "Amount in FC"
        .Range("I3").FormulaR1C1 = "Disc. Amount"
        .Range("K3").FormulaR1C1 = "Net Amount"
        .Range("M3").FormulaR1C1 = "Account No"
        .Range("N3").FormulaR1C1 = "Assignment"
        .Range("O3").FormulaR1C1 = "Text"
        .Range("R3").FormulaR1C1 = "Reference"
        .Range("S3").FormulaR1C1 = "Check Number"
        .Rows("4").Delete Shift:=xlUp

 
 End With
 
 
 
 
 
 
 '====================================================================================================================
 
 
 With FCHN
 
        .Columns("B:B").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        .Columns("A:A").EntireColumn.AutoFit
        .Range("B3").FormulaR1C1 = "DocumentNo"
        .Columns("B:B").ColumnWidth = 16.14

 End With
 
 
FCHN.DisplayPageBreaks = False
 
 

 
 
 '==================Format Header Starts Here============================================================================

 Range("A3:U3").Select

    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
              .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    
    
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.349986266670736
        .PatternTintAndShade = 0
    End With
    
    Selection.Font.Bold = True
    
    
    
    Rows("3:3").RowHeight = 28.5
    
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .AddIndent = False
        .ReadingOrder = xlContext
    End With
    

    
    Rows("3:3").RowHeight = 36.75
    
  

    Range("A2").Select
    ActiveCell.FormulaR1C1 = "1"
    Range("B2").Select
    ActiveCell.FormulaR1C1 = "2"
    Range("C2").Select
    ActiveCell.FormulaR1C1 = "3"
    Range("D2").Select
    ActiveCell.FormulaR1C1 = "4"
    Range("E2").Select
    ActiveCell.FormulaR1C1 = "5"
    Range("F2").Select
    ActiveCell.FormulaR1C1 = "6"
    Range("G2").Select
    ActiveCell.FormulaR1C1 = "7"
    Range("H2").Select
    ActiveCell.FormulaR1C1 = "8"
    Range("I2").Select
    ActiveCell.FormulaR1C1 = "9"
    Range("J2").Select
    ActiveCell.FormulaR1C1 = "10"
    Range("K2").Select
    ActiveCell.FormulaR1C1 = "11"
    Range("L2").Select
    ActiveCell.FormulaR1C1 = "12"
    Range("M2").Select
    ActiveCell.FormulaR1C1 = "13"
    Range("N2").Select
    ActiveCell.FormulaR1C1 = "14"
    Range("O2").Select
    ActiveCell.FormulaR1C1 = "15"
    Range("P2").Select
    ActiveCell.FormulaR1C1 = "16"
    Range("Q2").Select
    ActiveCell.FormulaR1C1 = "17"
    Range("R2").Select
    ActiveCell.FormulaR1C1 = "18"
    Range("S2").Select
    ActiveCell.FormulaR1C1 = "19"
    Range("T2").Select
    ActiveCell.FormulaR1C1 = "20"
    Range("U2").Select
    ActiveCell.FormulaR1C1 = "21"
'    Range("V2").Select
'    ActiveCell.FormulaR1C1 = "22"
'    Range("W2").Select
'    ActiveCell.FormulaR1C1 = "23"
'    Range("X2").Select
'    ActiveCell.FormulaR1C1 = "24"
'    Range("Y2").Select
'    ActiveCell.FormulaR1C1 = "25"
'    Range("Z2").Select
'    ActiveCell.FormulaR1C1 = "26"
'    Range("AA2").Select




    Range("A2:Z2").Select

    With Selection.Font
        .Color = -16776961
        .TintAndShade = 0
    End With

    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

Range("A1").Select

'==================Format Header Row Ends Here============================================================================


FCHN.Rows("4:4").Delete Shift:=xlUp

'==================Move DocumentNos from Column A to B starts here==========================================================

LastRow = FCHN.Cells.Find(What:="*", _
                            After:=FCHN.Range("A1"), _
                            LookAt:=xlPart, _
                            LookIn:=xlFormulas, _
                            SearchOrder:=xlByRows, _
                            SearchDirection:=xlPrevious, _
                            MatchCase:=False).row

   LastCol = FCHN.Cells.Find(What:="*", _
                            After:=FCHN.Range("A1"), _
                            LookAt:=xlPart, _
                            LookIn:=xlFormulas, _
                            SearchOrder:=xlByColumns, _
                            SearchDirection:=xlPrevious, _
                            MatchCase:=False).Column


'Index syntax is (row, column)

Dim DataRange
Set DataRange = FCHN.Range(Cells(3, 1), Cells(LastRow, LastCol))

FCHN.Range(Cells(4, 2), Cells(LastRow, 2)).Formula = "=A4"
FCHN.Range(Cells(4, 2), Cells(LastRow, 2)).Value = Range(Cells(4, 2), Cells(LastRow, 2)).Value


DataRange.AutoFilter Field:=2, Criteria1:="0"

FCHN.Range(Cells(4, 2), Cells(LastRow, 2)).SpecialCells(xlCellTypeVisible).ClearContents


FCHN.ShowAllData


'remove all blank rows
DataRange.AutoFilter Field:=1, Criteria1:="="

FCHN.Range(Cells(4, 1), Cells(LastRow, 1)).SpecialCells _
    (xlCellTypeVisible).EntireRow.Delete

FCHN.ShowAllData



'refresh lastrow and last column after deleting blank rows


LastRow = FCHN.Cells.Find(What:="*", _
                            After:=FCHN.Range("A1"), _
                            LookAt:=xlPart, _
                            LookIn:=xlFormulas, _
                            SearchOrder:=xlByRows, _
                            SearchDirection:=xlPrevious, _
                            MatchCase:=False).row

   LastCol = FCHN.Cells.Find(What:="*", _
                            After:=FCHN.Range("A1"), _
                            LookAt:=xlPart, _
                            LookIn:=xlFormulas, _
                            SearchOrder:=xlByColumns, _
                            SearchDirection:=xlPrevious, _
                            MatchCase:=False).Column


'Remove check numbers from col B





DataRange.AutoFilter Field:=4, Criteria1:="="
FCHN.Range(Cells(4, 2), Cells(LastRow, 2)).SpecialCells(xlCellTypeVisible).ClearContents
FCHN.ShowAllData



'insert formulas for all document #s in col A to refernce cell above it (which is check #)

DataRange.AutoFilter Field:=11, Criteria1:="="
DataRange.AutoFilter Field:=6, Criteria1:="="

FCHN.Range(Cells(4, 1), Cells(LastRow, 1)).SpecialCells(xlCellTypeVisible).FormulaR1C1 = "=R[-1]C[0]"
FCHN.ShowAllData
FCHN.Range(Cells(4, 1), Cells(LastRow, 1)).Value = Range(Cells(4, 1), Cells(LastRow, 1)).Value





'color code all the total lines

DataRange.AutoFilter Field:=2, Criteria1:="="
FCHN.Range(Cells(4, 1), Cells(LastRow, LastCol)).SpecialCells(xlCellTypeVisible).Select

With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent4
        .TintAndShade = 0.399975585192419
        .PatternTintAndShade = 0
    End With
    
    Range("A1").Select
    
    
    FCHN.ShowAllData



'==================Move DocumentNos from Column A to B ends here==========================================================

 Columns("J:J").ColumnWidth = 10.29
 Columns("F:F").ColumnWidth = 14.2
 Columns("P:P").ColumnWidth = 57
 Columns("E:E").Select
    With Selection
        .HorizontalAlignment = xlRight
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Columns("D:D").ColumnWidth = 15
    
    Columns("E:E").ColumnWidth = 5.29
    

    Columns("M:M").ColumnWidth = 26.57
    
    
    
    Range("A1").Select



MinutesElapsed = Format((Timer - StartTime) / 86400, "hh:mm:ss")
  
 MsgBox "Ran successfully in " & MinutesElapsed & " minutes." & vbNewLine & vbNewLine & _
"The macro has finished formatting the FCHN tab." & vbNewLine & vbNewLine & _
"Please press OK."



Exit Sub
Error:
 MsgBox "Something went wrong." & vbNewLine & vbNewLine & _
 "Make sure that the FCHN tab is present and labeled by Macro #2.1 with the Recon Month Prefix." & vbNewLine & vbNewLine & _
 "The macro will select the FCHN tab based on the Recon Month Prefix cell on the 'Macro Input' sheet, appended with the string '_FCHN YTD'.  This must match the name of the sheet.", vbExclamation

End Sub


