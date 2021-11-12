Attribute VB_Name = "Excel_Roll_Over_GL_Detail"
Sub RollOverGLDetail_AddFormulas()



Dim ORF_WB As Workbook
Set ORF_WB = ThisWorkbook

GLAccount = ORF_WB.Sheets("Macro Input").Range("GL_Account")
FiscalYear = ORF_WB.Sheets("Macro Input").Range("Fiscal_Year")
ReconMonth = ORF_WB.Sheets("Macro Input").Range("Recon_Month")
ReconMonth_Num = ORF_WB.Sheets("Macro Input").Range("ReconMonth_Num")


Dim StartTime As Double
Dim MinutesElapsed As String
'Remember time when macro starts
StartTime = Timer


Roll_Over_Formulas = MsgBox(Prompt:="Do you want to add the current month GL detail to the recon face sheet, and add all formulas? " & vbNewLine & vbNewLine & _
"Macros #1-4 should be run before this, or SUMIFS, VLOOKUPS to sheets added by these previous macros will error out, and cause this macro to break." & vbNewLine & vbNewLine & _
"This is because Excel will not know what sheet the VLOOKUPS are referencing as they will be missing from Macros #1-4,, and will open a dialog box to select an Excel file, interrupting the macro." & vbNewLine & vbNewLine & _
"If issues occurred with macro #4, you can add a sheet manually titled '_ORF Claim Info' prefixed by the Recon Month Prefix on the 'Macro Input' sheet.", Buttons:=vbQuestion + vbYesNo)

If Roll_Over_Formulas = vbNo Then
MsgBox ("Macro cancelled.")
End
End If



Dim ReconSheet As Worksheet
Set ReconSheet = ORF_WB.Sheets("1130_" & ReconMonth)

ReconSheet.Activate

'Finds last row on recon face sheet by going to bottom of column A and ctrl+shit+up.  Adding another column before A or some value in column A after GL detail will break this macro
lastrow2 = ActiveSheet.Cells(ActiveSheet.Rows.Count, "A").End(xlUp).row


 Dim GL2 As Worksheet
 Set GL2 = ORF_WB.Sheets(ReconMonth & "_GL 1130 Detail")

Debug.Print lastrow2


'Grabs last row and column from the GL Detail sheet
GL2.Activate



LastRow = ActiveSheet.Cells.Find(What:="*", _
                            After:=ActiveSheet.Range("A1"), _
                            LookAt:=xlPart, _
                            LookIn:=xlFormulas, _
                            SearchOrder:=xlByRows, _
                            SearchDirection:=xlPrevious, _
                            MatchCase:=False).row


   LastCol = ActiveSheet.Cells.Find(What:="*", _
                            After:=ActiveSheet.Range("A1"), _
                            LookAt:=xlPart, _
                            LookIn:=xlFormulas, _
                            SearchOrder:=xlByColumns, _
                            SearchDirection:=xlPrevious, _
                            MatchCase:=False).Column



Dim copyRng As Range
Set copyRng = GL2.Range("A2:Q" & LastRow)

Dim copyRng2 As Range
Set copyRng2 = GL2.Range("R2:R" & LastRow)

'Copies GL detail to bottom of the recon face sheet
ReconSheet.Activate

ReconSheet.Range("A" & lastrow2 + 1 & ":S" & copyRng.Rows.Count + lastrow2).Insert Shift:=xlDown
copyRng.Copy ReconSheet.Range("C" & lastrow2 + 1)



'Sets up range variable for current month GL detail on face sheet and color codes it purple
Dim CurrentMonthReconItems As Range

Set CurrentMonthReconItems = ReconSheet.Range("A" & lastrow2 + 1 & ":S" & copyRng.Rows.Count + lastrow2)

CurrentMonthReconItems.Select

  With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent4
        .TintAndShade = 0.599993896298105
        .PatternTintAndShade = 0
    End With
    
    
'labels all current month items as CM in column A
CurrentMonthReconItems.Columns(1).FormulaR1C1 = "CM"
CurrentMonthReconItems.Columns(12).Value = copyRng2.Value
    
    
'expands range to all columns to the right
Set CurrentMonthReconItems = CurrentMonthReconItems.Resize(, 29)



'This adds a COUNTIF formula on the recon face sheet which looks at the claims detail sheet. It sees how many rows exist on the claims detail for each claim on the recon face sheet listed (these lines on the face sheet are from the GL detail).
CurrentMonthReconItems.Select

CurrentMonthReconItems.Columns(2).FormulaR1C1 = _
        "=COUNTIF('" & ReconMonth & "_Claims Detail'!C4,'1130_" & ReconMonth & "'!RC[10])"





'Sets up another range variable c2 as column B, the range which contains the COUNTIF values, and other range variables
Dim c As Range

Dim c2 As Range
Set c2 = CurrentMonthReconItems.Columns(2)
c2.Select

'Sets up another range variable c3 - not used
Dim c3 As Range
Set c3 = CurrentMonthReconItems.Columns(11)


LoopEnd_For_AddingRows = c2.Rows.Count + lastrow2




  '============LOOP THAT ADDS ROWS AND SUMIFS FORMULAS DEPENDING ON VALUES IN COUNTIF FORMULAS=========================================
colNr = 2

    For rowNr = 1 To LoopEnd_For_AddingRows
        'For colNr = c2.Column To c2.Columns.Count
            Set cell = Cells(rowNr, colNr)
            
                            If cell.Value <> "" Then
                                     For COUNTIFNUMBER = 2 To 50
                    
                                            If cell.Value = COUNTIFNUMBER Then
                                                Call Add_Rows(rowNr, colNr, COUNTIFNUMBER)
                                                   rowNr = rowNr + (COUNTIFNUMBER - 1)
                                            End If
                                    
                                    Next COUNTIFNUMBER
                            End If
                           

        'Next colNr
    Next rowNr

'=======================END LOOP===========================================================================================================






'Adds formulas for various columns for the claim lines.  Only the claim line will have COUNTIFS in column B, hence, if c.value >0
For Each c In c2.Cells

                If c.Value > 0 Then
            
            
                            'Adds SUMIFS to grab amount for each claim
                            c.Offset(, 6).FormulaR1C1 = _
                            "=SUMIFS('" & ReconMonth & "_Claims Detail'!C5,'" & ReconMonth & "_Claims Detail'!C12,'1130_" & ReconMonth & "'!RC[-6],'" & ReconMonth & "_Claims Detail'!C4,'1130_" & ReconMonth & "'!RC[4])"
                            c.Offset(, 6).Interior.ThemeColor = xlThemeColorAccent4
                            c.Offset(, 6).Interior.TintAndShade = 0.399975585192419
                            ' c.Offset(, 6).Value = c.Offset(, 6).Value
                            
                            
                            
                            'Adds SUMIFS to grab ORF check # for each claim
                            c.Offset(, 19).FormulaR1C1 = _
                            "=SUMIFS('" & ReconMonth & "_Claims Detail'!C8,'" & ReconMonth & "_Claims Detail'!C12,'1130_" & ReconMonth & "'!RC[-19],'" & ReconMonth & "_Claims Detail'!C4,'1130_" & ReconMonth & "'!RC[-9])"
                            'c.Offset(, 19).Value = c.Offset(, 19).Value
                            
                            
                            
                            
                            'Adds SUMIFS to grab vendor # for each claim
                            c.Offset(, 11).FormulaR1C1 = _
                            "=SUMIFS('" & ReconMonth & "_Claims Detail'!C6,'" & ReconMonth & "_Claims Detail'!C12,'1130_" & ReconMonth & "'!RC[-11],'" & ReconMonth & "_Claims Detail'!C4,'1130_" & ReconMonth & "'!RC[-1])"
                            c.Offset(, 11).Interior.ThemeColor = xlThemeColorAccent4
                            c.Offset(, 11).Interior.TintAndShade = 0.399975585192419
                            'c.Offset(, 11).Value = c.Offset(, 11).Value
                            
                            
                            
                            
                            'pulls claim lines from the "Text" column into the yellow header "Claim" as these are the same.  For other lines, would have to look up to ORF check files for claim schedule number details
                            c.Offset(, 23).FormulaR1C1 = _
                            "=RC[-13]"
                            'c.Offset(, 23).Value = c.Offset(, 23).Value
                            
                          
                            
                            
                            'grab vendor name from FCHN (name of the payee with no location city attached, NOT recipient void/reason code
                            c.Offset(, 21).FormulaR1C1 = _
                            "=XLOOKUP(RC[-2],'" & ReconMonth & "_FCHN YTD'!C[-22],'" & ReconMonth & "_FCHN YTD'!C[-5],""Not Found"")"
                            
            
            
            
                            'grab Trip # from FCHN
                            c.Offset(, 22).FormulaR1C1 = _
                            "=OFFSET(XLOOKUP(RC[-3],'" & ReconMonth & "_FCHN YTD'!C[-23],'" & ReconMonth & "_FCHN YTD'!C[-9],""Not Found""),1,)"
            
                End If

Next c










'Adds formulas for all the check lines.  These are all the COUNTIF lines which have 0, since they are not claims on the claims detail
For Each c In c2.Cells

            If c.Value = 0 Then
                                                                                                                     
                                                                                                                     
                       'populates yellow header ORF Checks (FCHN) from Reference column as these checks already have the check # in the reference column, unlike claim lines
                        c.Offset(, 19).FormulaR1C1 = _
                        "=RC[-10]"
                        c.Offset(, 19).Value = c.Offset(, 19).Value                    'keep this hard code to remove leading zeroes for check #s
                   
                            
                            
                        
                        'look up vendor name from FCHN
                        c.Offset(, 21).FormulaR1C1 = _
                        "=XLOOKUP(RC[-2],'" & ReconMonth & "_FCHN YTD'!C[-22],'" & ReconMonth & "_FCHN YTD'!C[-5],""Not Found"")"
                        
                        
                        
                        
                        'look up vendor # from FCHN
                        c.Offset(, 11).FormulaR1C1 = _
                        "=XLOOKUP(RC[8],'" & ReconMonth & "_FCHN YTD'!C[-12],'" & ReconMonth & "_FCHN YTD'!C[8],""Not Found"")"
                        
                        
                        
                        'look up claim schedule  # from ORF check file
                        c.Offset(, 23).FormulaR1C1 = _
                        "=XLOOKUP(RC[-4],'" & ReconMonth & "_ORF Claim Info'!C[-24],'" & ReconMonth & "_ORF Claim Info'!C[-16],""Not Found"")"
                        
                        
                        
                        'look up trip # from reference column from FCHN
                        c.Offset(, 22).FormulaR1C1 = _
                        "=OFFSET(XLOOKUP(RC[-3],'" & ReconMonth & "_FCHN YTD'!C[-23],'" & ReconMonth & "_FCHN YTD'!C[-9],""Not Found""),1,)"
        
            End If

Next c








'add green divider line to column T, which is where SUMIFS will be added for Macro #7
 CurrentMonthReconItems.Columns(20).Select
 
 With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 5287936
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With



'This just select the last row so that when you go to the recon face sheet after the macro ends, you don't need to scroll down to see the check figures
ThisWorkbook.ActiveSheet.Range("H" & LastRow + 4).Select






 
MinutesElapsed = Format((Timer - StartTime) / 86400, "hh:mm:ss")

MsgBox "Ran successfully in " & MinutesElapsed & " minutes." & vbNewLine & vbNewLine & _
"The macro has finished running." & vbNewLine & vbNewLine & _
"Please click OK."



End Sub

Function Add_Rows(ByVal rowNr, colNr, loopnumber)

        Set cell = Cells(rowNr, colNr)
        cell.Interior.ThemeColor = xlThemeColorAccent4
        cell.EntireRow.Select
                
                
        For i = 1 To (loopnumber - 1)
            cell.EntireRow.Copy
            Selection.Insert Shift:=xlDown
            cell.Offset(-i).Value = loopnumber - i
        Next i


End Function
