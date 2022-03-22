Attribute VB_Name = "Add_CalATERS_Formulas_and_Rows"
Sub CALATERS_AddFormulas()


'On Error GoTo Error


Dim answer As Integer
 
answer = MsgBox("Did you hard code (paste values) over formulas for cleared items (green area) and non-cleared items?" & vbNewLine & vbNewLine & _
"This macro will delete all the COUNTIFs in column B.  If not hard coded, formulas pulling amounts from claim lines from the claims detail will break." & vbNewLine & vbNewLine & _
"Press YES to continue, or NO to cancel.", vbQuestion + vbYesNo + vbDefaultButton2, "Message Box Title")

If answer = vbNo Then
      MsgBox ("Macro cancelled by user.")
    Exit Sub
End If


Break_Out_CalATERS = MsgBox(Prompt:="Do you want break out CalATERS lines on the recon face sheet by GER #? " & vbNewLine & vbNewLine & _
"You need to input GER #s from each line's PDF, downloaded by Macro #8, before running this.", Buttons:=vbQuestion + vbYesNo)

If Break_Out_CalATERS = vbNo Then
      MsgBox ("Macro cancelled by user.")
    Exit Sub
End If




Dim StartTime As Double
Dim MinutesElapsed As String
'Remember time when macro starts
StartTime = Timer


Dim ORF_WB As Workbook
Set ORF_WB = ThisWorkbook

GLAccount = ORF_WB.Sheets("Macro Input").Range("GL_Account")
FiscalYear = ORF_WB.Sheets("Macro Input").Range("Fiscal_Year")
ReconMonth = ORF_WB.Sheets("Macro Input").Range("Recon_Month")
ReconMonth_Num = ORF_WB.Sheets("Macro Input").Range("ReconMonth_Num")

Dim CALATERS As Worksheet
Set CALATERS = ORF_WB.Sheets(ReconMonth & "_CalATERS Info")

CALATERS.Activate

Dim ReconSheet As Worksheet
Set ReconSheet = ORF_WB.Sheets("1130_" & ReconMonth)

ReconSheet.Activate


lastrow2 = ActiveSheet.Cells(ActiveSheet.Rows.Count, "A").End(xlUp).row
 
 Dim GL2 As Worksheet
 Set GL2 = ORF_WB.Sheets(ReconMonth & "_GL 1130 Detail")

Debug.Print lastrow2

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


lastrow3 = lastrow2 - LastRow


ReconSheet.Activate


Dim CurrentMonthReconItems As Range
Set CurrentMonthReconItems = ReconSheet.Range("A1" & ":S" & lastrow2)

CurrentMonthReconItems.Select
  
Set CurrentMonthReconItems = CurrentMonthReconItems.Resize(, 29)

CurrentMonthReconItems.Select

Dim c As Range
Dim c2 As Range
Set c2 = CurrentMonthReconItems.Columns(2)

Dim c3 As Range
Set c3 = CurrentMonthReconItems.Columns(11)


Columns("B:B").ClearContents




   '=======================LOOP THAT ENTERS COUNTIF FORMULA TO CALATERS MASTER SHEET==========================================================
For Each c In c2.Cells

            If c.Offset(, 9).Value = "CALATERS" Then
                        If c.Offset(, -1).Value = "CM" Then
                        
                                    
                        c.Offset(, 9).Interior.ThemeColor = xlThemeColorAccent4
                        c.Offset(, 9).Interior.TintAndShade = 0.399975585192419
                        
                        c.FormulaR1C1 = _
                        "=COUNTIF('" & ReconMonth & "_CalATERS Info'!C[5],'1130_" & ReconMonth & "'!RC[24])"
            
                        End If
            End If

Next c
  '=======================END LOOP===========================================================================================================




'Select last row of face sheet in amount column so that macro ends showing the check figures
c2.Select
ThisWorkbook.ActiveSheet.Range("H" & lastrow2).Select





'============LOOP THAT ADDS ROWS AND SUMIFS FORMULAS DEPENDING ON VALUES IN COUNTIF FORMULAS=========================================
colNr = 2

    'For rowNr = 1 To c2.Rows.Count
    For rowNr = 1 To 10000
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








'==============================================LOOP TO ADD SUMIFS FORMULAS FOR  CALATERS LINES==============================================
For Each c In c2.Cells

            If c.Value > 0 Then
        
            
                            c.Offset(, 6).Interior.ThemeColor = xlThemeColorAccent4
                            c.Offset(, 6).Interior.TintAndShade = 0.399975585192419
                            
                            '3/3changed this from c[-5] to c[0] for SUMIFS first sum range part; pulling GER Amount instead of Amount due to corrections
                            'Grab amount from the CALATERS tab
                            c.Offset(, 6).FormulaR1C1 = _
                            "=SUMIFS('" & ReconMonth & "_CalATERS Info'!C[0],'" & ReconMonth & "_CalATERS Info'!C[-7],'1130_" & ReconMonth & "'!R[]C[-6],'" & ReconMonth & "_CalATERS Info'!C[-1],'1130_" & ReconMonth & "'!R[]C[18])"
                            
                            
                            'Grab check # from the CALATERS tab
                            c.Offset(, 19).FormulaR1C1 = _
                            "=SUMIFS('" & ReconMonth & "_CalATERS Info'!C[-19],'" & ReconMonth & "_CalATERS Info'!C[-20],'1130_" & ReconMonth & "'!R[]C[-19],'" & ReconMonth & "_CalATERS Info'!C[-14],'1130_" & ReconMonth & "'!R[]C[5])"
                            
                            'Grab vendor name from the CALATERS tab
                            c.Offset(, 21).FormulaR1C1 = _
                            "=INDEX('" & ReconMonth & "_CalATERS Info'!C[-18],MATCH(1,INDEX(('" & ReconMonth & "_CalATERS Info'!C[-22]='1130_" & ReconMonth & "'!RC[-21])*('" & ReconMonth & "_CalATERS Info'!C[-16]='1130_" & ReconMonth & "'!RC[3]),,),0))"
                            
                            'Grab vendor # from the CALATERS tab
                            c.Offset(, 11).FormulaR1C1 = _
                            "=SUMIFS('" & ReconMonth & "_CalATERS Info'!C[-9],'" & ReconMonth & "_CalATERS Info'!C[-12],'1130_" & ReconMonth & "'!R[]C[-11],'" & ReconMonth & "_CalATERS Info'!C[-6],'1130_" & ReconMonth & "'!R[]C[13])"
                            
                            'Grab trip ID from the CALATERS tab
                            c.Offset(, 22).FormulaR1C1 = _
                            "=SUMIFS('" & ReconMonth & "_CalATERS Info'!C[-18],'" & ReconMonth & "_CalATERS Info'!C[-23],'1130_" & ReconMonth & "'!R[]C[-22],'" & ReconMonth & "_CalATERS Info'!C[-17],'1130_" & ReconMonth & "'!R[]C[2])"
                            

            End If

Next c
'=======================END LOOP===========================================================================================================


LastRow = ActiveSheet.Cells.Find(What:="*", _
                            After:=ActiveSheet.Range("A1"), _
                            LookAt:=xlPart, _
                            LookIn:=xlFormulas, _
                            SearchOrder:=xlByRows, _
                            SearchDirection:=xlPrevious, _
                            MatchCase:=False).row




ThisWorkbook.ActiveSheet.Range("H" & LastRow + 4).Select



MinutesElapsed = Format((Timer - StartTime) / 86400, "hh:mm:ss")

MsgBox "Ran successfully in " & MinutesElapsed & " minutes." & vbNewLine & vbNewLine & _
"The Macro has finished running. Please click OK. "




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




