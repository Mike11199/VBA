Attribute VB_Name = "Move_ORF_Checks_Net_to_Zero"
Sub Move_ORFCHECKS_NET_TO_ZERO_GREEN_AREA()


'On Error GoTo Error

Dim ORF_WB As Workbook
Set ORF_WB = ThisWorkbook

Dim StartTime As Double
Dim MinutesElapsed As String
'Remember time when macro starts
StartTime = Timer


Move_Net_Zero = MsgBox(Prompt:="Do you want to move ORF check lines which net to zero to the cleared items section, on the top of the recon face sheet? " & vbNewLine & vbNewLine & _
"This macro can be run multiple times.", Buttons:=vbQuestion + vbYesNo)

If Move_Net_Zero = vbNo Then
MsgBox ("Macro cancelled.")
End
End If

GLAccount = ORF_WB.Sheets("Macro Input").Range("GL_Account")
FiscalYear = ORF_WB.Sheets("Macro Input").Range("Fiscal_Year")
ReconMonth = ORF_WB.Sheets("Macro Input").Range("Recon_Month")
ReconMonth_Num = ORF_WB.Sheets("Macro Input").Range("ReconMonth_Num")



'grabs last row from recon face sheet
Dim ReconSheet As Worksheet
Set ReconSheet = ORF_WB.Sheets("1130_" & ReconMonth)
ReconSheet.Activate
lastrow2 = ActiveSheet.Cells(ActiveSheet.Rows.Count, "A").End(xlUp).row

 
'The header cell "Period" is named "Period_Header" as a named range.  This lets the macro only select the items below the header and not the cleared items in light green, if running multiple times
Dim rownumberheader As Long
rownumberheader = ReconSheet.Range("Period_Header").row


'Sets up range for recon items
Dim ReconItems As Range
Set ReconItems = ReconSheet.Range("A" & rownumberheader + 1 & ":AC" & lastrow2)
ReconItems.Select


'Sets up c2 range where the SUMIF formulas will go
Dim c As Range
Dim c2 As Range
Set c2 = ReconItems.Columns(20)
c2.Select



'This add the SUMIF formulas by check # in the green column divider, to see which checks have cleared against other checks
For Each c In c2.Cells
        c.FormulaR1C1 = "=SUMIF(C[1],R[0]C[1],C[-12])"
Next c
   
   
   
   
   
   
c2.Select
   
'This loops through each of the SUMIF formulas, and if the value is 0, or close to it and off by less than a cent for some reason, then moves it to the cleared items 'light green' section at the top of the recon work sheet
For rowNr = c2.row To lastrow2

        Set cell = Cells(rowNr, 20)
        
                If cell.Value > -0.001 And cell.Value < 0.001 Then           ' change to cell.value = 0 if issues with this
                        cell.EntireRow.Cut
                        Rows("1:1").Insert Shift:=xlDown
                        rowNr = rowNr - 1
                End If

Next rowNr
     
     
     
     
     
     
     
ActiveWindow.FreezePanes = False

MinutesElapsed = Format((Timer - StartTime) / 86400, "hh:mm:ss")

MsgBox "Ran successfully in " & MinutesElapsed & " minutes." & vbNewLine & vbNewLine & _
"The macro has finished running." & vbNewLine & vbNewLine & _
"Please click OK."



End Sub


