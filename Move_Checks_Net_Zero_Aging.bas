Attribute VB_Name = "Move_Checks_Net_Zero_Aging"
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

Sub Move_ORFCHECKS_NET_TO_ZERO_Aging()


Dim Recon_WB As Workbook
Set Recon_WB = ThisWorkbook


FiscalYear = Recon_WB.Sheets("Macro Input").Range("Fiscal_Year")
ReconMonth = Recon_WB.Sheets("Macro Input").Range("Recon_Month")
ReconMonth_Num = Recon_WB.Sheets("Macro Input").Range("ReconMonth_Num")


Dim StartTime As Double
Dim MinutesElapsed As String
'Remember time when macro starts
StartTime = Timer


Move_Net_Zero = MsgBox(Prompt:="Do you want to move check lines which net to zero to the 'Net Zero' sheet? " & vbNewLine & vbNewLine & _
"This macro first moves amounts between the same GL accounts with a SUMIFS formula, by check number and GL account, in case of timing issues with clearing.  It uses column S, or the green divider column for these formulas, and hard codes each SUMIFS as soon as it's added." & vbNewLine & vbNewLine & _
"It then uses a SUMIF formula to move amounts between identical check numbers only, which takes all GL accounts into consideration." & vbNewLine & vbNewLine & _
"This macro will take roughly three minutes to run, as cutting and moving individual rows can be slow." & vbNewLine & vbNewLine & _
"Please click 'Yes' to run the macro, or 'No' to cancel.", Buttons:=vbQuestion + vbYesNo)


If Move_Net_Zero = vbNo Then
        MsgBox ("Macro cancelled.")
        End
End If


Dim ReconSheet As Worksheet
Set ReconSheet = Recon_WB.Sheets(ReconMonth & "_ORF Aging")

Dim NetZero As Worksheet
Set NetZero = Recon_WB.Sheets("Net Zero")


ReconSheet.Activate


LastRow2 = ActiveSheet.Cells(ActiveSheet.Rows.Count, "B").End(xlUp).Row


 'This uses the named range of a header row cell to set up the range to loop through
Dim rownumberheader As Long
rownumberheader = ReconSheet.Range("Fund_Header").Row



Dim ReconItems As Range
Set ReconItems = ReconSheet.Range("A" & rownumberheader + 1 & ":AE" & LastRow2)

ReconItems.Select

  
   
Dim c As Range
Dim c2 As Range

Set c2 = ReconItems.Columns(19)
c2.Select



'Add SUMIFS for checks between same accounts in case of clearing issues
For Each c In c2.Cells

        c.FormulaR1C1 = "=SUMIFS(C[-12],C[1],RC[1],C[-16],RC[-16])"
        c.Value = c.Value

Next c
   
   
   
'Now move checks that clear between same accounts to the "Net Zero" Sheet
c2.Select

'Range to loop through is the last row minus the header cells, as we set the range using the header named range, not from A1.
rowNr = LastRow2 - rownumberheader
currentrow = 1


'Loop that moves SUMIFS which equal zero starts here
Do While currentrow <> rowNr

        Set cell = c2.Cells(currentrow, 1)
        
                Offset = cell.Offset(0, 1)
        
                 If cell.Value = 0 And Offset <> "" Then
                     cell.EntireRow.Cut
                     NetZero.Activate
                     Rows("4:4").Insert Shift:=xlDown
                     ReconSheet.Activate
                     Set cell = c2.Cells(currentrow, 1)
                     cell.Select
                     cell.EntireRow.Delete
                     currentrow = currentrow - 1
                     LastRow2 = ActiveSheet.Cells(ActiveSheet.Rows.Count, "B").End(xlUp).Row
                     rowNr = LastRow2 - rownumberheader
                End If
   
        currentrow = currentrow + 1
   
Loop
     
ActiveWindow.FreezePanes = False




'refresh range in case rows were deleted and moved to net zero tab between same accounts
 LastRow2 = ActiveSheet.Cells(ActiveSheet.Rows.Count, "B").End(xlUp).Row
Set ReconItems = ReconSheet.Range("A" & rownumberheader + 1 & ":AE" & LastRow2)
Set c2 = ReconItems.Columns(19)
c2.Select




'Add normal SUMIF for checks between with same check #
For Each c In c2.Cells

        c.FormulaR1C1 = "=SUMIF(C[1],R[0]C[1],C[-12])"
         c.Value = c.Value
        
Next c
   
   
   
   
   
   
   
 'Now move checks that clear to the "Net Zero" Sheet
c2.Select

rowNr = LastRow2 - rownumberheader
currentrow = 1


'Loop that moves SUMIFs which equal zero starts here
Do While currentrow <> rowNr

       
        Set cell = c2.Cells(currentrow, 1)
        
        Offset = cell.Offset(0, 1)
        
                 If cell.Value = 0 And Offset <> "" Then
                     cell.EntireRow.Cut
                     NetZero.Activate
                     Rows("4:4").Insert Shift:=xlDown
                     ReconSheet.Activate
                     Set cell = c2.Cells(currentrow, 1)
                     cell.Select
                     cell.EntireRow.Delete
                     currentrow = currentrow - 1
                     LastRow2 = ActiveSheet.Cells(ActiveSheet.Rows.Count, "B").End(xlUp).Row
                     rowNr = LastRow2 - rownumberheader
                End If
   
        currentrow = currentrow + 1
   
Loop
   


MinutesElapsed = Format((Timer - StartTime) / 86400, "hh:mm:ss")
 
 
 MsgBox "Ran successfully in " & MinutesElapsed & " minutes." & vbNewLine & vbNewLine & _
"The macro has finished moving lines to the 'Net Zero' sheet." & vbNewLine & vbNewLine & _
"Please press OK."



End Sub




