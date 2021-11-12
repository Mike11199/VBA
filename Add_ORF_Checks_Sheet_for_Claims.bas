Attribute VB_Name = "Add_ORF_Checks_Sheet_for_Claims"
Sub GrabData_ORF_Recon_Claims_Numbers()

On Error GoTo Error

Dim ORF_WB As Workbook
Set ORF_WB = ThisWorkbook

Dim StartTime As Double
Dim MinutesElapsed As String
'Remember time when macro starts
StartTime = Timer


'Change to G Drive and the file folder link in the named range on the Macro Input Sheet
ChDrive "G"
Dim ORF_FILE_LOCATION As String
ORF_FILE_LOCATION = ORF_WB.Sheets("Macro Input").Range("ORF_Files_Folder")
    
ChDrive Left(ORF_FILE_LOCATION, 1)
ChDir ORF_FILE_LOCATION


MsgBox "Select all Excel files in the relevant folder by holding CTRL, or clicking and dragging the mouse.  Then click 'Open' to add and merge files." & vbNewLine & vbNewLine & _
"The macro should automatically open the correct folder based on the 'ORF Files Folder' on the 'Macro Input' sheet."
    
    
Call MergeExcelFiles
Call Merge_Sheets
Call FormatORFClaimsHeader
    
 
MinutesElapsed = Format((Timer - StartTime) / 86400, "hh:mm:ss")

MsgBox "Ran successfully in " & MinutesElapsed & " minutes." & vbNewLine & vbNewLine & _
"The Macro has finished running. Please click OK. " & vbNewLine & vbNewLine & _
"Check that all data was copied to the 'ORF Claim Info' sheet for the current month. " & vbNewLine & vbNewLine & _
"The check numbers in column A have been converted to numbers from text, so you will need to omit the leading zeroes if searching for these to see if they were copied correctly.  " & vbNewLine & vbNewLine & _
"If everything looks correct, manually delete all other tabs besides 'ORF Claim Info' from between the 'ORF Files (Claim #s) -->' and '<-- ORF Files (Claims #s)' index tabs."
   



Exit Sub
Error:  MsgBox "Something went wrong.  Please try again."


End Sub

Sub Merge_Sheets()

Dim ORF_WB As Workbook
Set ORF_WB = ThisWorkbook

Dim startRow, startCol, LastRow, LastCol As Long
Dim headers As Range

'Set Master sheet for consolidation

Dim ReconMonth As Variant
ReconMonth = ORF_WB.Sheets("Macro Input").Range("Recon_Month")




ORF_WB.Sheets.Add(After:=Sheets("ORF Files (Claim #s) -->")).Name = ReconMonth & "_ORF Claim Info"
Dim mstr As Worksheet
Set mstr = ORF_WB.Sheets(ReconMonth & "_ORF Claim Info")

    With mstr.Tab
        .Color = 192
        .TintAndShade = 0
    End With


Dim aftersheetnumber As Long
aftersheetnumber = mstr.Index + 1



'Get Headers



ORF_WB.Sheets(aftersheetnumber).Activate
'Set headers = Application.InputBox("Select the Headers Cells", Type:=8)

Set headers = ActiveSheet.Range("A1:XFD1")


'Copy Headers into master
headers.Copy mstr.Range("A1")
startRow = headers.row + 1
startCol = headers.Column

Debug.Print startRow, startCol
'loop through all sheets



'MsgBox mstr.Index
'MsgBox ORF_WB.Sheets("<-- ORF Files (Claim #s)").Index


For i = aftersheetnumber To ORF_WB.Sheets("<-- ORF Files (Claim #s)").Index - 1
     'except the master sheet from looping
     
        Application.DisplayAlerts = False
        
        
        ORF_WB.Sheets(i).Activate
        


If ORF_WB.Sheets(i).AutoFilterMode Then
    ORF_WB.Sheets(i).AutoFilterMode = False
End If

        
        LastRow = Cells(Rows.Count, startCol).End(xlUp).row
        LastCol = Cells(startRow, Columns.Count).End(xlToLeft).Column
        'get data from each worksheet and copy it into Master sheet
        Range(Cells(startRow, startCol), Cells(LastRow, LastCol)).Copy _
        mstr.Range("A" & mstr.Cells(Rows.Count, 1).End(xlUp).row + 1)

        Application.DisplayAlerts = True
         
Next





mstr.Activate

Cells.Select
Selection.Columns.AutoFit
Columns("A:A").Value = Columns("A:A").Value
Range("A2").Select
ActiveWindow.FreezePanes = True


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
                    countSheets = countSheets + 1
                    wksCurSheet.Copy After:=wbkCurBook.Sheets("ORF Files (Claim #s) -->")
                Next
 
                wbkSrcBook.Close SaveChanges:=False
 
            Next
 
            Application.ScreenUpdating = True
            Application.Calculation = xlCalculationAutomatic
 
            MsgBox "Processed " & countFiles & " files" & vbCrLf & "Merged " & countSheets & " worksheets", Title:="Merge Excel files"
        End If
 
    Else
        MsgBox "No files selected.  Exiting the macro.", Title:="Merge Excel files"
        End
    End If
End Sub


Sub FormatORFClaimsHeader()

 
    Range("I1").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 15773696
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Range("A1").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 49407
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 5296274
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Range("I1").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 5296274
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Rows("1:1").Select
    Range("J1").Activate
    With Selection.Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.249977111117893
        .PatternTintAndShade = 0
    End With
    Range("A1").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 5296274
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Range("I1").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 5296274
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Rows("1:1").Select
    With Selection
        .VerticalAlignment = xlCenter
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
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
    Range("A1").Select


  Columns("A:A").ColumnWidth = 24.14
    Columns("A:A").ColumnWidth = 23
    Columns("B:B").ColumnWidth = 24.43
    Columns("I:I").ColumnWidth = 14.57


Range("A1").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Columns("A:B").Select
    With Selection
        .HorizontalAlignment = xlRight
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Range("A1:B1").Select
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

  Range("A1").Select


End Sub

