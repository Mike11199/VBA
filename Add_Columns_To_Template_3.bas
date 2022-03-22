Attribute VB_Name = "Add_Columns_To_Template_3"
Sub Add_Columns_To_Import_Template_3()


Dim wb As Workbook
Set wb = ThisWorkbook

Dim template As Worksheet
Set template = wb.Sheets("Template")


Dim answer As Integer
 
answer = MsgBox("Do you want to add lines to the import template from the current month work orders?" & vbNewLine & vbNewLine & _
"Make sure that the work order tabs have been added to this workbook.  This macro will all clear all previous data on the 'Template' tab." & vbNewLine & vbNewLine & _
"Click YES to continue, or NO to cancel.", vbQuestion + vbYesNo + vbDefaultButton2, "Message Box Title")

If answer = vbNo Then
      MsgBox ("Macro cancelled by user.")
    Exit Sub
End If

Template_End_Row = ThisWorkbook.Sheets("Template").Range("TEMPLATE_SUMMARY").row - 1
wb.Sheets("Template").Range("B8:U" & Template_End_Row).ClearContents

Call DeleteNamedRangesWithREF
Call ChangeLocalNameAndOrScope

Dim current_month_tab As Worksheet
Set current_month_tab = wb.Sheets("Current Month Tabs -->")
aftersheetnumber = current_month_tab.Index + 1

Dim RangeStart As Range
Dim CopyRange As Variant
Dim MapperRange As Variant

Named_Range_1 = wb.Sheets("Macro Input").Range("Named_Range_1")
Named_Range_2 = wb.Sheets("Macro Input").Range("Named_Range_2")
Named_Range_3 = wb.Sheets("Macro Input").Range("Named_Range_3")
Named_Range_4 = wb.Sheets("Macro Input").Range("Named_Range_4")

'=========Change both Named_Range_2 to other number or toother named range to map other tables to a template============
  If BET_RangeNameExists(Named_Range_3) = True Then
        Application.GoTo Reference:=Named_Range_3
        MapperRange = ThisWorkbook.ActiveSheet.Range(Named_Range_3)
Else
        MsgBox ("ERROR: The cell Named_Range_3 is empty on the 'Macro Input' tab, or the named range referenced there is misspelled or does not exist in this workbook." & vbNewLine & vbNewLine & "Please click OK to end the macro.")
        End
End If
'=====================================================================================================================
  
Dim nm
Dim PasteRange As Range
Dim PasteRange2 As Range


templaterows = wb.Sheets("Template").Range("TEMPLATE_SUMMARY").row
rows_added = 0
first_time = 0
rowpaste = 0
not_on_first_wo = 0

For i = 1 To UBound(MapperRange)

        NextWOIndicator = CStr(MapperRange(i, 4))
        nm = Trim(CStr(MapperRange(i, 1)))
        
        
        'this variable to only add payee type once per wo type
        first_run_through = 0

        
        
        
        'When jumping to a new category or work order, make all columns paste starting at the new location of the last row
         firstrowdata = wb.Sheets("Template").Range("B8").Value
        
        If NextWOIndicator = "Yes" And (Not IsEmpty(firstrowdata)) Then
                 rows_added = wb.Sheets("Template").Range("B7").End(xlDown).row + 1
                 not_on_first_wo = not_on_first_wo + 1
                 first_run_through = 1
        End If
            
        
        

        
        
        If Not CStr(MapperRange(i, 2)) = "" Then
        
                   
                        tmp_nm = CStr(MapperRange(i, 2))
                        payee_type = CStr(MapperRange(i, 3))
                        
                        Debug.Print nm

                        If BET_RangeNameExists(nm) = True Then
                
 
                                        Application.GoTo Reference:=nm
                                        Set RangeStart = ActiveSheet.Range(nm)
                                     
                                        RangeStart2 = ActiveSheet.Cells(RangeStart.row + 1, RangeStart.Column).Address
                                        RangeEnd = RangeStart.End(xlDown).Address
                                        RangeEnd = Split(RangeEnd, "$")
                                        RangeEnd = RangeEnd(1) & (RangeEnd(2) + 0)
                                        
                            
                                        CopyRange = wb.ActiveSheet.Range(RangeStart2 & ":" & RangeEnd)
                            
                            
                                        Application.GoTo Reference:=tmp_nm
                                        Set CopyStart = ActiveSheet.Range(tmp_nm)
                                        CopyStart2 = ActiveSheet.Cells(CopyStart.row + 1, CopyStart.Column).Address
                                      
                                        Dim columnpaste As Integer
                                        Dim rows_to_be_added_to_Template As Integer
                                      
                                        columnpaste = (CopyStart.Column)
                                        
                                        
                                        
                                       '===========Set up the 'rowpaste' variable, the last row where data is pasted in.  If it's not the first row, then add 8 rows for the header, and also add rows_added previously add if it's on a work order other than the first=======
                                        rowpaste = (rows_added)
                
                                        If not_on_first_wo = 0 Then
                                                     rowpaste = (rowpaste + 8)
                                        End If
                                        '=============================================================================================================================================================================================================

                                        CopyType = TypeName(CopyRange)
                                        
                
                                        If CopyType = "Variant()" Then
                                                    rows_to_be_added_to_Template = UBound(CopyRange)
                                        Else
                                                    rows_to_be_added_to_Template = 1
                                        End If
                                                    
                                                    
                                        PasteRangeEnd_Row = (rowpaste + rows_to_be_added_to_Template + 1)
                                                    
                                                    
                                        Set PasteRange = ActiveSheet.Range(Cells(rowpaste, columnpaste), Cells(PasteRangeEnd_Row, columnpaste))
                                        
                                        
                                        'Move screen to bottom of paste range
                                        ActiveSheet.Range(Cells(PasteRangeEnd_Row, columnpaste), Cells(PasteRangeEnd_Row, columnpaste)).Activate
                                        
                                        
                                           
                                           'Paste the payee type if that is on the template if it exists as a named range
                                           
                                           If BET_RangeNameExists("TEMPLATE_PAYEE") = True Then
                                                    payeetype = wb.Sheets("Template").Range("TEMPLATE_PAYEE").row - 1
                                                    Set PasteRange2 = ActiveSheet.Range(Cells(rowpaste, 3), Cells(PasteRangeEnd_Row, 3))
                                                    PasteRange2 = payee_type
                                            End If
                                        
                                        
                                        templaterows = wb.Sheets("Template").Range("TEMPLATE_SUMMARY").row - 1
                            
                            
                            
                                        'add rows if out of space on the template
                                                    Do While (rows_to_be_added_to_Template + rows_added) > templaterows
                                                                Rows((templaterows)).Select
                                                                Selection.Insert Shift:=xlDown
                                                                templaterows = templaterows + 1
                                                    Loop
                            
                            
                                      'copy to template
                                        PasteRange = CopyRange
                            
                            End If
            End If
            

        
Next i




'format row heights
Columns("A:A").Select
Selection.RowHeight = 15


'format rows pasted in with wrapped text to not have that, then bring back wrap text for cell A3 merged cell in header
  Cells.Select
    With Selection
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
    End With
    With Selection
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
    End With
    Range("A3:C6").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = True
    End With


ThisWorkbook.Sheets("Template").Range("TEMPLATE_SUMMARY").RowHeight = 22.5

'left align data range
templaterows = wb.Sheets("Template").Range("TEMPLATE_SUMMARY").row - 1

Range("B8:K" & templaterows).HorizontalAlignment = xlLeft

 Range("A3:C6").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = True
    End With


'==========================Clear extra lines pasted by the macro and duplicates===================================================================================

answer2 = MsgBox("Do you want to clear duplicate lines and/or '#N/A' lines that the macro may have added?", vbQuestion + vbYesNo + vbDefaultButton2, "Clear Extra Lines?")


If answer2 = vbYes Then
                
                lastrow5 = wb.Sheets("Template").Range("TEMPLATE_SUMMARY").row
                
                For X = 8 To lastrow5
                
                    cell = Cells(X, 2)
                    Cell_Prior = Cells(X - 1, 2)
                    
                    Amount_Cell = Cells(X, 6)
                    Amount_Cell_Prior = Cells(X - 1, 6)
                    
                     If CStr(cell) = CStr(Cell_Prior) Then
                        If CStr(Amount_Cell) = CStr(Amount_Cell_Prior) Then
                              Cells(X - 1, 2).EntireRow.ClearContents
                        End If
                    End If
                    
                    If IsError(cell) Or IsError(Amount_Cell) Then
                         Cells(X, 2).EntireRow.ClearContents
                    End If
                    
                    If CStr(cell) = "" Then
                        Cells(X, 2).EntireRow.ClearContents
                    End If
                    
                        If CStr(cell) = "Do not fill" Then
                        Cells(X, 2).EntireRow.ClearContents
                    End If
                
                Next X
                
End If
'=============================================================================================================================================================



'activate first sheet and scroll to the top after macro is finished
ThisWorkbook.Sheets("Template").Activate
ThisWorkbook.Sheets("Template").Range("A8").Activate


MsgBox ("The macro has finished" & vbNewLine & vbNewLine & "Please press OK.")


End Sub
Private Function BET_RangeNameExists(nname) As Boolean
Dim n As Name
    BET_RangeNameExists = False
    For Each n In ActiveWorkbook.Names
        If Trim(UCase(n.Name)) = Trim(UCase(nname)) Then
            BET_RangeNameExists = True
            Exit Function
        End If
    Next n
End Function

Sub DeleteNamedRangesWithREF()

    Dim nm As Name

    For Each nm In ActiveWorkbook.Names

        If InStr(nm.Value, "#REF!") > 0 Then

            nm.Delete

        End If

    Next nm

End Sub



Sub ChangeLocalNameAndOrScope()
'Programmatically change a sheet-level range name and/or scope to a new name and/or scope
Dim nm As Name, Ans As Integer, newNm As String
For Each nm In ActiveWorkbook.Names
    If nm.Name Like "*!*" Then 'It is sheet level
        newNm = Replace(nm.Name, "*!", "")
        Range(nm.RefersTo).Name = newNm
        nm.Delete
    End If
Next nm
End Sub



