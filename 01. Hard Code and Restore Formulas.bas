Sub Hard_Code_References()

    Dim sht As Worksheet
    Set sht = ActiveSheet
    Dim WS As Worksheet
    
     
      Question1 = MsgBox(Prompt:="Replace '=' character in formulas with '#$%' to hard code them?" & vbNewLine & vbNewLine & "You can then delete the original sheet the formulas were linked to." & vbNewLine & vbNewLine & _
      "Then add a new sheet and rename it to the original sheet. " & vbNewLine & vbNewLine & "Finally, click the green button to switch all formula references back, which will now link to the new sheet.", Buttons:=vbQuestion + vbYesNo)
   
      Question2 = MsgBox("Do this for EVERY sheet in the workbook?  This might take much longer but will ensure no #REFs occur.", Buttons:=vbQuestion + vbYesNo)
   

   
        If Question1 = vbYes And Question2 = vbNo Then
                         sht.Cells.Replace What:="=", Replacement:="#$%", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
        ElseIf Question1 = vbYes And Question2 = vbYes Then
                    For Each WS In Worksheets
                          WS.Cells.Replace What:="=", Replacement:="#$%", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
                     Next
        Else
            MsgBox "Macro cancelled by user."
        End If
        
        sht.Activate
        
        MsgBox "Finished!"
        
        

End Sub

Sub Restore_References()

     Dim sht As Worksheet
     Set sht = ActiveSheet
     
    
    Question1 = MsgBox(Prompt:="Replace '#$%' character in formulas with '=' to restore formulas?", Buttons:=vbQuestion + vbYesNo)
    
    Question2 = MsgBox("Do this for EVERY sheet in the workbook?  This might take much longer but will ensure no #REFs occur.", Buttons:=vbQuestion + vbYesNo)
    

    If Question1 = vbYes And Question2 = vbNo Then
                     sht.Cells.Replace What:="#$%", Replacement:="=", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
    ElseIf Question1 = vbYes And Question2 = vbYes Then
                For Each WS In Worksheets
                      WS.Cells.Replace What:="#$%", Replacement:="=", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
                 Next
    Else
        MsgBox "Macro cancelled by user."
    End If
    
    sht.Activate
    
    MsgBox "Finished!"

        
End Sub
