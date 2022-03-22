Attribute VB_Name = "Clear_Previous_Template_Table"
Sub Clear_Template_Rows()

Dim answer As Integer
 
answer = MsgBox("Are you sure you want to delete all data in the below table?" & vbNewLine & vbNewLine & _
"Press YES to continue, or NO to cancel.", vbQuestion + vbYesNo + vbDefaultButton2, "Clear Template Table?")

If answer = vbNo Then
    MsgBox ("Macro cancelled by user.")
    End
End If


Template_End_Row = ThisWorkbook.Sheets("Template").Range("TEMPLATE_SUMMARY").row - 1
ThisWorkbook.Sheets("Template").Range("B8:U" & Template_End_Row).ClearContents


End Sub

Sub Jump_To_Summary()

Application.GoTo Reference:="TEMPLATE_SUMMARY"

End Sub
