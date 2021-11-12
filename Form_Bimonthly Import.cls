VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Bimonthly Import"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Clear_Check_Boxes_Click()

Dim ctl As Control
    For Each ctl In Me.Controls
        If ctl.ControlType = acCheckBox Then
            ctl.Value = False
        End If
    Next ctl



End Sub

Private Sub Command10_Click()



CarryOn = MsgBox("This will import the 2nd BD Extract into the database and run queries to delete and append the 'Invoice' and 'Payment' tables.  Please select the 2nd BD Extract file.  All three BD Extracts should have the same format, so this can be any of the three (order doesn't matter).  Continue with the import?", vbYesNo, "Continue?")
        If CarryOn = vbNo Then
            Exit Sub
        End If

CarryOn = MsgBox("Delete the previous 'BD Raw' table?", vbYesNo, "Continue?")
        If CarryOn = vbNo Then
            Exit Sub
        End If


If Not IsNull(DLookup("Name", "MSysObjects", "Name='BD Raw' AND Type = 1")) Then
    DoCmd.DeleteObject acTable, "BD Raw"
End If


MsgBox ("Deleted the previous 'BD Raw' table.")




Const msoFileDialogFilePicker As Long = 3
Dim objDialog As Object

'===========OPENS FILE PICKER DIALOG BOX=====================================================================================
Set objDialog = Application.FileDialog(msoFileDialogFilePicker)

With objDialog
                            .AllowMultiSelect = False
                            .InitialFileName = ("\\MyFiles04\data\vol1\dptshare\FINANCIAL SERVICES\Pension Solution\Conversion\CB-RB Access\BD Extracts")               'Edit this to change default folder
                            .Show
        
            If .SelectedItems.Count = 0 Then
                            MsgBox "No file selected."
                            Exit Sub
            Else
                 
                 For Each varItem In .SelectedItems
                            strFile = Dir(varItem)
                            strFolder = Left(varItem, Len(varItem) - Len(strFile))
                            p = strFolder & strFile
                Next
                
            End If
End With


'===========EDITS PATH OF SAVED IMPORT TO FILE CHOSEN IN FILE PICKER==========================================================
    Dim ies As ImportExportSpecification, i As Long, oldXML() As String, newXML As String

     'newXlsxFileSpec = "C:\Users\Gord\Desktop\anotherTest.xlsx"  ' for testing
     newXlsxFileSpec = p

    Set ies = CurrentProject.ImportExportSpecifications("Import BD Extract_2")              'Change this to name of saved import
    oldXML = Split(ies.XML, vbCrLf, -1, vbBinaryCompare)
    newXML = ""
    For i = 0 To UBound(oldXML)
        If i = 1 Then
            ' re-write the second line of the existing XML
            newXML = newXML & _
                    "<ImportExportSpecification Path = """ & _
                    newXlsxFileSpec & _
                    """ xmlns=""urn:www.microsoft.com/office/access/imexspec"">" & _
                    vbCrLf
        Else
            newXML = newXML & oldXML(i) & vbCrLf
        End If
    Next
    ies.XML = newXML
    
'=======EXECUTE THE SAVED IMPORT WITH THE NEW FILE LOCATION================================================================
    ies.Execute
    Set ies = Nothing






'DoCmd.RunSavedImportExport "Import BD Extract_1"
MsgBox ("Imported BD Extract 2.")




'DoCmd.RunSavedImportExport "Import BD Extract_2"


CarryOn = MsgBox("Run the next query:  '2 Append Invoice' ?", vbYesNo, "Continue?")
        If CarryOn = vbNo Then
            Exit Sub
        End If


DoCmd.OpenQuery "2 Append Invoice"
MsgBox ("Ran Query '2 Append Invoice'.  This appended the 2nd BD extract to a fresh 'Invoice' Table.")

CarryOn = MsgBox("Run the next query:  '2 Append Payment' ?", vbYesNo, "Continue?")
        If CarryOn = vbNo Then
            Exit Sub
        End If

DoCmd.OpenQuery "2 Append Payment"
MsgBox ("Ran Query '2 Append Payment'.  This appended the 2nd BD extract to a fresh 'Payment' Table.")

CarryOn = MsgBox("Run the next query:  'MSI - Update Payment Table Warrant #' ?", vbYesNo, "Continue?")
        If CarryOn = vbNo Then
            Exit Sub
        End If

DoCmd.OpenQuery "MSI - Update Payment Table Warrant #"
MsgBox ("Ran Query 'MSI - Update Payment Table Warrant #'.  This trims the FY off warrant #s. if present")

MsgBox ("The macro has finished running.  Please click OK.")


End Sub

Private Sub Command11_Click()


CarryOn = MsgBox("This will import the 3rd BD Extract into the database and run queries to delete and append the 'Invoice' and 'Payment' tables.  Please select the 3rd BD Extract file.  All three BD Extracts should have the same format, so this can be any of the three (order doesn't matter).  Continue with the import?", vbYesNo, "Continue?")
        If CarryOn = vbNo Then
            Exit Sub
        End If

CarryOn = MsgBox("Delete the previous 'BD Raw' table?", vbYesNo, "Continue?")
        If CarryOn = vbNo Then
            Exit Sub
        End If


If Not IsNull(DLookup("Name", "MSysObjects", "Name='BD Raw' AND Type = 1")) Then
    DoCmd.DeleteObject acTable, "BD Raw"
End If


MsgBox ("Deleted the previous 'BD Raw' table.")
'




Const msoFileDialogFilePicker As Long = 3
Dim objDialog As Object

'===========OPENS FILE PICKER DIALOG BOX=====================================================================================
Set objDialog = Application.FileDialog(msoFileDialogFilePicker)

With objDialog
                            .AllowMultiSelect = False
                            .InitialFileName = ("\\MyFiles04\data\vol1\dptshare\FINANCIAL SERVICES\Pension Solution\Conversion\CB-RB Access\BD Extracts")               'Edit this to change default folder
                            .Show
        
            If .SelectedItems.Count = 0 Then
                            MsgBox "No file selected."
                            Exit Sub
            Else
                 
                 For Each varItem In .SelectedItems
                            strFile = Dir(varItem)
                            strFolder = Left(varItem, Len(varItem) - Len(strFile))
                            p = strFolder & strFile
                Next
                
            End If
End With


'===========EDITS PATH OF SAVED IMPORT TO FILE CHOSEN IN FILE PICKER==========================================================
    Dim ies As ImportExportSpecification, i As Long, oldXML() As String, newXML As String

     'newXlsxFileSpec = "C:\Users\Gord\Desktop\anotherTest.xlsx"  ' for testing
     newXlsxFileSpec = p

    Set ies = CurrentProject.ImportExportSpecifications("Import BD Extract_2")              'Change this to name of saved import
    oldXML = Split(ies.XML, vbCrLf, -1, vbBinaryCompare)
    newXML = ""
    For i = 0 To UBound(oldXML)
        If i = 1 Then
            ' re-write the second line of the existing XML
            newXML = newXML & _
                    "<ImportExportSpecification Path = """ & _
                    newXlsxFileSpec & _
                    """ xmlns=""urn:www.microsoft.com/office/access/imexspec"">" & _
                    vbCrLf
        Else
            newXML = newXML & oldXML(i) & vbCrLf
        End If
    Next
    ies.XML = newXML
    
'=======EXECUTE THE SAVED IMPORT WITH THE NEW FILE LOCATION================================================================
    ies.Execute
    Set ies = Nothing






'DoCmd.RunSavedImportExport "Import BD Extract_1"
MsgBox ("Imported BD Extract 3.")



'DoCmd.RunSavedImportExport "Import BD Extract_3"

CarryOn = MsgBox("Run the next query:  '2 Append Invoice' ?", vbYesNo, "Continue?")
        If CarryOn = vbNo Then
            Exit Sub
        End If

DoCmd.OpenQuery "2 Append Invoice"
MsgBox ("Ran Query '2 Append Invoice'.  This appended the 3rd BD extract to a fresh 'Invoice' Table.")

CarryOn = MsgBox("Run the next query:  '2 Append Payment' ?", vbYesNo, "Continue?")
        If CarryOn = vbNo Then
            Exit Sub
        End If

DoCmd.OpenQuery "2 Append Payment"
MsgBox ("Ran Query '2 Append Payment'.  This appended the 3rd BD extract to a fresh 'Payment' Table.")

CarryOn = MsgBox("Run the next query:  'MSI - Update Payment Table Warrant #' ?", vbYesNo, "Continue?")
        If CarryOn = vbNo Then
            Exit Sub
        End If

DoCmd.OpenQuery "MSI - Update Payment Table Warrant #"
MsgBox ("Ran Query 'MSI - Update Payment Table Warrant #'.  This trims the FY off warrant #s. if present")

MsgBox ("The macro has finished running.  Please click OK.")

End Sub






Private Sub Command8_Click()

CarryOn = MsgBox("This will import the 1st BD Extract into the database and run queries to delete and append the 'Invoice' and 'Payment' tables.  Please select the first BD Extract file.  All three BD Extracts should have the same format, so this can be any of the three (order doesn't matter).  Continue with the import?", vbYesNo, "Continue?")
        If CarryOn = vbNo Then
            Exit Sub
        End If


CarryOn = MsgBox("Delete the previous 'BD Raw' table?", vbYesNo, "Continue?")
        If CarryOn = vbNo Then
            Exit Sub
        End If



If Not IsNull(DLookup("Name", "MSysObjects", "Name='BD Raw' AND Type = 1")) Then
    DoCmd.DeleteObject acTable, "BD Raw"
End If


MsgBox ("Deleted the previous 'BD Raw' table.")




Const msoFileDialogFilePicker As Long = 3
Dim objDialog As Object

'===========OPENS FILE PICKER DIALOG BOX=====================================================================================
Set objDialog = Application.FileDialog(msoFileDialogFilePicker)

With objDialog
                            .AllowMultiSelect = False
                            .InitialFileName = ("\\MyFiles04\data\vol1\dptshare\FINANCIAL SERVICES\Pension Solution\Conversion\CB-RB Access\BD Extracts")               'Edit this to change default folder
                            .Show
        
            If .SelectedItems.Count = 0 Then
                            MsgBox "No file selected."
                            Exit Sub
            Else
                 
                 For Each varItem In .SelectedItems
                            strFile = Dir(varItem)
                            strFolder = Left(varItem, Len(varItem) - Len(strFile))
                            p = strFolder & strFile
                Next
                
            End If
End With


'===========EDITS PATH OF SAVED IMPORT TO FILE CHOSEN IN FILE PICKER==========================================================
    Dim ies As ImportExportSpecification, i As Long, oldXML() As String, newXML As String

     'newXlsxFileSpec = "C:\Users\Gord\Desktop\anotherTest.xlsx"  ' for testing
     newXlsxFileSpec = p

    Set ies = CurrentProject.ImportExportSpecifications("Import BD Extract_1")              'Change this to name of saved import
    oldXML = Split(ies.XML, vbCrLf, -1, vbBinaryCompare)
    newXML = ""
    For i = 0 To UBound(oldXML)
        If i = 1 Then
            ' re-write the second line of the existing XML
            newXML = newXML & _
                    "<ImportExportSpecification Path = """ & _
                    newXlsxFileSpec & _
                    """ xmlns=""urn:www.microsoft.com/office/access/imexspec"">" & _
                    vbCrLf
        Else
            newXML = newXML & oldXML(i) & vbCrLf
        End If
    Next
    ies.XML = newXML
    
'=======EXECUTE THE SAVED IMPORT WITH THE NEW FILE LOCATION================================================================
    ies.Execute
    Set ies = Nothing






'DoCmd.RunSavedImportExport "Import BD Extract_1"
MsgBox ("Imported BD Extract 1.")


CarryOn = MsgBox("Delete the previous 'Invoice' table by running the query 'Z Delete Invoice' ?", vbYesNo, "Continue?")
        If CarryOn = vbNo Then
            Exit Sub
        End If

DoCmd.OpenQuery "Z Delete Invoice"
MsgBox ("Ran Query 'Z Delete Invoice'.  This deleted the Invoice Table.")

CarryOn = MsgBox("Run the next query:  '2 Append Invoice' ?", vbYesNo, "Continue?")
        If CarryOn = vbNo Then
            Exit Sub
        End If

DoCmd.OpenQuery "2 Append Invoice"
MsgBox ("Ran Query '2 Append Invoice'.  This appended the 1st BD extract to a fresh 'Invoice' Table.")

CarryOn = MsgBox("Delete the previous 'Payment' table by running the query 'Z Delete Payment' ?", vbYesNo, "Continue?")
        If CarryOn = vbNo Then
            Exit Sub
        End If

DoCmd.OpenQuery "Z Delete Payment"
MsgBox ("Ran Query 'Z Delete Payment'.  This deleted the Payment Table")

CarryOn = MsgBox("Run the next query:  '2 Append Payment' ?", vbYesNo, "Continue?")
        If CarryOn = vbNo Then
            Exit Sub
        End If

DoCmd.OpenQuery "2 Append Payment"
MsgBox ("Ran Query '2 Append Payment'.  This appended the 1st BD extract to a fresh 'Payment' Table.")

CarryOn = MsgBox("Run the next query:  'MSI - Update Payment Table Warrant #' ?  This trims the FY off of warrant #s if present in the table.", vbYesNo, "Continue?")
        If CarryOn = vbNo Then
            Exit Sub
        End If

DoCmd.OpenQuery "MSI - Update Payment Table Warrant #"

MsgBox ("Ran Query 'MSI - Update Payment Table Warrant #'. ")

MsgBox ("The macro has finished running.  Please click OK.")

End Sub

Private Sub Ctl2_Append_AR_from_Invoice_Click()


CarryOn = MsgBox("Run the next query:  '2 Append AR from Invoice' ? ", vbYesNo, "Continue?")
If CarryOn = vbNo Then
    Exit Sub
End If

DoCmd.OpenQuery "2 Append AR from Invoice"
MsgBox ("Ran Query '2 Append AR from Invoice'.  ")



End Sub

Private Sub Ctl3_Append_Client_Tax_from_Payment_Click()


CarryOn = MsgBox("Run query '3 Append Client Tax from Payment' ?", vbYesNo, "Continue?")
      If CarryOn = vbNo Then
          Exit Sub
      End If


DoCmd.OpenQuery "3 Append Client Tax from Payment"
MsgBox ("Ran Query '3 Append Client Tax from Payment'.")



End Sub

Private Sub Ctl3_Append_Invoice_Corrections_Click()


CarryOn = MsgBox("Run query '3 Append Invoice Corrections' ?  This loads the records from the 'Invoice Corrections' table into the 'Invoice' table.", vbYesNo, "Continue?")
      If CarryOn = vbNo Then
          Exit Sub
      End If

DoCmd.OpenQuery "3 Append Invoice Corrections"
MsgBox ("Ran Query '3 Append Invoice Corrections'. ")




End Sub

Private Sub Ctl3_Append_Vendor_from_Payment_Click()


CarryOn = MsgBox("Run query '3 Append Vendor from Payment' ?", vbYesNo, "Continue?")
      If CarryOn = vbNo Then
          Exit Sub
      End If


DoCmd.OpenQuery "3 Append Vendor from Payment"
MsgBox ("Ran Query '3 Append Vendor from Payment'.  ")



End Sub

Private Sub Ctl3_Update_1042_S_with_START_account_Click()


CarryOn = MsgBox("Run the query:  '3 Update 1042-S with START account' ? ", vbYesNo, "Continue?")
If CarryOn = vbNo Then
    Exit Sub
End If

DoCmd.OpenQuery "3 Update 1042-S with START account"
MsgBox ("Ran Query '3 Update 1042-S with START account'.  ")



End Sub

Private Sub Ctl3_Update_AR_with_START_account_Click()

CarryOn = MsgBox("Run the query:  '3 Update AR with START account' ? ", vbYesNo, "Continue?")
If CarryOn = vbNo Then
    Exit Sub
End If

DoCmd.OpenQuery "3 Update AR with START account"
MsgBox ("Ran Query '3 Update AR with START account'.  ")


End Sub

Private Sub Ctl3_Update_Client_Tax_with_START_account_Click()


CarryOn = MsgBox("Run the query:  '3 Update Client Tax with START account' ? ", vbYesNo, "Continue?")
If CarryOn = vbNo Then
    Exit Sub
End If

DoCmd.OpenQuery "3 Update Client Tax with START account"
MsgBox ("Ran Query '3 Update Client Tax with START account'.  ")




End Sub

Private Sub Ctl3_Update_Invoice_Account_from_Client_Tax_Click()

CarryOn = MsgBox("Run the query:  '3 Update Invoice Account from Client Tax' ? ", vbYesNo, "Continue?")
If CarryOn = vbNo Then
    Exit Sub
End If

DoCmd.OpenQuery "3 Update Invoice Account from Client Tax"
MsgBox ("Ran Query '3 Update Invoice Account from Client Tax'.  ")



End Sub

Private Sub Ctl3_Update_Invoice_START_account_id_Click()

CarryOn = MsgBox("Run SQL code to convert recip_client_id field in START Account ID table to double format (needed for this query) ? ", vbYesNo, "Continue?")
If CarryOn = vbNo Then
    Exit Sub
End If


'This converts the recip_client_id table to double so that the next query can work.
Dim strSQL As String
   strSQL = "ALTER TABLE [START Account ID] ALTER COLUMN recip_client_id double"
   DoCmd.RunSQL strSQL
   
CarryOn = MsgBox("Run the query:  '3 Update Invoice START account id' ? ", vbYesNo, "Continue?")
If CarryOn = vbNo Then
    Exit Sub
End If


DoCmd.OpenQuery "3 Update Invoice START account id"
MsgBox ("Ran Query '3 Update Invoice START account id'.  ")






End Sub

Private Sub Ctl3_Update_Payment_START_account_id_Click()

CarryOn = MsgBox("Run the query:  '3 Update Payment START account id' ? ", vbYesNo, "Continue?")
If CarryOn = vbNo Then
    Exit Sub
End If

DoCmd.OpenQuery "3 Update Payment START account id"
MsgBox ("Ran Query '3 Update Payment START account id'.  ")



End Sub

Private Sub Ctl3_Update_START_Account_ID_tax_id_leading_zero_Click()


CarryOn = MsgBox("Run query '3 Update START Account ID tax id leading zero' ?  ", vbYesNo, "Continue?")
      If CarryOn = vbNo Then
          Exit Sub
      End If

DoCmd.OpenQuery "3 Update START Account ID tax id leading zero"
MsgBox ("Ran Query '3 Update START Account ID tax id leading zero'")



End Sub

Private Sub Ctl3_Update_Vendor_Tax_from_Client_Tax_Click()

CarryOn = MsgBox("Run the query:  '3 Update Vendor Tax from Client Tax' ? ", vbYesNo, "Continue?")
If CarryOn = vbNo Then
    Exit Sub
End If

DoCmd.OpenQuery "3 Update Vendor Tax from Client Tax"
MsgBox ("Ran Query '3 Update Vendor Tax from Client Tax'.  ")



End Sub

Private Sub Ctl3_Update_W2_Annual_with_START_account_Click()

CarryOn = MsgBox("Run the query:  '3 Update W2 Annual with START account' ? ", vbYesNo, "Continue?")
If CarryOn = vbNo Then
    Exit Sub
End If

DoCmd.OpenQuery "3 Update W2 Annual with START account"
MsgBox ("Ran Query '3 Update W2 Annual with START account'.  ")



End Sub

Private Sub Ctl3_Update_W2_Payment_with_START_account_Click()

CarryOn = MsgBox("Run the query:  '3 Update W2 Payment with START account' ? ", vbYesNo, "Continue?")
If CarryOn = vbNo Then
    Exit Sub
End If

DoCmd.OpenQuery "3 Update W2 Payment with START account"
MsgBox ("Ran Query '3 Update W2 Payment with START account'.  ")



End Sub

Private Sub Ctl3_Update_Worksheet_with_START_account_Click()

CarryOn = MsgBox("Run the query:  '3 Update Worksheet with START account' ? ", vbYesNo, "Continue?")
If CarryOn = vbNo Then
    Exit Sub
End If

DoCmd.OpenQuery "3 Update Worksheet with START account"
MsgBox ("Ran Query '3 Update Worksheet with START account'.  ")


End Sub

Private Sub Ctl5_Update_Vendor_Tax_Client_from_Client_Tax_Click()

CarryOn = MsgBox("Run the query:  '5 Update Vendor Tax Client from Client Tax' ? ", vbYesNo, "Continue?")
If CarryOn = vbNo Then
    Exit Sub
End If

DoCmd.OpenQuery "5 Update Vendor Tax Client from Client Tax"
MsgBox ("Ran Query '5 Update Vendor Tax Client from Client Tax'.  ")



End Sub

Private Sub Ctl8_0_Update_Invoice_Vendor_Client_Special_GL_Click()


CarryOn = MsgBox("Run query '8-0 Update Invoice Vendor Client Special GL' ?", vbYesNo, "Continue?")
      If CarryOn = vbNo Then
          Exit Sub
      End If

DoCmd.OpenQuery "8-0 Update Invoice Vendor Client Special GL"
MsgBox ("Ran Query '8-0 Update Invoice Vendor Client Special GL'.")

End Sub

Private Sub Ctl8_0_Update_Payment_from_Invoice_Click()

CarryOn = MsgBox("Run query '8-0 Update Payment from Invoice' ?", vbYesNo, "Continue?")
      If CarryOn = vbNo Then
          Exit Sub
      End If


DoCmd.OpenQuery "8-0 Update Payment from Invoice"
MsgBox ("Ran Query '8-0 Update Payment from Invoice'.")

End Sub

Private Sub Ctl8_0_Update_Payment_Vendor_Client_Account_Click()

CarryOn = MsgBox("Run query '8-0 Update Payment Vendor Client Account' ?", vbYesNo, "Continue?")
      If CarryOn = vbNo Then
          Exit Sub
      End If


DoCmd.OpenQuery "8-0 Update Payment Vendor Client Account"
MsgBox ("Ran Query '8-0 Update Payment Vendor Client Account'. ")


End Sub

Private Sub MSI___Medicare_Employer_I_Click()

CarryOn = MsgBox("Run the query:  'MSI - Medicare Employer I' ?", vbYesNo, "Continue?")
If CarryOn = vbNo Then
    Exit Sub
End If

DoCmd.OpenQuery "MSI - Medicare Employer I"
MsgBox ("Ran Query 'MSI - Medicare Employer I'.  ")

End Sub

Private Sub MSI___Medicare_Employer_II_Click()

CarryOn = MsgBox("Run the query:  'MSI - Medicare Employer II' ?", vbYesNo, "Continue?")
If CarryOn = vbNo Then
    Exit Sub
End If

DoCmd.OpenQuery "MSI - Medicare Employer II"
MsgBox ("Ran Query 'MSI - Medicare Employer II'.  ")

End Sub

Private Sub MSI___Medicare_Employer_III_Click()

CarryOn = MsgBox("Run the query:  'MSI - Medicare Employer III' ?", vbYesNo, "Continue?")
If CarryOn = vbNo Then
    Exit Sub
End If

DoCmd.OpenQuery "MSI - Medicare Employer III"
MsgBox ("Ran Query 'MSI - Medicare Employer III'.  ")

End Sub

Private Sub Run_All_BD_Extract_Queries_Click()

CarryOn = MsgBox("This macro will run all six BD Extract queries.  Start by running query '3 Append Invoice Corrections' ?  This loads the records from the 'Invoice Corrections' table into the 'Invoice' table.", vbYesNo, "Continue?")
      If CarryOn = vbNo Then
          Exit Sub
      End If

DoCmd.OpenQuery "3 Append Invoice Corrections"
MsgBox ("Ran Query '3 Append Invoice Corrections'. ")

CarryOn = MsgBox("Run the next query:  '3 Append Client Tax from Payment' ?", vbYesNo, "Continue?")
If CarryOn = vbNo Then
    Exit Sub
End If
        
DoCmd.OpenQuery "3 Append Client Tax from Payment"
MsgBox ("Ran Query '3 Append Client Tax from Payment'.")

CarryOn = MsgBox("Run the next query:  '3 Append Vendor from Payment' ?", vbYesNo, "Continue?")
If CarryOn = vbNo Then
    Exit Sub
End If

DoCmd.OpenQuery "3 Append Vendor from Payment"
MsgBox ("Ran Query '3 Append Vendor from Payment'.  ")

CarryOn = MsgBox("Run the next query:  '8-0 Update Invoice Vendor Client Special GL' ?", vbYesNo, "Continue?")
If CarryOn = vbNo Then
    Exit Sub
End If

DoCmd.OpenQuery "8-0 Update Invoice Vendor Client Special GL"
MsgBox ("Ran Query '8-0 Update Invoice Vendor Client Special GL'.")

CarryOn = MsgBox("Run the next query:  '8-0 Update Payment from Invoice' ?", vbYesNo, "Continue?")
If CarryOn = vbNo Then
    Exit Sub
End If

DoCmd.OpenQuery "8-0 Update Payment from Invoice"
MsgBox ("Ran Query '8-0 Update Payment from Invoice'.")

CarryOn = MsgBox("Run the next query:  '8-0 Update Payment Vendor Client Account' ?", vbYesNo, "Continue?")
If CarryOn = vbNo Then
    Exit Sub
End If

DoCmd.OpenQuery "8-0 Update Payment Vendor Client Account"

MsgBox ("Ran Query '8-0 Update Payment Vendor Client Account'. ")

MsgBox ("The macro has finished running.  Please click OK.")




End Sub

Private Sub Run_Start_Queries_Click()
 

CarryOn = MsgBox("This macro will run all fourteen START Extract queries.  Start by running query '3 Update START Account ID tax id leading zero' ?  ", vbYesNo, "Continue?")
      If CarryOn = vbNo Then
          Exit Sub
      End If

DoCmd.OpenQuery "3 Update START Account ID tax id leading zero"
MsgBox ("Ran Query '3 Update START Account ID tax id leading zero'")

CarryOn = MsgBox("Run the next query:  'Z Delete AR' ?  This deletes the AR Table.", vbYesNo, "Continue?")
If CarryOn = vbNo Then
    Exit Sub
End If


DoCmd.OpenQuery "Z Delete AR"
MsgBox ("Ran Query 'Z Delete AR'.")

CarryOn = MsgBox("Run the next query:  '2 Append AR from Invoice' ? ", vbYesNo, "Continue?")
If CarryOn = vbNo Then
    Exit Sub
End If

DoCmd.OpenQuery "2 Append AR from Invoice"
MsgBox ("Ran Query '2 Append AR from Invoice'.  ")

CarryOn = MsgBox("Run the next query:  '3 Update 1042-S with START account' ? ", vbYesNo, "Continue?")
If CarryOn = vbNo Then
    Exit Sub
End If

DoCmd.OpenQuery "3 Update 1042-S with START account"
MsgBox ("Ran Query '3 Update 1042-S with START account'.  ")

CarryOn = MsgBox("Run the next query:  '3 Update AR with START account' ? ", vbYesNo, "Continue?")
If CarryOn = vbNo Then
    Exit Sub
End If

DoCmd.OpenQuery "3 Update AR with START account"
MsgBox ("Ran Query '3 Update AR with START account'.  ")


CarryOn = MsgBox("Run SQL code to convert recip_client_id field in START Account ID table to double format (needed for the next query) ? ", vbYesNo, "Continue?")
If CarryOn = vbNo Then
    Exit Sub
End If


'This converts the recip_client_id table to double so that the next query can work.
Dim strSQL As String
   strSQL = "ALTER TABLE [START Account ID] ALTER COLUMN recip_client_id double"
   DoCmd.RunSQL strSQL
   
CarryOn = MsgBox("Run the next query:  '3 Update Invoice START account id' ? ", vbYesNo, "Continue?")
If CarryOn = vbNo Then
    Exit Sub
End If


DoCmd.OpenQuery "3 Update Invoice START account id"
MsgBox ("Ran Query '3 Update Invoice START account id'.  ")

CarryOn = MsgBox("Run the next query:  '3 Update Invoice Account from Client Tax' ? ", vbYesNo, "Continue?")
If CarryOn = vbNo Then
    Exit Sub
End If

DoCmd.OpenQuery "3 Update Invoice Account from Client Tax"
MsgBox ("Ran Query '3 Update Invoice Account from Client Tax'.  ")

CarryOn = MsgBox("Run the next query:  '3 Update Payment START account id' ? ", vbYesNo, "Continue?")
If CarryOn = vbNo Then
    Exit Sub
End If

DoCmd.OpenQuery "3 Update Payment START account id"
MsgBox ("Ran Query '3 Update Payment START account id'.  ")

CarryOn = MsgBox("Run the next query:  '3 Update Client Tax with START account' ? ", vbYesNo, "Continue?")
If CarryOn = vbNo Then
    Exit Sub
End If

DoCmd.OpenQuery "3 Update Client Tax with START account"
MsgBox ("Ran Query '3 Update Client Tax with START account'.  ")

CarryOn = MsgBox("Run the next query:  '3 Update Vendor Tax from Client Tax' ? ", vbYesNo, "Continue?")
If CarryOn = vbNo Then
    Exit Sub
End If

DoCmd.OpenQuery "3 Update Vendor Tax from Client Tax"
MsgBox ("Ran Query '3 Update Vendor Tax from Client Tax'.  ")

CarryOn = MsgBox("Run the next query:  '3 Update W2 Annual with START account' ? ", vbYesNo, "Continue?")
If CarryOn = vbNo Then
    Exit Sub
End If

DoCmd.OpenQuery "3 Update W2 Annual with START account"
MsgBox ("Ran Query '3 Update W2 Annual with START account'.  ")

CarryOn = MsgBox("Run the next query:  '3 Update W2 Payment with START account' ? ", vbYesNo, "Continue?")
If CarryOn = vbNo Then
    Exit Sub
End If

DoCmd.OpenQuery "3 Update W2 Payment with START account"
MsgBox ("Ran Query '3 Update W2 Payment with START account'.  ")

CarryOn = MsgBox("Run the next query:  '3 Update Worksheet with START account' ? ", vbYesNo, "Continue?")
If CarryOn = vbNo Then
    Exit Sub
End If

DoCmd.OpenQuery "3 Update Worksheet with START account"
MsgBox ("Ran Query '3 Update Worksheet with START account'.  ")

CarryOn = MsgBox("Run the next query:  '5 Update Vendor Tax Client from Client Tax' ? ", vbYesNo, "Continue?")
If CarryOn = vbNo Then
    Exit Sub
End If

DoCmd.OpenQuery "5 Update Vendor Tax Client from Client Tax"
MsgBox ("Ran Query '5 Update Vendor Tax Client from Client Tax'.  ")




MsgBox ("The macro has finished running.  Please click OK.")






End Sub

Private Sub START_Import_Click()


CarryOn = MsgBox("This will import the combined START Extract into the database.  First the previous 'START Account ID' table will be deleted, then select the START Extract file afterwards to import.  Continue with the import?", vbYesNo, "Continue?")
        If CarryOn = vbNo Then
            Exit Sub
        End If


CarryOn = MsgBox("Delete the previous 'START Account ID' table?", vbYesNo, "Continue?")
        If CarryOn = vbNo Then
            Exit Sub
        End If



If Not IsNull(DLookup("Name", "MSysObjects", "Name='START Account ID' AND Type = 1")) Then
    DoCmd.DeleteObject acTable, "START Account ID"
End If

MsgBox ("Deleted the previous 'START Account ID' table.  Please select the START Extract file for import.")




Const msoFileDialogFilePicker As Long = 3
Dim objDialog As Object

'===========OPENS FILE PICKER DIALOG BOX=====================================================================================
Set objDialog = Application.FileDialog(msoFileDialogFilePicker)

With objDialog
                            .AllowMultiSelect = False
                            .InitialFileName = ("\\MyFiles04\data\vol1\dptshare\FINANCIAL SERVICES\Pension Solution\Conversion\CB-RB Access\START Extracts")               'Edit this to change default folder
                            .Show
        
            If .SelectedItems.Count = 0 Then
                            MsgBox "No file selected."
                            Exit Sub
            Else
                 
                 For Each varItem In .SelectedItems
                            strFile = Dir(varItem)
                            strFolder = Left(varItem, Len(varItem) - Len(strFile))
                            p = strFolder & strFile
                Next
                
            End If
End With


'===========EDITS PATH OF SAVED IMPORT TO FILE CHOSEN IN FILE PICKER==========================================================
    Dim ies As ImportExportSpecification, i As Long, oldXML() As String, newXML As String

     'newXlsxFileSpec = "C:\Users\Gord\Desktop\anotherTest.xlsx"  ' for testing
     newXlsxFileSpec = p

    Set ies = CurrentProject.ImportExportSpecifications("Import-START EXTRACT")              'Change this to name of saved import
    oldXML = Split(ies.XML, vbCrLf, -1, vbBinaryCompare)
    newXML = ""
    For i = 0 To UBound(oldXML)
        If i = 1 Then
            ' re-write the second line of the existing XML
            newXML = newXML & _
                    "<ImportExportSpecification Path = """ & _
                    newXlsxFileSpec & _
                    """ xmlns=""urn:www.microsoft.com/office/access/imexspec"">" & _
                    vbCrLf
        Else
            newXML = newXML & oldXML(i) & vbCrLf
        End If
    Next
    ies.XML = newXML
    
'=======EXECUTE THE SAVED IMPORT WITH THE NEW FILE LOCATION================================================================
    ies.Execute
    Set ies = Nothing





'DoCmd.RunSavedImportExport "Import BD Extract_1"
MsgBox ("Imported the combined START Extract.")




'DoCmd.RunSavedImportExport "Import-START EXTRACT"

'Dim strSQL As String
'   strSQL = "ALTER TABLE [START Account ID] ALTER COLUMN recip_client_id double"
'   DoCmd.RunSQL strSQL
   
      strSQL = "ALTER TABLE [START Account ID] ALTER COLUMN acct_id double"
   DoCmd.RunSQL strSQL
   
   
   

MsgBox ("Imported the START EXTRACT file into a new 'START Account ID' table.")


MsgBox ("The macro has finished running.  Please click OK.")




End Sub

Private Sub Z_Delete_AR_Click()

CarryOn = MsgBox("Run query:  'Z Delete AR' ?  This deletes the AR Table.", vbYesNo, "Continue?")
If CarryOn = vbNo Then
    Exit Sub
End If

DoCmd.OpenQuery "Z Delete AR"
MsgBox ("Ran Query 'Z Delete AR'.")



End Sub
