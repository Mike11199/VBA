VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Export to BC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Command8_Click()


CarryOn = MsgBox("This will copy tables to the 'Export to BC' database.  Make sure the tables listed on the form are closed.  Please select the 'Export to BC' database from the following window.   Continue with the export?", vbYesNo, "Continue?")
        If CarryOn = vbNo Then
            Exit Sub
        End If



Call Export_to_BC_Database




End Sub
Function Export_to_BC_Database()
On Error GoTo Export_to_BC_Database_Err

destination = "K:\TECHNOLOGY SERVICES\Pension Solution\CB RB Conversion\Export to BC.accdb"
   

Dim f As Object
Set f = Application.FileDialog(3)

f.AllowMultiSelect = True
f.Title = "SELECT ACCESS EXPORT BC DATABASE"
'f.InitialFileName = "K:\TECHNOLOGY SERVICES\Pension Solution\CB RB Conversion\"
f.InitialFileName = "\\MyFiles04\data\vol1\dptshare\TECHNOLOGY SERVICES\Pension Solution\CB RB Conversion\"


If f.Show Then
    For i = 1 To f.SelectedItems.Count
        destination2 = (f.SelectedItems(i))
    Next
End If

 DoCmd.TransferDatabase acExport, "Microsoft Access", destination2, acTable, "1042-S Annual", "1042-S Annual", False
 DoCmd.TransferDatabase acExport, "Microsoft Access", destination2, acTable, "1099 Annual", "1099 Annual", False
 DoCmd.TransferDatabase acExport, "Microsoft Access", destination2, acTable, "1099 MISC", "1099 MISC", False
 DoCmd.TransferDatabase acExport, "Microsoft Access", destination2, acTable, "1099 Payment", "1099 Payment", False
 DoCmd.TransferDatabase acExport, "Microsoft Access", destination2, acTable, "AR", "AR", False
 DoCmd.TransferDatabase acExport, "Microsoft Access", destination2, acTable, "Client Tax", "Client Tax", False
 DoCmd.TransferDatabase acExport, "Microsoft Access", destination2, acTable, "Invoice", "Invoice", False
 DoCmd.TransferDatabase acExport, "Microsoft Access", destination2, acTable, "Payment", "Payment", False
 DoCmd.TransferDatabase acExport, "Microsoft Access", destination2, acTable, "Vendor Tax", "Vendor Tax", False
 DoCmd.TransferDatabase acExport, "Microsoft Access", destination2, acTable, "W2 Annual", "W2 Annual", False
 DoCmd.TransferDatabase acExport, "Microsoft Access", destination2, acTable, "W2 Payment", "W2 Payment", False
 DoCmd.TransferDatabase acExport, "Microsoft Access", destination2, acTable, "Worksheet", "Worksheet", False
 
 
   

Export_to_BC_Database_Exit:
    Exit Function

Export_to_BC_Database_Err:
    MsgBox Error$
    Resume Export_to_BC_Database_Exit

End Function

