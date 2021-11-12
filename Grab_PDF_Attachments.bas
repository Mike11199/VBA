Attribute VB_Name = "Grab_PDF_Attachments"
Sub GrabDocAttachments_ForNonNumberReferenceColumn()


On Error GoTo Error

Dim ORF_WB As Workbook
Set ORF_WB = ThisWorkbook

GLAccount = ORF_WB.Sheets("Macro Input").Range("GL_Account")
FiscalYear = ORF_WB.Sheets("Macro Input").Range("Fiscal_Year")
ReconMonth = ORF_WB.Sheets("Macro Input").Range("Recon_Month")
ReconMonth_Num = ORF_WB.Sheets("Macro Input").Range("ReconMonth_Num")
SapConnectionString = ORF_WB.Sheets("Macro Input").Range("SAP_Connection")

BD_LOG_ON.Show



Add_PDF_To_Excel_Sheets = MsgBox(Prompt:="This will download all PDF attachments to C:\TEMP." & vbNewLine & vbNewLine & "Do you also want to add each" & _
    " PDF to an Excel sheet as an object?" & vbNewLine & vbNewLine & "While this works, Excel can sometimes crash if opening a PDF from itself, and it's recommended to open" & _
     " PDFs from the C:\TEMP folder instead.", Buttons:=vbQuestion + vbYesNo)


Dim StartTime As Double
Dim MinutesElapsed As String
'Remember time when macro starts
StartTime = Timer





Set SAP_Application = CreateObject("Sapgui.ScriptingCtrl.1")


Rem Open a connection in synchronous mode
Set Connection = SAP_Application.OpenConnection(SapConnectionString, True)
Set SAP_Session = Connection.Children(0)



'=======================================================this code is for maximizing or minimizing the window==================
SessionHWND = SAP_Session.FindById("wnd[0]").Handle
ActivateWindow (SessionHWND)

SAP_Session.FindById("wnd[0]").Maximize


'=======================================================================================================================
        
BDUserName8 = BD_LOG_ON.BDUserBox.Value
BDPass8 = BD_LOG_ON.BDPasswordBox.Value

Application.Wait (Now + TimeValue("00:00:01") / 1000)
SAP_Session.FindById("wnd[0]/usr/txtRSYST-BNAME").Text = BDUserName8
Application.Wait (Now + TimeValue("00:00:01") / 1000)
SAP_Session.FindById("wnd[0]/usr/pwdRSYST-BCODE").Text = BDPass8
Application.Wait (Now + TimeValue("00:00:01") / 1000)
SAP_Session.FindById("wnd[0]").sendVKey 0
Application.Wait (Now + TimeValue("00:00:01") / 1000)


Unload BD_LOG_ON
BDUserName8 = ""
BDPass8 = ""

'========================================================================================================================

ActivateWindow (Application.hwnd)    'Brings Excel window to front

Dim ReconSheet As Worksheet
Set ReconSheet = ORF_WB.Sheets("1130_" & ReconMonth)

ReconSheet.Activate

lastrow2 = ActiveSheet.Cells(ActiveSheet.Rows.Count, "A").End(xlUp).row


Dim PeriodRange As Range
Set PeriodRange = ReconSheet.Range("A1:A" & lastrow2)

Dim c As Range
Dim ol As OLEObject
Dim p As Page



Set DotNetArray = CreateObject("System.Collections.ArrayList")



'Loop through current month REV FUND and CALATERS Cells
For Each c In PeriodRange.Cells
    If c.Value = "CM" Then
         If c.Offset(, 10).Value = "CALATERS" Or c.Offset(, 10).Value = "REV FUND" Then
                        
                        
                        
'========================This part, which includes an extra END IF before the next for 'each c in periodrange.cells' makes sure it doesn't grab the PDF for the same document twice=======================================================
                    Duplicate = 0
                    DocumentNumber = c.Offset(, 4).Value

                 
                                        If Not IsEmpty(DotNetArray) Then
                                                 For Each doc In DotNetArray
                                                         If doc = DocumentNumber Then
                                                                Duplicate = 1
                                                         End If
                                                    Next doc
                                         End If
                                                         
                                    
                    If Duplicate = 0 Then
                            DotNetArray.Add DocumentNumber
'================================================Make sure to delete extra 'END IF' if you get rid of this section===================================================================================================================
                     


                        
                            ActivateWindow (SessionHWND)

                            
                            SAP_Session.FindById("wnd[0]/tbar[0]/okcd").Text = "/nfb03"
                            SAP_Session.FindById("wnd[0]").sendVKey 0
                            SAP_Session.FindById("wnd[0]/usr/txtRF05L-BELNR").Text = DocumentNumber
                            SAP_Session.FindById("wnd[0]/usr/txtRF05L-GJAHR").Text = FiscalYear
                            SAP_Session.FindById("wnd[0]/usr/txtRF05L-GJAHR").SetFocus
                            SAP_Session.FindById("wnd[0]/usr/txtRF05L-GJAHR").CaretPosition = 4
                            SAP_Session.FindById("wnd[0]").sendVKey 0
                            SAP_Session.FindById("wnd[0]/usr/cntlCTRL_CONTAINERBSEG/shellcont/shell").CurrentCellColumn = "HKONT"
                            SAP_Session.FindById("wnd[0]/titl/shellcont/shell").PressContextButton "%GOS_TOOLBOX"
                            SAP_Session.FindById("wnd[0]/titl/shellcont/shell").SelectContextMenuItem "%GOS_VIEW_ATTA"
                            SAP_Session.FindById("wnd[1]/usr/cntlCONTAINER_0100/shellcont/shell").CurrentCellColumn = "BITM_FILENAME"
                            SAP_Session.FindById("wnd[1]/usr/cntlCONTAINER_0100/shellcont/shell").SelectedRows = "0"
                            SAP_Session.FindById("wnd[1]/usr/cntlCONTAINER_0100/shellcont/shell").PressToolbarButton "%ATTA_EXPORT"
                            SAP_Session.FindById("wnd[2]/usr/ctxtDY_PATH").Text = "C:\TEMP"
                            SAP_Session.FindById("wnd[2]/usr/ctxtDY_FILENAME").Text = DocumentNumber & ".pdf"
                            SAP_Session.FindById("wnd[2]/usr/ctxtDY_FILENAME").CaretPosition = 5
                            SAP_Session.FindById("wnd[2]/tbar[0]/btn[11]").press
                        
                            
                                                                        If Add_PDF_To_Excel_Sheets = vbYes Then
                                                                                        
                                                                                                        ActivateWindow (Application.hwnd) 'ExcelWBInFront
                                                                                                        
                                                                                                        
                                                                                                        Application.Wait (Now + TimeValue("00:00:01"))
                                                                                                        
                                                                                                        SheetNumberCount = ThisWorkbook.Sheets.Count
                                                                                                        
                                                                                                        Sheets.Add(After:=Sheets("PDFs -->")).Name = DocumentNumber & "_" & SheetNumberCount
                                                                                                        
                                                                                                        ActiveWorkbook.ActiveSheet.Tab.Color = 192
                                                                                                        
                                                                                                        
                                                                                                        
                                                                                                        
                                                                                                        
                                                                                                        
                                                                                                        Set ol = ActiveSheet.OLEObjects.Add(, "C:\TEMP\" & DocumentNumber & ".pdf")
                                                                                                        
                                                                                                        
                                                                                                        
                                                                                                        With ol
                                                                                                        .Left = ActiveSheet.Range("A1").Left
                                                                                                        .Height = ActiveSheet.Range("A1").Height
                                                                                                        .Width = ActiveSheet.Range("A1:i40").Width
                                                                                                        .Top = ActiveSheet.Range("A1").Top
                                                                                                        
                                                                                                        End With
                                                                                                        
                                                                                                        Application.Wait (Now + TimeValue("00:00:01"))
                                                                                                        
                                                                                                        
                                                                                                        
                                                                                                        c.Offset(, 10).Interior.ThemeColor = xlThemeColorAccent4
                                                                                                        c.Offset(, 10).Interior.TintAndShade = 0.399975585192419
                                                                                               
                                                                        End If       'End If for 'If add PDFS to Excel Sheets = Yes


           
            End If         'End If for 'If duplicate=0'
    End If                 'End If for 'If CALATERS or REV FUND
End If                     'End If for 'If CM'

Next c                  'Go to next document number for loop



ActivateWindow (Application.hwnd)    'Brings the Excel window to the front

Call Shell("explorer.exe" & " " & "C:\Temp", vbNormalFocus)

MinutesElapsed = Format((Timer - StartTime) / 86400, "hh:mm:ss")

MsgBox "Ran successfully in " & MinutesElapsed & " minutes." & vbNewLine & vbNewLine & _
"The macro has finished running.  It should have also opened the C:\Temp folder in file explorer for you, where the PDFs should be saved." & vbNewLine & vbNewLine & _
"Please click OK."



Exit Sub
Error:
MsgBox "Error.  Please press OK to end the macro."







End Sub
Public Function IsInArray(stringToBeFound As String, arr As Variant) As Boolean

'Checks if document number is in array (which means it was already used to grab a PDF attachment)

    Dim i
    For i = LBound(arr) To UBound(arr)
        If arr(i) = stringToBeFound Then
            IsInArray = True
            Exit Function
        End If
    Next i
    IsInArray = False

End Function
