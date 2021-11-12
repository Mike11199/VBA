VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} BD_LOG_ON 
   Caption         =   "BusinessDirect Log On"
   ClientHeight    =   6180
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8355
   OleObjectBlob   =   "BD_LOG_ON.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "BD_LOG_ON"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CancelButton_Click()


Me.Hide
End

End Sub



Private Sub EnterButton_Click()





Me.Hide


End Sub

Private Sub TextBox1_Change()




End Sub

Private Sub TextBox2_Change()

End Sub

Private Sub SingleSignOnButton_Click()

SingleSignOnValue = 1
Me.Hide



End Sub

Public Sub UserForm_Initialize()

Me.StartUpPosition = 0
Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)

Dim ORF_WB As Workbook
Set ORF_WB = ThisWorkbook

DefaultUser = ORF_WB.Sheets("Macro Input").Range("Default_User")
BD_LOG_ON.BDUserBox.Value = DefaultUser


End Sub
