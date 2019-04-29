VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmReqPasswordFrm 
   ClientHeight    =   2220
   ClientLeft      =   105
   ClientTop       =   435
   ClientWidth     =   5985
   OleObjectBlob   =   "frmReqPasswordFrm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmReqPasswordFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub confirmButton_Click()

    If (Me.pwTextBox.Value = "") Then
        MsgBox ("Password can not be an empty string.")
    Else
        Me.Hide
    End If
    
End Sub

Private Sub showCheckBox_Click()

    If Me.showCheckBox.Value = True Then
        Me.pwTextBox.PasswordChar = ""
    Else
        Me.pwTextBox.PasswordChar = "*"
    End If
    
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    
    If CloseMode = 0 Then
        Me.pwTextBox.Value = ""
    End If

End Sub

Private Sub cancelButton_Click()
    
    Me.pwTextBox.Value = ""
    Me.Hide
    
End Sub
