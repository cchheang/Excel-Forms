Attribute VB_Name = "frm"
'Form module

'Overview:
'Various forms for common tasks not available with default functions

'Available Functions:
    'frmReqPassword()
    'Requests password from user and returns string of input
    'Returns empty string if user cancels or closes window via "x" button
    'Custom userform allows masking of typed password which is not possible
    'with regular inputbox function
    

Option Explicit

Public Function frmReqPassword() As String
'Open custom form to request password from user
'and returns string of input entered by user
'Returns empty string if user cancels or closes window

'Requires: form frmReqPasswordFrm

    Dim reqPassword As frmReqPasswordFrm
    Set reqPassword = New frmReqPasswordFrm
    
    reqPassword.Show
    
    frmReqPassword = reqPassword.pwTextBox.Value
    
    Unload reqPassword

End Function
