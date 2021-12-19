Attribute VB_Name = "MainMod"
Option Explicit

Public rsMaxId As New ADODB.Recordset
Public rstFindPhoto As New ADODB.Recordset

Public max_width As Integer
Public MaxID As Integer

Public ctr As Control
Public flagLockText As Boolean
Global sSQL As String

Global Opt As Integer
'
'Public Sub open_connection()
'If CN.State = 1 Then CN.Close
'With CN
'.Provider = "microsoft.jet.oledb.4.0"
'.ConnectionString = App.Path & "\Employee_DB.mdb"
'.CursorLocation = adUseClient
'.Open
'End With
'End Sub
'
Public Function LockText(frm As Form)

For Each ctr In frm.Controls
    If TypeOf ctr Is TextBox Then
        ctr.Locked = True
    ElseIf TypeOf ctr Is ComboBox Then
        ctr.Locked = True
    ElseIf TypeOf ctr Is MaskEdBox Then
        ctr.Enabled = False
    End If
Next

flagLockText = True
End Function


Public Function UnLockText(frm As Form)

For Each ctr In frm.Controls
    
    If TypeOf ctr Is TextBox Then
        ctr.Locked = False
    ElseIf TypeOf ctr Is ComboBox Then
        ctr.Locked = False
    ElseIf TypeOf ctr Is MaskEdBox Then
        ctr.Enabled = True
    End If

Next
    flagLockText = False
End Function

Public Function ClearTextBox(frm As Form)

For Each ctr In frm.Controls
    If TypeOf ctr Is TextBox Then
        ctr.Text = ""
    ElseIf TypeOf ctr Is MaskEdBox Then
       ctr.Enabled = ""
    End If
Next

End Function

Public Function EnableText(frm As Form)

For Each ctr In frm.Controls
    If TypeOf ctr Is TextBox Then
        ctr.Enabled = True
    End If
Next

flagLockText = True

End Function

Public Function DisableText(frm As Form)

For Each ctr In frm.Controls
    
    If TypeOf ctr Is TextBox Then
        ctr.Enabled = False
   ElseIf TypeOf ctr Is ComboBox Then
       ctr.Locked = False
    End If

Next
    flagLockText = False
End Function

Public Sub num_only(KeyAscii As Integer)
    Select Case KeyAscii
            Case Asc("0") To Asc("9")
            Case 8
            Case 13
           ' Case 45
            Case Else
                KeyAscii = 0
    End Select
End Sub
