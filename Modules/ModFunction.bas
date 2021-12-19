Attribute VB_Name = "Module1"
Public Con As New ADODB.Connection
Public rs As New ADODB.Recordset
Public rptCon As New ADODB.Connection

Public rsview As New ADODB.Recordset
Public rsBranch As New ADODB.Recordset
Public rsFees As New ADODB.Recordset
Public rsStud As New ADODB.Recordset


Dim sSQL As String

Public Function Connect()
    If Con.State = 1 Then Con.Close
    With Con
        .Provider = "Microsoft.Jet.OleDB.4.0"
        .ConnectionString = App.Path & "\Database\HostelMgmt.mdb"
        .CursorLocation = adUseClient
        .Open
'        MsgBox "Connection Established Successfully", vbInformation, "Connection"
    End With
End Function


Public Function link1()
    If rptCon.State = 1 Then rptCon.Close
    With DataEnvironment1.conYCT
        .Provider = "Microsoft.Jet.OleDB.4.0"
        .ConnectionString = App.Path & "\Database\HostelMgmt.mdb"
        .Open
        MsgBox "Connection Established Successfully", vbInformation, "Connection"
    End With
End Function

