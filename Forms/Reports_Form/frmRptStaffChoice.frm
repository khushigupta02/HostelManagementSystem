VERSION 5.00
Begin VB.Form frmRptStaffChoice 
   BackColor       =   &H8000000D&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "View Staff Report."
   ClientHeight    =   1215
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4815
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1215
   ScaleWidth      =   4815
   Begin VB.TextBox TxtName 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Text            =   "SearchValue"
      Top             =   720
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   240
      Width           =   4455
   End
   Begin VB.CommandButton CmdView 
      Caption         =   "View"
      Height          =   375
      Left            =   2760
      TabIndex        =   0
      Top             =   720
      Width           =   1695
   End
End
Attribute VB_Name = "frmRptStaffChoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DetartName As Integer
Dim DesigName As Integer


'Private Sub Form_Load()
'    sSQL = "Select HostelName From tblHostelDetail"
'    Set rs = Con.Execute(sSQL)
'    Combo1.Clear
'    With rs
'        Do While Not .EOF
'            Combo1.AddItem .Fields(0)
'            .MoveNext
'        Loop
'    End With
'End Sub

Private Sub CmdView_Click()
    If Combo1.Text = "" Then
        MsgBox "Please Select a Correct Choice!", vbCritical
        Exit Sub
    End If

    With DataEnvironment1.rsCmdStaffInfoChoice
        If .State Then
            .Close
        End If
    End With
    
    DataEnvironment1.CmdStaffInfoChoice Combo1, DesigName, DetartName  ' TxtName, TxtName, TxtName  'DesigName, DetartName  'Combo1,
    DrptAllStaffInfoChoice.Show
       

'''SELECT tblStaffDetails.* FROM tblStaffDetails WHERE ( Gender = ? OR DesignationID = ? OR DeptID = ? )


End Sub



Private Sub Combo1_Click()

    If Opt = 1 Then
        If rs.State = 1 Then rs.Close
        rs.Open "Select DeptID from tblDepartment Where DeptName='" & Combo1 & "'", Con, 2, 3
        DetartName = rs.Fields("DeptID").Value
        TxtName = rs.Fields("DeptID").Value
    '    MsgBox rs.Fields("DeptID")
    ElseIf Opt = 2 Then
        If rs.State = 1 Then rs.Close
        rs.Open "Select DesignationID from tblDesignation Where DesignationName='" & Combo1 & "'", Con, 2, 3
        DesigName = rs.Fields("DesignationID").Value
    '    MsgBox DesigName
        TxtName = rs.Fields("DesignationID").Value
    Else
        TxtName = Combo1.Text
    End If

End Sub

Private Sub fetchName()
    If rs.State = 1 Then rs.Close
    rs.Open "Select DesignationName from tblDesignation where DesignationID= " & MSFStaFFDetails.TextMatrix(MSFStaFFDetails.RowSel, 9) & "", Con, 2, 3
    DesigName = rs.Fields(0).Value
    'MsgBox DesigName
    If rs.State = 1 Then rs.Close
    rs.Open "Select DeptName from tblDepartment Where DeptID IN (Select DeptID From tblStaffdetails Where StaffID= " & MSFStaFFDetails.TextMatrix(MSFStaFFDetails.RowSel, 0) & ")", Con, 2, 3
    DetartName = rs.Fields(0).Value
'    MsgBox DetartName

End Sub
