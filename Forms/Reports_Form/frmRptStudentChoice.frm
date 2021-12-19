VERSION 5.00
Begin VB.Form frmRptStudentChoice 
   BackColor       =   &H8000000D&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "View Student Report."
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
Attribute VB_Name = "frmRptStudentChoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


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

    With DataEnvironment1.rsCmdStudentInfoChoice
        If .State Then
            .Close
        End If
    End With
    DataEnvironment1.CmdStudentInfoChoice Combo1, Combo1, Combo1, Combo1
    DrptStudentInfoChoice.Show
       
End Sub
