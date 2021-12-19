VERSION 5.00
Begin VB.Form frmRptStaffInfoAll 
   Appearance      =   0  'Flat
   BackColor       =   &H8000000D&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "View Hostels Report."
   ClientHeight    =   1335
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5925
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1335
   ScaleWidth      =   5925
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CmdView 
      Caption         =   "View"
      Height          =   375
      Left            =   2640
      TabIndex        =   1
      Top             =   720
      Width           =   1695
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   240
      Width           =   4455
   End
   Begin VB.Image ImgPhoto 
      Height          =   975
      Left            =   4680
      Stretch         =   -1  'True
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "frmRptStaffInfoAll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As New Recordset

Private Sub CmdView_Click()
    If Combo1.Text = "" Then
        MsgBox "Please Select a Staff ID or Employee ID!", vbCritical
        Exit Sub
    End If
    
    With DataEnvironment1.rsCmdStaffInfo
        If .State Then
            .Close
        End If
    End With
    DataEnvironment1.CmdStaffInfo Combo1
    DrptStaffInfo.Show
    
'    DataEnvironment1.Command1 Combo1.Text
'    DrptStudentInfo.Show

End Sub

Private Sub Combo1_Click()
Dim rpt As DataReport
        If rs.State = 1 Then rs.Close
        rs.Open "Select EmpPhoto,EmpPhotoSize from tblStaffDetails where EmpID= '" & Combo1 & "'", Con, 2, 3
'        Call FillPhoto1(rs, "StuPhoto", "StuPhotoSize", rpt.Sections("Section1").Controls("ImgPhoto"))
'        rpt.Sections("Section1").Controls("ImgPhoto").Picture = FillPhoto1(rs, "StuPhoto", "StuPhotoSize", ImgPhoto)
          FillPhoto rs, "EmpPhoto", "EmpPhotoSize", ImgPhoto
'        MsgBox ImgPhoto

End Sub

Private Sub Form_Load()
    Me.Top = 100
    Me.Left = 100
    sSQL = "Select EmpID From tblStaffDetails"
    Set rs = Con.Execute(sSQL)
    Combo1.Clear
    With rs
        Do While Not .EOF
            Combo1.AddItem .Fields(0)
            .MoveNext
        Loop
    End With
End Sub
