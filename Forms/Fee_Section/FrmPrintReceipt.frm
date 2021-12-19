VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmPrintReceipt 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ADD DEPARTMENT FORM"
   ClientHeight    =   5715
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   13005
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5715
   ScaleWidth      =   13005
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   4335
      Left            =   -120
      TabIndex        =   1
      Top             =   1080
      Width           =   12855
      Begin VB.Frame Frame2 
         BackColor       =   &H00000000&
         Caption         =   "ADD DEPARTMENT"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   3375
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   6735
         Begin VB.CommandButton CmdFind 
            Caption         =   "Find"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   5400
            TabIndex        =   12
            Top             =   1440
            Width           =   1095
         End
         Begin VB.TextBox TxtReceiptId 
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   3240
            TabIndex        =   8
            ToolTipText     =   "Deparment ID"
            Top             =   480
            Width           =   3255
         End
         Begin VB.TextBox TxtEnrollment 
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   3240
            TabIndex        =   7
            ToolTipText     =   "Department Code"
            Top             =   1440
            Width           =   2055
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H00E0E0E0&
            Height          =   1095
            Left            =   2520
            TabIndex        =   5
            Top             =   2160
            Width           =   3975
            Begin VB.CommandButton CmdPrint 
               Caption         =   "PRINT"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   735
               Left            =   240
               TabIndex        =   11
               Top             =   240
               Width           =   1695
            End
            Begin VB.CommandButton CmdClose 
               Caption         =   "CLOSE"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   735
               Left            =   2160
               TabIndex        =   6
               Top             =   240
               Width           =   1575
            End
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "OR"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   4320
            TabIndex        =   13
            Top             =   1080
            Width           =   375
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "ENTER RECIEPT - ID "
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   480
            TabIndex        =   10
            Top             =   585
            Width           =   2595
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "ENROLLMENT NO."
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   480
            TabIndex        =   9
            Top             =   1545
            Width           =   2355
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00000000&
         Caption         =   "DETAIL"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   3375
         Left            =   7440
         TabIndex        =   2
         Top             =   360
         Width           =   5175
         Begin MSFlexGridLib.MSFlexGrid MSReceiptlist 
            Height          =   2535
            Left            =   240
            TabIndex        =   3
            Top             =   600
            Width           =   4695
            _ExtentX        =   8281
            _ExtentY        =   4471
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Line Line1 
            X1              =   2160
            X2              =   2280
            Y1              =   1920
            Y2              =   2040
         End
      End
      Begin VB.Line Line2 
         BorderWidth     =   5
         X1              =   90
         X2              =   12800
         Y1              =   3975
         Y2              =   3975
      End
      Begin VB.Line Line3 
         BorderWidth     =   3
         X1              =   7200
         X2              =   7200
         Y1              =   0
         Y2              =   4000
      End
   End
   Begin VB.Label Label5 
      BackColor       =   &H00000000&
      Caption         =   "PRINT RECEIPT "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   12135
   End
End
Attribute VB_Name = "FrmPrintReceipt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsview As New ADODB.Recordset
Dim Validate As Boolean
Dim MaxID As Integer

Private Sub CmdAddNew_Click()
Call UnLockText(Me)
Call ClearTextBox(Me)

CmdSave.Visible = True
CmdSave.Enabled = True

CmdAddNew.Enabled = False
CmdAddNew.Visible = False

Call fetchMaxID
TxtDeptId.Text = MaxID


End Sub


Private Sub CmdClose_Click()
Unload Me
End Sub


Private Sub CmdSave_Click()
Call chkValidate

If Validate = True Then
    If rs.State = 1 Then rs.Close
    rs.Open "Select * From tbldepartment", Con, 1, 3
    rs.AddNew
    rs.Fields(1) = TxtDeptCode.Text
    rs.Fields(2) = TxtDeptName.Text
    rs.Update
    MsgBox "NEW DEPARTMENT ADD"
    Call ShowData
    
CmdSave.Enabled = False
CmdEdit.Enabled = False

CmdAddNew.Enabled = True
CmdAddNew.Visible = True
     
    Call ClearTextBox(Me)
    Call LockText(Me)

End If

End Sub

Private Sub CmdUpdate_Click()
If rs.State = 1 Then rs.Close
    rs.Open "Select * From tbldepartment WHERE deptid=" & TxtDeptId.Text & "", Con, 1, 3
    rs.Fields(1) = TxtDeptCode.Text
    rs.Fields(2) = TxtDeptName.Text
    rs.Update
    MsgBox "UPDATION IS SUCCESFULY...."
    Call ShowData
    Call ClearTextBox(Me)
End Sub

Private Sub CmdFind_Click()
    If rs.State = 1 Then rs.Close
    rs.Open "Select * From tblPayReceipt WHERE EnrollNo='" & TxtEnrollment.Text & "'", Con, 1, 3
    If rs.RecordCount >= 1 Then
        MsgBox "RECORD FOUND IS SUCCESFULY...."
        Call ShowSearchData
    Else
        MsgBox "Enrollment Not Found"
    End If
End Sub

Private Sub CmdPrint_Click()
    If rs.State = 1 Then rs.Close
    rs.Open "Select * From tblPayReceipt WHERE ReceiptID=" & Val(TxtReceiptId.Text) & "", Con, 1, 3
        If rs.RecordCount >= 1 Then
            TxtReceiptId.Text = rs.Fields(2)
            MsgBox "RECORD FOUND IS SUCCESFULLY....."
        Else
            MsgBox "Reciept Enrollment Not Found"
        End If
End Sub

Private Sub Form_Load()
    Me.Top = 100
    Me.Left = 100

    CmdClose.Enabled = True
    
    Call ShowData
    
    Call fetchMaxID

End Sub

Private Sub ShowData()
MSReceiptlist.Clear
Dim i As Integer
MSReceiptlist.Cols = 3
MSReceiptlist.Rows = 1

MSReceiptlist.TextMatrix(0, 0) = "Receipt No."
MSReceiptlist.ColWidth(0) = max_width + 1350
MSReceiptlist.TextMatrix(0, 1) = "Enrollment No. "
MSReceiptlist.ColWidth(1) = max_width + 1500
MSReceiptlist.TextMatrix(0, 2) = "Student Name"
MSReceiptlist.ColWidth(2) = max_width + 2500

'MSDeptlist.TextMatrix(0, 0) = "Dept ID"
'MSDeptlist.TextMatrix(0, 1) = "Dept Code"
'MSDeptlist.TextMatrix(0, 2) = "Description"

If rsview.State = 1 Then rsview.Close
rsview.Open "Select * from tblPayReceipt", Con, adOpenForwardOnly, adLockReadOnly
MSReceiptlist.Rows = rsview.RecordCount + 1
For i = 1 To rsview.RecordCount
MSReceiptlist.TextMatrix(i, 0) = rsview.Fields(0).Value
MSReceiptlist.TextMatrix(i, 1) = rsview.Fields(2).Value
MSReceiptlist.TextMatrix(i, 2) = rsview.Fields(3).Value
rsview.MoveNext
Next i

End Sub


Private Sub fetchMaxID()
    If rs.State = 1 Then rs.Close
    rs.Open "Select Max(DeptID) From tbldepartment", Con, 2, 3

    If IsNull(rsMaxId(0)) Then
        MaxID = 1
    Else
        MaxID = rsMaxId.Fields(0) + 1
    End If

'For i = 0 To rs.RecordCount - 1
'  CmbdeptName.AddItem (rs.Fields(0).Value)
'   rs.MoveNext
'Next i

End Sub

'Private Sub TxtDeptId_LostFocus()
'TxtDeptId.Text = UCase(TxtDeptId)
'End Sub

Private Sub chkValidate()

If TxtDeptCode.Text = Trim("") Then
    MsgBox "Please Fill Correct Department Code"
    TxtDeptCode.SetFocus
    Validate = False
ElseIf TxtDeptName.Text = Trim("") Then
    MsgBox "Please Fill Correct Department Name"
    TxtDeptName.SetFocus
    Validate = False
Else
    Validate = True
End If

End Sub

Private Sub MSReceiptlist_Click()

TxtReceiptId.Text = MSReceiptlist.TextMatrix(MSReceiptlist.RowSel, 0)
TxtEnrollment.Text = MSReceiptlist.TextMatrix(MSReceiptlist.RowSel, 1)

End Sub

Private Sub MSDeptlist_DblClick()
'Dim frm As New FrmAddDesignation
'frm.TxtDesignID.Text = MSDeptlist.TextMatrix(MSDeptlist.RowSel, 0)
'frm.TxtDesignationName.Text = MSDeptlist.TextMatrix(MSDeptlist.RowSel, 2)
'frm.Show
End Sub


Private Sub ShowSearchData()
MSReceiptlist.Refresh
Dim i As Integer
MSReceiptlist.Cols = 3
MSReceiptlist.Rows = 1

MSReceiptlist.TextMatrix(0, 0) = "Receipt No."
MSReceiptlist.ColWidth(0) = max_width + 1350
MSReceiptlist.TextMatrix(0, 1) = "Enrollment No. "
MSReceiptlist.ColWidth(1) = max_width + 1500
MSReceiptlist.TextMatrix(0, 2) = "Student Name"
MSReceiptlist.ColWidth(2) = max_width + 2500

'If rsview.State = 1 Then rsview.Close
'rsview.Open "Select * from tblPayReceipt", con, adOpenForwardOnly, adLockReadOnly
MSReceiptlist.Rows = rs.RecordCount + 1
For i = 1 To rs.RecordCount
MSReceiptlist.TextMatrix(i, 0) = rs.Fields(0).Value
MSReceiptlist.TextMatrix(i, 1) = rs.Fields(2).Value
MSReceiptlist.TextMatrix(i, 2) = rs.Fields(3).Value
rs.MoveNext
Next i

End Sub

