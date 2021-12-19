VERSION 5.00
Begin VB.Form FrmAddHostel 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Create Hostal Form"
   ClientHeight    =   7830
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12525
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "FrmAddHostel.frx":0000
   ScaleHeight     =   7830
   ScaleWidth      =   12525
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6615
      Left            =   240
      TabIndex        =   12
      Top             =   960
      Width           =   12015
      Begin VB.Frame Frame2 
         BackColor       =   &H00000000&
         Caption         =   "ADD"
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
         Height          =   5895
         Left            =   240
         TabIndex        =   13
         Top             =   480
         Width           =   11535
         Begin VB.TextBox TxtHostalPrefix 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   9000
            TabIndex        =   2
            Top             =   600
            Width           =   2295
         End
         Begin VB.TextBox TxtNickName 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   2280
            TabIndex        =   3
            Top             =   1320
            Width           =   4200
         End
         Begin VB.ComboBox CmbHostalType 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            ItemData        =   "FrmAddHostel.frx":103B
            Left            =   9000
            List            =   "FrmAddHostel.frx":1048
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   1320
            Width           =   2295
         End
         Begin VB.TextBox TxtHostalName 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   2280
            TabIndex        =   1
            Top             =   600
            Width           =   4215
         End
         Begin VB.TextBox TxtAddress 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   2280
            TabIndex        =   5
            Top             =   2040
            Width           =   4215
         End
         Begin VB.TextBox TxtMobileNo 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   9000
            MaxLength       =   10
            TabIndex        =   6
            Top             =   2160
            Width           =   2295
         End
         Begin VB.TextBox TxtStuCapacity 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   2280
            TabIndex        =   7
            Text            =   "0"
            Top             =   3120
            Width           =   4215
         End
         Begin VB.TextBox TxtNoOfRoom 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   9000
            TabIndex        =   8
            Text            =   "0"
            Top             =   3120
            Width           =   2295
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H8000000C&
            Height          =   1215
            Left            =   2280
            TabIndex        =   14
            Top             =   4200
            Width           =   8055
            Begin VB.CommandButton CmdSave 
               Caption         =   "SAVE"
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   615
               Left            =   2280
               TabIndex        =   9
               Top             =   360
               Width           =   1575
            End
            Begin VB.CommandButton CmdCancel 
               Caption         =   "CANCEL"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   615
               Left            =   4200
               TabIndex        =   10
               Top             =   360
               Width           =   1575
            End
            Begin VB.CommandButton CmdAddNew 
               Caption         =   "ADD NEW"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   615
               Left            =   360
               TabIndex        =   0
               Top             =   360
               Width           =   1575
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
               Height          =   615
               Left            =   6120
               TabIndex        =   11
               Top             =   360
               Width           =   1575
            End
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "NICK NAME"
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
            Left            =   240
            TabIndex        =   22
            Top             =   1440
            Width           =   1515
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "HOSTEL PREFIX"
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
            Left            =   6840
            TabIndex        =   21
            Top             =   705
            Width           =   2130
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "HOSTEL NAME"
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
            Left            =   240
            TabIndex        =   20
            Top             =   705
            Width           =   1890
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ADDRESS"
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
            Left            =   240
            TabIndex        =   19
            Top             =   2325
            Width           =   1215
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "MOBILE NO"
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
            Left            =   6840
            TabIndex        =   18
            Top             =   2265
            Width           =   1515
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "NO OF ROOMS"
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
            Left            =   6840
            TabIndex        =   17
            Top             =   3240
            Width           =   1965
         End
         Begin VB.Label Label7 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "STUDENT CAPACITY"
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
            Height          =   525
            Left            =   240
            TabIndex        =   16
            Top             =   3098
            Width           =   1380
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "HOSTEL TYPE"
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
            Left            =   6840
            TabIndex        =   15
            Top             =   1395
            Width           =   1875
         End
      End
   End
   Begin VB.Label Label5 
      BackColor       =   &H00000000&
      Caption         =   "ADD HOSTEL"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   360
      TabIndex        =   23
      Top             =   240
      Width           =   11775
   End
End
Attribute VB_Name = "FrmAddHostel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Validate, res As Boolean
Dim rsHostel As New ADODB.Recordset

Private Sub CmbHostalType_Click()
    TxtHostalPrefix = UCase(Left$(TxtHostalName, 1)) & UCase(Left$(CmbHostalType, 1))
End Sub

Private Sub CmdAddNew_Click()
    CmdSave.Enabled = True
    CmdAddNew.Enabled = False
    UnLockText Me
End Sub

Private Sub CmdCancel_Click()
    CmdSave.Enabled = False
    CmdAddNew.Enabled = True
End Sub

Private Sub CmdClose_Click()
    Unload Me
End Sub

Private Sub CmdSave_Click()
'    On Error GoTo ErrorHandler
'    TxtHostalPrefix = UCase(Left$(TxtHostalName, 1)) & UCase(Left$(CmbHostalType, 1))
'    Exit Sub

    Call chkValidate
    If Validate = True Then
        If rsHostel.State = 1 Then rsHostel.Close
'            sSQL = "Insert into tblHostelDetail(HostelName, Gender, HostelNickname,  Prefix, Address, MobNo, Capacity, CapacityUsed) Values ('" & TxtHostalName & "','" & CmbHostalType & "','" & TxtNickName & "','" & TxtHostalPrefix & "','" & TxtAddress & "','" & TxtMobileNo & "', " & CInt(TxtStuCapacity) & ",0)"
'        rs.Open sSQL, con, 2, 3
       
        rsHostel.Open "Select * from tblHostelDetail", Con, 2, 3
        rsHostel.AddNew
        rsHostel.Fields("HostelName").Value = TxtHostalName.Text
        rsHostel.Fields("Gender").Value = CmbHostalType.Text
        rsHostel.Fields("HostelNickname").Value = TxtNickName.Text
        rsHostel.Fields("Prefix").Value = TxtHostalPrefix.Text
        rsHostel.Fields("Address").Value = TxtAddress.Text
        rsHostel.Fields("MobNo").Value = Val(TxtMobileNo.Text)
'        rsHostel.Fields("TotalBlocks").Value = TxtNoOfBlocks.Text
        rsHostel.Fields("TotalRooms").Value = TxtNoOfRoom.Text
        rsHostel.Fields("Capacity").Value = TxtStuCapacity.Text
        rsHostel.Fields("CapacityUsed").Value = 0

        rsHostel.Update
        MsgBox "New Hostal Added Successfully", vbInformation

        CmdSave.Enabled = False
        CmdAddNew.Enabled = True
         
        Call ClearTextBox(Me)
        Call LockText(Me)
    End If
    Exit Sub
ErrorHandler:
'    cn.RollbackTrans
    MsgBox "Hostel " & TxtHostalName & " Not Created.", vbInformation

End Sub

Private Sub Form_Load()
    Me.Top = 100
    Me.Left = 100
    LockText Me
End Sub

Private Sub TxtHostalName_Change()
    TxtHostalPrefix = UCase(Left$(TxtHostalName, 1))
'    Command1.Caption = "Create Hostel '" & txtHostelName & "'"
End Sub

Private Sub chkValidate()

    If TxtHostalName.Text = Trim("") Then
        MsgBox "Please Fill Correct Hostal Name "
        TxtHostalName.SetFocus
        Validate = False
    ElseIf TxtNickName.Text = Trim("") Then
        MsgBox "Please Fill Correct Nick Name"
        TxtNickName.SetFocus
        Validate = False
    ElseIf TxtAddress.Text = Trim("") Then
        MsgBox "Please Fill Correct Address "
        TxtAddress.SetFocus
        Validate = False
    ElseIf TxtMobileNo.Text = Trim("") Then
        MsgBox "Please Fill Correct Mobile No. "
        TxtMobileNo.SetFocus
        Validate = False
    ElseIf CmbHostalType.Text = Trim("") Then
        MsgBox "Please Fill Correct Hostal Type "
        CmbHostalType.SetFocus
        Validate = False
    ElseIf TxtStuCapacity.Text = Trim("") Then
        MsgBox "Please Fill Correct Student Capacity "
        TxtStuCapacity.SetFocus
        Validate = False
        
    Else
        Validate = True
    End If

End Sub

Private Sub TxtMobileNo_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
            Case Asc("0") To Asc("9")
            Case 8
            Case 13
           ' Case 45
            Case Else
                KeyAscii = 0
        End Select
        
'        num_only (KeyAscii)
End Sub

Private Sub TxtNoOfRoom_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
            Case Asc("0") To Asc("9")
            Case 8
            Case 13
           ' Case 45
            Case Else
                KeyAscii = 0
        End Select
        
'        num_only (KeyAscii)
End Sub

Private Sub TxtStuCapacity_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
            Case Asc("0") To Asc("9")
            Case 8
            Case 13
           ' Case 45
            Case Else
                KeyAscii = 0
        End Select
        
'        num_only (KeyAscii)
End Sub

