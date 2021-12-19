VERSION 5.00
Begin VB.Form FrmMenu 
   Caption         =   "menu"
   ClientHeight    =   8670
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13710
   LinkTopic       =   "Form1"
   Picture         =   "FrmMenu.frx":0000
   ScaleHeight     =   8670
   ScaleWidth      =   13710
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Height          =   11175
      Left            =   16800
      ScaleHeight     =   11115
      ScaleWidth      =   2115
      TabIndex        =   25
      Top             =   -120
      Width           =   2175
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00000000&
      Height          =   3615
      Left            =   11280
      TabIndex        =   19
      Top             =   4200
      Width           =   2295
      Begin VB.CommandButton CmdPrintReceipt 
         Caption         =   "Print Receipt"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   26
         Top             =   2880
         Width           =   1815
      End
      Begin VB.CommandButton CmdFeeStructure 
         Caption         =   "Fee Structure"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   22
         Top             =   720
         Width           =   1815
      End
      Begin VB.CommandButton CmdFeeReceipt 
         Caption         =   "Paid  Fee Receipt"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   21
         Top             =   1440
         Width           =   1815
      End
      Begin VB.CommandButton CmdFeeRefund 
         Caption         =   "Fee Refund "
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   20
         Top             =   2160
         Width           =   1815
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Fees Section"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   240
         TabIndex        =   23
         Top             =   240
         Width           =   1665
      End
   End
   Begin VB.CommandButton CmdExit 
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3000
      TabIndex        =   14
      Top             =   6960
      Width           =   7575
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00000000&
      Height          =   3855
      Left            =   11280
      TabIndex        =   11
      Top             =   120
      Width           =   2295
      Begin VB.CommandButton CmdViewHostel 
         Caption         =   "View Hostel"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   18
         Top             =   3120
         Width           =   1815
      End
      Begin VB.CommandButton CmdAllocateRoom 
         Caption         =   "Allocate Room"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   17
         Top             =   2280
         Width           =   1815
      End
      Begin VB.CommandButton CmdAddRoom 
         Caption         =   "Add Room "
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   16
         Top             =   1440
         Width           =   1815
      End
      Begin VB.CommandButton CmdAddHostel 
         Caption         =   "Add Hostel"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   15
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Room Section"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   240
         TabIndex        =   12
         Top             =   240
         Width           =   1830
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00000000&
      Height          =   3495
      Left            =   120
      TabIndex        =   6
      Top             =   4320
      Width           =   2295
      Begin VB.CommandButton CmdStaffDetails 
         Caption         =   "Staff Details"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   24
         Top             =   2760
         Width           =   1815
      End
      Begin VB.CommandButton CmdStaffRegistration 
         Caption         =   "Staff Registration"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   10
         Top             =   2040
         Width           =   1815
      End
      Begin VB.CommandButton CmdAddDpt 
         Caption         =   "Add Department"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   9
         Top             =   600
         Width           =   1815
      End
      Begin VB.CommandButton CmdAddDesignation 
         Caption         =   "Add Designation"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   8
         Top             =   1320
         Width           =   1815
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Staff Section"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   240
         TabIndex        =   7
         Top             =   120
         Width           =   1680
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Height          =   3855
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   2415
      Begin VB.CommandButton CmdUpdateStudent 
         Caption         =   "UPDATE STUDENT"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   13
         Top             =   3120
         Width           =   1815
      End
      Begin VB.CommandButton CmdSearchStudent 
         Caption         =   "Search Student"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   5
         Top             =   2280
         Width           =   1815
      End
      Begin VB.CommandButton CmdStuRegistration 
         Caption         =   "Student Registration"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   4
         Top             =   1440
         Width           =   1815
      End
      Begin VB.CommandButton CmdAddBranch 
         Caption         =   "Add Branch"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   3
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Student Section"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   2025
      End
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      X1              =   0
      X2              =   13695
      Y1              =   7935
      Y2              =   7935
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      X1              =   10800
      X2              =   13680
      Y1              =   4080
      Y2              =   4080
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      X1              =   10800
      X2              =   10800
      Y1              =   0
      Y2              =   7920
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      X1              =   0
      X2              =   2760
      Y1              =   4200
      Y2              =   4200
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      X1              =   2760
      X2              =   2760
      Y1              =   -120
      Y2              =   7920
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "HOSTEL MANAGEMENT SYSTEM"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   540
      Left            =   2880
      TabIndex        =   0
      Top             =   360
      Width           =   7800
   End
End
Attribute VB_Name = "FrmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CmbFeeRefund_Click()
FrmRefundFees.Show
End Sub

Private Sub CmdAddBranch_Click()
FrmAddBranch.Show
End Sub

Private Sub CmdAddDesignation_Click()
FrmAddDesignation.Show

End Sub

Private Sub CmdAddDpt_Click()
FrmAddDepartment.Show
End Sub

Private Sub CmdAddHostel_Click()
FrmAddHostel.Show
End Sub

Private Sub CmdAddRoom_Click()
FrmAddRoomDetails.Show
End Sub

Private Sub CmdAllocateRoom_Click()
FrmAllocateRoom.Show
End Sub

Private Sub CmdExit_Click()
End
End Sub

Private Sub CmdFeeReceipt_Click()
FrmFeeReciept.Show
End Sub

Private Sub CmdFeeRefund_Click()
FrmRefundFees.Show

End Sub

Private Sub CmdFeeStructure_Click()
FrmFeeStructure.Show
End Sub

Private Sub CmdPrintReceipt_Click()
''Call link1
'    DataEnvironment1.Command7_Grouping
'    DataReport4.Show
''    MsgBox DataReport1.Width

frmRptHostel.Show
End Sub

Private Sub CmdSearchStudent_Click()
FrmSearchStudent.Show
End Sub

Private Sub CmdStaffDetails_Click()
FrmViewStaffDetails.Show

End Sub

Private Sub CmdStaffRegistration_Click()
FrmStaffRegistration.Show
End Sub

Private Sub CmdStuRegistration_Click()
FrmStudentRegistration.Show
End Sub

Private Sub CmdUpdateStudent_Click()
FrmUpdateStuDetails.Show
End Sub

Private Sub CmdViewHostel_Click()
FrmDeAllocateRoom.Show
'FrmViewHostelDetail.Show
End Sub


Private Sub Form_Load()
    Me.Top = 100
    Me.Left = 100
    Call Connect
    
End Sub

