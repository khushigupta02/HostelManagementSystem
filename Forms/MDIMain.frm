VERSION 5.00
Begin VB.MDIForm MDIMain 
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   8250
   ClientLeft      =   165
   ClientTop       =   810
   ClientWidth     =   11850
   Icon            =   "MDIMain.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "MDIMain.frx":6852
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture2 
      Align           =   4  'Align Right
      Height          =   8250
      Left            =   9405
      Picture         =   "MDIMain.frx":32F42
      ScaleHeight     =   8190
      ScaleWidth      =   2385
      TabIndex        =   14
      Top             =   0
      Width           =   2450
      Begin VB.CommandButton CmdExit1 
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
         Height          =   735
         Left            =   120
         TabIndex        =   20
         Top             =   5040
         Width           =   2175
      End
      Begin VB.CommandButton CmdReports 
         Caption         =   "Reports"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   19
         Top             =   4080
         Width           =   2175
      End
      Begin VB.CommandButton CmdFees 
         Caption         =   "Fees"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   18
         Top             =   3120
         Width           =   2175
      End
      Begin VB.CommandButton CmdHostel 
         Caption         =   "Hostel"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   17
         Top             =   2160
         Width           =   2175
      End
      Begin VB.CommandButton CmdStaff 
         Caption         =   "Staff"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   16
         Top             =   1320
         Width           =   2175
      End
      Begin VB.CommandButton CmdStudent 
         Caption         =   "Student"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   15
         Top             =   480
         Width           =   2175
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   3  'Align Left
      Height          =   8250
      Left            =   0
      Picture         =   "MDIMain.frx":3FDF24
      ScaleHeight     =   8190
      ScaleWidth      =   2625
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   2685
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
         Height          =   735
         Left            =   240
         TabIndex        =   13
         Top             =   7440
         Width           =   2175
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00000000&
         Height          =   3495
         Left            =   120
         TabIndex        =   7
         Top             =   3960
         Width           =   2415
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
            TabIndex        =   11
            Top             =   1320
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
            TabIndex        =   10
            Top             =   600
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
            TabIndex        =   9
            Top             =   2040
            Width           =   1815
         End
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
            TabIndex        =   8
            Top             =   2760
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
            TabIndex        =   12
            Top             =   120
            Width           =   1680
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00000000&
         Height          =   3855
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   2415
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
            TabIndex        =   5
            Top             =   720
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
            TabIndex        =   3
            Top             =   2280
            Width           =   1815
         End
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
            TabIndex        =   2
            Top             =   3120
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
            TabIndex        =   6
            Top             =   240
            Width           =   2025
         End
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu smnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuManagement 
      Caption         =   "Management"
      Begin VB.Menu smnuStudent 
         Caption         =   "Student"
         Begin VB.Menu smnuStudentRegistration 
            Caption         =   "Student Registration"
         End
         Begin VB.Menu smnuUpdateStudent 
            Caption         =   "Update Student"
         End
         Begin VB.Menu smnuSearchStudent 
            Caption         =   "Search Student"
         End
         Begin VB.Menu smnuViewStudentRoomDetails 
            Caption         =   "View Student Room Details"
         End
         Begin VB.Menu smnuViewStudentFeeDetails 
            Caption         =   "View Student Fee Details"
         End
      End
      Begin VB.Menu mhypn1 
         Caption         =   "-"
      End
      Begin VB.Menu smnuStaff 
         Caption         =   "Staff"
         Begin VB.Menu smnuStaffRegistration 
            Caption         =   "Staff Registration"
         End
         Begin VB.Menu smnuStaffDetail 
            Caption         =   "View Staff Details"
         End
      End
      Begin VB.Menu mhypn2 
         Caption         =   "-"
      End
      Begin VB.Menu smnuFees 
         Caption         =   "Fees"
         Begin VB.Menu smnuCreateFeeStructure 
            Caption         =   "Create Fee Structure"
         End
         Begin VB.Menu smnuPaidFee 
            Caption         =   "Paid Fee"
         End
         Begin VB.Menu smnuFeeRefund 
            Caption         =   "Fee Refund"
         End
         Begin VB.Menu smnuPrintFeeReciept 
            Caption         =   "Print Fee Reciept"
         End
      End
      Begin VB.Menu mhypn3 
         Caption         =   "-"
      End
      Begin VB.Menu smnuHostel 
         Caption         =   "Hostel"
         Begin VB.Menu smnuCreateNewHostel 
            Caption         =   "Create / Add New Hostel"
         End
         Begin VB.Menu smnuViewHostel 
            Caption         =   "View Hostel"
         End
         Begin VB.Menu ssmnuRooms 
            Caption         =   "Rooms"
         End
         Begin VB.Menu smnuAllocateRoom 
            Caption         =   "Allocate Room"
         End
         Begin VB.Menu smnuDeAllocateRoom 
            Caption         =   "DeAllocate Room"
         End
      End
   End
   Begin VB.Menu mnuAllocation 
      Caption         =   "Allocation"
      Begin VB.Menu smnuAllocate 
         Caption         =   "Allocate"
      End
   End
   Begin VB.Menu mnuSystem 
      Caption         =   "System"
      Begin VB.Menu ssmnuBranch 
         Caption         =   "Add New Branch"
      End
      Begin VB.Menu ssmnuDepartment 
         Caption         =   "Add New Department"
      End
      Begin VB.Menu ssmnuDesignation 
         Caption         =   "Add New Designation"
      End
   End
   Begin VB.Menu mnuReports 
      Caption         =   "Reports"
      Begin VB.Menu mnuStudents 
         Caption         =   "Students"
         Begin VB.Menu smnuAllStudentDetails 
            Caption         =   "All Student Details"
         End
         Begin VB.Menu smnuStudentRegDetails 
            Caption         =   "Student Registration Details"
         End
         Begin VB.Menu smnuYearWise 
            Caption         =   "Semester / Year Wise"
         End
         Begin VB.Menu smnuStaffGenderWise 
            Caption         =   "Students List by Gender Wise"
         End
         Begin VB.Menu smnuBranchWise 
            Caption         =   "Students List by Branch Wise"
         End
         Begin VB.Menu smnuCategoryWise 
            Caption         =   "Students List by Category Wise"
         End
         Begin VB.Menu smnuComingSoon 
            Caption         =   "Coming Soon"
         End
      End
      Begin VB.Menu rh1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuStaff 
         Caption         =   "Staff"
         Begin VB.Menu smnuAllStaffDetails 
            Caption         =   "All Staff Details"
         End
         Begin VB.Menu smnuStaffDetails 
            Caption         =   "Staff Details"
         End
         Begin VB.Menu smnuDeptwise 
            Caption         =   "Staff List by Dept Wise"
         End
         Begin VB.Menu smnuDesignationWise 
            Caption         =   "Staff List by Designation Wise"
         End
         Begin VB.Menu smnuGenderWise 
            Caption         =   "Staff List by Gender Wise"
         End
      End
      Begin VB.Menu rh2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFees 
         Caption         =   "Fees"
         Begin VB.Menu smnuFeesStructureDetails 
            Caption         =   "Fees Structure Details"
         End
         Begin VB.Menu smnuPaidFeesDetails 
            Caption         =   "Paid Fees Details"
         End
         Begin VB.Menu smnuFeesReciept 
            Caption         =   "Fees Reciept"
         End
         Begin VB.Menu smnuRefundFeesDetails 
            Caption         =   "Refund Fees Details"
         End
      End
      Begin VB.Menu rh3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHostelRoom 
         Caption         =   "Hostel Room"
         Begin VB.Menu smnuHostelDetails 
            Caption         =   "Hostel Details"
         End
         Begin VB.Menu smnuHostelRoomDetails 
            Caption         =   "Hostel Room Details"
         End
         Begin VB.Menu smnuAllAllocatedRooms 
            Caption         =   "All Allocated Rooms"
         End
         Begin VB.Menu smnuAllUnallocatedRooms 
            Caption         =   "All Unallocated Rooms"
         End
         Begin VB.Menu smnuFullyAllocatedRooms 
            Caption         =   "Fully Allocated Rooms"
         End
         Begin VB.Menu smnuPartiallyAllocatedRooms 
            Caption         =   "Partially Allocated Rooms"
         End
      End
      Begin VB.Menu smnuBranch 
         Caption         =   "Branch List"
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "Window"
      Begin VB.Menu mnuCascade 
         Caption         =   "ReArrange"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu smnuAbout 
         Caption         =   "About"
      End
      Begin VB.Menu smnuDocumentation 
         Caption         =   "Documentation"
      End
   End
End
Attribute VB_Name = "MDIMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



'''''
'''''Private Sub CmdSearchStudent_Click()
'''''    frmRptHostel.Show
''''''    dRptHostel.Show
'''''End Sub
'''''
'''''Private Sub CmdStaffDetails_Click()
'''''    DataReport6.Show
'''''End Sub
'''''
'''''Private Sub CmdStaffRegistration_Click()
'''''    DataReport5.Show
'''''
'''''End Sub
'''''
'''''Private Sub CmdStuRegistration_Click()
'''''    FrmStudentRegistration.Show
'''''End Sub
'''''
'''''Private Sub CmdUpdateStudent_Click()
'''''    FrmUpdateStuDetails.Show
'''''
'''''End Sub
'''''
'''''
'''''
'''''Private Sub smnuBranch_Click()
'''''    dRptBranch.Show
'''''End Sub
'''''

Private Sub CmdExit_Click()
    Shell App.Path & "\DeleteTEMP.bat", vbHide
    End
End Sub



Private Sub CmdStudent_Click()
    FrmViewStudentRoomDetails.Show
End Sub

Private Sub smnuAllAllocatedRooms_Click()
    DataEnvironment1.Command7_Grouping
    DrptAllAllocatedRooms.Show
End Sub

Private Sub MDIForm_Terminate()
    Shell App.Path & "\DeleteTEMP.bat", vbHide

End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    Shell App.Path & "\DeleteTEMP.bat", vbHide

End Sub

Private Sub smnuAllocate_Click()
    MsgBox "Automatic Allocation waiting for Supervisor Recommendation!", vbInformation
End Sub

Private Sub smnuAllStaffDetails_Click()
    DrptAllStaffInfo.Show
End Sub

Private Sub smnuAllStudentDetails_Click()
    DrptAllStudentInfo.Show
End Sub

Private Sub smnuAllUnallocatedRooms_Click()
    DataEnvironment1.Command4_Grouping
    DrptAllUnAllocatedRooms.Show

End Sub

Private Sub smnuBranch_Click()
    dRptBranch.Show
End Sub

Private Sub smnuBranchWise_Click()
    sSQL = "Select DISTINCT tblBranchDetails.BranchName, tblStudentInfo.BranchCode From tblBranchDetails,tblStudentInfo"
    Set rs = Con.Execute(sSQL)
    frmRptStudentChoice.Combo1.Clear
    With rs
        Do While Not .EOF
            frmRptStudentChoice.Combo1.AddItem .Fields(0)
            .MoveNext
        Loop
    End With

End Sub

Private Sub smnuCategoryWise_Click()
        frmRptStudentChoice.Combo1.Clear
        frmRptStudentChoice.Combo1.AddItem "General"
        frmRptStudentChoice.Combo1.AddItem "OBC"
        frmRptStudentChoice.Combo1.AddItem "ST"
        frmRptStudentChoice.Combo1.AddItem "SC"
    frmRptStudentChoice.Show

End Sub

Private Sub smnuDeAllocateRoom_Click()
    FrmDeAllocateRoom.Show
End Sub

Private Sub smnuDeptwise_Click()
    sSQL = "Select DeptName From tblDepartment"
    Set rs = Con.Execute(sSQL)
    frmRptStaffChoice.Combo1.Clear
    Opt = 1
    With rs
        Do While Not .EOF
            frmRptStaffChoice.Combo1.AddItem .Fields(0)
            .MoveNext
        Loop
    End With

End Sub


Private Sub smnuDesignationWise_Click()
    sSQL = "Select DesignationName From tblDesignation"
    Set rs = Con.Execute(sSQL)
    frmRptStaffChoice.Combo1.Clear
    Opt = 2
    With rs
        Do While Not .EOF
            frmRptStaffChoice.Combo1.AddItem .Fields(0)
            .MoveNext
        Loop
    End With

End Sub



Private Sub smnuExit_Click()

    res = MsgBox("Are You Sure You Want to Exit Application", vbYesNo + vbQuestion, "Close Application")
    If res = vbYes Then
        End
    End If
    
End Sub

Private Sub smnuFeesReciept_Click()

    sSQL = "Select FeeReceiptNo,EnrollNo From tblPayReceipt"
    Set rs = Con.Execute(sSQL)
    frmRptPrintReciept.Combo1.Clear
    With rs
        Do While Not .EOF
            frmRptPrintReciept.Combo1.AddItem .Fields(0) '& " - " & .Fields(1)
            .MoveNext
        Loop
    End With


End Sub

Private Sub smnuFeesStructureDetails_Click()
    DrptFeeStructure.Show
End Sub

Private Sub smnuFullyAllocatedRooms_Click()
    DataEnvironment1.Command5_Grouping
    DrptFullyAllocatedRooms.Show
End Sub

Private Sub smnuGenderWise_Click()
        
    Opt = 3
    frmRptStaffChoice.Combo1.Clear
    frmRptStaffChoice.Combo1.AddItem "Male"
    frmRptStaffChoice.Combo1.AddItem "Female"
    frmRptStaffChoice.Show

End Sub

Private Sub smnuHostelDetails_Click()
    dRptHostelDetails.Show
End Sub

Private Sub smnuHostelRoomDetails_Click()
    frmRptHostel.Show
End Sub

Private Sub smnuPaidFeesDetails_Click()
    DrptPaidFeeDetails.Show
End Sub

Private Sub smnuPartiallyAllocatedRooms_Click()
    DataEnvironment1.Command3_Grouping
    DrptPartiallyAllocatedRooms.Show
End Sub

Private Sub smnuRefundFeesDetails_Click()

    With DataEnvironment1.conYCT
    If .State Then
        .Close
    End If
    End With

    DataEnvironment1.conYCT.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\DataBase\HostelMgmt.mdb;Persist Security Info=False"
    DataEnvironment1.conYCT.Open


        DrptRefundFeeDetails.Show
End Sub

Private Sub smnuStaffDetails_Click()
    frmRptStaffInfoAll.Show
End Sub


Private Sub smnuStaffGenderWise_Click()
        frmRptStudentChoice.Combo1.Clear
        frmRptStudentChoice.Combo1.AddItem "Male"
        frmRptStudentChoice.Combo1.AddItem "Female"
        frmRptStudentChoice.Show

End Sub
'''''
'''''
Private Sub smnuStudentRegDetails_Click()
    sSQL = "Select EnrollNo From tblStudentInfo"
    Set rs = Con.Execute(sSQL)
    frmRptStudentDetails.Combo1.Clear
    With rs
        Do While Not .EOF
            frmRptStudentDetails.Combo1.AddItem .Fields(0)
            .MoveNext
        Loop
    End With
'
'Dim str As String
'str = "A123"
'    DataEnvironment1.Command1 str
''    DrptStudentInfo.p
'    DrptStudentInfo.Show

End Sub


Private Sub smnuViewStudentFeeDetails_Click()
    FrmViewStudentFeeDetails.Show
End Sub

Private Sub smnuViewStudentRoomDetails_Click()
    FrmViewStudentRoomDetails.Show
End Sub

Private Sub smnuYearWise_Click()

        frmRptStudentChoice.Combo1.Clear
        frmRptStudentChoice.Combo1.AddItem "First Sem"
        frmRptStudentChoice.Combo1.AddItem "Second Sem"
        frmRptStudentChoice.Combo1.AddItem "Third Sem"
        frmRptStudentChoice.Combo1.AddItem "Fourth Sem"
        frmRptStudentChoice.Combo1.AddItem "Fifth Sem"
        frmRptStudentChoice.Combo1.AddItem "Sixth Sem"
    frmRptStudentChoice.Show

End Sub


Private Sub CmdExit1_Click()
    smnuExit_Click
'    Shell App.Path & "\DeleteTEMP.bat", vbHide
'    End
End Sub

Private Sub MDIForm_Load()
    Call Connect
    With DataEnvironment1.conYCT
    If .State Then
        .Close
    End If
    End With

    DataEnvironment1.conYCT.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\DataBase\HostelMgmt.mdb;Persist Security Info=False"
    DataEnvironment1.conYCT.Open

End Sub

Private Sub smnuAllocateRoom_Click()
    FrmAllocateRoom.Show
End Sub

Private Sub smnuCreateFeeStructure_Click()
    FrmFeeStructure.Show
End Sub

Private Sub smnuCreateNewHostel_Click()
    FrmAddHostel.Show
End Sub

Private Sub smnuFeeRefund_Click()
    FrmRefundFees.Show
End Sub

Private Sub smnuPaidFee_Click()
    FrmFeeReciept.Show
End Sub

Private Sub smnuSearchStudent_Click()
    FrmSearchStudent.Show
End Sub

Private Sub smnuStaffDetail_Click()
    FrmViewStaffDetails.Show
End Sub

Private Sub smnuStaffRegistration_Click()
    FrmStaffRegistration.Show
End Sub

Private Sub smnuStudentRegistration_Click()
'    Load FrmStudentRegistration
    FrmStudentRegistration.Show
End Sub

Private Sub smnuUpdateStudent_Click()
    FrmUpdateStuDetails.Show
End Sub

Private Sub smnuViewHostel_Click()
'    FrmViewHostelDetail.Show

End Sub

Private Sub ssmnuBranch_Click()
    FrmAddBranch.Show
End Sub

Private Sub ssmnuDepartment_Click()
    FrmAddDepartment.Show
End Sub

Private Sub ssmnuDesignation_Click()
    FrmAddDesignation.Show
End Sub


Private Sub ssmnuRooms_Click()
    FrmAddRoomDetails.Show
End Sub
