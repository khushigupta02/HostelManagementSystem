VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmAllocateRoom 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   9750
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13245
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "FrmAllocateRoom.frx":0000
   ScaleHeight     =   9750
   ScaleWidth      =   13245
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8655
      Left            =   240
      TabIndex        =   0
      Top             =   840
      Width           =   12735
      Begin VB.Frame Frame2 
         BackColor       =   &H00000000&
         Caption         =   "ROOM ALLOCATE"
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
         Height          =   7935
         Left            =   240
         TabIndex        =   6
         Top             =   480
         Width           =   12205
         Begin VB.TextBox TxtID 
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   405
            Left            =   10920
            TabIndex        =   24
            Top             =   2880
            Width           =   975
         End
         Begin VB.TextBox TxtFeeReceipt 
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   405
            Left            =   8880
            TabIndex        =   22
            Top             =   2880
            Width           =   2055
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Only Allocated Rooms"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   8880
            TabIndex        =   21
            Top             =   360
            Width           =   2895
         End
         Begin VB.ComboBox CmbStuEnrollment 
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
            ItemData        =   "FrmAllocateRoom.frx":103B
            Left            =   3240
            List            =   "FrmAllocateRoom.frx":1048
            Style           =   2  'Dropdown List
            TabIndex        =   18
            Top             =   1080
            Width           =   3135
         End
         Begin VB.ComboBox CmbHostalName 
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
            ItemData        =   "FrmAllocateRoom.frx":1061
            Left            =   3240
            List            =   "FrmAllocateRoom.frx":106E
            Style           =   2  'Dropdown List
            TabIndex        =   17
            Top             =   360
            Width           =   5535
         End
         Begin VB.TextBox TxtGender 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   405
            Left            =   8880
            TabIndex        =   1
            Top             =   1560
            Width           =   3015
         End
         Begin VB.TextBox TxtFatherName 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   405
            Left            =   3240
            TabIndex        =   2
            Top             =   2040
            Width           =   3135
         End
         Begin VB.TextBox TxtBranch 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   405
            Left            =   8880
            TabIndex        =   3
            Top             =   2040
            Width           =   3015
         End
         Begin VB.TextBox TxtStuName 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   405
            Left            =   3240
            TabIndex        =   4
            Top             =   1560
            Width           =   3135
         End
         Begin VB.TextBox TxtRoomNo 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   405
            Left            =   8880
            TabIndex        =   5
            Top             =   1080
            Width           =   3015
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H8000000C&
            Height          =   1215
            Left            =   240
            TabIndex        =   7
            Top             =   3600
            Width           =   11775
            Begin VB.CommandButton CmdClose 
               Caption         =   "CLOSE"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   795
               Left            =   7920
               TabIndex        =   16
               Top             =   240
               Width           =   2055
            End
            Begin VB.CommandButton CmdAllote 
               Caption         =   "ALLOT ROOM"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   795
               Left            =   1920
               TabIndex        =   9
               Top             =   240
               Width           =   2055
            End
            Begin VB.CommandButton CmdCancel 
               Caption         =   "CANCEL"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   795
               Left            =   5040
               TabIndex        =   8
               Top             =   240
               Width           =   2055
            End
         End
         Begin MSComCtl2.DTPicker DTPDOAllotement 
            Height          =   375
            Left            =   3240
            TabIndex        =   26
            Top             =   2880
            Width           =   3135
            _ExtentX        =   5530
            _ExtentY        =   661
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "dd/MM/yyyy"
            Format          =   118095875
            CurrentDate     =   43898
         End
         Begin MSFlexGridLib.MSFlexGrid MSFRoomDetails 
            Height          =   2655
            Left            =   240
            TabIndex        =   27
            Top             =   5040
            Width           =   11685
            _ExtentX        =   20611
            _ExtentY        =   4683
            _Version        =   393216
            SelectionMode   =   1
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
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "FEE RECIEPT NO"
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
            Left            =   6600
            TabIndex        =   23
            Top             =   2940
            Width           =   2100
         End
         Begin VB.Line Line2 
            BorderColor     =   &H80000005&
            BorderWidth     =   3
            X1              =   0
            X2              =   12200
            Y1              =   960
            Y2              =   960
         End
         Begin VB.Line Line1 
            BorderColor     =   &H80000005&
            BorderWidth     =   3
            X1              =   0
            X2              =   12200
            Y1              =   2640
            Y2              =   2640
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "GENDER"
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
            Left            =   6600
            TabIndex        =   20
            Top             =   1560
            Width           =   1095
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "BRANCH"
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
            Left            =   6600
            TabIndex        =   19
            Top             =   2100
            Width           =   1110
         End
         Begin VB.Label Label10 
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
            TabIndex        =   15
            Top             =   465
            Width           =   1890
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "DATE OF ALLOTMENT"
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
            TabIndex        =   14
            Top             =   2925
            Width           =   2805
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "SELECT ENROLLMENT"
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
            TabIndex        =   13
            Top             =   1185
            Width           =   2865
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "STUDENT NAME"
            DataField       =   "V"
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
            TabIndex        =   12
            Top             =   1620
            Width           =   2070
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ROOM NO"
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
            Left            =   6600
            TabIndex        =   11
            Top             =   1140
            Width           =   1275
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "FATHER NAME"
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
            TabIndex        =   10
            Top             =   2100
            Width           =   1875
         End
      End
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "ROOM ALLOCATION"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   360
      TabIndex        =   25
      Top             =   240
      Width           =   12495
   End
End
Attribute VB_Name = "FrmAllocateRoom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsHostel As New Recordset
Dim sSQL As String
Dim rs As New Recordset
Dim rsStud1 As New Recordset
Dim Validate As Boolean

Private Sub Check1_Click()
    loadRS
End Sub

Private Sub CmbHostalName_Click()
    MSFRoomDetails.Enabled = True
    loadRS

End Sub

Private Sub CmbStuEnrollment_Click()
    
    sSQL = "Select EnrollNo, StuName, FName, BranchCode, CurrentSem, Gender, PaidReceiptNo from tblStudentInfo where EnrollNo = '" & CmbStuEnrollment.Text & "'"
'    sSQL = "Select EnrollNo from tblStudentInfo where Allocated = " & False
    If rs.State = 1 Then rs.Close
    rs.Open sSQL, Con, 2, 3
    
    TxtStuName = rs.Fields("StuName")
    TxtFatherName = rs.Fields("FName")
    TxtBranch = rs.Fields("BranchCode")
    TxtGender = rs.Fields("Gender")
    TxtFeeReceipt = rs.Fields("PaidReceiptNo")
    
    Call ChkFee
    
End Sub

Private Sub CmdAllote_Click()
    Dim strStud As String
    Dim intAlloc As Integer

    If CmbStuEnrollment.Text = "" Then
        MsgBox "Select a Student's Enrollment number first!", vbExclamation
        Exit Sub
    End If
    
    If TxtRoomNo.Text = "" Then
        MsgBox "Select a Room Number From List first!", vbExclamation
        Exit Sub
    End If
    
    If MSFRoomDetails.row < 0 Then
        MsgBox "Please Select a room to allocate!", vbCritical
        Exit Sub
    End If
    
    If TxtFeeReceipt.Text = "0" Then
        MsgBox "You don't pay hostel fee " & ctrf & " Please Paid Your Hostel Fee first!", vbExclamation
        Exit Sub
    End If
    
    Call ChkFee
    If Validate = True Then
        
        strStud = CmbStuEnrollment.Text
        sSQL = "Select Allocated, Gender from tblStudentInfo where EnrollNo = '" & strStud & "'"
        If rsStud1.State = 1 Then rsStud1.Close
        rsStud1.Open sSQL, Con, 2, 3
        
        If rsStud1.BOF And rsStud1.EOF Then
            MsgBox "Sorry, no student like that..."
            Exit Sub
        End If
        
        If rsStud1.Fields("Allocated") = True Then
            MsgBox "Student already allocated to a room.", vbCritical
            Exit Sub
        Else
            'check students sex
            If LCase$(MSFRoomDetails.TextMatrix(MSFRoomDetails.row, 5)) <> LCase$(rsStud1.Fields("Gender")) Then
                MsgBox "Wrong Gender Allocation... Check Student's Gender", vbCritical
                Exit Sub
            End If
            
            'do the allocation here
            
            intAlloc = CInt(MSFRoomDetails.TextMatrix(MSFRoomDetails.row, 6))
            intCapacity = CInt(MSFRoomDetails.TextMatrix(MSFRoomDetails.row, 4))
            
            If intAlloc = intCapacity Then
                MsgBox "Sorry, Room is Full ! Please Select Another Room !", vbCritical
                Exit Sub
            End If
            
            intAlloc = CInt(MSFRoomDetails.TextMatrix(MSFRoomDetails.row, 6)) + 1
            
            'insert to room allocation table
            sSQL = "Insert into RoomAllocation( RoomId, EnrollNo, HostelName) Values ('" & MSFRoomDetails.TextMatrix(MSFRoomDetails.row, 1) & "','" & strStud & "','" & CmbHostalName.Text & "')"
            If rsStud1.State = 1 Then rsStud1.Close
            rsStud1.Open sSQL, Con, 2, 3
            
            sSQL = "Insert into tblRoomAllocation(RoomId, EnrollNo, HostelName, AlloteDate ) Values ('" & MSFRoomDetails.TextMatrix(MSFRoomDetails.row, 1) & "','" & strStud & "','" & CmbHostalName.Text & "', '" & DTPDOAllotement.Value & "')"
            If rsStud1.State = 1 Then rsStud1.Close
            rsStud1.Open sSQL, Con, 2, 3
            
            'indicate in student table that student is allocated
            sSQL = "Update tblStudentInfo Set Allocated = " & True & " Where EnrollNo = '" & strStud & "'"
            If rsStud1.State = 1 Then rsStud1.Close
            rsStud1.Open sSQL, Con, 2, 3
            
            'update current status of room allocation
            
            sSQL = "Update tblRoomDetails Set Allocated = " & intAlloc & " Where RoomId = '" & MSFRoomDetails.TextMatrix(MSFRoomDetails.row, 1) & "' and  Gender = '" & MSFRoomDetails.TextMatrix(MSFRoomDetails.row, 5) & "'"
'            MsgBoxs sSQL
            If rsStud1.State = 1 Then rsStud1.Close
            rsStud1.Open sSQL, Con, 2, 3
            
            
'            sSQL = "Update tblRoomDetails Set Allocated = " & intAlloc & " Where RoomId = '" & MSFRoomDetails.TextMatrix(MSFRoomDetails.row, 1) & "' and  Gender = '" & MSFRoomDetails.TextMatrix(MSFRoomDetails.row, 5) & "'"
''            MsgBoxs sSQL
'            If rsStud1.State = 1 Then rsStud1.Close
'            rsStud1.Open sSQL, Con, 2, 3
            
            
            
            MsgBox "Room Allocated Successfully", vbInformation
            loadRS
            ClearTextBox Me
            loadUnAllocated
        End If
    End If
    
End Sub

Private Sub CmdCancel_Click()
    Call ClearTextBox(Me)
End Sub

Private Sub CmdClose_Click()
    Unload Me
End Sub


Private Sub Form_Load()
    Me.Top = 100
    Me.Left = 100
'    Call LockText(Me)
    AddData
    loadUnAllocated
    DTPDOAllotement = Format(Date, "dd/MM/yyyy")
'    fetchMaxID
End Sub

'Private Sub fetchMaxID()
'    If rsMaxId.State = 1 Then rsMaxId.Close
'    rsMaxId.Open "Select max(ID) from tblRoomAllocation", Con, 2, 3
''    MaxID
'    If IsNull(rsMaxId(0)) Then
'        MaxID = 1
'    Else
'        MaxID = rsMaxId.Fields(0) + 1
'    End If
'
'End Sub


Private Sub loadUnAllocated()
    'On Error Resume Next
    CmbStuEnrollment.Clear
    sSQL = "Select EnrollNo from tblStudentInfo Where Allocated = " & False & " And PaidReceiptNo>=0"
    If rs.State = 1 Then rs.Close
    rs.Open sSQL, Con, 2, 3
    Do While Not rs.EOF
        CmbStuEnrollment.AddItem rs.Fields(0)
        rs.MoveNext
    Loop
End Sub

Private Sub AddData()
    
    If rsHostel.State = 1 Then rsHostel.Close
    rsHostel.Open "Select * from tblHostelDetail", Con, 2, 3
    
'    MsgBox rsHostel.RecordCount
    
    If IsNull(rsHostel(1)) Or rsHostel.RecordCount = 0 Then
        MsgBox " Hostel Or Room Information Not Available Please Add Hostel and Room Information Before Allocation", vbInformation
    Else
        rsHostel.MoveFirst
        CmbHostalName.Clear
        Do While Not rsHostel.EOF
            CmbHostalName.AddItem rsHostel.Fields("HostelName")         '   & " - ( " & rsHostel.Fields("Prefix") & " )"
            rsHostel.MoveNext
        Loop
    End If

    
'    rsHostel.MoveFirst
'    CmbHostalName.Clear
'    Do While Not rsHostel.EOF
'        CmbHostalName.AddItem rsHostel.Fields("HostelName")         '   & " - ( " & rsHostel.Fields("Prefix") & " )"
'        rsHostel.MoveNext
'    Loop

End Sub

Public Sub loadRS()
    If CmbHostalName.Text = "" Then Exit Sub
    sSQL = "Select RoomID,HostelName,RoomNo,Capacity,Gender,Allocated from tblRoomDetails where Hostelname = '" & CmbHostalName.Text & "'"
    
    If txtRoomSearch <> "" Then
        sSQL = sSQL & " and RoomNumber like '" & txtRoomSearch.Text & "%'"
    End If

    If Check1.Value = vbChecked Then
        sSQL = sSQL & " and allocated > 0"
    End If
    
loadRS1:

    If rs.State = 1 Then rs.Close
    rs.Open sSQL, Con, 2, 3
    If rs.EOF And rs.BOF Then
        MsgBox "Room Information not found", vbCritical
        sSQL = "Select * from tblRoomDetails where Hostelname = '" & CmbHostalName.Text & "'"
        'GoTo loadRS1
    Else
        LoadRecordsetIntoGrid rs, MSFRoomDetails
    End If
    
    loadUnAllocated
End Sub


Private Sub MSFRoomDetails_Click()
On Error Resume Next
'If MSFRoomDetails.se Then
'End If
    TxtRoomNo.Text = MSFRoomDetails.TextMatrix(MSFRoomDetails.row, 3)
End Sub

Private Sub MSFRoomDetails_DblClick()
    Dim rs1 As New Recordset
    sSQL = "Select RoomId,EnrollNo,ID From RoomAllocation Where RoomID = '" & MSFRoomDetails.TextMatrix(MSFRoomDetails.row, 1) & "'"
'    MsgBox sSQL
    If rs1.State = 1 Then rs1.Close
    rs1.Open sSQL, Con, 2, 3
    If rs1.EOF And rs1.BOF Then
        MsgBox "No Allocation for Room ", vbCritical
        Exit Sub
    End If
    Load FrmDeAllocateRoom
    With FrmDeAllocateRoom
        LoadRecordsetIntoGrid rs1, .grd
        .Show
    End With

End Sub

Private Sub ChkFee()
On Error Resume Next
    If rs.State = 1 Then rs.Close
    rs.Open "Select RemainAmt From tblPayReceipt Where EnrollNo='" & CmbStuEnrollment & "'", Con, 2, 3
    If rs.Fields("RemainAmt") >= 1 Then
        MsgBox "You Can't Allot Hostel or Room Because Fees is Not Paid ", vbInformation
        Validate = False
    Else
        Validate = True
    End If
End Sub


