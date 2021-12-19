VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmDeAllocateRoom 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   9045
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13230
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "FrmDeAllocateRoom.frx":0000
   ScaleHeight     =   9045
   ScaleWidth      =   13230
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
      Height          =   7935
      Left            =   240
      TabIndex        =   5
      Top             =   840
      Width           =   12735
      Begin VB.Frame Frame2 
         BackColor       =   &H00000000&
         Caption         =   "ROOM DE-ALLOCATE"
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
         Height          =   7335
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   12205
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
            Left            =   9360
            TabIndex        =   16
            Top             =   1680
            Visible         =   0   'False
            Width           =   2535
         End
         Begin VB.ComboBox CmbRoomNo 
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
            Left            =   9360
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   360
            Width           =   2535
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
            Left            =   2520
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   360
            Width           =   5055
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
            Left            =   2520
            TabIndex        =   6
            Top             =   1080
            Width           =   5055
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H8000000C&
            Height          =   1215
            Left            =   240
            TabIndex        =   8
            Top             =   1800
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
               TabIndex        =   14
               Top             =   240
               Width           =   2055
            End
            Begin VB.CommandButton CmdDeAllote 
               Caption         =   "DE-ALLOT ROOM"
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
               TabIndex        =   4
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
               TabIndex        =   9
               Top             =   240
               Width           =   2055
            End
         End
         Begin MSFlexGridLib.MSFlexGrid grd 
            Height          =   3855
            Left            =   240
            TabIndex        =   3
            Top             =   3240
            Width           =   11775
            _ExtentX        =   20770
            _ExtentY        =   6800
            _Version        =   393216
            Enabled         =   -1  'True
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
         Begin MSComCtl2.DTPicker DTPDeAllotement 
            Height          =   375
            Left            =   9360
            TabIndex        =   2
            Top             =   1080
            Width           =   2535
            _ExtentX        =   4471
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
            Format          =   93847555
            CurrentDate     =   43898
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
            Left            =   7920
            TabIndex        =   15
            Top             =   1740
            Visible         =   0   'False
            Width           =   1275
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
            Y1              =   1680
            Y2              =   1680
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
            TabIndex        =   13
            Top             =   465
            Width           =   1890
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "DE-ALLOT DATE"
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
            Left            =   7680
            TabIndex        =   12
            Top             =   1005
            Width           =   1755
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ROOM ID"
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
            Left            =   7800
            TabIndex        =   11
            Top             =   465
            Width           =   1170
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
            TabIndex        =   10
            Top             =   1140
            Width           =   2070
         End
      End
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "ROOM DE-ALLOCATION"
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
      TabIndex        =   17
      Top             =   240
      Width           =   12495
   End
End
Attribute VB_Name = "FrmDeAllocateRoom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsHostel As New ADODB.Recordset


Public Sub ShowData()

    Dim rs1 As New ADODB.Recordset
    sSQL = "Select RoomId,EnrollNo,ID from RoomAllocation Where HostelName = '" & CmbHostalName & "' AND RoomID = '" & CmbRoomNo & "'"    'where RoomID = '" & grd.TextMatrix(grd.Row, 2) & "'"
'    Set rs1 = cn.Execute(sSQL)
    If rs1.State = 1 Then rs1.Close
    rs1.Open sSQL, Con, 2, 3

    If rs1.EOF And rs1.BOF Then
        MsgBox "No Allocation for Room ", vbCritical
        grd.Clear
        Exit Sub
    End If
'    Load frmRoomDetails
    With FrmDeAllocateRoom
        LoadRecordsetIntoGrid rs1, .grd
        .Show
    End With
    
    grd_Add
    
    
End Sub


Private Sub CmbHostalName_Click()
    
    If rsHostel.State = 1 Then rsHostel.Close
    rsHostel.Open "Select * from tblRoomDetails where HostelName='" & CmbHostalName & "'", Con, 2, 3
    
        If rsHostel.EOF And rsHostel.BOF Then
        MsgBox "Rooms Not Found ", vbCritical
        Exit Sub
    Else
        rsHostel.MoveFirst
        CmbRoomNo.Clear
        Do While Not rsHostel.EOF
            CmbRoomNo.AddItem rsHostel.Fields("RoomID")         '   & " - ( " & rsHostel.Fields("Prefix") & " )"
            rsHostel.MoveNext
        Loop
    End If
'    ShowData

End Sub

Private Sub CmbRoomNo_Click()
    ShowData
End Sub

Private Sub CmdCancel_Click()

'    sSQL = "Update tblRoomAllocation Set DeAlloteDate = '" & DTPDeAllotement.Value & "' Where EnrollNo='A123' AND ID = 3"
'    If rs.State = 1 Then rs.Close
'    rs.Open sSQL, con, 2, 3
''rs.MoveFirst
''MsgBox rs.Fields(0)
End Sub

Private Sub CmdClose_Click()
    Unload Me
End Sub

Private Sub CmdDeAllote_Click()
    Dim strStud As String
    Dim intAlloc As Integer

    If grd.row < 1 Then
        MsgBox "Please Select Student to de-allocate!", vbCritical
        Exit Sub
    End If
    
    strStud = grd.TextMatrix(grd.row, 2)
    
    sSQL = "Delete From RoomAllocation Where RoomID = '" & grd.TextMatrix(grd.row, 1) & "' and EnrollNo = '" & grd.TextMatrix(grd.row, 2) & "'"
    If rs.State = 1 Then rs.Close
    rs.Open sSQL, Con, 2, 3
    
    sSQL = "Update tblStudentInfo Set Allocated = " & False & " Where EnrollNo = '" & strStud & "'"
    If rs.State = 1 Then rs.Close
    rs.Open sSQL, Con, 2, 3
    
    sSQL = "Update tblRoomAllocation Set DeAlloteDate = '" & DTPDeAllotement.Value & "' Where EnrollNo = '" & strStud & "' AND ID=" & grd.TextMatrix(grd.row, 3) & ""
    If rs.State = 1 Then rs.Close
    rs.Open sSQL, Con, 2, 3
    
    sSQL = "Select Allocated from tblRoomDetails Where RoomId = '" & grd.TextMatrix(grd.row, 1) & "'"
    If rs.State = 1 Then rs.Close
    rs.Open sSQL, Con, 2, 3
    
    intAlloc = CInt(rs.Fields(0)) - 1
    
    sSQL = "Update tblRoomDetails Set Allocated = " & intAlloc & " Where RoomId = '" & grd.TextMatrix(grd.row, 1) & "'"
    If rs.State = 1 Then rs.Close
    rs.Open sSQL, Con, 2, 3

'    frmHostelMgt.grd_dblClick
'    Dim frm As New FrmAllocateRoom
'       frm.MSFRoomDetails_dbl
    ShowData
End Sub

Private Sub Form_Load()
    Me.Top = 100
    Me.Left = 100

    DTPDeAllotement = Format(Date, "dd/MM/yyyy")
    AddData
    grd_Add
End Sub

Private Sub AddData()
    
    If rsHostel.State = 1 Then rsHostel.Close
    rsHostel.Open "Select * from tblHostelDetail", Con, 2, 3
    
    If IsNull(rsHostel(1)) Or rsHostel.RecordCount = 0 Then
        MsgBox " Hostel Or Room Information Not Available Please Allocate Room and Add Hostel and Room Information Before De Allocation", vbInformation
        Exit Sub
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

Private Sub grd_Click()

    If grd.row < 1 Then
        MsgBox "Please Select Hostel Name and Room No for Student to de-allocate!", vbCritical
        Exit Sub
    End If

    If rsHostel.State = 1 Then rsHostel.Close
    rsHostel.Open "Select StuName from tblStudentInfo Where EnrollNo= '" & grd.TextMatrix(grd.row, 2) & "'", Con, 2, 3
'        TxtStuName = rsHostel.Fields("StuName")        '   & " - ( " & rsHostel.Fields("Prefix") & " )"
    
    If rsHostel.RecordCount = 0 Then            ' IsNull(TxtStuName) Or IsEmpty(TxtStuName) Or TxtStuName = "" Then
        MsgBox " Student Information Not Available Please Allocate Room Before De Allocation", vbInformation
        Exit Sub
    Else
        TxtStuName = rsHostel.Fields("StuName")        '   & " - ( " & rsHostel.Fields("Prefix") & " )"
    End If




End Sub


Private Sub grd_Add()
'  grd.Cols(3).
    grd.Cols = grd.Cols + 1
    grd.Cols = grd.Cols + 1
End Sub

