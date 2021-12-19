VERSION 5.00
Begin VB.Form FrmAddRoomDetails 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "VIEW OR ADD ROOM DETAILS FORM"
   ClientHeight    =   8175
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10950
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "FrmAddRoomDetails.frx":0000
   ScaleHeight     =   8175
   ScaleWidth      =   10950
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
      Height          =   7095
      Left            =   360
      TabIndex        =   0
      Top             =   840
      Width           =   10335
      Begin VB.Frame sss 
         BackColor       =   &H00000000&
         Caption         =   "ADD ROOM "
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
         Height          =   6255
         Left            =   240
         TabIndex        =   9
         Top             =   600
         Width           =   9855
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
            ItemData        =   "FrmAddRoomDetails.frx":103B
            Left            =   2640
            List            =   "FrmAddRoomDetails.frx":1048
            Style           =   2  'Dropdown List
            TabIndex        =   21
            Top             =   480
            Width           =   6975
         End
         Begin VB.ComboBox CmbRoomType 
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
            ItemData        =   "FrmAddRoomDetails.frx":1061
            Left            =   7320
            List            =   "FrmAddRoomDetails.frx":106B
            Style           =   2  'Dropdown List
            TabIndex        =   19
            Top             =   1320
            Width           =   2295
         End
         Begin VB.Frame Frame2 
            BackColor       =   &H8000000C&
            Height          =   1095
            Left            =   240
            TabIndex        =   14
            Top             =   4800
            Width           =   9375
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
               Left            =   7080
               TabIndex        =   18
               Top             =   240
               Width           =   1900
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
               Left            =   4800
               TabIndex        =   17
               Top             =   240
               Width           =   1900
            End
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
               Left            =   2520
               TabIndex        =   16
               Top             =   240
               Width           =   1900
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
               Left            =   240
               TabIndex        =   15
               Top             =   240
               Width           =   1900
            End
         End
         Begin VB.TextBox TxtRoomNo 
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
            Left            =   2640
            TabIndex        =   1
            Top             =   1320
            Width           =   2535
         End
         Begin VB.TextBox TxtRoomCapacity 
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
            Left            =   2640
            MaxLength       =   1
            TabIndex        =   2
            Text            =   "0"
            Top             =   2040
            Width           =   2535
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H00FFFFFF&
            Caption         =   "ROOM FACILITY "
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
            Height          =   1455
            Left            =   960
            TabIndex        =   10
            Top             =   2880
            Width           =   8175
            Begin VB.CheckBox ChkChair 
               BackColor       =   &H00FFFFFF&
               Caption         =   "CHAIR"
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
               Height          =   285
               Left            =   3360
               TabIndex        =   4
               Top             =   960
               Value           =   1  'Checked
               Width           =   1575
            End
            Begin VB.CheckBox ChkTable 
               BackColor       =   &H00FFFFFF&
               Caption         =   "TABLE"
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
               Height          =   285
               Left            =   360
               TabIndex        =   6
               Top             =   960
               Value           =   1  'Checked
               Width           =   2055
            End
            Begin VB.CheckBox ChkTubelight 
               BackColor       =   &H00FFFFFF&
               Caption         =   "TUBELIGHT"
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
               Height          =   285
               Left            =   360
               TabIndex        =   3
               Top             =   480
               Value           =   1  'Checked
               Width           =   2055
            End
            Begin VB.CheckBox ChkHolder 
               BackColor       =   &H00FFFFFF&
               Caption         =   "HOLDER"
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
               Height          =   285
               Left            =   3360
               TabIndex        =   5
               Top             =   480
               Value           =   1  'Checked
               Width           =   1575
            End
            Begin VB.CheckBox ChkCellingFan 
               BackColor       =   &H00FFFFFF&
               Caption         =   "CELLING FAN"
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
               Height          =   285
               Left            =   5640
               TabIndex        =   7
               Top             =   480
               Value           =   1  'Checked
               Width           =   2415
            End
            Begin VB.CheckBox ChkSleepingBed 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Caption         =   "SLEEPING BED"
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
               Height          =   285
               Left            =   5640
               TabIndex        =   8
               Top             =   960
               Value           =   1  'Checked
               Width           =   2415
            End
         End
         Begin VB.Line Line3 
            BorderColor     =   &H00FFFFFF&
            BorderWidth     =   2
            X1              =   0
            X2              =   9850
            Y1              =   1080
            Y2              =   1095
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ROOM TYPE"
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
            Left            =   5400
            TabIndex        =   20
            Top             =   1395
            Width           =   1530
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00FFFFFF&
            BorderWidth     =   2
            X1              =   0
            X2              =   9850
            Y1              =   4560
            Y2              =   4575
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00FFFFFF&
            BorderWidth     =   2
            X1              =   0
            X2              =   9850
            Y1              =   2640
            Y2              =   2655
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "HOSTEL NAME "
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
            Top             =   555
            Width           =   1845
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ROOM NO. "
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
            Top             =   1395
            Width           =   1425
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ROOM CAPACITY "
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
            TabIndex        =   11
            Top             =   2115
            Width           =   2250
         End
      End
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "VIEW && ADD ROOM"
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
      Left            =   480
      TabIndex        =   22
      Top             =   240
      Width           =   10095
   End
End
Attribute VB_Name = "FrmAddRoomDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Validate As Boolean
Dim rsHostel As New Recordset

    Dim intCapacity As Integer
    Dim strRoomNumber, strSex As String
    Dim TotalCapacity As Integer

Private Sub CmbHostalName_Click()

    sSQL = "Select Gender from tblHostelDetail where hostelname = '" & CmbHostalName & "'"
'    Set rsHostel = cn.Execute(sSQL)
    If rsHostel.State = 1 Then rsHostel.Close
    rsHostel.Open sSQL, Con, 2, 3
    
    If LCase$(rsHostel.Fields(0)) = "male" Then
        CmbRoomType.Clear
        CmbRoomType.AddItem "Male"
    ElseIf LCase$(rsHostel.Fields(0)) = "female" Then
        CmbRoomType.Clear
        CmbRoomType.AddItem "Female"
    ElseIf LCase$(rsHostel.Fields(0)) = "both" Then
        CmbRoomType.Clear
        CmbRoomType.AddItem "Both"
    Else
        CmbRoomType.Clear
        CmbRoomType.AddItem "Male"
        CmbRoomType.AddItem "Female"
    End If
    
            fetchMaxID
        TxtRoomNo.Text = MaxID
        TxtRoomNo.Locked = True


End Sub

Private Sub CmdAddNew_Click()

    If CmbHostalName.Text = "" Then
        MsgBox "Please Select Hostal Name First "
        CmbHostalName.SetFocus
        Exit Sub
    Else
    
        CmdSave.Enabled = True
        CmdAddNew.Enabled = False
        UnLockText Me
        
        fetchMaxID
        TxtRoomNo.Text = MaxID
        TxtRoomNo.Locked = True
    End If

End Sub

Private Sub fetchMaxID()

    If rs.State = 1 Then rs.Close
    rs.Open "Select Count(RoomNo) from tblRoomDetails where HostelName = '" & CmbHostalName.Text & "'", Con, 2, 3
'    MsgBox rs.RecordCount
    If rs.BOF = True And rs.EOF = True Then
        MaxID = 1
    Else
        MaxID = rs.Fields(0).Value + 1
    End If
    
'    If IsNull(rsMaxId(0)) Then
'        MaxID = 1
'    Else
'        MaxID = rsMaxId.Fields(0) + 1
'    End If


End Sub


Private Sub CmdCancel_Click()

    CmdSave.Enabled = False
    CmdAddNew.Enabled = True
    Clear_Fields
End Sub

Private Sub CmdClose_Click()

''         sSQL = "Update tblRoomDetails Set (Chair=" & ChkChair.Value & ", Table=" & ChkTable.Value & ", Tubelight=" & ChkTubelight.Value & ", Holder=" & ChkHolder.Value & ", CellingFan= " & ChkCellingFan.Value & ", SleepingBed=" & ChkSleepingBed.Value & ") Where Roomno=" & TxtRoomNo - 1 & " AND HostelType='" & CmbRoomType & "'"
'         sSQL = "Select * From tblRoomDetails Where Roomno = " & TxtRoomNo - 1 & " AND HostelType = '" & CmbRoomType & "'"
'  MsgBox sSQL
'  If rsHostel.State = 1 Then rsHostel.Close
'        rsHostel.Open sSQL, con, 2, 3
'  MsgBox sSQL & vbCrLf & rsHostel.Fields(1)

'Exit Sub
    Unload Me
End Sub

Private Sub CmdSave_Click()
 
    chkValidate
    If Validate = True Then
        mess = MsgBox("Create room entry - number:" & strRoomNumber & " Capacity:" & intCapacity & " Members sex:" & strSex & " - in hostel:" & CmbHostalName.Text & "?", vbYesNo)
        If mess = vbNo Then
            Exit Sub
        End If
        
       
''        sSQL = "Select Capacity from tblHostelDetail Where hostelname = '" & CmbHostalName & "'"
''        If rsHostel.State = 1 Then rsHostel.Close
''        rsHostel.Open sSQL, Con, 2, 3
''
''        cap = CInt(rsHostel.Fields("Capacity"))
''
''        sSQL = "Select Capacity from tblRoomDetails where Hostelname = '" & CmbHostalName & "'"
''        If rsHostel.State = 1 Then rsHostel.Close
''        rsHostel.Open sSQL, Con, 2, 3
''
''        For i = 0 To rsHostel.RecordCount - 1
''            TotalCapacity = TotalCapacity + CInt(rsHostel.Fields("Capacity"))
''        Next i
''
''        If TotalCapacity >= cap Then
''                MsgBox (" Stu Create room entry - number ")
''        End If
         
     
'        sSQL = "Insert into tblRoomDetails(RoomID, HostelName, RoomNo, Capacity, Allocated, HostelType,Chair, Table, Tubelight, Holder, CellingFan, SleepingBed) Values ('" & Left$(CmbHostalName.Text, 1) & "-" & strRoomNumber & "','" & CmbHostalName.Text & "','" & strRoomNumber & "'," & intCapacity & ",0,'" & strSex & "'," & chkCheck(ChkChair) & "," & chkCheck(ChkTable) & "," & ChkTubelight.Value & "," & ChkHolder.Value & "," & ChkCellingFan.Value & "," & ChkSleepingBed.Value & ")"
        sSQL = "Insert into tblRoomDetails(RoomID, HostelName, RoomNo, Capacity, Allocated, Gender) Values ('" & Left$(CmbHostalName.Text, 1) & "-" & strRoomNumber & "','" & CmbHostalName.Text & "'," & strRoomNumber & "," & intCapacity & ",0,'" & strSex & "')"
        If rsHostel.State = 1 Then rsHostel.Close
'        MsgBox sSQL
        rsHostel.Open sSQL, Con, 2, 3
                  
'        sSQL = "Update tblRoomDetails Set (Chair=" & ChkChair.Value & ", Table=" & ChkTable.Value & ", Tubelight=" & ChkTubelight.Value & ", Holder=" & ChkHolder.Value & ", CellingFan = " & ChkCellingFan.Value & ", SleepingBed=" & ChkSleepingBed.Value & ") Where id= max(tblRoomDetails.id) "
'        If rsHostel.State = 1 Then rsHostel.Close
'        rsHostel.Open sSQL, con, 2, 3

        sSQL = "Select Capacity from tblHostelDetail where Hostelname = '" & CmbHostalName & "'"
        If rsHostel.State = 1 Then rsHostel.Close
        rsHostel.Open sSQL, Con, 2, 3
        
        cap = CInt(rsHostel.Fields("Capacity"))
        
        'update hostel parent record - total capacity
'
'        sSQL = "Update tblHostelDetail Set Capacity = " & Cap + CInt(intCapacity) & " Where HostelName = '" & CmbHostalName & "'"
'        If rsHostel.State = 1 Then rsHostel.Close
'        rsHostel.Open sSQL, Con, 2, 3
'
        MsgBox "New Room Added Successfully", vbInformation

        CmdSave.Enabled = False
        CmdAddNew.Enabled = True
         
        Clear_Fields
        Call LockText(Me)
        
    End If
End Sub

Private Sub Form_Load()
    Me.Top = 100
    Me.Left = 100
    LockText Me
    AddData
    CmbHostalName.Locked = False
'    CmbHostalName.SetFocus
End Sub

Private Sub AddData()
    
'    If rs.State = 1 Then rs.Close
'    rs.Open "Select HostelName from tblHostelDetail", con, 2, 3
''    rs.MoveFirst
'    For i = 0 To rs.RecordCount - 1
'       CmbHostalName.AddItem (rs.Fields(0).Value)
'       rs.MoveNext
'    Next i
    
    If rsHostel.State = 1 Then rsHostel.Close
    rsHostel.Open "Select * from tblHostelDetail", Con, 2, 3
'    rsHostel.MoveFirst
    CmbHostalName.Clear
    Do While Not rsHostel.EOF
        CmbHostalName.AddItem rsHostel.Fields("HostelName")         '   & " - ( " & rsHostel.Fields("Prefix") & " )"
        rsHostel.MoveNext
    Loop

End Sub


Private Sub chkValidate()
   
    strRoomNumber = Me.TxtRoomNo
    intCapacity = Val(TxtRoomCapacity)
    strSex = Me.CmbRoomType
    
    If strRoomNumber = "" Then
        MsgBox "Please enter an entry for the room number"
        TxtRoomNo.SetFocus
        Validate = False
        Exit Sub
    ElseIf intCapacity = 0 Then
        MsgBox "Please enter an entry for the room capacity"
        TxtRoomCapacity.SetFocus
        Validate = False
        Exit Sub
    ElseIf strSex = "" Then
        MsgBox "Please enter an entry for the room sex"
        CmbRoomType.SetFocus
        Validate = False
        Exit Sub
    Else
        Validate = True
    End If
    
    
End Sub

Sub Clear_Fields()
    Me.CmbHostalName.Refresh
'    CmbHostalName.Text = CmbHostalName.ListIndex
    Me.TxtRoomNo = ""
    Me.TxtRoomCapacity = ""
    Me.CmbRoomType.Clear
    
End Sub


Private Sub TxtRoomCapacity_KeyPress(KeyAscii As Integer)

    Select Case KeyAscii
            Case Asc("0") To Asc("9")
            Case 8
            Case 13
           ' Case 45
            Case Else
                KeyAscii = 0
    End Select
        
End Sub

