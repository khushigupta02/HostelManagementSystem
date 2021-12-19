VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmFeeReciept 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fee Receipt Form"
   ClientHeight    =   8610
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   13410
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "FrmFeeReciept.frx":0000
   ScaleHeight     =   8610
   ScaleWidth      =   13410
   Begin VB.Frame Frame4 
      BackColor       =   &H80000007&
      Height          =   855
      Left            =   5160
      TabIndex        =   21
      Top             =   120
      Visible         =   0   'False
      Width           =   7935
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
         Height          =   525
         Left            =   2520
         TabIndex        =   23
         Top             =   240
         Width           =   3255
      End
      Begin VB.CommandButton CmdPrint 
         Caption         =   "Print Fee Receipt"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5880
         TabIndex        =   22
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "RECEIPT ID"
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
         Left            =   120
         TabIndex        =   24
         Top             =   300
         Width           =   1440
      End
   End
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
      ForeColor       =   &H00000000&
      Height          =   7215
      Left            =   360
      TabIndex        =   0
      Top             =   1080
      Width           =   12735
      Begin VB.Frame Frame2 
         BackColor       =   &H00000000&
         Caption         =   "FEE PAID "
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   6615
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   12255
         Begin VB.TextBox TxtStudentName1 
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   525
            Left            =   11280
            TabIndex        =   30
            Top             =   1200
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.TextBox TxtTotalAmt1 
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   525
            Left            =   5040
            TabIndex        =   26
            Text            =   "0"
            Top             =   4200
            Visible         =   0   'False
            Width           =   3255
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H8000000C&
            Height          =   1335
            Left            =   240
            TabIndex        =   14
            Top             =   5040
            Width           =   11775
            Begin VB.CommandButton CmdAddNew 
               Caption         =   "ADD NEW"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   735
               Left            =   960
               TabIndex        =   18
               Top             =   360
               Width           =   1695
            End
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
               Height          =   735
               Left            =   8880
               TabIndex        =   17
               Top             =   360
               Width           =   1695
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
               Height          =   735
               Left            =   6360
               TabIndex        =   16
               Top             =   360
               Width           =   1695
            End
            Begin VB.CommandButton CmdSubmit 
               Caption         =   "SUBMIT"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   735
               Left            =   3600
               TabIndex        =   15
               Top             =   360
               Width           =   1695
            End
         End
         Begin VB.TextBox TxtRemark 
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1455
            Left            =   2760
            MultiLine       =   -1  'True
            TabIndex        =   6
            Top             =   2640
            Width           =   9255
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
            Height          =   405
            Left            =   2760
            TabIndex        =   3
            Top             =   1200
            Width           =   3255
         End
         Begin VB.TextBox TxtReceiptNo 
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
            Left            =   2760
            TabIndex        =   1
            Top             =   480
            Width           =   3255
         End
         Begin VB.TextBox TxtPaidAmount 
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
            Left            =   2760
            MaxLength       =   5
            TabIndex        =   4
            Text            =   "0"
            Top             =   1920
            Width           =   3255
         End
         Begin VB.TextBox TxtDueAmount1 
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   525
            Left            =   11280
            TabIndex        =   5
            Text            =   "0"
            Top             =   1920
            Visible         =   0   'False
            Width           =   735
         End
         Begin MSComCtl2.DTPicker DTPPaidDate 
            Height          =   375
            Left            =   8760
            TabIndex        =   2
            Top             =   480
            Width           =   3255
            _ExtentX        =   5741
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
            CurrentDate     =   43907
         End
         Begin VB.Label TxtDueAmount 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "DUE-AMOUNT"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   285
            Left            =   8760
            TabIndex        =   29
            Top             =   2040
            Width           =   1815
         End
         Begin VB.Label TxtTotalAmt 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "TOTAL FEE"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   285
            Left            =   2760
            TabIndex        =   28
            Top             =   4320
            Width           =   1410
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "TOTAL FEE"
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
            Left            =   600
            TabIndex        =   27
            Top             =   4320
            Width           =   1410
         End
         Begin VB.Label TxtStudentName 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "STUDENT NAME"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   285
            Left            =   8760
            TabIndex        =   25
            Top             =   1320
            Width           =   2070
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "PAID DATE"
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
            Left            =   6360
            TabIndex        =   19
            Top             =   525
            Width           =   1395
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   " ENROLLMENT NO."
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
            Top             =   1305
            Width           =   2430
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "STUDENT NAME"
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
            Left            =   6360
            TabIndex        =   12
            Top             =   1320
            Width           =   2070
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "REMARK"
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
            Left            =   360
            TabIndex        =   11
            Top             =   3225
            Width           =   1110
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "PAID- AMOUNT"
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
            Left            =   360
            TabIndex        =   10
            Top             =   2040
            Width           =   1950
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "DUE-AMOUNT"
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
            Left            =   6360
            TabIndex        =   9
            Top             =   2025
            Width           =   1815
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "RECEIPT NO."
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
            Height          =   405
            Left            =   360
            TabIndex        =   8
            Top             =   540
            Width           =   1620
         End
      End
   End
   Begin VB.Label Label9 
      BackColor       =   &H00000000&
      Caption         =   "FEE PAID RECEIPT"
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
      Left            =   480
      TabIndex        =   20
      Top             =   240
      Width           =   12495
   End
End
Attribute VB_Name = "FrmFeeReciept"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MaxID, Totalamt As Integer
Dim Validate As Boolean
Dim MaxFeeID As Integer
Dim sSQL As String

Private Sub CmdAddNew_Click()
    Call UnLockText(Me)
    Call ClearTextBox(Me)
    
    CmdSubmit.Visible = True
    CmdSubmit.Enabled = True
    
    CmdClose.Visible = True
    CmdClose.Enabled = True
    
    CmdCancel.Visible = True
    CmdCancel.Enabled = True
    
    CmdAddNew.Enabled = False
    
    Call fetchMaxID
    TxtReceiptNo.Text = Val(MaxID)
    TxtReceiptNo.Locked = True
    TxtTotalAmt.Caption = Totalamt
    
    TxtPaidAmount.Text = "0"
    TxtDueAmount.Caption = "0"

End Sub

Private Sub CmdCancel_Click()
    Call ClearTextBox(Me)
    CmdAddNew.Enabled = True
    CmdAddNew.Visible = True
    CmdSubmit.Enabled = False
    TxtPaidAmount.Text = "0"
    TxtDueAmount.Caption = "0"

End Sub

Private Sub CmdClose_Click()
    Unload Me
End Sub


Private Sub CmdSubmit_Click()
    
'    Dim msgresult As Integer
'    msgresult = MsgBox("FEE PAID SUCCESSFULY........" & vbNewLine & "Do you want to Print Reciept", vbQuestion + vbYesNo)
'    If msgresult = vbYes Then
'        MsgBox "Gr8..... Write Print Report Code Here", vbInformation
'    End If
'
'    Exit Sub
    
If Val(TxtPaidAmount.Text) > Val(TxtTotalAmt) Or Val(TxtPaidAmount.Text) < Val(TxtTotalAmt) Then
    MsgBox "Paid Amount Not Greater or Less then Total Fee", vbInformation
    TxtPaidAmount.SetFocus
    Exit Sub
End If

    
chkStudFee
If Validate = False Then
    Exit Sub
Else
    Call chkValidate
    
    If Validate = True Then
            
        If rs.State = 1 Then rs.Close
        rs.Open "Select * From tblPayReceipt", Con, 1, 3
        rs.AddNew
        rs.Fields("FEERECEIPTNO") = Val(TxtReceiptNo.Text)
        rs.Fields("ENROLLNO") = TxtEnrollment.Text
        rs.Fields("STUNAME") = TxtStudentName.Caption
        rs.Fields("PAIDDATE") = DTPPaidDate.Value
        rs.Fields("PAIDAMT") = Val(TxtPaidAmount.Text)
        rs.Fields("RemainAmt") = Val(TxtDueAmount.Caption)
        rs.Fields("REMARK") = TxtRemark.Text
        rs.Update
        
        
        msgresult = MsgBox("FEE PAID SUCCESSFULY........" & vbNewLine & "Do you want to Print Reciept", vbQuestion + vbYesNo)
        If msgresult = vbYes Then
            
            With DataEnvironment1.rsCmdPrintFeeReciept
                If .State Then
                    .Close
                End If
            End With
                        
            DataEnvironment1.CmdPrintFeeReciept TxtReceiptNo
            dRptPrintReciept.Show
            
        End If
         
        TxtStudentName.Caption = "STUDENT NAME"
        TxtDueAmount.Caption = "DUE AMOUNT"
 
        sSQL = "Update tblStudentInfo Set PaidReceiptNo = " & Val(TxtReceiptNo.Text) & " Where EnrollNo = '" & TxtEnrollment.Text & "'"
    '    MsgBox sSQL
        If rsStud.State = 1 Then rsStud.Close
        rsStud.Open sSQL, Con, 2, 3
         
        Call ClearTextBox(Me)
        Call LockText(Me)
         
        CmdAddNew.Enabled = True
        CmdAddNew.Visible = True
        
        CmdSubmit.Enabled = False
        CmdClose.Enabled = True
        CmdCancel.Enabled = True
    
    End If
End If

End Sub

Private Sub Form_Load()
    Me.Top = 100
    Me.Left = 100
    
    CmdSubmit.Enabled = False
    
    Call LockText(Me)
    Call fetchMaxID
    
    DTPPaidDate.Value = Date
End Sub

Private Sub fetchMaxID()
    If rsMaxId.State = 1 Then rsMaxId.Close
    rsMaxId.Open "Select Max(FeeReceiptNo) from tblPayReceipt", Con, 2, 3
    
'    MsgBox rsMaxId.RecordCount
'    Exit Sub
    
    If IsNull(rsMaxId(0)) Then
        MaxID = 1
    Else
        MaxID = rsMaxId.Fields(0) + 1
    End If

'    MaxID = rsMaxId.Fields(0).Value + 1
    

    If rsMaxId.State = 1 Then rsMaxId.Close
    rsMaxId.Open "Select Max(FeeID)from tblFeeStructure", Con, 2, 3
    MaxFeeID = rsMaxId.Fields(0).Value
    
    If IsNull(rsMaxId(0)) Then
        MaxFeeID = 0
    Else
        MaxFeeID = rsMaxId.Fields(0)
    End If

    
    If rsMaxId.State = 1 Then rsMaxId.Close
    rsMaxId.Open "Select TotalFee from tblFeeStructure where FeeID = " & MaxFeeID & "", Con, 2, 3
            
    If rsMaxId.BOF = True And rsMaxId.EOF = True Then
'        MaxID = 1
    Else
        Totalamt = rsMaxId.Fields(0).Value
        TxtTotalAmt.Caption = Totalamt
'        MaxID = rsMaxId.Fields(0).Value + 1
    End If

    
'    If IsNull(rsMaxId(0)) Then
'        MaxFeeID = 1
'    Else
'        MaxFeeID = rsMaxId.Fields(0) + 1
'    End If
    
    
End Sub

Private Sub chkValidate()
If TxtReceiptNo.Text = Trim("") Then
    MsgBox "Please Fill Up Correct Receipt No. "
    TxtReceiptNo.SetFocus
    Validate = False
    
ElseIf TxtPaidAmount.Text = Trim("") Then
    MsgBox "Please Fill Up Correct Paid Amount "
    TxtPaidAmount.SetFocus
    Validate = False
    
ElseIf TxtEnrollment.Text = Trim("") Then
    MsgBox "Please Fill Up Correct Enrollment No. "
    TxtEnrollment.SetFocus
    Validate = False

ElseIf TxtDueAmount.Caption = Trim("") Then
    MsgBox "Please Fill Up Correct Due Amount "
    TxtPaidAmount.SetFocus
    Validate = False

ElseIf TxtStudentName.Caption = Trim("") Then
    MsgBox "Please Fill Up Correct Student Name "
    TxtEnrollment.SetFocus
    Validate = False

ElseIf TxtRemark.Text = Trim("") Then
    MsgBox "Please Fill Up Correct Remark "
    TxtRemark.SetFocus
    Validate = False

Else
    Validate = True
End If

End Sub


Private Sub TxtDueAmount1_Click()
'Dim temp, temp1 As Integer
'
'temp = Totalamt Mod 10
'temp1 = Totalamt / 10
'MsgBox temp1
End Sub

Private Sub TxtEnrollment_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If rs.State = 1 Then rs.Close
        
        rs.Open "Select * from tblStudentInfo where EnrollNo= '" & TxtEnrollment.Text & "'", Con, 1, 3
        If rs.RecordCount >= 1 Then
            TxtStudentName.Caption = rs.Fields("StuName")
        Else
            MsgBox "Enrollment Not Found"
            TxtEnrollment.SetFocus
        End If
    End If
End Sub

Private Sub TxtEnrollment_LostFocus()
        If rs.State = 1 Then rs.Close
        rs.Open "Select * from tblStudentInfo where EnrollNo= '" & TxtEnrollment.Text & "'", Con, 1, 3
        If rs.RecordCount >= 1 Then
            TxtStudentName.Caption = rs.Fields("StuName")
        Else
            MsgBox "Enrollment Not Found"
            
        End If
End Sub

Private Sub TxtPaidAmount_Change()
   TxtDueAmount.Caption = Val(TxtTotalAmt) - Val(TxtPaidAmount.Text)
End Sub


Private Sub TxtPaidAmount_KeyPress(KeyAscii As Integer)
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

Private Sub chkStudFee()
   
   If rs.State = 1 Then rs.Close
        rs.Open "Select PaidReceiptNo from tblStudentInfo where EnrollNo= '" & TxtEnrollment.Text & "'", Con, 1, 3
'        rs.Open "Select RemainAmt from tblPayReceipt Where EnrollNo= '" & TxtEnrollment.Text & "'", Con, 1, 3
        
        If rs.Fields(0).Value > 0 Then
            MsgBox "Fees Already Paid For " & TxtEnrollment, vbCritical
            
'            MsgBox "Your Fees Already Paid For " & rs.Fields(0).Value, vbCritical                '& TxtEnrollment, vbCritical
            
            
            Validate = False
        Else
            Validate = True
    End If
   
End Sub
