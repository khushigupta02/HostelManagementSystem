VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmRefundFees 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fee Refund Form"
   ClientHeight    =   8370
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   14385
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "FrmRefundFees.frx":0000
   ScaleHeight     =   8370
   ScaleWidth      =   14385
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
      Left            =   240
      TabIndex        =   9
      Top             =   960
      Width           =   13815
      Begin VB.Frame Frame3 
         BackColor       =   &H8000000C&
         Height          =   1215
         Left            =   240
         TabIndex        =   17
         Top             =   5760
         Width           =   13335
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
            Height          =   855
            Left            =   960
            TabIndex        =   22
            Top             =   240
            Width           =   2295
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
            Height          =   855
            Left            =   10200
            TabIndex        =   20
            Top             =   240
            Width           =   2295
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
            Height          =   855
            Left            =   7320
            TabIndex        =   19
            Top             =   240
            Width           =   2055
         End
         Begin VB.CommandButton CmdRefund 
            Caption         =   "REFUND"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   4080
            TabIndex        =   18
            Top             =   240
            Width           =   2295
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00000000&
         Caption         =   " REFUND"
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
         Height          =   5055
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Width           =   13335
         Begin VB.TextBox TxtDepositAmt 
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
            ForeColor       =   &H00000040&
            Height          =   495
            Left            =   3360
            TabIndex        =   23
            Text            =   "0"
            Top             =   2040
            Width           =   3495
         End
         Begin VB.TextBox TxtStudentName 
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
            Height          =   525
            Left            =   9960
            TabIndex        =   3
            Top             =   1320
            Width           =   3135
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
            Height          =   525
            Left            =   3360
            TabIndex        =   2
            Top             =   1320
            Width           =   3495
         End
         Begin VB.TextBox TxtRefundId 
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
            Left            =   3360
            TabIndex        =   1
            Top             =   600
            Width           =   3495
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
            Height          =   1245
            Left            =   3360
            TabIndex        =   7
            Top             =   3480
            Width           =   9735
         End
         Begin VB.TextBox TxtDeductAmount 
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
            Height          =   525
            Left            =   9960
            TabIndex        =   6
            Text            =   "0"
            Top             =   2760
            Width           =   3135
         End
         Begin VB.TextBox TxtRefundAmount 
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   495
            Left            =   3360
            TabIndex        =   4
            Text            =   "0"
            Top             =   2760
            Width           =   3495
         End
         Begin MSComCtl2.DTPicker DTPRefundDate 
            Height          =   495
            Left            =   9960
            TabIndex        =   5
            Top             =   2040
            Width           =   3135
            _ExtentX        =   5530
            _ExtentY        =   873
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
            Format          =   150601731
            CurrentDate     =   43907
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "DEPOSIT-AMOUNT"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   360
            TabIndex        =   24
            Top             =   2115
            Width           =   2730
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "DEDUCT-AMOUNT"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   345
            Left            =   7080
            TabIndex        =   16
            Top             =   2850
            Width           =   2910
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "REMARK"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   345
            Left            =   360
            TabIndex        =   15
            Top             =   3930
            Width           =   1965
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "REFUND-AMOUNT"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   345
            Left            =   360
            TabIndex        =   14
            Top             =   2835
            Width           =   2865
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "STUDENT NAME"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   345
            Left            =   7080
            TabIndex        =   13
            Top             =   1410
            Width           =   2430
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ENROLLMENT NO."
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   345
            Left            =   360
            TabIndex        =   12
            Top             =   1410
            Width           =   2865
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "REFUND-DATE"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   345
            Left            =   7080
            TabIndex        =   11
            Top             =   2115
            Width           =   2430
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "REFUND-ID"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   345
            Left            =   360
            TabIndex        =   10
            Top             =   690
            Width           =   1800
         End
      End
   End
   Begin VB.Label Label9 
      BackColor       =   &H00000000&
      Caption         =   "REFUND MONEY DETAIL"
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
      TabIndex        =   21
      Top             =   120
      Width           =   13575
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   495
      Left            =   3720
      TabIndex        =   0
      Top             =   -1080
      Width           =   3015
   End
End
Attribute VB_Name = "FrmRefundFees"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MaxID  As Integer
Dim Validate As Boolean
Dim sSQL As String

Private Sub CmdAddNew_Click()
    Call UnLockText(Me)
    
    CmdAddNew.Enabled = False
    CmdRefund.Enabled = True
    
    Call fetchMaxID
    TxtRefundId.Text = MaxID
End Sub

Private Sub CmdCancel_Click()

    Call ClearTextBox(Me)
    CmdAddNew.Enabled = True
    CmdRefund.Enabled = False
'    CmdCancel.Enabled = False

End Sub

Private Sub CmdClose_Click()
    Unload Me
End Sub

Private Sub CmdRefund_Click()

    If Val(TxtRefundAmount.Text) > Val(TxtDepositAmt.Text) Or Val(TxtRefundAmount.Text) < 0 Then
        MsgBox "Refund Amount Not Greater then Deposit Amount or Less then 0", vbInformation
        TxtRefundAmount.SetFocus
        Exit Sub
    End If



chkStudFee
If Validate = False Then
    Exit Sub
Else
    Call chkValidate
    If Validate = True Then
    
        If rs.State = 1 Then rs.Close
        rs.Open "Select * From tblRefund", Con, 1, 3
        rs.AddNew
        rs.Fields("RefundID") = Val(TxtRefundId.Text)
        rs.Fields("RefundReceiptNo") = MaxID
        rs.Fields("RefundDate") = DTPRefundDate.Value
        rs.Fields("EnrollNo") = TxtEnrollment.Text
        rs.Fields("StuName") = TxtStudentName.Text
'        rs.Fields("DepositAmt") = Val(TxtDepositAmt.Text)
        rs.Fields("RefundAmt") = Val(TxtRefundAmount.Text)
        rs.Fields("DeductAmt") = Val(TxtDeductAmount.Text)
        rs.Fields("Remark") = TxtRemark.Text
        rs.Update
        msgresult = MsgBox("FEE PAID SUCCESSFULY........" & vbNewLine & "Do you want to Print Reciept", vbQuestion + vbYesNo)
        If msgresult = vbYes Then
            MsgBox "Gr8..... Write Print Report Code Here", vbInformation
            
    '        TxtStudentName.Text = "STUDENT NAME"
    '        TxtDueAmount.Caption = "DUE AMOUNT"
        End If
         
        sSQL = "Update tblStudentInfo Set PaidReceiptNo = " & 0 & ",  RefundReceiptNo = " & Val(TxtRefundId.Text) & " Where EnrollNo = '" & TxtEnrollment.Text & "'"
    '    MsgBox sSQL
        If rsStud.State = 1 Then rsStud.Close
        rsStud.Open sSQL, Con, 2, 3
 
         
        Call ClearTextBox(Me)
        Call LockText(Me)
         
        CmdAddNew.Enabled = True
        CmdAddNew.Visible = True
        
        CmdRefund.Enabled = False
        CmdClose.Enabled = True
        CmdCancel.Enabled = True
    
    End If
End If

End Sub

Private Sub Form_Load()
    Me.Top = 100
    Me.Left = 100
    
    Call LockText(Me)
'    Call fetchMaxID
    CmdRefund.Enabled = False
    DTPRefundDate.Value = Now

End Sub

Private Sub fetchMaxID()
    If rsMaxId.State = 1 Then rsMaxId.Close
    rsMaxId.Open "Select Max(RefundReceiptNo) from tblRefund", Con, 2, 3
'    MaxID = rs.Fields(0).Value + 1

    If IsNull(rsMaxId(0)) Then
        MaxID = 1
    Else
        MaxID = rsMaxId.Fields(0) + 1
    End If

End Sub


Private Sub TxtDeductAmount_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
            Case Asc("0") To Asc("9")
            Case 8
            Case 13
           ' Case 45
            Case Else
                KeyAscii = 0
    End Select
End Sub

Private Sub TxtDepositAmt_KeyPress(KeyAscii As Integer)

    Select Case KeyAscii
            Case Asc("0") To Asc("9")
            Case 8
            Case 13
           ' Case 45
            Case Else
                KeyAscii = 0
    End Select

End Sub

Private Sub TxtEnrollment_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If rs.State = 1 Then rs.Close
'        rs.Open "Select * from tblStudentRegistration where EnrollNo= '" & TxtEnrollment.Text & "'", con, 1, 3
        rs.Open "Select * from tblPayReceipt where EnrollNo= '" & TxtEnrollment.Text & "'", Con, 1, 3
        If rs.RecordCount >= 1 Then
            TxtStudentName.Text = rs.Fields("StuName")
            TxtDepositAmt.Text = rs.Fields("PaidAmt").Value
        Else
            MsgBox "Enrollment Not Found"
        End If
    End If
End Sub

Private Sub TxtEnrollment_LostFocus()

        If rs.State = 1 Then rs.Close
'        rs.Open "Select * from tblStudentRegistration where EnrollNo= '" & TxtEnrollment.Text & "'", con, 1, 3
        rs.Open "Select * from tblPayReceipt where EnrollNo= '" & TxtEnrollment.Text & "'", Con, 1, 3
        If rs.RecordCount >= 1 Then
            TxtStudentName.Text = rs.Fields("StuName")
            TxtDepositAmt.Text = rs.Fields("PaidAmt").Value
        Else
            MsgBox "Enrollment Not Found"
        End If

End Sub

Private Sub TxtRefundAmount_Change()
   TxtDeductAmount.Text = Val(TxtDepositAmt.Text) - Val(TxtRefundAmount.Text)

End Sub

Private Sub chkValidate()
    If TxtRefundId.Text = Trim("") Then
        MsgBox "Please Fill Up Correct Receipt No. "
        Validate = False
          
    ElseIf TxtEnrollment.Text = Trim("") Then
        MsgBox "Please Fill Up Correct Enrollment No. "
        TxtEnrollment.SetFocus
        Validate = False
    
    ElseIf TxtStudentName.Text = Trim("") Then
        MsgBox "Please Fill Up Correct Student Name "
        TxtEnrollment.SetFocus
        Validate = False
    
    ElseIf TxtDepositAmt.Text = Trim("") Then
        MsgBox "Please Fill Up Correct Amount "
        TxtEnrollment.SetFocus
        Validate = False
        
    ElseIf TxtRefundAmount.Text = Trim("") Then
        MsgBox "Please Fill Up Correct Due Amount "
        TxtRefundAmount.SetFocus
        Validate = False
    
    ElseIf TxtDeductAmount.Text = Trim("") Then
        MsgBox "Please Fill Up Correct Due Amount "
        TxtDeductAmount.SetFocus
        Validate = False
    
    ElseIf TxtRemark.Text = Trim("") Then
        MsgBox "Please Fill Up Correct Remark "
        TxtRemark.SetFocus
        Validate = False
    
    Else
        Validate = True
    End If

End Sub


Private Sub TxtRefundAmount_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
            Case Asc("0") To Asc("9")
            Case 8
            Case 13
           ' Case 45
            Case Else
                KeyAscii = 0
    End Select
End Sub


Private Sub chkStudFee()
   
   If rs.State = 1 Then rs.Close
        rs.Open "Select PaidReceiptNo from tblStudentInfo where EnrollNo= '" & TxtEnrollment.Text & "'", Con, 1, 3
        If rs.Fields(0).Value = 0 Then
            MsgBox "No Fees Paid or Balance For this Student " & TxtEnrollment, vbCritical
            Validate = False
        Else
            Validate = True
    End If
   
End Sub

