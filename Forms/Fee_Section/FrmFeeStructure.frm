VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmFeeStructure 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fee Structure Form"
   ClientHeight    =   9330
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   13620
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "FrmFeeStructure.frx":0000
   ScaleHeight     =   9330
   ScaleWidth      =   13620
   Begin VB.Frame Frame1 
      Height          =   6135
      Left            =   240
      TabIndex        =   11
      Top             =   960
      Width           =   13095
      Begin VB.TextBox TxtFeeID 
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
         Height          =   495
         Left            =   0
         TabIndex        =   33
         Text            =   "0.00"
         Top             =   0
         Visible         =   0   'False
         Width           =   3015
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H8000000C&
         Height          =   1335
         Left            =   240
         TabIndex        =   23
         Top             =   4560
         Width           =   12615
         Begin VB.CommandButton CmdAddNew 
            Caption         =   "ADD NEW"
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
            Left            =   1080
            TabIndex        =   29
            Top             =   360
            Width           =   1800
         End
         Begin VB.CommandButton CmdEdit 
            Caption         =   "EDIT"
            Enabled         =   0   'False
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
            Left            =   3840
            TabIndex        =   28
            Top             =   360
            Width           =   1800
         End
         Begin VB.CommandButton CmdSave 
            Caption         =   "SAVE"
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
            Left            =   1080
            TabIndex        =   27
            Top             =   360
            Width           =   1800
         End
         Begin VB.CommandButton CmdUpdate 
            Caption         =   "UPDATE"
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
            Left            =   3840
            TabIndex        =   26
            Top             =   360
            Width           =   1800
         End
         Begin VB.CommandButton CmdCancel 
            Caption         =   "CANCEL"
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
            Left            =   6840
            TabIndex        =   25
            Top             =   360
            Width           =   1800
         End
         Begin VB.CommandButton CmdClose 
            Caption         =   "CLOSE"
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
            Left            =   9840
            TabIndex        =   24
            Top             =   360
            Width           =   1800
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00000000&
         Caption         =   "FEE STRUCTURE"
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
         Height          =   4215
         Left            =   240
         TabIndex        =   12
         Top             =   240
         Width           =   12615
         Begin VB.TextBox TxtMessFee 
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
            Left            =   2400
            MaxLength       =   5
            TabIndex        =   0
            Text            =   "0"
            Top             =   720
            Width           =   3255
         End
         Begin VB.TextBox TxtNewsPaperFee 
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
            Left            =   9360
            MaxLength       =   5
            TabIndex        =   5
            Text            =   "0"
            Top             =   720
            Width           =   3015
         End
         Begin VB.TextBox TxtLightFee 
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
            Left            =   2400
            MaxLength       =   5
            TabIndex        =   1
            Text            =   "0"
            Top             =   1320
            Width           =   3255
         End
         Begin VB.TextBox TxtDevelopmentFee 
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
            Left            =   9360
            MaxLength       =   5
            TabIndex        =   6
            Text            =   "0"
            Top             =   1320
            Width           =   3015
         End
         Begin VB.TextBox TxtFoodFee 
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
            Left            =   2400
            MaxLength       =   5
            TabIndex        =   2
            Text            =   "0"
            Top             =   1920
            Width           =   3255
         End
         Begin VB.TextBox TxtStuRegistrationFee 
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
            Left            =   9360
            MaxLength       =   5
            TabIndex        =   7
            Text            =   "0"
            Top             =   1920
            Width           =   3015
         End
         Begin VB.TextBox TxtWaterCharge 
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
            Left            =   2400
            MaxLength       =   5
            TabIndex        =   3
            Text            =   "0"
            Top             =   2520
            Width           =   3255
         End
         Begin VB.TextBox TxtEletricBillFee 
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
            Left            =   9360
            MaxLength       =   5
            TabIndex        =   8
            Text            =   "0"
            Top             =   2520
            Width           =   3015
         End
         Begin VB.TextBox TxtFanCharge 
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
            Left            =   2400
            MaxLength       =   5
            TabIndex        =   4
            Text            =   "0"
            Top             =   3120
            Width           =   3255
         End
         Begin VB.TextBox TxtCommonCharge 
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
            Left            =   9360
            MaxLength       =   5
            TabIndex        =   9
            Text            =   "0"
            Top             =   3120
            Width           =   3015
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Common  Charge"
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
            Left            =   6480
            TabIndex        =   22
            Top             =   3210
            Width           =   2205
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Fan Charge"
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
            Height          =   375
            Left            =   240
            TabIndex        =   21
            Top             =   3240
            Width           =   1935
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Electricity Bill"
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
            Left            =   6480
            TabIndex        =   20
            Top             =   2610
            Width           =   1935
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Water Charge"
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
            Left            =   240
            TabIndex        =   19
            Top             =   2610
            Width           =   1785
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Student Registration"
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
            Left            =   6480
            TabIndex        =   18
            Top             =   2010
            Width           =   2610
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Food Fee"
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
            Height          =   375
            Left            =   240
            TabIndex        =   17
            Top             =   1980
            Width           =   1695
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Development Fee"
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
            Left            =   6480
            TabIndex        =   16
            Top             =   1410
            Width           =   2280
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Light Fee"
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
            Left            =   240
            TabIndex        =   15
            Top             =   1410
            Width           =   1260
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Newspaper Fee"
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
            Left            =   6480
            TabIndex        =   14
            Top             =   810
            Width           =   2025
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Mees Fee"
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
            Left            =   240
            TabIndex        =   13
            Top             =   810
            Width           =   1245
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00FFFFFF&
            BorderWidth     =   3
            X1              =   6030
            X2              =   6030
            Y1              =   240
            Y2              =   4230
         End
      End
   End
   Begin VB.TextBox TxtTotalAmt 
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
      Height          =   495
      Left            =   10200
      TabIndex        =   30
      Text            =   "0.00"
      Top             =   6720
      Visible         =   0   'False
      Width           =   3015
   End
   Begin MSFlexGridLib.MSFlexGrid MSFFeeDetails 
      Height          =   1935
      Left            =   240
      TabIndex        =   32
      Top             =   7200
      Width           =   13095
      _ExtentX        =   23098
      _ExtentY        =   3413
      _Version        =   393216
      Enabled         =   -1  'True
      SelectionMode   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "FEE STRUCTURE DETAIL"
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
      Left            =   360
      TabIndex        =   10
      Top             =   240
      Width           =   12855
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Fee"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   375
      Left            =   8760
      TabIndex        =   31
      Top             =   6960
      Visible         =   0   'False
      Width           =   1335
   End
End
Attribute VB_Name = "FrmFeeStructure"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Validate As Boolean
Dim MaxID As Integer


Private Sub CmdAddNew_Click()
    Call UnLockText(Me)
    Call ClearTextBox(Me)
    Call FeeStructureDemo
'    Call fetchMaxID
    
    CmdSave.Visible = True
    CmdSave.Enabled = True
    
    CmdClose.Enabled = True
    CmdCancel.Enabled = True
    CmdEdit.Enabled = False
    CmdUpdate.Enabled = False
    
    CmdAddNew.Enabled = False
    CmdAddNew.Visible = False

End Sub

Private Sub FeeStructureDemo()
    TxtMessFee.Text = "0"
    TxtLightFee.Text = "0"
    TxtFoodFee.Text = "0"
    TxtWaterCharge.Text = "0"
    TxtFanCharge.Text = "0"
    TxtNewsPaperFee.Text = "0"
    TxtDevelopmentFee.Text = "0"
    TxtStuRegistrationFee.Text = "0"
    TxtEletricBillFee.Text = "0"
    TxtCommonCharge.Text = "0"
    TxtTotalAmt.Text = "0"
End Sub


Private Sub CmdCancel_Click()

    Call ClearTextBox(Me)
    Call FeeStructureDemo
    
    CmdSave.Visible = False
    CmdSave.Enabled = False
    
    CmdAddNew.Enabled = True
    CmdAddNew.Visible = True
    
    CmdUpdate.Enabled = False
    
    CmdSave.Enabled = False
    CmdEdit.Enabled = False
    
    Call fetchMaxID
    ShowData
End Sub

Private Sub CmdClose_Click()
    Unload Me
End Sub

Private Sub CmdEdit_Click()
    CmdUpdate.Enabled = True
    CmdUpdate.Visible = True
    
    Call chkValidate
    If Validate = True Then
        CmdEdit.Enabled = False
        CmdEdit.Visible = False
        
        Call UnLockText(Me)
    End If

End Sub

Private Sub CmdSave_Click()
    Call chkValidate
    If Validate = True Then
        If rsFees.State = 1 Then rsFees.Close
        rsFees.Open "Select * From tblfeeStructure", Con, 1, 3
        rsFees.AddNew
        rsFees.Fields("MessFEE") = Val(TxtMessFee.Text)
        rsFees.Fields("NewsFee") = Val(TxtNewsPaperFee.Text)
        rsFees.Fields("DevelopFee") = Val(TxtDevelopmentFee.Text)
        rsFees.Fields("FoodFee") = Val(TxtFoodFee.Text)
        rsFees.Fields("ElectBil") = Val(TxtEletricBillFee.Text)
        rsFees.Fields("FanCharge") = Val(TxtFanCharge.Text)
        rsFees.Fields("WaterCharge") = Val(TxtWaterCharge.Text)
        rsFees.Fields("StuRegFee") = Val(TxtStuRegistrationFee.Text)
        rsFees.Fields("RoomCharge") = Val(TxtCommonCharge.Text)
        rsFees.Fields("LightFee") = Val(TxtLightFee.Text)
        rsFees.Fields("TotalFee") = Val(TxtTotalAmt.Text)
        rsFees.Fields("CreationDate") = Now
        rsFees.Fields("UpdatationDate") = Now
        rsFees.Update
        MsgBox "FEES ADDED SUCCESFULY......."
        
        CmdSave.Enabled = False
        CmdEdit.Enabled = False
        
        CmdAddNew.Enabled = True
        CmdAddNew.Visible = True
             
        Call ClearTextBox(Me)
        Call LockText(Me)
        Call ShowData
    
    End If

End Sub

Private Sub CmdUpdate_Click()
    Call chkValidate
    If Validate = True Then
        If rsFees.State = 1 Then rsFees.Close
        rsFees.Open "Select * From tblfeeStructure Where FeeID=" & TxtFeeID & "", Con, 2, 3
        rsFees.Fields("MessFEE") = Val(TxtMessFee.Text)
        rsFees.Fields("NewsFee") = Val(TxtNewsPaperFee.Text)
        rsFees.Fields("DevelopFee") = Val(TxtDevelopmentFee.Text)
        rsFees.Fields("FoodFee") = Val(TxtFoodFee.Text)
        rsFees.Fields("ElectBil") = Val(TxtEletricBillFee.Text)
        rsFees.Fields("FanCharge") = Val(TxtFanCharge.Text)
        rsFees.Fields("WaterCharge") = Val(TxtWaterCharge.Text)
        rsFees.Fields("StuRegFee") = Val(TxtStuRegistrationFee.Text)
        rsFees.Fields("RoomCharge") = Val(TxtCommonCharge.Text)
        rsFees.Fields("LightFee") = Val(TxtLightFee.Text)
        rsFees.Fields("TotalFee") = Val(TxtTotalAmt.Text)
        rsFees.Fields("UpdatationDate") = Now
        rsFees.Update
        MsgBox "FEES UPDATED SUCCESFULY......."
        
        Call ShowData
        CmdUpdate.Enabled = False
        CmdEdit.Enabled = True
        CmdEdit.Visible = True
        
        CmdAddNew.Enabled = True
        CmdAddNew.Visible = True
             
        Call ClearTextBox(Me)
        Call LockText(Me)
        Call fetchMaxID
    
    End If


End Sub

Private Sub Form_Load()
    Me.Top = 100
    Me.Left = 100
    
    CmdSave.Enabled = False
    CmdEdit.Enabled = False
    CmdUpdate.Enabled = False

    
    Call fetchMaxID
'    Call loadfetchData
    Call LockText(Me)
    ShowData
End Sub


Private Sub chkValidate()

    If TxtMessFee.Text = Trim("") Then
        MsgBox "Please Fill Up Correct Mees Fee "
        TxtMessFee.SetFocus
        Validate = False
        
    ElseIf TxtLightFee.Text = Trim("") Then
        MsgBox "Please Fill Up Correct Light Fee"
        TxtLightFee.SetFocus
        Validate = False
        
    ElseIf TxtFoodFee.Text = Trim("") Then
        MsgBox "Please Fill Up Correct Food Fee"
        TxtFoodFee.SetFocus
        Validate = False
          
    ElseIf TxtWaterCharge.Text = Trim("") Then
        MsgBox "Please Fill Up Correct Water Fee"
        TxtWaterCharge.SetFocus
        Validate = False
          
    ElseIf TxtFanCharge.Text = Trim("") Then
        MsgBox "Please Fill Up Correct Fan Fee"
        TxtFanCharge.SetFocus
        Validate = False
        
    ElseIf TxtNewsPaperFee.Text = Trim("") Then
        MsgBox "Please Fill Up Correct Newspaper Fee"
        TxtNewsPaperFee.SetFocus
        Validate = False
        
    ElseIf TxtDevelopmentFee.Text = Trim("") Then
        MsgBox "Please Fill Up Correct Development Fee"
        TxtDevelopmentFee.SetFocus
        Validate = False
            
    ElseIf TxtStuRegistrationFee.Text = Trim("") Then
        MsgBox "Please Fill Up Correct Student Registration Fee"
        TxtStuRegistrationFee.SetFocus
        Validate = False
           
    ElseIf TxtEletricBillFee.Text = Trim("") Then
        MsgBox "Please Fill Up Correct Electricity Fee"
        TxtEletricBillFee.SetFocus
        Validate = False
          
    ElseIf TxtCommonCharge.Text = Trim("") Then
        MsgBox "Please Fill Up Correct Common Charges"
        TxtCommonCharge.SetFocus
        Validate = False
           
    Else
        Validate = True
    End If

End Sub

Private Sub fetchMaxID()
    If rsMaxId.State = 1 Then rsMaxId.Close
    rsMaxId.Open "Select Max(FeeID) from tblFeeStructure", Con, 2, 3
    
    If IsNull(rsMaxId(0)) Then
        MaxID = 1
    Else
        MaxID = rsMaxId.Fields(0) + 1
    End If

End Sub

Private Sub loadfetchData()
'

    TxtFeeID.Text = MSFFeeDetails.TextMatrix(MSFFeeDetails.RowSel, 0)

    TxtMessFee.Text = MSFFeeDetails.TextMatrix(MSFFeeDetails.RowSel, 1)
    TxtLightFee.Text = MSFFeeDetails.TextMatrix(MSFFeeDetails.RowSel, 2)
    TxtFoodFee.Text = MSFFeeDetails.TextMatrix(MSFFeeDetails.RowSel, 3)
    TxtWaterCharge.Text = MSFFeeDetails.TextMatrix(MSFFeeDetails.RowSel, 4)
    TxtFanCharge.Text = MSFFeeDetails.TextMatrix(MSFFeeDetails.RowSel, 5)
    TxtNewsPaperFee.Text = MSFFeeDetails.TextMatrix(MSFFeeDetails.RowSel, 6)
    TxtDevelopmentFee.Text = MSFFeeDetails.TextMatrix(MSFFeeDetails.RowSel, 7)
    TxtStuRegistrationFee.Text = MSFFeeDetails.TextMatrix(MSFFeeDetails.RowSel, 8)
    TxtEletricBillFee.Text = MSFFeeDetails.TextMatrix(MSFFeeDetails.RowSel, 9)
    TxtCommonCharge.Text = MSFFeeDetails.TextMatrix(MSFFeeDetails.RowSel, 10)
    TxtTotalAmt.Text = MSFFeeDetails.TextMatrix(MSFFeeDetails.RowSel, 11)

End Sub

Public Sub ShowData()
    MSFFeeDetails.Refresh
    Dim i As Integer
    MSFFeeDetails.Cols = 12
    MSFFeeDetails.Rows = 1
    
'     MSFFeeDetails.f
   
    
    MSFFeeDetails.TextMatrix(0, 0) = "Fees ID"
    MSFFeeDetails.ColWidth(0) = max_width + 800
    MSFFeeDetails.TextMatrix(0, 1) = "Mess Fee"
    MSFFeeDetails.ColWidth(1) = max_width + 1000
    MSFFeeDetails.TextMatrix(0, 2) = "Light Fee"
    MSFFeeDetails.ColWidth(2) = max_width + 1000
    MSFFeeDetails.TextMatrix(0, 3) = "Food Fee"
    MSFFeeDetails.ColWidth(3) = max_width + 1000
    MSFFeeDetails.TextMatrix(0, 4) = "Water Charge"
    MSFFeeDetails.ColWidth(4) = max_width + 1300

    MSFFeeDetails.TextMatrix(0, 5) = "Fan Charge"
    MSFFeeDetails.ColWidth(5) = max_width + 1100
    MSFFeeDetails.TextMatrix(0, 6) = "News Paper Fee"
    MSFFeeDetails.ColWidth(6) = max_width + 1500
    MSFFeeDetails.TextMatrix(0, 7) = "Development Fee"
    MSFFeeDetails.ColWidth(7) = max_width + 1600
    MSFFeeDetails.TextMatrix(0, 8) = "Student Registration Fee"
    MSFFeeDetails.ColWidth(8) = max_width + 2200
    MSFFeeDetails.TextMatrix(0, 9) = "Electricity Bill Charge"
    MSFFeeDetails.ColWidth(9) = max_width + 2000
    MSFFeeDetails.TextMatrix(0, 10) = "Room Charge"
    MSFFeeDetails.ColWidth(10) = max_width + 1300
    MSFFeeDetails.TextMatrix(0, 11) = "Total Fees"
    MSFFeeDetails.ColWidth(11) = max_width + 1300
    
    If rsview.State = 1 Then rsview.Close
    rsview.Open "select * from tblFeeStructure", Con, adOpenForwardOnly, adLockReadOnly
    MSFFeeDetails.Rows = rsview.RecordCount + 1
    For i = 1 To rsview.RecordCount
    MSFFeeDetails.TextMatrix(i, 0) = rsview.Fields("FeeID").Value
    MSFFeeDetails.TextMatrix(i, 1) = rsview.Fields("MessFee").Value
    MSFFeeDetails.TextMatrix(i, 2) = rsview.Fields("LightFee").Value
    MSFFeeDetails.TextMatrix(i, 3) = rsview.Fields("FoodFee").Value
    MSFFeeDetails.TextMatrix(i, 4) = rsview.Fields("WaterCharge").Value
    MSFFeeDetails.TextMatrix(i, 5) = rsview.Fields("FanCharge").Value
    MSFFeeDetails.TextMatrix(i, 6) = rsview.Fields("NewsFee").Value
    MSFFeeDetails.TextMatrix(i, 7) = rsview.Fields("DevelopFee").Value
    MSFFeeDetails.TextMatrix(i, 8) = rsview.Fields("StuRegFee").Value
    MSFFeeDetails.TextMatrix(i, 9) = rsview.Fields("ElectBil").Value
    MSFFeeDetails.TextMatrix(i, 10) = rsview.Fields("RoomCharge").Value
    MSFFeeDetails.TextMatrix(i, 11) = rsview.Fields("TotalFee").Value

    rsview.MoveNext
    Next i

End Sub


Private Sub MSFFeeDetails_Click()
    CmdEdit.Enabled = True
    loadfetchData
End Sub

Private Sub TxtCommonCharge_Change()
    TxtTotalAmt = Val(TxtMessFee.Text) + Val(TxtLightFee.Text) + Val(TxtFoodFee.Text) + Val(TxtWaterCharge.Text) + Val(TxtFanCharge.Text) + Val(TxtNewsPaperFee.Text) + Val(TxtDevelopmentFee.Text) + Val(TxtStuRegistrationFee.Text) + Val(TxtEletricBillFee.Text) + Val(TxtCommonCharge.Text)
End Sub

Private Sub TxtDevelopmentFee_Change()
    TxtTotalAmt = Val(TxtMessFee.Text) + Val(TxtLightFee.Text) + Val(TxtFoodFee.Text) + Val(TxtWaterCharge.Text) + Val(TxtFanCharge.Text) + Val(TxtNewsPaperFee.Text) + Val(TxtDevelopmentFee.Text) + Val(TxtStuRegistrationFee.Text) + Val(TxtEletricBillFee.Text) + Val(TxtCommonCharge.Text)
End Sub

Private Sub TxtEletricBillFee_Change()
    TxtTotalAmt = Val(TxtMessFee.Text) + Val(TxtLightFee.Text) + Val(TxtFoodFee.Text) + Val(TxtWaterCharge.Text) + Val(TxtFanCharge.Text) + Val(TxtNewsPaperFee.Text) + Val(TxtDevelopmentFee.Text) + Val(TxtStuRegistrationFee.Text) + Val(TxtEletricBillFee.Text) + Val(TxtCommonCharge.Text)
End Sub

Private Sub TxtFanCharge_Change()
    TxtTotalAmt = Val(TxtMessFee.Text) + Val(TxtLightFee.Text) + Val(TxtFoodFee.Text) + Val(TxtWaterCharge.Text) + Val(TxtFanCharge.Text) + Val(TxtNewsPaperFee.Text) + Val(TxtDevelopmentFee.Text) + Val(TxtStuRegistrationFee.Text) + Val(TxtEletricBillFee.Text) + Val(TxtCommonCharge.Text)
End Sub

Private Sub TxtFoodFee_Change()
    TxtTotalAmt = Val(TxtMessFee.Text) + Val(TxtLightFee.Text) + Val(TxtFoodFee.Text) + Val(TxtWaterCharge.Text) + Val(TxtFanCharge.Text) + Val(TxtNewsPaperFee.Text) + Val(TxtDevelopmentFee.Text) + Val(TxtStuRegistrationFee.Text) + Val(TxtEletricBillFee.Text) + Val(TxtCommonCharge.Text)
End Sub

Private Sub TxtLightFee_Change()
    TxtTotalAmt = Val(TxtMessFee.Text) + Val(TxtLightFee.Text) + Val(TxtFoodFee.Text) + Val(TxtWaterCharge.Text) + Val(TxtFanCharge.Text) + Val(TxtNewsPaperFee.Text) + Val(TxtDevelopmentFee.Text) + Val(TxtStuRegistrationFee.Text) + Val(TxtEletricBillFee.Text) + Val(TxtCommonCharge.Text)
End Sub

Private Sub TxtMessFee_Change()
    TxtTotalAmt = Val(TxtMessFee.Text) + Val(TxtLightFee.Text) + Val(TxtFoodFee.Text) + Val(TxtWaterCharge.Text) + Val(TxtFanCharge.Text) + Val(TxtNewsPaperFee.Text) + Val(TxtDevelopmentFee.Text) + Val(TxtStuRegistrationFee.Text) + Val(TxtEletricBillFee.Text) + Val(TxtCommonCharge.Text)
End Sub

Private Sub TxtNewsPaperFee_Change()
    TxtTotalAmt = Val(TxtMessFee.Text) + Val(TxtLightFee.Text) + Val(TxtFoodFee.Text) + Val(TxtWaterCharge.Text) + Val(TxtFanCharge.Text) + Val(TxtNewsPaperFee.Text) + Val(TxtDevelopmentFee.Text) + Val(TxtStuRegistrationFee.Text) + Val(TxtEletricBillFee.Text) + Val(TxtCommonCharge.Text)
End Sub

Private Sub TxtStuRegistrationFee_Change()
    TxtTotalAmt = Val(TxtMessFee.Text) + Val(TxtLightFee.Text) + Val(TxtFoodFee.Text) + Val(TxtWaterCharge.Text) + Val(TxtFanCharge.Text) + Val(TxtNewsPaperFee.Text) + Val(TxtDevelopmentFee.Text) + Val(TxtStuRegistrationFee.Text) + Val(TxtEletricBillFee.Text) + Val(TxtCommonCharge.Text)
End Sub

Private Sub TxtWaterCharge_Change()
    TxtTotalAmt = Val(TxtMessFee.Text) + Val(TxtLightFee.Text) + Val(TxtFoodFee.Text) + Val(TxtWaterCharge.Text) + Val(TxtFanCharge.Text) + Val(TxtNewsPaperFee.Text) + Val(TxtDevelopmentFee.Text) + Val(TxtStuRegistrationFee.Text) + Val(TxtEletricBillFee.Text) + Val(TxtCommonCharge.Text)
End Sub

Private Sub TxtWaterCharge_KeyPress(KeyAscii As Integer)
    num_only (KeyAscii)
End Sub

Private Sub TxtCommonCharge_KeyPress(KeyAscii As Integer)
    num_only (KeyAscii)
End Sub

Private Sub TxtDevelopmentFee_KeyPress(KeyAscii As Integer)
    num_only (KeyAscii)
End Sub

Private Sub TxtEletricBillFee_KeyPress(KeyAscii As Integer)
    num_only (KeyAscii)
End Sub

Private Sub TxtFanCharge_KeyPress(KeyAscii As Integer)
    num_only (KeyAscii)
End Sub

Private Sub TxtFoodFee_KeyPress(KeyAscii As Integer)
    num_only (KeyAscii)
End Sub

Private Sub TxtLightFee_KeyPress(KeyAscii As Integer)
    num_only (KeyAscii)
End Sub

Private Sub TxtMessFee_KeyPress(KeyAscii As Integer)
    num_only (KeyAscii)
End Sub

Private Sub TxtNewsPaperFee_KeyPress(KeyAscii As Integer)
    num_only (KeyAscii)
End Sub

Private Sub TxtStuRegistrationFee_KeyPress(KeyAscii As Integer)
    num_only (KeyAscii)
End Sub


