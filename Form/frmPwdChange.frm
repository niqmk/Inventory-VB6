VERSION 5.00
Begin VB.Form frmPwdChange 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3990
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3030
   Icon            =   "frmPwdChange.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3990
   ScaleWidth      =   3030
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "Ubah"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   3
      Top             =   3480
      Width           =   1335
   End
   Begin VB.Frame fraDetail 
      Height          =   2655
      Left            =   120
      TabIndex        =   10
      Top             =   720
      Width           =   2775
      Begin VB.TextBox txtUserPwdNewConf 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   120
         MaxLength       =   20
         PasswordChar    =   "·"
         TabIndex        =   2
         Top             =   2280
         Width           =   2535
      End
      Begin VB.TextBox txtUserPwdNew 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   120
         MaxLength       =   20
         PasswordChar    =   "·"
         TabIndex        =   1
         Top             =   1440
         Width           =   2535
      End
      Begin VB.TextBox txtUserPwdOld 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   120
         PasswordChar    =   "·"
         TabIndex        =   0
         Top             =   600
         Width           =   2535
      End
      Begin VB.Label lblUserPwdNewConf 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Password Baru (Konfirmasi)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   8
         Top             =   1920
         Width           =   2415
      End
      Begin VB.Label lblUserPwdNew 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Password Baru"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   7
         Top             =   1080
         Width           =   1320
      End
      Begin VB.Label lblUserPwdOld 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Password Lama"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   1395
      End
   End
   Begin VB.Frame fraHeader 
      Height          =   615
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   2775
      Begin VB.Label txtUserId 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1440
         TabIndex        =   5
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label lblUserId 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Id Pemakai"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmPwdChange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    SetInitialization
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmPwdChange = Nothing
End Sub

Private Sub txtUserPwdOld_GotFocus()
    mdlProcedures.GotFocus Me.txtUserPwdOld
End Sub

Private Sub txtUserPwdNew_GotFocus()
    mdlProcedures.GotFocus Me.txtUserPwdNew
End Sub

Private Sub txtUserPwdNewConf_GotFocus()
    mdlProcedures.GotFocus Me.txtUserPwdNewConf
End Sub

Private Sub cmdUpdate_Click()
    If Not CheckValidation Then Exit Sub
    
    Dim rstTemp As ADODB.Recordset
    
    Set rstTemp = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "*", mdlTable.CreateTMUSER, , "UserId='" & mdlGlobal.UserAuthority.UserId & "'")
    
    With rstTemp
        If .RecordCount > 0 Then
            !UserPwd = mdlSecurity.EncryptText(mdlProcedures.RepDupText(Me.txtUserPwdNew.Text), !UserId)
            
            !UpdateId = mdlGlobal.UserAuthority.UserId
            !UpdateDate = mdlProcedures.FormatDate(Now)
            
           .Update
        End If
    End With
    
    mdlDatabase.CloseRecordset rstTemp
    
    MsgBox "Password Berhasil Diubah", vbInformation + vbOKOnly, Me.Caption
    
    Unload Me
End Sub

Private Sub SetInitialization()
    mdlProcedures.CenterWindows Me

    Me.Caption = mdlText.strPwdChange
    
    FillText
End Sub

Private Sub FillText()
    Me.txtUserId.Caption = mdlGlobal.UserAuthority.UserId
End Sub

Private Function CheckValidation() As Boolean
    If Trim(Me.txtUserPwdOld.Text) = "" Then
        Me.txtUserPwdOld.SetFocus
        
        CheckValidation = False
        
        Exit Function
    ElseIf Trim(Me.txtUserPwdNew.Text) = "" Then
        Me.txtUserPwdNew.SetFocus
        
        CheckValidation = False
        
        Exit Function
    ElseIf Trim(Me.txtUserPwdNewConf.Text) = "" Then
        Me.txtUserPwdNewConf.SetFocus
        
        CheckValidation = False
        
        Exit Function
    ElseIf Not Len(Trim(Me.txtUserPwdNew.Text)) > 5 Then
        Me.txtUserPwdNew.SetFocus
        
        CheckValidation = False
        
        Exit Function
    End If
    
    If Not mdlDatabase.IsDataCorrect( _
        mdlGlobal.conInventory, _
        "UserPwd", _
        mdlTable.CreateTMUSER, _
        "UserId='" & mdlGlobal.UserAuthority.UserId & "'", _
        mdlSecurity.EncryptText(mdlProcedures.RepDupText(Me.txtUserPwdOld.Text), mdlGlobal.UserAuthority.UserId), _
        True) Then
        Me.txtUserPwdOld.SetFocus
        
        CheckValidation = False
        
        Exit Function
    End If
    
    If Not Me.txtUserPwdNew.Text = Me.txtUserPwdNewConf.Text Then
        Me.txtUserPwdNew.SetFocus
        
        CheckValidation = False
        
        Exit Function
    End If
    
    CheckValidation = True
End Function
