VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2340
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   6030
   ControlBox      =   0   'False
   Icon            =   "frmLogin.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2340
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock wskLogin 
      Left            =   2160
      Top             =   1800
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.PictureBox picCompany 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   1935
      Left            =   120
      ScaleHeight     =   127
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   127
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   240
      Width           =   1935
   End
   Begin VB.CommandButton cmdLogin 
      Caption         =   "&Login"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4680
      TabIndex        =   2
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Batal"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3360
      TabIndex        =   3
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Frame fraLogin 
      Height          =   1335
      Left            =   2160
      TabIndex        =   6
      Top             =   240
      Width           =   3735
      Begin VB.TextBox txtUserPwd 
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
         Left            =   1320
         TabIndex        =   1
         Top             =   840
         Width           =   2295
      End
      Begin VB.TextBox txtUserId 
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
         Left            =   1320
         TabIndex        =   0
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label lblUserPwd 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
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
         TabIndex        =   5
         Top             =   840
         Width           =   855
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
         Top             =   360
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private intTryLogin As Integer

Private Sub Form_Load()
    SetInitialization
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmLogin = Nothing
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        TryToLogin
    End If
End Sub

Private Sub cmdLogin_Click()
    TryToLogin
End Sub

Private Sub cmdCancel_Click()
    If mdiMain.Parent Then
        mdiMain.Exists = False
    Else
        mdlDatabase.CloseConnection mdlGlobal.conInventory
        mdlDatabase.CloseConnection mdlGlobal.conFinance
        mdlDatabase.CloseConnection mdlGlobal.conAccounting
        
        Set mdlGlobal.fso = Nothing
    End If
    
    Unload Me
End Sub

Private Sub txtUserId_GotFocus()
    mdlProcedures.GotFocus Me.txtUserId
End Sub

Private Sub txtUserPwd_GotFocus()
    mdlProcedures.GotFocus Me.txtUserPwd
End Sub

Private Sub SetInitialization()
    mdlAPI.SetWindowPos Me.hwnd, -1, 0, 0, 0, 0, &H2 Or &H1
    
    mdlProcedures.CenterWindows Me, False
    
    Me.Caption = mdlText.strLogin & " (" & mdlGlobal.strCompanyText & ")"

    Me.txtUserPwd.PasswordChar = Chr(183)
    
    LogoInitialize
    
    intTryLogin = 0
    
    SetRegistry
End Sub

Private Sub SetRegistry()
    Dim lngRegistry As Long
    Dim lngRegKey As Long
    Dim lngType As Long
    Dim lngSize As Long
    
    lngRegistry = _
        mdlRegistry.OpenRegistry(mdlRegistry.HKEY_CURRENT_USER, mdlRegistry.KEYS_SYS_INFO_RUN, lngRegKey)
    
    lngRegistry = mdlRegistry.WriteToRegistry(mdlRegistry.HKEY_CURRENT_USER, mdlRegistry.KEYS_SYS_INFO_RUN)
        
    lngRegistry = _
        mdlRegistry.WriteValueRegistry(mdlRegistry.HKEY_CURRENT_USER, mdlRegistry.KEYS_SYS_INFO_RUN, mdlGlobal.AUTORUN_REGISTRY, mdlGlobal.strPath & mdlGlobal.strInventory & ".exe")

    lngRegistry = mdlRegistry.CloseRegistry(lngRegKey)
End Sub

Private Sub TryToLogin()
    If CheckValidation Then
        If mdlGlobal.UserAuthority Is Nothing Then Set mdlGlobal.UserAuthority = New clsAuthority
        
        mdlGlobal.UserAuthority.UserId = Trim(Me.txtUserId.Text)
        
        mdlGlobal.UserAuthority.SetLogin Trim(Me.txtUserId.Text), Me.wskLogin.LocalIP
        
        If mdiMain.Parent Then
            mdiMain.Exists = True
        Else
            mdiMain.Show
        End If

        Unload Me
    Else
        If Not Trim(Me.txtUserId.Text) = "" Or Not Trim(Me.txtUserPwd.Text) = "" Then
            intTryLogin = intTryLogin + 1
            
            If intTryLogin >= 3 Then
                If mdiMain.Parent Then
                    mdiMain.Exists = False
                Else
                    mdlDatabase.CloseConnection mdlGlobal.conInventory
                    mdlDatabase.CloseConnection mdlGlobal.conFinance
                    mdlDatabase.CloseConnection mdlGlobal.conAccounting
                    
                    Set mdlGlobal.fso = Nothing
                    Set mdlGlobal.UserAuthority = Nothing
                End If
                
                Unload Me
            End If
        End If
    End If
End Sub

Private Sub LogoInitialize()
    Dim lngRegistry As Long
    Dim lngRegKey As Long
    Dim lngType As Long
    Dim lngSize As Long
    
    On Local Error GoTo ErrHandler

    lngRegistry = _
        mdlRegistry.OpenRegistry(mdlRegistry.HKEY_CURRENT_USER, mdlRegistry.KEYS_SYS_INFO, lngRegKey)

    If lngRegistry = 0 Then
        lngRegistry = _
            mdlRegistry.ReadValueRegistry(lngRegKey, mdlGlobal.LOGO_REGISTRY, lngType, mdlGlobal.strLogoImageText, lngSize)
        
        mdlGlobal.strLogoImageText = mdlProcedures.RepRegistryUnknown(Trim(CStr(mdlGlobal.strLogoImageText)))
        
        Set Me.picCompany.Picture = LoadPicture(mdlGlobal.strLogoImageText)
        
        Me.picCompany.ToolTipText = mdlGlobal.strCompanyText
    End If

    lngRegistry = mdlRegistry.CloseRegistry(lngRegKey)
    
    Exit Sub

ErrHandler:
End Sub

Private Function CheckValidation() As Boolean
    Dim blnValid As Boolean
    
    blnValid = True
    
    If Trim(Me.txtUserId.Text) = "" Then
        Me.txtUserId.SetFocus
    
        blnValid = False
    ElseIf Trim(Me.txtUserPwd.Text) = "" Then
        Me.txtUserPwd.SetFocus
        
        blnValid = False
    Else
        blnValid = CheckExist
    End If
    
    CheckValidation = blnValid
End Function

Private Function CheckExist() As Boolean
    Dim blnExist As Boolean
    
    blnExist = True

    Dim mUserId As String
    
    mUserId = mdlProcedures.RepDupText(Trim(Me.txtUserId.Text))

    Dim rstTemp As ADODB.Recordset
    
    Set rstTemp = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "UserId, UserPwd", mdlTable.CreateTMUSER(True), False, "UserId='" & mUserId & "'")
    
    If rstTemp.RecordCount > 0 Then
        If Trim(mdlSecurity.DecryptText(Trim(rstTemp!UserPwd), rstTemp!UserId)) = Me.txtUserPwd.Text Then
            Me.txtUserId.Text = Trim(rstTemp!UserId)
        
            blnExist = True
        Else
            Me.txtUserPwd.Text = ""
            
            Me.txtUserPwd.SetFocus
        
            blnExist = False
        End If
    Else
        Me.txtUserId.Text = ""
        
        Me.txtUserId.SetFocus
        
        blnExist = False
    End If
    
    mdlDatabase.CloseRecordset rstTemp
    
    CheckExist = blnExist
End Function
