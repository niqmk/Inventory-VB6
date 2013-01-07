VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmProfile 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7470
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7455
   Icon            =   "frmProfile.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7470
   ScaleWidth      =   7455
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdRegistration 
      Caption         =   "&Registrasi"
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
      Left            =   5520
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   6960
      Width           =   1815
   End
   Begin TabDlg.SSTab sstMain 
      Height          =   6735
      Left            =   120
      TabIndex        =   30
      Top             =   120
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   11880
      _Version        =   393216
      Tab             =   1
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Data Perusahaan"
      TabPicture(0)   =   "frmProfile.frx":1F8A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "fraMain(0)"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Logo"
      TabPicture(1)   =   "frmProfile.frx":1FA6
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "fraMain(1)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Administrator"
      TabPicture(2)   =   "frmProfile.frx":1FC2
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraMain(2)"
      Tab(2).ControlCount=   1
      Begin VB.Frame fraMain 
         Height          =   6255
         Index           =   1
         Left            =   120
         TabIndex        =   28
         Top             =   360
         Width           =   6975
         Begin MSComDlg.CommonDialog cdlPicture 
            Left            =   240
            Top             =   4080
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.CommandButton cmdBrowse 
            Caption         =   "Gambar"
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
            Index           =   1
            Left            =   5880
            TabIndex        =   7
            Top             =   5760
            Width           =   975
         End
         Begin VB.PictureBox picWallpaper 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   4935
            Left            =   2640
            ScaleHeight     =   327
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   279
            TabIndex        =   23
            TabStop         =   0   'False
            Top             =   720
            Width           =   4215
         End
         Begin VB.CommandButton cmdBrowse 
            Caption         =   "Gambar"
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
            Index           =   0
            Left            =   1440
            TabIndex        =   8
            Top             =   2760
            Width           =   975
         End
         Begin VB.PictureBox picLogo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   1935
            Left            =   240
            ScaleHeight     =   127
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   143
            TabIndex        =   22
            TabStop         =   0   'False
            Top             =   720
            Width           =   2175
         End
         Begin VB.Label lblWallpaper 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Gambar Wallpaper"
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
            Left            =   2640
            TabIndex        =   20
            Top             =   360
            Width           =   1620
         End
         Begin VB.Label lblLogo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "127 x 127 px"
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
            Index           =   1
            Left            =   240
            TabIndex        =   21
            Top             =   2760
            Width           =   1125
         End
         Begin VB.Label lblLogo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Logo Perusahaan"
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
            Index           =   0
            Left            =   240
            TabIndex        =   19
            Top             =   360
            Width           =   1515
         End
      End
      Begin VB.Frame fraMain 
         Height          =   6255
         Index           =   2
         Left            =   -74880
         TabIndex        =   29
         Top             =   360
         Width           =   6975
         Begin VB.TextBox txtUserPwdConf 
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
            Left            =   240
            MaxLength       =   20
            PasswordChar    =   "·"
            TabIndex        =   11
            Top             =   2400
            Width           =   2535
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
            Left            =   240
            MaxLength       =   8
            TabIndex        =   9
            Top             =   720
            Width           =   1095
         End
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
            IMEMode         =   3  'DISABLE
            Left            =   240
            MaxLength       =   20
            PasswordChar    =   "·"
            TabIndex        =   10
            Top             =   1560
            Width           =   2535
         End
         Begin VB.Label lblUserPwdConf 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Administrator Password (Konfirmasi)"
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
            Left            =   240
            TabIndex        =   26
            Top             =   2040
            Width           =   3165
         End
         Begin VB.Label lblUserPwd 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Administrator Password"
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
            Left            =   240
            TabIndex        =   25
            Top             =   1200
            Width           =   2070
         End
         Begin VB.Label lblUserId 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Administrator Id"
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
            Left            =   240
            TabIndex        =   24
            Top             =   360
            Width           =   1365
         End
      End
      Begin VB.Frame fraMain 
         Height          =   6255
         Index           =   0
         Left            =   -74880
         TabIndex        =   27
         Top             =   360
         Width           =   6975
         Begin VB.TextBox txtNPWP 
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
            Left            =   120
            MaxLength       =   50
            TabIndex        =   6
            Top             =   5880
            Width           =   6735
         End
         Begin VB.TextBox txtFax 
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
            Left            =   120
            MaxLength       =   50
            TabIndex        =   5
            Top             =   5160
            Width           =   6735
         End
         Begin VB.TextBox txtPhone 
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
            Left            =   120
            MaxLength       =   50
            TabIndex        =   4
            Top             =   4440
            Width           =   6735
         End
         Begin VB.TextBox txtEmail 
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
            Left            =   120
            MaxLength       =   50
            TabIndex        =   3
            Top             =   3720
            Width           =   6735
         End
         Begin VB.TextBox txtWebsite 
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
            Left            =   120
            MaxLength       =   50
            TabIndex        =   2
            Top             =   3000
            Width           =   6735
         End
         Begin VB.TextBox txtAddress 
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   885
            Left            =   120
            MaxLength       =   150
            MultiLine       =   -1  'True
            TabIndex        =   1
            Top             =   1680
            Width           =   6735
         End
         Begin VB.TextBox txtName 
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   120
            MaxLength       =   100
            MultiLine       =   -1  'True
            TabIndex        =   0
            Top             =   600
            Width           =   6735
         End
         Begin VB.Label lblNPWP 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "NPWP"
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
            TabIndex        =   18
            Top             =   5520
            Width           =   600
         End
         Begin VB.Label lblFax 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fax"
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
            TabIndex        =   17
            Top             =   4800
            Width           =   330
         End
         Begin VB.Label lblPhone 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Telepon"
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
            TabIndex        =   16
            Top             =   4080
            Width           =   675
         End
         Begin VB.Label lblEmail 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "E-Mail"
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
            TabIndex        =   15
            Top             =   3360
            Width           =   555
         End
         Begin VB.Label lblWebsite 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Website"
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
            TabIndex        =   14
            Top             =   2640
            Width           =   720
         End
         Begin VB.Label lblAddress 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Alamat"
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
            TabIndex        =   13
            Top             =   1320
            Width           =   615
         End
         Begin VB.Label lblName 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nama Perusahaan"
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
            TabIndex        =   12
            Top             =   240
            Width           =   1605
         End
      End
   End
End
Attribute VB_Name = "frmProfile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private strLogo As String
Private strWallpaper As String

Private Sub Form_Activate()
    Me.sstMain.Tab = 0
End Sub

Private Sub Form_Load()
    SetInitialization
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    mdlDatabase.CloseConnection mdlGlobal.conInventory
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmProfile = Nothing
End Sub

Private Sub sstMain_Click(PreviousTab As Integer)
    If Not Me.sstMain.Tab = PreviousTab Then
        Select Case sstMain.Tab
            Case 0:
                Me.txtName.SetFocus
            Case 1:
                Me.cmdBrowse(0).SetFocus
            Case 2:
                If Me.fraMain(2).Enabled Then
                    Me.txtUserId.SetFocus
                End If
        End Select
    End If
End Sub

Private Sub txtName_GotFocus()
    mdlProcedures.GotFocus Me.txtName
End Sub

Private Sub txtAddress_GotFocus()
    mdlProcedures.GotFocus Me.txtAddress
End Sub

Private Sub txtWebsite_GotFocus()
    mdlProcedures.GotFocus Me.txtWebsite
End Sub

Private Sub txtEmail_GotFocus()
    mdlProcedures.GotFocus Me.txtEmail
End Sub

Private Sub txtUserId_GotFocus()
    mdlProcedures.GotFocus Me.txtUserId
End Sub

Private Sub txtUserPwd_GotFocus()
    mdlProcedures.GotFocus Me.txtUserPwd
End Sub

Private Sub txtUserPwdConf_GotFocus()
    mdlProcedures.GotFocus Me.txtUserPwdConf
End Sub

Private Sub cmdBrowse_Click(Index As Integer)
    On Local Error GoTo ErrHandler
    
    With Me.cdlPicture
        .ShowOpen
        
        If Not Trim(.FileName) = "" Then
            Dim strExtension() As String
            
            strExtension = Split(.FileTitle, ".")
            
            Select Case Index
                Case 0:
                    strLogo = Trim(.FileName)
                    'strLogo = mdlGlobal.strPath & mdlGlobal.LOGO_IMAGE_FILE & "." & strExtension(UBound(strExtension))
                    
                    'mdlGlobal.fso.CopyFile .FileName, strLogo
                    
                    Set Me.picLogo.Picture = LoadPicture(strLogo)
                Case 1:
                    strWallpaper = Trim(.FileName)
                    'strWallpaper = mdlGlobal.strPath & mdlGlobal.WALLPAPER_IMAGE_FILE & "." & strExtension(UBound(strExtension))
                    
                    'mdlGlobal.fso.CopyFile .FileName, strWallpaper
                    
                    Set Me.picWallpaper.Picture = LoadPicture(strWallpaper)
            End Select
        End If
    End With
    
ErrHandler:
End Sub

Private Sub cmdRegistration_Click()
    If Not IsValidRegistration Then Exit Sub
    
    If Me.fraMain(2).Enabled Then
        If SetAdministration Then
            SaveAdministrator
            
            SetRegistry
            
            Unload Me
        End If
    Else
        SetRegistry
        
        Unload Me
    End If
End Sub

Private Sub SetInitialization()
    mdlProcedures.CenterWindows Me, False
    
    Me.Caption = mdlText.strProfile

    With Me.cdlPicture
        .CancelError = True
        .Filter = "All Images |*.JPG;*.BMP"
    End With
    
    Dim rstTemp As ADODB.Recordset
    
    Set rstTemp = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "*", mdlTable.CreateTMUSER(True), False)
    
    With rstTemp
        While Not .EOF
            If Trim(mdlSecurity.DecryptText(Trim(!UserType), !UserId)) = mdlGlobal.strAdministrator Then
                Me.fraMain(2).Enabled = False
                
                GoTo ExitSub
            End If
            
            .MoveNext
        Wend
    End With
    
ExitSub:
End Sub

Private Function IsValidRegistration() As Boolean
    Dim blnValid As Boolean
    
    blnValid = True
    
    If Trim(Me.txtName.Text) = "" Then
        Me.sstMain.Tab = 0
    
        Me.txtName.SetFocus
        
        blnValid = False
    ElseIf Trim(strLogo) = "" Then
        Me.sstMain.Tab = 1
        
        Me.cmdBrowse(0).SetFocus
        
        blnValid = False
    End If
    
    IsValidRegistration = blnValid
End Function

Private Sub SetRegistry()
    Dim lngRegistry As Long
    Dim lngRegKey As Long
    Dim lngType As Long
    Dim lngSize As Long
    
    On Local Error GoTo ErrHandler

    lngRegistry = _
        mdlRegistry.OpenRegistry(mdlRegistry.HKEY_CURRENT_USER, mdlRegistry.KEYS_SYS_INFO, lngRegKey)
        
    If Not lngRegistry = 0 Then
        lngRegistry = mdlRegistry.WriteToRegistry(mdlRegistry.HKEY_CURRENT_USER, mdlRegistry.KEYS_SYS_INFO)
        
        lngRegistry = _
            mdlRegistry.WriteValueRegistry(mdlRegistry.HKEY_CURRENT_USER, mdlRegistry.KEYS_SYS_INFO, mdlGlobal.COMPANY_REGISTRY, Trim(Me.txtName.Text))
        lngRegistry = _
            mdlRegistry.WriteValueRegistry(mdlRegistry.HKEY_CURRENT_USER, mdlRegistry.KEYS_SYS_INFO, mdlGlobal.ADDRESS_REGISTRY, Trim(Me.txtAddress.Text))
        lngRegistry = _
            mdlRegistry.WriteValueRegistry(mdlRegistry.HKEY_CURRENT_USER, mdlRegistry.KEYS_SYS_INFO, mdlGlobal.WEBSITE_REGISTRY, Trim(Me.txtWebsite.Text))
        lngRegistry = _
            mdlRegistry.WriteValueRegistry(mdlRegistry.HKEY_CURRENT_USER, mdlRegistry.KEYS_SYS_INFO, mdlGlobal.EMAIL_REGISTRY, Trim(Me.txtEmail.Text))
        lngRegistry = _
            mdlRegistry.WriteValueRegistry(mdlRegistry.HKEY_CURRENT_USER, mdlRegistry.KEYS_SYS_INFO, mdlGlobal.PHONE_REGISTRY, Trim(Me.txtPhone.Text))
        lngRegistry = _
            mdlRegistry.WriteValueRegistry(mdlRegistry.HKEY_CURRENT_USER, mdlRegistry.KEYS_SYS_INFO, mdlGlobal.FAX_REGISTRY, Trim(Me.txtFax.Text))
        lngRegistry = _
            mdlRegistry.WriteValueRegistry(mdlRegistry.HKEY_CURRENT_USER, mdlRegistry.KEYS_SYS_INFO, mdlGlobal.NPWP_REGISTRY, Trim(Me.txtNPWP.Text))
        lngRegistry = _
            mdlRegistry.WriteValueRegistry(mdlRegistry.HKEY_CURRENT_USER, mdlRegistry.KEYS_SYS_INFO, mdlGlobal.LOGO_REGISTRY, strLogo)
        lngRegistry = _
            mdlRegistry.WriteValueRegistry(mdlRegistry.HKEY_CURRENT_USER, mdlRegistry.KEYS_SYS_INFO, mdlGlobal.WALLPAPER_REGISTRY, strWallpaper)
    Else
        GoTo ErrHandler
    End If

    lngRegistry = mdlRegistry.CloseRegistry(lngRegKey)
    
    Exit Sub

ErrHandler:
    MsgBox "Profile Tidak Dapat Disimpan", vbCritical, Me.Caption
End Sub

Private Function SetAdministration() As Boolean
    If Trim(Me.txtUserId.Text) = "" Then
        Me.sstMain.Tab = 2
        
        Me.txtUserId.SetFocus
        
        SetAdministration = False
        
        Exit Function
    ElseIf Trim(Me.txtUserPwd.Text) = "" Then
        Me.sstMain.Tab = 2
        
        Me.txtUserPwd.SetFocus
        
        SetAdministration = False
        
        Exit Function
    ElseIf Trim(Me.txtUserPwdConf.Text) = "" Then
        Me.sstMain.Tab = 2
        
        Me.txtUserPwd.SetFocus
        
        SetAdministration = False
        
        Exit Function
    ElseIf Not Len(Trim(Me.txtUserPwd.Text)) > 5 Then
        Me.sstMain.Tab = 2
        
        Me.txtUserPwd.SetFocus
        
        SetAdministration = False
        
        Exit Function
    End If
    
    If Not Me.txtUserPwd.Text = Me.txtUserPwdConf.Text Then
        MsgBox "Password Tidak Sama", vbExclamation, Me.Caption
        
        Me.sstMain.Tab = 2
        
        Me.txtUserPwd.SetFocus
        
        SetAdministration = False
        
        Exit Function
    End If
    
    SetAdministration = True
End Function

Private Sub SaveAdministrator()
    Dim mUserId As String
    Dim mUserPwd As String
    
    mUserId = mdlProcedures.RepDupText(Trim(Me.txtUserId.Text))
    mUserPwd = mdlProcedures.RepDupText(Me.txtUserPwd.Text)

    Dim rstTemp As ADODB.Recordset
    
    Set rstTemp = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "*", mdlTable.CreateTMUSER, , "UserId='" & mUserId & "'")
    
    With rstTemp
        If Not .RecordCount > 0 Then
            .AddNew
            
            !UserId = mUserId
            !UserName = mUserId
            
            !UserType = mdlSecurity.EncryptText(mdlGlobal.strAdministrator, !UserId)
            
            !CreateId = mUserId
            !CreateDate = mdlProcedures.FormatDate(Now)
        End If
        
        !UserPwd = mdlSecurity.EncryptText(mUserPwd, !UserId)
        !UpdateId = mUserId
        !UpdateDate = mdlProcedures.FormatDate(Now)
        
        .Update
    End With
    
    mdlDatabase.CloseRecordset rstTemp
End Sub
