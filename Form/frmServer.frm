VERSION 5.00
Begin VB.Form frmServer 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5655
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5790
   Icon            =   "frmServer.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5655
   ScaleWidth      =   5790
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSave 
      Caption         =   "Simpan"
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
      Left            =   4560
      TabIndex        =   8
      Top             =   5160
      Width           =   1095
   End
   Begin VB.Frame fraServer 
      Height          =   4935
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   5535
      Begin VB.TextBox txtPassword 
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
         TabIndex        =   7
         Top             =   4560
         Width           =   5055
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
         TabIndex        =   6
         Top             =   3840
         Width           =   5055
      End
      Begin VB.OptionButton optServer 
         Caption         =   "MySQL ODBC 5.1"
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
         Index           =   4
         Left            =   240
         TabIndex        =   4
         Top             =   2160
         Width           =   2415
      End
      Begin VB.OptionButton optServer 
         Caption         =   "MICROSOFT ACCESS"
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
         Index           =   3
         Left            =   240
         TabIndex        =   3
         Top             =   1680
         Width           =   2415
      End
      Begin VB.OptionButton optServer 
         Caption         =   "SQL SERVER EXPRESS"
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
         Index           =   2
         Left            =   240
         TabIndex        =   2
         Top             =   1200
         Width           =   2535
      End
      Begin VB.OptionButton optServer 
         Caption         =   "SQL SERVER 2000"
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
         Left            =   240
         TabIndex        =   1
         Top             =   720
         Width           =   2415
      End
      Begin VB.OptionButton optServer 
         Caption         =   "SQL SERVER 7"
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
         Left            =   240
         TabIndex        =   0
         Top             =   240
         Width           =   2415
      End
      Begin VB.TextBox txtServer 
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
         TabIndex        =   5
         Top             =   3120
         Width           =   5055
      End
      Begin VB.Label lblPassword 
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
         Left            =   240
         TabIndex        =   11
         Top             =   4200
         Width           =   855
      End
      Begin VB.Label lblUserId 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User Id"
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
         TabIndex        =   10
         Top             =   3480
         Width           =   615
      End
      Begin VB.Label lblServer 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Server / IP Address"
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
         TabIndex        =   9
         Top             =   2760
         Width           =   2250
      End
   End
End
Attribute VB_Name = "frmServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private objMode As SQLDATABASE

Private Sub Form_Load()
    SetInitialization
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmServer = Nothing
End Sub

Private Sub optServer_Click(Index As Integer)
    Me.txtServer.Text = ""
    
    Select Case Index
        Case SQLSERVER7:
            Me.txtServer.Enabled = True
            
            objMode = SQLSERVER7
        Case SQLSERVER2000:
            objMode = SQLSERVER2000
        Case SQLEXPRESS:
            objMode = SQLEXPRESS
        Case MSACCESS:
            Dim strFile As String
            
            strFile = mdlGlobal.strPath & mdlGlobal.strInventory & ".mdb"
            
            If mdlGlobal.fso.FileExists(strFile) Then
                Me.txtServer.Enabled = False
                
                Me.txtServer.Text = strFile
                Me.txtServer.Enabled = False
                
                objMode = MSACCESS
            Else
                MsgBox "Database Access di " & strFile & " Tidak Ditemukan" & vbCrLf & _
                    "Silahkan Buat Database Access Terlebih Dahulu", vbExclamation + vbOKOnly, Me.Caption
                    
                Me.optServer(0).SetFocus
            End If
        Case MYSQL:
            objMode = MYSQL
    End Select
End Sub

Private Sub cmdSave_Click()
    If objMode = SQLSERVER7 Or objMode = SQLSERVER2000 Or objMode = SQLEXPRESS Then
        If Trim(Me.txtServer.Text) = "" Then
            MsgBox "Server : 127.0.0.1", vbInformation + vbOKOnly, Me.Caption
            
            If Not SetRegistry("127.0.0.1", Trim(Me.txtUserId.Text), Trim(Me.txtPassword.Text)) Then Exit Sub
        Else
            If Not SetRegistry(Trim(Me.txtServer.Text), Trim(Me.txtUserId.Text), Trim(Me.txtPassword.Text)) Then Exit Sub
        End If
    Else
        If Not SetRegistry(Trim(Me.txtServer.Text), Trim(Me.txtUserId.Text), Trim(Me.txtPassword.Text)) Then Exit Sub
    End If
    
    Unload Me
End Sub

Private Sub SetInitialization()
    mdlProcedures.CenterWindows Me, False
    
    Me.txtPassword.PasswordChar = Chr(183)
    
    Me.Caption = mdlText.strServerSetting
End Sub

Private Function SetRegistry( _
    Optional ByVal strServerName As String = "", _
    Optional ByVal strUserId As String = "", _
    Optional ByVal strPassword As String = "") As Boolean
    Dim lngRegistry As Long
    Dim lngRegKey As Long
    Dim lngType As Long
    Dim lngSize As Long
    
    On Local Error GoTo ErrHandler

    lngRegistry = _
        mdlRegistry.OpenRegistry(mdlRegistry.HKEY_CURRENT_USER, mdlRegistry.KEYS_SYS_INFO_SERVER1, lngRegKey)
        
    If Not lngRegistry = 0 Then
        lngRegistry = mdlRegistry.WriteToRegistry(mdlRegistry.HKEY_CURRENT_USER, mdlRegistry.KEYS_SYS_INFO_SERVER1)
        
        lngRegistry = _
            mdlRegistry.WriteValueRegistry(mdlRegistry.HKEY_CURRENT_USER, mdlRegistry.KEYS_SYS_INFO_SERVER1, mdlGlobal.SERVER_REGISTRY, strServerName)
            
        If Trim(strUserId) = "" Then
            lngRegistry = _
                mdlRegistry.WriteValueRegistry(mdlRegistry.HKEY_CURRENT_USER, mdlRegistry.KEYS_SYS_INFO_SERVER1, mdlGlobal.USERID_REGISTRY, "")
        Else
            lngRegistry = _
                mdlRegistry.WriteValueRegistry(mdlRegistry.HKEY_CURRENT_USER, mdlRegistry.KEYS_SYS_INFO_SERVER1, mdlGlobal.USERID_REGISTRY, mdlSecurity.EncryptText(strUserId, mdlGlobal.PUBLIC_KEY))
        End If
        
        If Trim(strPassword) = "" Then
            lngRegistry = _
                mdlRegistry.WriteValueRegistry(mdlRegistry.HKEY_CURRENT_USER, mdlRegistry.KEYS_SYS_INFO_SERVER1, mdlGlobal.PASSWORD_REGISTRY, "")
        Else
            lngRegistry = _
                mdlRegistry.WriteValueRegistry(mdlRegistry.HKEY_CURRENT_USER, mdlRegistry.KEYS_SYS_INFO_SERVER1, mdlGlobal.PASSWORD_REGISTRY, mdlSecurity.EncryptText(strPassword, mdlGlobal.PUBLIC_KEY))
        End If
        
        lngRegistry = _
            mdlRegistry.WriteValueRegistry(mdlRegistry.HKEY_CURRENT_USER, mdlRegistry.KEYS_SYS_INFO_SERVER1, mdlGlobal.DATABASE_REGISTRY, CStr(objMode))
            
        SetRegistry = True
    Else
        GoTo ErrHandler
    End If

    lngRegistry = mdlRegistry.CloseRegistry(lngRegKey)
    
    Exit Function

ErrHandler:
    MsgBox "Server Tidak Dapat Disimpan", vbCritical, Me.Caption
    
    SetRegistry = False
End Function
