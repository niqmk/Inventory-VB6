VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmFax 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1950
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   7950
   Icon            =   "frmFax.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1950
   ScaleWidth      =   7950
   ShowInTaskbar   =   0   'False
   Begin MSComDlg.CommonDialog cdlFile 
      Left            =   120
      Top             =   1440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Kirim"
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
      Left            =   6720
      TabIndex        =   4
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Frame fraMain 
      Height          =   1215
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   7695
      Begin VB.CommandButton cmdFileName 
         Caption         =   "..."
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
         Left            =   7200
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   720
         Width           =   375
      End
      Begin VB.CommandButton cmdFax 
         Caption         =   "..."
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
         Left            =   7200
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox txtFileName 
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
         Left            =   960
         TabIndex        =   2
         Top             =   720
         Width           =   6135
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
         Left            =   960
         MaxLength       =   50
         TabIndex        =   0
         Top             =   240
         Width           =   6135
      End
      Begin VB.Label lblFileName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Filename"
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
         Top             =   720
         Width           =   795
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
         TabIndex        =   5
         Top             =   240
         Width           =   330
      End
   End
End
Attribute VB_Name = "frmFax"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private blnParent As Boolean

Private Sub Form_Load()
    SetInitialization
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmFax = Nothing
End Sub

Private Sub txtFax_GotFocus()
    mdlProcedures.GotFocus Me.txtFax
End Sub

Private Sub txtFileName_GotFocus()
    mdlProcedures.GotFocus Me.txtFileName
End Sub

Private Sub cmdFax_Click()
    If blnParent Then Exit Sub
    
    blnParent = True
    
    mdlProcedures.CornerWindows Me
    
    mdlProcedures.ShowForm frmBRWTMCUSTOMER, False, True
End Sub

Private Sub cmdFileName_Click()
    On Local Error GoTo ErrHandler
    
    With Me.cdlFile
        .ShowOpen
        
        If Not Trim(.FileName) = "" Then
            Dim strExtension() As String
            
            strExtension = Split(.FileTitle, ".")
        End If
    End With
    
ErrHandler:
End Sub

Private Sub cmdSend_Click()
    If Not CheckValidation Then Exit Sub
    
    SendFax
End Sub

Private Sub SetInitialization()
    mdlProcedures.CenterWindows Me
    
    Me.Caption = mdlText.strFax
    
    With Me.cdlFile
        .CancelError = True
        .Filter = "All Files |*.*"
    End With
    
    blnParent = False
End Sub

Private Function CheckValidation() As Boolean
    If mdlProcedures.RepDupText(Trim(Me.txtFax.Text)) = "" Then
        MsgBox "Fax Harap Diisi", vbOKOnly + vbExclamation, Me.Caption
        
        Me.txtFax.SetFocus

        CheckValidation = False

        Exit Function
    ElseIf mdlProcedures.RepDupText(Trim(Me.txtFileName.Text)) = "" Then
        MsgBox "Filename Harap Diisi", vbOKOnly + vbExclamation, Me.Caption
        
        Me.txtFileName.SetFocus

        CheckValidation = False

        Exit Function
    End If
    
    CheckValidation = True
End Function

Private Sub SendFax()
'
End Sub

Public Property Get Parent() As Boolean
    Parent = blnParent
End Property

Public Property Get Fax() As String
    Fax = Trim(Me.txtFax.Text)
End Property

Public Property Let Parent(ByVal blnEnable As Boolean)
    blnParent = blnEnable
    
    If Not blnEnable Then mdlProcedures.CenterWindows Me
End Property

Public Property Let Fax(ByVal strFax As String)
    Me.txtFax.Text = strFax
End Property
