VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmChatDetail 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   870
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   7365
   Icon            =   "frmChatDetail.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   870
   ScaleWidth      =   7365
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtChat 
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
      TabIndex        =   0
      Top             =   480
      Width           =   5895
   End
   Begin VB.CommandButton cmdChat 
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
      Left            =   6120
      TabIndex        =   1
      Top             =   360
      Width           =   1095
   End
   Begin MSWinsockLib.Winsock wskChat 
      Left            =   5520
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label lblMain 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Kepada :"
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
      TabIndex        =   2
      Top             =   120
      Width           =   780
   End
End
Attribute VB_Name = "frmChatDetail"
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

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Me.wskChat.Close
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmChatDetail = Nothing
End Sub

Private Sub wskChat_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    MsgBox Number & " : " & Description, vbOKOnly + vbExclamation, Me.Caption
End Sub

Private Sub wskChat_SendComplete()
    Me.wskChat.Close
    
    Unload Me
End Sub

Private Sub cmdChat_Click()
    frmChat.ReceiveChatText mdlGlobal.UserAuthority.UserId, Me.txtChat.Text
    
    Me.wskChat.SendData mdlGlobal.UserAuthority.UserId & " | " & Me.txtChat.Text
End Sub

Private Sub SetInitialization()
    mdlProcedures.CenterWindows Me, False
    
    Me.Caption = mdlText.strChat
    
    Me.lblMain.Caption = Me.lblMain.Caption & " " & frmChat.UserIP
    
    Me.wskChat.Connect frmChat.UserIP, mdlGlobal.intPort
End Sub
