VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2295
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5325
   Icon            =   "frmAbout.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2295
   ScaleWidth      =   5325
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   3840
      TabIndex        =   2
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Frame fraMain 
      Height          =   1095
      Left            =   1320
      TabIndex        =   3
      Top             =   600
      Width           =   3855
      Begin VB.Label lblCompany 
         ForeColor       =   &H00C00000&
         Height          =   735
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   3615
      End
   End
   Begin VB.PictureBox picProvider 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   240
      Picture         =   "frmAbout.frx":1F8A
      ScaleHeight     =   825
      ScaleWidth      =   825
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   240
      Width           =   855
   End
   Begin VB.Label lblProvider 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MJStone Software Production"
      Height          =   195
      Left            =   1320
      TabIndex        =   0
      Top             =   240
      Width           =   2115
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    SetInitialization
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmAbout = Nothing
End Sub

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub SetInitialization()
    mdlProcedures.CenterWindows Me, False
    
    Me.Caption = mdlText.strAbout
    
    Me.lblCompany.Caption = mdlGlobal.strCompanyText
End Sub
