VERSION 5.00
Begin VB.Form frmBackup 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4335
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8430
   Icon            =   "frmBackup.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4335
   ScaleWidth      =   8430
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdBackup 
      Caption         =   "Backup"
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
      Left            =   6840
      TabIndex        =   3
      Top             =   600
      Width           =   1455
   End
   Begin VB.TextBox txtFileBackup 
      Height          =   285
      Left            =   4080
      TabIndex        =   2
      Top             =   240
      Width           =   4215
   End
   Begin VB.Frame fraMain 
      Height          =   4095
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   3855
      Begin VB.DirListBox dirBackup 
         Height          =   3015
         Left            =   240
         TabIndex        =   1
         Top             =   840
         Width           =   3375
      End
      Begin VB.DriveListBox drvBackup 
         Height          =   315
         Left            =   240
         TabIndex        =   0
         Top             =   360
         Width           =   3375
      End
   End
End
Attribute VB_Name = "frmBackup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    SetInitialization
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmBackup = Nothing
End Sub

Private Sub drvBackup_Change()
    On Local Error GoTo ErrHandler
    
    Me.dirBackup.Path = Me.drvBackup.Drive
    
    Exit Sub
    
ErrHandler:
    Me.dirBackup.Path = mdlGlobal.strPath
End Sub

Private Sub dirBackup_Change()
    On Local Error GoTo ErrHandler
    
    Dim strDirPath As String
    
    strDirPath = dirBackup.Path
    
    If Not Right(strDirPath, 1) = "\" Then strDirPath = strDirPath & "\"
    
    Me.txtFileBackup.Text = strDirPath & mdlGlobal.strInventory & mdlProcedures.FormatDate(Now, "ddmmyyyy") & ".bak"
    
    Exit Sub
    
ErrHandler:
    Me.txtFileBackup.Text = ""
End Sub

Private Sub cmdBackup_Click()
    mdlDatabase.BackupDatabase mdlGlobal.conInventory, mdlGlobal.strInventory, Trim(Me.txtFileBackup.Text), mdlGlobal.objDatabaseInit
    
    MsgBox "Backup Database Sukses", vbInformation + vbOKOnly, Me.Caption
End Sub

Private Sub SetInitialization()
    mdlProcedures.CenterWindows Me
    
    Me.Caption = mdlText.strBackup
    
    Me.drvBackup.Drive = mdlGlobal.strPath
    Me.dirBackup.Path = mdlGlobal.strPath
    
    Dim strDirPath As String
    
    strDirPath = dirBackup.Path
    
    If Not Right(strDirPath, 1) = "\" Then strDirPath = strDirPath & "\"
    
    Me.txtFileBackup.Text = strDirPath & mdlGlobal.strInventory & mdlProcedures.FormatDate(Now, "ddmmyyyy") & ".bak"
End Sub
