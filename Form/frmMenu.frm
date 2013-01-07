VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMenu 
   BackColor       =   &H00808080&
   BorderStyle     =   0  'None
   ClientHeight    =   4725
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2460
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4725
   ScaleWidth      =   2460
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picRecycle 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   840
      ScaleHeight     =   705
      ScaleWidth      =   705
      TabIndex        =   5
      Top             =   3600
      Width           =   735
   End
   Begin MSComctlLib.ImageList imlRecycle 
      Left            =   1560
      Top             =   3720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":1B52
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lblRecycle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   3240
      Width           =   1695
   End
   Begin VB.Label lblMenu 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Master"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   1695
   End
   Begin VB.Label lblMenu 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Manajemen"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   1
      Left            =   360
      TabIndex        =   1
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Label lblMenu 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Transaksi"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   2
      Left            =   360
      TabIndex        =   2
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Label lblMenu 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Laporan"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   3
      Left            =   360
      TabIndex        =   3
      Top             =   2520
      Width           =   1695
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum MenuMode
    [MasterMode]
    [MISMode]
    [TransactionMode]
    [ReportMode]
End Enum

Private Enum RecycleStatus
    [EmptyRecycle] = 1
    [FillRecycle]
End Enum

Private objRecycleStatus As RecycleStatus

Private intLeftMenu As Integer
Private intTopMenu As Integer

Private Sub Form_Load()
    SetInitialization
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmMenu = Nothing
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim intCounter As Integer
    
    For intCounter = 0 To Me.lblMenu.Count - 1
        Me.lblMenu(intCounter).BackColor = &HFFFF00
    Next intCounter
End Sub

Private Sub picRecycle_Click()
    If Not mdlGlobal.UserAuthority.IsMenuAccess(mdlTable.CreateTHRECYCLE) Then Exit Sub
    If Not mdlTransaction.IsRecycleExists Then Exit Sub
    
    mdlProcedures.ShowForm frmTHRECYCLE, mdiMain.mnuCloseAll.Checked
End Sub

Private Sub lblMenu_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.lblMenu(Index).BackColor = &HFF00&
End Sub

Private Sub lblMenu_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = MouseButtonConstants.vbLeftButton Then
        intLeftMenu = Me.lblMenu(Index).Left
        intTopMenu = Me.lblMenu(Index).Top + Me.Top + mdiMain.tlbMain.Height
    
        Select Case Index
            Case MasterMode:
                mdlProcedures.ShowForm frmMaster, False
            Case MISMode:
                mdlProcedures.ShowForm frmMIS, False
            Case TransactionMode:
                mdlProcedures.ShowForm frmTransaction, False
            Case ReportMode:
                mdlProcedures.ShowForm frmReport, False
        End Select
    End If
End Sub

Private Sub SetInitialization()
    mdlProcedures.CenterWindows Me, False, True
    
    mdlAPI.SetLayeredWindow Me.hwnd, True
    mdlAPI.SetLayeredWindowAttributes Me.hwnd, 0, 70, &H2
    
    mdlAPI.SmoothForm Me, 30
    
    SetRecycle
End Sub

Public Sub SetRecycle()
    If mdlTransaction.IsRecycleExists Then
        Me.lblRecycle.Caption = "Recycle Bin"
        
        objRecycleStatus = FillRecycle
    Else
        Me.lblRecycle.Caption = ""
        
        objRecycleStatus = EmptyRecycle
    End If
    
    Me.picRecycle.Picture = Me.imlRecycle.ListImages(objRecycleStatus).Picture
End Sub

Public Property Get LeftMenu() As Integer
    LeftMenu = intLeftMenu
End Property

Public Property Get TopMenu() As Integer
    TopMenu = intTopMenu
End Property
