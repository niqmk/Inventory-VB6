VERSION 5.00
Begin VB.Form frmMaster 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   0  'None
   ClientHeight    =   3060
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8325
   Icon            =   "frmMaster.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3060
   ScaleWidth      =   8325
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Label lblMenu 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Gudang"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   13
      Left            =   5640
      TabIndex        =   13
      Top             =   1920
      Width           =   2535
   End
   Begin VB.Label lblMenu 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mata Uang"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   12
      Left            =   5640
      TabIndex        =   12
      Top             =   1320
      Width           =   2535
   End
   Begin VB.Label lblMenu 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Divisi"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   11
      Left            =   5640
      TabIndex        =   11
      Top             =   720
      Width           =   2535
   End
   Begin VB.Label lblMenu 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Jabatan"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   10
      Left            =   5640
      TabIndex        =   10
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label lblMenu 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Karyawan"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   9
      Left            =   2880
      TabIndex        =   9
      Top             =   2520
      Width           =   2535
   End
   Begin VB.Label lblMenu 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Pemasok"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   8
      Left            =   2880
      TabIndex        =   8
      Top             =   1920
      Width           =   2535
   End
   Begin VB.Label lblMenu 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Customer"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   7
      Left            =   2880
      TabIndex        =   7
      Top             =   1320
      Width           =   2535
   End
   Begin VB.Label lblMenu 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Satuan"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   6
      Left            =   2880
      TabIndex        =   6
      Top             =   720
      Width           =   2535
   End
   Begin VB.Label lblMenu 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Merk"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   5
      Left            =   2880
      TabIndex        =   5
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label lblMenu 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Harga Jual"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   2
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   2535
   End
   Begin VB.Label lblMenu 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Daftar Harga"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   2535
   End
   Begin VB.Label lblMenu 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Jenis"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   4
      Left            =   120
      TabIndex        =   4
      Top             =   2520
      Width           =   2535
   End
   Begin VB.Label lblMenu 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Grup"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   3
      Left            =   120
      TabIndex        =   3
      Top             =   1920
      Width           =   2535
   End
   Begin VB.Label lblMenu 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Barang"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2535
   End
End
Attribute VB_Name = "frmMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum MenuList
    [Item]
    [PriceList]
    [ItemPrice]
    [Group]
    [Category]
    [Brand]
    [Unity]
    [Customer]
    [Vendor]
    [Employee]
    [JobType]
    [Division]
    [Currenci]
    [Warehouse]
End Enum

Private Sub Form_Load()
    SetInitialization
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmMaster = Nothing
End Sub

Private Sub Form_LostFocus()
    Unload Me
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim intCounter As Integer
    
    For intCounter = 0 To Me.lblMenu.Count - 1
        Me.lblMenu(intCounter).BackColor = &HC0E0FF
    Next intCounter
End Sub

Private Sub lblMenu_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = MouseButtonConstants.vbLeftButton Then
        Unload Me
        
        Select Case Index
            Case Item:
                mdlProcedures.ShowForm frmTMITEM, mdiMain.CloseAll, , frmMenu.Name
            Case PriceList:
                mdlProcedures.ShowForm frmTMPRICELIST, mdiMain.CloseAll, , frmMenu.Name
            Case ItemPrice:
                mdlProcedures.ShowForm frmTMITEMPRICE, mdiMain.CloseAll, , frmMenu.Name
            Case Group:
                mdlProcedures.ShowForm frmTMGROUP, mdiMain.CloseAll, , frmMenu.Name
            Case Category:
                mdlProcedures.ShowForm frmTMCATEGORY, mdiMain.CloseAll, , frmMenu.Name
            Case Brand:
                mdlProcedures.ShowForm frmTMBRAND, mdiMain.CloseAll, , frmMenu.Name
            Case Unity:
                mdlProcedures.ShowForm frmTMUNITY, mdiMain.CloseAll, , frmMenu.Name
            Case Customer:
                mdlProcedures.ShowForm frmTMCUSTOMER, mdiMain.CloseAll, , frmMenu.Name
            Case Vendor:
                mdlProcedures.ShowForm frmTMVENDOR, mdiMain.CloseAll, , frmMenu.Name
            Case Employee:
                mdlProcedures.ShowForm frmTMEMPLOYEE, mdiMain.CloseAll, , frmMenu.Name
            Case JobType:
                mdlProcedures.ShowForm frmTMJOBTYPE, mdiMain.CloseAll, , frmMenu.Name
            Case Division:
                mdlProcedures.ShowForm frmTMDIVISION, mdiMain.CloseAll, , frmMenu.Name
            Case Currenci:
                mdlProcedures.ShowForm frmTMCURRENCY, mdiMain.CloseAll, , frmMenu.Name
            Case Warehouse:
                mdlProcedures.ShowForm frmTMWAREHOUSE, mdiMain.CloseAll, , frmMenu.Name
        End Select
    End If
End Sub

Private Sub lblMenu_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.lblMenu(Index).BackColor = &HFFC0FF
End Sub

Private Sub SetInitialization()
    Me.Move frmMenu.LeftMenu, frmMenu.TopMenu
    
    mdlAPI.SetLayeredWindow Me.hwnd, True
    
    mdlAPI.SetLayeredWindowAttributes Me.hwnd, 0, (255 * 60) / 100, &H2
    
    mdlAPI.SmoothForm Me, 30
    
    SetClient
End Sub

Private Sub SetClient()
    Me.lblMenu(Item).Visible = mdiMain.mnuItem.Visible
    Me.lblMenu(PriceList).Visible = mdiMain.mnuItem.Visible
    Me.lblMenu(ItemPrice).Visible = mdiMain.mnuItem.Visible
    Me.lblMenu(Group).Visible = mdiMain.mnuItem.Visible
    Me.lblMenu(Category).Visible = mdiMain.mnuItem.Visible
    Me.lblMenu(Brand).Visible = mdiMain.mnuItem.Visible
    Me.lblMenu(Unity).Visible = mdiMain.mnuItem.Visible
        
    If mdiMain.mnuItem.Visible Then
        Me.lblMenu(Item).Visible = mdiMain.mnuTMITEM.Visible
        Me.lblMenu(PriceList).Visible = mdiMain.mnuTMITEMPRICE.Visible
        Me.lblMenu(ItemPrice).Visible = mdiMain.mnuTMITEMPRICE.Visible
        Me.lblMenu(Group).Visible = mdiMain.mnuTMGROUP.Visible
        Me.lblMenu(Category).Visible = mdiMain.mnuTMCATEGORY.Visible
        Me.lblMenu(Brand).Visible = mdiMain.mnuTMBRAND.Visible
        Me.lblMenu(Unity).Visible = mdiMain.mnuTMUNITY.Visible
    End If
    
    Me.lblMenu(Customer).Visible = mdiMain.mnuExternal.Visible
    Me.lblMenu(Vendor).Visible = mdiMain.mnuExternal.Visible
    
    If mdiMain.mnuExternal.Visible Then
        Me.lblMenu(Customer).Visible = mdiMain.mnuTMCUSTOMER.Visible
        Me.lblMenu(Vendor).Visible = mdiMain.mnuTMVENDOR.Visible
    End If
    
    Me.lblMenu(Employee).Visible = mdiMain.mnuInternal.Visible
    Me.lblMenu(JobType).Visible = mdiMain.mnuInternal.Visible
    Me.lblMenu(Division).Visible = mdiMain.mnuInternal.Visible
    Me.lblMenu(Currenci).Visible = mdiMain.mnuInternal.Visible
    Me.lblMenu(Warehouse).Visible = mdiMain.mnuInternal.Visible
    
    If mdiMain.mnuInternal.Visible Then
        Me.lblMenu(Employee).Visible = mdiMain.mnuTMEMPLOYEE.Visible
        Me.lblMenu(JobType).Visible = mdiMain.mnuTMJOBTYPE.Visible
        Me.lblMenu(Division).Visible = mdiMain.mnuTMDIVISION.Visible
        Me.lblMenu(Currenci).Visible = mdiMain.mnuTMCURRENCY.Visible
        Me.lblMenu(Warehouse).Visible = mdiMain.mnuTMWAREHOUSE.Visible
    End If
End Sub
