VERSION 5.00
Begin VB.Form frmMIS 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   ClientHeight    =   7590
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8085
   Icon            =   "frmMIS.frx":0000
   KeyPreview      =   -1  'True
   ScaleHeight     =   7590
   ScaleWidth      =   8085
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Label lblMenuSell 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Sales Summary"
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
      Left            =   5640
      TabIndex        =   28
      Top             =   4080
      Width           =   2295
   End
   Begin VB.Label lblMenuMaster 
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
      TabIndex        =   17
      Top             =   2400
      Width           =   2295
   End
   Begin VB.Label lblMenuWarehouse 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mutasi Barang"
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
      TabIndex        =   22
      Top             =   6480
      Width           =   2295
   End
   Begin VB.Label lblMenuBuy 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Purchase Order"
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
      Left            =   2880
      TabIndex        =   23
      Top             =   4080
      Width           =   2295
   End
   Begin VB.Label lblMenu 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Gudang"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   3600
      Width           =   2355
   End
   Begin VB.Label lblMenuBuy 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Delivery Order"
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
      Left            =   2880
      TabIndex        =   24
      Top             =   4680
      Width           =   2295
   End
   Begin VB.Label lblMenu 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Pembelian"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   2880
      TabIndex        =   2
      Top             =   3600
      Width           =   2355
   End
   Begin VB.Label lblMenuSell 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Purchase Order"
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
      Left            =   5640
      TabIndex        =   29
      Top             =   4680
      Width           =   2295
   End
   Begin VB.Label lblMenuSell 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Sales Order"
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
      Left            =   5640
      TabIndex        =   30
      Top             =   5280
      Width           =   2295
   End
   Begin VB.Label lblMenuSell 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Surat Jalan"
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
      Left            =   5640
      TabIndex        =   31
      Top             =   5880
      Width           =   2295
   End
   Begin VB.Label lblMenuSell 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Faktur"
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
      Left            =   5640
      TabIndex        =   32
      Top             =   6480
      Width           =   2295
   End
   Begin VB.Label lblMenuBuy 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Surat Jalan"
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
      Left            =   2880
      TabIndex        =   25
      Top             =   5280
      Width           =   2295
   End
   Begin VB.Label lblMenuBuy 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Faktur"
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
      Left            =   2880
      TabIndex        =   26
      Top             =   5880
      Width           =   2295
   End
   Begin VB.Label lblMenu 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Penjualan"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   5640
      TabIndex        =   3
      Top             =   3600
      Width           =   2355
   End
   Begin VB.Label lblMenuWarehouse 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Stok Awal"
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
      TabIndex        =   18
      Top             =   4080
      Width           =   2295
   End
   Begin VB.Label lblMenuWarehouse 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Sisa Stok"
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
      TabIndex        =   19
      Top             =   4680
      Width           =   2295
   End
   Begin VB.Label lblMenuWarehouse 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Pemasukkan Barang"
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
      TabIndex        =   20
      Top             =   5280
      Width           =   2295
   End
   Begin VB.Label lblMenuWarehouse 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Pengeluaran Barang"
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
      TabIndex        =   21
      Top             =   5880
      Width           =   2295
   End
   Begin VB.Label lblMenuBuy 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Retur"
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
      Left            =   2880
      TabIndex        =   27
      Top             =   6480
      Width           =   2295
   End
   Begin VB.Label lblMenuSell 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Retur"
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
      Left            =   5640
      TabIndex        =   33
      Top             =   7080
      Width           =   2295
   End
   Begin VB.Label lblMenuMaster 
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
      TabIndex        =   4
      Top             =   600
      Width           =   2295
   End
   Begin VB.Label lblMenuMaster 
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
      Index           =   3
      Left            =   120
      TabIndex        =   7
      Top             =   2400
      Width           =   2295
   End
   Begin VB.Label lblMenuMaster 
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
      Index           =   4
      Left            =   120
      TabIndex        =   8
      Top             =   3000
      Width           =   2295
   End
   Begin VB.Label lblMenuMaster 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Harga Barang"
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
      TabIndex        =   5
      Top             =   1200
      Width           =   2295
   End
   Begin VB.Label lblMenuMaster 
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
      Index           =   2
      Left            =   120
      TabIndex        =   6
      Top             =   1800
      Width           =   2295
   End
   Begin VB.Label lblMenuMaster 
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
      Index           =   5
      Left            =   2880
      TabIndex        =   9
      Top             =   600
      Width           =   2295
   End
   Begin VB.Label lblMenuMaster 
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
      Index           =   6
      Left            =   2880
      TabIndex        =   10
      Top             =   1200
      Width           =   2295
   End
   Begin VB.Label lblMenuMaster 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Transaksi Customer"
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
      TabIndex        =   11
      Top             =   1800
      Width           =   2295
   End
   Begin VB.Label lblMenuMaster 
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
      TabIndex        =   12
      Top             =   2400
      Width           =   2295
   End
   Begin VB.Label lblMenuMaster 
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
      TabIndex        =   13
      Top             =   3000
      Width           =   2295
   End
   Begin VB.Label lblMenuMaster 
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
      TabIndex        =   14
      Top             =   600
      Width           =   2295
   End
   Begin VB.Label lblMenuMaster 
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
      TabIndex        =   15
      Top             =   1200
      Width           =   2295
   End
   Begin VB.Label lblMenuMaster 
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
      TabIndex        =   16
      Top             =   1800
      Width           =   2295
   End
   Begin VB.Label lblMenu 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Master"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7995
   End
End
Attribute VB_Name = "frmMIS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum MenuMasterList
    [Item]
    [ItemPrice]
    [Group]
    [Category]
    [Brand]
    [Unity]
    [Customer]
    [CustomerTrans]
    [Vendor]
    [Employee]
    [JobType]
    [Division]
    [Currenci]
    [Warehouse]
End Enum

Private Enum MenuWarehouseList
    [STOCKINIT]
    [STOCK]
    [ITEMIN]
    [ITEMOUT]
    [MUTITEM]
End Enum

Private Enum MenuBuyList
    [POBUY]
    [DOBUY]
    [SJBUY]
    [FKTBUY]
    [RTRBUY]
End Enum

Private Enum MenuSellList
    [SALESSUM]
    [POSELL]
    [SOSELL]
    [SJSELL]
    [FKTSELL]
    [RTRSELL]
End Enum

Private Sub Form_Load()
    SetInitialization
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmMIS = Nothing
End Sub

Private Sub Form_LostFocus()
    Unload Me
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim intCounter As Integer
    
    For intCounter = 0 To Me.lblMenuMaster.Count - 1
        Me.lblMenuMaster(intCounter).BackColor = &HC0E0FF
    Next intCounter
    
    For intCounter = 0 To Me.lblMenuWarehouse.Count - 1
        Me.lblMenuWarehouse(intCounter).BackColor = &HC0FFC0
    Next intCounter
    
    For intCounter = 0 To Me.lblMenuBuy.Count - 1
        Me.lblMenuBuy(intCounter).BackColor = &HC0C0FF
    Next intCounter
    
    For intCounter = 0 To Me.lblMenuSell.Count - 1
        Me.lblMenuSell(intCounter).BackColor = &HC0FFFF
    Next intCounter
End Sub

Private Sub lblMenuMaster_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = MouseButtonConstants.vbLeftButton Then
        Unload Me
        
        Select Case Index
            Case Item:
                mdlProcedures.ShowForm frmMISTMITEM, mdiMain.CloseAll, , frmMenu.Name
            Case ItemPrice:
                mdlProcedures.ShowForm frmMISTMITEMPRICE, mdiMain.CloseAll, , frmMenu.Name
            Case Group:
                mdlProcedures.ShowForm frmMISTMGROUP, mdiMain.CloseAll, , frmMenu.Name
            Case Category:
                mdlProcedures.ShowForm frmMISTMCATEGORY, mdiMain.CloseAll, , frmMenu.Name
            Case Brand:
                mdlProcedures.ShowForm frmMISTMBRAND, mdiMain.CloseAll, , frmMenu.Name
            Case Unity:
                mdlProcedures.ShowForm frmMISTMUNITY, mdiMain.CloseAll, , frmMenu.Name
            Case Customer:
                mdlProcedures.ShowForm frmMISTMCUSTOMER, mdiMain.CloseAll, , frmMenu.Name
            Case CustomerTrans:
                mdlProcedures.ShowForm frmMISTMCUSTOMERTRANS, mdiMain.CloseAll, , frmMenu.Name
            Case Vendor:
                mdlProcedures.ShowForm frmMISTMVENDOR, mdiMain.CloseAll, , frmMenu.Name
            Case Employee:
                mdlProcedures.ShowForm frmMISTMEMPLOYEE, mdiMain.CloseAll, , frmMenu.Name
            Case JobType:
                mdlProcedures.ShowForm frmMISTMJOBTYPE, mdiMain.CloseAll, , frmMenu.Name
            Case Division:
                mdlProcedures.ShowForm frmMISTMDIVISION, mdiMain.CloseAll, , frmMenu.Name
            Case Currenci:
                mdlProcedures.ShowForm frmMISTMCURRENCY, mdiMain.CloseAll, , frmMenu.Name
            Case Warehouse:
                mdlProcedures.ShowForm frmMISTMWAREHOUSE, mdiMain.CloseAll, , frmMenu.Name
        End Select
    End If
End Sub

Private Sub lblMenuMaster_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.lblMenuMaster(Index).BackColor = &HFFC0FF
End Sub

Private Sub lblMenuWarehouse_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = MouseButtonConstants.vbLeftButton Then
        Unload Me
        
        Select Case Index
            Case STOCKINIT:
                mdlProcedures.ShowForm frmMISTMSTOCKINIT, mdiMain.CloseAll, , frmMenu.Name
            Case STOCK:
                mdlProcedures.ShowForm frmMISTHSTOCK, mdiMain.CloseAll, , frmMenu.Name
            Case ITEMIN:
                mdlProcedures.ShowForm frmMISTHITEMIN, mdiMain.CloseAll, , frmMenu.Name
            Case ITEMOUT:
                mdlProcedures.ShowForm frmMISTHITEMOUT, mdiMain.CloseAll, , frmMenu.Name
            Case MUTITEM:
                mdlProcedures.ShowForm frmMISTHMUTITEM, mdiMain.CloseAll, , frmMenu.Name
        End Select
    End If
End Sub

Private Sub lblMenuWarehouse_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.lblMenuWarehouse(Index).BackColor = &HFF8080
End Sub

Private Sub lblMenuSell_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = MouseButtonConstants.vbLeftButton Then
        Unload Me
        
        Select Case Index
            Case SALESSUM:
                mdlProcedures.ShowForm frmMISTHSALESSUM, mdiMain.CloseAll, , frmMenu.Name
            Case POSELL:
                mdlProcedures.ShowForm frmMISTHPOSELL, mdiMain.CloseAll, , frmMenu.Name
            Case SOSELL:
                mdlProcedures.ShowForm frmMISTHSOSELL, mdiMain.CloseAll, , frmMenu.Name
            Case SJSELL:
                mdlProcedures.ShowForm frmMISTHSJSELL, mdiMain.CloseAll, , frmMenu.Name
            Case FKTSELL:
                mdlProcedures.ShowForm frmMISTHFKTSELL, mdiMain.CloseAll, , frmMenu.Name
            Case RTRSELL:
                mdlProcedures.ShowForm frmMISTHRTRSELL, mdiMain.CloseAll, , frmMenu.Name
        End Select
    End If
End Sub

Private Sub lblMenuBuy_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = MouseButtonConstants.vbLeftButton Then
        Unload Me
        
        Select Case Index
            Case POBUY:
                mdlProcedures.ShowForm frmMISTHPOBUY, mdiMain.CloseAll, , frmMenu.Name
            Case DOBUY:
                mdlProcedures.ShowForm frmMISTHDOBUY, mdiMain.CloseAll, , frmMenu.Name
            Case SJBUY:
                mdlProcedures.ShowForm frmMISTHSJBUY, mdiMain.CloseAll, , frmMenu.Name
            Case FKTBUY:
                mdlProcedures.ShowForm frmMISTHFKTBUY, mdiMain.CloseAll, , frmMenu.Name
            Case RTRBUY:
                mdlProcedures.ShowForm frmMISTHRTRBUY, mdiMain.CloseAll, , frmMenu.Name
        End Select
    End If
End Sub

Private Sub lblMenuBuy_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.lblMenuBuy(Index).BackColor = &HC0FFC0
End Sub

Private Sub lblMenuSell_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.lblMenuSell(Index).BackColor = &H80FF&
End Sub

Private Sub SetInitialization()
    Me.Move frmMenu.LeftMenu, frmMenu.TopMenu
    
    mdlAPI.SetLayeredWindow Me.hwnd, True
    
    mdlAPI.SetLayeredWindowAttributes Me.hwnd, 0, (255 * 60) / 100, &H2
    
    mdlAPI.SmoothForm Me, 30
    
    SetClient
End Sub

Private Sub SetClient()
    Me.lblMenu(0).Visible = mdiMain.mnuMaster.Visible
    Me.lblMenuMaster(Item).Visible = mdiMain.mnuMISItem.Visible
    Me.lblMenuMaster(ItemPrice).Visible = mdiMain.mnuMISItem.Visible
    Me.lblMenuMaster(Group).Visible = mdiMain.mnuMISItem.Visible
    Me.lblMenuMaster(Category).Visible = mdiMain.mnuMISItem.Visible
    Me.lblMenuMaster(Brand).Visible = mdiMain.mnuMISItem.Visible
    Me.lblMenuMaster(Unity).Visible = mdiMain.mnuMISItem.Visible
        
    If mdiMain.mnuMISItem.Visible Then
        Me.lblMenuMaster(Item).Visible = mdiMain.mnuMISTMITEM.Visible
        Me.lblMenuMaster(ItemPrice).Visible = mdiMain.mnuMISTMITEMPRICE.Visible
        Me.lblMenuMaster(Group).Visible = mdiMain.mnuMISTMGROUP.Visible
        Me.lblMenuMaster(Category).Visible = mdiMain.mnuMISTMCATEGORY.Visible
        Me.lblMenuMaster(Brand).Visible = mdiMain.mnuMISTMBRAND.Visible
        Me.lblMenuMaster(Unity).Visible = mdiMain.mnuMISTMUNITY.Visible
    End If
    
    Me.lblMenuMaster(Customer).Visible = mdiMain.mnuMISExternal.Visible
    Me.lblMenuMaster(Vendor).Visible = mdiMain.mnuMISExternal.Visible
    
    If mdiMain.mnuMISExternal.Visible Then
        Me.lblMenuMaster(Customer).Visible = mdiMain.mnuMISTMCUSTOMER.Visible
        Me.lblMenuMaster(Vendor).Visible = mdiMain.mnuMISTMVENDOR.Visible
    End If
    
    Me.lblMenuMaster(Employee).Visible = mdiMain.mnuMISInternal.Visible
    Me.lblMenuMaster(JobType).Visible = mdiMain.mnuMISInternal.Visible
    Me.lblMenuMaster(Division).Visible = mdiMain.mnuMISInternal.Visible
    Me.lblMenuMaster(Currenci).Visible = mdiMain.mnuMISInternal.Visible
    Me.lblMenuMaster(Warehouse).Visible = mdiMain.mnuMISInternal.Visible
    
    If mdiMain.mnuMISInternal.Visible Then
        Me.lblMenuMaster(Employee).Visible = mdiMain.mnuMISTMEMPLOYEE.Visible
        Me.lblMenuMaster(JobType).Visible = mdiMain.mnuMISTMJOBTYPE.Visible
        Me.lblMenuMaster(Division).Visible = mdiMain.mnuMISTMDIVISION.Visible
        Me.lblMenuMaster(Currenci).Visible = mdiMain.mnuMISTMCURRENCY.Visible
        Me.lblMenuMaster(Warehouse).Visible = mdiMain.mnuMISTMWAREHOUSE.Visible
    End If
    
    Me.lblMenu(1).Visible = mdiMain.mnuMISWarehouse.Visible
    Me.lblMenuWarehouse(STOCKINIT).Visible = mdiMain.mnuMISWarehouse.Visible
    Me.lblMenuWarehouse(STOCK).Visible = mdiMain.mnuMISWarehouse.Visible
    Me.lblMenuWarehouse(ITEMIN).Visible = mdiMain.mnuMISWarehouse.Visible
    Me.lblMenuWarehouse(ITEMOUT).Visible = mdiMain.mnuMISWarehouse.Visible
    Me.lblMenuWarehouse(MUTITEM).Visible = mdiMain.mnuMISWarehouse.Visible
    
    If mdiMain.mnuMISWarehouse.Visible Then
        Me.lblMenuWarehouse(STOCKINIT).Visible = mdiMain.mnuMISTMSTOCKINIT.Visible
        Me.lblMenuWarehouse(STOCK).Visible = mdiMain.mnuMISTHSTOCK.Visible
        Me.lblMenuWarehouse(ITEMIN).Visible = mdiMain.mnuMISTHITEMIN.Visible
        Me.lblMenuWarehouse(ITEMOUT).Visible = mdiMain.mnuMISTHITEMOUT.Visible
        Me.lblMenuWarehouse(MUTITEM).Visible = mdiMain.mnuMISTHMUTITEM.Visible
    End If
    
    Me.lblMenu(2).Visible = mdiMain.mnuMISBuy.Visible
    Me.lblMenuBuy(POBUY).Visible = mdiMain.mnuMISBuy.Visible
    Me.lblMenuBuy(DOBUY).Visible = mdiMain.mnuMISBuy.Visible
    Me.lblMenuBuy(SJBUY).Visible = mdiMain.mnuMISBuy.Visible
    Me.lblMenuBuy(FKTBUY).Visible = mdiMain.mnuMISBuy.Visible
    Me.lblMenuBuy(RTRBUY).Visible = mdiMain.mnuMISBuy.Visible
    
    If mdiMain.mnuMISBuy.Visible Then
        Me.lblMenuBuy(POBUY).Visible = mdiMain.mnuMISTHPOBUY.Visible
        Me.lblMenuBuy(DOBUY).Visible = mdiMain.mnuMISTHDOBUY.Visible
        Me.lblMenuBuy(SJBUY).Visible = mdiMain.mnuMISTHSJBUY.Visible
        Me.lblMenuBuy(FKTBUY).Visible = mdiMain.mnuMISTHFKTBUY.Visible
        Me.lblMenuBuy(RTRBUY).Visible = mdiMain.mnuMISTHRTRBUY.Visible
    End If
    
    Me.lblMenu(3).Visible = mdiMain.mnuMISSell.Visible
    Me.lblMenuSell(SALESSUM).Visible = mdiMain.mnuMISSell.Visible
    Me.lblMenuSell(POSELL).Visible = mdiMain.mnuMISSell.Visible
    Me.lblMenuSell(SOSELL).Visible = mdiMain.mnuMISSell.Visible
    Me.lblMenuSell(SJSELL).Visible = mdiMain.mnuMISSell.Visible
    Me.lblMenuSell(FKTSELL).Visible = mdiMain.mnuMISSell.Visible
    Me.lblMenuSell(RTRSELL).Visible = mdiMain.mnuMISSell.Visible
    
    If mdiMain.mnuMISSell.Visible Then
        Me.lblMenuSell(SALESSUM).Visible = mdiMain.mnuMISTHSALESSUM.Visible
        Me.lblMenuSell(POSELL).Visible = mdiMain.mnuMISTHPOSELL.Visible
        Me.lblMenuSell(SOSELL).Visible = mdiMain.mnuMISTHSOSELL.Visible
        Me.lblMenuSell(SJSELL).Visible = mdiMain.mnuMISTHSJSELL.Visible
        Me.lblMenuSell(FKTSELL).Visible = mdiMain.mnuMISTHFKTSELL.Visible
        Me.lblMenuSell(RTRSELL).Visible = mdiMain.mnuMISTHRTRSELL.Visible
    End If
End Sub
