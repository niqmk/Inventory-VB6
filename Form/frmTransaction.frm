VERSION 5.00
Begin VB.Form frmTransaction 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   0  'None
   ClientHeight    =   4125
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8085
   Icon            =   "frmTransaction.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4125
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
      TabIndex        =   12
      Top             =   600
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
      TabIndex        =   17
      Top             =   3600
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
      TabIndex        =   11
      Top             =   3000
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
      Index           =   3
      Left            =   120
      TabIndex        =   6
      Top             =   2400
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
      Index           =   2
      Left            =   120
      TabIndex        =   5
      Top             =   1800
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
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   2295
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
      TabIndex        =   3
      Top             =   600
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
      Index           =   2
      Left            =   5640
      TabIndex        =   2
      Top             =   120
      Width           =   2355
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
      TabIndex        =   10
      Top             =   2400
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
      TabIndex        =   9
      Top             =   1800
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
      TabIndex        =   16
      Top             =   3000
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
      TabIndex        =   15
      Top             =   2400
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
      TabIndex        =   14
      Top             =   1800
      Width           =   2295
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
      TabIndex        =   13
      Top             =   1200
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
      Index           =   1
      Left            =   2880
      TabIndex        =   1
      Top             =   120
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
      TabIndex        =   8
      Top             =   1200
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
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2355
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
      TabIndex        =   7
      Top             =   600
      Width           =   2295
   End
End
Attribute VB_Name = "frmTransaction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum MenuWarehouseList
    [STOCKINIT]
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
    Set frmTransaction = Nothing
End Sub

Private Sub Form_LostFocus()
    Unload Me
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim intCounter As Integer
    
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

Private Sub lblMenuWarehouse_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = MouseButtonConstants.vbLeftButton Then
        Unload Me
        
        Select Case Index
            Case STOCKINIT:
                mdlProcedures.ShowForm frmTMSTOCKINIT, mdiMain.CloseAll, , frmMenu.Name
            Case ITEMIN:
                mdlProcedures.ShowForm frmTHITEMIN, mdiMain.CloseAll, , frmMenu.Name
            Case ITEMOUT:
                mdlProcedures.ShowForm frmTHITEMOUT, mdiMain.CloseAll, , frmMenu.Name
            Case MUTITEM:
                mdlProcedures.ShowForm frmTHMUTITEM, mdiMain.CloseAll, , frmMenu.Name
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
                mdlProcedures.ShowForm frmTHSALESSUM, mdiMain.CloseAll, , frmMenu.Name
            Case POSELL:
                mdlProcedures.ShowForm frmTHPOSELL, mdiMain.CloseAll, , frmMenu.Name
            Case SOSELL:
                mdlProcedures.ShowForm frmTHSOSELL, mdiMain.CloseAll, , frmMenu.Name
            Case SJSELL:
                mdlProcedures.ShowForm frmTHSJSELL, mdiMain.CloseAll, , frmMenu.Name
            Case FKTSELL:
                mdlProcedures.ShowForm frmTHFKTSELL, mdiMain.CloseAll, , frmMenu.Name
        End Select
    End If
End Sub

Private Sub lblMenuBuy_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = MouseButtonConstants.vbLeftButton Then
        Unload Me
        
        Select Case Index
            Case POBUY:
                mdlProcedures.ShowForm frmTHPOBUY, mdiMain.CloseAll, , frmMenu.Name
            Case DOBUY:
                mdlProcedures.ShowForm frmTHDOBUY, mdiMain.CloseAll, , frmMenu.Name
            Case SJBUY:
                mdlProcedures.ShowForm frmTHSJBUY, mdiMain.CloseAll, , frmMenu.Name
            Case FKTBUY:
                mdlProcedures.ShowForm frmTHFKTBUY, mdiMain.CloseAll, , frmMenu.Name
            Case RTRBUY:
                mdlProcedures.ShowForm frmTHRTRBUY, mdiMain.CloseAll, , frmMenu.Name
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
    Me.lblMenu(0).Visible = mdiMain.mnuWarehouse.Visible
    Me.lblMenuWarehouse(STOCKINIT).Visible = mdiMain.mnuWarehouse.Visible
    Me.lblMenuWarehouse(ITEMIN).Visible = mdiMain.mnuWarehouse.Visible
    Me.lblMenuWarehouse(ITEMOUT).Visible = mdiMain.mnuWarehouse.Visible
    Me.lblMenuWarehouse(MUTITEM).Visible = mdiMain.mnuWarehouse.Visible
    
    If mdiMain.mnuWarehouse.Visible Then
        Me.lblMenuWarehouse(STOCKINIT).Visible = mdiMain.mnuTMSTOCKINIT.Visible
        Me.lblMenuWarehouse(ITEMIN).Visible = mdiMain.mnuTHITEMIN.Visible
        Me.lblMenuWarehouse(ITEMOUT).Visible = mdiMain.mnuTHITEMOUT.Visible
        Me.lblMenuWarehouse(MUTITEM).Visible = mdiMain.mnuTHMUTITEM.Visible
    End If
    
    Me.lblMenu(1).Visible = mdiMain.mnuBuy.Visible
    Me.lblMenuBuy(POBUY).Visible = mdiMain.mnuBuy.Visible
    Me.lblMenuBuy(DOBUY).Visible = mdiMain.mnuBuy.Visible
    Me.lblMenuBuy(SJBUY).Visible = mdiMain.mnuBuy.Visible
    Me.lblMenuBuy(FKTBUY).Visible = mdiMain.mnuBuy.Visible
    Me.lblMenuBuy(RTRBUY).Visible = mdiMain.mnuBuy.Visible
    
    If mdiMain.mnuBuy.Visible Then
        Me.lblMenuBuy(POBUY).Visible = mdiMain.mnuTHPOBUY.Visible
        Me.lblMenuBuy(DOBUY).Visible = mdiMain.mnuTHDOBUY.Visible
        Me.lblMenuBuy(SJBUY).Visible = mdiMain.mnuTHSJBUY.Visible
        Me.lblMenuBuy(FKTBUY).Visible = mdiMain.mnuTHFKTBUY.Visible
        Me.lblMenuBuy(RTRBUY).Visible = mdiMain.mnuTHRTRBUY.Visible
    End If
    
    Me.lblMenu(2).Visible = mdiMain.mnuSell.Visible
    Me.lblMenuSell(SALESSUM).Visible = mdiMain.mnuSell.Visible
    Me.lblMenuSell(POSELL).Visible = mdiMain.mnuSell.Visible
    Me.lblMenuSell(SOSELL).Visible = mdiMain.mnuSell.Visible
    Me.lblMenuSell(SJSELL).Visible = mdiMain.mnuSell.Visible
    Me.lblMenuSell(FKTSELL).Visible = mdiMain.mnuSell.Visible
    Me.lblMenuSell(RTRSELL).Visible = mdiMain.mnuSell.Visible
    
    If mdiMain.mnuSell.Visible Then
        Me.lblMenuSell(SALESSUM).Visible = mdiMain.mnuTHPOSELL.Visible
        Me.lblMenuSell(POSELL).Visible = mdiMain.mnuTHPOSELL.Visible
        Me.lblMenuSell(SOSELL).Visible = mdiMain.mnuTHSOSELL.Visible
        Me.lblMenuSell(SJSELL).Visible = mdiMain.mnuTHSJSELL.Visible
        Me.lblMenuSell(FKTSELL).Visible = mdiMain.mnuTHFKTSELL.Visible
        Me.lblMenuSell(RTRSELL).Visible = mdiMain.mnuTHRTRSELL.Visible
    End If
End Sub
