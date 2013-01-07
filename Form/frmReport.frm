VERSION 5.00
Begin VB.Form frmReport 
   BackColor       =   &H00808000&
   BorderStyle     =   0  'None
   ClientHeight    =   4620
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8100
   Icon            =   "frmReport.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4620
   ScaleWidth      =   8100
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Label lblMenuSell 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
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
      TabIndex        =   14
      Top             =   600
      Width           =   2295
   End
   Begin VB.Label lblMenuCustomer 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Total Order Penjualan"
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
      TabIndex        =   20
      Top             =   4080
      Width           =   2295
   End
   Begin VB.Label lblMenu 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Customer"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   3
      Top             =   3600
      Width           =   2355
   End
   Begin VB.Label lblMenuWarehouse 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0FF&
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
      TabIndex        =   8
      Top             =   3000
      Width           =   2295
   End
   Begin VB.Label lblMenuBuy 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
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
      TabIndex        =   9
      Top             =   600
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
      ForeColor       =   &H0080C0FF&
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
      BackColor       =   &H00FFC0C0&
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
      TabIndex        =   10
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
      ForeColor       =   &H0080C0FF&
      Height          =   255
      Index           =   1
      Left            =   2880
      TabIndex        =   1
      Top             =   120
      Width           =   2355
   End
   Begin VB.Label lblMenuSell 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
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
      TabIndex        =   15
      Top             =   1200
      Width           =   2295
   End
   Begin VB.Label lblMenuSell 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
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
      TabIndex        =   16
      Top             =   1800
      Width           =   2295
   End
   Begin VB.Label lblMenuSell 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
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
      TabIndex        =   17
      Top             =   2400
      Width           =   2295
   End
   Begin VB.Label lblMenuSell 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
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
      TabIndex        =   18
      Top             =   3000
      Width           =   2295
   End
   Begin VB.Label lblMenuBuy 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
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
      TabIndex        =   11
      Top             =   1800
      Width           =   2295
   End
   Begin VB.Label lblMenuBuy 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
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
      TabIndex        =   12
      Top             =   2400
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
      ForeColor       =   &H0080C0FF&
      Height          =   255
      Index           =   2
      Left            =   5640
      TabIndex        =   2
      Top             =   120
      Width           =   2355
   End
   Begin VB.Label lblMenuWarehouse 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0FF&
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
      TabIndex        =   4
      Top             =   600
      Width           =   2295
   End
   Begin VB.Label lblMenuWarehouse 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0FF&
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
      TabIndex        =   5
      Top             =   1200
      Width           =   2295
   End
   Begin VB.Label lblMenuWarehouse 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0FF&
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
      TabIndex        =   6
      Top             =   1800
      Width           =   2295
   End
   Begin VB.Label lblMenuWarehouse 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0FF&
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
      TabIndex        =   7
      Top             =   2400
      Width           =   2295
   End
   Begin VB.Label lblMenuBuy 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
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
      TabIndex        =   13
      Top             =   3000
      Width           =   2295
   End
   Begin VB.Label lblMenuSell 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
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
      TabIndex        =   19
      Top             =   3600
      Width           =   2295
   End
End
Attribute VB_Name = "frmReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

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

Private Enum MenuCustomerList
    [SUMTHPOSELL]
End Enum

Private Sub Form_Load()
    SetInitialization
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmReport = Nothing
End Sub

Private Sub Form_LostFocus()
    Unload Me
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim intCounter As Integer
    
    For intCounter = 0 To Me.lblMenuWarehouse.Count - 1
        Me.lblMenuWarehouse(intCounter).BackColor = &HFFC0FF
    Next intCounter
    
    For intCounter = 0 To Me.lblMenuBuy.Count - 1
        Me.lblMenuBuy(intCounter).BackColor = &HFFC0C0
    Next intCounter
    
    For intCounter = 0 To Me.lblMenuSell.Count - 1
        Me.lblMenuSell(intCounter).BackColor = &HC0FFC0
    Next intCounter
    
    For intCounter = 0 To Me.lblMenuCustomer.Count - 1
        Me.lblMenuCustomer(intCounter).BackColor = &HC0FFFF
    Next intCounter
End Sub

Private Sub lblMenuWarehouse_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = MouseButtonConstants.vbLeftButton Then
        Unload Me
        
        Select Case Index
            Case STOCKINIT:
                mdlProcedures.ShowForm frmRPTTMSTOCKINIT, mdiMain.CloseAll, , frmMenu.Name
            Case STOCK:
                mdlProcedures.ShowForm frmRPTTHSTOCK, mdiMain.CloseAll, , frmMenu.Name
            Case ITEMIN:
                mdlProcedures.ShowForm frmRPTTHITEMIN, mdiMain.CloseAll, , frmMenu.Name
            Case ITEMOUT:
                mdlProcedures.ShowForm frmRPTTHITEMOUT, mdiMain.CloseAll, , frmMenu.Name
            Case MUTITEM:
                mdlProcedures.ShowForm frmRPTTHMUTITEM, mdiMain.CloseAll, , frmMenu.Name
        End Select
    End If
End Sub

Private Sub lblMenuSell_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = MouseButtonConstants.vbLeftButton Then
        Unload Me
        
        Select Case Index
            Case SALESSUM:
                mdlProcedures.ShowForm frmRPTTHSALESSUM, mdiMain.CloseAll, , frmMenu.Name
            Case POSELL:
                mdlProcedures.ShowForm frmRPTTHPOSELL, mdiMain.CloseAll, , frmMenu.Name
            Case SOSELL:
                mdlProcedures.ShowForm frmRPTTHSOSELL, mdiMain.CloseAll, , frmMenu.Name
            Case SJSELL:
                mdlProcedures.ShowForm frmRPTTHSJSELL, mdiMain.CloseAll, , frmMenu.Name
            Case FKTSELL:
                mdlProcedures.ShowForm frmRPTTHFKTSELL, mdiMain.CloseAll, , frmMenu.Name
            Case RTRSELL:
                mdlProcedures.ShowForm frmRPTTHRTRSELL, mdiMain.CloseAll, , frmMenu.Name
        End Select
    End If
End Sub

Private Sub lblMenuBuy_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = MouseButtonConstants.vbLeftButton Then
        Unload Me
        
        Select Case Index
            Case POBUY:
                mdlProcedures.ShowForm frmRPTTHPOBUY, mdiMain.CloseAll, , frmMenu.Name
            Case DOBUY:
                mdlProcedures.ShowForm frmRPTTHDOBUY, mdiMain.CloseAll, , frmMenu.Name
            Case SJBUY:
                mdlProcedures.ShowForm frmRPTTHSJBUY, mdiMain.CloseAll, , frmMenu.Name
            Case FKTBUY:
                mdlProcedures.ShowForm frmRPTTHFKTBUY, mdiMain.CloseAll, , frmMenu.Name
            Case RTRBUY:
                mdlProcedures.ShowForm frmRPTTHRTRBUY, mdiMain.CloseAll, , frmMenu.Name
        End Select
    End If
End Sub

Private Sub lblMenuCustomer_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = MouseButtonConstants.vbLeftButton Then
        Unload Me
        
        Select Case Index
            Case SUMTHPOSELL:
                mdlProcedures.ShowForm frmRPTSUMTHPOSELL, mdiMain.CloseAll, , frmMenu.Name
        End Select
    End If
End Sub

Private Sub lblMenuWarehouse_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.lblMenuWarehouse(Index).BackColor = &HC0C0C0
End Sub

Private Sub lblMenuBuy_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.lblMenuBuy(Index).BackColor = &HFFFFFF
End Sub

Private Sub lblMenuSell_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.lblMenuSell(Index).BackColor = &HFFFFC0
End Sub

Private Sub lblMenuCustomer_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.lblMenuCustomer(Index).BackColor = &HC0E0FF
End Sub

Private Sub SetInitialization()
    Me.Move frmMenu.LeftMenu, frmMenu.TopMenu
    
    mdlAPI.SetLayeredWindow Me.hwnd, True
    
    mdlAPI.SetLayeredWindowAttributes Me.hwnd, 0, (255 * 60) / 100, &H2
    
    mdlAPI.SmoothForm Me, 30
    
    SetClient
End Sub

Private Sub SetClient()
    Me.lblMenu(0).Visible = mdiMain.mnuWarehouseReport.Visible
    Me.lblMenuWarehouse(STOCKINIT).Visible = mdiMain.mnuWarehouseReport.Visible
    Me.lblMenuWarehouse(ITEMIN).Visible = mdiMain.mnuWarehouseReport.Visible
    Me.lblMenuWarehouse(ITEMOUT).Visible = mdiMain.mnuWarehouseReport.Visible
    Me.lblMenuWarehouse(MUTITEM).Visible = mdiMain.mnuWarehouseReport.Visible
    
    If mdiMain.mnuWarehouseReport.Visible Then
        Me.lblMenuWarehouse(STOCKINIT).Visible = mdiMain.mnuRPTTMSTOCKINIT.Visible
        Me.lblMenuWarehouse(STOCK).Visible = mdiMain.mnuRPTTHSTOCK.Visible
        Me.lblMenuWarehouse(ITEMIN).Visible = mdiMain.mnuRPTTHITEMIN.Visible
        Me.lblMenuWarehouse(ITEMOUT).Visible = mdiMain.mnuRPTTHITEMOUT.Visible
        Me.lblMenuWarehouse(MUTITEM).Visible = mdiMain.mnuRPTTHMUTITEM.Visible
    End If
    
    Me.lblMenu(1).Visible = mdiMain.mnuBuyReport.Visible
    Me.lblMenuBuy(POBUY).Visible = mdiMain.mnuBuyReport.Visible
    Me.lblMenuBuy(DOBUY).Visible = mdiMain.mnuBuyReport.Visible
    Me.lblMenuBuy(SJBUY).Visible = mdiMain.mnuBuyReport.Visible
    Me.lblMenuBuy(FKTBUY).Visible = mdiMain.mnuBuyReport.Visible
    Me.lblMenuBuy(RTRBUY).Visible = mdiMain.mnuBuyReport.Visible
    
    If mdiMain.mnuBuyReport.Visible Then
        Me.lblMenuBuy(POBUY).Visible = mdiMain.mnuRPTTHPOBUY.Visible
        Me.lblMenuBuy(DOBUY).Visible = mdiMain.mnuRPTTHDOBUY.Visible
        Me.lblMenuBuy(SJBUY).Visible = mdiMain.mnuRPTTHSJBUY.Visible
        Me.lblMenuBuy(FKTBUY).Visible = mdiMain.mnuRPTTHFKTBUY.Visible
        Me.lblMenuBuy(RTRBUY).Visible = mdiMain.mnuRPTTHRTRBUY.Visible
    End If
    
    Me.lblMenu(2).Visible = mdiMain.mnuSellReport.Visible
    Me.lblMenuSell(SALESSUM).Visible = mdiMain.mnuSellReport.Visible
    Me.lblMenuSell(POSELL).Visible = mdiMain.mnuSellReport.Visible
    Me.lblMenuSell(SOSELL).Visible = mdiMain.mnuSellReport.Visible
    Me.lblMenuSell(SJSELL).Visible = mdiMain.mnuSellReport.Visible
    Me.lblMenuSell(FKTSELL).Visible = mdiMain.mnuSellReport.Visible
    Me.lblMenuSell(RTRSELL).Visible = mdiMain.mnuSellReport.Visible
    
    If mdiMain.mnuSellReport.Visible Then
        Me.lblMenuSell(SALESSUM).Visible = mdiMain.mnuRPTTHPOSELL.Visible
        Me.lblMenuSell(POSELL).Visible = mdiMain.mnuRPTTHPOSELL.Visible
        Me.lblMenuSell(SOSELL).Visible = mdiMain.mnuRPTTHSOSELL.Visible
        Me.lblMenuSell(SJSELL).Visible = mdiMain.mnuRPTTHSJSELL.Visible
        Me.lblMenuSell(FKTSELL).Visible = mdiMain.mnuRPTTHFKTSELL.Visible
        Me.lblMenuSell(RTRSELL).Visible = mdiMain.mnuRPTTHRTRSELL.Visible
    End If
    
    Me.lblMenu(3).Visible = mdiMain.mnuCustomerReport.Visible
    Me.lblMenuCustomer(SUMTHPOSELL).Visible = mdiMain.mnuCustomerReport.Visible
    
    If mdiMain.mnuCustomerReport.Visible Then
        Me.lblMenuSell(SUMTHPOSELL).Visible = mdiMain.mnuRPTSUMTHPOSELL.Visible
    End If
End Sub
