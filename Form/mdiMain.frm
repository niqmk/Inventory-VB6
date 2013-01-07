VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.MDIForm mdiMain 
   BackColor       =   &H8000000C&
   ClientHeight    =   4695
   ClientLeft      =   165
   ClientTop       =   255
   ClientWidth     =   6930
   Icon            =   "mdiMain.frx":0000
   Begin MSWinsockLib.Winsock wskListen 
      Left            =   120
      Top             =   720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.PictureBox picCurrencyToolbar 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   570
      Left            =   0
      ScaleHeight     =   570
      ScaleWidth      =   6930
      TabIndex        =   0
      Top             =   3510
      Width           =   6930
      Begin VB.CommandButton cmdConvert 
         Caption         =   "Ubah"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3600
         TabIndex        =   4
         Top             =   120
         Width           =   975
      End
      Begin VB.ComboBox cmbCurrencyToId 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   120
         Width           =   975
      End
      Begin VB.ComboBox cmbCurrencyFromId 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   120
         Width           =   975
      End
      Begin VB.TextBox txtConvertValue 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2280
         MaxLength       =   10
         TabIndex        =   3
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.Timer tmrMain 
      Left            =   6480
      Top             =   4200
   End
   Begin MSComctlLib.Toolbar tlbMain 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   6930
      _ExtentX        =   12224
      _ExtentY        =   1111
      ButtonWidth     =   609
      ButtonHeight    =   953
      Appearance      =   1
      _Version        =   393216
      Begin MSComctlLib.ImageList imlMain 
         Left            =   6240
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   48
         ImageHeight     =   48
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   4
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mdiMain.frx":1242
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mdiMain.frx":2D94
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mdiMain.frx":48E6
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mdiMain.frx":8D38
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.StatusBar stsMain 
      Align           =   2  'Align Bottom
      Height          =   615
      Left            =   0
      TabIndex        =   6
      Top             =   4080
      Width           =   6930
      _ExtentX        =   12224
      _ExtentY        =   1085
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Georgia"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList imlStatus 
      Left            =   6360
      Top             =   2520
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":A88A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":CCDC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":E82E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":10380
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuMaster 
      Caption         =   "Master"
      Begin VB.Menu mnuItem 
         Caption         =   "Barang"
         Begin VB.Menu mnuTMITEM 
            Caption         =   "Barang"
         End
         Begin VB.Menu mnuTMPRICELIST 
            Caption         =   "Daftar Harga"
         End
         Begin VB.Menu mnuTMITEMPRICE 
            Caption         =   "Harga Jual"
         End
         Begin VB.Menu mnuTMGROUP 
            Caption         =   "Grup"
         End
         Begin VB.Menu mnuTMCATEGORY 
            Caption         =   "Jenis"
         End
         Begin VB.Menu mnuTMBRAND 
            Caption         =   "Merk"
         End
         Begin VB.Menu mnuTMUNITY 
            Caption         =   "Satuan"
         End
      End
      Begin VB.Menu mnuExternal 
         Caption         =   "Eksternal"
         Begin VB.Menu mnuTMCUSTOMER 
            Caption         =   "Customer"
         End
         Begin VB.Menu mnuTMVENDOR 
            Caption         =   "Pemasok"
         End
      End
      Begin VB.Menu mnuInternal 
         Caption         =   "Internal"
         Begin VB.Menu mnuTMEMPLOYEE 
            Caption         =   "Karyawan"
         End
         Begin VB.Menu mnuTMJOBTYPE 
            Caption         =   "Jabatan"
         End
         Begin VB.Menu mnuTMDIVISION 
            Caption         =   "Divisi"
         End
         Begin VB.Menu mnuTMCURRENCY 
            Caption         =   "Mata Uang"
         End
         Begin VB.Menu mnuTMWAREHOUSE 
            Caption         =   "Gudang"
         End
      End
   End
   Begin VB.Menu mnuMIS 
      Caption         =   "Manajemen Informasi"
      Begin VB.Menu mnuMISItem 
         Caption         =   "Barang"
         Begin VB.Menu mnuMISTMITEM 
            Caption         =   "Barang"
         End
         Begin VB.Menu mnuMISTMPRICELIST 
            Caption         =   "Daftar Harga"
         End
         Begin VB.Menu mnuMISTMITEMPRICE 
            Caption         =   "Harga Jual"
         End
         Begin VB.Menu mnuMISTMGROUP 
            Caption         =   "Grup"
         End
         Begin VB.Menu mnuMISTMCATEGORY 
            Caption         =   "Jenis"
         End
         Begin VB.Menu mnuMISTMBRAND 
            Caption         =   "Merk"
         End
         Begin VB.Menu mnuMISTMUNITY 
            Caption         =   "Satuan"
         End
      End
      Begin VB.Menu mnuMISExternal 
         Caption         =   "Eksternal"
         Begin VB.Menu mnuMISTMCUSTOMER 
            Caption         =   "Customer"
         End
         Begin VB.Menu mnuMISTMCUSTOMERTRANS 
            Caption         =   "Transaksi Customer"
         End
         Begin VB.Menu mnuMISTMVENDOR 
            Caption         =   "Pemasok"
         End
      End
      Begin VB.Menu mnuMISInternal 
         Caption         =   "Internal"
         Begin VB.Menu mnuMISTMEMPLOYEE 
            Caption         =   "Karyawan"
         End
         Begin VB.Menu mnuMISTMJOBTYPE 
            Caption         =   "Jabatan"
         End
         Begin VB.Menu mnuMISTMDIVISION 
            Caption         =   "Divisi"
         End
         Begin VB.Menu mnuMISTMCURRENCY 
            Caption         =   "Mata Uang"
         End
         Begin VB.Menu mnuMISTMWAREHOUSE 
            Caption         =   "Gudang"
         End
      End
      Begin VB.Menu mnuMISWarehouse 
         Caption         =   "Gudang"
         Begin VB.Menu mnuMISTMSTOCKINIT 
            Caption         =   "Stok Awal"
         End
         Begin VB.Menu mnuMISTHSTOCK 
            Caption         =   "Sisa Stok"
         End
         Begin VB.Menu mnuMISTHITEMIN 
            Caption         =   "Pemasukkan Barang"
         End
         Begin VB.Menu mnuMISTHITEMOUT 
            Caption         =   "Pengeluaran Barang"
         End
         Begin VB.Menu mnuMISTHMUTITEM 
            Caption         =   "Mutasi Barang"
         End
      End
      Begin VB.Menu mnuMISBuy 
         Caption         =   "Pembelian"
         Begin VB.Menu mnuMISTHPOBUY 
            Caption         =   "Purchase Order"
         End
         Begin VB.Menu mnuMISTHDOBUY 
            Caption         =   "Delivery Order"
         End
         Begin VB.Menu mnuMISTHSJBUY 
            Caption         =   "Surat Jalan"
         End
         Begin VB.Menu mnuMISTHFKTBUY 
            Caption         =   "Faktur"
         End
         Begin VB.Menu mnuMISTHRTRBUY 
            Caption         =   "Retur"
         End
      End
      Begin VB.Menu mnuMISSell 
         Caption         =   "Penjualan"
         Begin VB.Menu mnuMISTHSALESSUM 
            Caption         =   "Sales Summary"
         End
         Begin VB.Menu mnuMISTHPOSELL 
            Caption         =   "Purchase Order"
         End
         Begin VB.Menu mnuMISTHSOSELL 
            Caption         =   "Sales Order"
         End
         Begin VB.Menu mnuMISTHSJSELL 
            Caption         =   "Surat Jalan"
         End
         Begin VB.Menu mnuMISTHFKTSELL 
            Caption         =   "Faktur"
         End
         Begin VB.Menu mnuMISTHRTRSELL 
            Caption         =   "Retur"
         End
      End
   End
   Begin VB.Menu mnuTransaction 
      Caption         =   "Transaksi"
      Begin VB.Menu mnuWarehouse 
         Caption         =   "Gudang"
         Begin VB.Menu mnuTMSTOCKINIT 
            Caption         =   "Stok Awal"
         End
         Begin VB.Menu mnuTHITEMIN 
            Caption         =   "Pemasukkan Barang"
         End
         Begin VB.Menu mnuTHITEMOUT 
            Caption         =   "Pengeluaran Barang"
         End
         Begin VB.Menu mnuTHMUTITEM 
            Caption         =   "Mutasi Barang"
         End
      End
      Begin VB.Menu mnuBuy 
         Caption         =   "Pembelian"
         Begin VB.Menu mnuTHPOBUY 
            Caption         =   "Purchase Order"
         End
         Begin VB.Menu mnuTHDOBUY 
            Caption         =   "Delivery Order"
         End
         Begin VB.Menu mnuTHSJBUY 
            Caption         =   "Surat Jalan"
         End
         Begin VB.Menu mnuTHFKTBUY 
            Caption         =   "Faktur"
         End
         Begin VB.Menu mnuTHRTRBUY 
            Caption         =   "Retur"
         End
      End
      Begin VB.Menu mnuSell 
         Caption         =   "Penjualan"
         Begin VB.Menu mnuTHSALESSUM 
            Caption         =   "Sales Summary"
         End
         Begin VB.Menu mnuTHPOSELL 
            Caption         =   "Purchase Order"
         End
         Begin VB.Menu mnuTHSOSELL 
            Caption         =   "Sales Order"
         End
         Begin VB.Menu mnuTHSJSELL 
            Caption         =   "Surat Jalan"
         End
         Begin VB.Menu mnuTHFKTSELL 
            Caption         =   "Faktur"
         End
         Begin VB.Menu mnuTHRTRSELL 
            Caption         =   "Retur"
         End
      End
   End
   Begin VB.Menu mnuReport 
      Caption         =   "Laporan"
      Begin VB.Menu mnuWarehouseReport 
         Caption         =   "Gudang"
         Begin VB.Menu mnuRPTTMSTOCKINIT 
            Caption         =   "Stok Awal"
         End
         Begin VB.Menu mnuRPTTHSTOCK 
            Caption         =   "Sisa Stok"
         End
         Begin VB.Menu mnuRPTTHITEMIN 
            Caption         =   "Penerimaan Barang"
         End
         Begin VB.Menu mnuRPTTHITEMOUT 
            Caption         =   "Pengeluaran Barang"
         End
         Begin VB.Menu mnuRPTTHMUTITEM 
            Caption         =   "Mutasi Barang"
         End
      End
      Begin VB.Menu mnuBuyReport 
         Caption         =   "Pembelian"
         Begin VB.Menu mnuRPTTHPOBUY 
            Caption         =   "Purchase Order"
         End
         Begin VB.Menu mnuRPTTHDOBUY 
            Caption         =   "Delivery Order"
         End
         Begin VB.Menu mnuRPTTHSJBUY 
            Caption         =   "Surat Jalan"
         End
         Begin VB.Menu mnuRPTTHFKTBUY 
            Caption         =   "Faktur"
         End
         Begin VB.Menu mnuRPTTHRTRBUY 
            Caption         =   "Retur"
         End
      End
      Begin VB.Menu mnuSellReport 
         Caption         =   "Penjualan"
         Begin VB.Menu mnuRPTTHSALESSUM 
            Caption         =   "Sales Summary"
         End
         Begin VB.Menu mnuRPTTHPOSELL 
            Caption         =   "Purchase Order"
         End
         Begin VB.Menu mnuRPTTHSOSELL 
            Caption         =   "Sales Order"
         End
         Begin VB.Menu mnuRPTTHSJSELL 
            Caption         =   "Surat Jalan"
         End
         Begin VB.Menu mnuRPTTHFKTSELL 
            Caption         =   "Faktur"
         End
         Begin VB.Menu mnuRPTTHRTRSELL 
            Caption         =   "Retur"
         End
      End
      Begin VB.Menu mnuExternalReport 
         Caption         =   "Eksternal"
         Begin VB.Menu mnuCustomerReport 
            Caption         =   "Customer"
            Begin VB.Menu mnuRPTSUMTHPOSELL 
               Caption         =   "Total Order Penjualan"
            End
         End
      End
   End
   Begin VB.Menu mnuSettings 
      Caption         =   "Pengaturan"
      Begin VB.Menu mnuAddUser 
         Caption         =   "Tambah Pemakai"
      End
      Begin VB.Menu mnuSeparatorSettings1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTMREMINDERCUSTOMER 
         Caption         =   "Pengingat Customer"
      End
      Begin VB.Menu mnuSeparatorSettings2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPwdChange 
         Caption         =   "Ubah Password"
      End
      Begin VB.Menu mnuBackup 
         Caption         =   "Backup Database"
      End
      Begin VB.Menu mnuToolbar 
         Caption         =   "Toolbar"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuCurrencyToolbar 
         Caption         =   "Toolbar Mata Uang"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuMenu 
         Caption         =   "Menu"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuSeparatorSettings3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSyncFinance 
         Caption         =   "Synchronize Finance"
      End
      Begin VB.Menu mnuSeparatorSettings4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSyncAccounting 
         Caption         =   "Synchronize Accounting"
      End
   End
   Begin VB.Menu mnuWindows 
      Caption         =   "Windows"
      Begin VB.Menu mnuHelp 
         Caption         =   "Bantuan"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuCloseAll 
         Caption         =   "Tutup Semua Windows"
      End
      Begin VB.Menu mnuSeparatorWindows 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "Tentang Program"
      End
   End
   Begin VB.Menu mnuMouseRight 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnuWallpaper 
         Caption         =   "Ubah Wallpaper"
      End
      Begin VB.Menu mnuLogOut 
         Caption         =   "Log Out"
         Shortcut        =   {F2}
      End
   End
End
Attribute VB_Name = "mdiMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum StatusMode
    [UserMode] = 1
    [ReminderMode]
    [TimeMode]
End Enum

Private Enum MenuMode
    [MasterMode]
    [MISMode]
    [TransactionMode]
    [ReportMode]
End Enum

Private Enum ButtonMode
    [FaxButton] = 1
    [AnalysisButton]
    [PrinterButton]
    [ChatButton]
End Enum

Private blnParent As Boolean
Private blnExists As Boolean
Private blnFormExists As Boolean
Private blnReminder As Boolean
Private blnConvertCurrency As Boolean

Private Sub MDIForm_Load()
    SetInitialization
End Sub

Private Sub MDIForm_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = MouseButtonConstants.vbRightButton Then
        PopupMenu mnuMouseRight
    End If
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If mdlGlobal.blnFill Then
        Cancel = 1
    Else
        blnFormExists = False
        Me.tmrMain.Enabled = False
        
        mdlGlobal.UserAuthority.SetLogout mdlGlobal.UserAuthority.UserId, Me.wskListen.LocalIP
        
        CloseListener
        
        mdlProcedures.CloseAllForms Me
        
        mdlDatabase.CloseConnection mdlGlobal.conInventory
        mdlDatabase.CloseConnection mdlGlobal.conFinance
        mdlDatabase.CloseConnection mdlGlobal.conAccounting
        
        Set mdlGlobal.fso = Nothing
    End If
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    Set mdiMain = Nothing
End Sub

Private Sub wskListen_ConnectionRequest(ByVal requestID As Long)
    If Not Me.wskListen.State = sckClosed Then
        Me.wskListen.Close
    End If
    
    Me.wskListen.Accept requestID
End Sub

Private Sub wskListen_DataArrival(ByVal bytesTotal As Long)
    If Not mdlGlobal.blnChat Then
        mdlProcedures.ShowForm frmChat, False
    End If
    
    Dim strData As String
    
    Me.wskListen.GetData strData, vbString
    
    frmChat.ReceiveChatText mdlProcedures.SplitData(strData) & " (" & Me.wskListen.RemoteHostIP & ")", mdlProcedures.SplitData(strData, 1)
End Sub

Private Sub wskListen_Close()
    If Not Me.wskListen.State = sckClosed Then
        CloseListener
    End If
End Sub

Private Sub txtConvertValue_Change()
    Me.txtConvertValue.Text = mdlProcedures.FormatCurrency(Me.txtConvertValue.Text)
    
    Me.txtConvertValue.SelStart = Len(Me.txtConvertValue.Text)
End Sub

Private Sub txtConvertValue_GotFocus()
    mdlProcedures.GotFocus Me.txtConvertValue
End Sub

Private Sub cmbCurrencyFromId_Click()
    If Not mdlProcedures.IsValidComboData(Me.cmbCurrencyFromId) Then Exit Sub
    
    CheckConvertValue
End Sub

Private Sub cmbCurrencyToId_Click()
    If Not mdlProcedures.IsValidComboData(Me.cmbCurrencyToId) Then Exit Sub
    
    CheckConvertValue
End Sub

Private Sub cmdConvert_Click()
    If Not mdlProcedures.IsValidComboData(Me.cmbCurrencyFromId) Then Exit Sub
    If Not mdlProcedures.IsValidComboData(Me.cmbCurrencyToId) Then Exit Sub
    If Not blnConvertCurrency Then Exit Sub
    If Not mdlGlobal.UserAuthority.IsMenuAccess(Me.mnuTMCURRENCY.Name) Then Exit Sub
    
    Dim strCurrencyFromId As String
    Dim strCurrencyToId As String
    
    strCurrencyFromId = mdlProcedures.GetComboData(Me.cmbCurrencyFromId)
    strCurrencyToId = mdlProcedures.GetComboData(Me.cmbCurrencyToId)
    
    Dim rstTemp As ADODB.Recordset
    
    Set rstTemp = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "*", mdlTable.CreateTMCONVERTCURRENCY)
    
    mdlDatabase.SearchRecordset rstTemp, "ConvertId", strCurrencyFromId & strCurrencyToId & mdlProcedures.FormatDate(Now, "ddMMyyyy")
        
    With rstTemp
        If .EOF Then
            .AddNew
            
            !ConvertId = strCurrencyFromId & strCurrencyToId & mdlProcedures.FormatDate(Now, "ddMMyyyy")
            !ConvertDate = mdlProcedures.FormatDate(Now, mdlGlobal.strFormatDate)
            !CurrencyFromId = strCurrencyFromId
            !CurrencyToId = strCurrencyToId
            
            !CreateId = Trim(mdlGlobal.UserAuthority.UserId)
            !CreateDate = mdlProcedures.FormatDate(Now)
        End If
        
        !ConvertValue = mdlProcedures.GetCurrency(Me.txtConvertValue.Text)
        
        !UpdateId = Trim(mdlGlobal.UserAuthority.UserId)
        !UpdateDate = mdlProcedures.FormatDate(Now)
        
        .Update
    End With
    
    MsgBox "Dari Mata Uang : " & _
        mdlDatabase.GetFieldData(mdlGlobal.conInventory, "Name", mdlTable.CreateTMCURRENCY, "CurrencyId='" & strCurrencyFromId & "'") & vbCrLf & _
        "Ke Mata Uang : " & _
        mdlDatabase.GetFieldData(mdlGlobal.conInventory, "Name", mdlTable.CreateTMCURRENCY, "CurrencyId='" & strCurrencyToId & "'") & vbCrLf & _
        "Tanggal : " & mdlProcedures.FormatDate(Now, mdlGlobal.strFormatDate) & vbCrLf & _
        "Nilai Tukar : " & mdlProcedures.GetCurrency(Me.txtConvertValue.Text), vbOKOnly + vbInformation, Me.Caption

    mdlDatabase.CloseRecordset rstTemp
End Sub

Private Sub tlbMain_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case FaxButton:
            mdlProcedures.ShowForm frmFax, Me.CloseAll, , frmMenu.Name
        Case AnalysisButton:
            mdlProcedures.ShowForm frmAnalysis, Me.CloseAll, , frmMenu.Name
        Case PrinterButton:
            mdlProcedures.ShowForm frmPrinter, Me.CloseAll, , frmMenu.Name
        Case ChatButton:
            mdlProcedures.ShowForm frmChat, Me.CloseAll, , frmMenu.Name
    End Select
End Sub

Private Sub mnuTMITEM_Click()
    mdlProcedures.ShowForm frmTMITEM, Me.CloseAll, , frmMenu.Name
End Sub

Private Sub mnuTMPRICELIST_Click()
    mdlProcedures.ShowForm frmTMPRICELIST, Me.CloseAll, , frmMenu.Name
End Sub

Private Sub mnuTMITEMPRICE_Click()
    mdlProcedures.ShowForm frmTMITEMPRICE, Me.CloseAll, , frmMenu.Name
End Sub

Private Sub mnuTMGROUP_Click()
    mdlProcedures.ShowForm frmTMGROUP, Me.CloseAll, , frmMenu.Name
End Sub

Private Sub mnuTMCATEGORY_Click()
    mdlProcedures.ShowForm frmTMCATEGORY, Me.CloseAll, , frmMenu.Name
End Sub

Private Sub mnuTMBRAND_Click()
    mdlProcedures.ShowForm frmTMBRAND, Me.CloseAll, , frmMenu.Name
End Sub

Private Sub mnuTMUNITY_Click()
    mdlProcedures.ShowForm frmTMUNITY, Me.CloseAll, , frmMenu.Name
End Sub

Private Sub mnuTMCUSTOMER_Click()
    mdlProcedures.ShowForm frmTMCUSTOMER, Me.CloseAll, , frmMenu.Name
End Sub

Private Sub mnuTMVENDOR_Click()
    mdlProcedures.ShowForm frmTMVENDOR, Me.CloseAll, , frmMenu.Name
End Sub

Private Sub mnuTMEMPLOYEE_Click()
    mdlProcedures.ShowForm frmTMEMPLOYEE, Me.CloseAll, , frmMenu.Name
End Sub

Private Sub mnuTMJOBTYPE_Click()
    mdlProcedures.ShowForm frmTMJOBTYPE, Me.CloseAll, , frmMenu.Name
End Sub

Private Sub mnuTMDIVISION_Click()
    mdlProcedures.ShowForm frmTMDIVISION, Me.CloseAll, , frmMenu.Name
End Sub

Private Sub mnuTMCURRENCY_Click()
    mdlProcedures.ShowForm frmTMCURRENCY, Me.CloseAll, , frmMenu.Name
End Sub

Private Sub mnuTMWAREHOUSE_Click()
    mdlProcedures.ShowForm frmTMWAREHOUSE, Me.CloseAll, , frmMenu.Name
End Sub

Private Sub mnuTMSTOCKINIT_Click()
    mdlProcedures.ShowForm frmTMSTOCKINIT, Me.CloseAll, , frmMenu.Name
End Sub

Private Sub mnuTHITEMIN_Click()
    mdlProcedures.ShowForm frmTHITEMIN, Me.CloseAll, , frmMenu.Name
End Sub

Private Sub mnuTHITEMOUT_Click()
    mdlProcedures.ShowForm frmTHITEMOUT, Me.CloseAll, , frmMenu.Name
End Sub

Private Sub mnuTHMUTITEM_Click()
    mdlProcedures.ShowForm frmTHMUTITEM, Me.CloseAll, , frmMenu.Name
End Sub

Private Sub mnuMISTMITEM_Click()
    mdlProcedures.ShowForm frmMISTMITEM, Me.CloseAll, , frmMenu.Name
End Sub

Private Sub mnuMISTMITEMPRICE_Click()
    mdlProcedures.ShowForm frmMISTMITEMPRICE, Me.CloseAll, , frmMenu.Name
End Sub

Private Sub mnuMISTMGROUP_Click()
    mdlProcedures.ShowForm frmMISTMGROUP, Me.CloseAll, , frmMenu.Name
End Sub

Private Sub mnuMISTMCATEGORY_Click()
    mdlProcedures.ShowForm frmMISTMCATEGORY, Me.CloseAll, , frmMenu.Name
End Sub

Private Sub mnuMISTMBRAND_Click()
    mdlProcedures.ShowForm frmMISTMBRAND, Me.CloseAll, , frmMenu.Name
End Sub

Private Sub mnuMISTMUNITY_Click()
    mdlProcedures.ShowForm frmMISTMUNITY, Me.CloseAll, , frmMenu.Name
End Sub

Private Sub mnuMISTMCUSTOMER_Click()
    mdlProcedures.ShowForm frmMISTMCUSTOMER, Me.CloseAll, , frmMenu.Name
End Sub

Private Sub mnuMISTMCUSTOMERTRANS_Click()
    mdlProcedures.ShowForm frmMISTMCUSTOMERTRANS, Me.CloseAll, , frmMenu.Name
End Sub

Private Sub mnuMISTMVENDOR_Click()
    mdlProcedures.ShowForm frmMISTMVENDOR, Me.CloseAll, , frmMenu.Name
End Sub

Private Sub mnuMISTMEMPLOYEE_Click()
    mdlProcedures.ShowForm frmMISTMEMPLOYEE, Me.CloseAll, , frmMenu.Name
End Sub

Private Sub mnuMISTMJOBTYPE_Click()
    mdlProcedures.ShowForm frmMISTMJOBTYPE, Me.CloseAll, , frmMenu.Name
End Sub

Private Sub mnuMISTMDIVISION_Click()
    mdlProcedures.ShowForm frmMISTMDIVISION, Me.CloseAll, , frmMenu.Name
End Sub

Private Sub mnuMISTMCURRENCY_Click()
    mdlProcedures.ShowForm frmMISTMCURRENCY, Me.CloseAll, , frmMenu.Name
End Sub

Private Sub mnuMISTMWAREHOUSE_Click()
    mdlProcedures.ShowForm frmMISTMWAREHOUSE, Me.CloseAll, , frmMenu.Name
End Sub

Private Sub mnuMISTMSTOCKINIT_Click()
    mdlProcedures.ShowForm frmMISTMSTOCKINIT, Me.CloseAll, , frmMenu.Name
End Sub

Private Sub mnuMISTHSTOCK_Click()
    mdlProcedures.ShowForm frmMISTHSTOCK, Me.CloseAll, , frmMenu.Name
End Sub

Private Sub mnuMISTHITEMIN_Click()
    mdlProcedures.ShowForm frmMISTHITEMIN, Me.CloseAll, , frmMenu.Name
End Sub

Private Sub mnuMISTHITEMOUT_Click()
    mdlProcedures.ShowForm frmMISTHITEMOUT, Me.CloseAll, , frmMenu.Name
End Sub

Private Sub mnuMISTHMUTITEM_Click()
    mdlProcedures.ShowForm frmMISTHMUTITEM, Me.CloseAll, , frmMenu.Name
End Sub

Private Sub mnuMISTHPOBUY_Click()
    mdlProcedures.ShowForm frmMISTHPOBUY, Me.CloseAll, , frmMenu.Name
End Sub

Private Sub mnuMISTHDOBUY_Click()
    mdlProcedures.ShowForm frmMISTHDOBUY, Me.CloseAll, , frmMenu.Name
End Sub

Private Sub mnuMISTHSJBUY_Click()
    mdlProcedures.ShowForm frmMISTHSJBUY, Me.CloseAll, , frmMenu.Name
End Sub

Private Sub mnuMISTHFKTBUY_Click()
    mdlProcedures.ShowForm frmMISTHFKTBUY, Me.CloseAll, , frmMenu.Name
End Sub

Private Sub mnuMISTHRTRBUY_Click()
    mdlProcedures.ShowForm frmMISTHRTRBUY, Me.CloseAll, , frmMenu.Name
End Sub

Private Sub mnuMISTHSALESSUM_Click()
    mdlProcedures.ShowForm frmMISTHSALESSUM, Me.CloseAll, , frmMenu.Name
End Sub

Private Sub mnuMISTHPOSELL_Click()
    mdlProcedures.ShowForm frmMISTHPOSELL, Me.CloseAll, , frmMenu.Name
End Sub

Private Sub mnuMISTHSOSELL_Click()
    mdlProcedures.ShowForm frmMISTHSOSELL, Me.CloseAll, , frmMenu.Name
End Sub

Private Sub mnuMISTHSJSELL_Click()
    mdlProcedures.ShowForm frmMISTHSJSELL, Me.CloseAll, , frmMenu.Name
End Sub

Private Sub mnuMISTHFKTSELL_Click()
    mdlProcedures.ShowForm frmMISTHFKTSELL, Me.CloseAll, , frmMenu.Name
End Sub

Private Sub mnuMISTHRTRSELL_Click()
    mdlProcedures.ShowForm frmMISTHRTRSELL, Me.CloseAll, , frmMenu.Name
End Sub

Private Sub mnuTHPOBUY_Click()
    mdlProcedures.ShowForm frmTHPOBUY, Me.CloseAll, , frmMenu.Name
End Sub

Private Sub mnuTHDOBUY_Click()
    mdlProcedures.ShowForm frmTHDOBUY, Me.CloseAll, , frmMenu.Name
End Sub

Private Sub mnuTHSJBUY_Click()
    mdlProcedures.ShowForm frmTHSJBUY, Me.CloseAll, , frmMenu.Name
End Sub

Private Sub mnuTHFKTBUY_Click()
    mdlProcedures.ShowForm frmTHFKTBUY, Me.CloseAll, , frmMenu.Name
End Sub

Private Sub mnuTHRTRBUY_Click()
    mdlProcedures.ShowForm frmTHRTRBUY, Me.CloseAll, , frmMenu.Name
End Sub

Private Sub mnuTHSALESSUM_Click()
    mdlProcedures.ShowForm frmTHSALESSUM, Me.CloseAll, , frmMenu.Name
End Sub

Private Sub mnuTHPOSELL_Click()
    mdlProcedures.ShowForm frmTHPOSELL, Me.CloseAll, , frmMenu.Name
End Sub

Private Sub mnuTHSOSELL_Click()
    mdlProcedures.ShowForm frmTHSOSELL, Me.CloseAll, , frmMenu.Name
End Sub

Private Sub mnuTHSJSELL_Click()
    mdlProcedures.ShowForm frmTHSJSELL, Me.CloseAll, , frmMenu.Name
End Sub

Private Sub mnuTHFKTSELL_Click()
    mdlProcedures.ShowForm frmTHFKTSELL, Me.CloseAll, , frmMenu.Name
End Sub

Private Sub mnuTHRTRSELL_Click()
    mdlProcedures.ShowForm frmTHRTRSELL, Me.CloseAll, , frmMenu.Name
End Sub

Private Sub mnuRPTTMSTOCKINIT_Click()
    mdlProcedures.ShowForm frmRPTTMSTOCKINIT, Me.CloseAll, , frmMenu.Name
End Sub

Private Sub mnuRPTTHSTOCK_Click()
    mdlProcedures.ShowForm frmRPTTHSTOCK, Me.CloseAll, , frmMenu.Name
End Sub

Private Sub mnuRPTTHITEMIN_Click()
    mdlProcedures.ShowForm frmRPTTHITEMIN, Me.CloseAll, , frmMenu.Name
End Sub

Private Sub mnuRPTTHITEMOUT_Click()
    mdlProcedures.ShowForm frmRPTTHITEMOUT, Me.CloseAll, , frmMenu.Name
End Sub

Private Sub mnuRPTTHMUTITEM_Click()
    mdlProcedures.ShowForm frmRPTTHMUTITEM, Me.CloseAll, , frmMenu.Name
End Sub

Private Sub mnuRPTTHPOBUY_Click()
    mdlProcedures.ShowForm frmRPTTHPOBUY, Me.CloseAll, , frmMenu.Name
End Sub

Private Sub mnuRPTTHDOBUY_Click()
    mdlProcedures.ShowForm frmRPTTHDOBUY, Me.CloseAll, , frmMenu.Name
End Sub

Private Sub mnuRPTTHSJBUY_Click()
    mdlProcedures.ShowForm frmRPTTHSJBUY, Me.CloseAll, , frmMenu.Name
End Sub

Private Sub mnuRPTTHFKTBUY_Click()
    mdlProcedures.ShowForm frmRPTTHFKTBUY, Me.CloseAll, , frmMenu.Name
End Sub

Private Sub mnuRPTTHRTRBUY_Click()
    mdlProcedures.ShowForm frmRPTTHRTRBUY, Me.CloseAll, , frmMenu.Name
End Sub

Private Sub mnuRPTTHSALESSUM_Click()
    mdlProcedures.ShowForm frmRPTTHSALESSUM, Me.CloseAll, , frmMenu.Name
End Sub

Private Sub mnuRPTTHPOSELL_Click()
    mdlProcedures.ShowForm frmRPTTHPOSELL, Me.CloseAll, , frmMenu.Name
End Sub

Private Sub mnuRPTTHSOSELL_Click()
    mdlProcedures.ShowForm frmRPTTHSOSELL, Me.CloseAll, , frmMenu.Name
End Sub

Private Sub mnuRPTTHSJSELL_Click()
    mdlProcedures.ShowForm frmRPTTHSJSELL, Me.CloseAll, , frmMenu.Name
End Sub

Private Sub mnuRPTTHFKTSELL_Click()
    mdlProcedures.ShowForm frmRPTTHFKTSELL, Me.CloseAll, , frmMenu.Name
End Sub

Private Sub mnuRPTTHRTRSELL_Click()
    mdlProcedures.ShowForm frmRPTTHRTRSELL, Me.CloseAll, , frmMenu.Name
End Sub

Private Sub mnuRPTSUMTHPOSELL_Click()
    mdlProcedures.ShowForm frmRPTSUMTHPOSELL, Me.CloseAll, , frmMenu.Name
End Sub

Private Sub mnuTMREMINDERCUSTOMER_Click()
    mdlProcedures.ShowForm frmTMREMINDERCUSTOMER, Me.CloseAll, , frmMenu.Name
End Sub

Private Sub mnuAddUser_Click()
    mdlProcedures.ShowForm frmAddUser, Me.CloseAll, , frmMenu.Name
End Sub

Private Sub mnuPwdChange_Click()
    mdlProcedures.ShowForm frmPwdChange, Me.CloseAll, , frmMenu.Name
End Sub

Private Sub mnuBackup_Click()
    mdlProcedures.ShowForm frmBackup, Me.CloseAll, , frmMenu.Name
End Sub

Private Sub mnuToolbar_Click()
    Me.mnuToolbar.Checked = Not Me.mnuToolbar.Checked
    
    Me.tlbMain.Visible = Me.mnuToolbar.Checked
    Me.stsMain.Visible = Me.mnuToolbar.Checked
End Sub

Private Sub mnuCurrencyToolbar_Click()
    Me.mnuCurrencyToolbar.Checked = Not Me.mnuCurrencyToolbar.Checked
    
    Me.picCurrencyToolbar.Visible = Me.mnuCurrencyToolbar.Checked
End Sub

Private Sub mnuMenu_Click()
    Me.mnuMenu.Checked = Not mnuMenu.Checked
    
    If Me.mnuMenu.Checked Then
        frmMenu.Show
    Else
        frmMenu.Hide
    End If
End Sub

Private Sub mnuSyncFinance_Click()
    mdlProcedures.ShowForm frmSyncFinance, Me.CloseAll, , frmMenu.Name
End Sub

Private Sub mnuSyncAccounting_Click()
    mdlProcedures.ShowForm frmSyncAccounting, Me.CloseAll, , frmMenu.Name
End Sub

Private Sub mnuHelp_Click()
    If mdlGlobal.fso.FileExists(mdlGlobal.strPath & "help.chm") Then
        mdlAPI.ShellExecute Me.hwnd, "", mdlGlobal.strPath & "help.chm", "", "", 1
    End If
End Sub

Private Sub mnuCloseAll_Click()
    mnuCloseAll.Checked = Not mnuCloseAll.Checked
    
    If mnuCloseAll.Checked Then mdlProcedures.CloseAllForms Me, frmMenu.Name
End Sub

Private Sub mnuAbout_Click()
    mdlProcedures.ShowForm frmAbout, Me.CloseAll, , frmMenu.Name
End Sub

Private Sub mnuWallpaper_Click()
    mdlProcedures.ShowForm frmWallpaper, Me.CloseAll, , frmMenu.Name
End Sub

Private Sub mnuLogOut_Click()
    If MsgBox("Apakah Anda Yakin? (Log Off)", vbQuestion + vbYesNo, Me.Caption) = vbNo Then Exit Sub
    
    mdlGlobal.UserAuthority.SetLogout mdlGlobal.UserAuthority.UserId, Me.wskListen.LocalIP
    
    Me.stsMain.Panels(UserMode).Bevel = sbrRaised
    
    SetLog
End Sub

Private Sub stsMain_PanelClick(ByVal Panel As MSComctlLib.Panel)
    Select Case Panel.Index
        Case UserMode:
            If MsgBox("Apakah Anda Yakin? (Log Off)", vbQuestion + vbYesNo, Me.Caption) = vbNo Then Exit Sub
            
            Me.stsMain.Panels(UserMode).Bevel = sbrRaised
            
            SetLog
        Case ReminderMode:
            If Not (ReminderCustomer + ReminderPOSELL) > 0 Then Exit Sub
            
            mdlProcedures.ShowForm frmReminderList, False, , frmMenu.Name
    End Select
End Sub

Private Sub SetInitialization()
    mdlProcedures.CornerWindows Me
    
    Me.Caption = mdlGlobal.strCompanyText
    
    Me.WindowState = FormWindowStateConstants.vbMaximized
    
    With Me.tlbMain
        .AllowCustomize = False
        .ImageList = Me.imlMain
        
        .Buttons.Add FaxButton, , "Fax", , FaxButton
        .Buttons.Add AnalysisButton, , "Analisis", , AnalysisButton
        .Buttons.Add PrinterButton, , "Set Printer", , PrinterButton
        .Buttons.Add ChatButton, , "Chat", , ChatButton
    End With
    
    With Me.stsMain
        .Panels.Add UserMode, , , sbrText, imlStatus.ListImages(1).Picture
        .Panels.Add ReminderMode, , , sbrText, imlStatus.ListImages(2).Picture
        .Panels.Add TimeMode, , , sbrText, imlStatus.ListImages(4).Picture
        
        .Panels(TimeMode).Text = mdlProcedures.FormatDate(Now, "dd-MMMM-yyyy (hh:mm:ss)")
    End With
    
    With Me.tmrMain
        .Interval = 1000
        .Enabled = True
    End With
    
    WallpaperInitialize
    SetListener
    
    frmMenu.Show , Me
    
    blnFormExists = True
    blnReminder = False
    blnConvertCurrency = False
    
    SetStatus
    SetClient
    
    Me.CheckReminder
    Me.CheckConvertCurrency
End Sub

Private Sub SetListener()
    Me.wskListen.LocalPort = mdlGlobal.intPort
    Me.wskListen.Listen
End Sub

Private Sub CloseListener()
    Me.wskListen.Close
    
    If blnFormExists Then
        SetListener
    End If
End Sub

Private Sub WallpaperInitialize()
    Dim lngRegistry As Long
    Dim lngRegKey As Long
    Dim lngType As Long
    Dim lngSize As Long
    
    On Local Error GoTo ErrHandler

    lngRegistry = _
        mdlRegistry.OpenRegistry(mdlRegistry.HKEY_CURRENT_USER, mdlRegistry.KEYS_SYS_INFO, lngRegKey)

    If lngRegistry = 0 Then
        lngRegistry = _
            mdlRegistry.ReadValueRegistry(lngRegKey, mdlGlobal.WALLPAPER_REGISTRY, lngType, mdlGlobal.strWallpaperImageText, lngSize)
            
        mdlGlobal.strWallpaperImageText = Trim(CStr(mdlGlobal.strWallpaperImageText))
        
        If mdlGlobal.fso.FileExists(mdlGlobal.strWallpaperImageText) Then Set Me.Picture = LoadPicture(mdlGlobal.strWallpaperImageText)
    End If

    lngRegistry = mdlRegistry.CloseRegistry(lngRegKey)
    
    Exit Sub

ErrHandler:
End Sub

Private Sub SetStatus()
    With Me.stsMain
        .Panels(UserMode).Text = Trim(mdlDatabase.GetFieldData(mdlGlobal.conInventory, "UserName", mdlTable.CreateTMUSER, "UserId='" & mdlGlobal.UserAuthority.UserId & "'")) & " ( " & UserAuthority.GetType & " )"
        .Panels(UserMode).AutoSize = sbrContents
        .Panels(UserMode).Bevel = sbrInset
        .Panels(ReminderMode).Alignment = sbrLeft
        .Panels(ReminderMode).AutoSize = sbrContents
        .Panels(TimeMode).Alignment = sbrLeft
        .Panels(TimeMode).AutoSize = sbrContents
        .Panels(TimeMode).Bevel = sbrNoBevel
    End With
End Sub

Private Sub SetLog()
    blnExists = False
    
    blnParent = True
    
    mdlProcedures.ShowForm frmLogin, , True, frmMenu.Name
    
    If Not blnExists Then
        Unload Me
    Else
        SetStatus
        
        SetClient
    End If
End Sub

Private Sub SetClient()
    SetMaster
    
    SetMIS
    
    SetTransaction
    
    SetSettings
    
    SetReport
    
    If Not mdlGlobal.UserAuthority.IsAdmin Then
        Me.mnuAddUser.Visible = False
    Else
        Me.mnuAddUser.Visible = True
    End If
    
    Me.mnuSeparatorSettings1.Visible = Me.mnuTMREMINDERCUSTOMER.Visible And Me.mnuAddUser.Visible
    
    If Me.mnuSettings.Visible Then
        Me.tlbMain.Visible = Me.mnuToolbar.Visible
        Me.stsMain.Visible = Me.mnuToolbar.Visible
    Else
        Me.tlbMain.Visible = False
        Me.stsMain.Visible = False
    End If
End Sub

Private Sub SetMaster()
    Dim strMasterRoot(12) As String
    
    strMasterRoot(0) = Me.mnuTMITEM.Name
    strMasterRoot(1) = Me.mnuTMITEMPRICE.Name
    strMasterRoot(2) = Me.mnuTMGROUP.Name
    strMasterRoot(3) = Me.mnuTMCATEGORY.Name
    strMasterRoot(4) = Me.mnuTMBRAND.Name
    strMasterRoot(5) = Me.mnuTMUNITY.Name
    strMasterRoot(6) = Me.mnuTMCUSTOMER.Name
    strMasterRoot(7) = Me.mnuTMVENDOR.Name
    strMasterRoot(8) = Me.mnuTMEMPLOYEE.Name
    strMasterRoot(9) = Me.mnuTMJOBTYPE.Name
    strMasterRoot(10) = Me.mnuTMDIVISION.Name
    strMasterRoot(11) = Me.mnuTMCURRENCY.Name
    strMasterRoot(12) = Me.mnuTMWAREHOUSE.Name
    
    If mdlGlobal.UserAuthority.IsMenuRoot(strMasterRoot) Then
        Me.mnuMaster.Visible = True
        
        frmMenu.lblMenu(MasterMode).Visible = True
        
        Dim strItemRoot(5) As String
        
        strItemRoot(0) = Me.mnuTMITEM.Name
        strItemRoot(1) = Me.mnuTMITEMPRICE.Name
        strItemRoot(2) = Me.mnuTMGROUP.Name
        strItemRoot(3) = Me.mnuTMCATEGORY.Name
        strItemRoot(4) = Me.mnuTMBRAND.Name
        strItemRoot(5) = Me.mnuTMUNITY.Name
        
        If mdlGlobal.UserAuthority.IsMenuRoot(strItemRoot) Then
            Me.mnuItem.Visible = True
            
            Me.mnuTMITEM.Visible = _
                mdlGlobal.UserAuthority.IsMenuAccess(Me.mnuTMITEM.Name)
            Me.mnuTMITEMPRICE.Visible = _
                mdlGlobal.UserAuthority.IsMenuAccess(Me.mnuTMITEMPRICE.Name)
            Me.mnuTMGROUP.Visible = _
                mdlGlobal.UserAuthority.IsMenuAccess(Me.mnuTMGROUP.Name)
            Me.mnuTMCATEGORY.Visible = _
                mdlGlobal.UserAuthority.IsMenuAccess(Me.mnuTMCATEGORY.Name)
            Me.mnuTMBRAND.Visible = _
                mdlGlobal.UserAuthority.IsMenuAccess(Me.mnuTMBRAND.Name)
            Me.mnuTMUNITY.Visible = _
                mdlGlobal.UserAuthority.IsMenuAccess(Me.mnuTMUNITY.Name)
        Else
            Me.mnuItem.Visible = False
        End If
        
        Dim strExternalRoot(1) As String
        
        strExternalRoot(0) = Me.mnuTMCUSTOMER.Name
        strExternalRoot(1) = Me.mnuTMVENDOR.Name
        
        If mdlGlobal.UserAuthority.IsMenuRoot(strExternalRoot) Then
            Me.mnuExternal.Visible = True
            
            Me.mnuTMCUSTOMER.Visible = _
                mdlGlobal.UserAuthority.IsMenuAccess(Me.mnuTMCUSTOMER.Name)
            Me.mnuTMVENDOR.Visible = _
                mdlGlobal.UserAuthority.IsMenuAccess(Me.mnuTMVENDOR.Name)
        Else
            Me.mnuExternal.Visible = False
        End If
        
        Dim strInternalRoot(4) As String
        
        strInternalRoot(0) = Me.mnuTMEMPLOYEE.Name
        strInternalRoot(1) = Me.mnuTMJOBTYPE.Name
        strInternalRoot(2) = Me.mnuTMDIVISION.Name
        strInternalRoot(3) = Me.mnuTMCURRENCY.Name
        strInternalRoot(4) = Me.mnuTMWAREHOUSE.Name
        
        If mdlGlobal.UserAuthority.IsMenuRoot(strInternalRoot) Then
            Me.mnuInternal.Visible = True
            
            Me.mnuTMEMPLOYEE.Visible = _
                mdlGlobal.UserAuthority.IsMenuAccess(Me.mnuTMEMPLOYEE.Name)
            Me.mnuTMJOBTYPE.Visible = _
                mdlGlobal.UserAuthority.IsMenuAccess(Me.mnuTMJOBTYPE.Name)
            Me.mnuTMDIVISION.Visible = _
                mdlGlobal.UserAuthority.IsMenuAccess(Me.mnuTMDIVISION.Name)
            Me.mnuTMCURRENCY.Visible = _
                mdlGlobal.UserAuthority.IsMenuAccess(Me.mnuTMCURRENCY.Name)
            Me.mnuTMWAREHOUSE.Visible = _
                mdlGlobal.UserAuthority.IsMenuAccess(Me.mnuTMWAREHOUSE.Name)
        Else
            Me.mnuInternal.Visible = False
        End If
    Else
        Me.mnuMaster.Visible = False
        
        frmMenu.lblMenu(MasterMode).Visible = False
    End If
End Sub

Private Sub SetMIS()
    Dim strMISRoot(28) As String
    
    strMISRoot(0) = Me.mnuMISTMITEM.Name
    strMISRoot(1) = Me.mnuMISTMITEMPRICE.Name
    strMISRoot(2) = Me.mnuMISTMGROUP.Name
    strMISRoot(3) = Me.mnuMISTMCATEGORY.Name
    strMISRoot(4) = Me.mnuMISTMBRAND.Name
    strMISRoot(5) = Me.mnuMISTMUNITY.Name
    strMISRoot(6) = Me.mnuMISTMCUSTOMER.Name
    strMISRoot(7) = Me.mnuMISTMVENDOR.Name
    strMISRoot(8) = Me.mnuMISTMEMPLOYEE.Name
    strMISRoot(9) = Me.mnuMISTMJOBTYPE.Name
    strMISRoot(10) = Me.mnuMISTMDIVISION.Name
    strMISRoot(11) = Me.mnuMISTMCURRENCY.Name
    strMISRoot(12) = Me.mnuMISTMWAREHOUSE.Name
    strMISRoot(13) = Me.mnuMISTMSTOCKINIT.Name
    strMISRoot(14) = Me.mnuMISTHSTOCK.Name
    strMISRoot(15) = Me.mnuMISTHITEMIN.Name
    strMISRoot(16) = Me.mnuMISTHITEMOUT.Name
    strMISRoot(17) = Me.mnuMISTHMUTITEM.Name
    strMISRoot(18) = Me.mnuMISTHPOBUY.Name
    strMISRoot(19) = Me.mnuMISTHDOBUY.Name
    strMISRoot(20) = Me.mnuMISTHSJBUY.Name
    strMISRoot(21) = Me.mnuMISTHFKTBUY.Name
    strMISRoot(22) = Me.mnuMISTHRTRBUY.Name
    strMISRoot(23) = Me.mnuMISTHSALESSUM.Name
    strMISRoot(24) = Me.mnuMISTHPOSELL.Name
    strMISRoot(25) = Me.mnuMISTHSOSELL.Name
    strMISRoot(26) = Me.mnuMISTHSJSELL.Name
    strMISRoot(27) = Me.mnuMISTHFKTSELL.Name
    strMISRoot(28) = Me.mnuMISTHRTRSELL.Name
    
    If mdlGlobal.UserAuthority.IsMenuRoot(strMISRoot) Then
        Me.mnuMIS.Visible = True
        
        frmMenu.lblMenu(MISMode).Visible = True
        
        Dim strMISItemRoot(5) As String
        
        strMISItemRoot(0) = Me.mnuMISTMITEM.Name
        strMISItemRoot(1) = Me.mnuMISTMITEMPRICE.Name
        strMISItemRoot(2) = Me.mnuMISTMGROUP.Name
        strMISItemRoot(3) = Me.mnuMISTMCATEGORY.Name
        strMISItemRoot(4) = Me.mnuMISTMBRAND.Name
        strMISItemRoot(5) = Me.mnuMISTMUNITY.Name
        
        If mdlGlobal.UserAuthority.IsMenuRoot(strMISItemRoot) Then
            Me.mnuMISItem.Visible = True
            
            Me.mnuMISTMITEM.Visible = _
                mdlGlobal.UserAuthority.IsMenuAccess(Me.mnuMISTMITEM.Name)
            Me.mnuMISTMITEMPRICE.Visible = _
                mdlGlobal.UserAuthority.IsMenuAccess(Me.mnuMISTMITEM.Name)
            Me.mnuMISTMGROUP.Visible = _
                mdlGlobal.UserAuthority.IsMenuAccess(Me.mnuMISTMGROUP.Name)
            Me.mnuMISTMCATEGORY.Visible = _
                mdlGlobal.UserAuthority.IsMenuAccess(Me.mnuMISTMCATEGORY.Name)
            Me.mnuMISTMBRAND.Visible = _
                mdlGlobal.UserAuthority.IsMenuAccess(Me.mnuMISTMBRAND.Name)
            Me.mnuMISTMUNITY.Visible = _
                mdlGlobal.UserAuthority.IsMenuAccess(Me.mnuMISTMUNITY.Name)
        Else
            Me.mnuMISItem.Visible = False
        End If
        
        Dim strMISExternalRoot(1) As String
        
        strMISExternalRoot(0) = Me.mnuMISTMCUSTOMER.Name
        strMISExternalRoot(1) = Me.mnuMISTMVENDOR.Name
        
        If mdlGlobal.UserAuthority.IsMenuRoot(strMISExternalRoot) Then
            Me.mnuMISExternal.Visible = True
            
            Me.mnuMISTMVENDOR.Visible = _
                mdlGlobal.UserAuthority.IsMenuAccess(Me.mnuMISTMVENDOR.Name)
            Me.mnuMISTMCUSTOMER.Visible = _
                mdlGlobal.UserAuthority.IsMenuAccess(Me.mnuMISTMCUSTOMER.Name)
        Else
            Me.mnuMISExternal.Visible = False
        End If
        
        Dim strMISInternalRoot(4) As String
        
        strMISInternalRoot(0) = Me.mnuMISTMEMPLOYEE.Name
        strMISInternalRoot(1) = Me.mnuMISTMJOBTYPE.Name
        strMISInternalRoot(2) = Me.mnuMISTMDIVISION.Name
        strMISInternalRoot(3) = Me.mnuMISTMCURRENCY.Name
        strMISInternalRoot(4) = Me.mnuMISTMWAREHOUSE.Name
        
        If mdlGlobal.UserAuthority.IsMenuRoot(strMISInternalRoot) Then
            Me.mnuMISInternal.Visible = True
            
            Me.mnuMISTMEMPLOYEE.Visible = _
                mdlGlobal.UserAuthority.IsMenuAccess(Me.mnuMISTMEMPLOYEE.Name)
            Me.mnuMISTMJOBTYPE.Visible = _
                mdlGlobal.UserAuthority.IsMenuAccess(Me.mnuMISTMJOBTYPE.Name)
            Me.mnuMISTMDIVISION.Visible = _
                mdlGlobal.UserAuthority.IsMenuAccess(Me.mnuMISTMDIVISION.Name)
            Me.mnuMISTMCURRENCY.Visible = _
                mdlGlobal.UserAuthority.IsMenuAccess(Me.mnuMISTMCURRENCY.Name)
            Me.mnuMISTMWAREHOUSE.Visible = _
                mdlGlobal.UserAuthority.IsMenuAccess(Me.mnuMISTMWAREHOUSE.Name)
        Else
            Me.mnuMISInternal.Visible = False
        End If
        
        Dim strMISWarehouseRoot(4) As String
        
        strMISWarehouseRoot(0) = Me.mnuMISTMSTOCKINIT.Name
        strMISWarehouseRoot(1) = Me.mnuMISTHSTOCK.Name
        strMISWarehouseRoot(2) = Me.mnuMISTHITEMIN.Name
        strMISWarehouseRoot(3) = Me.mnuMISTHITEMOUT.Name
        strMISWarehouseRoot(4) = Me.mnuMISTHMUTITEM.Name
        
        If mdlGlobal.UserAuthority.IsMenuRoot(strMISWarehouseRoot) Then
            Me.mnuMISWarehouse.Visible = True
            
            Me.mnuMISTMSTOCKINIT.Visible = _
                mdlGlobal.UserAuthority.IsMenuAccess(Me.mnuMISTMSTOCKINIT.Name)
            Me.mnuMISTHSTOCK.Visible = _
                mdlGlobal.UserAuthority.IsMenuAccess(Me.mnuMISTHSTOCK.Name)
            Me.mnuMISTHITEMIN.Visible = _
                mdlGlobal.UserAuthority.IsMenuAccess(Me.mnuMISTHITEMIN.Name)
            Me.mnuMISTHITEMOUT.Visible = _
                mdlGlobal.UserAuthority.IsMenuAccess(Me.mnuMISTHITEMOUT.Name)
            Me.mnuMISTHMUTITEM.Visible = _
                mdlGlobal.UserAuthority.IsMenuAccess(Me.mnuMISTHMUTITEM.Name)
        Else
            Me.mnuMISWarehouse.Visible = False
        End If
        
        Dim strMISBuyRoot(4) As String
        
        strMISBuyRoot(0) = Me.mnuMISTHPOBUY.Name
        strMISBuyRoot(1) = Me.mnuMISTHDOBUY.Name
        strMISBuyRoot(2) = Me.mnuMISTHSJBUY.Name
        strMISBuyRoot(3) = Me.mnuMISTHFKTBUY.Name
        strMISBuyRoot(4) = Me.mnuMISTHRTRBUY.Name
        
        If mdlGlobal.UserAuthority.IsMenuRoot(strMISBuyRoot) Then
            Me.mnuMISBuy.Visible = True
            
            Me.mnuMISTHPOBUY.Visible = _
                mdlGlobal.UserAuthority.IsMenuAccess(Me.mnuMISTHPOBUY.Name)
            Me.mnuMISTHDOBUY.Visible = _
                mdlGlobal.UserAuthority.IsMenuAccess(Me.mnuMISTHDOBUY.Name)
            Me.mnuMISTHSJBUY.Visible = _
                mdlGlobal.UserAuthority.IsMenuAccess(Me.mnuMISTHSJBUY.Name)
            Me.mnuMISTHFKTBUY.Visible = _
                mdlGlobal.UserAuthority.IsMenuAccess(Me.mnuMISTHFKTBUY.Name)
            Me.mnuMISTHRTRBUY.Visible = _
                mdlGlobal.UserAuthority.IsMenuAccess(Me.mnuMISTHRTRBUY.Name)
        Else
            Me.mnuMISBuy.Visible = False
        End If
        
        Dim strMISSellRoot(5) As String
        
        strMISSellRoot(0) = Me.mnuMISTHSALESSUM.Name
        strMISSellRoot(1) = Me.mnuMISTHPOSELL.Name
        strMISSellRoot(2) = Me.mnuMISTHSOSELL.Name
        strMISSellRoot(3) = Me.mnuMISTHSJSELL.Name
        strMISSellRoot(4) = Me.mnuMISTHFKTSELL.Name
        strMISSellRoot(5) = Me.mnuMISTHRTRSELL.Name
        
        If mdlGlobal.UserAuthority.IsMenuRoot(strMISSellRoot) Then
            Me.mnuMISSell.Visible = True
            
            Me.mnuMISTHSALESSUM.Visible = _
                mdlGlobal.UserAuthority.IsMenuAccess(Me.mnuMISTHSALESSUM.Name)
            Me.mnuMISTHPOSELL.Visible = _
                mdlGlobal.UserAuthority.IsMenuAccess(Me.mnuMISTHPOSELL.Name)
            Me.mnuMISTHSOSELL.Visible = _
                mdlGlobal.UserAuthority.IsMenuAccess(Me.mnuMISTHSOSELL.Name)
            Me.mnuMISTHSJSELL.Visible = _
                mdlGlobal.UserAuthority.IsMenuAccess(Me.mnuMISTHSJSELL.Name)
            Me.mnuMISTHFKTSELL.Visible = _
                mdlGlobal.UserAuthority.IsMenuAccess(Me.mnuMISTHFKTSELL.Name)
            Me.mnuMISTHRTRSELL.Visible = _
                mdlGlobal.UserAuthority.IsMenuAccess(Me.mnuMISTHRTRSELL.Name)
        Else
            Me.mnuMISSell.Visible = False
        End If
    Else
        Me.mnuMIS.Visible = False
        
        frmMenu.lblMenu(MISMode).Visible = False
    End If
End Sub

Private Sub SetTransaction()
    Dim strTransactionRoot(14) As String
    
    strTransactionRoot(0) = Me.mnuTMSTOCKINIT.Name
    strTransactionRoot(1) = Me.mnuTHITEMIN.Name
    strTransactionRoot(2) = Me.mnuTHITEMOUT.Name
    strTransactionRoot(3) = Me.mnuTHMUTITEM.Name
    strTransactionRoot(4) = Me.mnuTHPOBUY.Name
    strTransactionRoot(5) = Me.mnuTHDOBUY.Name
    strTransactionRoot(6) = Me.mnuTHSJBUY.Name
    strTransactionRoot(7) = Me.mnuTHFKTBUY.Name
    strTransactionRoot(8) = Me.mnuTHRTRBUY.Name
    strTransactionRoot(9) = Me.mnuTHSALESSUM.Name
    strTransactionRoot(10) = Me.mnuTHPOSELL.Name
    strTransactionRoot(11) = Me.mnuTHSOSELL.Name
    strTransactionRoot(12) = Me.mnuTHSJSELL.Name
    strTransactionRoot(13) = Me.mnuTHFKTSELL.Name
    strTransactionRoot(14) = Me.mnuTHRTRSELL.Name
    
    If mdlGlobal.UserAuthority.IsMenuRoot(strTransactionRoot) Then
        Me.mnuTransaction.Visible = True
        
        frmMenu.lblMenu(TransactionMode).Visible = True
        
        Dim strWarehouseRoot(3) As String
        
        strWarehouseRoot(0) = Me.mnuTMSTOCKINIT.Name
        strWarehouseRoot(1) = Me.mnuTHITEMIN.Name
        strWarehouseRoot(2) = Me.mnuTHITEMOUT.Name
        strWarehouseRoot(3) = Me.mnuTHMUTITEM.Name
        
        If mdlGlobal.UserAuthority.IsMenuRoot(strWarehouseRoot) Then
            Me.mnuWarehouse.Visible = True
            
            Me.mnuTMSTOCKINIT.Visible = _
                mdlGlobal.UserAuthority.IsMenuAccess(Me.mnuTMSTOCKINIT.Name)
            Me.mnuTHITEMIN.Visible = _
                mdlGlobal.UserAuthority.IsMenuAccess(Me.mnuTHITEMIN.Name)
            Me.mnuTHITEMOUT.Visible = _
                mdlGlobal.UserAuthority.IsMenuAccess(Me.mnuTHITEMOUT.Name)
            Me.mnuTHMUTITEM.Visible = _
                mdlGlobal.UserAuthority.IsMenuAccess(Me.mnuTHMUTITEM.Name)
        Else
            Me.mnuWarehouse.Visible = False
        End If
            
        Dim strBuyRoot(4) As String
        
        strBuyRoot(0) = Me.mnuTHPOBUY.Name
        strBuyRoot(1) = Me.mnuTHDOBUY.Name
        strBuyRoot(2) = Me.mnuTHSJBUY.Name
        strBuyRoot(3) = Me.mnuTHFKTBUY.Name
        strBuyRoot(4) = Me.mnuTHRTRBUY.Name
        
        If mdlGlobal.UserAuthority.IsMenuRoot(strBuyRoot) Then
            Me.mnuBuy.Visible = True
            
            Me.mnuTHPOBUY.Visible = _
                mdlGlobal.UserAuthority.IsMenuAccess(Me.mnuTHPOBUY.Name)
            Me.mnuTHDOBUY.Visible = _
                mdlGlobal.UserAuthority.IsMenuAccess(Me.mnuTHDOBUY.Name)
            Me.mnuTHSJBUY.Visible = _
                mdlGlobal.UserAuthority.IsMenuAccess(Me.mnuTHSJBUY.Name)
            Me.mnuTHFKTBUY.Visible = _
                mdlGlobal.UserAuthority.IsMenuAccess(Me.mnuTHFKTBUY.Name)
            Me.mnuTHRTRBUY.Visible = _
                mdlGlobal.UserAuthority.IsMenuAccess(Me.mnuTHRTRBUY.Name)
        Else
            Me.mnuBuy.Visible = False
        End If
        
        Dim strSellRoot(5) As String
        
        strSellRoot(0) = Me.mnuTHSALESSUM.Name
        strSellRoot(1) = Me.mnuTHPOSELL.Name
        strSellRoot(2) = Me.mnuTHSOSELL.Name
        strSellRoot(3) = Me.mnuTHSJSELL.Name
        strSellRoot(4) = Me.mnuTHFKTSELL.Name
        strSellRoot(5) = Me.mnuTHRTRSELL.Name
        
        If mdlGlobal.UserAuthority.IsMenuRoot(strSellRoot) Then
            Me.mnuSell.Visible = True
            
            Me.mnuTHSALESSUM.Visible = _
                mdlGlobal.UserAuthority.IsMenuAccess(Me.mnuTHSALESSUM.Name)
            Me.mnuTHPOSELL.Visible = _
                mdlGlobal.UserAuthority.IsMenuAccess(Me.mnuTHPOSELL.Name)
            Me.mnuTHSOSELL.Visible = _
                mdlGlobal.UserAuthority.IsMenuAccess(Me.mnuTHSOSELL.Name)
            Me.mnuTHSJSELL.Visible = _
                mdlGlobal.UserAuthority.IsMenuAccess(Me.mnuTHSJSELL.Name)
            Me.mnuTHFKTSELL.Visible = _
                mdlGlobal.UserAuthority.IsMenuAccess(Me.mnuTHRTRSELL.Name)
            Me.mnuTHRTRSELL.Visible = _
                mdlGlobal.UserAuthority.IsMenuAccess(Me.mnuTHRTRSELL.Name)
        Else
            Me.mnuSell.Visible = False
        End If
    Else
        Me.mnuTransaction.Visible = False
        
        frmMenu.lblMenu(TransactionMode).Visible = False
    End If
End Sub

Private Sub SetSettings()
    Dim strSettingsRoot(6) As String
    
    strSettingsRoot(0) = Me.mnuTMREMINDERCUSTOMER.Name
    strSettingsRoot(1) = Me.mnuPwdChange.Name
    strSettingsRoot(2) = Me.mnuBackup.Name
    strSettingsRoot(3) = Me.mnuToolbar.Name
    strSettingsRoot(4) = Me.mnuCurrencyToolbar.Name
    strSettingsRoot(5) = Me.mnuSyncFinance.Name
    strSettingsRoot(6) = Me.mnuSyncAccounting.Name
    
    If mdlGlobal.UserAuthority.IsMenuRoot(strSettingsRoot) Then
        Me.mnuSettings.Visible = True
        
        Me.mnuTMREMINDERCUSTOMER.Visible = _
            mdlGlobal.UserAuthority.IsMenuAccess(Me.mnuTMREMINDERCUSTOMER.Name)
        Me.stsMain.Panels(ReminderMode).Visible = _
            mdlGlobal.UserAuthority.IsMenuAccess(Me.mnuTMREMINDERCUSTOMER.Name)
        Me.mnuPwdChange.Visible = _
            mdlGlobal.UserAuthority.IsMenuAccess(Me.mnuPwdChange.Name)
        Me.mnuBackup.Visible = _
            mdlGlobal.UserAuthority.IsMenuAccess(Me.mnuBackup.Name)
        Me.mnuToolbar.Visible = _
            mdlGlobal.UserAuthority.IsMenuAccess(Me.mnuToolbar.Name)
        Me.mnuCurrencyToolbar.Visible = _
            mdlGlobal.UserAuthority.IsMenuAccess(Me.mnuCurrencyToolbar.Name)
        Me.mnuMenu.Visible = _
            mdlGlobal.UserAuthority.IsMenuAccess(Me.mnuMenu.Name)
        Me.mnuSeparatorSettings3.Visible = _
            mdlGlobal.UserAuthority.IsMenuAccess(Me.mnuSyncFinance.Name) And Not mdlGlobal.conFinance Is Nothing
        Me.mnuSyncFinance.Visible = _
            mdlGlobal.UserAuthority.IsMenuAccess(Me.mnuSyncFinance.Name) And Not mdlGlobal.conFinance Is Nothing
        Me.mnuSeparatorSettings4.Visible = _
            mdlGlobal.UserAuthority.IsMenuAccess(Me.mnuSyncAccounting.Name) And Not mdlGlobal.conAccounting Is Nothing
        Me.mnuSyncAccounting.Visible = _
            mdlGlobal.UserAuthority.IsMenuAccess(Me.mnuSyncAccounting.Name) And Not mdlGlobal.conAccounting Is Nothing
        
        If Me.mnuMenu.Visible Then
            frmMenu.Show
        Else
            frmMenu.Hide
        End If
    Else
        frmMenu.Hide
        
        Me.mnuSettings.Visible = False
    End If
End Sub

Private Sub SetReport()
    Dim strReportRoot(16) As String
    
    strReportRoot(0) = Me.mnuRPTTMSTOCKINIT.Name
    strReportRoot(1) = Me.mnuRPTTHSTOCK.Name
    strReportRoot(2) = Me.mnuRPTTHITEMIN.Name
    strReportRoot(3) = Me.mnuRPTTHITEMOUT.Name
    strReportRoot(4) = Me.mnuRPTTHMUTITEM.Name
    strReportRoot(5) = Me.mnuRPTTHPOBUY.Name
    strReportRoot(6) = Me.mnuRPTTHDOBUY.Name
    strReportRoot(7) = Me.mnuRPTTHSJBUY.Name
    strReportRoot(8) = Me.mnuRPTTHFKTBUY.Name
    strReportRoot(9) = Me.mnuRPTTHRTRBUY.Name
    strReportRoot(10) = Me.mnuRPTTHSALESSUM.Name
    strReportRoot(11) = Me.mnuRPTTHPOSELL.Name
    strReportRoot(12) = Me.mnuRPTTHSOSELL.Name
    strReportRoot(13) = Me.mnuRPTTHSJSELL.Name
    strReportRoot(14) = Me.mnuRPTTHFKTSELL.Name
    strReportRoot(15) = Me.mnuRPTTHRTRSELL.Name
    strReportRoot(16) = Me.mnuRPTSUMTHPOSELL.Name
    
    If mdlGlobal.UserAuthority.IsMenuRoot(strReportRoot) Then
        Me.mnuReport.Visible = True
        
        frmMenu.lblMenu(ReportMode).Visible = True
        
        Dim strWarehouseReportRoot(4) As String
        
        strWarehouseReportRoot(0) = Me.mnuRPTTMSTOCKINIT.Name
        strWarehouseReportRoot(1) = Me.mnuRPTTHSTOCK.Name
        strWarehouseReportRoot(2) = Me.mnuRPTTHITEMIN.Name
        strWarehouseReportRoot(3) = Me.mnuRPTTHITEMOUT.Name
        strWarehouseReportRoot(4) = Me.mnuRPTTHMUTITEM.Name
        
        If mdlGlobal.UserAuthority.IsMenuRoot(strWarehouseReportRoot) Then
            Me.mnuWarehouseReport.Visible = True
            
            Me.mnuRPTTMSTOCKINIT.Visible = _
                mdlGlobal.UserAuthority.IsMenuAccess(Me.mnuRPTTMSTOCKINIT.Name)
            Me.mnuRPTTHSTOCK.Visible = _
                mdlGlobal.UserAuthority.IsMenuAccess(Me.mnuRPTTHSTOCK.Name)
            Me.mnuRPTTHITEMIN.Visible = _
                mdlGlobal.UserAuthority.IsMenuAccess(Me.mnuRPTTHITEMIN.Name)
            Me.mnuRPTTHITEMOUT.Visible = _
                mdlGlobal.UserAuthority.IsMenuAccess(Me.mnuRPTTHITEMOUT.Name)
            Me.mnuRPTTHMUTITEM.Visible = _
                mdlGlobal.UserAuthority.IsMenuAccess(Me.mnuRPTTHMUTITEM.Name)
        Else
            Me.mnuWarehouseReport.Visible = False
        End If
        
        Dim strBuyReportRoot(4) As String
        
        strBuyReportRoot(0) = Me.mnuRPTTHPOBUY.Name
        strBuyReportRoot(1) = Me.mnuRPTTHDOBUY.Name
        strBuyReportRoot(2) = Me.mnuRPTTHSJBUY.Name
        strBuyReportRoot(3) = Me.mnuRPTTHFKTBUY.Name
        strBuyReportRoot(4) = Me.mnuRPTTHRTRBUY.Name
        
        If mdlGlobal.UserAuthority.IsMenuRoot(strBuyReportRoot) Then
            Me.mnuBuyReport.Visible = True
            
            Me.mnuRPTTHPOBUY.Visible = _
                mdlGlobal.UserAuthority.IsMenuAccess(Me.mnuRPTTHPOBUY.Name)
            Me.mnuRPTTHDOBUY.Visible = _
                mdlGlobal.UserAuthority.IsMenuAccess(Me.mnuRPTTHDOBUY.Name)
            Me.mnuRPTTHSJBUY.Visible = _
                mdlGlobal.UserAuthority.IsMenuAccess(Me.mnuRPTTHSJBUY.Name)
            Me.mnuRPTTHFKTBUY.Visible = _
                mdlGlobal.UserAuthority.IsMenuAccess(Me.mnuRPTTHFKTBUY.Name)
            Me.mnuRPTTHRTRBUY.Visible = _
                mdlGlobal.UserAuthority.IsMenuAccess(Me.mnuRPTTHRTRBUY.Name)
        Else
            Me.mnuBuyReport.Visible = False
        End If
        
        Dim strSellReportRoot(5) As String
        
        strSellReportRoot(0) = Me.mnuRPTTHSALESSUM.Name
        strSellReportRoot(1) = Me.mnuRPTTHPOSELL.Name
        strSellReportRoot(2) = Me.mnuRPTTHSOSELL.Name
        strSellReportRoot(3) = Me.mnuRPTTHSJSELL.Name
        strSellReportRoot(4) = Me.mnuRPTTHFKTSELL.Name
        strSellReportRoot(5) = Me.mnuRPTTHRTRSELL.Name
        
        If mdlGlobal.UserAuthority.IsMenuRoot(strSellReportRoot) Then
            Me.mnuSellReport.Visible = True
            
            Me.mnuRPTTHSALESSUM.Visible = _
                mdlGlobal.UserAuthority.IsMenuAccess(Me.mnuRPTTHSALESSUM.Name)
            Me.mnuRPTTHPOSELL.Visible = _
                mdlGlobal.UserAuthority.IsMenuAccess(Me.mnuRPTTHPOSELL.Name)
            Me.mnuRPTTHSOSELL.Visible = _
                mdlGlobal.UserAuthority.IsMenuAccess(Me.mnuRPTTHSOSELL.Name)
            Me.mnuRPTTHSJSELL.Visible = _
                mdlGlobal.UserAuthority.IsMenuAccess(Me.mnuRPTTHSJSELL.Name)
            Me.mnuRPTTHFKTSELL.Visible = _
                mdlGlobal.UserAuthority.IsMenuAccess(Me.mnuRPTTHFKTSELL.Name)
            Me.mnuRPTTHRTRSELL.Visible = _
                mdlGlobal.UserAuthority.IsMenuAccess(Me.mnuRPTTHRTRSELL.Name)
        Else
            Me.mnuSellReport.Visible = False
        End If
        
        Dim strExternalReportRoot(0) As String
        
        strExternalReportRoot(0) = Me.mnuExternalReport.Name
        
        If mdlGlobal.UserAuthority.IsMenuRoot(strExternalReportRoot) Then
            Me.mnuExternalReport.Visible = True
            
            Dim strCustomerReportRoot(0) As String
            
            strCustomerReportRoot(0) = Me.mnuRPTSUMTHPOSELL.Name
            
            If mdlGlobal.UserAuthority.IsMenuRoot(strCustomerReportRoot) Then
                Me.mnuCustomerReport.Visible = True
                
                Me.mnuRPTSUMTHPOSELL.Visible = _
                    mdlGlobal.UserAuthority.IsMenuAccess(Me.mnuRPTSUMTHPOSELL.Name)
            Else
                Me.mnuCustomerReport.Visible = False
            End If
        Else
            Me.mnuExternalReport.Visible = False
        End If
    Else
        Me.mnuReport.Visible = False
        
        frmMenu.lblMenu(ReportMode).Visible = False
    End If
End Sub

Public Property Let Exists(ByVal blnEnable As Boolean)
    blnExists = blnEnable
End Property

Public Property Let Parent(ByVal blnEnable As Boolean)
    blnParent = blnEnable
End Property

Public Property Let Reminder(ByVal blnEnable As Boolean)
    blnReminder = blnEnable
End Property

Public Property Get Parent() As Boolean
    Parent = blnParent
End Property

Public Property Get Reminder() As Boolean
    Reminder = blnReminder
End Property

Private Sub tmrMain_Timer()
    Me.stsMain.Panels(TimeMode).Text = mdlProcedures.FormatDate(Now, "dd-MMMM-yyyy (hh:mm:ss)")
End Sub

Public Sub CheckReminder()
    Dim curCount As Currency
    
    curCount = mdlProcedures.GetCurrency(ReminderCustomer) + mdlProcedures.GetCurrency(ReminderPOSELL)
    
    If curCount > 0 Then
        Me.stsMain.Panels(ReminderMode).Text = curCount & " Pengingat "
        Me.stsMain.Panels(ReminderMode).Picture = Me.imlStatus.ListImages(3).Picture
        Me.stsMain.Panels(ReminderMode).Bevel = sbrRaised
        
        DoEvents
        
        mdlAPI.Beep 2300, 500
    Else
        Me.stsMain.Panels(ReminderMode).Text = "Tidak Ada Pengingat "
        Me.stsMain.Panels(ReminderMode).Picture = Me.imlStatus.ListImages(2).Picture
        Me.stsMain.Panels(ReminderMode).Bevel = sbrRaised
    End If
End Sub

Public Sub CheckConvertCurrency()
    Dim rstTemp As ADODB.Recordset
    
    Set rstTemp = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "CurrencyId", mdlTable.CreateTMCURRENCY, False)
    
    Dim strCurrencyFromId As String
    Dim strCurrencyToId As String
    
    strCurrencyFromId = ""
    strCurrencyToId = ""
    
    If mdlProcedures.IsValidComboData(Me.cmbCurrencyFromId) Then
        strCurrencyFromId = mdlProcedures.GetComboData(Me.cmbCurrencyFromId)
    End If
    
    If mdlProcedures.IsValidComboData(Me.cmbCurrencyToId) Then
        strCurrencyToId = mdlProcedures.GetComboData(Me.cmbCurrencyToId)
    End If
    
    mdlProcedures.FillComboData Me.cmbCurrencyFromId, rstTemp
    mdlProcedures.FillComboData Me.cmbCurrencyToId, rstTemp
    
    If Not mdlProcedures.RepDupText(Trim(strCurrencyFromId)) = "" Then
        mdlProcedures.SetComboData Me.cmbCurrencyFromId, strCurrencyFromId
    End If
    
    If Not mdlProcedures.RepDupText(Trim(strCurrencyToId)) = "" Then
        mdlProcedures.SetComboData Me.cmbCurrencyToId, strCurrencyToId
    End If
    
    mdlDatabase.CloseRecordset rstTemp
    
    If mdlProcedures.RepDupText(Trim(strCurrencyToId)) = "" And mdlProcedures.RepDupText(Trim(strCurrencyFromId)) = "" Then
        Set rstTemp = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "*", mdlTable.CreateTMCONVERTCURRENCY, False, "", "UpdateDate DESC")
        
        With rstTemp
            If .RecordCount > 0 Then
                mdlProcedures.SetComboData Me.cmbCurrencyFromId, !CurrencyFromId
                mdlProcedures.SetComboData Me.cmbCurrencyToId, !CurrencyToId
                
                Me.txtConvertValue.Text = mdlProcedures.GetCurrency(!ConvertValue)
            End If
        End With
        
        mdlDatabase.CloseRecordset rstTemp
    End If
End Sub

Private Sub CheckConvertValue()
    Dim strCurrencyFromId As String
    Dim strCurrencyToId As String
    
    strCurrencyFromId = ""
    strCurrencyToId = ""
    
    If mdlProcedures.IsValidComboData(Me.cmbCurrencyFromId) Then
        strCurrencyFromId = mdlProcedures.GetComboData(Me.cmbCurrencyFromId)
    End If
    
    If mdlProcedures.IsValidComboData(Me.cmbCurrencyToId) Then
        strCurrencyToId = mdlProcedures.GetComboData(Me.cmbCurrencyToId)
    End If
    
    If mdlProcedures.RepDupText(Trim(strCurrencyFromId)) = "" Or mdlProcedures.RepDupText(Trim(strCurrencyToId)) = "" Then
        Me.txtConvertValue.Text = "0"
        
        blnConvertCurrency = False
    Else
        If mdlProcedures.RepDupText(Trim(strCurrencyFromId)) = mdlProcedures.RepDupText(Trim(strCurrencyToId)) Then
            Me.txtConvertValue.Text = "0"
            
            blnConvertCurrency = False
        Else
            Me.txtConvertValue.Text = mdlDatabase.GetFieldData(mdlGlobal.conInventory, "ConvertValue", mdlTable.CreateTMCONVERTCURRENCY, "CurrencyFromId='" & strCurrencyFromId & "' AND CurrencyToId='" & strCurrencyToId & "'", "ConvertDate ASC")
            
            blnConvertCurrency = True
        End If
    End If
End Sub

Private Function ReminderCustomer() As Integer
    Dim rstTemp As ADODB.Recordset
    
    If mdlGlobal.objDatabaseInit = SQLSERVER7 Or _
        mdlGlobal.objDatabaseInit = SQLSERVER2000 Or _
        mdlGlobal.objDatabaseInit = SQLEXPRESS Then
        Set rstTemp = _
            mdlDatabase.OpenRecordset( _
                mdlGlobal.conInventory, _
                "*", _
                mdlTable.CreateTMREMINDERCUSTOMER, _
                False, _
                "ValidateType<>'" & ValidateType.NoneValidate & "' AND ValidateDate<='" & mdlProcedures.FormatDate(Now) & "'")
    ElseIf mdlGlobal.objDatabaseInit = MSACCESS Then
        Set rstTemp = _
            mdlDatabase.OpenRecordset( _
                mdlGlobal.conInventory, _
                "*", _
                mdlTable.CreateTMREMINDERCUSTOMER, _
                False, _
                "ValidateType<>'" & ValidateType.NoneValidate & "' AND ValidateDate<=#" & mdlProcedures.FormatDate(Now) & "#")
    ElseIf mdlGlobal.objDatabaseInit = MYSQL Then
        Set rstTemp = _
            mdlDatabase.OpenRecordset( _
                mdlGlobal.conInventory, _
                "*", _
                mdlTable.CreateTMREMINDERCUSTOMER, _
                False, _
                "ValidateType<>'" & ValidateType.NoneValidate & "' AND ValidateDate<='" & mdlProcedures.FormatDate(Now) & "'")
    End If
    
    Dim intCount As Integer
    
    intCount = 0
    
    intCount = intCount + rstTemp.RecordCount
    
    mdlDatabase.CloseRecordset rstTemp
    
    ReminderCustomer = intCount
End Function

Private Function ReminderPOSELL() As Integer
    Dim rstTemp As ADODB.Recordset
    
    If mdlGlobal.objDatabaseInit = SQLSERVER7 Or _
        mdlGlobal.objDatabaseInit = SQLSERVER2000 Or _
        mdlGlobal.objDatabaseInit = SQLEXPRESS Then
        Set rstTemp = _
            mdlDatabase.OpenRecordset( _
                mdlGlobal.conInventory, _
                "*", _
                mdlTable.CreateTHPOSELL, _
                False, _
                "DateLine<='" & mdlProcedures.FormatDate(Now) & "'")
    ElseIf mdlGlobal.objDatabaseInit = MSACCESS Then
        Set rstTemp = _
            mdlDatabase.OpenRecordset( _
                mdlGlobal.conInventory, _
                "*", _
                mdlTable.CreateTHPOSELL, _
                False, _
                "DateLine<=#" & mdlProcedures.FormatDate(Now) & "#")
    ElseIf mdlGlobal.objDatabaseInit = MYSQL Then
        Set rstTemp = _
            mdlDatabase.OpenRecordset( _
                mdlGlobal.conInventory, _
                "*", _
                mdlTable.CreateTHPOSELL, _
                False, _
                "DateLine<='" & mdlProcedures.FormatDate(Now) & "'")
    End If
    
    Dim curQtyPO As Currency
    Dim curQtySJ As Currency
    
    Dim intCount As Integer
    
    intCount = 0
    
    With rstTemp
        While Not .EOF
            curQtyPO = mdlTHPOSELL.GetTotalQtyPOSELL(!POId)
            curQtySJ = mdlTHSJSELL.GetQtyPOFromSJSELL(!POId)
            
            If curQtySJ < curQtyPO Then
                intCount = intCount + 1
            End If
            
            .MoveNext
        Wend
    End With
    
    mdlDatabase.CloseRecordset rstTemp
    
    ReminderPOSELL = intCount
End Function

Public Property Get CloseAll() As Boolean
    CloseAll = Me.mnuCloseAll.Checked
End Property
