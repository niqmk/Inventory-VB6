VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmAddUser 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7335
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   8190
   Icon            =   "frmAddUser.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7335
   ScaleWidth      =   8190
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Hapus"
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
      Left            =   5040
      TabIndex        =   6
      Top             =   6840
      Width           =   1455
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Simpan"
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
      Left            =   6600
      TabIndex        =   5
      Top             =   6840
      Width           =   1455
   End
   Begin MSFlexGridLib.MSFlexGrid flxMain 
      Height          =   4095
      Left            =   120
      TabIndex        =   4
      Top             =   2640
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   7223
      _Version        =   393216
      FocusRect       =   2
      GridLines       =   0
      MergeCells      =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame fraDetail 
      Height          =   1575
      Left            =   120
      TabIndex        =   13
      Top             =   960
      Width           =   7935
      Begin VB.TextBox txtUserPwd 
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
         Left            =   1680
         MaxLength       =   20
         TabIndex        =   3
         Top             =   1200
         Width           =   2535
      End
      Begin VB.TextBox txtUserName 
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
         Left            =   1680
         MaxLength       =   50
         TabIndex        =   1
         Top             =   240
         Width           =   6135
      End
      Begin VB.ComboBox cmbUserType 
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
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   720
         Width           =   6135
      End
      Begin VB.Label lblUserPwd 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
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
         TabIndex        =   11
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label lblUserType 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipe Pemakai"
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
         TabIndex        =   10
         Top             =   720
         Width           =   1185
      End
      Begin VB.Label lblUserName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Pemakai"
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
         TabIndex        =   9
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame fraHeader 
      Height          =   735
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   7935
      Begin VB.CommandButton cmdUserId 
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
         Left            =   3000
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox txtUserId 
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
         Left            =   1680
         MaxLength       =   8
         TabIndex        =   0
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label lblUserId 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Id Pemakai"
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
         TabIndex        =   8
         Top             =   240
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmAddUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum ColumnConstants
    [BlankColumn]
    [MenuRootColumn]
    [MenuIdColumn]
    [MenuNameColumn]
    [AccessYNColumn]
End Enum

Private Const TMITEM_Row As Integer = 2
Private Const TMPRICELIST_Row As Integer = 3
Private Const TMITEMPRICE_Row As Integer = 4
Private Const TMGROUP_Row As Integer = 5
Private Const TMCATEGORY_Row As Integer = 6
Private Const TMBRAND_Row As Integer = 7
Private Const TMUNITY_Row As Integer = 8
Private Const TMCUSTOMER_Row As Integer = 9
Private Const TMVENDOR_Row As Integer = 10
Private Const TMEMPLOYEE_Row As Integer = 11
Private Const TMJOBTYPE_Row As Integer = 12
Private Const TMDIVISION_Row As Integer = 13
Private Const TMCURRENCY_Row As Integer = 14
Private Const TMWAREHOUSE_Row As Integer = 15

Private Const SEPARATOR1_Row As Integer = 16

Private Const MISTMITEM_Row As Integer = 17
Private Const MISTMPRICELIST_Row As Integer = 18
Private Const MISTMITEMPRICE_Row As Integer = 19
Private Const MISTMGROUP_Row As Integer = 20
Private Const MISTMCATEGORY_Row As Integer = 21
Private Const MISTMBRAND_Row As Integer = 22
Private Const MISTMUNITY_Row As Integer = 23
Private Const MISTMCUSTOMER_Row As Integer = 24
Private Const MISTMCUSTOMERTRANS_Row As Integer = 25
Private Const MISTMVENDOR_Row As Integer = 26
Private Const MISTMEMPLOYEE_Row As Integer = 27
Private Const MISTMJOBTYPE_Row As Integer = 28
Private Const MISTMDIVISION_Row As Integer = 29
Private Const MISTMCURRENCY_Row As Integer = 30
Private Const MISTMWAREHOUSE_Row As Integer = 31
Private Const MISTMSTOCKINIT_Row As Integer = 32
Private Const MISTHSTOCK_Row As Integer = 33
Private Const MISTHITEMIN_Row As Integer = 34
Private Const MISTHITEMOUT_Row As Integer = 35
Private Const MISTHMUTITEM_Row As Integer = 36
Private Const MISTHPOBUY_Row As Integer = 37
Private Const MISTHDOBUY_Row As Integer = 38
Private Const MISTHSJBUY_Row As Integer = 39
Private Const MISTHFKTBUY_Row As Integer = 40
Private Const MISTHRTRBUY_Row As Integer = 41
Private Const MISTHSALESSUM_Row As Integer = 42
Private Const MISTHPOSELL_Row As Integer = 43
Private Const MISTHSOSELL_Row As Integer = 44
Private Const MISTHSJSELL_Row As Integer = 45
Private Const MISTHFKTSELL_Row As Integer = 46
Private Const MISTHRTRSELL_Row As Integer = 47

Private Const SEPARATOR2_Row As Integer = 48

Private Const TMSTOCKINIT_Row As Integer = 49
Private Const THITEMIN_Row As Integer = 50
Private Const THITEMOUT_Row As Integer = 51
Private Const THMUTITEM_Row As Integer = 52
Private Const THPOBUY_Row As Integer = 53
Private Const THDOBUY_Row As Integer = 54
Private Const THSJBUY_Row As Integer = 55
Private Const THFKTBUY_Row As Integer = 56
Private Const THRTRBUY_Row As Integer = 57
Private Const THSALESSUM_Row As Integer = 58
Private Const THPOSELL_Row As Integer = 59
Private Const THSOSELL_Row As Integer = 60
Private Const THSJSELL_Row As Integer = 61
Private Const THFKTSELL_Row As Integer = 62
Private Const THRTRSELL_Row As Integer = 63

Private Const SEPARATOR3_Row As Integer = 64

Private Const RPTTMSTOCKINIT_Row As Integer = 65
Private Const RPTTHSTOCK_Row As Integer = 66
Private Const RPTTHITEMIN_Row As Integer = 67
Private Const RPTTHITEMOUT_Row As Integer = 68
Private Const RPTTHMUTITEM_Row As Integer = 69
Private Const RPTTHPOBUY_Row As Integer = 70
Private Const RPTTHDOBUY_Row As Integer = 71
Private Const RPTTHSJBUY_Row As Integer = 72
Private Const RPTTHFKTBUY_Row As Integer = 73
Private Const RPTTHRTRBUY_Row As Integer = 74
Private Const RPTTHPOSELL_Row As Integer = 75
Private Const RPTTHSOSELL_Row As Integer = 76
Private Const RPTTHSJSELL_Row As Integer = 77
Private Const RPTTHFKTSELL_Row As Integer = 78
Private Const RPTTHRTRSELL_Row As Integer = 79
Private Const RPTSUMTHPOSELL_Row As Integer = 80

Private Const SEPARATOR4_Row As Integer = 81

Private Const TMREMINDERCUSTOMER_Row As Integer = 82
Private Const PwdChange_Row As Integer = 83
Private Const Backup_Row As Integer = 84
Private Const Toolbar_Row As Integer = 85
Private Const CurrencyToolbar_Row As Integer = 86
Private Const Menu_Row As Integer = 87
Private Const THRECYCLE_Row As Integer = 88
Private Const SyncFinance_Row As Integer = 89
Private Const SyncAccounting_Row As Integer = 90

Private blnParent As Boolean

Private Sub Form_Load()
    SetInitialization
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If blnParent Then
        Cancel = 1
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmAddUser = Nothing
End Sub

Private Sub txtUserId_GotFocus()
    mdlProcedures.GotFocus Me.txtUserId
End Sub

Private Sub txtUserName_GotFocus()
    mdlProcedures.GotFocus Me.txtUserName
End Sub

Private Sub txtUserPwd_GotFocus()
    mdlProcedures.GotFocus Me.txtUserPwd
End Sub

Private Sub txtUserId_Validate(Cancel As Boolean)
    If Trim(Me.txtUserId.Text) = "" Then
        ClearText
        
        Exit Sub
    End If
    
    FillText Trim(Me.txtUserId.Text)
End Sub

Private Sub flxMain_DblClick()
    With Me.flxMain
        If .Col = AccessYNColumn Then
            If Not .TextMatrix(.Row, MenuNameColumn) = "" Then
                If .TextMatrix(.Row, .Col) = Chr(254) Then
                    .TextMatrix(.Row, .Col) = ""
                Else
                    .TextMatrix(.Row, .Col) = Chr(254)
                End If
            End If
        End If
    End With
End Sub

Private Sub cmdUserId_Click()
    If blnParent Then Exit Sub
    
    blnParent = True
    
    mdlProcedures.CornerWindows Me
    
    mdlProcedures.ShowForm frmBRWTMUSER, False, True
End Sub

Private Sub cmdSave_Click()
    SaveFunction
End Sub

Private Sub cmdDelete_Click()
    mdlDatabase.DeleteRecordQuery mdlGlobal.conInventory, mdlTable.CreateTMUSER, "UserId='" & mdlProcedures.RepDupText(Me.txtUserId.Text) & "'"
    mdlDatabase.DeleteRecordQuery mdlGlobal.conInventory, mdlTable.CreateTMMENUAUTHORITY, "UserId='" & mdlProcedures.RepDupText(Me.txtUserId.Text) & "'"
    
    ClearText True
    
    Me.txtUserId.SetFocus
End Sub

Private Sub SetInitialization()
    mdlProcedures.CenterWindows Me
    
    Me.Caption = mdlText.strAddUser
    
    blnParent = False
    
    With Me.cmbUserType
        .AddItem mdlGlobal.strAdministrator
        .AddItem mdlGlobal.strUser
    End With
    
    If Me.cmbUserType.ListCount > 0 Then
        Me.cmbUserType.ListIndex = 0
    End If
    
    Me.txtUserPwd.PasswordChar = Chr(183)
    
    ArrangeGrid
End Sub

Private Sub ArrangeGrid()
    With Me.flxMain
        .Rows = 89
        .Cols = AccessYNColumn + 1
        
        .ColWidth(BlankColumn) = 300
        .ColWidth(MenuRootColumn) = 2000
        .ColWidth(MenuIdColumn) = 0
        .ColWidth(MenuNameColumn) = 3800
        .ColWidth(AccessYNColumn) = 1530
        
        .ColAlignment(MenuNameColumn) = flexAlignLeftCenter
        .ColAlignment(AccessYNColumn) = flexAlignCenterCenter
        
        .BackColorSel = &HC0E000
        .ForeColorSel = &HFF8099
        
        .TextMatrix(0, MenuRootColumn) = "Root Menu"
        .TextMatrix(0, MenuNameColumn) = "Nama Menu"
        .TextMatrix(0, AccessYNColumn) = "Hak Akses"
        
        .TextMatrix(1, MenuRootColumn) = mdiMain.mnuMaster.Caption
        
        SetMenuMatrix TMITEM_Row, mdiMain.mnuTMITEM.Name, mdiMain.mnuTMITEM.Caption
        SetMenuMatrix TMPRICELIST_Row, mdiMain.mnuTMPRICELIST.Name, mdiMain.mnuTMPRICELIST.Caption
        SetMenuMatrix TMITEMPRICE_Row, mdiMain.mnuTMITEMPRICE.Name, mdiMain.mnuTMITEMPRICE.Caption
        SetMenuMatrix TMGROUP_Row, mdiMain.mnuTMGROUP.Name, mdiMain.mnuTMGROUP.Caption
        SetMenuMatrix TMCATEGORY_Row, mdiMain.mnuTMCATEGORY.Name, mdiMain.mnuTMCATEGORY.Caption
        SetMenuMatrix TMBRAND_Row, mdiMain.mnuTMBRAND.Name, mdiMain.mnuTMBRAND.Caption
        SetMenuMatrix TMUNITY_Row, mdiMain.mnuTMUNITY.Name, mdiMain.mnuTMUNITY.Caption
        SetMenuMatrix TMCUSTOMER_Row, mdiMain.mnuTMCUSTOMER.Name, mdiMain.mnuTMCUSTOMER.Caption
        SetMenuMatrix TMVENDOR_Row, mdiMain.mnuTMVENDOR.Name, mdiMain.mnuTMVENDOR.Caption
        SetMenuMatrix TMEMPLOYEE_Row, mdiMain.mnuTMEMPLOYEE.Name, mdiMain.mnuTMEMPLOYEE.Caption
        SetMenuMatrix TMJOBTYPE_Row, mdiMain.mnuTMJOBTYPE.Name, mdiMain.mnuTMJOBTYPE.Caption
        SetMenuMatrix TMDIVISION_Row, mdiMain.mnuTMDIVISION.Name, mdiMain.mnuTMDIVISION.Caption
        SetMenuMatrix TMCURRENCY_Row, mdiMain.mnuTMCURRENCY.Name, mdiMain.mnuTMCURRENCY.Caption
        SetMenuMatrix TMWAREHOUSE_Row, mdiMain.mnuTMWAREHOUSE.Name, mdiMain.mnuTMWAREHOUSE.Caption
        
        .TextMatrix(SEPARATOR1_Row, MenuRootColumn) = mdiMain.mnuMIS.Caption
        
        SetMenuMatrix MISTMITEM_Row, mdiMain.mnuMISTMITEM.Name, mdiMain.mnuMISTMITEM.Caption
        SetMenuMatrix MISTMPRICELIST_Row, mdiMain.mnuMISTMPRICELIST.Name, mdiMain.mnuMISTMPRICELIST.Caption
        SetMenuMatrix MISTMITEMPRICE_Row, mdiMain.mnuMISTMITEMPRICE.Name, mdiMain.mnuMISTMITEMPRICE.Caption
        SetMenuMatrix MISTMGROUP_Row, mdiMain.mnuMISTMGROUP.Name, mdiMain.mnuMISTMGROUP.Caption
        SetMenuMatrix MISTMCATEGORY_Row, mdiMain.mnuMISTMCATEGORY.Name, mdiMain.mnuMISTMCATEGORY.Caption
        SetMenuMatrix MISTMBRAND_Row, mdiMain.mnuMISTMBRAND.Name, mdiMain.mnuMISTMBRAND.Caption
        SetMenuMatrix MISTMUNITY_Row, mdiMain.mnuMISTMUNITY.Name, mdiMain.mnuMISTMUNITY.Caption
        SetMenuMatrix MISTMCUSTOMER_Row, mdiMain.mnuMISTMCUSTOMER.Name, mdiMain.mnuMISTMCUSTOMER.Caption
        SetMenuMatrix MISTMCUSTOMERTRANS_Row, mdiMain.mnuMISTMCUSTOMERTRANS.Name, mdiMain.mnuMISTMCUSTOMERTRANS.Caption
        SetMenuMatrix MISTMVENDOR_Row, mdiMain.mnuMISTMVENDOR.Name, mdiMain.mnuMISTMVENDOR.Caption
        SetMenuMatrix MISTMEMPLOYEE_Row, mdiMain.mnuMISTMEMPLOYEE.Name, mdiMain.mnuMISTMEMPLOYEE.Caption
        SetMenuMatrix MISTMJOBTYPE_Row, mdiMain.mnuMISTMJOBTYPE.Name, mdiMain.mnuMISTMJOBTYPE.Caption
        SetMenuMatrix MISTMDIVISION_Row, mdiMain.mnuMISTMDIVISION.Name, mdiMain.mnuMISTMDIVISION.Caption
        SetMenuMatrix MISTMCURRENCY_Row, mdiMain.mnuMISTMCURRENCY.Name, mdiMain.mnuMISTMCURRENCY.Caption
        SetMenuMatrix MISTMWAREHOUSE_Row, mdiMain.mnuMISTMWAREHOUSE.Name, mdiMain.mnuMISTMWAREHOUSE.Caption
        SetMenuMatrix MISTMSTOCKINIT_Row, mdiMain.mnuMISTMSTOCKINIT.Name, mdiMain.mnuMISTMSTOCKINIT.Caption & " (" & mdiMain.mnuMISWarehouse.Caption & ")"
        SetMenuMatrix MISTHSTOCK_Row, mdiMain.mnuMISTHSTOCK.Name, mdiMain.mnuMISTHSTOCK.Caption & " (" & mdiMain.mnuMISWarehouse.Caption & ")"
        SetMenuMatrix MISTHITEMIN_Row, mdiMain.mnuMISTHITEMIN.Name, mdiMain.mnuMISTHITEMIN.Caption & " (" & mdiMain.mnuMISWarehouse.Caption & ")"
        SetMenuMatrix MISTHITEMOUT_Row, mdiMain.mnuMISTHITEMOUT.Name, mdiMain.mnuMISTHITEMOUT.Caption & " (" & mdiMain.mnuMISWarehouse.Caption & ")"
        SetMenuMatrix MISTHMUTITEM_Row, mdiMain.mnuMISTHMUTITEM.Name, mdiMain.mnuMISTHMUTITEM.Caption & " (" & mdiMain.mnuMISWarehouse.Caption & ")"
        SetMenuMatrix MISTHPOBUY_Row, mdiMain.mnuMISTHPOBUY.Name, mdiMain.mnuMISTHPOBUY.Caption & " (" & mdiMain.mnuMISBuy.Caption & ")"
        SetMenuMatrix MISTHDOBUY_Row, mdiMain.mnuMISTHDOBUY.Name, mdiMain.mnuMISTHDOBUY.Caption & " (" & mdiMain.mnuMISBuy.Caption & ")"
        SetMenuMatrix MISTHSJBUY_Row, mdiMain.mnuMISTHSJBUY.Name, mdiMain.mnuMISTHSJBUY.Caption & " (" & mdiMain.mnuMISBuy.Caption & ")"
        SetMenuMatrix MISTHFKTBUY_Row, mdiMain.mnuMISTHFKTBUY.Name, mdiMain.mnuMISTHFKTBUY.Caption & " (" & mdiMain.mnuMISBuy.Caption & ")"
        SetMenuMatrix MISTHRTRBUY_Row, mdiMain.mnuMISTHRTRBUY.Name, mdiMain.mnuMISTHRTRBUY.Caption & " (" & mdiMain.mnuMISBuy.Caption & ")"
        SetMenuMatrix MISTHSALESSUM_Row, mdiMain.mnuMISTHSALESSUM.Name, mdiMain.mnuMISTHSALESSUM.Caption & " (" & mdiMain.mnuMISSell.Caption & ")"
        SetMenuMatrix MISTHPOSELL_Row, mdiMain.mnuMISTHPOSELL.Name, mdiMain.mnuMISTHPOSELL.Caption & " (" & mdiMain.mnuMISSell.Caption & ")"
        SetMenuMatrix MISTHSOSELL_Row, mdiMain.mnuMISTHSOSELL.Name, mdiMain.mnuMISTHSOSELL.Caption & " (" & mdiMain.mnuMISSell.Caption & ")"
        SetMenuMatrix MISTHSJSELL_Row, mdiMain.mnuMISTHSJSELL.Name, mdiMain.mnuMISTHSJSELL.Caption & " (" & mdiMain.mnuMISSell.Caption & ")"
        SetMenuMatrix MISTHFKTSELL_Row, mdiMain.mnuMISTHFKTSELL.Name, mdiMain.mnuMISTHFKTSELL.Caption & " (" & mdiMain.mnuMISSell.Caption & ")"
        SetMenuMatrix MISTHRTRSELL_Row, mdiMain.mnuMISTHRTRSELL.Name, mdiMain.mnuMISTHRTRSELL.Caption & " (" & mdiMain.mnuMISSell.Caption & ")"
        
        .TextMatrix(SEPARATOR2_Row, MenuRootColumn) = mdiMain.mnuTransaction.Caption
        
        SetMenuMatrix TMSTOCKINIT_Row, mdiMain.mnuTMSTOCKINIT.Name, mdiMain.mnuTMSTOCKINIT.Caption & " (" & mdiMain.mnuWarehouse.Caption & ")"
        SetMenuMatrix THITEMIN_Row, mdiMain.mnuTHITEMIN.Name, mdiMain.mnuTHITEMIN.Caption & " (" & mdiMain.mnuWarehouse.Caption & ")"
        SetMenuMatrix THITEMOUT_Row, mdiMain.mnuTHITEMOUT.Name, mdiMain.mnuTHITEMOUT.Caption & " (" & mdiMain.mnuWarehouse.Caption & ")"
        SetMenuMatrix THMUTITEM_Row, mdiMain.mnuTHMUTITEM.Name, mdiMain.mnuTHMUTITEM.Caption & " (" & mdiMain.mnuWarehouse.Caption & ")"
        
        SetMenuMatrix THPOBUY_Row, mdiMain.mnuTHPOBUY.Name, mdiMain.mnuTHPOBUY.Caption & " (" & mdiMain.mnuBuy.Caption & ")"
        SetMenuMatrix THDOBUY_Row, mdiMain.mnuTHDOBUY.Name, mdiMain.mnuTHDOBUY.Caption & " (" & mdiMain.mnuBuy.Caption & ")"
        SetMenuMatrix THSJBUY_Row, mdiMain.mnuTHSJBUY.Name, mdiMain.mnuTHSJBUY.Caption & " (" & mdiMain.mnuBuy.Caption & ")"
        SetMenuMatrix THFKTBUY_Row, mdiMain.mnuTHFKTBUY.Name, mdiMain.mnuTHFKTBUY.Caption & " (" & mdiMain.mnuBuy.Caption & ")"
        SetMenuMatrix THRTRBUY_Row, mdiMain.mnuTHRTRBUY.Name, mdiMain.mnuTHRTRBUY.Caption & " (" & mdiMain.mnuBuy.Caption & ")"
        
        SetMenuMatrix THSALESSUM_Row, mdiMain.mnuTHSALESSUM.Name, mdiMain.mnuTHSALESSUM.Caption & " (" & mdiMain.mnuSell.Caption & ")"
        SetMenuMatrix THPOSELL_Row, mdiMain.mnuTHPOSELL.Name, mdiMain.mnuTHPOSELL.Caption & " (" & mdiMain.mnuSell.Caption & ")"
        SetMenuMatrix THSOSELL_Row, mdiMain.mnuTHSOSELL.Name, mdiMain.mnuTHSOSELL.Caption & " (" & mdiMain.mnuSell.Caption & ")"
        SetMenuMatrix THSJSELL_Row, mdiMain.mnuTHSJSELL.Name, mdiMain.mnuTHSJSELL.Caption & " (" & mdiMain.mnuSell.Caption & ")"
        SetMenuMatrix THFKTSELL_Row, mdiMain.mnuTHFKTSELL.Name, mdiMain.mnuTHFKTSELL.Caption & " (" & mdiMain.mnuSell.Caption & ")"
        SetMenuMatrix THRTRSELL_Row, mdiMain.mnuTHRTRSELL.Name, mdiMain.mnuTHRTRSELL.Caption & " (" & mdiMain.mnuSell.Caption & ")"
        
        .TextMatrix(SEPARATOR3_Row, MenuRootColumn) = mdiMain.mnuReport.Caption
        
        SetMenuMatrix RPTTMSTOCKINIT_Row, mdiMain.mnuRPTTMSTOCKINIT.Name, mdiMain.mnuRPTTMSTOCKINIT.Caption & " (" & mdiMain.mnuWarehouseReport.Caption & ")"
        SetMenuMatrix RPTTHSTOCK_Row, mdiMain.mnuRPTTHSTOCK.Name, mdiMain.mnuRPTTHSTOCK.Caption & " (" & mdiMain.mnuWarehouseReport.Caption & ")"
        SetMenuMatrix RPTTHITEMIN_Row, mdiMain.mnuRPTTHITEMIN.Name, mdiMain.mnuRPTTHITEMIN.Caption & " (" & mdiMain.mnuWarehouseReport.Caption & ")"
        SetMenuMatrix RPTTHITEMOUT_Row, mdiMain.mnuRPTTHITEMOUT.Name, mdiMain.mnuRPTTHITEMOUT.Caption & " (" & mdiMain.mnuWarehouseReport.Caption & ")"
        SetMenuMatrix RPTTHMUTITEM_Row, mdiMain.mnuRPTTHMUTITEM.Name, mdiMain.mnuRPTTHMUTITEM.Caption & " (" & mdiMain.mnuWarehouseReport.Caption & ")"
        SetMenuMatrix RPTTHPOBUY_Row, mdiMain.mnuRPTTHPOBUY.Name, mdiMain.mnuRPTTHPOBUY.Caption & " (" & mdiMain.mnuBuyReport.Caption & ")"
        SetMenuMatrix RPTTHDOBUY_Row, mdiMain.mnuRPTTHDOBUY.Name, mdiMain.mnuRPTTHDOBUY.Caption & " (" & mdiMain.mnuBuyReport.Caption & ")"
        SetMenuMatrix RPTTHSJBUY_Row, mdiMain.mnuRPTTHSJBUY.Name, mdiMain.mnuRPTTHSJBUY.Caption & " (" & mdiMain.mnuBuyReport.Caption & ")"
        SetMenuMatrix RPTTHFKTBUY_Row, mdiMain.mnuRPTTHFKTBUY.Name, mdiMain.mnuRPTTHFKTBUY.Caption & " (" & mdiMain.mnuBuyReport.Caption & ")"
        SetMenuMatrix RPTTHRTRBUY_Row, mdiMain.mnuRPTTHRTRBUY.Name, mdiMain.mnuRPTTHRTRBUY.Caption & " (" & mdiMain.mnuBuyReport.Caption & ")"
        SetMenuMatrix RPTTHPOSELL_Row, mdiMain.mnuRPTTHPOSELL.Name, mdiMain.mnuRPTTHPOSELL.Caption & " (" & mdiMain.mnuSellReport.Caption & ")"
        SetMenuMatrix RPTTHSOSELL_Row, mdiMain.mnuRPTTHSOSELL.Name, mdiMain.mnuRPTTHSOSELL.Caption & " (" & mdiMain.mnuSellReport.Caption & ")"
        SetMenuMatrix RPTTHSJSELL_Row, mdiMain.mnuRPTTHSJSELL.Name, mdiMain.mnuRPTTHSJSELL.Caption & " (" & mdiMain.mnuSellReport.Caption & ")"
        SetMenuMatrix RPTTHFKTSELL_Row, mdiMain.mnuRPTTHFKTSELL.Name, mdiMain.mnuRPTTHFKTSELL.Caption & " (" & mdiMain.mnuSellReport.Caption & ")"
        SetMenuMatrix RPTTHRTRSELL_Row, mdiMain.mnuRPTTHRTRSELL.Name, mdiMain.mnuRPTTHRTRSELL.Caption & " (" & mdiMain.mnuSellReport.Caption & ")"
        SetMenuMatrix RPTSUMTHPOSELL_Row, mdiMain.mnuRPTSUMTHPOSELL.Name, mdiMain.mnuRPTSUMTHPOSELL.Caption & " (" & mdiMain.mnuExternalReport.Caption & " - " & mdiMain.mnuCustomerReport.Caption & ")"
        
        .TextMatrix(SEPARATOR4_Row, MenuRootColumn) = mdiMain.mnuSettings.Caption
        
        SetMenuMatrix TMREMINDERCUSTOMER_Row, mdiMain.mnuTMREMINDERCUSTOMER.Name, mdiMain.mnuTMREMINDERCUSTOMER.Caption
        SetMenuMatrix PwdChange_Row, mdiMain.mnuPwdChange.Name, mdiMain.mnuPwdChange.Caption
        SetMenuMatrix Backup_Row, mdiMain.mnuBackup.Name, mdiMain.mnuBackup.Caption
        SetMenuMatrix Toolbar_Row, mdiMain.mnuToolbar.Name, mdiMain.mnuToolbar.Caption
        SetMenuMatrix CurrencyToolbar_Row, mdiMain.mnuCurrencyToolbar.Name, mdiMain.mnuCurrencyToolbar.Caption
        SetMenuMatrix Menu_Row, mdiMain.mnuMenu.Name, mdiMain.mnuMenu.Caption
        SetMenuMatrix THRECYCLE_Row, mdlTable.CreateTHRECYCLE, mdlText.strTHRECYCLE
        
        If Not mdlGlobal.conFinance Is Nothing Then
            .Rows = .Rows + 1
        
            SetMenuMatrix SyncFinance_Row, mdiMain.mnuSyncFinance.Name, mdiMain.mnuSyncFinance.Caption
        End If
        
        If Not mdlGlobal.conAccounting Is Nothing Then
            .Rows = .Rows + 1
            
            SetMenuMatrix SyncAccounting_Row, mdiMain.mnuSyncAccounting.Name, mdiMain.mnuSyncAccounting.Caption
        End If
    End With
End Sub

Private Sub SetMenuMatrix(ByVal intRow As Integer, ByVal strId As String, ByVal strValue As String)
    With Me.flxMain
        .TextMatrix(intRow, MenuIdColumn) = strId
        .TextMatrix(intRow, MenuNameColumn) = strValue
    End With
End Sub

Private Sub SaveFunction()
    If Not CheckValidation Then Exit Sub
    
    Dim blnExists As Boolean
    
    blnExists = True
    
    Dim rstTemp As ADODB.Recordset
    
    Set rstTemp = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "*", mdlTable.CreateTMUSER, , "UserId='" & mdlProcedures.RepDupText(Me.txtUserId.Text) & "'")
    
    With rstTemp
        If Not .RecordCount > 0 Then
            .AddNew
            
            !UserId = mdlProcedures.RepDupText(Trim(Me.txtUserId.Text))
            
            !CreateId = mdlGlobal.UserAuthority.UserId
            !CreateDate = mdlProcedures.FormatDate(Now)
            
            blnExists = False
        End If
        
        !UserName = mdlProcedures.RepDupText(Trim(Me.txtUserName.Text))
        !UserType = mdlSecurity.EncryptText(mdlProcedures.GetComboData(Me.cmbUserType), !UserId)
        !UserPwd = mdlSecurity.EncryptText(Me.txtUserPwd.Text, !UserId)
        
        !UpdateId = mdlGlobal.UserAuthority.UserId
        !UpdateDate = mdlProcedures.FormatDate(Now)
        
        .Update
    End With
    
    If Me.cmbUserType.ListIndex = 0 Then
        mdlDatabase.DeleteRecordQuery mdlGlobal.conInventory, mdlTable.CreateTMMENUAUTHORITY, "UserId='" & mdlProcedures.RepDupText(Me.txtUserId.Text) & "'"
    Else
        SaveAuthority rstTemp!UserId
    End If
    
    If blnExists Then
        If mdlProcedures.SetMsgYesNo("Pemakai Diubah" & vbCrLf & "Ingin Hapus Layar ?", Me.Caption) Then
            ClearText True
            
            Me.txtUserId.SetFocus
        End If
    Else
        If mdlProcedures.SetMsgYesNo("Pemakai Bertambah" & vbCrLf & "Ingin Hapus Layar ?", Me.Caption) Then
            ClearText True
            
            Me.txtUserId.SetFocus
        End If
    End If
    
    mdlDatabase.CloseRecordset rstTemp
End Sub

Private Sub SaveAuthority(ByVal strUserId As String)
    With Me.flxMain
        Dim rstTemp As ADODB.Recordset
        
        Dim intCounter As Integer
        
        For intCounter = 0 To .Rows - 1
            If Not Trim(.TextMatrix(intCounter, MenuIdColumn)) = "" Then
                Set rstTemp = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "*", mdlTable.CreateTMMENU, , "MenuId='" & mdlProcedures.RepDupText(.TextMatrix(intCounter, MenuIdColumn)) & "'")
                
                If Not rstTemp.RecordCount > 0 Then
                    rstTemp.AddNew
                    
                    rstTemp!MenuId = Trim(.TextMatrix(intCounter, MenuIdColumn))
                    
                    rstTemp!CreateId = mdlGlobal.UserAuthority.UserId
                    rstTemp!CreateDate = mdlProcedures.FormatDate(Now)
                End If
                
                rstTemp!MenuName = Trim(.TextMatrix(intCounter, MenuNameColumn))
                
                rstTemp!UpdateId = mdlGlobal.UserAuthority.UserId
                rstTemp!UpdateDate = mdlProcedures.FormatDate(Now)
                
                rstTemp.Update
                
                Set rstTemp = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "*", mdlTable.CreateTMMENUAUTHORITY, , "UserId='" & strUserId & "' AND MenuId='" & mdlProcedures.RepDupText(.TextMatrix(intCounter, MenuIdColumn)) & "'")
                
                If Not rstTemp.RecordCount > 0 Then
                    rstTemp.AddNew
                    
                    rstTemp!UserId = strUserId
                    rstTemp!MenuId = Trim(.TextMatrix(intCounter, MenuIdColumn))
                    
                    rstTemp!AuthorityId = rstTemp!UserId & rstTemp!MenuId
                    
                    rstTemp!CreateId = mdlGlobal.UserAuthority.UserId
                    rstTemp!CreateDate = mdlProcedures.FormatDate(Now)
                End If
                
                If Trim(.TextMatrix(intCounter, AccessYNColumn)) = Chr(254) Then
                    rstTemp!AccessYN = mdlSecurity.EncryptText(mdlGlobal.strYes, rstTemp!UserId)
                Else
                    rstTemp!AccessYN = mdlSecurity.EncryptText(mdlGlobal.strNo, rstTemp!UserId)
                End If
                
                rstTemp!UpdateId = mdlGlobal.UserAuthority.UserId
                rstTemp!UpdateDate = mdlProcedures.FormatDate(Now)
                
                rstTemp.Update
            End If
        Next intCounter
        
        mdlDatabase.CloseRecordset rstTemp
    End With
End Sub

Private Function CheckValidation() As Boolean
    If Trim(Me.txtUserId.Text) = "" Then
        Me.txtUserId.SetFocus
        
        CheckValidation = False
        
        Exit Function
    ElseIf Trim(Me.txtUserName.Text) = "" Then
        Me.txtUserName.SetFocus
        
        CheckValidation = False
        
        Exit Function
    ElseIf Trim(Me.txtUserPwd.Text) = "" Then
        Me.txtUserPwd.SetFocus
        
        CheckValidation = False
        
        Exit Function
    ElseIf Not Len(Trim(Me.txtUserPwd.Text)) > 5 Then
        Me.txtUserPwd.SetFocus
        
        CheckValidation = False
        
        Exit Function
    End If
    
    If Not Me.cmbUserType.ListIndex > -1 Then
        Me.cmbUserType.SetFocus
        
        CheckValidation = False
        
        Exit Function
    End If
    
    CheckValidation = True
End Function

Private Sub ClearText(Optional ByVal blnPrimaryKey As Boolean = False)
    If blnPrimaryKey Then Me.txtUserId.Text = ""
    
    Me.txtUserName.Text = ""
    Me.cmbUserType.ListIndex = 0
    Me.txtUserPwd.Text = ""
    
    ClearMatrix
End Sub

Private Sub ClearMatrix()
    With Me.flxMain
        .TextMatrix(TMITEM_Row, AccessYNColumn) = ""
        .TextMatrix(TMPRICELIST_Row, AccessYNColumn) = ""
        .TextMatrix(TMITEMPRICE_Row, AccessYNColumn) = ""
        .TextMatrix(TMGROUP_Row, AccessYNColumn) = ""
        .TextMatrix(TMCATEGORY_Row, AccessYNColumn) = ""
        .TextMatrix(TMBRAND_Row, AccessYNColumn) = ""
        .TextMatrix(TMUNITY_Row, AccessYNColumn) = ""
        .TextMatrix(TMCUSTOMER_Row, AccessYNColumn) = ""
        .TextMatrix(TMVENDOR_Row, AccessYNColumn) = ""
        .TextMatrix(TMEMPLOYEE_Row, AccessYNColumn) = ""
        .TextMatrix(TMJOBTYPE_Row, AccessYNColumn) = ""
        .TextMatrix(TMDIVISION_Row, AccessYNColumn) = ""
        .TextMatrix(TMCURRENCY_Row, AccessYNColumn) = ""
        .TextMatrix(TMWAREHOUSE_Row, AccessYNColumn) = ""
        
        .TextMatrix(MISTMITEM_Row, AccessYNColumn) = ""
        .TextMatrix(MISTMPRICELIST_Row, AccessYNColumn) = ""
        .TextMatrix(MISTMITEMPRICE_Row, AccessYNColumn) = ""
        .TextMatrix(MISTMGROUP_Row, AccessYNColumn) = ""
        .TextMatrix(MISTMCATEGORY_Row, AccessYNColumn) = ""
        .TextMatrix(MISTMBRAND_Row, AccessYNColumn) = ""
        .TextMatrix(MISTMUNITY_Row, AccessYNColumn) = ""
        .TextMatrix(MISTMCUSTOMER_Row, AccessYNColumn) = ""
        .TextMatrix(MISTMCUSTOMERTRANS_Row, AccessYNColumn) = ""
        .TextMatrix(MISTMVENDOR_Row, AccessYNColumn) = ""
        .TextMatrix(MISTMEMPLOYEE_Row, AccessYNColumn) = ""
        .TextMatrix(MISTMJOBTYPE_Row, AccessYNColumn) = ""
        .TextMatrix(MISTMDIVISION_Row, AccessYNColumn) = ""
        .TextMatrix(MISTMCURRENCY_Row, AccessYNColumn) = ""
        .TextMatrix(MISTMWAREHOUSE_Row, AccessYNColumn) = ""
        .TextMatrix(MISTMSTOCKINIT_Row, AccessYNColumn) = ""
        .TextMatrix(MISTHSTOCK_Row, AccessYNColumn) = ""
        .TextMatrix(MISTHITEMIN_Row, AccessYNColumn) = ""
        .TextMatrix(MISTHITEMOUT_Row, AccessYNColumn) = ""
        .TextMatrix(MISTHMUTITEM_Row, AccessYNColumn) = ""
        .TextMatrix(MISTHPOBUY_Row, AccessYNColumn) = ""
        .TextMatrix(MISTHDOBUY_Row, AccessYNColumn) = ""
        .TextMatrix(MISTHSJBUY_Row, AccessYNColumn) = ""
        .TextMatrix(MISTHFKTBUY_Row, AccessYNColumn) = ""
        .TextMatrix(MISTHRTRBUY_Row, AccessYNColumn) = ""
        .TextMatrix(MISTHSALESSUM_Row, AccessYNColumn) = ""
        .TextMatrix(MISTHPOSELL_Row, AccessYNColumn) = ""
        .TextMatrix(MISTHSOSELL_Row, AccessYNColumn) = ""
        .TextMatrix(MISTHSJSELL_Row, AccessYNColumn) = ""
        .TextMatrix(MISTHFKTSELL_Row, AccessYNColumn) = ""
        .TextMatrix(MISTHRTRSELL_Row, AccessYNColumn) = ""
        
        .TextMatrix(TMSTOCKINIT_Row, AccessYNColumn) = ""
        .TextMatrix(THITEMIN_Row, AccessYNColumn) = ""
        .TextMatrix(THITEMOUT_Row, AccessYNColumn) = ""
        .TextMatrix(THMUTITEM_Row, AccessYNColumn) = ""
        .TextMatrix(THPOBUY_Row, AccessYNColumn) = ""
        .TextMatrix(THDOBUY_Row, AccessYNColumn) = ""
        .TextMatrix(THSJBUY_Row, AccessYNColumn) = ""
        .TextMatrix(THFKTBUY_Row, AccessYNColumn) = ""
        .TextMatrix(THRTRBUY_Row, AccessYNColumn) = ""
        .TextMatrix(THSALESSUM_Row, AccessYNColumn) = ""
        .TextMatrix(THPOSELL_Row, AccessYNColumn) = ""
        .TextMatrix(THSOSELL_Row, AccessYNColumn) = ""
        .TextMatrix(THSJSELL_Row, AccessYNColumn) = ""
        .TextMatrix(THFKTSELL_Row, AccessYNColumn) = ""
        .TextMatrix(THRTRSELL_Row, AccessYNColumn) = ""
        
        .TextMatrix(RPTTMSTOCKINIT_Row, AccessYNColumn) = ""
        .TextMatrix(RPTTHSTOCK_Row, AccessYNColumn) = ""
        .TextMatrix(RPTTHITEMIN_Row, AccessYNColumn) = ""
        .TextMatrix(RPTTHITEMOUT_Row, AccessYNColumn) = ""
        .TextMatrix(RPTTHMUTITEM_Row, AccessYNColumn) = ""
        .TextMatrix(RPTTHPOBUY_Row, AccessYNColumn) = ""
        .TextMatrix(RPTTHDOBUY_Row, AccessYNColumn) = ""
        .TextMatrix(RPTTHSJBUY_Row, AccessYNColumn) = ""
        .TextMatrix(RPTTHFKTBUY_Row, AccessYNColumn) = ""
        .TextMatrix(RPTTHRTRBUY_Row, AccessYNColumn) = ""
        .TextMatrix(RPTTHPOSELL_Row, AccessYNColumn) = ""
        .TextMatrix(RPTTHSOSELL_Row, AccessYNColumn) = ""
        .TextMatrix(RPTTHSJSELL_Row, AccessYNColumn) = ""
        .TextMatrix(RPTTHFKTSELL_Row, AccessYNColumn) = ""
        .TextMatrix(RPTTHRTRSELL_Row, AccessYNColumn) = ""
        .TextMatrix(RPTSUMTHPOSELL_Row, AccessYNColumn) = ""
        
        .TextMatrix(TMREMINDERCUSTOMER_Row, AccessYNColumn) = ""
        .TextMatrix(PwdChange_Row, AccessYNColumn) = ""
        .TextMatrix(Backup_Row, AccessYNColumn) = ""
        .TextMatrix(Toolbar_Row, AccessYNColumn) = ""
        .TextMatrix(CurrencyToolbar_Row, AccessYNColumn) = ""
        .TextMatrix(Menu_Row, AccessYNColumn) = ""
        .TextMatrix(THRECYCLE_Row, AccessYNColumn) = ""
        
        If Not mdlGlobal.conFinance Is Nothing Then
            .TextMatrix(SyncFinance_Row, AccessYNColumn) = ""
        End If
        
        If Not mdlGlobal.conAccounting Is Nothing Then
            .TextMatrix(SyncAccounting_Row, AccessYNColumn) = ""
        End If
    End With
End Sub

Private Sub FillText(ByVal strUserId As String)
    Dim rstTemp As ADODB.Recordset
    
    Set rstTemp = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "*", mdlTable.CreateTMUSER, False, "UserId='" & mdlProcedures.RepDupText(strUserId) & "'")
    
    With rstTemp
        If .RecordCount > 0 Then
            Me.txtUserName.Text = Trim(!UserName)
            
            If mdlGlobal.UserAuthority.IsAdmin(Trim(strUserId)) Then
                Me.cmbUserType.ListIndex = 0
                
                ClearMatrix
            Else
                Me.cmbUserType.ListIndex = 1
                
                FillMenuAuthority !UserId
            End If
            
            Me.txtUserPwd.Text = mdlSecurity.DecryptText(Trim(!UserPwd), !UserId)
        Else
            ClearText
        End If
    End With
    
    mdlDatabase.CloseRecordset rstTemp
End Sub

Private Sub FillMenuAuthority(ByVal strUserId As String)
    With Me.flxMain
        .TextMatrix(TMITEM_Row, AccessYNColumn) = _
            IIf(mdlGlobal.UserAuthority.IsMenuAccess(.TextMatrix(TMITEM_Row, MenuIdColumn), strUserId), Chr(254), "")
        .TextMatrix(TMPRICELIST_Row, AccessYNColumn) = _
            IIf(mdlGlobal.UserAuthority.IsMenuAccess(.TextMatrix(TMPRICELIST_Row, MenuIdColumn), strUserId), Chr(254), "")
        .TextMatrix(TMITEMPRICE_Row, AccessYNColumn) = _
            IIf(mdlGlobal.UserAuthority.IsMenuAccess(.TextMatrix(TMITEMPRICE_Row, MenuIdColumn), strUserId), Chr(254), "")
        .TextMatrix(TMGROUP_Row, AccessYNColumn) = _
            IIf(mdlGlobal.UserAuthority.IsMenuAccess(.TextMatrix(TMGROUP_Row, MenuIdColumn), strUserId), Chr(254), "")
        .TextMatrix(TMCATEGORY_Row, AccessYNColumn) = _
            IIf(mdlGlobal.UserAuthority.IsMenuAccess(.TextMatrix(TMCATEGORY_Row, MenuIdColumn), strUserId), Chr(254), "")
        .TextMatrix(TMBRAND_Row, AccessYNColumn) = _
            IIf(mdlGlobal.UserAuthority.IsMenuAccess(.TextMatrix(TMBRAND_Row, MenuIdColumn), strUserId), Chr(254), "")
        .TextMatrix(TMUNITY_Row, AccessYNColumn) = _
            IIf(mdlGlobal.UserAuthority.IsMenuAccess(.TextMatrix(TMUNITY_Row, MenuIdColumn), strUserId), Chr(254), "")
        .TextMatrix(TMCUSTOMER_Row, AccessYNColumn) = _
            IIf(mdlGlobal.UserAuthority.IsMenuAccess(.TextMatrix(TMCUSTOMER_Row, MenuIdColumn), strUserId), Chr(254), "")
        .TextMatrix(TMVENDOR_Row, AccessYNColumn) = _
            IIf(mdlGlobal.UserAuthority.IsMenuAccess(.TextMatrix(TMVENDOR_Row, MenuIdColumn), strUserId), Chr(254), "")
        .TextMatrix(TMEMPLOYEE_Row, AccessYNColumn) = _
            IIf(mdlGlobal.UserAuthority.IsMenuAccess(.TextMatrix(TMEMPLOYEE_Row, MenuIdColumn), strUserId), Chr(254), "")
        .TextMatrix(TMJOBTYPE_Row, AccessYNColumn) = _
            IIf(mdlGlobal.UserAuthority.IsMenuAccess(.TextMatrix(TMJOBTYPE_Row, MenuIdColumn), strUserId), Chr(254), "")
        .TextMatrix(TMDIVISION_Row, AccessYNColumn) = _
            IIf(mdlGlobal.UserAuthority.IsMenuAccess(.TextMatrix(TMDIVISION_Row, MenuIdColumn), strUserId), Chr(254), "")
        .TextMatrix(TMCURRENCY_Row, AccessYNColumn) = _
            IIf(mdlGlobal.UserAuthority.IsMenuAccess(.TextMatrix(TMWAREHOUSE_Row, MenuIdColumn), strUserId), Chr(254), "")
        .TextMatrix(TMWAREHOUSE_Row, AccessYNColumn) = _
            IIf(mdlGlobal.UserAuthority.IsMenuAccess(.TextMatrix(TMWAREHOUSE_Row, MenuIdColumn), strUserId), Chr(254), "")
        
        .TextMatrix(MISTMITEM_Row, AccessYNColumn) = _
            IIf(mdlGlobal.UserAuthority.IsMenuAccess(.TextMatrix(MISTMITEM_Row, MenuIdColumn), strUserId), Chr(254), "")
        .TextMatrix(MISTMPRICELIST_Row, AccessYNColumn) = _
            IIf(mdlGlobal.UserAuthority.IsMenuAccess(.TextMatrix(MISTMPRICELIST_Row, MenuIdColumn), strUserId), Chr(254), "")
        .TextMatrix(MISTMITEMPRICE_Row, AccessYNColumn) = _
            IIf(mdlGlobal.UserAuthority.IsMenuAccess(.TextMatrix(MISTMITEMPRICE_Row, MenuIdColumn), strUserId), Chr(254), "")
        .TextMatrix(MISTMGROUP_Row, AccessYNColumn) = _
            IIf(mdlGlobal.UserAuthority.IsMenuAccess(.TextMatrix(MISTMGROUP_Row, MenuIdColumn), strUserId), Chr(254), "")
        .TextMatrix(MISTMCATEGORY_Row, AccessYNColumn) = _
            IIf(mdlGlobal.UserAuthority.IsMenuAccess(.TextMatrix(MISTMCATEGORY_Row, MenuIdColumn), strUserId), Chr(254), "")
        .TextMatrix(MISTMBRAND_Row, AccessYNColumn) = _
            IIf(mdlGlobal.UserAuthority.IsMenuAccess(.TextMatrix(MISTMBRAND_Row, MenuIdColumn), strUserId), Chr(254), "")
        .TextMatrix(MISTMUNITY_Row, AccessYNColumn) = _
            IIf(mdlGlobal.UserAuthority.IsMenuAccess(.TextMatrix(MISTMUNITY_Row, MenuIdColumn), strUserId), Chr(254), "")
        .TextMatrix(MISTMCUSTOMER_Row, AccessYNColumn) = _
            IIf(mdlGlobal.UserAuthority.IsMenuAccess(.TextMatrix(MISTMCUSTOMER_Row, MenuIdColumn), strUserId), Chr(254), "")
        .TextMatrix(MISTMCUSTOMERTRANS_Row, AccessYNColumn) = _
            IIf(mdlGlobal.UserAuthority.IsMenuAccess(.TextMatrix(MISTMCUSTOMERTRANS_Row, MenuIdColumn), strUserId), Chr(254), "")
        .TextMatrix(MISTMVENDOR_Row, AccessYNColumn) = _
            IIf(mdlGlobal.UserAuthority.IsMenuAccess(.TextMatrix(MISTMVENDOR_Row, MenuIdColumn), strUserId), Chr(254), "")
        .TextMatrix(MISTMEMPLOYEE_Row, AccessYNColumn) = _
            IIf(mdlGlobal.UserAuthority.IsMenuAccess(.TextMatrix(MISTMEMPLOYEE_Row, MenuIdColumn), strUserId), Chr(254), "")
        .TextMatrix(MISTMJOBTYPE_Row, AccessYNColumn) = _
            IIf(mdlGlobal.UserAuthority.IsMenuAccess(.TextMatrix(MISTMJOBTYPE_Row, MenuIdColumn), strUserId), Chr(254), "")
        .TextMatrix(MISTMDIVISION_Row, AccessYNColumn) = _
            IIf(mdlGlobal.UserAuthority.IsMenuAccess(.TextMatrix(MISTMDIVISION_Row, MenuIdColumn), strUserId), Chr(254), "")
        .TextMatrix(MISTMCURRENCY_Row, AccessYNColumn) = _
            IIf(mdlGlobal.UserAuthority.IsMenuAccess(.TextMatrix(MISTMCURRENCY_Row, MenuIdColumn), strUserId), Chr(254), "")
        .TextMatrix(MISTMWAREHOUSE_Row, AccessYNColumn) = _
            IIf(mdlGlobal.UserAuthority.IsMenuAccess(.TextMatrix(MISTMWAREHOUSE_Row, MenuIdColumn), strUserId), Chr(254), "")
        .TextMatrix(MISTMSTOCKINIT_Row, AccessYNColumn) = _
            IIf(mdlGlobal.UserAuthority.IsMenuAccess(.TextMatrix(MISTMSTOCKINIT_Row, MenuIdColumn), strUserId), Chr(254), "")
        .TextMatrix(MISTHSTOCK_Row, AccessYNColumn) = _
            IIf(mdlGlobal.UserAuthority.IsMenuAccess(.TextMatrix(MISTHSTOCK_Row, MenuIdColumn), strUserId), Chr(254), "")
        .TextMatrix(MISTHITEMIN_Row, AccessYNColumn) = _
            IIf(mdlGlobal.UserAuthority.IsMenuAccess(.TextMatrix(MISTHITEMIN_Row, MenuIdColumn), strUserId), Chr(254), "")
        .TextMatrix(MISTHITEMOUT_Row, AccessYNColumn) = _
            IIf(mdlGlobal.UserAuthority.IsMenuAccess(.TextMatrix(MISTHITEMOUT_Row, MenuIdColumn), strUserId), Chr(254), "")
        .TextMatrix(MISTHMUTITEM_Row, AccessYNColumn) = _
            IIf(mdlGlobal.UserAuthority.IsMenuAccess(.TextMatrix(MISTHMUTITEM_Row, MenuIdColumn), strUserId), Chr(254), "")
        .TextMatrix(MISTHPOBUY_Row, AccessYNColumn) = _
            IIf(mdlGlobal.UserAuthority.IsMenuAccess(.TextMatrix(MISTHPOBUY_Row, MenuIdColumn), strUserId), Chr(254), "")
        .TextMatrix(MISTHDOBUY_Row, AccessYNColumn) = _
            IIf(mdlGlobal.UserAuthority.IsMenuAccess(.TextMatrix(MISTHDOBUY_Row, MenuIdColumn), strUserId), Chr(254), "")
        .TextMatrix(MISTHSJBUY_Row, AccessYNColumn) = _
            IIf(mdlGlobal.UserAuthority.IsMenuAccess(.TextMatrix(MISTHSJBUY_Row, MenuIdColumn), strUserId), Chr(254), "")
        .TextMatrix(MISTHFKTBUY_Row, AccessYNColumn) = _
            IIf(mdlGlobal.UserAuthority.IsMenuAccess(.TextMatrix(MISTHFKTBUY_Row, MenuIdColumn), strUserId), Chr(254), "")
        .TextMatrix(MISTHRTRBUY_Row, AccessYNColumn) = _
            IIf(mdlGlobal.UserAuthority.IsMenuAccess(.TextMatrix(MISTHRTRBUY_Row, MenuIdColumn), strUserId), Chr(254), "")
        .TextMatrix(MISTHSALESSUM_Row, AccessYNColumn) = _
            IIf(mdlGlobal.UserAuthority.IsMenuAccess(.TextMatrix(MISTHSALESSUM_Row, MenuIdColumn), strUserId), Chr(254), "")
        .TextMatrix(MISTHPOSELL_Row, AccessYNColumn) = _
            IIf(mdlGlobal.UserAuthority.IsMenuAccess(.TextMatrix(MISTHPOSELL_Row, MenuIdColumn), strUserId), Chr(254), "")
        .TextMatrix(MISTHSOSELL_Row, AccessYNColumn) = _
            IIf(mdlGlobal.UserAuthority.IsMenuAccess(.TextMatrix(MISTHSOSELL_Row, MenuIdColumn), strUserId), Chr(254), "")
        .TextMatrix(MISTHSJSELL_Row, AccessYNColumn) = _
            IIf(mdlGlobal.UserAuthority.IsMenuAccess(.TextMatrix(MISTHSJSELL_Row, MenuIdColumn), strUserId), Chr(254), "")
        .TextMatrix(MISTHFKTSELL_Row, AccessYNColumn) = _
            IIf(mdlGlobal.UserAuthority.IsMenuAccess(.TextMatrix(MISTHFKTSELL_Row, MenuIdColumn), strUserId), Chr(254), "")
        .TextMatrix(MISTHRTRSELL_Row, AccessYNColumn) = _
            IIf(mdlGlobal.UserAuthority.IsMenuAccess(.TextMatrix(MISTHRTRSELL_Row, MenuIdColumn), strUserId), Chr(254), "")
        
        .TextMatrix(TMSTOCKINIT_Row, AccessYNColumn) = _
            IIf(mdlGlobal.UserAuthority.IsMenuAccess(.TextMatrix(TMSTOCKINIT_Row, MenuIdColumn), strUserId), Chr(254), "")
        .TextMatrix(THITEMIN_Row, AccessYNColumn) = _
            IIf(mdlGlobal.UserAuthority.IsMenuAccess(.TextMatrix(THITEMIN_Row, MenuIdColumn), strUserId), Chr(254), "")
        .TextMatrix(THITEMOUT_Row, AccessYNColumn) = _
            IIf(mdlGlobal.UserAuthority.IsMenuAccess(.TextMatrix(THITEMOUT_Row, MenuIdColumn), strUserId), Chr(254), "")
        .TextMatrix(THMUTITEM_Row, AccessYNColumn) = _
            IIf(mdlGlobal.UserAuthority.IsMenuAccess(.TextMatrix(THMUTITEM_Row, MenuIdColumn), strUserId), Chr(254), "")
            
        .TextMatrix(THPOBUY_Row, AccessYNColumn) = _
            IIf(mdlGlobal.UserAuthority.IsMenuAccess(.TextMatrix(THPOBUY_Row, MenuIdColumn), strUserId), Chr(254), "")
        .TextMatrix(THDOBUY_Row, AccessYNColumn) = _
            IIf(mdlGlobal.UserAuthority.IsMenuAccess(.TextMatrix(THDOBUY_Row, MenuIdColumn), strUserId), Chr(254), "")
        .TextMatrix(THSJBUY_Row, AccessYNColumn) = _
            IIf(mdlGlobal.UserAuthority.IsMenuAccess(.TextMatrix(THSJBUY_Row, MenuIdColumn), strUserId), Chr(254), "")
        .TextMatrix(THFKTBUY_Row, AccessYNColumn) = _
            IIf(mdlGlobal.UserAuthority.IsMenuAccess(.TextMatrix(THFKTBUY_Row, MenuIdColumn), strUserId), Chr(254), "")
        .TextMatrix(THRTRBUY_Row, AccessYNColumn) = _
            IIf(mdlGlobal.UserAuthority.IsMenuAccess(.TextMatrix(THRTRBUY_Row, MenuIdColumn), strUserId), Chr(254), "")
        
        .TextMatrix(THSALESSUM_Row, AccessYNColumn) = _
            IIf(mdlGlobal.UserAuthority.IsMenuAccess(.TextMatrix(THSALESSUM_Row, MenuIdColumn), strUserId), Chr(254), "")
        .TextMatrix(THPOSELL_Row, AccessYNColumn) = _
            IIf(mdlGlobal.UserAuthority.IsMenuAccess(.TextMatrix(THPOSELL_Row, MenuIdColumn), strUserId), Chr(254), "")
        .TextMatrix(THSOSELL_Row, AccessYNColumn) = _
            IIf(mdlGlobal.UserAuthority.IsMenuAccess(.TextMatrix(THSOSELL_Row, MenuIdColumn), strUserId), Chr(254), "")
        .TextMatrix(THSJSELL_Row, AccessYNColumn) = _
            IIf(mdlGlobal.UserAuthority.IsMenuAccess(.TextMatrix(THSJSELL_Row, MenuIdColumn), strUserId), Chr(254), "")
        .TextMatrix(THFKTSELL_Row, AccessYNColumn) = _
            IIf(mdlGlobal.UserAuthority.IsMenuAccess(.TextMatrix(THFKTSELL_Row, MenuIdColumn), strUserId), Chr(254), "")
        .TextMatrix(THRTRSELL_Row, AccessYNColumn) = _
            IIf(mdlGlobal.UserAuthority.IsMenuAccess(.TextMatrix(THRTRSELL_Row, MenuIdColumn), strUserId), Chr(254), "")
        
        .TextMatrix(RPTTMSTOCKINIT_Row, AccessYNColumn) = _
            IIf(mdlGlobal.UserAuthority.IsMenuAccess(.TextMatrix(RPTTMSTOCKINIT_Row, MenuIdColumn), strUserId), Chr(254), "")
        .TextMatrix(RPTTHSTOCK_Row, AccessYNColumn) = _
            IIf(mdlGlobal.UserAuthority.IsMenuAccess(.TextMatrix(RPTTHSTOCK_Row, MenuIdColumn), strUserId), Chr(254), "")
        .TextMatrix(RPTTHITEMIN_Row, AccessYNColumn) = _
            IIf(mdlGlobal.UserAuthority.IsMenuAccess(.TextMatrix(RPTTHITEMIN_Row, MenuIdColumn), strUserId), Chr(254), "")
        .TextMatrix(RPTTHITEMOUT_Row, AccessYNColumn) = _
            IIf(mdlGlobal.UserAuthority.IsMenuAccess(.TextMatrix(RPTTHITEMOUT_Row, MenuIdColumn), strUserId), Chr(254), "")
        .TextMatrix(RPTTHMUTITEM_Row, AccessYNColumn) = _
            IIf(mdlGlobal.UserAuthority.IsMenuAccess(.TextMatrix(RPTTHMUTITEM_Row, MenuIdColumn), strUserId), Chr(254), "")
        .TextMatrix(RPTTHPOBUY_Row, AccessYNColumn) = _
            IIf(mdlGlobal.UserAuthority.IsMenuAccess(.TextMatrix(RPTTHPOBUY_Row, MenuIdColumn), strUserId), Chr(254), "")
        .TextMatrix(RPTTHDOBUY_Row, AccessYNColumn) = _
            IIf(mdlGlobal.UserAuthority.IsMenuAccess(.TextMatrix(RPTTHDOBUY_Row, MenuIdColumn), strUserId), Chr(254), "")
        .TextMatrix(RPTTHSJBUY_Row, AccessYNColumn) = _
            IIf(mdlGlobal.UserAuthority.IsMenuAccess(.TextMatrix(RPTTHSJBUY_Row, MenuIdColumn), strUserId), Chr(254), "")
        .TextMatrix(RPTTHFKTBUY_Row, AccessYNColumn) = _
            IIf(mdlGlobal.UserAuthority.IsMenuAccess(.TextMatrix(RPTTHFKTBUY_Row, MenuIdColumn), strUserId), Chr(254), "")
        .TextMatrix(RPTTHRTRBUY_Row, AccessYNColumn) = _
            IIf(mdlGlobal.UserAuthority.IsMenuAccess(.TextMatrix(RPTTHRTRBUY_Row, MenuIdColumn), strUserId), Chr(254), "")
        .TextMatrix(RPTTHPOSELL_Row, AccessYNColumn) = _
            IIf(mdlGlobal.UserAuthority.IsMenuAccess(.TextMatrix(RPTTHPOSELL_Row, MenuIdColumn), strUserId), Chr(254), "")
        .TextMatrix(RPTTHSOSELL_Row, AccessYNColumn) = _
            IIf(mdlGlobal.UserAuthority.IsMenuAccess(.TextMatrix(RPTTHSOSELL_Row, MenuIdColumn), strUserId), Chr(254), "")
        .TextMatrix(RPTTHSJSELL_Row, AccessYNColumn) = _
            IIf(mdlGlobal.UserAuthority.IsMenuAccess(.TextMatrix(RPTTHSJSELL_Row, MenuIdColumn), strUserId), Chr(254), "")
        .TextMatrix(RPTTHFKTSELL_Row, AccessYNColumn) = _
            IIf(mdlGlobal.UserAuthority.IsMenuAccess(.TextMatrix(RPTTHFKTSELL_Row, MenuIdColumn), strUserId), Chr(254), "")
        .TextMatrix(RPTTHRTRSELL_Row, AccessYNColumn) = _
            IIf(mdlGlobal.UserAuthority.IsMenuAccess(.TextMatrix(RPTTHRTRSELL_Row, MenuIdColumn), strUserId), Chr(254), "")
        .TextMatrix(RPTSUMTHPOSELL_Row, AccessYNColumn) = _
            IIf(mdlGlobal.UserAuthority.IsMenuAccess(.TextMatrix(RPTSUMTHPOSELL_Row, MenuIdColumn), strUserId), Chr(254), "")
        
        .TextMatrix(TMREMINDERCUSTOMER_Row, AccessYNColumn) = _
            IIf(mdlGlobal.UserAuthority.IsMenuAccess(.TextMatrix(TMREMINDERCUSTOMER_Row, MenuIdColumn), strUserId), Chr(254), "")
        .TextMatrix(PwdChange_Row, AccessYNColumn) = _
            IIf(mdlGlobal.UserAuthority.IsMenuAccess(.TextMatrix(PwdChange_Row, MenuIdColumn), strUserId), Chr(254), "")
        .TextMatrix(Backup_Row, AccessYNColumn) = _
            IIf(mdlGlobal.UserAuthority.IsMenuAccess(.TextMatrix(Backup_Row, MenuIdColumn), strUserId), Chr(254), "")
        .TextMatrix(Toolbar_Row, AccessYNColumn) = _
            IIf(mdlGlobal.UserAuthority.IsMenuAccess(.TextMatrix(Toolbar_Row, MenuIdColumn), strUserId), Chr(254), "")
        .TextMatrix(CurrencyToolbar_Row, AccessYNColumn) = _
            IIf(mdlGlobal.UserAuthority.IsMenuAccess(.TextMatrix(CurrencyToolbar_Row, MenuIdColumn), strUserId), Chr(254), "")
        .TextMatrix(Menu_Row, AccessYNColumn) = _
            IIf(mdlGlobal.UserAuthority.IsMenuAccess(.TextMatrix(Menu_Row, MenuIdColumn), strUserId), Chr(254), "")
        .TextMatrix(THRECYCLE_Row, AccessYNColumn) = _
            IIf(mdlGlobal.UserAuthority.IsMenuAccess(.TextMatrix(THRECYCLE_Row, MenuIdColumn), strUserId), Chr(254), "")
        
        If Not mdlGlobal.conFinance Is Nothing Then
            .TextMatrix(SyncFinance_Row, AccessYNColumn) = _
                IIf(mdlGlobal.UserAuthority.IsMenuAccess(.TextMatrix(SyncFinance_Row, MenuIdColumn), strUserId), Chr(254), "")
        End If
        
        If Not mdlGlobal.conAccounting Is Nothing Then
            .TextMatrix(SyncAccounting_Row, AccessYNColumn) = _
                IIf(mdlGlobal.UserAuthority.IsMenuAccess(.TextMatrix(SyncAccounting_Row, MenuIdColumn), strUserId), Chr(254), "")
        End If
    End With
End Sub

Public Property Get Parent() As Boolean
    Parent = blnParent
End Property

Public Property Get UserId() As String
    UserId = Trim(Me.txtUserId.Text)
End Property

Public Property Get UserName() As String
    UserName = Trim(Me.txtUserName.Text)
End Property

Public Property Let Parent(ByVal blnEnable As Boolean)
    blnParent = blnEnable
    
    If Not blnEnable Then mdlProcedures.CenterWindows Me
End Property

Public Property Let UserId(ByVal strUserId As String)
    Me.txtUserId.Text = strUserId
End Property
