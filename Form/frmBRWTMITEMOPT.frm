VERSION 5.00
Begin VB.Form frmBRWTMITEMOPT 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3390
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   9015
   Icon            =   "frmBRWTMITEMOPT.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3390
   ScaleWidth      =   9015
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Batal"
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
      Left            =   6360
      TabIndex        =   6
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton cmdOptional 
      Caption         =   "Optional"
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
      Left            =   7680
      TabIndex        =   5
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Frame fraSearch 
      Height          =   2655
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   8775
      Begin VB.ComboBox cmbUnityId 
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
         TabIndex        =   4
         Top             =   2160
         Width           =   7215
      End
      Begin VB.ComboBox cmbGroupId 
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
         TabIndex        =   1
         Top             =   720
         Width           =   7215
      End
      Begin VB.ComboBox cmbBrandId 
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
         TabIndex        =   3
         Top             =   1680
         Width           =   7215
      End
      Begin VB.ComboBox cmbCategoryId 
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
         TabIndex        =   2
         Top             =   1200
         Width           =   7215
      End
      Begin VB.ComboBox cmbVendorId 
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
         TabIndex        =   0
         Top             =   240
         Width           =   7455
      End
      Begin VB.Label lblUnityId 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Height          =   240
         Left            =   120
         TabIndex        =   11
         Top             =   2160
         Width           =   615
      End
      Begin VB.Label lblBrandId 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Height          =   240
         Left            =   120
         TabIndex        =   10
         Top             =   1680
         Width           =   435
      End
      Begin VB.Label lblVendorId 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Height          =   240
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   825
      End
      Begin VB.Label lblCategoryId 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Height          =   240
         Left            =   120
         TabIndex        =   9
         Top             =   1200
         Width           =   450
      End
      Begin VB.Label lblGroupId 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Height          =   240
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   420
      End
   End
End
Attribute VB_Name = "frmBRWTMITEMOPT"
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
    If frmBRWTMITEM.Parent Then
        frmBRWTMITEM.Parent = False
    End If
    
    If frmBRWTMPRICELIST.Parent Then
        frmBRWTMPRICELIST.Parent = False
    End If
    
    If frmBRWTMITEMPRICE.Parent Then
        frmBRWTMITEMPRICE.Parent = False
    End If
    
    If frmBRWTMSTOCKINIT.Parent Then
        frmBRWTMSTOCKINIT.Parent = False
    End If
    
    If frmMISTMITEM.Parent Then
        frmMISTMITEM.Parent = False
    End If
    
    If frmMISTMPRICELIST.Parent Then
        frmMISTMPRICELIST.Parent = False
    End If
    
    If frmMISTMITEMPRICE.Parent Then
        frmMISTMITEMPRICE.Parent = False
    End If
    
    If frmMISTMSTOCKINIT.Parent Then
        frmMISTMSTOCKINIT.Parent = False
    End If
    
    If frmMISTHSTOCK.Parent Then
        frmMISTHSTOCK.Parent = False
    End If
    
    If frmRPTTHSTOCK.Parent Then
        frmRPTTHSTOCK.Parent = False
    End If
    
    If frmRPTTMSTOCKINIT.Parent Then
        frmRPTTMSTOCKINIT.Parent = False
    End If
    
    If frmTDPOBUY.Parent Then
        frmTDPOBUY.Parent = False
    End If
    
    If frmTDDOBUY.Parent Then
        frmTDDOBUY.Parent = False
    End If
    
    If frmTDSJBUY.Parent Then
        frmTDSJBUY.Parent = False
    End If
    
    If frmTDRTRBUY.Parent Then
        frmTDRTRBUY.Parent = False
    End If
    
    If frmTDPOSELL.Parent Then
        frmTDPOSELL.Parent = False
    End If
    
    If frmTDSOSELL.Parent Then
        frmTDSOSELL.Parent = False
    End If
    
    If frmTDSJSELL.Parent Then
        frmTDSJSELL.Parent = False
    End If
    
    If frmTDRTRSELL.Parent Then
        frmTDRTRSELL.Parent = False
    End If
    
    If frmTDMUTITEM.Parent Then
        frmTDMUTITEM.Parent = False
    End If
    
    If frmTDITEMIN.Parent Then
        frmTDITEMIN.Parent = False
    End If
    
    If frmTDITEMOUT.Parent Then
        frmTDITEMOUT.Parent = False
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmBRWTMITEMOPT = Nothing
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOptional_Click()
    If frmBRWTMITEM.Parent Then
        frmBRWTMITEM.VendorOptional = Me.cmbVendorId.Text
        frmBRWTMITEM.GroupOptional = Me.cmbGroupId.Text
        frmBRWTMITEM.CategoryOptional = Me.cmbCategoryId.Text
        frmBRWTMITEM.BrandOptional = Me.cmbBrandId.Text
        frmBRWTMITEM.UnityOptional = Me.cmbUnityId.Text
    ElseIf frmBRWTMPRICELIST.Parent Then
        frmBRWTMPRICELIST.VendorOptional = Me.cmbVendorId.Text
        frmBRWTMPRICELIST.GroupOptional = Me.cmbGroupId.Text
        frmBRWTMPRICELIST.CategoryOptional = Me.cmbCategoryId.Text
        frmBRWTMPRICELIST.BrandOptional = Me.cmbBrandId.Text
        frmBRWTMPRICELIST.UnityOptional = Me.cmbUnityId.Text
    ElseIf frmBRWTMITEMPRICE.Parent Then
        frmBRWTMITEMPRICE.VendorOptional = Me.cmbVendorId.Text
        frmBRWTMITEMPRICE.GroupOptional = Me.cmbGroupId.Text
        frmBRWTMITEMPRICE.CategoryOptional = Me.cmbCategoryId.Text
        frmBRWTMITEMPRICE.BrandOptional = Me.cmbBrandId.Text
        frmBRWTMITEMPRICE.UnityOptional = Me.cmbUnityId.Text
    ElseIf frmBRWTMSTOCKINIT.Parent Then
        frmBRWTMSTOCKINIT.VendorOptional = Me.cmbVendorId.Text
        frmBRWTMSTOCKINIT.GroupOptional = Me.cmbGroupId.Text
        frmBRWTMSTOCKINIT.CategoryOptional = Me.cmbCategoryId.Text
        frmBRWTMSTOCKINIT.BrandOptional = Me.cmbBrandId.Text
        frmBRWTMSTOCKINIT.UnityOptional = Me.cmbUnityId.Text
    ElseIf frmMISTHSTOCK.Parent Then
        frmMISTHSTOCK.VendorOptional = Me.cmbVendorId.Text
        frmMISTHSTOCK.GroupOptional = Me.cmbGroupId.Text
        frmMISTHSTOCK.CategoryOptional = Me.cmbCategoryId.Text
        frmMISTHSTOCK.BrandOptional = Me.cmbBrandId.Text
        frmMISTHSTOCK.UnityOptional = Me.cmbUnityId.Text
    ElseIf frmMISTMITEM.Parent Then
        frmMISTMITEM.VendorOptional = Me.cmbVendorId.Text
        frmMISTMITEM.GroupOptional = Me.cmbGroupId.Text
        frmMISTMITEM.CategoryOptional = Me.cmbCategoryId.Text
        frmMISTMITEM.BrandOptional = Me.cmbBrandId.Text
        frmMISTMITEM.UnityOptional = Me.cmbUnityId.Text
    ElseIf frmMISTMSTOCKINIT.Parent Then
        frmMISTMSTOCKINIT.VendorOptional = Me.cmbVendorId.Text
        frmMISTMSTOCKINIT.GroupOptional = Me.cmbGroupId.Text
        frmMISTMSTOCKINIT.CategoryOptional = Me.cmbCategoryId.Text
        frmMISTMSTOCKINIT.BrandOptional = Me.cmbBrandId.Text
        frmMISTMSTOCKINIT.UnityOptional = Me.cmbUnityId.Text
    ElseIf frmMISTMPRICELIST.Parent Then
        frmMISTMPRICELIST.VendorOptional = Me.cmbVendorId.Text
        frmMISTMPRICELIST.GroupOptional = Me.cmbGroupId.Text
        frmMISTMPRICELIST.CategoryOptional = Me.cmbCategoryId.Text
        frmMISTMPRICELIST.BrandOptional = Me.cmbBrandId.Text
        frmMISTMPRICELIST.UnityOptional = Me.cmbUnityId.Text
    ElseIf frmMISTMITEMPRICE.Parent Then
        frmMISTMITEMPRICE.VendorOptional = Me.cmbVendorId.Text
        frmMISTMITEMPRICE.GroupOptional = Me.cmbGroupId.Text
        frmMISTMITEMPRICE.CategoryOptional = Me.cmbCategoryId.Text
        frmMISTMITEMPRICE.BrandOptional = Me.cmbBrandId.Text
        frmMISTMITEMPRICE.UnityOptional = Me.cmbUnityId.Text
    ElseIf frmRPTTHSTOCK.Parent Then
        frmRPTTHSTOCK.VendorOptional = Me.cmbVendorId.Text
        frmRPTTHSTOCK.GroupOptional = Me.cmbGroupId.Text
        frmRPTTHSTOCK.CategoryOptional = Me.cmbCategoryId.Text
        frmRPTTHSTOCK.BrandOptional = Me.cmbBrandId.Text
        frmRPTTHSTOCK.UnityOptional = Me.cmbUnityId.Text
    ElseIf frmRPTTMSTOCKINIT.Parent Then
        frmRPTTMSTOCKINIT.VendorOptional = Me.cmbVendorId.Text
        frmRPTTMSTOCKINIT.GroupOptional = Me.cmbGroupId.Text
        frmRPTTMSTOCKINIT.CategoryOptional = Me.cmbCategoryId.Text
        frmRPTTMSTOCKINIT.BrandOptional = Me.cmbBrandId.Text
        frmRPTTMSTOCKINIT.UnityOptional = Me.cmbUnityId.Text
    ElseIf frmTDPOBUY.Parent Then
        frmTDPOBUY.VendorOptional = Me.cmbVendorId.Text
        frmTDPOBUY.GroupOptional = Me.cmbGroupId.Text
        frmTDPOBUY.CategoryOptional = Me.cmbCategoryId.Text
        frmTDPOBUY.BrandOptional = Me.cmbBrandId.Text
        frmTDPOBUY.UnityOptional = Me.cmbUnityId.Text
    ElseIf frmTDDOBUY.Parent Then
        frmTDDOBUY.VendorOptional = Me.cmbVendorId.Text
        frmTDDOBUY.GroupOptional = Me.cmbGroupId.Text
        frmTDDOBUY.CategoryOptional = Me.cmbCategoryId.Text
        frmTDDOBUY.BrandOptional = Me.cmbBrandId.Text
        frmTDDOBUY.UnityOptional = Me.cmbUnityId.Text
    ElseIf frmTDSJBUY.Parent Then
        frmTDSJBUY.VendorOptional = Me.cmbVendorId.Text
        frmTDSJBUY.GroupOptional = Me.cmbGroupId.Text
        frmTDSJBUY.CategoryOptional = Me.cmbCategoryId.Text
        frmTDSJBUY.BrandOptional = Me.cmbBrandId.Text
        frmTDSJBUY.UnityOptional = Me.cmbUnityId.Text
    ElseIf frmTDRTRBUY.Parent Then
        frmTDRTRBUY.VendorOptional = Me.cmbVendorId.Text
        frmTDRTRBUY.GroupOptional = Me.cmbGroupId.Text
        frmTDRTRBUY.CategoryOptional = Me.cmbCategoryId.Text
        frmTDRTRBUY.BrandOptional = Me.cmbBrandId.Text
        frmTDRTRBUY.UnityOptional = Me.cmbUnityId.Text
    ElseIf frmTDPOSELL.Parent Then
        frmTDPOSELL.VendorOptional = Me.cmbVendorId.Text
        frmTDPOSELL.GroupOptional = Me.cmbGroupId.Text
        frmTDPOSELL.CategoryOptional = Me.cmbCategoryId.Text
        frmTDPOSELL.BrandOptional = Me.cmbBrandId.Text
        frmTDPOSELL.UnityOptional = Me.cmbUnityId.Text
    ElseIf frmTDSOSELL.Parent Then
        frmTDSOSELL.VendorOptional = Me.cmbVendorId.Text
        frmTDSOSELL.GroupOptional = Me.cmbGroupId.Text
        frmTDSOSELL.CategoryOptional = Me.cmbCategoryId.Text
        frmTDSOSELL.BrandOptional = Me.cmbBrandId.Text
        frmTDSOSELL.UnityOptional = Me.cmbUnityId.Text
    ElseIf frmTDSJSELL.Parent Then
        frmTDSJSELL.VendorOptional = Me.cmbVendorId.Text
        frmTDSJSELL.GroupOptional = Me.cmbGroupId.Text
        frmTDSJSELL.CategoryOptional = Me.cmbCategoryId.Text
        frmTDSJSELL.BrandOptional = Me.cmbBrandId.Text
        frmTDSJSELL.UnityOptional = Me.cmbUnityId.Text
    ElseIf frmTDRTRSELL.Parent Then
        frmTDRTRSELL.VendorOptional = Me.cmbVendorId.Text
        frmTDRTRSELL.GroupOptional = Me.cmbGroupId.Text
        frmTDRTRSELL.CategoryOptional = Me.cmbCategoryId.Text
        frmTDRTRSELL.BrandOptional = Me.cmbBrandId.Text
        frmTDRTRSELL.UnityOptional = Me.cmbUnityId.Text
    ElseIf frmTDMUTITEM.Parent Then
        frmTDMUTITEM.VendorOptional = Me.cmbVendorId.Text
        frmTDMUTITEM.GroupOptional = Me.cmbGroupId.Text
        frmTDMUTITEM.CategoryOptional = Me.cmbCategoryId.Text
        frmTDMUTITEM.BrandOptional = Me.cmbBrandId.Text
        frmTDMUTITEM.UnityOptional = Me.cmbUnityId.Text
    ElseIf frmTDITEMIN.Parent Then
        frmTDITEMIN.VendorOptional = Me.cmbVendorId.Text
        frmTDITEMIN.GroupOptional = Me.cmbGroupId.Text
        frmTDITEMIN.CategoryOptional = Me.cmbCategoryId.Text
        frmTDITEMIN.BrandOptional = Me.cmbBrandId.Text
        frmTDITEMIN.UnityOptional = Me.cmbUnityId.Text
    ElseIf frmTDITEMOUT.Parent Then
        frmTDITEMOUT.VendorOptional = Me.cmbVendorId.Text
        frmTDITEMOUT.GroupOptional = Me.cmbGroupId.Text
        frmTDITEMOUT.CategoryOptional = Me.cmbCategoryId.Text
        frmTDITEMOUT.BrandOptional = Me.cmbBrandId.Text
        frmTDITEMOUT.UnityOptional = Me.cmbUnityId.Text
    End If
    
    Unload Me
End Sub

Private Sub SetInitialization()
    mdlProcedures.CenterWindows Me, False
    
    Me.Caption = mdlText.strBRWTMITEMOPT
    
    FillCombo
    
    If frmBRWTMITEM.Parent Then
        mdlProcedures.SetComboData Me.cmbVendorId, mdlProcedures.SplitData(frmBRWTMITEM.VendorOptional)
        mdlProcedures.SetComboData Me.cmbGroupId, mdlProcedures.SplitData(frmBRWTMITEM.GroupOptional)
        mdlProcedures.SetComboData Me.cmbCategoryId, mdlProcedures.SplitData(frmBRWTMITEM.CategoryOptional)
        mdlProcedures.SetComboData Me.cmbBrandId, mdlProcedures.SplitData(frmBRWTMITEM.BrandOptional)
        mdlProcedures.SetComboData Me.cmbUnityId, mdlProcedures.SplitData(frmBRWTMITEM.UnityOptional)
    ElseIf frmBRWTMPRICELIST.Parent Then
        mdlProcedures.SetComboData Me.cmbVendorId, mdlProcedures.SplitData(frmBRWTMPRICELIST.VendorOptional)
        mdlProcedures.SetComboData Me.cmbGroupId, mdlProcedures.SplitData(frmBRWTMPRICELIST.GroupOptional)
        mdlProcedures.SetComboData Me.cmbCategoryId, mdlProcedures.SplitData(frmBRWTMPRICELIST.CategoryOptional)
        mdlProcedures.SetComboData Me.cmbBrandId, mdlProcedures.SplitData(frmBRWTMPRICELIST.BrandOptional)
        mdlProcedures.SetComboData Me.cmbUnityId, mdlProcedures.SplitData(frmBRWTMPRICELIST.UnityOptional)
    ElseIf frmBRWTMITEMPRICE.Parent Then
        mdlProcedures.SetComboData Me.cmbVendorId, mdlProcedures.SplitData(frmBRWTMITEMPRICE.VendorOptional)
        mdlProcedures.SetComboData Me.cmbGroupId, mdlProcedures.SplitData(frmBRWTMITEMPRICE.GroupOptional)
        mdlProcedures.SetComboData Me.cmbCategoryId, mdlProcedures.SplitData(frmBRWTMITEMPRICE.CategoryOptional)
        mdlProcedures.SetComboData Me.cmbBrandId, mdlProcedures.SplitData(frmBRWTMITEMPRICE.BrandOptional)
        mdlProcedures.SetComboData Me.cmbUnityId, mdlProcedures.SplitData(frmBRWTMITEMPRICE.UnityOptional)
    ElseIf frmBRWTMSTOCKINIT.Parent Then
        mdlProcedures.SetComboData Me.cmbVendorId, mdlProcedures.SplitData(frmBRWTMSTOCKINIT.VendorOptional)
        mdlProcedures.SetComboData Me.cmbGroupId, mdlProcedures.SplitData(frmBRWTMSTOCKINIT.GroupOptional)
        mdlProcedures.SetComboData Me.cmbCategoryId, mdlProcedures.SplitData(frmBRWTMSTOCKINIT.CategoryOptional)
        mdlProcedures.SetComboData Me.cmbBrandId, mdlProcedures.SplitData(frmBRWTMSTOCKINIT.BrandOptional)
        mdlProcedures.SetComboData Me.cmbUnityId, mdlProcedures.SplitData(frmBRWTMSTOCKINIT.UnityOptional)
    ElseIf frmMISTHSTOCK.Parent Then
        mdlProcedures.SetComboData Me.cmbVendorId, mdlProcedures.SplitData(frmMISTHSTOCK.VendorOptional)
        mdlProcedures.SetComboData Me.cmbGroupId, mdlProcedures.SplitData(frmMISTHSTOCK.GroupOptional)
        mdlProcedures.SetComboData Me.cmbCategoryId, mdlProcedures.SplitData(frmMISTHSTOCK.CategoryOptional)
        mdlProcedures.SetComboData Me.cmbBrandId, mdlProcedures.SplitData(frmMISTHSTOCK.BrandOptional)
        mdlProcedures.SetComboData Me.cmbUnityId, mdlProcedures.SplitData(frmMISTHSTOCK.UnityOptional)
    ElseIf frmMISTMITEM.Parent Then
        mdlProcedures.SetComboData Me.cmbVendorId, mdlProcedures.SplitData(frmMISTMITEM.VendorOptional)
        mdlProcedures.SetComboData Me.cmbGroupId, mdlProcedures.SplitData(frmMISTMITEM.GroupOptional)
        mdlProcedures.SetComboData Me.cmbCategoryId, mdlProcedures.SplitData(frmMISTMITEM.CategoryOptional)
        mdlProcedures.SetComboData Me.cmbBrandId, mdlProcedures.SplitData(frmMISTMITEM.BrandOptional)
        mdlProcedures.SetComboData Me.cmbUnityId, mdlProcedures.SplitData(frmMISTMITEM.UnityOptional)
    ElseIf frmMISTMPRICELIST.Parent Then
        mdlProcedures.SetComboData Me.cmbVendorId, mdlProcedures.SplitData(frmMISTMPRICELIST.VendorOptional)
        mdlProcedures.SetComboData Me.cmbGroupId, mdlProcedures.SplitData(frmMISTMPRICELIST.GroupOptional)
        mdlProcedures.SetComboData Me.cmbCategoryId, mdlProcedures.SplitData(frmMISTMPRICELIST.CategoryOptional)
        mdlProcedures.SetComboData Me.cmbBrandId, mdlProcedures.SplitData(frmMISTMPRICELIST.BrandOptional)
        mdlProcedures.SetComboData Me.cmbUnityId, mdlProcedures.SplitData(frmMISTMPRICELIST.UnityOptional)
    ElseIf frmMISTMITEMPRICE.Parent Then
        mdlProcedures.SetComboData Me.cmbVendorId, mdlProcedures.SplitData(frmMISTMITEMPRICE.VendorOptional)
        mdlProcedures.SetComboData Me.cmbGroupId, mdlProcedures.SplitData(frmMISTMITEMPRICE.GroupOptional)
        mdlProcedures.SetComboData Me.cmbCategoryId, mdlProcedures.SplitData(frmMISTMITEMPRICE.CategoryOptional)
        mdlProcedures.SetComboData Me.cmbBrandId, mdlProcedures.SplitData(frmMISTMITEMPRICE.BrandOptional)
        mdlProcedures.SetComboData Me.cmbUnityId, mdlProcedures.SplitData(frmMISTMITEMPRICE.UnityOptional)
    ElseIf frmMISTMSTOCKINIT.Parent Then
        mdlProcedures.SetComboData Me.cmbVendorId, mdlProcedures.SplitData(frmMISTMSTOCKINIT.VendorOptional)
        mdlProcedures.SetComboData Me.cmbGroupId, mdlProcedures.SplitData(frmMISTMSTOCKINIT.GroupOptional)
        mdlProcedures.SetComboData Me.cmbCategoryId, mdlProcedures.SplitData(frmMISTMSTOCKINIT.CategoryOptional)
        mdlProcedures.SetComboData Me.cmbBrandId, mdlProcedures.SplitData(frmMISTMSTOCKINIT.BrandOptional)
        mdlProcedures.SetComboData Me.cmbUnityId, mdlProcedures.SplitData(frmMISTMSTOCKINIT.UnityOptional)
    ElseIf frmRPTTHSTOCK.Parent Then
        mdlProcedures.SetComboData Me.cmbVendorId, mdlProcedures.SplitData(frmRPTTHSTOCK.VendorOptional)
        mdlProcedures.SetComboData Me.cmbGroupId, mdlProcedures.SplitData(frmRPTTHSTOCK.GroupOptional)
        mdlProcedures.SetComboData Me.cmbCategoryId, mdlProcedures.SplitData(frmRPTTHSTOCK.CategoryOptional)
        mdlProcedures.SetComboData Me.cmbBrandId, mdlProcedures.SplitData(frmRPTTHSTOCK.BrandOptional)
        mdlProcedures.SetComboData Me.cmbUnityId, mdlProcedures.SplitData(frmRPTTHSTOCK.UnityOptional)
    ElseIf frmRPTTMSTOCKINIT.Parent Then
        mdlProcedures.SetComboData Me.cmbVendorId, mdlProcedures.SplitData(frmRPTTMSTOCKINIT.VendorOptional)
        mdlProcedures.SetComboData Me.cmbGroupId, mdlProcedures.SplitData(frmRPTTMSTOCKINIT.GroupOptional)
        mdlProcedures.SetComboData Me.cmbCategoryId, mdlProcedures.SplitData(frmRPTTMSTOCKINIT.CategoryOptional)
        mdlProcedures.SetComboData Me.cmbBrandId, mdlProcedures.SplitData(frmRPTTMSTOCKINIT.BrandOptional)
        mdlProcedures.SetComboData Me.cmbUnityId, mdlProcedures.SplitData(frmRPTTMSTOCKINIT.UnityOptional)
    ElseIf frmTDPOBUY.Parent Then
        mdlProcedures.SetComboData Me.cmbVendorId, mdlProcedures.SplitData(frmTDPOBUY.VendorOptional)
        mdlProcedures.SetComboData Me.cmbGroupId, mdlProcedures.SplitData(frmTDPOBUY.GroupOptional)
        mdlProcedures.SetComboData Me.cmbCategoryId, mdlProcedures.SplitData(frmTDPOBUY.CategoryOptional)
        mdlProcedures.SetComboData Me.cmbBrandId, mdlProcedures.SplitData(frmTDPOBUY.BrandOptional)
        mdlProcedures.SetComboData Me.cmbUnityId, mdlProcedures.SplitData(frmTDPOBUY.UnityOptional)
    ElseIf frmTDDOBUY.Parent Then
        mdlProcedures.SetComboData Me.cmbVendorId, mdlProcedures.SplitData(frmTDDOBUY.VendorOptional)
        mdlProcedures.SetComboData Me.cmbGroupId, mdlProcedures.SplitData(frmTDDOBUY.GroupOptional)
        mdlProcedures.SetComboData Me.cmbCategoryId, mdlProcedures.SplitData(frmTDDOBUY.CategoryOptional)
        mdlProcedures.SetComboData Me.cmbBrandId, mdlProcedures.SplitData(frmTDDOBUY.BrandOptional)
        mdlProcedures.SetComboData Me.cmbUnityId, mdlProcedures.SplitData(frmTDDOBUY.UnityOptional)
    ElseIf frmTDSJBUY.Parent Then
        mdlProcedures.SetComboData Me.cmbVendorId, mdlProcedures.SplitData(frmTDSJBUY.VendorOptional)
        mdlProcedures.SetComboData Me.cmbGroupId, mdlProcedures.SplitData(frmTDSJBUY.GroupOptional)
        mdlProcedures.SetComboData Me.cmbCategoryId, mdlProcedures.SplitData(frmTDSJBUY.CategoryOptional)
        mdlProcedures.SetComboData Me.cmbBrandId, mdlProcedures.SplitData(frmTDSJBUY.BrandOptional)
        mdlProcedures.SetComboData Me.cmbUnityId, mdlProcedures.SplitData(frmTDSJBUY.UnityOptional)
    ElseIf frmTDRTRBUY.Parent Then
        mdlProcedures.SetComboData Me.cmbVendorId, mdlProcedures.SplitData(frmTDRTRBUY.VendorOptional)
        mdlProcedures.SetComboData Me.cmbGroupId, mdlProcedures.SplitData(frmTDRTRBUY.GroupOptional)
        mdlProcedures.SetComboData Me.cmbCategoryId, mdlProcedures.SplitData(frmTDRTRBUY.CategoryOptional)
        mdlProcedures.SetComboData Me.cmbBrandId, mdlProcedures.SplitData(frmTDRTRBUY.BrandOptional)
        mdlProcedures.SetComboData Me.cmbUnityId, mdlProcedures.SplitData(frmTDRTRBUY.UnityOptional)
    ElseIf frmTDPOSELL.Parent Then
        mdlProcedures.SetComboData Me.cmbVendorId, mdlProcedures.SplitData(frmTDPOSELL.VendorOptional)
        mdlProcedures.SetComboData Me.cmbGroupId, mdlProcedures.SplitData(frmTDPOSELL.GroupOptional)
        mdlProcedures.SetComboData Me.cmbCategoryId, mdlProcedures.SplitData(frmTDPOSELL.CategoryOptional)
        mdlProcedures.SetComboData Me.cmbBrandId, mdlProcedures.SplitData(frmTDPOSELL.BrandOptional)
        mdlProcedures.SetComboData Me.cmbUnityId, mdlProcedures.SplitData(frmTDPOSELL.UnityOptional)
    ElseIf frmTDSOSELL.Parent Then
        mdlProcedures.SetComboData Me.cmbVendorId, mdlProcedures.SplitData(frmTDSOSELL.VendorOptional)
        mdlProcedures.SetComboData Me.cmbGroupId, mdlProcedures.SplitData(frmTDSOSELL.GroupOptional)
        mdlProcedures.SetComboData Me.cmbCategoryId, mdlProcedures.SplitData(frmTDSOSELL.CategoryOptional)
        mdlProcedures.SetComboData Me.cmbBrandId, mdlProcedures.SplitData(frmTDSOSELL.BrandOptional)
        mdlProcedures.SetComboData Me.cmbUnityId, mdlProcedures.SplitData(frmTDSOSELL.UnityOptional)
    ElseIf frmTDSJSELL.Parent Then
        mdlProcedures.SetComboData Me.cmbVendorId, mdlProcedures.SplitData(frmTDSJSELL.VendorOptional)
        mdlProcedures.SetComboData Me.cmbGroupId, mdlProcedures.SplitData(frmTDSJSELL.GroupOptional)
        mdlProcedures.SetComboData Me.cmbCategoryId, mdlProcedures.SplitData(frmTDSJSELL.CategoryOptional)
        mdlProcedures.SetComboData Me.cmbBrandId, mdlProcedures.SplitData(frmTDSJSELL.BrandOptional)
        mdlProcedures.SetComboData Me.cmbUnityId, mdlProcedures.SplitData(frmTDSJSELL.UnityOptional)
    ElseIf frmTDRTRSELL.Parent Then
        mdlProcedures.SetComboData Me.cmbVendorId, mdlProcedures.SplitData(frmTDRTRSELL.VendorOptional)
        mdlProcedures.SetComboData Me.cmbGroupId, mdlProcedures.SplitData(frmTDRTRSELL.GroupOptional)
        mdlProcedures.SetComboData Me.cmbCategoryId, mdlProcedures.SplitData(frmTDRTRSELL.CategoryOptional)
        mdlProcedures.SetComboData Me.cmbBrandId, mdlProcedures.SplitData(frmTDRTRSELL.BrandOptional)
        mdlProcedures.SetComboData Me.cmbUnityId, mdlProcedures.SplitData(frmTDRTRSELL.UnityOptional)
    ElseIf frmTDMUTITEM.Parent Then
        mdlProcedures.SetComboData Me.cmbVendorId, mdlProcedures.SplitData(frmTDMUTITEM.VendorOptional)
        mdlProcedures.SetComboData Me.cmbGroupId, mdlProcedures.SplitData(frmTDMUTITEM.GroupOptional)
        mdlProcedures.SetComboData Me.cmbCategoryId, mdlProcedures.SplitData(frmTDMUTITEM.CategoryOptional)
        mdlProcedures.SetComboData Me.cmbBrandId, mdlProcedures.SplitData(frmTDMUTITEM.BrandOptional)
        mdlProcedures.SetComboData Me.cmbUnityId, mdlProcedures.SplitData(frmTDMUTITEM.UnityOptional)
    ElseIf frmTDITEMIN.Parent Then
        mdlProcedures.SetComboData Me.cmbVendorId, mdlProcedures.SplitData(frmTDITEMIN.VendorOptional)
        mdlProcedures.SetComboData Me.cmbGroupId, mdlProcedures.SplitData(frmTDITEMIN.GroupOptional)
        mdlProcedures.SetComboData Me.cmbCategoryId, mdlProcedures.SplitData(frmTDITEMIN.CategoryOptional)
        mdlProcedures.SetComboData Me.cmbBrandId, mdlProcedures.SplitData(frmTDITEMIN.BrandOptional)
        mdlProcedures.SetComboData Me.cmbUnityId, mdlProcedures.SplitData(frmTDITEMIN.UnityOptional)
    ElseIf frmTDITEMOUT.Parent Then
        mdlProcedures.SetComboData Me.cmbVendorId, mdlProcedures.SplitData(frmTDITEMOUT.VendorOptional)
        mdlProcedures.SetComboData Me.cmbGroupId, mdlProcedures.SplitData(frmTDITEMOUT.GroupOptional)
        mdlProcedures.SetComboData Me.cmbCategoryId, mdlProcedures.SplitData(frmTDITEMOUT.CategoryOptional)
        mdlProcedures.SetComboData Me.cmbBrandId, mdlProcedures.SplitData(frmTDITEMOUT.BrandOptional)
        mdlProcedures.SetComboData Me.cmbUnityId, mdlProcedures.SplitData(frmTDITEMOUT.UnityOptional)
    End If
End Sub

Private Sub FillCombo()
    Me.FillComboTMVENDOR
    Me.FillComboTMGROUP
    Me.FillComboTMCATEGORY
    Me.FillComboTMBRAND
    Me.FillComboTMUNITY
End Sub

Public Sub FillComboTMVENDOR()
    Dim rstTemp As ADODB.Recordset
    
    Set rstTemp = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "VendorId, Name", mdlTable.CreateTMVENDOR, False)
    
    mdlProcedures.FillComboData Me.cmbVendorId, rstTemp
    
    mdlDatabase.CloseRecordset rstTemp
End Sub

Public Sub FillComboTMGROUP()
    Dim rstTemp As ADODB.Recordset
    
    Set rstTemp = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "GroupId, Name", mdlTable.CreateTMGROUP, False)
    
    mdlProcedures.FillComboData Me.cmbGroupId, rstTemp
    
    mdlDatabase.CloseRecordset rstTemp
End Sub

Public Sub FillComboTMCATEGORY()
    Dim rstTemp As ADODB.Recordset
    
    Set rstTemp = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "CategoryId, Name", mdlTable.CreateTMCATEGORY, False)
    
    mdlProcedures.FillComboData Me.cmbCategoryId, rstTemp
    
    mdlDatabase.CloseRecordset rstTemp
End Sub

Public Sub FillComboTMBRAND()
    Dim rstTemp As ADODB.Recordset
    
    Set rstTemp = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "BrandId, Name", mdlTable.CreateTMBRAND, False)
    
    mdlProcedures.FillComboData Me.cmbBrandId, rstTemp
    
    mdlDatabase.CloseRecordset rstTemp
End Sub

Public Sub FillComboTMUNITY()
    Dim rstTemp As ADODB.Recordset
    
    Set rstTemp = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "UnityId, Name", mdlTable.CreateTMUNITY, False)
    
    mdlProcedures.FillComboData Me.cmbUnityId, rstTemp
    
    mdlDatabase.CloseRecordset rstTemp
End Sub
