VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmTMCONVERTPRICE 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   8205
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8775
   Icon            =   "frmTMCONVERTPRICE.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8205
   ScaleWidth      =   8775
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdChoose 
      Caption         =   "Pilih"
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
      Left            =   7440
      TabIndex        =   28
      Top             =   7680
      Width           =   1215
   End
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
      Left            =   6120
      TabIndex        =   29
      Top             =   7680
      Width           =   1215
   End
   Begin VB.Frame fraDetail 
      Height          =   975
      Left            =   120
      TabIndex        =   31
      Top             =   3840
      Width           =   8535
      Begin VB.OptionButton optType 
         Caption         =   "Non Heavy"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   4800
         TabIndex        =   4
         Top             =   600
         Width           =   1455
      End
      Begin VB.OptionButton optType 
         Caption         =   "Heavy"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   4800
         TabIndex        =   3
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox txtFreight 
         Alignment       =   1  'Right Justify
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
         Left            =   1320
         MaxLength       =   10
         TabIndex        =   5
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox txtNonHeavy 
         Alignment       =   1  'Right Justify
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
         Left            =   7080
         MaxLength       =   10
         TabIndex        =   7
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox txtHeavy 
         Alignment       =   1  'Right Justify
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
         Left            =   7080
         MaxLength       =   10
         TabIndex        =   6
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox txtWeight 
         Alignment       =   1  'Right Justify
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
         Left            =   1320
         MaxLength       =   10
         TabIndex        =   2
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lblFreight 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Freight"
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
         TabIndex        =   24
         Top             =   600
         Width           =   600
      End
      Begin VB.Label lblNonHeavy 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MF"
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
         Left            =   6600
         TabIndex        =   23
         Top             =   600
         Width           =   285
      End
      Begin VB.Label lblHeavy 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MF"
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
         Left            =   6600
         TabIndex        =   22
         Top             =   240
         Width           =   285
      End
      Begin VB.Label lblType 
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
         Left            =   3600
         TabIndex        =   21
         Top             =   240
         Width           =   450
      End
      Begin VB.Label lblWeight 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Berat"
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
         TabIndex        =   20
         Top             =   240
         Width           =   465
      End
   End
   Begin VB.Frame fraHeader 
      Height          =   2655
      Left            =   120
      TabIndex        =   30
      Top             =   1200
      Width           =   8535
      Begin VB.ComboBox cmbPriceListId 
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
         Left            =   1320
         TabIndex        =   0
         Top             =   1320
         Width           =   3735
      End
      Begin MSComCtl2.DTPicker dtpConvertDate 
         Height          =   375
         Left            =   1320
         TabIndex        =   1
         Top             =   2160
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   78446595
         CurrentDate     =   39331
      End
      Begin VB.Label lblConvertDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal"
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
         TabIndex        =   19
         Top             =   2160
         Width           =   675
      End
      Begin VB.Label txtCurrencyId 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   1320
         TabIndex        =   18
         Top             =   1800
         Width           =   7095
      End
      Begin VB.Label lblCurrency 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Height          =   240
         Left            =   120
         TabIndex        =   17
         Top             =   1800
         Width           =   945
      End
      Begin VB.Label txtPriceListValue 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   6960
         TabIndex        =   16
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label lblPriceListValue 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Harga"
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
         Left            =   5880
         TabIndex        =   15
         Top             =   1320
         Width           =   510
      End
      Begin VB.Label lblPriceListId 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Height          =   240
         Left            =   120
         TabIndex        =   14
         Top             =   1320
         Width           =   1080
      End
      Begin VB.Label txtItemId 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1320
         TabIndex        =   9
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lblPartNumber 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nomor Part"
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
         Top             =   600
         Width           =   990
      End
      Begin VB.Label txtPartNumber 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1320
         TabIndex        =   11
         Top             =   600
         Width           =   4935
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama"
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
         TabIndex        =   12
         Top             =   960
         Width           =   510
      End
      Begin VB.Label txtName 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   1320
         TabIndex        =   13
         Top             =   960
         Width           =   6060
      End
      Begin VB.Label lblItemId 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kode Barang"
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
         Width           =   1125
      End
   End
   Begin MSComctlLib.Toolbar tlbMain 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   32
      Top             =   0
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   1111
      ButtonWidth     =   609
      ButtonHeight    =   953
      Appearance      =   1
      _Version        =   393216
      Begin MSComctlLib.ImageList imlMain 
         Left            =   8040
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   48
         ImageHeight     =   48
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   5
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTMCONVERTPRICE.frx":1F8A
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTMCONVERTPRICE.frx":43DC
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTMCONVERTPRICE.frx":682E
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTMCONVERTPRICE.frx":8C80
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTMCONVERTPRICE.frx":B0D2
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin MSDataGridLib.DataGrid dgdMain 
      Height          =   2655
      Left            =   120
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   4920
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   4683
      _Version        =   393216
      HeadLines       =   2
      RowHeight       =   20
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Label lblTotal 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total"
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
      TabIndex        =   25
      Top             =   7680
      Width           =   420
   End
   Begin VB.Label txtTotal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   720
      TabIndex        =   26
      Top             =   7680
      Width           =   2100
   End
End
Attribute VB_Name = "frmTMCONVERTPRICE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum ButtonMode
    [AddButton] = 1
    [UpdateButton]
    [DeleteButton]
    [SaveButton]
    [CancelButton]
End Enum

Private rstMain As ADODB.Recordset

Private objMode As FunctionMode

Private strFormCaption As String

Private Sub Form_Load()
    SetInitialization
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If frmTMITEMPRICE.Parent Then
        frmTMITEMPRICE.Parent = False
    End If
    
    mdlDatabase.CloseRecordset rstMain
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmTMCONVERTPRICE = Nothing
End Sub

Private Sub tlbMain_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case AddButton:
            objMode = AddMode
            
            SetMode
        Case UpdateButton:
            objMode = UpdateMode
            
            SetMode
        Case DeleteButton:
            DeleteFunction
        Case SaveButton:
            SaveFunction
        Case CancelButton:
            objMode = ViewMode
            
            SetMode
            
            FillText
    End Select
End Sub

Private Sub optType_Click(Index As Integer)
    Select Case Index
        Case 0:
            Me.txtHeavy.Locked = False
            Me.txtNonHeavy.Locked = True
        Case 1:
            Me.txtHeavy.Locked = True
            Me.txtNonHeavy.Locked = False
    End Select
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdChoose_Click()
    If Not rstMain.RecordCount > 0 Then Exit Sub

    If frmTMITEMPRICE.Parent Then
        frmTMITEMPRICE.ItemPriceText = Me.txtTotal.Caption
        frmTMITEMPRICE.CurrencyIdCombo = mdlProcedures.SplitData(Me.txtCurrencyId.Caption)
    End If
    
    Unload Me
End Sub

Private Sub cmbPriceListId_Click()
    If objMode = ViewMode Then Exit Sub
    
    If Not mdlProcedures.IsValidComboData(Me.cmbPriceListId) Then
        Me.txtPriceListValue.Caption = ""
        Me.txtCurrencyId.Caption = ""
    Else
        Dim strPriceListId As String
        Dim strCurrencyId As String
        
        strPriceListId = Me.PriceListIdCombo
        strCurrencyId = mdlDatabase.GetFieldData(mdlGlobal.conInventory, "CurrencyId", mdlTable.CreateTMPRICELIST, "PriceListId='" & strPriceListId & "'")
        
        Me.txtPriceListValue.Caption = mdlProcedures.FormatCurrency(mdlDatabase.GetFieldData(mdlGlobal.conInventory, "PriceListValue", mdlTable.CreateTMPRICELIST, "PriceListId='" & strPriceListId & "'"), "#,##0.00")
        
        If Trim(strCurrencyId) = "" Then
            Me.txtCurrencyId.Caption = strCurrencyId
        Else
            Me.txtCurrencyId.Caption = strCurrencyId & " | " & mdlDatabase.GetFieldData(mdlGlobal.conInventory, "Name", mdlTable.CreateTMCURRENCY, "CurrencyId='" & strCurrencyId & "'")
        End If
    End If
End Sub

Private Sub txtWeight_GotFocus()
    mdlProcedures.GotFocus Me.txtWeight
End Sub

Private Sub txtHeavy_GotFocus()
    mdlProcedures.GotFocus Me.txtHeavy
End Sub

Private Sub txtNonHeavy_GotFocus()
    mdlProcedures.GotFocus Me.txtNonHeavy
End Sub

Private Sub txtFreight_GotFocus()
    mdlProcedures.GotFocus Me.txtFreight
End Sub

Private Sub txtWeight_Validate(Cancel As Boolean)
    Me.txtWeight.Text = mdlProcedures.FormatCurrency(Me.txtWeight.Text, "#,##0.00")
End Sub

Private Sub txtHeavy_Validate(Cancel As Boolean)
    Me.txtHeavy.Text = mdlProcedures.FormatCurrency(Me.txtHeavy.Text, "#,##0.00")
End Sub

Private Sub txtNonHeavy_Validate(Cancel As Boolean)
    Me.txtNonHeavy.Text = mdlProcedures.FormatCurrency(Me.txtNonHeavy.Text, "#,##0.00")
End Sub

Private Sub txtFreight_Validate(Cancel As Boolean)
    Me.txtFreight.Text = mdlProcedures.FormatCurrency(Me.txtFreight.Text, "#,##0.00")
End Sub

Private Sub cmbPriceListId_Validate(Cancel As Boolean)
    If Not mdlProcedures.IsValidComboData(Me.cmbPriceListId) Then
        Me.txtPriceListValue.Caption = ""
        Me.txtCurrencyId.Caption = ""
    Else
        Dim strPriceListId As String
        Dim strCurrencyId As String
        
        strPriceListId = Me.PriceListIdCombo
        strCurrencyId = mdlDatabase.GetFieldData(mdlGlobal.conInventory, "CurrencyId", mdlTable.CreateTMPRICELIST, "PriceListId='" & strPriceListId & "'")
        
        Me.txtPriceListValue.Caption = mdlProcedures.FormatCurrency(mdlDatabase.GetFieldData(mdlGlobal.conInventory, "PriceListValue", mdlTable.CreateTMPRICELIST, "PriceListId='" & strPriceListId & "'"), "#,##0.00")
        Me.txtCurrencyId.Caption = strCurrencyId & " | " & mdlDatabase.GetFieldData(mdlGlobal.conInventory, "Name", mdlTable.CreateTMCURRENCY, "CurrencyId='" & strCurrencyId & "'")
    End If
End Sub

Private Sub dgdMain_HeadClick(ByVal ColIndex As Integer)
    If Not objMode = ViewMode Then Exit Sub
    
    rstMain.Sort = rstMain.Fields(ColIndex)
    
    If rstMain.RecordCount > 0 Then
        FillText
    End If
End Sub

Private Sub dgdMain_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If Not objMode = ViewMode Then Exit Sub
    
    If rstMain.RecordCount > 0 Then
        FillText
    End If
End Sub

Private Sub SetInitialization()
    mdlProcedures.CenterWindows Me, False
    
    With Me.tlbMain
        .AllowCustomize = False
        
        .ImageList = Me.imlMain
        
        .Buttons.Add AddButton, , "Tambah", , AddButton
        .Buttons.Add UpdateButton, , "Ubah", , UpdateButton
        .Buttons.Add DeleteButton, , "Hapus", , DeleteButton
        .Buttons.Add SaveButton, , "Simpan", , SaveButton
        .Buttons.Add CancelButton, , "Batal", , CancelButton
    End With
    
    Me.dtpConvertDate.CustomFormat = mdlGlobal.strFormatDate
    
    Dim strCriteria As String
    
    strCriteria = ""
    
    If frmTMITEMPRICE.Parent Then
        Me.txtItemId.Caption = frmTMITEMPRICE.ItemIdText
        Me.txtPartNumber.Caption = frmTMITEMPRICE.PartNumberText
        Me.txtName.Caption = frmTMITEMPRICE.ItemNameText
        
        strCriteria = "ItemId='" & frmTMITEMPRICE.ItemIdText & "'"
    End If
    
    Me.txtHeavy.Locked = True
    Me.txtNonHeavy.Locked = True
    
    FillCombo
    
    strFormCaption = mdlText.strTMCONVERTPRICE
    
    Set rstMain = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "*", mdlTable.CreateTMCONVERTPRICE, , strCriteria, "ConvertDate ASC")
    
    ArrangeGrid
    
    objMode = ViewMode
    
    SetMode
    
    FillText
End Sub

Private Sub ArrangeGrid()
    Set Me.dgdMain.DataSource = rstMain
    
    With Me.dgdMain
        .Columns(1).Width = 2000
        .Columns(1).Locked = True
        .Columns(1).Caption = "Tanggal"
        .Columns(1).NumberFormat = "dd MMMM yyyy"
        .Columns(4).Width = 1800
        .Columns(4).Locked = True
        .Columns(4).Caption = "Berat"
        .Columns(4).NumberFormat = "#,##0.00"
        .Columns(6).Width = 1800
        .Columns(6).Locked = True
        .Columns(6).Caption = "MF"
        .Columns(6).NumberFormat = "#,##0.00"
        .Columns(7).Width = 1800
        .Columns(7).Locked = True
        .Columns(7).Caption = "Freight"
        .Columns(7).NumberFormat = "#,##0.00"
        
        .Columns(0).Width = 0
        .Columns(0).Visible = False
        .Columns(0).Locked = True
        .Columns(5).Width = 0
        .Columns(5).Visible = False
        .Columns(5).Locked = True
        
        Dim intCounter As Integer
        
        For intCounter = 2 To 3
            .Columns(intCounter).Width = 0
            .Columns(intCounter).Visible = False
            .Columns(intCounter).Locked = True
        Next intCounter
        
        For intCounter = 8 To .Columns.Count - 1
            .Columns(intCounter).Width = 0
            .Columns(intCounter).Visible = False
            .Columns(intCounter).Locked = True
        Next intCounter
    End With
End Sub

Private Sub SetMode()
    Dim intCounter As Integer
    
    Dim blnFront As Boolean
    Dim blnBack As Boolean
    
    If objMode = ViewMode Then
        blnFront = True
        blnBack = False
    Else
        blnFront = False
        blnBack = True
    End If
    
    Me.dgdMain.Enabled = blnFront
    
    With Me.tlbMain
        .Buttons(AddButton).Visible = blnFront
        .Buttons(UpdateButton).Visible = blnFront
        .Buttons(DeleteButton).Visible = blnFront
        .Buttons(SaveButton).Visible = blnBack
        .Buttons(CancelButton).Visible = blnBack
    End With
    
    Me.fraDetail.Enabled = blnBack
    
    If objMode = AddMode Then
        Me.fraHeader.Enabled = True
    
        mdlProcedures.SetControlMode Me, objMode
        
        Me.txtPriceListValue.Caption = ""
        Me.txtCurrencyId.Caption = ""
        Me.txtTotal.Caption = ""
        Me.txtHeavy.Text = mdlProcedures.FormatCurrency(CStr(mdlGlobal.curHeavy), "#,##0.00")
        Me.txtNonHeavy.Text = mdlProcedures.FormatCurrency(CStr(mdlGlobal.curNonHeavy), "#,##0.00")
        
        For intCounter = 0 To Me.optType.Count - 1
            Me.optType(intCounter).Value = False
        Next intCounter
        
        Me.txtHeavy.Locked = True
        Me.txtNonHeavy.Locked = True
        
        Me.cmbPriceListId.SetFocus
    ElseIf objMode = UpdateMode Then
        Me.fraHeader.Enabled = False
        
        mdlProcedures.SetControlMode Me, objMode, False
        
        Me.txtWeight.SetFocus
    Else
        Me.fraHeader.Enabled = False
    End If
    
    If rstMain.RecordCount > 0 Then
        With Me.tlbMain
            .Buttons(UpdateButton).Enabled = True
            .Buttons(DeleteButton).Enabled = True
        End With
        
        If objMode = ViewMode Then
            mdlProcedures.SetControlMode Me, objMode, False
        End If
    Else
        With Me.tlbMain
            .Buttons(UpdateButton).Enabled = False
            .Buttons(DeleteButton).Enabled = False
        End With
        
        If objMode = ViewMode Then
            mdlProcedures.SetControlMode Me, objMode
            
            Me.txtPriceListValue.Caption = ""
            Me.txtCurrencyId.Caption = ""
            Me.txtTotal.Caption = ""
            Me.txtHeavy.Text = ""
            Me.txtNonHeavy.Text = ""
            
            For intCounter = 0 To Me.optType.Count - 1
                Me.optType(intCounter).Value = False
            Next intCounter
            
            Me.txtHeavy.Locked = True
            Me.txtNonHeavy.Locked = True
        End If
    End If
    
    Me.Caption = mdlProcedures.SetCaptionMode(strFormCaption, objMode)
End Sub

Private Sub FillCombo()
    Me.FillComboTMPRICELIST
End Sub

Public Sub FillComboTMPRICELIST()
    Dim strItemId As String
    
    strItemId = ""
    
    If frmTMITEMPRICE.Parent Then
        strItemId = frmTMITEMPRICE.ItemIdText
    End If
    
    Dim rstTemp As ADODB.Recordset
    
    Set rstTemp = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "PriceListId, PriceListDate", mdlTable.CreateTMPRICELIST, False, "ItemId='" & strItemId & "'", "PriceListDate DESC")
    
    mdlProcedures.FillComboData Me.cmbPriceListId, rstTemp, " | ", True
    
    mdlDatabase.CloseRecordset rstTemp
End Sub

Private Sub DeleteFunction()
    objMode = DeleteMode
    
    If mdlProcedures.SetMsgYesNo("Apakah Anda Yakin ?", mdlProcedures.SetCaptionMode(strFormCaption, objMode)) Then
        mdlDatabase.DeleteSingleRecord rstMain
    End If
    
    objMode = ViewMode
    
    SetMode
    
    FillText
End Sub

Private Sub SaveFunction()
    If Not CheckValidation Then Exit Sub
    
    With rstMain
        Dim strItemId As String
        Dim strPriceListId As String
        
        strItemId = ""
        strPriceListId = ""
        
        If frmTMITEMPRICE.Parent Then
            strItemId = frmTMITEMPRICE.ItemIdText
        End If
        
        strItemId = _
            strItemId & _
            Space(mdlDatabase.GetColumnSize(mdlGlobal.conInventory, "ItemId", mdlTable.CreateTMCONVERTPRICE) - Len(strItemId))
        
        strPriceListId = Me.PriceListIdCombo
        
        mdlDatabase.SearchRecordset rstMain, "ConvertId", strItemId & strPriceListId & mdlProcedures.FormatDate(Me.dtpConvertDate.Value, "ddMMyyyy")
        
        If .EOF Then
            .AddNew
            
            !ConvertId = strItemId & strPriceListId & mdlProcedures.FormatDate(Me.dtpConvertDate.Value, "ddMMyyyy")
            !ItemId = Trim(strItemId)
            !PriceListId = Trim(strPriceListId)
            !ConvertDate = mdlProcedures.FormatDate(Me.dtpConvertDate.Value)
            
            !CreateId = Trim(mdlGlobal.UserAuthority.UserId)
            !CreateDate = mdlProcedures.FormatDate(Now)
        End If
        
        !Weight = mdlProcedures.GetCurrency(Me.txtWeight.Text)
        
        If Me.optType(0).Value Then
            !TypemfValue = mdlProcedures.GetCurrency(Me.txtHeavy.Text)
            !Typemf = mdlGlobal.strHeavy
        ElseIf Me.optType(1).Value Then
            !TypemfValue = mdlProcedures.GetCurrency(Me.txtNonHeavy.Text)
            !Typemf = mdlGlobal.strNonHeavy
        End If
        
        !Freight = mdlProcedures.GetCurrency(Me.txtFreight.Text)
        
        !UpdateId = Trim(mdlGlobal.UserAuthority.UserId)
        !UpdateDate = mdlProcedures.FormatDate(Now)
        
        .Update
    End With
    
    If objMode = AddMode Then
        mdlProcedures.SetControlMode Me, objMode
        
        Me.txtPriceListValue.Caption = ""
        Me.txtCurrencyId.Caption = ""
        Me.txtHeavy.Text = mdlProcedures.FormatCurrency(CStr(mdlGlobal.curHeavy), "#,##0.00")
        Me.txtNonHeavy.Text = mdlProcedures.FormatCurrency(CStr(mdlGlobal.curNonHeavy), "#,##0.00")
        
        Dim intCounter As Integer
        
        For intCounter = 0 To Me.optType.Count - 1
            Me.optType(intCounter).Value = False
        Next intCounter
        
        Me.txtHeavy.Locked = True
        Me.txtNonHeavy.Locked = True
        
        Me.dtpConvertDate.SetFocus
    ElseIf objMode = UpdateMode Then
        objMode = ViewMode
        
        SetMode
    End If
End Sub

Private Function CheckValidation() As Boolean
    If Not mdlProcedures.IsValidComboData(Me.cmbPriceListId) Then
        MsgBox "Daftar Harga Harap Diisi", vbOKOnly + vbExclamation, mdlProcedures.SetCaptionMode(strFormCaption, objMode)
        
        Me.cmbPriceListId.SetFocus
        
        CheckValidation = False
        
        Exit Function
    End If
    
    If objMode = AddMode Then
        With rstMain
            Dim strItemId As String
            Dim strPriceListId As String
            
            strItemId = ""
            strPriceListId = ""
            
            If frmTMITEMPRICE.Parent Then
                strItemId = frmTMITEMPRICE.ItemIdText
            End If
            
            strItemId = _
                strItemId & _
                Space(mdlDatabase.GetColumnSize(mdlGlobal.conInventory, "ItemId", mdlTable.CreateTMCONVERTPRICE) - Len(strItemId))
            
            strPriceListId = Me.PriceListIdCombo
            
            mdlDatabase.SearchRecordset rstMain, "ConvertId", strItemId & strPriceListId & mdlProcedures.FormatDate(Me.dtpConvertDate.Value, "ddMMyyyy")
            
            If Not .EOF Then
                MsgBox "Konversi Harga Sudah Ada", vbOKOnly + vbExclamation, mdlProcedures.SetCaptionMode(strFormCaption, objMode)
                
                Me.dtpConvertDate.SetFocus
                
                CheckValidation = False
                
                Exit Function
            End If
        End With
    End If
    
    CheckValidation = True
End Function

Private Sub FillText()
    With rstMain
        If .RecordCount > 0 Then
            Me.PriceListIdCombo = !PriceListId
                        
            Me.dtpConvertDate.Value = mdlProcedures.FormatDate(!ConvertDate, mdlGlobal.strFormatDate)
            
            Me.txtWeight.Text = mdlProcedures.FormatCurrency(!Weight, "#,##0.00")
            
            If Trim(!Typemf) = mdlGlobal.strHeavy Then
                Me.optType(0).Value = True
                
                Me.txtHeavy.Text = mdlProcedures.FormatCurrency(!TypemfValue, "#,##0.00")
                Me.txtNonHeavy.Text = mdlProcedures.FormatCurrency(CStr(curNonHeavy), "#,##0.00")
            Else
                Me.optType(1).Value = False
                
                Me.txtHeavy.Text = mdlProcedures.FormatCurrency(CStr(curHeavy), "#,##0.00")
                Me.txtNonHeavy.Text = mdlProcedures.FormatCurrency(!TypemfValue, "#,##0.00")
            End If
            
            Me.txtFreight.Text = mdlProcedures.FormatCurrency(!Freight, "#,##0.00")
            
            Dim curPriceListValue As Currency
            Dim curMF As Currency
            Dim curWeight As Currency
            Dim curFreight As Currency
            Dim curTotal As Currency
            
            curPriceListValue = mdlProcedures.GetCurrency(mdlDatabase.GetFieldData(mdlGlobal.conInventory, "PriceListValue", mdlTable.CreateTMPRICELIST, "PriceListId='" & mdlProcedures.GetComboData(Me.cmbPriceListId) & "'"))
            curMF = mdlProcedures.GetCurrency("0")
            curWeight = mdlProcedures.GetCurrency(Me.txtWeight.Text)
            curFreight = mdlProcedures.GetCurrency(Me.txtFreight.Text)
            
            If Me.optType(0).Value Then
                curMF = mdlProcedures.GetCurrency(Me.txtHeavy.Text)
            ElseIf Me.optType(1).Value Then
                curMF = mdlProcedures.GetCurrency(Me.txtNonHeavy.Text)
            End If
            
            curTotal = ((curPriceListValue * curMF) + (curWeight * curFreight)) / 0.9
            
            Me.txtTotal.Caption = mdlProcedures.FormatCurrency(CStr(curTotal), "#,##0.00")
        End If
    End With
End Sub

Public Property Get PriceListIdCombo() As String
    PriceListIdCombo = mdlProcedures.GetComboData(Me.cmbPriceListId)
End Property

Public Property Let PriceListIdCombo(ByVal strPriceListId As String)
    mdlProcedures.SetComboData Me.cmbPriceListId, strPriceListId
End Property
