VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmTDPOBUY 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   8775
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8895
   Icon            =   "frmTDPOBUY.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8775
   ScaleWidth      =   8895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraHeader 
      Height          =   2175
      Left            =   120
      TabIndex        =   35
      Top             =   6000
      Width           =   8655
      Begin VB.TextBox txtQty 
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
         Left            =   1200
         MaxLength       =   10
         TabIndex        =   6
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton optItemPrice 
         Caption         =   "Berdasarkan History"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   1200
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   600
         Width           =   3375
      End
      Begin VB.OptionButton optItemPrice 
         Caption         =   "Custom"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   5160
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   600
         Width           =   3375
      End
      Begin VB.Frame fraItemPrice 
         Height          =   1095
         Index           =   0
         Left            =   120
         TabIndex        =   36
         Top             =   960
         Width           =   4815
         Begin VB.ComboBox cmbPriceDate 
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
            Left            =   1080
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   240
            Width           =   3615
         End
         Begin VB.TextBox txtItemPrice 
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
            Index           =   0
            Left            =   1080
            MaxLength       =   10
            TabIndex        =   8
            Top             =   720
            Width           =   1335
         End
         Begin VB.Label lblPriceDate 
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
            TabIndex        =   28
            Top             =   240
            Width           =   675
         End
         Begin VB.Label lblItemPrice 
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
            Index           =   1
            Left            =   120
            TabIndex        =   29
            Top             =   720
            Width           =   510
         End
      End
      Begin VB.Frame fraItemPrice 
         Height          =   1095
         Index           =   1
         Left            =   5040
         TabIndex        =   37
         Top             =   960
         Width           =   3495
         Begin VB.TextBox txtItemPrice 
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
            Index           =   1
            Left            =   1200
            MaxLength       =   10
            TabIndex        =   9
            Top             =   240
            Width           =   1335
         End
         Begin VB.ComboBox cmbCurrencyId 
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
            TabIndex        =   10
            Top             =   600
            Width           =   975
         End
         Begin VB.Label lblItemPrice 
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
            Index           =   2
            Left            =   120
            TabIndex        =   30
            Top             =   240
            Width           =   510
         End
         Begin VB.Label lblCurrencyId 
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
            TabIndex        =   31
            Top             =   600
            Width           =   945
         End
      End
      Begin VB.Label lblQty 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah"
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
         TabIndex        =   22
         Top             =   240
         Width           =   615
      End
      Begin VB.Label lblItemPrice 
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
         Index           =   0
         Left            =   120
         TabIndex        =   25
         Top             =   600
         Width           =   510
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
         Left            =   5160
         TabIndex        =   23
         Top             =   240
         Width           =   420
      End
      Begin VB.Label txtTotal 
         Alignment       =   1  'Right Justify
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
         Left            =   6360
         TabIndex        =   24
         Top             =   240
         Width           =   2175
      End
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
      Left            =   7440
      TabIndex        =   11
      Top             =   8280
      Width           =   1335
   End
   Begin VB.Frame fraSearch 
      Height          =   1335
      Left            =   120
      TabIndex        =   33
      Top             =   120
      Width           =   8655
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
         Left            =   7440
         TabIndex        =   3
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox txtPartNumber 
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
         Left            =   1200
         TabIndex        =   1
         Top             =   600
         Width           =   4935
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "Filter"
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
         TabIndex        =   4
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox txtItemId 
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
         Left            =   1200
         TabIndex        =   0
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtName 
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
         Left            =   1200
         TabIndex        =   2
         Top             =   960
         Width           =   6135
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
         TabIndex        =   13
         Top             =   600
         Width           =   990
      End
      Begin VB.Label lblItemId 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kode"
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
         Top             =   240
         Width           =   450
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
         TabIndex        =   14
         Top             =   960
         Width           =   510
      End
   End
   Begin VB.Frame fraInfo 
      Height          =   1815
      Left            =   120
      TabIndex        =   34
      Top             =   4200
      Width           =   8655
      Begin VB.ComboBox cmbWarehouseId 
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
         TabIndex        =   5
         Top             =   960
         Width           =   7335
      End
      Begin VB.Label lblWarehouseId 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Height          =   240
         Left            =   120
         TabIndex        =   19
         Top             =   960
         Width           =   675
      End
      Begin VB.Label lblStockQty 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Stok"
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
         Top             =   1440
         Width           =   405
      End
      Begin VB.Label txtStockQty 
         Alignment       =   1  'Right Justify
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
         Left            =   1200
         TabIndex        =   21
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label txtUnityId 
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
         Left            =   1200
         TabIndex        =   18
         Top             =   600
         Width           =   6975
      End
      Begin VB.Label txtItemIdText 
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
         Left            =   1200
         TabIndex        =   16
         Top             =   240
         Width           =   7335
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
         TabIndex        =   17
         Top             =   600
         Width           =   615
      End
      Begin VB.Label lblItemIdText 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Height          =   240
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   615
      End
   End
   Begin MSDataGridLib.DataGrid dgdMain 
      Height          =   2535
      Left            =   120
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   1560
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   4471
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
End
Attribute VB_Name = "frmTDPOBUY"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private rstMain As ADODB.Recordset

Private strVendor As String
Private strGroup As String
Private strCategory As String
Private strBrand As String
Private strUnity As String

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
    Else
        If frmTHPOBUY.Parent Then
            frmTHPOBUY.Parent = False
        End If
        
        mdlDatabase.CloseRecordset rstMain
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmTDPOBUY = Nothing
End Sub

Private Sub cmdSearch_Click()
    SetGrid False
End Sub

Private Sub txtItemId_GotFocus()
    mdlProcedures.GotFocus Me.txtItemId
End Sub

Private Sub txtName_GotFocus()
    mdlProcedures.GotFocus Me.txtName
End Sub

Private Sub txtQty_GotFocus()
    mdlProcedures.GotFocus Me.txtQty
End Sub

Private Sub cmbWarehouseId_Validate(Cancel As Boolean)
    If rstMain.RecordCount > 0 Then
        Me.txtStockQty.Caption = mdlProcedures.FormatCurrency(mdlTransaction.CheckStock(rstMain!ItemId, mdlProcedures.GetComboData(Me.cmbWarehouseId)))
    End If
End Sub

Private Sub txtQty_Change()
    Me.txtQty.Text = mdlProcedures.FormatCurrency(Me.txtQty.Text)
    
    Me.txtQty.SelStart = Len(Me.txtQty.Text)
End Sub

Private Sub txtQty_Validate(Cancel As Boolean)
    If Me.optItemPrice(0).Value Then
        CalculateTotalByPriceId mdlProcedures.GetComboData(Me.cmbPriceDate, , , " : "), mdlProcedures.GetCurrency(Me.txtItemPrice(0).Text)
    ElseIf Me.optItemPrice(1).Value Then
        CalculateTotalByCurrencyId mdlProcedures.GetComboData(Me.cmbCurrencyId), mdlProcedures.GetCurrency(Me.txtItemPrice(1).Text)
    Else
        Me.txtTotal.Caption = "0"
    End If
End Sub

Private Sub txtItemPrice_Change(Index As Integer)
    Me.txtItemPrice(Index).Text = mdlProcedures.FormatCurrency(Me.txtItemPrice(Index).Text)
    
    Me.txtItemPrice(Index).SelStart = Len(Me.txtItemPrice(Index).Text)
End Sub

Private Sub txtItemPrice_Validate(Index As Integer, Cancel As Boolean)
    If Not Me.optItemPrice(Index).Value Then Me.optItemPrice(Index).Value = True
    
    Select Case Index
        Case 0:
            CalculateTotalByPriceId mdlProcedures.GetComboData(Me.cmbPriceDate, , , " : "), mdlProcedures.GetCurrency(Me.txtItemPrice(0).Text)
        Case 1:
            CalculateTotalByCurrencyId mdlProcedures.GetComboData(Me.cmbCurrencyId), mdlProcedures.GetCurrency(Me.txtItemPrice(1).Text)
    End Select
End Sub

Private Sub cmbPriceDate_Validate(Cancel As Boolean)
    If Not mdlProcedures.IsValidComboData(Me.cmbPriceDate) Then Exit Sub
    
    If Not Me.optItemPrice(0).Value Then Me.optItemPrice(0).Value = True
    
    Me.txtItemPrice(0).Text = mdlDatabase.GetFieldData(mdlGlobal.conInventory, "ItemPrice", mdlTable.CreateTMPRICEBUY, "PriceId='" & mdlProcedures.GetComboData(Me.cmbPriceDate, , , " : ") & "'")
    
    CalculateTotalByPriceId mdlProcedures.GetComboData(Me.cmbPriceDate, , , " : "), mdlProcedures.GetCurrency(Me.txtItemPrice(0).Text)
End Sub

Private Sub cmbCurrencyId_Validate(Cancel As Boolean)
    CalculateTotalByCurrencyId mdlProcedures.GetComboData(Me.cmbCurrencyId), mdlProcedures.GetCurrency(Me.txtItemPrice(1).Text)
End Sub

Private Sub optItemPrice_Click(Index As Integer)
    Select Case Index
        Case 0:
            Me.cmbPriceDate.SetFocus
        Case 1:
            Me.txtItemPrice(1).SetFocus
    End Select
End Sub

Private Sub dgdMain_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If rstMain.RecordCount > 0 Then
        FillText
    End If
End Sub

Private Sub cmdOptional_Click()
    If blnParent Then Exit Sub
    
    blnParent = True
    
    mdlProcedures.CornerWindows Me, , False
    
    mdlProcedures.ShowForm frmBRWTMITEMOPT, False, True
End Sub

Private Sub cmdSave_Click()
    If Not CheckValidation Then Exit Sub
    
    Dim mPriceId As String
    
    If Me.optItemPrice(0).Value Then
        mPriceId = mdlDatabase.GetFieldData(mdlGlobal.conInventory, "PriceId", mdlTable.CreateTMPRICEBUY, "PriceId='" & mdlProcedures.GetComboData(Me.cmbPriceDate, , , " : ") & "'")
    ElseIf Me.optItemPrice(1).Value Then
        Dim rstTemp As ADODB.Recordset
        
        Set rstTemp = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "*", mdlTable.CreateTMPRICEBUY, , "PriceId='" & rstMain!ItemId & mdlProcedures.FormatDate(Now, "ddMMyyyy") & "'")
        
        With rstTemp
            If Not .RecordCount > 0 Then
                .AddNew
            
                !PriceId = rstMain!ItemId & mdlProcedures.FormatDate(Now, "ddMMyyyy")
                
                !ItemId = rstMain!ItemId
                !PriceDate = mdlProcedures.FormatDate(Now)
                
                !CreateId = Trim(mdlGlobal.UserAuthority.UserId)
                !CreateDate = mdlProcedures.FormatDate(Now)
            End If
            
            !ItemPrice = mdlProcedures.GetCurrency(Trim(Me.txtItemPrice(1).Text))
            !CurrencyId = mdlProcedures.RepDupText(Trim(mdlProcedures.GetComboData(Me.cmbCurrencyId)))
            
            !UpdateId = Trim(mdlGlobal.UserAuthority.UserId)
            !UpdateDate = mdlProcedures.FormatDate(Now)
            
            .Update
        End With
        
        mdlDatabase.CloseRecordset rstTemp
        
        mPriceId = rstMain!ItemId & mdlProcedures.FormatDate(Now, "ddMMyyyy")
    End If
    
    frmTHPOBUY.SaveDetail rstMain!ItemId, mdlProcedures.GetCurrency(Me.txtQty.Text), mPriceId
End Sub

Private Function CheckValidation() As Boolean
    If rstMain.RecordCount > 0 Then
        If Not CheckOption Then
            MsgBox "Harga Harap Dipilih Terlebih Dahulu", vbOKOnly + vbExclamation, Me.Caption
            
            Me.optItemPrice(0).SetFocus
            
            CheckValidation = False
            
            Exit Function
        End If
        
        If Me.optItemPrice(0).Value Then
            If Not mdlProcedures.IsValidComboData(Me.cmbPriceDate) Then
                MsgBox "Tanggal Harga Harap Dipilih Terlebih Dahulu", vbOKOnly + vbExclamation, Me.Caption
                
                Me.cmbPriceDate.SetFocus
                
                CheckValidation = False
                
                Exit Function
            End If
        ElseIf Me.optItemPrice(1).Value Then
            If Not mdlGlobal.UserAuthority.IsMenuAccess(mdiMain.mnuTMITEMPRICE.Name) Then
                Me.optItemPrice(1).SetFocus
                
                CheckValidation = False
                
                Exit Function
            ElseIf Not mdlProcedures.IsValidComboData(Me.cmbCurrencyId) Then
                MsgBox "Mata Uang Harap Dipilih Terlebih Dahulu", vbOKOnly + vbExclamation, Me.Caption
                
                Me.cmbCurrencyId.SetFocus
                
                CheckValidation = False
                
                Exit Function
            ElseIf Not mdlProcedures.GetCurrency(Me.txtItemPrice(1).Text) > 0 Then
                MsgBox "Harga Harap Diisi Terlebih Dahulu", vbOKOnly + vbExclamation, Me.Caption
                
                Me.txtItemPrice(1).SetFocus
                
                CheckValidation = False
                
                Exit Function
            End If
        End If
        
        Dim curQty As Currency
        Dim curStockQty As Currency
        Dim curMaxStock As Currency
        
        curQty = mdlProcedures.GetCurrency(Me.txtQty.Text)
        curStockQty = mdlProcedures.GetCurrency(Me.txtStockQty.Caption)
        curMaxStock = mdlProcedures.GetCurrency(mdlDatabase.GetFieldData(mdlGlobal.conInventory, "MaxStock", mdlTable.CreateTMITEM, "ItemId='" & rstMain!ItemId & "'"))
        
        If curQty > 0 Then
            If curMaxStock > 0 Then
                If (curStockQty + curQty) > curMaxStock Then
                    MsgBox "Stok Maksimum : " & CStr(curMaxStock), vbOKOnly + vbExclamation, Me.Caption
                    
                    Me.txtQty.SetFocus
                    
                    CheckValidation = False
                    
                    Exit Function
                End If
            End If
        Else
            MsgBox "Qty Harap Diisi Terlebih Dahulu", vbOKOnly + vbExclamation, Me.Caption
            
            Me.txtQty.SetFocus
            
            CheckValidation = False
            
            Exit Function
        End If
    Else
        Me.txtItemId.SetFocus
        
        CheckValidation = False
        
        Exit Function
    End If
    
    CheckValidation = True
End Function

Private Function CheckOption() As Boolean
    Dim blnFound As Boolean
    
    blnFound = False
    
    Dim intCounter As Integer
    
    For intCounter = 0 To Me.optItemPrice.Count - 1
        If Me.optItemPrice(intCounter).Value Then
            blnFound = True
            
            Exit For
        End If
    Next intCounter
    
    CheckOption = blnFound
End Function

Private Sub SetInitialization()
    mdlProcedures.CenterWindows Me, False
    
    Me.Caption = mdlText.strTDPOBUY
    
    blnParent = False
    
    FillCombo
    
    SetGrid False
End Sub

Private Sub FillCombo()
    Me.FillComboTMWAREHOUSE
    Me.FillComboTMCURRENCY
End Sub

Private Sub FillComboTMPRICEBUY(Optional ByVal strItemId As String = "")
    Me.cmbPriceDate.Clear
    
    Dim rstTemp As ADODB.Recordset
    
    Set rstTemp = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "PriceId, PriceDate", mdlTable.CreateTMPRICEBUY, False, "ItemId='" & strItemId & "'")
    
    With rstTemp
        While Not .EOF
            Me.cmbPriceDate.AddItem !PriceId & " : " & mdlProcedures.FormatDate(!PriceDate, mdlGlobal.strFormatDate)
            
            .MoveNext
        Wend
    End With
    
    mdlDatabase.CloseRecordset rstTemp
End Sub

Public Sub FillComboTMCURRENCY()
    Dim rstTemp As ADODB.Recordset
    
    Set rstTemp = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "CurrencyId", mdlTable.CreateTMCURRENCY)
    
    mdlProcedures.FillComboData Me.cmbCurrencyId, rstTemp
    
    mdlDatabase.CloseRecordset rstTemp
End Sub

Public Sub FillComboTMWAREHOUSE()
    Dim rstTemp As ADODB.Recordset
    
    Set rstTemp = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "WarehouseId, Name", mdlTable.CreateTMWAREHOUSE, False, , "WarehouseSet DESC")
    
    mdlProcedures.FillComboData Me.cmbWarehouseId, rstTemp
    
    If Me.cmbWarehouseId.ListCount > 0 Then Me.cmbWarehouseId.ListIndex = 0
    
    mdlDatabase.CloseRecordset rstTemp
End Sub

Private Sub SetGrid(Optional ByVal blnInitialize As Boolean = True)
    Dim strTableFirst As String
    Dim strTableSecond As String
    Dim strTableThird As String
    Dim strTableFourth As String
    Dim strTableFifth As String
    Dim strTableSixth As String
    Dim strTable As String
    
    strTableFirst = mdlTable.CreateTMITEM
    strTableSecond = mdlTable.CreateTMVENDOR
    strTableThird = mdlTable.CreateTMGROUP
    strTableFourth = mdlTable.CreateTMCATEGORY
    strTableFifth = mdlTable.CreateTMBRAND
    strTableSixth = mdlTable.CreateTMUNITY
    
    strTable = "(((((" & strTableFirst & " LEFT JOIN " & strTableSecond & _
        " ON " & strTableFirst & ".VendorId=" & strTableSecond & ".VendorId) LEFT JOIN " & strTableThird & _
        " ON " & strTableFirst & ".GroupId=" & strTableThird & ".GroupId) LEFT JOIN " & strTableFourth & _
        " ON " & strTableFirst & ".CategoryId=" & strTableFourth & ".CategoryId) LEFT JOIN " & strTableFifth & _
        " ON " & strTableFirst & ".BrandId=" & strTableFifth & ".BrandId) LEFT JOIN " & strTableSixth & _
        " ON " & strTableFirst & ".UnityId=" & strTableSixth & ".UnityId)"
        
    Dim strCriteria As String
    
    strCriteria = ""
    
    If blnInitialize Then
        strCriteria = strTableFirst & ".ItemId=''"
    Else
        If Not Trim(Me.txtItemId.Text) = "" Then
            strCriteria = mdlProcedures.QueryLikeCriteria(strTableFirst & ".ItemId", mdlProcedures.RepDupText(Me.txtItemId.Text))
        End If
        
        If Not Trim(Me.txtPartNumber.Text) = "" Then
            If Not Trim(strCriteria) = "" Then
                strCriteria = strCriteria & " AND "
            End If
            
            strCriteria = strCriteria & mdlProcedures.QueryLikeCriteria("PartNumber", mdlProcedures.RepDupText(Me.txtPartNumber.Text))
        End If
        
        If Not Trim(Me.txtName.Text) = "" Then
            If Not Trim(strCriteria) = "" Then
                strCriteria = strCriteria & " AND "
            End If
            
            strCriteria = strCriteria & mdlProcedures.QueryLikeCriteria(strTableFirst & ".Name", mdlProcedures.RepDupText(Me.txtName.Text))
        End If
        
        Dim strOptional(1) As String
        
        If Not Trim(strVendor) = "" Then
            strOptional(0) = mdlProcedures.SplitData(strVendor)
            strOptional(1) = mdlProcedures.SplitData(strVendor, 1)
            
            If Not Trim(strOptional(0)) = "" Then
                If Not Trim(strCriteria) = "" Then
                    strCriteria = strCriteria & " AND "
                End If
                
                strCriteria = strCriteria & mdlProcedures.QueryLikeCriteria(strTableFirst & ".VendorId", mdlProcedures.RepDupText(strOptional(0)))
            End If
            
            If Not Trim(strOptional(1)) = "" Then
                If Not Trim(strCriteria) = "" Then
                    strCriteria = strCriteria & " AND "
                End If
                
                strCriteria = strCriteria & mdlProcedures.QueryLikeCriteria(strTableSecond & ".Name", mdlProcedures.RepDupText(strOptional(1)))
            End If
        End If
        
        If Not Trim(strGroup) = "" Then
            strOptional(0) = mdlProcedures.SplitData(strGroup)
            strOptional(1) = mdlProcedures.SplitData(strGroup, 1)
            
            If Not Trim(strOptional(0)) = "" Then
                If Not Trim(strCriteria) = "" Then
                    strCriteria = strCriteria & " AND "
                End If
                
                strCriteria = strCriteria & mdlProcedures.QueryLikeCriteria(strTableFirst & ".GroupId", mdlProcedures.RepDupText(strOptional(0)))
            End If
            
            If Not Trim(strOptional(1)) = "" Then
                If Not Trim(strCriteria) = "" Then
                    strCriteria = strCriteria & " AND "
                End If
                
                strCriteria = strCriteria & mdlProcedures.QueryLikeCriteria(strTableThird & ".Name", mdlProcedures.RepDupText(strOptional(1)))
            End If
        End If
        
        If Not Trim(strCategory) = "" Then
            strOptional(0) = mdlProcedures.SplitData(strCategory)
            strOptional(1) = mdlProcedures.SplitData(strCategory, 1)
            
            If Not Trim(strOptional(0)) = "" Then
                If Not Trim(strCriteria) = "" Then
                    strCriteria = strCriteria & " AND "
                End If
                
                strCriteria = strCriteria & mdlProcedures.QueryLikeCriteria(strTableFirst & ".CategoryId", mdlProcedures.RepDupText(strOptional(0)))
            End If
            
            If Not Trim(strOptional(1)) = "" Then
                If Not Trim(strCriteria) = "" Then
                    strCriteria = strCriteria & " AND "
                End If
                
                strCriteria = strCriteria & mdlProcedures.QueryLikeCriteria(strTableFourth & ".Name", mdlProcedures.RepDupText(strOptional(1)))
            End If
        End If
        
        If Not Trim(strBrand) = "" Then
            strOptional(0) = mdlProcedures.SplitData(strBrand)
            strOptional(1) = mdlProcedures.SplitData(strBrand, 1)
            
            If Not Trim(strOptional(0)) = "" Then
                If Not Trim(strCriteria) = "" Then
                    strCriteria = strCriteria & " AND "
                End If
                
                strCriteria = strCriteria & mdlProcedures.QueryLikeCriteria(strTableFirst & ".BrandId", mdlProcedures.RepDupText(strOptional(0)))
            End If
            
            If Not Trim(strOptional(1)) = "" Then
                If Not Trim(strCriteria) = "" Then
                    strCriteria = strCriteria & " AND "
                End If
                
                strCriteria = strCriteria & mdlProcedures.QueryLikeCriteria(strTableFifth & ".Name", mdlProcedures.RepDupText(strOptional(1)))
            End If
        End If
        
        If Not Trim(strUnity) = "" Then
            strOptional(0) = mdlProcedures.SplitData(strUnity)
            strOptional(1) = mdlProcedures.SplitData(strUnity, 1)
            
            If Not Trim(strOptional(0)) = "" Then
                If Not Trim(strCriteria) = "" Then
                    strCriteria = strCriteria & " AND "
                End If
                
                strCriteria = strCriteria & mdlProcedures.QueryLikeCriteria(strTableFirst & ".UnityId", mdlProcedures.RepDupText(strOptional(0)))
            End If
            
            If Not Trim(strOptional(1)) = "" Then
                If Not Trim(strCriteria) = "" Then
                    strCriteria = strCriteria & " AND "
                End If
                
                strCriteria = strCriteria & mdlProcedures.QueryLikeCriteria(strTableSixth & ".Name", mdlProcedures.RepDupText(strOptional(1)))
            End If
        End If
    End If
    
    Set rstMain = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, strTableFirst & ".ItemId, PartNumber, " & strTableFirst & ".Name, " & strTableFirst & ".UnityId", strTable, False, strCriteria, strTableFirst & ".ItemId ASC")
    
    If rstMain.RecordCount > 0 Then
        FillText
    Else
        FillText True
    End If
    
    Set Me.dgdMain.DataSource = rstMain
    
    With Me.dgdMain
        .RowHeight = 500
        
        .Columns(0).Width = 1200
        .Columns(0).Locked = True
        .Columns(0).Caption = "Kode"
        .Columns(1).Width = 2500
        .Columns(1).Locked = True
        .Columns(1).Caption = "Nomor Part"
        .Columns(1).WrapText = True
        .Columns(2).Width = 3100
        .Columns(2).Locked = True
        .Columns(2).Caption = "Nama"
        .Columns(2).WrapText = True
        .Columns(3).Width = 1200
        .Columns(3).Locked = True
        .Columns(3).Caption = "Satuan"
    End With
End Sub

Private Sub FillText(Optional ByVal blnClear As Boolean = False)
    If blnClear Then
        Me.txtItemIdText.Caption = ""
        Me.txtUnityId.Caption = ""
        Me.txtStockQty.Caption = ""
        
        FillComboTMPRICEBUY
    Else
        With rstMain
            Me.txtItemIdText.Caption = !ItemId & " | " & !Name
            
            If Trim(!UnityId) = "" Then
                Me.txtUnityId.Caption = !UnityId
            Else
                Me.txtUnityId.Caption = !UnityId & " | " & mdlDatabase.GetFieldData(mdlGlobal.conInventory, "Name", mdlTable.CreateTMUNITY, "UnityId='" & !UnityId & "'")
            End If
            
            Me.txtStockQty.Caption = mdlProcedures.FormatCurrency(mdlTransaction.CheckStock(!ItemId, mdlProcedures.GetComboData(Me.cmbWarehouseId)))
            
            FillComboTMPRICEBUY !ItemId
        End With
    End If
End Sub

Private Sub CalculateTotalByPriceId(ByVal strPriceId As String, ByVal curItemPrice As Currency)
    Dim strCurrencyToId As String
    
    strCurrencyToId = mdlDatabase.GetFieldData(mdlGlobal.conInventory, "CurrencyId", mdlTable.CreateTMPRICEBUY, "PriceId='" & strPriceId & "'")
    
    Dim mItemPrice As Currency
    
    mItemPrice = _
        mdlTransaction.ConvertCurrency( _
            frmTHPOBUY.CurrencyIdCombo, _
            strCurrencyToId, _
            mdlProcedures.GetCurrency(mdlDatabase.GetFieldData(mdlGlobal.conInventory, "ItemPrice", mdlTable.CreateTMPRICEBUY, "PriceId='" & strPriceId & "'")))
    
    Me.txtTotal.Caption = mdlProcedures.FormatCurrency(CStr(mItemPrice * mdlProcedures.GetCurrency(Me.txtQty.Text)))
End Sub

Private Sub CalculateTotalByCurrencyId(ByVal strCurrencyId As String, ByVal curItemPrice As Currency)
    Dim mItemPrice As Currency
    
    mItemPrice = mdlTransaction.ConvertCurrency(frmTHPOBUY.CurrencyIdCombo, strCurrencyId, curItemPrice)
    
    Me.txtTotal.Caption = mdlProcedures.FormatCurrency(CStr(mItemPrice * mdlProcedures.GetCurrency(Me.txtQty.Text)))
End Sub

Public Property Get Parent() As Boolean
    Parent = blnParent
End Property

Public Property Get VendorOptional() As String
    VendorOptional = strVendor
End Property

Public Property Get GroupOptional() As String
    GroupOptional = strGroup
End Property

Public Property Get CategoryOptional() As String
    CategoryOptional = strCategory
End Property

Public Property Get BrandOptional() As String
    BrandOptional = strBrand
End Property

Public Property Get UnityOptional() As String
    UnityOptional = strUnity
End Property

Public Property Let Parent(ByVal blnEnable As Boolean)
    blnParent = blnEnable
    
    If Not blnEnable Then mdlProcedures.CenterWindows Me, False
End Property

Public Property Let VendorOptional(ByVal strValue As String)
    strVendor = strValue
End Property

Public Property Let GroupOptional(ByVal strValue As String)
    strGroup = strValue
End Property

Public Property Let CategoryOptional(ByVal strValue As String)
    strCategory = strValue
End Property

Public Property Let BrandOptional(ByVal strValue As String)
    strBrand = strValue
End Property

Public Property Let UnityOptional(ByVal strValue As String)
    strUnity = strValue
End Property
