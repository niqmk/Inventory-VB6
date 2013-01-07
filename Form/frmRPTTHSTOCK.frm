VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmRPTTHSTOCK 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3735
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   9390
   Icon            =   "frmRPTTHSTOCK.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   9390
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraMain 
      Caption         =   "Detail"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   1
      Left            =   120
      TabIndex        =   18
      Top             =   2400
      Width           =   3735
      Begin MSComCtl2.DTPicker dtpStartDate 
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   600
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
         Format          =   93323267
         CurrentDate     =   39335
      End
      Begin MSComCtl2.DTPicker dtpFinishDate 
         Height          =   375
         Left            =   2040
         TabIndex        =   9
         Top             =   600
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
         Format          =   93323267
         CurrentDate     =   39335
      End
      Begin VB.Label lblStartDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dari Tanggal"
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
         Width           =   1080
      End
      Begin VB.Label lblFinishDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sampai Tanggal"
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
         Left            =   2040
         TabIndex        =   16
         Top             =   240
         Width           =   1395
      End
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Cetak"
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
      Left            =   8040
      TabIndex        =   10
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Frame fraMain 
      Height          =   2295
      Index           =   0
      Left            =   120
      TabIndex        =   17
      Top             =   120
      Width           =   9135
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
         Left            =   7920
         TabIndex        =   4
         Top             =   1800
         Width           =   1095
      End
      Begin VB.CheckBox chkPrintAll 
         Caption         =   "Cetak Semua"
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
         Left            =   3960
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   1800
         Width           =   1695
      End
      Begin VB.CheckBox chkStockEmpty 
         Caption         =   "Stok Kosong"
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
         Left            =   2040
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   1800
         Width           =   1455
      End
      Begin VB.CheckBox chkDetail 
         Caption         =   "Cetak Detail"
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
         Left            =   120
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   1800
         Width           =   1695
      End
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
         Left            =   1560
         TabIndex        =   3
         Top             =   1320
         Width           =   7455
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
         Left            =   1560
         TabIndex        =   2
         Top             =   960
         Width           =   6135
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
         Left            =   1560
         TabIndex        =   1
         Top             =   600
         Width           =   2535
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
         Left            =   1560
         TabIndex        =   0
         Top             =   240
         Width           =   975
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
         TabIndex        =   14
         Top             =   1320
         Width           =   675
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
         TabIndex        =   13
         Top             =   960
         Width           =   510
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
         TabIndex        =   12
         Top             =   600
         Width           =   990
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
         TabIndex        =   11
         Top             =   240
         Width           =   1125
      End
   End
End
Attribute VB_Name = "frmRPTTHSTOCK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private PrintReport As clsRPTTHSTOCK

Private strVendor As String
Private strGroup As String
Private strCategory As String
Private strBrand As String
Private strUnity As String

Private blnParent As Boolean

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub Form_Load()
    SetInitialization
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If mdlGlobal.blnFill Then
        Cancel = 1
    Else
        If blnParent Then
            Cancel = 1
        Else
            If Not PrintReport Is Nothing Then
                Set PrintReport = Nothing
            End If
        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmRPTTHSTOCK = Nothing
End Sub

Private Sub chkDetail_Click()
    If Me.chkDetail.Value = vbChecked Then
        Me.fraMain(1).Enabled = True
    Else
        Me.fraMain(1).Enabled = False
    End If
End Sub

Private Sub cmdOptional_Click()
    If blnParent Then Exit Sub
    
    blnParent = True
    
    mdlProcedures.CornerWindows Me
    
    mdlProcedures.ShowForm frmBRWTMITEMOPT, False, True
End Sub

Private Sub cmdPrint_Click()
    If mdlGlobal.blnFill Then Exit Sub
    
    If Not mdlProcedures.IsValidComboData(Me.cmbWarehouseId) Then
        MsgBox "Gudang Harap Diisi Terlebih Dahulu", vbOKOnly + vbExclamation, Me.Caption
        
        Exit Sub
    End If
    
    SetPrint
End Sub

Private Sub txtItemId_GotFocus()
    mdlProcedures.GotFocus Me.txtItemId
End Sub

Private Sub txtPartNumber_GotFocus()
    mdlProcedures.GotFocus Me.txtPartNumber
End Sub

Private Sub txtName_GotFocus()
    mdlProcedures.GotFocus Me.txtName
End Sub

Private Sub SetInitialization()
    mdlProcedures.CenterWindows Me
    
    Me.Caption = mdlText.strRPTTHSTOCK
    
    blnParent = False
    
    Me.dtpStartDate.CustomFormat = mdlGlobal.strFormatDate
    Me.dtpFinishDate.CustomFormat = mdlGlobal.strFormatDate
    
    FillCombo
    
    Dim strMonth As String
    Dim strYear As String
    
    strMonth = mdlProcedures.FormatDate(Now, "MM")
    strYear = mdlProcedures.FormatDate(Now, "yyyy")
    
    Me.dtpStartDate.Value = mdlProcedures.SetDate(strMonth, strYear)
    Me.dtpFinishDate.Value = mdlProcedures.SetDate(strMonth, strYear, , True)
    
    Me.chkDetail.Value = vbChecked
    Me.chkPrintAll.Value = vbChecked
End Sub

Private Sub FillCombo()
    Dim rstTemp As ADODB.Recordset
    
    Set rstTemp = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "WarehouseId, Name", mdlTable.CreateTMWAREHOUSE, False)
    
    mdlProcedures.FillComboData Me.cmbWarehouseId, rstTemp
    
    mdlDatabase.CloseRecordset rstTemp
End Sub

Private Sub SetPrint()
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
    
    Dim rstMain As ADODB.Recordset
    
    Set rstMain = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, strTableFirst & ".ItemId, PartNumber, " & strTableFirst & ".Name", strTable, False, strCriteria, strTableFirst & ".ItemId ASC")
    
    If Not rstMain.RecordCount > 0 Then
        MsgBox "Data Tidak Ada", vbOKOnly + vbExclamation, Me.Caption
    Else
        Dim strWarehouseId As String
        
        If mdlProcedures.IsValidComboData(Me.cmbWarehouseId) Then
            strWarehouseId = mdlProcedures.GetComboData(Me.cmbWarehouseId)
        Else
            strWarehouseId = mdlTransaction.GetWarehouseIdSet
        End If
        
        If Not PrintReport Is Nothing Then
            Set PrintReport = Nothing
        End If
        
        Set PrintReport = New clsRPTTHSTOCK
        
        mdlGlobal.blnFill = True
        
        PrintReport.ImportToExcel _
            rstMain, _
            Me.dtpStartDate.Value, _
            Me.dtpFinishDate.Value, _
            strWarehouseId, _
            Me.chkStockEmpty.Value, _
            Me.chkDetail.Value, _
            Me.chkPrintAll.Value
            
        mdlGlobal.blnFill = False
    End If
    
    mdlDatabase.CloseRecordset rstMain
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
    
    If Not blnEnable Then mdlProcedures.CenterWindows Me
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
