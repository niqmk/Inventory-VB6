VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmMISTHSTOCK 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7230
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9390
   Icon            =   "frmMISTHSTOCK.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7230
   ScaleWidth      =   9390
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraMain 
      Height          =   2295
      Left            =   120
      TabIndex        =   12
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
         Left            =   6720
         TabIndex        =   5
         Top             =   1800
         Width           =   1095
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
         Left            =   7920
         TabIndex        =   6
         Top             =   1800
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
         Left            =   1560
         TabIndex        =   0
         Top             =   240
         Width           =   975
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
         Left            =   120
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   1800
         Width           =   1455
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
         TabIndex        =   7
         Top             =   240
         Width           =   450
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
         TabIndex        =   8
         Top             =   600
         Width           =   990
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
         TabIndex        =   9
         Top             =   960
         Width           =   510
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
         TabIndex        =   10
         Top             =   1320
         Width           =   675
      End
   End
   Begin MSDataGridLib.DataGrid dgdMain 
      Height          =   4575
      Left            =   120
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   2520
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   8070
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
Attribute VB_Name = "frmMISTHSTOCK"
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

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys "{TAB}"
    ElseIf KeyCode = vbKeyEscape Then
        mdlGlobal.blnFill = False
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
            mdlDatabase.CloseRecordset rstMain
        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmMISTHSTOCK = Nothing
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

Private Sub cmdOptional_Click()
    If mdlGlobal.blnFill Then Exit Sub
    If blnParent Then Exit Sub
    
    blnParent = True
    
    mdlProcedures.CornerWindows Me
    
    mdlProcedures.ShowForm frmBRWTMITEMOPT, False, True
End Sub

Private Sub cmdSearch_Click()
    If mdlGlobal.blnFill Then Exit Sub
    
    If Not mdlProcedures.IsValidComboData(Me.cmbWarehouseId) Then
        MsgBox "Gudang Harap Dipilih Terlebih Dahulu", vbOKOnly + vbExclamation, Me.Caption
        
        Exit Sub
    End If
    
    SetGrid False
End Sub

Private Sub dgdMain_HeadClick(ByVal ColIndex As Integer)
    rstMain.Sort = rstMain.Fields(ColIndex).Name
End Sub

Private Sub SetInitialization()
    mdlProcedures.CenterWindows Me
    
    Me.Caption = mdlText.strMISTHSTOCK
    
    blnParent = False
    
    FillCombo
    
    Me.chkStockEmpty.Value = vbChecked
    
    SetGrid
End Sub

Private Sub FillCombo()
    Dim rstTemp As ADODB.Recordset
    
    Set rstTemp = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "WarehouseId, Name", mdlTable.CreateTMWAREHOUSE, False, , "WarehouseId ASC")
    
    mdlProcedures.FillComboData Me.cmbWarehouseId, rstTemp
    
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
    
    Dim rstTemp As ADODB.Recordset
    
    Set rstTemp = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, strTableFirst & ".ItemId, PartNumber, " & strTableFirst & ".Name", strTable, False, strCriteria, strTableFirst & ".ItemId ASC")
    
    SetVirtualRecordset
    
    Set Me.dgdMain.DataSource = rstMain
    
    With Me.dgdMain
        .RowHeight = 500
        
        .Columns(0).Width = 1100
        .Columns(0).Locked = True
        .Columns(0).Caption = "Kode"
        .Columns(1).Width = 2600
        .Columns(1).Locked = True
        .Columns(1).Caption = "Nomor Part"
        .Columns(1).WrapText = True
        .Columns(2).Width = 3500
        .Columns(2).Locked = True
        .Columns(2).Caption = "Nama"
        .Columns(2).WrapText = True
        .Columns(3).Width = 1300
        .Columns(3).Locked = True
        .Columns(3).Caption = "Qty"
        .Columns(3).NumberFormat = "#,##0"
        .Columns(3).Alignment = dbgRight
    End With
    
    mdlGlobal.blnFill = True
    
    Me.cmbWarehouseId.Enabled = Not mdlGlobal.blnFill
    Me.chkStockEmpty.Enabled = Not mdlGlobal.blnFill
    
    Dim curQty As Currency
    
    With rstTemp
        While Not .EOF
            curQty = mdlTransaction.CheckStock(!ItemId, mdlProcedures.GetComboData(Me.cmbWarehouseId))
            
            If Me.chkStockEmpty.Value = vbChecked Then
                If Not curQty > 0 Then
                    rstMain.AddNew
                    
                    rstMain!ItemId = !ItemId
                    rstMain!PartNumber = !PartNumber
                    rstMain!Name = !Name
                    rstMain!Qty = curQty
                    
                    rstMain.Update
                End If
            Else
                rstMain.AddNew
                
                rstMain!ItemId = !ItemId
                rstMain!PartNumber = !PartNumber
                rstMain!Name = !Name
                rstMain!Qty = curQty
                
                rstMain.Update
            End If
            
            DoEvents
            
            If Not mdlGlobal.blnFill Then
                GoTo HotKeys
            End If
            
            .MoveNext
        Wend
    End With
    
HotKeys:
    
    mdlGlobal.blnFill = False
    
    Me.cmbWarehouseId.Enabled = Not mdlGlobal.blnFill
    Me.chkStockEmpty.Enabled = Not mdlGlobal.blnFill
    
    mdlDatabase.CloseRecordset rstTemp
    
    If rstMain.RecordCount > 0 Then
        rstMain.MoveFirst
    End If
End Sub

Private Sub SetVirtualRecordset()
    mdlDatabase.CloseRecordset rstMain
    
    Set rstMain = New ADODB.Recordset
    rstMain.CursorLocation = adUseClient
    
    With rstMain
        .Fields.Append "ItemId", adChar, mdlDatabase.GetColumnSize(mdlGlobal.conInventory, "ItemId", mdlTable.CreateTMITEM)
        .Fields.Append "PartNumber", adChar, mdlDatabase.GetColumnSize(mdlGlobal.conInventory, "PartNumber", mdlTable.CreateTMITEM)
        .Fields.Append "Name", adChar, mdlDatabase.GetColumnSize(mdlGlobal.conInventory, "Name", mdlTable.CreateTMITEM)
        .Fields.Append "Qty", adCurrency
        
        .Open
    End With
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
