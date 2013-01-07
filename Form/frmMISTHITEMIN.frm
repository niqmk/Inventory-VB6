VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmMISTHITEMIN 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7590
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9030
   Icon            =   "frmMISTHITEMIN.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7590
   ScaleWidth      =   9030
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraSearch 
      Height          =   1215
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   8775
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
         Left            =   1320
         TabIndex        =   0
         Top             =   240
         Width           =   7335
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
         Left            =   7560
         TabIndex        =   3
         Top             =   720
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker dtpStartDate 
         Height          =   375
         Left            =   1320
         TabIndex        =   1
         Top             =   720
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
         Format          =   82247683
         CurrentDate     =   39335
      End
      Begin MSComCtl2.DTPicker dtpFinishDate 
         Height          =   375
         Left            =   5520
         TabIndex        =   2
         Top             =   720
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
         Format          =   82247683
         CurrentDate     =   39335
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
         TabIndex        =   4
         Top             =   240
         Width           =   675
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
         Left            =   3960
         TabIndex        =   6
         Top             =   720
         Width           =   1395
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
         TabIndex        =   5
         Top             =   720
         Width           =   1080
      End
   End
   Begin MSDataGridLib.DataGrid dgdHeader 
      Height          =   3135
      Left            =   120
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1440
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   5530
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
   Begin MSDataGridLib.DataGrid dgdDetail 
      Height          =   2775
      Left            =   120
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   4680
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   4895
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
Attribute VB_Name = "frmMISTHITEMIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private rstHeader As ADODB.Recordset
Private rstDetail As ADODB.Recordset

Private Sub Form_Load()
    SetInitialization
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    mdlDatabase.CloseRecordset rstDetail
    mdlDatabase.CloseRecordset rstHeader
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmMISTHITEMIN = Nothing
End Sub

Private Sub dgdHeader_HeadClick(ByVal ColIndex As Integer)
    rstHeader.Sort = rstHeader.Fields(ColIndex).Name
    
    If rstHeader.RecordCount > 0 Then
        SetGridDetail rstHeader!ItemInId
    Else
        SetGridDetail
    End If
End Sub

Private Sub dgdDetail_HeadClick(ByVal ColIndex As Integer)
    rstDetail.Sort = rstDetail.Fields(ColIndex).Name
End Sub

Private Sub dgdHeader_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If rstHeader.RecordCount > 0 Then
        SetGridDetail rstHeader!ItemInId
    Else
        SetGridDetail
    End If
End Sub

Private Sub cmdSearch_Click()
    SetGridHeader False
End Sub

Private Sub SetInitialization()
    mdlProcedures.CenterWindows Me
    
    Me.Caption = mdlText.strMISTHITEMIN
    
    Me.dtpStartDate.CustomFormat = mdlGlobal.strFormatDate
    Me.dtpFinishDate.CustomFormat = mdlGlobal.strFormatDate
    
    FillCombo
    
    Dim strYear As String
    Dim strMonth As String
        
    strYear = mdlProcedures.FormatDate(Now, "yyyy")
    strMonth = mdlProcedures.FormatDate(Now, "MM")
        
    Me.dtpStartDate.Value = mdlProcedures.SetDate(strMonth, strYear)
    Me.dtpFinishDate.Value = mdlProcedures.SetDate(strMonth, strYear, , True)
    
    SetGridHeader
End Sub

Private Sub FillCombo()
    Dim rstTemp As ADODB.Recordset
    
    Set rstTemp = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "WarehouseId, Name", mdlTable.CreateTMWAREHOUSE, False)
    
    mdlProcedures.FillComboData Me.cmbWarehouseId, rstTemp
    
    mdlDatabase.CloseRecordset rstTemp
End Sub

Private Sub SetGridHeader(Optional ByVal blnInitialize As Boolean = True)
    Dim strCriteria As String
    
    If blnInitialize Then
        strCriteria = "ItemInId=''"
    Else
        strCriteria = ""
        
        If mdlProcedures.IsValidComboData(Me.cmbWarehouseId) Then
            strCriteria = mdlProcedures.QueryLikeCriteria("WarehouseId", mdlProcedures.GetComboData(Me.cmbWarehouseId))
        End If
        
        If Not Trim(strCriteria) = "" Then
            strCriteria = strCriteria & " AND "
        End If
        
        If mdlGlobal.objDatabaseInit = SQLSERVER7 Or mdlGlobal.objDatabaseInit = SQLSERVER2000 Or mdlGlobal.objDatabaseInit = SQLEXPRESS Then
            strCriteria = strCriteria & "ItemInDate>='" & mdlProcedures.FormatDate(Me.dtpStartDate.Value) & "' AND ItemInDate<='" & mdlProcedures.FormatDate(Me.dtpFinishDate.Value) & "'"
        ElseIf mdlGlobal.objDatabaseInit = MSACCESS Then
            strCriteria = strCriteria & "ItemInDate>=#" & mdlProcedures.FormatDate(Me.dtpStartDate.Value) & "# AND ItemInDate<=#" & mdlProcedures.FormatDate(Me.dtpFinishDate.Value) & "#"
        ElseIf mdlGlobal.objDatabaseInit = MYSQL Then
            strCriteria = strCriteria & "ItemInDate>='" & mdlProcedures.FormatDate(Me.dtpStartDate.Value) & "' AND ItemInDate<='" & mdlProcedures.FormatDate(Me.dtpFinishDate.Value) & "'"
        End If
    End If
    
    Dim rstTemp As ADODB.Recordset
    
    Set rstTemp = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "ItemInId, ItemInDate, WarehouseId", mdlTable.CreateTHITEMIN, False, strCriteria, "ItemInDate ASC")
    
    SetVirtualRecordset
    
    With rstTemp
        While Not .EOF
            rstHeader.AddNew
            
            rstHeader!ItemInId = !ItemInId
            rstHeader!ItemInDate = mdlProcedures.FormatDate(!ItemInDate)
            rstHeader!WarehouseId = !WarehouseId
            rstHeader!WarehouseName = mdlDatabase.GetFieldData(mdlGlobal.conInventory, "Name", mdlTable.CreateTMWAREHOUSE, "WarehouseId='" & !WarehouseId & "'")
            
            rstHeader.Update
            
            .MoveNext
        Wend
    End With
    
    mdlDatabase.CloseRecordset rstTemp
    
    If rstHeader.RecordCount > 0 Then
        rstHeader.MoveFirst
        
        SetGridDetail rstHeader!ItemInId
    Else
        SetGridDetail
    End If
    
    Set Me.dgdHeader.DataSource = rstHeader
    
    With Me.dgdHeader
        .RowHeight = 500
        
        .Columns(0).Width = 0
        .Columns(0).Locked = True
        .Columns(0).Visible = False
        .Columns(1).Width = 2400
        .Columns(1).Locked = True
        .Columns(1).Caption = "Tanggal"
        .Columns(1).NumberFormat = "dd MMMM yyyy"
        .Columns(2).Width = 1000
        .Columns(2).Locked = True
        .Columns(2).Caption = "Gudang"
        .Columns(3).Width = 4600
        .Columns(3).Locked = True
        .Columns(3).Caption = "Nama"
        .Columns(3).WrapText = True
    End With
End Sub

Private Sub SetVirtualRecordset()
    mdlDatabase.CloseRecordset rstHeader
    
    Set rstHeader = New ADODB.Recordset
    rstHeader.CursorLocation = adUseClient
    
    With rstHeader
        .Fields.Append "ItemInId", adChar, mdlDatabase.GetColumnSize(mdlGlobal.conInventory, "ItemInId", mdlTable.CreateTHITEMIN)
        .Fields.Append "ItemInDate", adDate
        .Fields.Append "WarehouseId", adChar, mdlDatabase.GetColumnSize(mdlGlobal.conInventory, "WarehouseId", mdlTable.CreateTHITEMIN)
        .Fields.Append "WarehouseName", adChar, mdlDatabase.GetColumnSize(mdlGlobal.conInventory, "Name", mdlTable.CreateTMWAREHOUSE)
        
        .Open
    End With
End Sub

Private Sub SetGridDetail(Optional ByVal strItemInId As String = "")
    Dim strTableFirst As String
    Dim strTableSecond As String
    Dim strTable As String
    
    strTableFirst = mdlTable.CreateTDITEMIN
    strTableSecond = mdlTable.CreateTMITEM
    
    strTable = strTableFirst & " LEFT JOIN " & strTableSecond & _
        " ON " & strTableFirst & ".ItemId=" & strTableSecond & ".ItemId"
    
    Dim strCriteria As String
    
    strCriteria = "ItemInId='" & strItemInId & "'"
    
    Set rstDetail = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, strTableFirst & ".ItemId, Name, Qty, UnityId", strTable, False, strCriteria, strTableFirst & ".ItemId ASC")
    
    Set Me.dgdDetail.DataSource = rstDetail
    
    With Me.dgdDetail
        .RowHeight = 500
        
        .Columns(0).Width = 1100
        .Columns(0).Locked = True
        .Columns(0).Caption = "Kode"
        .Columns(1).Width = 4800
        .Columns(1).Locked = True
        .Columns(1).Caption = "Nama"
        .Columns(1).WrapText = True
        .Columns(2).Width = 1200
        .Columns(2).Locked = True
        .Columns(2).Caption = "Qty"
        .Columns(2).NumberFormat = "#,##0"
        .Columns(2).Alignment = dbgRight
        .Columns(3).Width = 1000
        .Columns(3).Locked = True
        .Columns(3).Caption = "Satuan"
    End With
End Sub
