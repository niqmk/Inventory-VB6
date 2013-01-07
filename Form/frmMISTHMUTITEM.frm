VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmMISTHMUTITEM 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7470
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   9270
   Icon            =   "frmMISTHMUTITEM.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7470
   ScaleWidth      =   9270
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraSearch 
      Height          =   1695
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   9015
      Begin VB.ComboBox cmbWarehouseTo 
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
         TabIndex        =   1
         Top             =   720
         Width           =   7335
      End
      Begin VB.ComboBox cmbWarehouseFrom 
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
         Left            =   7800
         TabIndex        =   4
         Top             =   1200
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker dtpStartDate 
         Height          =   375
         Left            =   1560
         TabIndex        =   2
         Top             =   1200
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
         Format          =   82313219
         CurrentDate     =   39335
      End
      Begin MSComCtl2.DTPicker dtpFinishDate 
         Height          =   375
         Left            =   5760
         TabIndex        =   3
         Top             =   1200
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
         Format          =   82313219
         CurrentDate     =   39335
      End
      Begin VB.Label lblWarehouseTo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Gudang Tujuan"
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
         TabIndex        =   6
         Top             =   720
         Width           =   1305
      End
      Begin VB.Label lblWarehouseFrom 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Gudang Asal"
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
         Top             =   240
         Width           =   1125
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
         Left            =   4200
         TabIndex        =   8
         Top             =   1200
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
         TabIndex        =   7
         Top             =   1200
         Width           =   1080
      End
   End
   Begin MSDataGridLib.DataGrid dgdDetail 
      Height          =   2775
      Left            =   120
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   4560
      Width           =   9015
      _ExtentX        =   15901
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
   Begin MSDataGridLib.DataGrid dgdHeader 
      Height          =   2535
      Left            =   120
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1920
      Width           =   9015
      _ExtentX        =   15901
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
Attribute VB_Name = "frmMISTHMUTITEM"
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
    mdlDatabase.CloseRecordset rstHeader
    mdlDatabase.CloseRecordset rstDetail
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmMISTHMUTITEM = Nothing
End Sub

Private Sub cmdSearch_Click()
    SetGridHeader False
End Sub

Private Sub dgdHeader_HeadClick(ByVal ColIndex As Integer)
    rstHeader.Sort = rstHeader.Fields(ColIndex).Name
    
    If rstHeader.RecordCount > 0 Then
        SetGridDetail rstHeader!MutId
    End If
End Sub

Private Sub dgdDetail_HeadClick(ByVal ColIndex As Integer)
    rstDetail.Sort = rstDetail.Fields(ColIndex).Name
End Sub

Private Sub dgdHeader_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If rstHeader.RecordCount > 0 Then
        SetGridDetail rstHeader!MutId
    End If
End Sub

Private Sub SetInitialization()
    mdlProcedures.CenterWindows Me
    
    Me.Caption = mdlText.strMISTHMUTITEM
    
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
    
    mdlProcedures.FillComboData Me.cmbWarehouseFrom, rstTemp
    mdlProcedures.FillComboData Me.cmbWarehouseTo, rstTemp
    
    mdlDatabase.CloseRecordset rstTemp
End Sub

Private Sub SetGridHeader(Optional ByVal blnInitialize As Boolean = True)
    Dim strCriteria As String
    
    If blnInitialize Then
        strCriteria = "MutId=''"
    Else
        strCriteria = ""
        
        If mdlProcedures.IsValidComboData(Me.cmbWarehouseFrom) Then
            strCriteria = mdlProcedures.QueryLikeCriteria("WarehouseFrom", mdlProcedures.GetComboData(Me.cmbWarehouseFrom))
        End If
        
        If mdlProcedures.IsValidComboData(Me.cmbWarehouseTo) Then
            If Not Trim(strCriteria) = "" Then
                strCriteria = strCriteria & " AND "
            End If
            
            strCriteria = mdlProcedures.QueryLikeCriteria("WarehouseTo", mdlProcedures.GetComboData(Me.cmbWarehouseTo))
        End If
        
        If Not Trim(strCriteria) = "" Then
            strCriteria = strCriteria & " AND "
        End If
        
        If mdlGlobal.objDatabaseInit = SQLSERVER7 Or mdlGlobal.objDatabaseInit = SQLSERVER2000 Or mdlGlobal.objDatabaseInit = SQLEXPRESS Then
            strCriteria = strCriteria & "MutDate>='" & mdlProcedures.FormatDate(Me.dtpStartDate.Value) & "' AND MutDate<='" & mdlProcedures.FormatDate(Me.dtpFinishDate.Value) & "'"
        ElseIf mdlGlobal.objDatabaseInit = MSACCESS Then
            strCriteria = strCriteria & "MutDate>=#" & mdlProcedures.FormatDate(Me.dtpStartDate.Value) & "# AND MutDate<=#" & mdlProcedures.FormatDate(Me.dtpFinishDate.Value) & "#"
        ElseIf mdlGlobal.objDatabaseInit = MYSQL Then
            strCriteria = strCriteria & "MutDate>='" & mdlProcedures.FormatDate(Me.dtpStartDate.Value) & "' AND MutDate<='" & mdlProcedures.FormatDate(Me.dtpFinishDate.Value) & "'"
        End If
    End If
    
    Dim rstTemp As ADODB.Recordset
    
    Set rstTemp = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "MutId, MutDate, WarehouseFrom, WarehouseTo", mdlTable.CreateTHMUTITEM, False, strCriteria, "MutDate ASC")
    
    SetVirtualRecordset
    
    With rstTemp
        While Not .EOF
            rstHeader.AddNew
            
            rstHeader!MutId = !MutId
            rstHeader!MutDate = mdlProcedures.FormatDate(!MutDate)
            rstHeader!WarehouseFrom = !WarehouseFrom
            rstHeader!WarehouseNameFrom = mdlDatabase.GetFieldData(mdlGlobal.conInventory, "Name", mdlTable.CreateTMWAREHOUSE, "WarehouseId='" & !WarehouseFrom & "'")
            rstHeader!WarehouseTo = !WarehouseTo
            rstHeader!WarehouseNameTo = mdlDatabase.GetFieldData(mdlGlobal.conInventory, "Name", mdlTable.CreateTMWAREHOUSE, "WarehouseId='" & !WarehouseTo & "'")
            
            rstHeader.Update
            
            .MoveNext
        Wend
    End With
    
    mdlDatabase.CloseRecordset rstTemp
    
    If rstHeader.RecordCount > 0 Then
        rstHeader.MoveFirst
        
        SetGridDetail rstHeader!MutId
    Else
        SetGridDetail
    End If
    
    Set Me.dgdHeader.DataSource = rstHeader
    
    With Me.dgdHeader
        .RowHeight = 1000
        
        .Columns(0).Width = 0
        .Columns(0).Locked = True
        .Columns(0).Visible = False
        .Columns(1).Width = 2400
        .Columns(1).Locked = True
        .Columns(1).Caption = "Tanggal"
        .Columns(1).NumberFormat = "dd MMMM yyyy"
        .Columns(2).Width = 1000
        .Columns(2).Locked = True
        .Columns(2).Caption = "Gudang Asal"
        .Columns(3).Width = 1950
        .Columns(3).Locked = True
        .Columns(3).Caption = "Nama"
        .Columns(3).WrapText = True
        .Columns(4).Width = 1000
        .Columns(4).Locked = True
        .Columns(4).Caption = "Gudang Tujuan"
        .Columns(5).Width = 1950
        .Columns(5).Locked = True
        .Columns(5).Caption = "Nama"
        .Columns(5).WrapText = True
    End With
End Sub

Private Sub SetVirtualRecordset()
    mdlDatabase.CloseRecordset rstHeader
    
    Set rstHeader = New ADODB.Recordset
    rstHeader.CursorLocation = adUseClient
    
    With rstHeader
        .Fields.Append "MutId", adChar, mdlDatabase.GetColumnSize(mdlGlobal.conInventory, "MutId", mdlTable.CreateTHMUTITEM)
        .Fields.Append "MutDate", adDate
        .Fields.Append "WarehouseFrom", adChar, mdlDatabase.GetColumnSize(mdlGlobal.conInventory, "WarehouseFrom", mdlTable.CreateTHMUTITEM)
        .Fields.Append "WarehouseNameFrom", adChar, mdlDatabase.GetColumnSize(mdlGlobal.conInventory, "Name", mdlTable.CreateTMWAREHOUSE)
        .Fields.Append "WarehouseTo", adChar, mdlDatabase.GetColumnSize(mdlGlobal.conInventory, "WarehouseTo", mdlTable.CreateTHMUTITEM)
        .Fields.Append "WarehouseNameTo", adChar, mdlDatabase.GetColumnSize(mdlGlobal.conInventory, "Name", mdlTable.CreateTMWAREHOUSE)
        
        .Open
    End With
End Sub

Private Sub SetGridDetail(Optional ByVal strMutId As String = "")
    Dim strTableFirst As String
    Dim strTableSecond As String
    Dim strTable As String
    
    strTableFirst = mdlTable.CreateTDMUTITEM
    strTableSecond = mdlTable.CreateTMITEM
    
    strTable = strTableFirst & " LEFT JOIN " & strTableSecond & " ON " & _
        strTableFirst & ".ItemId=" & strTableSecond & ".ItemId"
    
    Dim strCriteria As String
    
    strCriteria = "MutId='" & strMutId & "'"
    
    Set rstDetail = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, strTableFirst & ".ItemId, PartNumber, Name, Qty, UnityId", strTable, False, strCriteria, strTableFirst & ".ItemId ASC")
    
    Set Me.dgdDetail.DataSource = rstDetail
    
    With Me.dgdDetail
        .RowHeight = 730
        
        .Columns(0).Width = 1200
        .Columns(0).Locked = True
        .Columns(0).Caption = "Kode"
        .Columns(1).Width = 2300
        .Columns(1).Locked = True
        .Columns(1).Caption = "Nomor Part"
        .Columns(1).WrapText = True
        .Columns(2).Width = 2500
        .Columns(2).Locked = True
        .Columns(2).Caption = "Nama"
        .Columns(2).WrapText = True
        .Columns(3).Width = 1300
        .Columns(3).Locked = True
        .Columns(3).Caption = "Qty"
        .Columns(3).Alignment = dbgRight
        .Columns(3).NumberFormat = "#,##0"
        .Columns(4).Width = 1000
        .Columns(4).Locked = True
        .Columns(4).Caption = "Satuan"
    End With
End Sub
