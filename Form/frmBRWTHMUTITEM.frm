VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmBRWTHMUTITEM 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5070
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   9270
   Icon            =   "frmBRWTHMUTITEM.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5070
   ScaleWidth      =   9270
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraSearch 
      Height          =   1695
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   9015
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
         TabIndex        =   9
         Top             =   1200
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
         Left            =   4080
         TabIndex        =   10
         Top             =   1200
         Width           =   1395
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
         TabIndex        =   7
         Top             =   240
         Width           =   1125
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
         TabIndex        =   8
         Top             =   720
         Width           =   1305
      End
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
      Left            =   6600
      TabIndex        =   6
      Top             =   4560
      Width           =   1215
   End
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
      Left            =   7920
      TabIndex        =   5
      Top             =   4560
      Width           =   1215
   End
   Begin MSDataGridLib.DataGrid dgdMain 
      Height          =   2535
      Left            =   120
      TabIndex        =   11
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
Attribute VB_Name = "frmBRWTHMUTITEM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private rstMain As ADODB.Recordset

Private strParent As String

Private Sub Form_Load()
    SetInitialization
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If frmTHMUTITEM.Parent Then
        frmTHMUTITEM.Parent = False
    End If
    
    mdlDatabase.CloseRecordset rstMain
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmBRWTHMUTITEM = Nothing
End Sub

Private Sub dgdMain_HeadClick(ByVal ColIndex As Integer)
    rstMain.Sort = rstMain.Fields(ColIndex).Name
    
    If rstMain.RecordCount > 0 Then
        If frmTHMUTITEM.Parent Then
            frmTHMUTITEM.MutId = rstMain!MutId
        End If
    End If
End Sub

Private Sub dgdMain_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If rstMain.RecordCount > 0 Then
        If frmTHMUTITEM.Parent Then
            frmTHMUTITEM.MutId = rstMain!MutId
        End If
    End If
End Sub

Private Sub cmdSearch_Click()
    SetGrid
End Sub

Private Sub cmdChoose_Click()
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    If frmTHMUTITEM.Parent Then
        frmTHMUTITEM.MutId = strParent
    End If
    
    Unload Me
End Sub

Private Sub SetInitialization()
    mdlProcedures.CenterWindows Me, False
    
    Me.Caption = mdlText.strBRWTHMUTITEM
    
    Me.dtpStartDate.CustomFormat = mdlGlobal.strFormatDate
    Me.dtpFinishDate.CustomFormat = mdlGlobal.strFormatDate
    
    FillCombo
    
    If frmTHMUTITEM.Parent Then
        mdlProcedures.SetComboData Me.cmbWarehouseFrom, frmTHMUTITEM.WarehouseFromCombo
        mdlProcedures.SetComboData Me.cmbWarehouseTo, frmTHMUTITEM.WarehouseToCombo
        
        Dim strYear As String
        Dim strMonth As String
        
        strYear = frmTHMUTITEM.YearTrans
        strMonth = frmTHMUTITEM.MonthTrans
        
        Me.dtpStartDate.Value = mdlProcedures.SetDate(strMonth, strYear)
        Me.dtpFinishDate.Value = mdlProcedures.SetDate(strMonth, strYear, , True)
        
        strParent = Trim(frmTHMUTITEM.MutId)
    End If
    
    SetGrid
End Sub

Private Sub FillCombo()
    Dim rstTemp As ADODB.Recordset
    
    Set rstTemp = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "WarehouseId, Name", mdlTable.CreateTMWAREHOUSE, False)
    
    mdlProcedures.FillComboData Me.cmbWarehouseFrom, rstTemp
    mdlProcedures.FillComboData Me.cmbWarehouseTo, rstTemp
    
    mdlDatabase.CloseRecordset rstTemp
End Sub

Private Sub SetGrid()
    Dim strCriteria As String
    
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
    
    If frmTHMUTITEM.Parent Then
        If Not Trim(strCriteria) = "" Then
            strCriteria = strCriteria & " AND "
        End If
        
        strCriteria = strCriteria & "MONTH(MutDate)=" & frmTHMUTITEM.MonthTrans & " AND YEAR(MutDate)=" & frmTHMUTITEM.YearTrans
    End If
    
    Dim rstTemp As ADODB.Recordset
    
    Set rstTemp = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "MutId, MutDate, WarehouseFrom, WarehouseTo", mdlTable.CreateTHMUTITEM, False, strCriteria, "MutDate ASC")
    
    SetVirtualRecordset
    
    With rstTemp
        While Not .EOF
            rstMain.AddNew
            
            rstMain!MutId = !MutId
            rstMain!MutDate = mdlProcedures.FormatDate(!MutDate)
            rstMain!WarehouseFrom = !WarehouseFrom
            rstMain!WarehouseNameFrom = mdlDatabase.GetFieldData(mdlGlobal.conInventory, "Name", mdlTable.CreateTMWAREHOUSE, "WarehouseId='" & !WarehouseFrom & "'")
            rstMain!WarehouseTo = !WarehouseTo
            rstMain!WarehouseNameTo = mdlDatabase.GetFieldData(mdlGlobal.conInventory, "Name", mdlTable.CreateTMWAREHOUSE, "WarehouseId='" & !WarehouseTo & "'")
            
            rstMain.Update
            
            .MoveNext
        Wend
    End With
    
    mdlDatabase.CloseRecordset rstTemp
    
    If rstMain.RecordCount > 0 Then
        rstMain.MoveFirst
        
        If frmTHMUTITEM.Parent Then
            frmTHMUTITEM.MutId = rstMain!MutId
        End If
    End If
    
    Set Me.dgdMain.DataSource = rstMain
    
    With Me.dgdMain
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
    mdlDatabase.CloseRecordset rstMain
    
    Set rstMain = New ADODB.Recordset
    rstMain.CursorLocation = adUseClient
    
    With rstMain
        .Fields.Append "MutId", adChar, mdlDatabase.GetColumnSize(mdlGlobal.conInventory, "MutId", mdlTable.CreateTHMUTITEM)
        .Fields.Append "MutDate", adDate
        .Fields.Append "WarehouseFrom", adChar, mdlDatabase.GetColumnSize(mdlGlobal.conInventory, "WarehouseFrom", mdlTable.CreateTHMUTITEM)
        .Fields.Append "WarehouseNameFrom", adChar, mdlDatabase.GetColumnSize(mdlGlobal.conInventory, "Name", mdlTable.CreateTMWAREHOUSE)
        .Fields.Append "WarehouseTo", adChar, mdlDatabase.GetColumnSize(mdlGlobal.conInventory, "WarehouseTo", mdlTable.CreateTHMUTITEM)
        .Fields.Append "WarehouseNameTo", adChar, mdlDatabase.GetColumnSize(mdlGlobal.conInventory, "Name", mdlTable.CreateTMWAREHOUSE)
        
        .Open
    End With
End Sub
