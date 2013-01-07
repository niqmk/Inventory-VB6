VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmMISTHRTRBUY 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7605
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9270
   Icon            =   "frmMISTHRTRBUY.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7605
   ScaleWidth      =   9270
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraSearch 
      Height          =   1935
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   9015
      Begin VB.TextBox txtRtrId 
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
         Left            =   1440
         TabIndex        =   0
         Top             =   240
         Width           =   2535
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
         TabIndex        =   5
         Top             =   1440
         Width           =   1095
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
         Left            =   1440
         TabIndex        =   2
         Top             =   960
         Width           =   7455
      End
      Begin VB.TextBox txtSJId 
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
         Left            =   1440
         TabIndex        =   1
         Top             =   600
         Width           =   2535
      End
      Begin MSComCtl2.DTPicker dtpStartDate 
         Height          =   375
         Left            =   1440
         TabIndex        =   3
         Top             =   1440
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
         Left            =   6000
         TabIndex        =   4
         Top             =   1440
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
      Begin VB.Label lblRtrId 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nomor Retur"
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
         Top             =   240
         Width           =   1095
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
         TabIndex        =   8
         Top             =   960
         Width           =   825
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
         Top             =   1440
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
         Left            =   4440
         TabIndex        =   10
         Top             =   1440
         Width           =   1395
      End
      Begin VB.Label lblSJId 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nomor SJ"
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
         Top             =   600
         Width           =   855
      End
   End
   Begin MSDataGridLib.DataGrid dgdHeader 
      Height          =   2415
      Left            =   120
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   2160
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   4260
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
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   4680
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
End
Attribute VB_Name = "frmMISTHRTRBUY"
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
    Set frmMISTHRTRBUY = Nothing
End Sub

Private Sub txtRtrId_GotFocus()
    mdlProcedures.GotFocus Me.txtRtrId
End Sub

Private Sub txtSJId_GotFocus()
    mdlProcedures.GotFocus Me.txtSJId
End Sub

Private Sub cmdSearch_Click()
    SetGridHeader False
End Sub

Private Sub dgdHeader_HeadClick(ByVal ColIndex As Integer)
    rstHeader.Sort = rstHeader.Fields(ColIndex).Name
    
    If rstHeader.RecordCount > 0 Then
        SetGridDetail rstHeader!RtrId
    Else
        SetGridDetail
    End If
End Sub

Private Sub dgdDetail_HeadClick(ByVal ColIndex As Integer)
    rstDetail.Sort = rstDetail.Fields(ColIndex).Name
End Sub

Private Sub dgdHeader_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If rstHeader.RecordCount > 0 Then
        SetGridDetail rstHeader!RtrId
    Else
        SetGridDetail
    End If
End Sub

Private Sub SetInitialization()
    mdlProcedures.CenterWindows Me
    
    Me.Caption = mdlText.strMISTHRTRBUY
    
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
    
    Set rstTemp = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "VendorId, Name", mdlTable.CreateTMVENDOR, False, , "VendorId ASC")
    
    mdlProcedures.FillComboData Me.cmbVendorId, rstTemp
    
    mdlDatabase.CloseRecordset rstTemp
End Sub

Private Sub SetGridHeader(Optional ByVal blnInitialize As Boolean = True)
    Dim strCriteria As String
    
    If blnInitialize Then
        strCriteria = "RtrId=''"
    Else
        strCriteria = ""
        
        If Not Trim(Me.txtRtrId.Text) = "" Then
            strCriteria = mdlProcedures.QueryLikeCriteria("RtrId", mdlProcedures.RepDupText(Me.txtRtrId.Text))
        End If
        
        If Not Trim(Me.txtSJId.Text) = "" Then
            If Not Trim(strCriteria) = "" Then
                strCriteria = strCriteria & " AND "
            End If
            
            strCriteria = strCriteria & mdlProcedures.QueryLikeCriteria("SJId", mdlProcedures.RepDupText(Me.txtSJId.Text))
        End If
        
        If mdlProcedures.IsValidComboData(Me.cmbVendorId) Then
            If Not Trim(strCriteria) = "" Then
                strCriteria = strCriteria & " AND "
            End If
            
            strCriteria = strCriteria & "SJId IN (SELECT SJId FROM " & mdlTable.CreateTHSJBUY & " WHERE DOId IN (SELECT DOId FROM " & mdlTable.CreateTHDOBUY & " WHERE POId IN (SELECT POId FROM " & mdlTable.CreateTHPOBUY & " WHERE VendorId='" & mdlProcedures.GetComboData(Me.cmbVendorId) & "')))"
        End If
        
        If Not Trim(strCriteria) = "" Then
            strCriteria = strCriteria & " AND "
        End If
        
        If mdlGlobal.objDatabaseInit = SQLSERVER7 Or mdlGlobal.objDatabaseInit = SQLSERVER2000 Or mdlGlobal.objDatabaseInit = SQLEXPRESS Then
            strCriteria = strCriteria & "RtrDate>='" & mdlProcedures.FormatDate(Me.dtpStartDate.Value) & "' AND RtrDate<='" & mdlProcedures.FormatDate(Me.dtpFinishDate.Value) & "'"
        ElseIf mdlGlobal.objDatabaseInit = MSACCESS Then
            strCriteria = strCriteria & "RtrDate>=#" & mdlProcedures.FormatDate(Me.dtpStartDate.Value) & "# AND RtrDate<=#" & mdlProcedures.FormatDate(Me.dtpFinishDate.Value) & "#"
        ElseIf mdlGlobal.objDatabaseInit = MYSQL Then
            strCriteria = strCriteria & "RtrDate>='" & mdlProcedures.FormatDate(Me.dtpStartDate.Value) & "' AND RtrDate<='" & mdlProcedures.FormatDate(Me.dtpFinishDate.Value) & "'"
        End If
    End If
    
    Set rstHeader = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "RtrId, RtrDate, SJId", mdlTable.CreateTHRTRBUY, False, strCriteria, "RtrId")
    
    If rstHeader.RecordCount > 0 Then
        SetGridDetail rstHeader!RtrId
    Else
        SetGridDetail
    End If
    
    Set Me.dgdHeader.DataSource = rstHeader
    
    With Me.dgdHeader
        .Columns(0).Width = 2950
        .Columns(0).Locked = True
        .Columns(0).Caption = "Nomor Retur"
        .Columns(1).Width = 2400
        .Columns(1).Locked = True
        .Columns(1).Caption = "Tanggal"
        .Columns(1).NumberFormat = "dd MMMM yyyy"
        .Columns(2).Width = 2950
        .Columns(2).Locked = True
        .Columns(2).Caption = "Nomor SJ"
    End With
End Sub

Private Sub SetGridDetail(Optional ByVal strRtrId As String = "")
    Dim strTableFirst As String
    Dim strTableSecond As String
    Dim strTable As String
    
    strTableFirst = mdlTable.CreateTDRTRBUY
    strTableSecond = mdlTable.CreateTMITEM
    
    strTable = strTableFirst & " LEFT JOIN " & strTableSecond & _
        " ON " & strTableFirst & ".ItemId=" & strTableSecond & ".ItemId"
    
    Dim strCriteria As String
    
    strCriteria = "RtrId='" & strRtrId & "'"
    
    Set rstDetail = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, strTableFirst & ".ItemId, Name, Qty, UnityId", strTable, False, strCriteria, strTableFirst & ".ItemId ASC")
    
    Set Me.dgdDetail.DataSource = rstDetail
    
    With Me.dgdDetail
        .RowHeight = 500
        
        .Columns(0).Width = 1100
        .Columns(0).Locked = True
        .Columns(0).Caption = "Kode"
        .Columns(1).Width = 5000
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
