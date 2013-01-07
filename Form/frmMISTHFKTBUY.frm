VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmMISTHFKTBUY 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7575
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9375
   Icon            =   "frmMISTHFKTBUY.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7575
   ScaleWidth      =   9375
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraSearch 
      Height          =   1575
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   9135
      Begin VB.TextBox txtFktId 
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
         Left            =   7920
         TabIndex        =   4
         Top             =   1080
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
         Left            =   1560
         TabIndex        =   1
         Top             =   600
         Width           =   7455
      End
      Begin MSComCtl2.DTPicker dtpStartDate 
         Height          =   375
         Left            =   1560
         TabIndex        =   2
         Top             =   1080
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
         Left            =   6120
         TabIndex        =   3
         Top             =   1080
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
      Begin VB.Label lblFktId 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nomor Faktur"
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
         Width           =   1185
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
         TabIndex        =   6
         Top             =   600
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
         TabIndex        =   7
         Top             =   1080
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
         TabIndex        =   8
         Top             =   1080
         Width           =   1395
      End
   End
   Begin MSDataGridLib.DataGrid dgdHeader 
      Height          =   2775
      Left            =   120
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1800
      Width           =   9135
      _ExtentX        =   16113
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
   Begin MSDataGridLib.DataGrid dgdDetail 
      Height          =   2775
      Left            =   120
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   4680
      Width           =   9135
      _ExtentX        =   16113
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
Attribute VB_Name = "frmMISTHFKTBUY"
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
    Set frmMISTHFKTBUY = Nothing
End Sub

Private Sub txtFktId_GotFocus()
    mdlProcedures.GotFocus Me.txtFktId
End Sub

Private Sub cmdSearch_Click()
    SetGridHeader False
End Sub

Private Sub dgdHeader_HeadClick(ByVal ColIndex As Integer)
    rstHeader.Sort = rstHeader.Fields(ColIndex).Name
    
    If rstHeader.RecordCount > 0 Then
        SetGridDetail rstHeader!FktId
    Else
        SetGridDetail
    End If
End Sub

Private Sub dgdDetail_HeadClick(ByVal ColIndex As Integer)
    rstDetail.Sort = rstDetail.Fields(ColIndex).Name
End Sub

Private Sub dgdHeader_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If rstHeader.RecordCount > 0 Then
        SetGridDetail rstHeader!FktId
    Else
        SetGridDetail
    End If
End Sub

Private Sub SetInitialization()
    mdlProcedures.CenterWindows Me
    
    Me.Caption = mdlText.strMISTHFKTBUY
    
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
    Dim strTableFirst As String
    Dim strTableSecond As String
    Dim strTable As String
    
    strTableFirst = mdlTable.CreateTHFKTBUY
    strTableSecond = mdlTable.CreateTMVENDOR
    
    strTable = strTableFirst & " LEFT JOIN " & strTableSecond & _
        " ON " & strTableFirst & ".VendorId=" & strTableSecond & ".VendorId"
    
    Dim strCriteria As String
    
    If blnInitialize Then
        strCriteria = "FktId=''"
    Else
        strCriteria = ""
        
        If Not Trim(Me.txtFktId.Text) = "" Then
            strCriteria = mdlProcedures.QueryLikeCriteria("FktId", mdlProcedures.RepDupText(Me.txtFktId.Text))
        End If
        
        If mdlProcedures.IsValidComboData(Me.cmbVendorId) Then
            If Not Trim(strCriteria) = "" Then
                strCriteria = strCriteria & " AND "
            End If
            
            strCriteria = strCriteria & strTableFirst & ".VendorId='" & mdlProcedures.GetComboData(Me.cmbVendorId) & "'"
        End If
        
        If Not Trim(strCriteria) = "" Then
            strCriteria = strCriteria & " AND "
        End If
        
        If mdlGlobal.objDatabaseInit = SQLSERVER7 Or mdlGlobal.objDatabaseInit = SQLSERVER2000 Or mdlGlobal.objDatabaseInit = SQLEXPRESS Then
            strCriteria = strCriteria & "FktDate>='" & mdlProcedures.FormatDate(Me.dtpStartDate.Value) & "' AND FktDate<='" & mdlProcedures.FormatDate(Me.dtpFinishDate.Value) & "'"
        ElseIf mdlGlobal.objDatabaseInit = MSACCESS Then
            strCriteria = strCriteria & "FktDate>=#" & mdlProcedures.FormatDate(Me.dtpStartDate.Value) & "# AND FktDate<=#" & mdlProcedures.FormatDate(Me.dtpFinishDate.Value) & "#"
        ElseIf mdlGlobal.objDatabaseInit = MYSQL Then
            strCriteria = strCriteria & "FktDate>='" & mdlProcedures.FormatDate(Me.dtpStartDate.Value) & "' AND FktDate<='" & mdlProcedures.FormatDate(Me.dtpFinishDate.Value) & "'"
        End If
    End If
    
    Set rstHeader = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "FktId, FktDate, Name", strTable, False, strCriteria, "FktId ASC")
    
    If rstHeader.RecordCount > 0 Then
        SetGridDetail rstHeader!FktId
    Else
        SetGridDetail
    End If
    
    Set Me.dgdHeader.DataSource = rstHeader
    
    With Me.dgdHeader
        .RowHeight = 500
        
        .Columns(0).Width = 2500
        .Columns(0).Locked = True
        .Columns(0).Caption = "Nomor"
        .Columns(1).Width = 2500
        .Columns(1).Locked = True
        .Columns(1).Caption = "Tanggal"
        .Columns(1).NumberFormat = "dd MMMM yyyy"
        .Columns(2).Width = 3500
        .Columns(2).Locked = True
        .Columns(2).Caption = "Pemasok"
        .Columns(2).WrapText = True
    End With
End Sub

Private Sub SetGridDetail(Optional ByVal strFktId As String = "")
    Dim strCriteria As String
    
    strCriteria = "FktId='" & strFktId & "'"
    
    Dim rstTemp As ADODB.Recordset
    
    Set rstTemp = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "SJId", mdlTable.CreateTDFKTBUY, False, strCriteria, "SJId ASC")
    
    SetVirtualRecordset
    
    With rstTemp
        While Not .EOF
            rstDetail.AddNew
            
            rstDetail!SJId = !SJId
            rstDetail!SJDate = mdlDatabase.GetFieldData(mdlGlobal.conInventory, "SJDate", mdlTable.CreateTHSJBUY, "SJId='" & !SJId & "'")
            rstDetail!Qty = mdlTHSJBUY.GetTotalQtySJBUY(!SJId)
            
            rstDetail.Update
            
            .MoveNext
        Wend
    End With
    
    mdlDatabase.CloseRecordset rstTemp
    
    Set Me.dgdDetail.DataSource = rstDetail
    
    With Me.dgdDetail
        .Columns(0).Width = 2700
        .Columns(0).Locked = True
        .Columns(0).Caption = "Nomor SJ"
        .Columns(1).Width = 2800
        .Columns(1).Locked = True
        .Columns(1).Caption = "Tanggal"
        .Columns(1).NumberFormat = "dd MMMM yyyy"
        .Columns(2).Width = 3000
        .Columns(2).Locked = True
        .Columns(2).Caption = "Qty. SJ"
        .Columns(2).NumberFormat = "#,##0"
        .Columns(2).Alignment = dbgRight
    End With
End Sub

Private Sub SetVirtualRecordset()
    mdlDatabase.CloseRecordset rstDetail
    
    Set rstDetail = New ADODB.Recordset
    rstDetail.CursorLocation = adUseClient
    
    With rstDetail
        .Fields.Append "SJId", adChar, mdlDatabase.GetColumnSize(mdlGlobal.conInventory, "SJId", mdlTable.CreateTHSJBUY)
        .Fields.Append "SJDate", adDate
        .Fields.Append "Qty", adCurrency
        
        .Open
    End With
End Sub
