VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmBRWTHSJSELL 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6270
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9285
   Icon            =   "frmBRWTHSJSELL.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6270
   ScaleWidth      =   9285
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
      Left            =   7920
      TabIndex        =   6
      Top             =   5760
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
      Left            =   6720
      TabIndex        =   7
      Top             =   5760
      Width           =   1095
   End
   Begin VB.Frame fraSearch 
      Height          =   1935
      Left            =   120
      TabIndex        =   14
      Top             =   120
      Width           =   9015
      Begin VB.ComboBox cmbCustomerId 
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
      Begin VB.TextBox txtSOId 
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
         TabIndex        =   0
         Top             =   240
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
         Format          =   82313219
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
         Format          =   82313219
         CurrentDate     =   39335
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
         TabIndex        =   12
         Top             =   1440
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
         TabIndex        =   11
         Top             =   1440
         Width           =   1080
      End
      Begin VB.Label lblCustomerId 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Customer"
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
         Top             =   960
         Width           =   840
      End
      Begin VB.Label lblSOId 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nomor SO"
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
         Top             =   600
         Width           =   915
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
         TabIndex        =   8
         Top             =   240
         Width           =   855
      End
   End
   Begin MSDataGridLib.DataGrid dgdMain 
      Height          =   3495
      Left            =   120
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   2160
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   6165
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
Attribute VB_Name = "frmBRWTHSJSELL"
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
    If frmTHSJSELL.Parent Then
        frmTHSJSELL.Parent = False
    End If
    
    If frmTHRTRSELL.Parent Then
        frmTHRTRSELL.Parent = False
    End If
    
    mdlDatabase.CloseRecordset rstMain
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmBRWTHSJSELL = Nothing
End Sub

Private Sub txtSJId_GotFocus()
    mdlProcedures.GotFocus Me.txtSJId
End Sub

Private Sub txtSOId_GotFocus()
    mdlProcedures.GotFocus Me.txtSOId
End Sub

Private Sub cmdSearch_Click()
    SetGrid
End Sub

Private Sub dgdMain_HeadClick(ByVal ColIndex As Integer)
    rstMain.Sort = rstMain.Fields(ColIndex).Name
    
    If rstMain.RecordCount > 0 Then
        If frmTHSJSELL.Parent Then
            frmTHSJSELL.SJId = rstMain!SJId
        ElseIf frmTHRTRSELL.Parent Then
            frmTHRTRSELL.SJIdCombo = rstMain!SJId
        End If
    End If
End Sub

Private Sub dgdMain_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If rstMain.RecordCount > 0 Then
        If frmTHSJSELL.Parent Then
            frmTHSJSELL.SJId = rstMain!SJId
        ElseIf frmTHRTRSELL.Parent Then
            frmTHRTRSELL.SJIdCombo = rstMain!SJId
        End If
    End If
End Sub

Private Sub cmdChoose_Click()
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    If frmTHSJSELL.Parent Then
        frmTHSJSELL.SJId = strParent
    ElseIf frmTHRTRSELL.Parent Then
        frmTHRTRSELL.SJIdCombo = strParent
    End If
    
    Unload Me
End Sub

Private Sub SetInitialization()
    mdlProcedures.CenterWindows Me, False
    
    Me.Caption = mdlText.strBRWTHSJSELL
    
    Me.dtpStartDate.CustomFormat = mdlGlobal.strFormatDate
    Me.dtpFinishDate.CustomFormat = mdlGlobal.strFormatDate
    
    FillCombo
    
    Dim strYear As String
    Dim strMonth As String
    
    If frmTHSJSELL.Parent Then
        Me.txtSJId.Text = Trim(frmTHSJSELL.SJId)
        Me.txtSOId.Text = Trim(frmTHSJSELL.SOIdCombo)
        
        mdlProcedures.SetComboData Me.cmbCustomerId, frmTHSJSELL.CustomerIdCombo
        
        strYear = frmTHSJSELL.YearTrans
        strMonth = frmTHSJSELL.MonthTrans
        
        strParent = Trim(frmTHSJSELL.SJId)
    ElseIf frmTHRTRSELL.Parent Then
        mdlProcedures.SetComboData Me.cmbCustomerId, frmTHRTRSELL.CustomerIdCombo
        
        Me.cmbCustomerId.Enabled = False
    
        strYear = frmTHRTRSELL.YearTrans
        strMonth = frmTHRTRSELL.MonthTrans
        
        strParent = Trim(frmTHRTRSELL.SJIdCombo)
    End If
    
    Me.dtpStartDate.Value = mdlProcedures.SetDate(strMonth, strYear)
    Me.dtpFinishDate.Value = mdlProcedures.SetDate(strMonth, strYear, , True)
    
    SetGrid
End Sub

Private Sub FillCombo()
    Dim rstTemp As ADODB.Recordset
    
    Set rstTemp = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "CustomerId, Name", mdlTable.CreateTMCUSTOMER, False, , "CustomerId ASC")
    
    mdlProcedures.FillComboData Me.cmbCustomerId, rstTemp
    
    mdlDatabase.CloseRecordset rstTemp
End Sub

Private Sub SetGrid()
    Dim strCriteria As String
    
    strCriteria = ""
    
    If Not Trim(Me.txtSJId.Text) = "" Then
        strCriteria = mdlProcedures.QueryLikeCriteria("SJId", mdlProcedures.RepDupText(Me.txtSJId.Text))
    End If
    
    If Not Trim(Me.txtSOId.Text) = "" Then
        If Not Trim(strCriteria) = "" Then
            strCriteria = strCriteria & " AND "
        End If
        
        strCriteria = strCriteria & mdlProcedures.QueryLikeCriteria("SOId", mdlProcedures.RepDupText(Me.txtSOId.Text))
    End If
    
    If mdlProcedures.IsValidComboData(Me.cmbCustomerId) Then
        If Not Trim(strCriteria) = "" Then
            strCriteria = strCriteria & " AND "
        End If
        
        strCriteria = strCriteria & "SOId IN (SELECT SOId FROM " & mdlTable.CreateTHSOSELL & " WHERE POId IN (SELECT POId FROM " & mdlTable.CreateTHPOSELL & " WHERE CustomerId='" & mdlProcedures.GetComboData(Me.cmbCustomerId) & "'))"
    End If
    
    If frmTHSJSELL.Parent Then
        If Not Trim(strCriteria) = "" Then
            strCriteria = strCriteria & " AND "
        End If
        
        strCriteria = strCriteria & "MONTH(SJDate)=" & frmTHSJSELL.MonthTrans & " AND YEAR(SJDate)=" & frmTHSJSELL.YearTrans
    End If
    
    If Not Trim(strCriteria) = "" Then
        strCriteria = strCriteria & " AND "
    End If
    
    If mdlGlobal.objDatabaseInit = SQLSERVER7 Or mdlGlobal.objDatabaseInit = SQLSERVER2000 Or mdlGlobal.objDatabaseInit = SQLEXPRESS Then
        strCriteria = strCriteria & "SJDate>='" & mdlProcedures.FormatDate(Me.dtpStartDate.Value) & "' AND SJDate<='" & mdlProcedures.FormatDate(Me.dtpFinishDate.Value) & "'"
    ElseIf mdlGlobal.objDatabaseInit = MSACCESS Then
        strCriteria = strCriteria & "SJDate>=#" & mdlProcedures.FormatDate(Me.dtpStartDate.Value) & "# AND SJDate<=#" & mdlProcedures.FormatDate(Me.dtpFinishDate.Value) & "#"
    ElseIf mdlGlobal.objDatabaseInit = MYSQL Then
        strCriteria = strCriteria & "SJDate>='" & mdlProcedures.FormatDate(Me.dtpStartDate.Value) & "' AND SJDate<='" & mdlProcedures.FormatDate(Me.dtpFinishDate.Value) & "'"
    End If
    
    Set rstMain = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "SJId, SJDate, SOId", mdlTable.CreateTHSJSELL, False, strCriteria, "SJId ASC")
    
    If rstMain.RecordCount > 0 Then
        If frmTHSJSELL.Parent Then
            frmTHSJSELL.SJId = rstMain!SJId
        ElseIf frmTHRTRSELL.Parent Then
            frmTHRTRSELL.SJIdCombo = rstMain!SJId
        End If
    End If
    
    Set Me.dgdMain.DataSource = rstMain
    
    With Me.dgdMain
        .Columns(0).Width = 2950
        .Columns(0).Locked = True
        .Columns(0).Caption = "Nomor SJ"
        .Columns(1).Width = 2400
        .Columns(1).Locked = True
        .Columns(1).Caption = "Tanggal"
        .Columns(1).NumberFormat = "dd MMMM yyyy"
        .Columns(2).Width = 2950
        .Columns(2).Locked = True
        .Columns(2).Caption = "Nomor SO"
    End With
End Sub
