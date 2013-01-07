VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmTDFKTBUY 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6615
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7950
   Icon            =   "frmTDFKTBUY.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6615
   ScaleWidth      =   7950
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraSearch 
      Height          =   1455
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   7695
      Begin VB.TextBox txtPOId 
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
         Left            =   6480
         TabIndex        =   4
         Top             =   960
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker dtpStartDate 
         Height          =   375
         Left            =   1440
         TabIndex        =   2
         Top             =   960
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
         Left            =   4800
         TabIndex        =   3
         Top             =   960
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
      Begin VB.Label lblPOId 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nomor PO"
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
         TabIndex        =   6
         Top             =   240
         Width           =   855
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
         Left            =   3240
         TabIndex        =   9
         Top             =   960
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
         TabIndex        =   8
         Top             =   960
         Width           =   1080
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
      Left            =   6600
      TabIndex        =   5
      Top             =   6120
      Width           =   1215
   End
   Begin MSDataGridLib.DataGrid dgdMain 
      Height          =   4335
      Left            =   120
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   1680
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   7646
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
Attribute VB_Name = "frmTDFKTBUY"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private rstMain As ADODB.Recordset

Private Sub Form_Load()
    SetInitialization
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If frmTHFKTBUY.Parent Then
        frmTHFKTBUY.Parent = False
    End If
    
    mdlDatabase.CloseRecordset rstMain
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmTDFKTBUY = Nothing
End Sub

Private Sub txtSJId_GotFocus()
    mdlProcedures.GotFocus Me.txtSJId
End Sub

Private Sub cmdSearch_Click()
    SetGrid False
End Sub

Private Sub cmdSave_Click()
    If rstMain.RecordCount > 0 Then
        If mdlProcedures.GetCurrency(rstMain!Qty) > 0 Then
            frmTHFKTBUY.SaveDetail rstMain!SJId
        Else
            Me.dgdMain.SetFocus
        End If
    End If
End Sub

Private Sub SetInitialization()
    mdlProcedures.CenterWindows Me, False
    
    Me.Caption = mdlText.strTDFKTBUY
    
    Me.dtpStartDate.CustomFormat = mdlGlobal.strFormatDate
    Me.dtpFinishDate.CustomFormat = mdlGlobal.strFormatDate
    
    Dim strMonth As String
    Dim strYear As String
    
    If frmTHFKTBUY.Parent Then
        strMonth = frmTHFKTBUY.MonthTrans
        strYear = frmTHFKTBUY.YearTrans
    End If
    
    Me.dtpStartDate.Value = mdlProcedures.SetDate(strMonth, strYear)
    Me.dtpFinishDate.Value = mdlProcedures.SetDate(strMonth, strYear, , True)
    
    SetGrid False
End Sub

Private Sub SetGrid(Optional ByVal blnInitialize As Boolean = True)
    Dim strTableFirst As String
    Dim strTableSecond As String
    Dim strTableThird As String
    Dim strTable As String
    
    strTableFirst = mdlTable.CreateTHSJBUY
    strTableSecond = mdlTable.CreateTHDOBUY
    strTableThird = mdlTable.CreateTHPOBUY
    
    strTable = "(" & strTableFirst & " LEFT JOIN " & strTableSecond & _
        " ON " & strTableFirst & ".DOId=" & strTableSecond & ".DOId) LEFT JOIN " & strTableThird & _
        " ON " & strTableSecond & ".POId=" & strTableThird & ".POId"
    
    Dim strCriteria As String
    
    strCriteria = ""
    
    If blnInitialize Then
        strCriteria = "SJId=''"
    Else
        If Not Trim(Me.txtSJId.Text) = "" Then
            strCriteria = mdlProcedures.QueryLikeCriteria("SJId", mdlProcedures.RepDupText(Me.txtSJId.Text))
        End If
        
        If Not Trim(Me.txtPOId.Text) = "" Then
            If Not Trim(strCriteria) = "" Then
                strCriteria = strCriteria & " AND "
            End If
            
            strCriteria = strCriteria & mdlProcedures.QueryLikeCriteria("POId", mdlProcedures.RepDupText(Me.txtPOId.Text))
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
    End If
    
    If Not Trim(strCriteria) = "" Then
        strCriteria = strCriteria & " AND "
    End If
    
    strCriteria = strCriteria & "VendorId='" & frmTHFKTBUY.VendorIdCombo & "'"
    strCriteria = strCriteria & " AND " & strTableFirst & ".SJId NOT IN (SELECT SJId FROM " & mdlTable.CreateTDFKTBUY & ")"
    
    Dim rstTemp As ADODB.Recordset
    
    Set rstTemp = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "SJId, SJDate", strTable, False, strCriteria, "SJId ASC")
    
    SetVirtualRecordset
    
    With rstTemp
        While Not .EOF
            rstMain.AddNew
            
            rstMain!SJId = !SJId
            rstMain!SJDate = !SJDate
            rstMain!Qty = mdlTHSJBUY.GetTotalQtySJBUY(!SJId)
            
            rstMain.Update
            
            .MoveNext
        Wend
    End With
    
    mdlDatabase.CloseRecordset rstTemp
    
    If rstMain.RecordCount > 0 Then
        rstMain.MoveFirst
    End If
    
    Set Me.dgdMain.DataSource = rstMain
    
    With Me.dgdMain
        .Columns(0).Width = 2500
        .Columns(0).Locked = True
        .Columns(0).Caption = "Nomor SJ"
        .Columns(1).Width = 2500
        .Columns(1).Locked = True
        .Columns(1).Caption = "Tanggal"
        .Columns(1).NumberFormat = "dd MMMM yyyy"
        .Columns(2).Width = 2000
        .Columns(2).Locked = True
        .Columns(2).Caption = "Qty. SJ"
        .Columns(2).NumberFormat = "#,##0"
        .Columns(2).Alignment = dbgRight
    End With
End Sub

Private Sub SetVirtualRecordset()
    mdlDatabase.CloseRecordset rstMain
    
    Set rstMain = New ADODB.Recordset
    rstMain.CursorLocation = adUseClient
    
    With rstMain
        .Fields.Append "SJId", adChar, mdlDatabase.GetColumnSize(mdlGlobal.conInventory, "SJId", mdlTable.CreateTHSJBUY)
        .Fields.Append "SJDate", adDate
        .Fields.Append "Qty", adCurrency
        
        .Open
    End With
End Sub
