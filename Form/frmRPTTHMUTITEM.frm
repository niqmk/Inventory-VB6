VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmRPTTHMUTITEM 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   7830
   Icon            =   "frmRPTTHMUTITEM.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   7830
   ShowInTaskbar   =   0   'False
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
      Left            =   6600
      TabIndex        =   4
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Frame fraSearch 
      Height          =   2295
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   7575
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
         Left            =   120
         TabIndex        =   1
         Top             =   1320
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
         Left            =   120
         TabIndex        =   0
         Top             =   600
         Width           =   7335
      End
      Begin MSComCtl2.DTPicker dtpStartDate 
         Height          =   375
         Left            =   1440
         TabIndex        =   2
         Top             =   1800
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
         Format          =   44695555
         CurrentDate     =   39335
      End
      Begin MSComCtl2.DTPicker dtpFinishDate 
         Height          =   375
         Left            =   5880
         TabIndex        =   3
         Top             =   1800
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
         Format          =   44695555
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
         Top             =   960
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
         Top             =   1800
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
         Top             =   1800
         Width           =   1080
      End
   End
End
Attribute VB_Name = "frmRPTTHMUTITEM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private PrintReport As clsRPTTHMUTITEM

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub Form_Load()
    SetInitialization
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Not PrintReport Is Nothing Then
        Set PrintReport = Nothing
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmRPTTHMUTITEM = Nothing
End Sub

Private Sub cmdPrint_Click()
    SetPrint
End Sub

Private Sub SetInitialization()
    mdlProcedures.CenterWindows Me
    
    Me.Caption = mdlText.strRPTTHMUTITEM
    
    Me.dtpStartDate.CustomFormat = mdlGlobal.strFormatDate
    Me.dtpFinishDate.CustomFormat = mdlGlobal.strFormatDate
    
    FillCombo
    
    Dim strMonth As String
    Dim strYear As String
    
    strMonth = mdlProcedures.FormatDate(Now, "MM")
    strYear = mdlProcedures.FormatDate(Now, "yyyy")
    
    Me.dtpStartDate.Value = mdlProcedures.SetDate(strMonth, strYear)
    Me.dtpFinishDate.Value = mdlProcedures.SetDate(strMonth, strYear, , True)
End Sub

Private Sub FillCombo()
    Dim rstTemp As ADODB.Recordset
    
    Set rstTemp = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "WarehouseId, Name", mdlTable.CreateTMWAREHOUSE, False)
    
    mdlProcedures.FillComboData Me.cmbWarehouseFrom, rstTemp
    mdlProcedures.FillComboData Me.cmbWarehouseTo, rstTemp
    
    mdlDatabase.CloseRecordset rstTemp
End Sub

Private Sub SetPrint()
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
    
    Dim rstMain As ADODB.Recordset
    
    Set rstMain = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "MutId, MutDate, WarehouseFrom, WarehouseTo", mdlTable.CreateTHMUTITEM, False, strCriteria, "MutDate ASC")
    
    If Not rstMain.RecordCount > 0 Then
        MsgBox "Data Tidak Ada", vbOKOnly + vbExclamation, Me.Caption
    Else
        If Not PrintReport Is Nothing Then
            Set PrintReport = Nothing
        End If
        
        Set PrintReport = New clsRPTTHMUTITEM
        
        PrintReport.ImportToExcel rstMain
    End If
    
    mdlDatabase.CloseRecordset rstMain
End Sub
