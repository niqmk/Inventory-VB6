VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmRPTSUMTHPOSELL 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3915
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9165
   Icon            =   "frmRPTSUMTHPOSELL.frx":0000
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3915
   ScaleWidth      =   9165
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraDetail 
      Height          =   735
      Left            =   120
      TabIndex        =   11
      Top             =   2280
      Width           =   8895
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
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame fraSearch 
      Height          =   2175
      Index           =   0
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   8895
      Begin VB.OptionButton optSearch 
         Caption         =   "Berdasarkan Bulan ini"
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
         Index           =   1
         Left            =   6360
         TabIndex        =   2
         Top             =   840
         Width           =   2295
      End
      Begin VB.OptionButton optSearch 
         Caption         =   "Berdasarkan Tanggal"
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
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   840
         Width           =   2295
      End
      Begin VB.Frame fraSearch 
         Caption         =   "Berdasarkan Tanggal"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Index           =   1
         Left            =   2640
         TabIndex        =   12
         Top             =   720
         Width           =   3375
         Begin MSComCtl2.DTPicker dtpStartDate 
            Height          =   375
            Left            =   1680
            TabIndex        =   4
            Top             =   360
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
            Format          =   45416451
            CurrentDate     =   39335
         End
         Begin MSComCtl2.DTPicker dtpFinishDate 
            Height          =   375
            Left            =   1680
            TabIndex        =   5
            Top             =   840
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
            Format          =   45416451
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
            TabIndex        =   8
            Top             =   360
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
            Left            =   120
            TabIndex        =   9
            Top             =   840
            Width           =   1395
         End
      End
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
         Left            =   1320
         TabIndex        =   0
         Top             =   240
         Width           =   7455
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
         TabIndex        =   7
         Top             =   240
         Width           =   840
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
      Left            =   7920
      TabIndex        =   6
      Top             =   3120
      Width           =   1095
   End
End
Attribute VB_Name = "frmRPTSUMTHPOSELL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private PrintReport As clsRPTSUMTHPOSELL

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
    Set frmRPTSUMTHPOSELL = Nothing
End Sub

Private Sub optSearch_Click(Index As Integer)
    Select Case Index
        Case 0:
            Me.fraSearch(1).Enabled = True
            
            Me.dtpStartDate.SetFocus
        Case 1:
            Me.fraSearch(1).Enabled = False
    End Select
End Sub

Private Sub cmdPrint_Click()
    SetPrint
End Sub

Private Sub SetInitialization()
    mdlProcedures.CenterWindows Me
    
    Me.Caption = mdlText.strRPTSUMTHPOSELL
    
    Me.dtpStartDate.CustomFormat = mdlGlobal.strFormatDate
    Me.dtpFinishDate.CustomFormat = mdlGlobal.strFormatDate
    
    FillCombo
    
    Dim strMonth As String
    Dim strYear As String
    
    strMonth = mdlProcedures.FormatDate(Now, "MM")
    strYear = mdlProcedures.FormatDate(Now, "yyyy")
    
    Me.dtpStartDate.Value = mdlProcedures.SetDate(strMonth, strYear)
    Me.dtpFinishDate.Value = mdlProcedures.SetDate(strMonth, strYear, , True)
    
    Me.fraSearch(1).Enabled = False
End Sub

Private Sub FillCombo()
    Dim rstTemp As ADODB.Recordset
    
    Set rstTemp = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "CustomerId, Name", mdlTable.CreateTMCUSTOMER, False, , "CustomerId ASC")
    
    mdlProcedures.FillComboData Me.cmbCustomerId, rstTemp
    
    mdlDatabase.CloseRecordset rstTemp
End Sub

Private Sub SetPrint()
    Dim strCriteria As String
    
    strCriteria = ""
    
    If mdlProcedures.IsValidComboData(Me.cmbCustomerId) Then
        strCriteria = "CustomerId='" & mdlProcedures.GetComboData(Me.cmbCustomerId) & "'"
    End If
    
    If Me.optSearch(0).Value Then
        If Not Trim(strCriteria) = "" Then
            strCriteria = strCriteria & " AND "
        End If
        
        If mdlGlobal.objDatabaseInit = SQLSERVER7 Or mdlGlobal.objDatabaseInit = SQLSERVER2000 Or mdlGlobal.objDatabaseInit = SQLEXPRESS Then
            strCriteria = strCriteria & "PODate>='" & mdlProcedures.FormatDate(Me.dtpStartDate.Value) & "' AND PODate<='" & mdlProcedures.FormatDate(Me.dtpFinishDate.Value) & "'"
        ElseIf mdlGlobal.objDatabaseInit = MSACCESS Then
            strCriteria = strCriteria & "PODate>=#" & mdlProcedures.FormatDate(Me.dtpStartDate.Value) & "# AND PODate<=#" & mdlProcedures.FormatDate(Me.dtpFinishDate.Value) & "#"
        ElseIf mdlGlobal.objDatabaseInit = MYSQL Then
            strCriteria = strCriteria & "PODate>='" & mdlProcedures.FormatDate(Me.dtpStartDate.Value) & "' AND PODate<='" & mdlProcedures.FormatDate(Me.dtpFinishDate.Value) & "'"
        End If
    ElseIf Me.optSearch(1).Value Then
        If Not Trim(strCriteria) = "" Then
            strCriteria = strCriteria & " AND "
        End If
    
        If mdlGlobal.objDatabaseInit = SQLSERVER7 Or mdlGlobal.objDatabaseInit = SQLSERVER2000 Or mdlGlobal.objDatabaseInit = SQLEXPRESS Then
            strCriteria = strCriteria & "PODate>='" & mdlProcedures.FormatDate(mdlProcedures.SetDate(Month(Now), Year(Now))) & "' AND PODate<='" & mdlProcedures.FormatDate(mdlProcedures.SetDate(Month(Now), Year(Now), , True)) & "'"
        ElseIf mdlGlobal.objDatabaseInit = MSACCESS Then
            strCriteria = strCriteria & "PODate>=#" & mdlProcedures.FormatDate(mdlProcedures.SetDate(Month(Now), Year(Now))) & "# AND PODate<=#" & mdlProcedures.FormatDate(mdlProcedures.SetDate(Month(Now), Year(Now), , True)) & "#"
        ElseIf mdlGlobal.objDatabaseInit = MYSQL Then
            strCriteria = strCriteria & "PODate>='" & mdlProcedures.FormatDate(mdlProcedures.SetDate(Month(Now), Year(Now))) & "' AND PODate<='" & mdlProcedures.FormatDate(mdlProcedures.SetDate(Month(Now), Year(Now), , True)) & "'"
        End If
    End If
    
    Dim rstMain As ADODB.Recordset
    
    Set rstMain = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "*", mdlTable.CreateTHPOSELL, False, strCriteria, "POId ASC")
    
    If Not rstMain.RecordCount > 0 Then
        MsgBox "Data Tidak Ada", vbOKOnly + vbExclamation, Me.Caption
    Else
        If Not PrintReport Is Nothing Then
            Set PrintReport = Nothing
        End If
        
        Set PrintReport = New clsRPTSUMTHPOSELL
        
        PrintReport.ImportToExcel rstMain, Me.chkDetail.Value
    End If
    
    mdlDatabase.CloseRecordset rstMain
End Sub
