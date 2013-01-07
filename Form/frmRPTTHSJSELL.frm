VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmRPTTHSJSELL 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2655
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   9270
   Icon            =   "frmRPTTHSJSELL.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2655
   ScaleWidth      =   9270
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
      Left            =   8040
      TabIndex        =   5
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Frame fraSearch 
      Height          =   1935
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   9015
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
         Left            =   7320
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
         TabIndex        =   7
         Top             =   600
         Width           =   915
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
         TabIndex        =   8
         Top             =   960
         Width           =   840
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
         Left            =   5640
         TabIndex        =   10
         Top             =   1440
         Width           =   1395
      End
   End
End
Attribute VB_Name = "frmRPTTHSJSELL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private PrintReport As clsRPTTHSJSELL

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
    Set frmRPTTHSJSELL = Nothing
End Sub

Private Sub cmdPrint_Click()
    SetPrint
End Sub

Private Sub txtSJId_GotFocus()
    mdlProcedures.GotFocus Me.txtSJId
End Sub

Private Sub txtSOId_GotFocus()
    mdlProcedures.GotFocus Me.txtSOId
End Sub

Private Sub SetInitialization()
    mdlProcedures.CenterWindows Me
    
    Me.Caption = mdlText.strRPTTHSJSELL
    
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
    
    Set rstTemp = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "CustomerId, Name", mdlTable.CreateTMCUSTOMER, False)
    
    mdlProcedures.FillComboData Me.cmbCustomerId, rstTemp
    
    mdlDatabase.CloseRecordset rstTemp
End Sub

Private Sub SetPrint()
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
    
    Dim rstMain As ADODB.Recordset
    
    Set rstMain = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "*", mdlTable.CreateTHSJSELL, False, strCriteria, "SJId ASC")
    
    If Not rstMain.RecordCount > 0 Then
        MsgBox "Data Tidak Ada", vbOKOnly + vbExclamation, Me.Caption
    Else
        If Not PrintReport Is Nothing Then
            Set PrintReport = Nothing
        End If
        
        Set PrintReport = New clsRPTTHSJSELL
        
        PrintReport.ImportToExcel rstMain
    End If
    
    mdlDatabase.CloseRecordset rstMain
End Sub
