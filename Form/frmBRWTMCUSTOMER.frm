VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmBRWTMCUSTOMER 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5175
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8670
   Icon            =   "frmBRWTMCUSTOMER.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   8670
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
      Left            =   7320
      TabIndex        =   3
      Top             =   4680
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
      Left            =   6000
      TabIndex        =   4
      Top             =   4680
      Width           =   1215
   End
   Begin VB.Frame fraSearch 
      Height          =   975
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   8415
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
         Left            =   7200
         TabIndex        =   2
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox txtName 
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
         Left            =   960
         TabIndex        =   1
         Top             =   600
         Width           =   6135
      End
      Begin VB.TextBox txtCustomerId 
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
         Left            =   960
         TabIndex        =   0
         Top             =   240
         Width           =   855
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama"
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
         Width           =   510
      End
      Begin VB.Label lblCustomerId 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kode"
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
         Width           =   450
      End
   End
   Begin MSDataGridLib.DataGrid dgdMain 
      Height          =   3375
      Left            =   120
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1200
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   5953
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
Attribute VB_Name = "frmBRWTMCUSTOMER"
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
    If frmTMCUSTOMER.Parent Then
        frmTMCUSTOMER.Parent = False
    End If
    
    If frmTHPOSELL.Parent Then
        frmTHPOSELL.Parent = False
    End If
    
    If frmTHSOSELL.Parent Then
        frmTHSOSELL.Parent = False
    End If
    
    If frmTHSJSELL.Parent Then
        frmTHSJSELL.Parent = False
    End If
    
    If frmTHFKTSELL.Parent Then
        frmTHFKTSELL.Parent = False
    End If
    
    If frmTHRTRSELL.Parent Then
        frmTHRTRSELL.Parent = False
    End If
    
    If frmFax.Parent Then
        frmFax.Parent = False
    End If
    
    mdlDatabase.CloseRecordset rstMain
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmBRWTMCUSTOMER = Nothing
End Sub

Private Sub txtCustomerId_GotFocus()
    mdlProcedures.GotFocus Me.txtCustomerId
End Sub

Private Sub txtName_GotFocus()
    mdlProcedures.GotFocus Me.txtName
End Sub

Private Sub cmdSearch_Click()
    SetGrid
End Sub

Private Sub dgdMain_HeadClick(ByVal ColIndex As Integer)
    rstMain.Sort = rstMain.Fields(ColIndex).Name
    
    If rstMain.RecordCount > 0 Then
        If frmTMCUSTOMER.Parent Then
            frmTMCUSTOMER.CustomerId = rstMain!CustomerId
        ElseIf frmTHPOSELL.Parent Then
            frmTHPOSELL.CustomerIdCombo = rstMain!CustomerId
        ElseIf frmTHSOSELL.Parent Then
            frmTHSOSELL.CustomerIdCombo = rstMain!CustomerId
        ElseIf frmTHSJSELL.Parent Then
            frmTHSJSELL.CustomerIdCombo = rstMain!CustomerId
        ElseIf frmTHFKTSELL.Parent Then
            frmTHFKTSELL.CustomerIdCombo = rstMain!CustomerId
        ElseIf frmTHRTRSELL.Parent Then
            frmTHRTRSELL.CustomerIdCombo = rstMain!CustomerId
        End If
    End If
End Sub

Private Sub dgdMain_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If rstMain.RecordCount > 0 Then
        If frmTMCUSTOMER.Parent Then
            frmTMCUSTOMER.CustomerId = rstMain!CustomerId
        ElseIf frmTHPOSELL.Parent Then
            frmTHPOSELL.CustomerIdCombo = rstMain!CustomerId
        ElseIf frmTHSOSELL.Parent Then
            frmTHSOSELL.CustomerIdCombo = rstMain!CustomerId
        ElseIf frmTHSJSELL.Parent Then
            frmTHSJSELL.CustomerIdCombo = rstMain!CustomerId
        ElseIf frmTHFKTSELL.Parent Then
            frmTHFKTSELL.CustomerIdCombo = rstMain!CustomerId
        ElseIf frmTHRTRSELL.Parent Then
            frmTHRTRSELL.CustomerIdCombo = rstMain!CustomerId
        End If
    End If
End Sub

Private Sub cmdChoose_Click()
    If rstMain.RecordCount > 0 Then
        If frmFax.Parent Then
            frmFax.Fax = rstMain!Fax
        End If
    End If

    Unload Me
End Sub

Private Sub cmdCancel_Click()
    If frmTMCUSTOMER.Parent Then
        frmTMCUSTOMER.CustomerId = strParent
    ElseIf frmTHPOSELL.Parent Then
        frmTHPOSELL.CustomerIdCombo = strParent
    ElseIf frmTHSOSELL.Parent Then
        frmTHSOSELL.CustomerIdCombo = strParent
    ElseIf frmTHSJSELL.Parent Then
        frmTHSJSELL.CustomerIdCombo = strParent
    ElseIf frmTHFKTSELL.Parent Then
        frmTHFKTSELL.CustomerIdCombo = strParent
    ElseIf frmTHRTRSELL.Parent Then
        frmTHRTRSELL.CustomerIdCombo = strParent
    End If
    
    Unload Me
End Sub

Private Sub SetInitialization()
    mdlProcedures.CenterWindows Me, False
    
    Me.Caption = mdlText.strBRWTMCUSTOMER

    If frmTMCUSTOMER.Parent Then
        Me.txtCustomerId.Text = Trim(frmTMCUSTOMER.CustomerId)
        Me.txtName.Text = Trim(frmTMCUSTOMER.CustomerName)
        
        strParent = Trim(frmTMCUSTOMER.CustomerId)
    ElseIf frmTHPOSELL.Parent Then
        strParent = Trim(frmTHPOSELL.CustomerIdCombo)
    ElseIf frmTHSOSELL.Parent Then
        strParent = Trim(frmTHSOSELL.CustomerIdCombo)
    ElseIf frmTHSJSELL.Parent Then
        strParent = Trim(frmTHSJSELL.CustomerIdCombo)
    ElseIf frmTHFKTSELL.Parent Then
        strParent = Trim(frmTHFKTSELL.CustomerIdCombo)
    ElseIf frmTHRTRSELL.Parent Then
        strParent = Trim(frmTHRTRSELL.CustomerIdCombo)
    End If
    
    SetGrid
End Sub

Private Sub SetGrid()
    Dim strCriteria As String
    
    strCriteria = ""
    
    If Not Trim(Me.txtCustomerId.Text) = "" Then
        strCriteria = mdlProcedures.QueryLikeCriteria("CustomerId", mdlProcedures.RepDupText(Me.txtCustomerId.Text))
    End If
    
    If Not Trim(Me.txtName.Text) = "" Then
        If Not Trim(strCriteria) = "" Then
            strCriteria = strCriteria & " AND "
        End If
        
        strCriteria = strCriteria & mdlProcedures.QueryLikeCriteria("Name", mdlProcedures.RepDupText(Me.txtName.Text))
    End If
    
    If frmFax.Parent Then
        Set rstMain = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "CustomerId, Name, Fax", mdlTable.CreateTMCUSTOMER, False, strCriteria, "CustomerId ASC")
    Else
        Set rstMain = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "CustomerId, Name", mdlTable.CreateTMCUSTOMER, False, strCriteria, "CustomerId ASC")
    End If
    
    If rstMain.RecordCount > 0 Then
        If frmTMCUSTOMER.Parent Then
            frmTMCUSTOMER.CustomerId = rstMain!CustomerId
        ElseIf frmTHPOSELL.Parent Then
            frmTHPOSELL.CustomerIdCombo = rstMain!CustomerId
        ElseIf frmTHSOSELL.Parent Then
            frmTHSOSELL.CustomerIdCombo = rstMain!CustomerId
        ElseIf frmTHSJSELL.Parent Then
            frmTHSJSELL.CustomerIdCombo = rstMain!CustomerId
        ElseIf frmTHFKTSELL.Parent Then
            frmTHFKTSELL.CustomerIdCombo = rstMain!CustomerId
        ElseIf frmTHRTRSELL.Parent Then
            frmTHRTRSELL.CustomerIdCombo = rstMain!CustomerId
        End If
    End If
    
    Set Me.dgdMain.DataSource = rstMain
    
    With Me.dgdMain
        .Columns(0).Width = 800
        .Columns(0).Locked = True
        .Columns(0).Caption = "Kode"
        
        If frmFax.Parent Then
            .Columns(1).Width = 4000
            .Columns(1).Locked = True
            .Columns(1).Caption = "Nama"
            .Columns(2).Width = 2900
            .Columns(2).Locked = True
            .Columns(2).Caption = "Fax"
        Else
            .Columns(1).Width = 6900
            .Columns(1).Locked = True
            .Columns(1).Caption = "Nama"
        End If
    End With
End Sub
