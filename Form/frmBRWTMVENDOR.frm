VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmBRWTMVENDOR 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5310
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8790
   Icon            =   "frmBRWTMVENDOR.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5310
   ScaleWidth      =   8790
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
      Left            =   7440
      TabIndex        =   3
      Top             =   4800
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
      Left            =   6120
      TabIndex        =   4
      Top             =   4800
      Width           =   1215
   End
   Begin VB.Frame fraSearch 
      Height          =   975
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   8535
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
         Left            =   7320
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
      Begin VB.TextBox txtVendorId 
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
      Begin VB.Label lblVendorId 
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
      Height          =   3495
      Left            =   120
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1200
      Width           =   8535
      _ExtentX        =   15055
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
Attribute VB_Name = "frmBRWTMVENDOR"
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
    If frmTMVENDOR.Parent Then
        frmTMVENDOR.Parent = False
    End If
    
    If frmTMITEM.Parent Then
        frmTMITEM.Parent = False
    End If
    
    If frmTHPOBUY.Parent Then
        frmTHPOBUY.Parent = False
    End If
    
    If frmTHDOBUY.Parent Then
        frmTHDOBUY.Parent = False
    End If
    
    If frmTHSJBUY.Parent Then
        frmTHSJBUY.Parent = False
    End If
    
    If frmTHFKTBUY.Parent Then
        frmTHFKTBUY.Parent = False
    End If
    
    If frmTHRTRBUY.Parent Then
        frmTHRTRBUY.Parent = False
    End If
    
    mdlDatabase.CloseRecordset rstMain
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmBRWTMVENDOR = Nothing
End Sub

Private Sub txtVendorId_GotFocus()
    mdlProcedures.GotFocus Me.txtVendorId
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
        If frmTMVENDOR.Parent Then
            frmTMVENDOR.VendorId = rstMain!VendorId
        ElseIf frmTMITEM.Parent Then
            frmTMITEM.VendorIdCombo = rstMain!VendorId
        ElseIf frmTHPOBUY.Parent Then
            frmTHPOBUY.VendorIdCombo = rstMain!VendorId
        ElseIf frmTHDOBUY.Parent Then
            frmTHDOBUY.VendorIdCombo = rstMain!VendorId
        ElseIf frmTHSJBUY.Parent Then
            frmTHSJBUY.VendorIdCombo = rstMain!VendorId
        ElseIf frmTHFKTBUY.Parent Then
            frmTHFKTBUY.VendorIdCombo = rstMain!VendorId
        ElseIf frmTHRTRBUY.Parent Then
            frmTHRTRBUY.VendorIdCombo = rstMain!VendorId
        End If
    End If
End Sub

Private Sub dgdMain_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If rstMain.RecordCount > 0 Then
        If frmTMVENDOR.Parent Then
            frmTMVENDOR.VendorId = rstMain!VendorId
        ElseIf frmTMITEM.Parent Then
            frmTMITEM.VendorIdCombo = rstMain!VendorId
        ElseIf frmTHPOBUY.Parent Then
            frmTHPOBUY.VendorIdCombo = rstMain!VendorId
        ElseIf frmTHDOBUY.Parent Then
            frmTHDOBUY.VendorIdCombo = rstMain!VendorId
        ElseIf frmTHSJBUY.Parent Then
            frmTHSJBUY.VendorIdCombo = rstMain!VendorId
        ElseIf frmTHFKTBUY.Parent Then
            frmTHFKTBUY.VendorIdCombo = rstMain!VendorId
        ElseIf frmTHRTRBUY.Parent Then
            frmTHRTRBUY.VendorIdCombo = rstMain!VendorId
        End If
    End If
End Sub

Private Sub cmdChoose_Click()
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    If frmTMVENDOR.Parent Then
        frmTMVENDOR.VendorId = strParent
    ElseIf frmTMITEM.Parent Then
        frmTMITEM.VendorIdCombo = strParent
    ElseIf frmTHPOBUY.Parent Then
        frmTHPOBUY.VendorIdCombo = strParent
    ElseIf frmTHDOBUY.Parent Then
        frmTHDOBUY.VendorIdCombo = strParent
    ElseIf frmTHSJBUY.Parent Then
        frmTHSJBUY.VendorIdCombo = strParent
    ElseIf frmTHFKTBUY.Parent Then
        frmTHFKTBUY.VendorIdCombo = strParent
    ElseIf frmTHRTRBUY.Parent Then
        frmTHRTRBUY.VendorIdCombo = strParent
    End If
    
    Unload Me
End Sub

Private Sub SetInitialization()
    mdlProcedures.CenterWindows Me, False
    
    Me.Caption = mdlText.strBRWTMVENDOR

    If frmTMVENDOR.Parent Then
        Me.txtVendorId.Text = Trim(frmTMVENDOR.VendorId)
        Me.txtName.Text = Trim(frmTMVENDOR.VendorName)
        
        strParent = Trim(frmTMVENDOR.VendorId)
    ElseIf frmTMITEM.Parent Then
        strParent = Trim(frmTMITEM.VendorIdCombo)
    ElseIf frmTHPOBUY.Parent Then
        strParent = Trim(frmTHPOBUY.VendorIdCombo)
    ElseIf frmTHDOBUY.Parent Then
        strParent = Trim(frmTHDOBUY.VendorIdCombo)
    ElseIf frmTHSJBUY.Parent Then
        strParent = Trim(frmTHSJBUY.VendorIdCombo)
    ElseIf frmTHFKTBUY.Parent Then
        strParent = Trim(frmTHFKTBUY.VendorIdCombo)
    ElseIf frmTHRTRBUY.Parent Then
        strParent = Trim(frmTHRTRBUY.VendorIdCombo)
    End If
    
    SetGrid
End Sub

Private Sub SetGrid()
    Dim strCriteria As String
    
    strCriteria = ""
    
    If Not Trim(Me.txtVendorId.Text) = "" Then
        strCriteria = mdlProcedures.QueryLikeCriteria("VendorId", mdlProcedures.RepDupText(Me.txtVendorId.Text))
    End If
    
    If Not Trim(Me.txtName.Text) = "" Then
        If Not Trim(strCriteria) = "" Then
            strCriteria = strCriteria & " AND "
        End If
        
        strCriteria = strCriteria & mdlProcedures.QueryLikeCriteria("Name", mdlProcedures.RepDupText(Me.txtName.Text))
    End If
    
    Set rstMain = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "VendorId, Name", mdlTable.CreateTMVENDOR, False, strCriteria, "VendorId ASC")
    
    If rstMain.RecordCount > 0 Then
        If frmTMVENDOR.Parent Then
            frmTMVENDOR.VendorId = rstMain!VendorId
        ElseIf frmTMITEM.Parent Then
            frmTMITEM.VendorIdCombo = rstMain!VendorId
        ElseIf frmTHPOBUY.Parent Then
            frmTHPOBUY.VendorIdCombo = rstMain!VendorId
        ElseIf frmTHDOBUY.Parent Then
            frmTHDOBUY.VendorIdCombo = rstMain!VendorId
        ElseIf frmTHSJBUY.Parent Then
            frmTHSJBUY.VendorIdCombo = rstMain!VendorId
        ElseIf frmTHFKTBUY.Parent Then
            frmTHFKTBUY.VendorIdCombo = rstMain!VendorId
        ElseIf frmTHRTRBUY.Parent Then
            frmTHRTRBUY.VendorIdCombo = rstMain!VendorId
        End If
    End If
    
    Set Me.dgdMain.DataSource = rstMain
    
    With Me.dgdMain
        .Columns(0).Width = 800
        .Columns(0).Locked = True
        .Columns(0).Caption = "Kode"
        .Columns(1).Width = 7000
        .Columns(1).Locked = True
        .Columns(1).Caption = "Nama"
    End With
End Sub
