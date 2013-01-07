VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmBRWTMEMPLOYEE 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5310
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8670
   Icon            =   "frmBRWTMEMPLOYEE.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5310
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
      Left            =   6000
      TabIndex        =   4
      Top             =   4800
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
      Begin VB.TextBox txtEmployeeId 
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
         Width           =   975
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
      Begin VB.Label lblEmployeeId 
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
      Width           =   8415
      _ExtentX        =   14843
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
Attribute VB_Name = "frmBRWTMEMPLOYEE"
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
    If frmTMEMPLOYEE.Parent Then
        frmTMEMPLOYEE.Parent = False
    End If
    
    If frmTMWAREHOUSE.Parent Then
        frmTMWAREHOUSE.Parent = False
    End If
    
    If frmTHPOBUY.Parent Then
        frmTHPOBUY.Parent = False
    End If
    
    mdlDatabase.CloseRecordset rstMain
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmBRWTMEMPLOYEE = Nothing
End Sub

Private Sub txtEmployeeId_GotFocus()
    mdlProcedures.GotFocus Me.txtEmployeeId
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
        If frmTMEMPLOYEE.Parent Then
            frmTMEMPLOYEE.EmployeeId = rstMain!EmployeeId
        ElseIf frmTMWAREHOUSE.Parent Then
            frmTMWAREHOUSE.EmployeeIdCombo = rstMain!EmployeeId
        ElseIf frmTHPOBUY.Parent Then
            If frmTHPOBUY.ByBoolean Then
                frmTHPOBUY.EmployeeByCombo = rstMain!EmployeeId
            ElseIf frmTHPOBUY.AgreeBoolean Then
                frmTHPOBUY.EmployeeAgreeCombo = rstMain!EmployeeId
            End If
        End If
    End If
End Sub

Private Sub dgdMain_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If rstMain.RecordCount > 0 Then
        If frmTMEMPLOYEE.Parent Then
            frmTMEMPLOYEE.EmployeeId = rstMain!EmployeeId
        ElseIf frmTMWAREHOUSE.Parent Then
            frmTMWAREHOUSE.EmployeeIdCombo = rstMain!EmployeeId
        ElseIf frmTHPOBUY.Parent Then
            If frmTHPOBUY.ByBoolean Then
                frmTHPOBUY.EmployeeByCombo = rstMain!EmployeeId
            ElseIf frmTHPOBUY.AgreeBoolean Then
                frmTHPOBUY.EmployeeAgreeCombo = rstMain!EmployeeId
            End If
        End If
    End If
End Sub

Private Sub cmdChoose_Click()
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    If frmTMEMPLOYEE.Parent Then
        frmTMEMPLOYEE.EmployeeId = strParent
    ElseIf frmTMWAREHOUSE.Parent Then
        frmTMWAREHOUSE.EmployeeIdCombo = strParent
    ElseIf frmTHPOBUY.Parent Then
        If frmTHPOBUY.ByBoolean Then
            frmTHPOBUY.EmployeeByCombo = strParent
        ElseIf frmTHPOBUY.AgreeBoolean Then
            frmTHPOBUY.EmployeeAgreeCombo = strParent
        End If
    End If
    
    Unload Me
End Sub

Private Sub SetInitialization()
    mdlProcedures.CenterWindows Me, False
    
    Me.Caption = mdlText.strBRWTMEMPLOYEE

    If frmTMEMPLOYEE.Parent Then
        Me.txtEmployeeId.Text = Trim(frmTMEMPLOYEE.EmployeeId)
        Me.txtName.Text = Trim(frmTMEMPLOYEE.EmployeeName)
        
        strParent = Trim(frmTMEMPLOYEE.EmployeeId)
    ElseIf frmTMWAREHOUSE.Parent Then
        strParent = Trim(frmTMWAREHOUSE.EmployeeIdCombo)
    ElseIf frmTHPOBUY.Parent Then
        If frmTHPOBUY.ByBoolean Then
            strParent = Trim(frmTHPOBUY.EmployeeByCombo)
        ElseIf frmTHPOBUY.AgreeBoolean Then
            strParent = Trim(frmTHPOBUY.EmployeeAgreeCombo)
        End If
    End If
    
    SetGrid
End Sub

Private Sub SetGrid()
    Dim strCriteria As String
    
    strCriteria = ""
    
    If Not Trim(Me.txtEmployeeId.Text) = "" Then
        strCriteria = mdlProcedures.QueryLikeCriteria("EmployeeId", mdlProcedures.RepDupText(Me.txtEmployeeId.Text))
    End If
    
    If Not Trim(Me.txtName.Text) = "" Then
        If Not Trim(strCriteria) = "" Then
            strCriteria = strCriteria & " AND "
        End If
        
        strCriteria = strCriteria & mdlProcedures.QueryLikeCriteria("Name", mdlProcedures.RepDupText(Me.txtName.Text))
    End If
    
    Set rstMain = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "EmployeeId, Name", mdlTable.CreateTMEMPLOYEE, False, strCriteria, "EmployeeId ASC")
    
    If rstMain.RecordCount > 0 Then
        If frmTMEMPLOYEE.Parent Then
            frmTMEMPLOYEE.EmployeeId = rstMain!EmployeeId
        ElseIf frmTMWAREHOUSE.Parent Then
            frmTMWAREHOUSE.EmployeeIdCombo = rstMain!EmployeeId
        ElseIf frmTHPOBUY.Parent Then
            If frmTHPOBUY.ByBoolean Then
                frmTHPOBUY.EmployeeByCombo = rstMain!EmployeeId
            ElseIf frmTHPOBUY.AgreeBoolean Then
                frmTHPOBUY.EmployeeAgreeCombo = rstMain!EmployeeId
            End If
        End If
    End If
    
    Set Me.dgdMain.DataSource = rstMain
    
    With Me.dgdMain
        .Columns(0).Width = 900
        .Columns(0).Locked = True
        .Columns(0).Caption = "Kode"
        .Columns(1).Width = 6800
        .Columns(1).Locked = True
        .Columns(1).Caption = "Nama"
    End With
End Sub
