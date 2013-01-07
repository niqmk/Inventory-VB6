VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmBRWTMWAREHOUSE 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5190
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8670
   Icon            =   "frmBRWTMWAREHOUSE.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5190
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
      Begin VB.TextBox txtWarehouseId 
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
         Width           =   735
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
      Begin VB.Label lblWarehouseId 
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
Attribute VB_Name = "frmBRWTMWAREHOUSE"
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
    If frmTMWAREHOUSE.Parent Then
        frmTMWAREHOUSE.Parent = False
    End If
    
    If frmTHITEMOUT.Parent Then
        frmTHITEMOUT.Parent = False
    End If
    
    If frmTHITEMIN.Parent Then
        frmTHITEMIN.Parent = False
    End If
    
    If frmTHMUTITEM.Parent Then
        frmTHMUTITEM.Parent = False
    End If
    
    If frmTHDOBUY.Parent Then
        frmTHDOBUY.Parent = False
    End If
    
    If frmTHRTRSELL.Parent Then
        frmTHRTRSELL.Parent = False
    End If
    
    mdlDatabase.CloseRecordset rstMain
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmBRWTMWAREHOUSE = Nothing
End Sub

Private Sub txtWarehouseId_GotFocus()
    mdlProcedures.GotFocus Me.txtWarehouseId
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
        If frmTMWAREHOUSE.Parent Then
            frmTMWAREHOUSE.WarehouseId = rstMain!WarehouseId
        ElseIf frmTHITEMIN.Parent Then
            frmTHITEMIN.WarehouseIdCombo = rstMain!WarehouseId
        ElseIf frmTHITEMOUT.Parent Then
            frmTHITEMOUT.WarehouseIdCombo = rstMain!WarehouseId
        ElseIf frmTHMUTITEM.Parent Then
            If frmTHMUTITEM.FromBoolean Then
                frmTHMUTITEM.WarehouseFromCombo = rstMain!WarehouseId
            ElseIf frmTHMUTITEM.ToBoolean Then
                frmTHMUTITEM.WarehouseToCombo = rstMain!WarehouseId
            End If
        ElseIf frmTHDOBUY.Parent Then
            frmTHDOBUY.WarehouseIdCombo = rstMain!WarehouseId
        ElseIf frmTHRTRSELL.Parent Then
            frmTHRTRSELL.WarehouseIdCombo = rstMain!WarehouseId
        End If
    End If
End Sub

Private Sub dgdMain_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If rstMain.RecordCount > 0 Then
        If frmTMWAREHOUSE.Parent Then
            frmTMWAREHOUSE.WarehouseId = rstMain!WarehouseId
        ElseIf frmTHITEMIN.Parent Then
            frmTHITEMIN.WarehouseIdCombo = rstMain!WarehouseId
        ElseIf frmTHITEMOUT.Parent Then
            frmTHITEMOUT.WarehouseIdCombo = rstMain!WarehouseId
        ElseIf frmTHMUTITEM.Parent Then
            If frmTHMUTITEM.FromBoolean Then
                frmTHMUTITEM.WarehouseFromCombo = rstMain!WarehouseId
            ElseIf frmTHMUTITEM.ToBoolean Then
                frmTHMUTITEM.WarehouseToCombo = rstMain!WarehouseId
            End If
        ElseIf frmTHDOBUY.Parent Then
            frmTHDOBUY.WarehouseIdCombo = rstMain!WarehouseId
        ElseIf frmTHRTRSELL.Parent Then
            frmTHRTRSELL.WarehouseIdCombo = rstMain!WarehouseId
        End If
    End If
End Sub

Private Sub cmdChoose_Click()
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    If frmTMWAREHOUSE.Parent Then
        frmTMWAREHOUSE.WarehouseId = strParent
    ElseIf frmTHITEMIN.Parent Then
        frmTHITEMIN.WarehouseIdCombo = strParent
    ElseIf frmTHITEMOUT.Parent Then
        frmTHITEMOUT.WarehouseIdCombo = strParent
    ElseIf frmTHMUTITEM.Parent Then
        If frmTHMUTITEM.FromBoolean Then
            frmTHMUTITEM.WarehouseFromCombo = strParent
        ElseIf frmTHMUTITEM.ToBoolean Then
            frmTHMUTITEM.WarehouseToCombo = strParent
        End If
    ElseIf frmTHDOBUY.Parent Then
        frmTHDOBUY.WarehouseIdCombo = strParent
    ElseIf frmTHRTRSELL.Parent Then
        frmTHRTRSELL.WarehouseIdCombo = strParent
    End If
    
    Unload Me
End Sub

Private Sub SetInitialization()
    mdlProcedures.CenterWindows Me, False
    
    Me.Caption = mdlText.strBRWTMWAREHOUSE

    If frmTMWAREHOUSE.Parent Then
        Me.txtWarehouseId.Text = Trim(frmTMWAREHOUSE.WarehouseId)
        Me.txtName.Text = Trim(frmTMWAREHOUSE.WarehouseName)
        
        strParent = Trim(frmTMWAREHOUSE.WarehouseId)
    ElseIf frmTHITEMIN.Parent Then
        strParent = Trim(frmTHITEMIN.WarehouseIdCombo)
    ElseIf frmTHITEMOUT.Parent Then
        strParent = Trim(frmTHITEMOUT.WarehouseIdCombo)
    ElseIf frmTHMUTITEM.Parent Then
        If frmTHMUTITEM.FromBoolean Then
            strParent = Trim(frmTHMUTITEM.WarehouseFromCombo)
        ElseIf frmTHMUTITEM.ToBoolean Then
            strParent = Trim(frmTHMUTITEM.WarehouseToCombo)
        End If
    ElseIf frmTHDOBUY.Parent Then
        strParent = Trim(frmTHDOBUY.WarehouseIdCombo)
    ElseIf frmTHRTRSELL.Parent Then
        strParent = Trim(frmTHRTRSELL.WarehouseIdCombo)
    End If
    
    SetGrid
End Sub

Private Sub SetGrid()
    Dim strCriteria As String
    
    strCriteria = ""
    
    If Not Trim(Me.txtWarehouseId.Text) = "" Then
        strCriteria = mdlProcedures.QueryLikeCriteria("WarehouseId", mdlProcedures.RepDupText(Me.txtWarehouseId.Text))
    End If
    
    If Not Trim(Me.txtName.Text) = "" Then
        If Not Trim(strCriteria) = "" Then
            strCriteria = strCriteria & " AND "
        End If
        
        strCriteria = strCriteria & mdlProcedures.QueryLikeCriteria("Name", mdlProcedures.RepDupText(Me.txtName.Text))
    End If
    
    Set rstMain = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "WarehouseId, Name, WarehouseSet", mdlTable.CreateTMWAREHOUSE, False, strCriteria, "WarehouseId ASC")
    
    If rstMain.RecordCount > 0 Then
        If frmTMWAREHOUSE.Parent Then
            frmTMWAREHOUSE.WarehouseId = rstMain!WarehouseId
        ElseIf frmTHITEMIN.Parent Then
            frmTHITEMIN.WarehouseIdCombo = rstMain!WarehouseId
        ElseIf frmTHITEMOUT.Parent Then
            frmTHITEMOUT.WarehouseIdCombo = rstMain!WarehouseId
        ElseIf frmTHMUTITEM.Parent Then
            If frmTHMUTITEM.FromBoolean Then
                frmTHMUTITEM.WarehouseFromCombo = rstMain!WarehouseId
            ElseIf frmTHMUTITEM.ToBoolean Then
                frmTHMUTITEM.WarehouseToCombo = rstMain!WarehouseId
            End If
        ElseIf frmTHDOBUY.Parent Then
            frmTHDOBUY.WarehouseIdCombo = rstMain!WarehouseId
        ElseIf frmTHRTRSELL.Parent Then
            frmTHRTRSELL.WarehouseIdCombo = rstMain!WarehouseId
        End If
    End If
    
    Set Me.dgdMain.DataSource = rstMain
    
    With Me.dgdMain
        .Columns(0).Width = 800
        .Columns(0).Locked = True
        .Columns(0).Caption = "Kode"
        .Columns(1).Width = 6200
        .Columns(1).Locked = True
        .Columns(1).Caption = "Nama"
        .Columns(1).WrapText = True
        .Columns(2).Width = 700
        .Columns(2).Locked = True
        .Columns(2).Caption = "SET"
        .Columns(2).Alignment = dbgCenter
    End With
End Sub
