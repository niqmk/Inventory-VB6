VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmMISTMWAREHOUSE 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5205
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   7470
   Icon            =   "frmMISTMWAREHOUSE.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5205
   ScaleWidth      =   7470
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraSearch 
      Height          =   1455
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   7215
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
         Left            =   6000
         TabIndex        =   2
         Top             =   960
         Width           =   1095
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
         TabIndex        =   3
         Top             =   240
         Width           =   450
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
         TabIndex        =   4
         Top             =   600
         Width           =   510
      End
   End
   Begin MSDataGridLib.DataGrid dgdMain 
      Height          =   3375
      Left            =   120
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1680
      Width           =   7215
      _ExtentX        =   12726
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
Attribute VB_Name = "frmMISTMWAREHOUSE"
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
    mdlDatabase.CloseRecordset rstMain
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmMISTMWAREHOUSE = Nothing
End Sub

Private Sub txtWarehouseId_GotFocus()
    mdlProcedures.GotFocus Me.txtWarehouseId
End Sub

Private Sub txtName_GotFocus()
    mdlProcedures.GotFocus Me.txtName
End Sub

Private Sub cmdSearch_Click()
    SetGrid False
End Sub

Private Sub dgdMain_HeadClick(ByVal ColIndex As Integer)
    rstMain.Sort = rstMain.Fields(ColIndex).Name
End Sub

Private Sub SetInitialization()
    mdlProcedures.CenterWindows Me
    
    Me.Caption = mdlText.strMISTMWAREHOUSE
    
    SetGrid
End Sub

Private Sub SetGrid(Optional ByVal blnInitialize As Boolean = True)
    Dim strCriteria As String
    
    If blnInitialize Then
        strCriteria = "WarehouseId=''"
    Else
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
    End If
    
    Set rstMain = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "WarehouseId, Name, WarehouseSet", mdlTable.CreateTMWAREHOUSE, False, strCriteria, "WarehouseId ASC")
    
    Set Me.dgdMain.DataSource = rstMain
    
    With Me.dgdMain
        .RowHeight = 500
        
        .Columns(0).Width = 800
        .Columns(0).Locked = True
        .Columns(0).Caption = "Kode"
        .Columns(1).Width = 5200
        .Columns(1).Locked = True
        .Columns(1).Caption = "Nama"
        .Columns(1).WrapText = True
        .Columns(2).Width = 500
        .Columns(2).Locked = True
        .Columns(2).Caption = "SET"
        .Columns(2).Alignment = dbgCenter
    End With
End Sub
