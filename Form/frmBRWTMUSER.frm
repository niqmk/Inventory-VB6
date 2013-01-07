VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmBRWTMUSER 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5295
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   7815
   Icon            =   "frmBRWTMUSER.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5295
   ScaleWidth      =   7815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraSearch 
      Height          =   975
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   7575
      Begin VB.TextBox txtUserId 
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
         Left            =   1320
         TabIndex        =   0
         Top             =   240
         Width           =   2535
      End
      Begin VB.TextBox txtUserName 
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
         Left            =   1320
         TabIndex        =   1
         Top             =   600
         Width           =   6135
      End
      Begin VB.Label lblUserId 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Id Pemakai"
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
      Left            =   5160
      TabIndex        =   3
      Top             =   4800
      Width           =   1215
   End
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
      Left            =   6480
      TabIndex        =   2
      Top             =   4800
      Width           =   1215
   End
   Begin MSDataGridLib.DataGrid dgdMain 
      Height          =   3495
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   7575
      _ExtentX        =   13361
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
Attribute VB_Name = "frmBRWTMUSER"
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
    If frmAddUser.Parent Then
        frmAddUser.Parent = False
    End If
    
    mdlDatabase.CloseRecordset rstMain
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmBRWTMUSER = Nothing
End Sub

Private Sub txtUserId_GotFocus()
    mdlProcedures.GotFocus Me.txtUserId
End Sub

Private Sub txtUserId_Validate(Cancel As Boolean)
    SetGrid
End Sub

Private Sub txtUserName_GotFocus()
    mdlProcedures.GotFocus Me.txtUserName
End Sub

Private Sub txtUserName_Validate(Cancel As Boolean)
    SetGrid
End Sub

Private Sub dgdMain_HeadClick(ByVal ColIndex As Integer)
    rstMain.Sort = rstMain.Fields(ColIndex).Name
End Sub

Private Sub cmdChoose_Click()
    If rstMain.RecordCount > 0 Then
        If frmAddUser.Parent Then
            frmAddUser.UserId = Trim(rstMain!UserId)
        End If
    End If
    
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub SetInitialization()
    mdlProcedures.CenterWindows Me, False
    
    Me.Caption = mdlText.strBRWTMUSER

    If frmAddUser.Parent Then
        Me.txtUserId.Text = Trim(frmAddUser.UserId)
        Me.txtUserName.Text = Trim(frmAddUser.UserName)
    End If
        
    SetGrid
End Sub

Private Sub SetGrid()
    Dim strCriteria As String
    
    strCriteria = ""
    
    If Not Trim(Me.txtUserId.Text) = "" Then
        strCriteria = mdlProcedures.QueryLikeCriteria("UserId", mdlProcedures.RepDupText(Me.txtUserId.Text))
    End If
    
    If Not Trim(Me.txtUserName.Text) = "" Then
        If Not Trim(strCriteria) = "" Then
            strCriteria = strCriteria & " AND "
        End If
        
        strCriteria = strCriteria & mdlProcedures.QueryLikeCriteria("UserName", mdlProcedures.RepDupText(Me.txtUserName.Text))
    End If
    
    Set rstMain = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "UserId, UserName", mdlTable.CreateTMUSER, False, strCriteria, "UserId ASC")
    
    Set Me.dgdMain.DataSource = rstMain
    
    With Me.dgdMain
        .Columns(0).Width = 1500
        .Columns(0).Locked = True
        .Columns(0).Caption = "Id Pemakai"
        .Columns(1).Width = 5300
        .Columns(1).Locked = True
        .Columns(1).Caption = "Nama"
    End With
End Sub
