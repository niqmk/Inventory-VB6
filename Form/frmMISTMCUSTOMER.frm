VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmMISTMCUSTOMER 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7710
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14790
   Icon            =   "frmMISTMCUSTOMER.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7710
   ScaleWidth      =   14790
   ShowInTaskbar   =   0   'False
   Begin MSDataGridLib.DataGrid dgdDetail 
      Height          =   3735
      Left            =   120
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   3840
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   6588
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
   Begin VB.Frame fraSearch 
      Height          =   975
      Left            =   120
      TabIndex        =   9
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
   Begin MSDataGridLib.DataGrid dgdHeader 
      Height          =   2535
      Left            =   120
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1200
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   4471
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
   Begin MSDataGridLib.DataGrid dgdHeaderNotes 
      Height          =   3495
      Left            =   8640
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   240
      Width           =   6015
      _ExtentX        =   10610
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
   Begin MSDataGridLib.DataGrid dgdDetailNotes 
      Height          =   3735
      Left            =   8640
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   3840
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   6588
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
Attribute VB_Name = "frmMISTMCUSTOMER"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private rstHeader As ADODB.Recordset
Private rstDetail As ADODB.Recordset
Private rstHeaderNotes As ADODB.Recordset
Private rstDetailNotes As ADODB.Recordset

Private blnParent As Boolean

Private Sub Form_Load()
    SetInitialization
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If blnParent Then
        Cancel = 1
    Else
        If frmReminderList.Parent Then
            frmReminderList.Parent = False
        End If
    
        mdlDatabase.CloseRecordset rstHeaderNotes
        mdlDatabase.CloseRecordset rstDetailNotes
        mdlDatabase.CloseRecordset rstDetail
        mdlDatabase.CloseRecordset rstHeader
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmMISTMCUSTOMER = Nothing
End Sub

Private Sub txtCustomerId_GotFocus()
    mdlProcedures.GotFocus Me.txtCustomerId
End Sub

Private Sub txtName_GotFocus()
    mdlProcedures.GotFocus Me.txtName
End Sub

Private Sub cmdSearch_Click()
    SetGridHeader False
End Sub

Private Sub dgdHeader_DblClick()
    If blnParent Then Exit Sub
    
    If rstHeader.RecordCount > 0 Then
        If mdlGlobal.UserAuthority.IsMenuAccess(mdiMain.mnuTMCUSTOMER.Name) Then
            blnParent = True
            
            mdlProcedures.CornerWindows Me
            
            mdlProcedures.ShowForm frmTMCUSTOMERNOTES, False, True
        End If
    End If
End Sub

Private Sub dgdHeader_HeadClick(ByVal ColIndex As Integer)
    rstHeader.Sort = rstHeader.Fields(ColIndex).Name
    
    If rstHeader.RecordCount > 0 Then
        SetGridDetail rstHeader!CustomerId
        SetGridHeaderNotes rstHeader!CustomerId
    End If
End Sub

Private Sub dgdHeader_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If rstHeader.RecordCount > 0 Then
        SetGridDetail rstHeader!CustomerId
        SetGridHeaderNotes rstHeader!CustomerId
    End If
End Sub

Private Sub dgdDetail_DblClick()
    If blnParent Then Exit Sub
    
    If rstDetail.RecordCount > 0 Then
        If mdlGlobal.UserAuthority.IsMenuAccess(mdiMain.mnuTMCUSTOMER.Name) Then
            blnParent = True
            
            mdlProcedures.CornerWindows Me
            
            mdlProcedures.ShowForm frmTMCONTACTNOTES, False, True
        End If
    End If
End Sub

Private Sub dgdDetail_HeadClick(ByVal ColIndex As Integer)
    rstDetail.Sort = rstDetail.Fields(ColIndex).Name
    
    If rstDetail.RecordCount > 0 Then
        SetGridDetailNotes rstDetail!ContactId
    End If
End Sub

Private Sub dgdHeaderNotes_HeadClick(ByVal ColIndex As Integer)
    rstHeaderNotes.Sort = rstHeaderNotes.Fields(ColIndex).Name
End Sub

Private Sub dgdDetailNotes_HeadClick(ByVal ColIndex As Integer)
    rstDetailNotes.Sort = rstDetailNotes.Fields(ColIndex).Name
End Sub

Private Sub dgdDetail_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If rstDetail.RecordCount > 0 Then
        SetGridDetailNotes rstDetail!ContactId
    End If
End Sub

Private Sub SetInitialization()
    mdlProcedures.CenterWindows Me
    
    Me.Caption = mdlText.strMISTMCUSTOMER
    
    If frmReminderList.Parent Then
        Me.txtCustomerId.Text = frmReminderList.CustomerId
        
        Me.txtName.Text = frmReminderList.CustomerName
    End If
    
    SetGridHeader
End Sub

Private Sub SetGridHeader(Optional ByVal blnInitialize As Boolean = True)
    Dim strCriteria As String
    
    If blnInitialize Then
        strCriteria = "CustomerId=''"
    Else
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
    End If
    
    Set rstHeader = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "CustomerId, Name, Phone, Fax", mdlTable.CreateTMCUSTOMER, False, strCriteria, "CustomerId ASC")
    
    If rstHeader.RecordCount > 0 Then
        SetGridDetail rstHeader!CustomerId
        SetGridHeaderNotes rstHeader!CustomerId
    Else
        SetGridDetail
        SetGridHeaderNotes
    End If
    
    Set Me.dgdHeader.DataSource = rstHeader
    
    With Me.dgdHeader
        .RowHeight = 1000
        
        .Columns(0).Width = 800
        .Columns(0).Locked = True
        .Columns(0).Caption = "Kode"
        .Columns(1).Width = 3000
        .Columns(1).Locked = True
        .Columns(1).Caption = "Nama"
        .Columns(1).WrapText = True
        .Columns(2).Width = 2000
        .Columns(2).Locked = True
        .Columns(2).Caption = "Telepon"
        .Columns(2).WrapText = True
        .Columns(3).Width = 2000
        .Columns(3).Locked = True
        .Columns(3).Caption = "Fax"
        .Columns(3).WrapText = True
    End With
End Sub

Private Sub SetGridDetail(Optional ByVal strCustomerId As String = "")
    Dim strCriteria As String
    
    strCriteria = "CustomerId='" & strCustomerId & "'"
    
    Set rstDetail = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "ContactId, Name, Phone, HandPhone, Email", mdlTable.CreateTMCONTACTCUSTOMER, False, strCriteria, "Name ASC")
    
    If rstDetail.RecordCount > 0 Then
        SetGridDetailNotes rstDetail!ContactId
    Else
        SetGridDetailNotes
    End If
    
    Set Me.dgdDetail.DataSource = rstDetail
    
    With Me.dgdDetail
        .RowHeight = 800
        
        .Columns(0).Width = 0
        .Columns(0).Locked = True
        .Columns(0).Visible = False
        .Columns(1).Width = 2800
        .Columns(1).Locked = True
        .Columns(1).Caption = "Nama"
        .Columns(1).WrapText = True
        .Columns(2).Width = 1650
        .Columns(2).Locked = True
        .Columns(2).Caption = "Telepon"
        .Columns(2).WrapText = True
        .Columns(3).Width = 1650
        .Columns(3).Locked = True
        .Columns(3).Caption = "HP"
        .Columns(3).WrapText = True
        .Columns(4).Width = 1650
        .Columns(4).Locked = True
        .Columns(4).Caption = "Email"
        .Columns(4).WrapText = True
    End With
End Sub

Private Sub SetGridHeaderNotes(Optional ByVal strCustomerId As String = "")
    Dim strCriteria As String
    
    strCriteria = "CustomerId='" & strCustomerId & "'"
    
    Set rstHeaderNotes = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "NotesDate, Notes", mdlTable.CreateTMCUSTOMERNOTES, False, strCriteria, "NotesDate DESC")
    
    Set Me.dgdHeaderNotes.DataSource = rstHeaderNotes
    
    With Me.dgdHeaderNotes
        .RowHeight = 1000
        
        .Columns(0).Width = 2100
        .Columns(0).Locked = True
        .Columns(0).Caption = "Tanggal"
        .Columns(0).NumberFormat = "dd MMMM yyyy"
        .Columns(1).Width = 3200
        .Columns(1).Locked = True
        .Columns(1).Caption = "Catatan"
        .Columns(1).WrapText = True
    End With
End Sub

Private Sub SetGridDetailNotes(Optional ByVal strContactId As String = "")
    Dim strCriteria As String
    
    strCriteria = "ContactId='" & strContactId & "'"
    
    Set rstDetailNotes = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "NotesDate, Notes", mdlTable.CreateTMCONTACTNOTES, False, strCriteria, "NotesDate DESC")
    
    Set Me.dgdDetailNotes.DataSource = rstDetailNotes
    
    With Me.dgdDetailNotes
        .RowHeight = 1000
        
        .Columns(0).Width = 2100
        .Columns(0).Locked = True
        .Columns(0).Caption = "Tanggal"
        .Columns(0).NumberFormat = "dd MMMM yyyy"
        .Columns(1).Width = 3200
        .Columns(1).Locked = True
        .Columns(1).Caption = "Catatan"
        .Columns(1).WrapText = True
    End With
End Sub

Public Property Get Parent() As Boolean
    Parent = blnParent
End Property

Public Property Get CustomerId() As String
    If rstHeader Is Nothing Then Exit Property
    
    If rstHeader.RecordCount > 0 Then
        CustomerId = rstHeader!CustomerId
    End If
End Property

Public Property Get CustomerName() As String
    If rstHeader Is Nothing Then Exit Property
    
    If rstHeader.RecordCount > 0 Then
        CustomerName = rstHeader!Name
    End If
End Property

Public Property Get ContactId() As String
    If rstDetail Is Nothing Then Exit Property
    
    If rstDetail.RecordCount > 0 Then
        ContactId = rstDetail!ContactId
    End If
End Property

Public Property Get ContactName() As String
    If rstDetail Is Nothing Then Exit Property
    
    If rstDetail.RecordCount > 0 Then
        ContactName = rstDetail!Name
    End If
End Property

Public Property Let Parent(ByVal blnEnable As Boolean)
    blnParent = blnEnable
    
    If Not blnEnable Then mdlProcedures.CenterWindows Me
End Property

Public Property Let CustomerId(ByVal strCustomerId As String)
    SetGridDetail strCustomerId
    SetGridHeaderNotes strCustomerId
End Property

Public Property Let ContactId(ByVal strContactId As String)
    SetGridDetailNotes strContactId
End Property
