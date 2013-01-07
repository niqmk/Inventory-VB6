VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmMISTMCUSTOMERTRANS 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7710
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   8685
   Icon            =   "frmMISTMCUSTOMERTRANS.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7710
   ScaleWidth      =   8685
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraSearch 
      Height          =   975
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   8415
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
   End
   Begin MSDataGridLib.DataGrid dgdHeader 
      Height          =   2535
      Left            =   120
      TabIndex        =   7
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
   Begin MSDataGridLib.DataGrid dgdDetail 
      Height          =   3375
      Left            =   120
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   3840
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
   Begin VB.Label txtTotal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   6480
      TabIndex        =   6
      Top             =   7320
      Width           =   2100
   End
   Begin VB.Label lblTotal 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total"
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
      Left            =   5880
      TabIndex        =   5
      Top             =   7320
      Width           =   420
   End
End
Attribute VB_Name = "frmMISTMCUSTOMERTRANS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private rstHeader As ADODB.Recordset
Private rstDetail As ADODB.Recordset

Private Sub Form_Load()
    SetInitialization
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If frmReminderList.Parent Then
        frmReminderList.Parent = False
    End If
    
    mdlDatabase.CloseRecordset rstDetail
    mdlDatabase.CloseRecordset rstHeader
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmMISTMCUSTOMERTRANS = Nothing
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

Private Sub dgdHeader_HeadClick(ByVal ColIndex As Integer)
    rstHeader.Sort = rstHeader.Fields(ColIndex).Name
    
    If rstHeader.RecordCount > 0 Then
        SetGridDetail rstHeader!CustomerId
    End If
End Sub

Private Sub dgdHeader_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If rstHeader.RecordCount > 0 Then
        SetGridDetail rstHeader!CustomerId
    End If
End Sub

Private Sub dgdDetail_HeadClick(ByVal ColIndex As Integer)
    rstDetail.Sort = rstDetail.Fields(ColIndex).Name
End Sub

Private Sub SetInitialization()
    mdlProcedures.CenterWindows Me
    
    Me.Caption = mdlText.strMISTMCUSTOMERTRANS
    
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
    Else
        SetGridDetail
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
    
    Dim rstTemp As ADODB.Recordset
    
    Set rstTemp = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "POId, PODate, POCustomerId", mdlTable.CreateTHPOSELL, False, strCriteria, "POId ASC")
    
    SetVirtualRecordset
    
    Dim curTotal As Currency
    
    curTotal = 0
    
    With rstTemp
        While Not .EOF
            rstDetail.AddNew
            
            rstDetail!POId = !POId
            rstDetail!PODate = !PODate
            rstDetail!POCustomerId = !POCustomerId
            rstDetail!Qty = mdlTHPOSELL.GetTotalQtyPOSELL(!POId)
            
            curTotal = curTotal + mdlProcedures.GetCurrency(rstDetail!Qty)
            
            rstDetail.Update
            
            .MoveNext
        Wend
    End With
    
    mdlDatabase.CloseRecordset rstTemp
    
    Me.txtTotal.Caption = mdlProcedures.FormatCurrency(CStr(curTotal))
    
    If rstDetail.RecordCount > 0 Then
        rstDetail.MoveFirst
    End If
    
    Set Me.dgdDetail.DataSource = rstDetail
    
    With Me.dgdDetail
        .RowHeight = 300
        
        .Columns(0).Width = 2500
        .Columns(0).Locked = True
        .Columns(0).Caption = "Nomor PO"
        .Columns(1).Width = 2200
        .Columns(1).Locked = True
        .Columns(1).Caption = "Tanggal"
        .Columns(1).NumberFormat = "dd MMMM yyyy"
        .Columns(2).Width = 1550
        .Columns(2).Locked = True
        .Columns(2).Caption = "Nomor PO Customer"
        .Columns(2).WrapText = True
        .Columns(3).Width = 1500
        .Columns(3).Locked = True
        .Columns(3).Caption = "Qty"
        .Columns(3).NumberFormat = "#,##0"
        .Columns(3).Alignment = dbgRight
    End With
End Sub

Private Sub SetVirtualRecordset()
    mdlDatabase.CloseRecordset rstDetail
    
    Set rstDetail = New ADODB.Recordset
    rstDetail.CursorLocation = adUseClient
    
    With rstDetail
        .Fields.Append "POId", adChar, mdlDatabase.GetColumnSize(mdlGlobal.conInventory, "POId", mdlTable.CreateTHPOSELL)
        .Fields.Append "PODate", adDate
        .Fields.Append "POCustomerId", adChar, mdlDatabase.GetColumnSize(mdlGlobal.conInventory, "POCustomerId", mdlTable.CreateTHPOSELL)
        .Fields.Append "Qty", adCurrency
        
        .Open
    End With
End Sub
