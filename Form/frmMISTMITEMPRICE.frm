VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmMISTMITEMPRICE 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5670
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   9270
   Icon            =   "frmMISTMITEMPRICE.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5670
   ScaleWidth      =   9270
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraSearch 
      Height          =   1815
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   9015
      Begin VB.CommandButton cmdOptional 
         Caption         =   "Optional"
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
         Left            =   7800
         TabIndex        =   5
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox txtItemId 
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
         Width           =   975
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
         Left            =   1440
         TabIndex        =   2
         Top             =   960
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
         Left            =   7800
         TabIndex        =   6
         Top             =   1320
         Width           =   1095
      End
      Begin VB.TextBox txtPartNumber 
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
         Width           =   4935
      End
      Begin MSComCtl2.DTPicker dtpStartDate 
         Height          =   375
         Left            =   1440
         TabIndex        =   3
         Top             =   1320
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
         Format          =   99287043
         CurrentDate     =   39335
      End
      Begin MSComCtl2.DTPicker dtpFinishDate 
         Height          =   375
         Left            =   6000
         TabIndex        =   4
         Top             =   1320
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
         Format          =   99287043
         CurrentDate     =   39335
      End
      Begin VB.Label lblItemId 
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
         TabIndex        =   7
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
         TabIndex        =   9
         Top             =   960
         Width           =   510
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
         TabIndex        =   10
         Top             =   1320
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
         Left            =   4320
         TabIndex        =   11
         Top             =   1320
         Width           =   1395
      End
      Begin VB.Label lblPartNumber 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nomor Part"
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
         Top             =   600
         Width           =   990
      End
   End
   Begin MSDataGridLib.DataGrid dgdMain 
      Height          =   3495
      Left            =   120
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   2040
      Width           =   9015
      _ExtentX        =   15901
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
Attribute VB_Name = "frmMISTMITEMPRICE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private rstMain As ADODB.Recordset

Private strVendor As String
Private strGroup As String
Private strCategory As String
Private strBrand As String
Private strUnity As String

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
        mdlDatabase.CloseRecordset rstMain
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmMISTMITEMPRICE = Nothing
End Sub

Private Sub txtItemId_GotFocus()
    mdlProcedures.GotFocus Me.txtItemId
End Sub

Private Sub txtPartNumber_GotFocus()
    mdlProcedures.GotFocus Me.txtPartNumber
End Sub

Private Sub txtName_GotFocus()
    mdlProcedures.GotFocus Me.txtName
End Sub

Private Sub cmdOptional_Click()
    If blnParent Then Exit Sub
    
    blnParent = True
    
    mdlProcedures.CornerWindows Me
    
    mdlProcedures.ShowForm frmBRWTMITEMOPT, False, True
End Sub

Private Sub cmdSearch_Click()
    SetGrid False
End Sub

Private Sub dgdMain_HeadClick(ByVal ColIndex As Integer)
    rstMain.Sort = rstMain.Fields(ColIndex).Name
End Sub

Private Sub SetInitialization()
    mdlProcedures.CenterWindows Me
    
    Me.Caption = mdlText.strMISTMITEMPRICE
    
    blnParent = False
    
    Me.dtpStartDate.CustomFormat = mdlGlobal.strFormatDate
    Me.dtpFinishDate.CustomFormat = mdlGlobal.strFormatDate
    
    Dim strYear As String
    Dim strMonth As String
    
    strYear = mdlProcedures.FormatDate(Now, "yyyy")
    strMonth = mdlProcedures.FormatDate(Now, "MM")
    
    Me.dtpStartDate.Value = mdlProcedures.SetDate(strMonth, strYear)
    Me.dtpFinishDate.Value = mdlProcedures.SetDate(strMonth, strYear, , True)
    
    SetGrid
End Sub

Private Sub SetGrid(Optional ByVal blnInitialize As Boolean = True)
    Dim strTableFirst As String
    Dim strTableSecond As String
    Dim strTableThird As String
    Dim strTableFourth As String
    Dim strTableFifth As String
    Dim strTableSixth As String
    Dim strTableSeventh As String
    Dim strTable As String
    
    strTableFirst = mdlTable.CreateTMITEMPRICE
    strTableSecond = mdlTable.CreateTMITEM
    strTableThird = mdlTable.CreateTMVENDOR
    strTableFourth = mdlTable.CreateTMGROUP
    strTableFifth = mdlTable.CreateTMCATEGORY
    strTableSixth = mdlTable.CreateTMBRAND
    strTableSeventh = mdlTable.CreateTMUNITY
    
    strTable = "((((((" & strTableFirst & " LEFT JOIN " & strTableSecond & _
        " ON " & strTableFirst & ".ItemId=" & strTableSecond & ".ItemId) LEFT JOIN " & strTableThird & _
        " ON " & strTableSecond & ".VendorId=" & strTableThird & ".VendorId) LEFT JOIN " & strTableFourth & _
        " ON " & strTableSecond & ".GroupId=" & strTableFourth & ".GroupId) LEFT JOIN " & strTableFifth & _
        " ON " & strTableSecond & ".CategoryId=" & strTableFifth & ".CategoryId) LEFT JOIN " & strTableSixth & _
        " ON " & strTableSecond & ".BrandId=" & strTableSixth & ".BrandId) LEFT JOIN " & strTableSeventh & _
        " ON " & strTableSecond & ".UnityId=" & strTableSeventh & ".UnityId)"
        
    Dim strCriteria As String
    
    If blnInitialize Then
        strCriteria = strTableFirst & ".ItemId=''"
    Else
        strCriteria = ""
        
        If Not Trim(Me.txtItemId.Text) = "" Then
            strCriteria = mdlProcedures.QueryLikeCriteria(strTableFirst & ".ItemId", mdlProcedures.RepDupText(Me.txtItemId.Text))
        End If
        
        If Not Trim(Me.txtPartNumber.Text) = "" Then
            If Not Trim(strCriteria) = "" Then
                strCriteria = strCriteria & " AND "
            End If
            
            strCriteria = strCriteria & mdlProcedures.QueryLikeCriteria("PartNumber", mdlProcedures.RepDupText(Me.txtPartNumber.Text))
        End If
        
        If Not Trim(Me.txtName.Text) = "" Then
            If Not Trim(strCriteria) = "" Then
                strCriteria = strCriteria & " AND "
            End If
            
            strCriteria = strCriteria & mdlProcedures.QueryLikeCriteria(strTableSecond & ".Name", mdlProcedures.RepDupText(Me.txtName.Text))
        End If
        
        Dim strOptional(1) As String
        
        If Not Trim(strVendor) = "" Then
            strOptional(0) = mdlProcedures.SplitData(strVendor)
            strOptional(1) = mdlProcedures.SplitData(strVendor, 1)
            
            If Not Trim(strOptional(0)) = "" Then
                If Not Trim(strCriteria) = "" Then
                    strCriteria = strCriteria & " AND "
                End If
                
                strCriteria = strCriteria & mdlProcedures.QueryLikeCriteria(strTableSecond & ".VendorId", mdlProcedures.RepDupText(strOptional(0)))
            End If
            
            If Not Trim(strOptional(1)) = "" Then
                If Not Trim(strCriteria) = "" Then
                    strCriteria = strCriteria & " AND "
                End If
                
                strCriteria = strCriteria & mdlProcedures.QueryLikeCriteria(strTableThird & ".Name", mdlProcedures.RepDupText(strOptional(1)))
            End If
        End If
        
        If Not Trim(strGroup) = "" Then
            strOptional(0) = mdlProcedures.SplitData(strGroup)
            strOptional(1) = mdlProcedures.SplitData(strGroup, 1)
            
            If Not Trim(strOptional(0)) = "" Then
                If Not Trim(strCriteria) = "" Then
                    strCriteria = strCriteria & " AND "
                End If
                
                strCriteria = strCriteria & mdlProcedures.QueryLikeCriteria(strTableSecond & ".GroupId", mdlProcedures.RepDupText(strOptional(0)))
            End If
            
            If Not Trim(strOptional(1)) = "" Then
                If Not Trim(strCriteria) = "" Then
                    strCriteria = strCriteria & " AND "
                End If
                
                strCriteria = strCriteria & mdlProcedures.QueryLikeCriteria(strTableFourth & ".Name", mdlProcedures.RepDupText(strOptional(1)))
            End If
        End If
        
        If Not Trim(strCategory) = "" Then
            strOptional(0) = mdlProcedures.SplitData(strCategory)
            strOptional(1) = mdlProcedures.SplitData(strCategory, 1)
            
            If Not Trim(strOptional(0)) = "" Then
                If Not Trim(strCriteria) = "" Then
                    strCriteria = strCriteria & " AND "
                End If
                
                strCriteria = strCriteria & mdlProcedures.QueryLikeCriteria(strTableSecond & ".CategoryId", mdlProcedures.RepDupText(strOptional(0)))
            End If
            
            If Not Trim(strOptional(1)) = "" Then
                If Not Trim(strCriteria) = "" Then
                    strCriteria = strCriteria & " AND "
                End If
                
                strCriteria = strCriteria & mdlProcedures.QueryLikeCriteria(strTableFifth & ".Name", mdlProcedures.RepDupText(strOptional(1)))
            End If
        End If
        
        If Not Trim(strBrand) = "" Then
            strOptional(0) = mdlProcedures.SplitData(strBrand)
            strOptional(1) = mdlProcedures.SplitData(strBrand, 1)
            
            If Not Trim(strOptional(0)) = "" Then
                If Not Trim(strCriteria) = "" Then
                    strCriteria = strCriteria & " AND "
                End If
                
                strCriteria = strCriteria & mdlProcedures.QueryLikeCriteria(strTableSecond & ".BrandId", mdlProcedures.RepDupText(strOptional(0)))
            End If
            
            If Not Trim(strOptional(1)) = "" Then
                If Not Trim(strCriteria) = "" Then
                    strCriteria = strCriteria & " AND "
                End If
                
                strCriteria = strCriteria & mdlProcedures.QueryLikeCriteria(strTableSixth & ".Name", mdlProcedures.RepDupText(strOptional(1)))
            End If
        End If
        
        If Not Trim(strUnity) = "" Then
            strOptional(0) = mdlProcedures.SplitData(strUnity)
            strOptional(1) = mdlProcedures.SplitData(strUnity, 1)
            
            If Not Trim(strOptional(0)) = "" Then
                If Not Trim(strCriteria) = "" Then
                    strCriteria = strCriteria & " AND "
                End If
                
                strCriteria = strCriteria & mdlProcedures.QueryLikeCriteria(strTableSecond & ".UnityId", mdlProcedures.RepDupText(strOptional(0)))
            End If
            
            If Not Trim(strOptional(1)) = "" Then
                If Not Trim(strCriteria) = "" Then
                    strCriteria = strCriteria & " AND "
                End If
                
                strCriteria = strCriteria & mdlProcedures.QueryLikeCriteria(strTableSeventh & ".Name", mdlProcedures.RepDupText(strOptional(1)))
            End If
        End If
        
        If Not Trim(strCriteria) = "" Then
            strCriteria = strCriteria & " AND "
        End If
        
        If mdlGlobal.objDatabaseInit = SQLSERVER7 Or mdlGlobal.objDatabaseInit = SQLSERVER2000 Or mdlGlobal.objDatabaseInit = SQLEXPRESS Then
            strCriteria = strCriteria & "PriceDate>='" & mdlProcedures.FormatDate(Me.dtpStartDate.Value) & "' AND PriceDate<='" & mdlProcedures.FormatDate(Me.dtpFinishDate.Value) & "'"
        ElseIf mdlGlobal.objDatabaseInit = MSACCESS Then
            strCriteria = strCriteria & "PriceDate>=#" & mdlProcedures.FormatDate(Me.dtpStartDate.Value) & "# AND PriceDate<=#" & mdlProcedures.FormatDate(Me.dtpFinishDate.Value) & "#"
        ElseIf mdlGlobal.objDatabaseInit = MYSQL Then
            strCriteria = strCriteria & "PriceDate>='" & mdlProcedures.FormatDate(Me.dtpStartDate.Value) & "' AND PriceDate<='" & mdlProcedures.FormatDate(Me.dtpFinishDate.Value) & "'"
        End If
    End If
    
    Set rstMain = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "PriceId, " & strTableFirst & ".ItemId, PartNumber, " & strTableSecond & ".Name, PriceDate, ItemPrice, CurrencyId", strTable, False, strCriteria, "PriceDate, " & strTableFirst & ".ItemId ASC")
    
    Set Me.dgdMain.DataSource = rstMain
    
    With Me.dgdMain
        .RowHeight = 1000
        
        .Columns(0).Width = 0
        .Columns(0).Locked = True
        .Columns(0).Visible = False
        .Columns(1).Width = 1000
        .Columns(1).Locked = True
        .Columns(1).Caption = "Kode"
        .Columns(2).Width = 1400
        .Columns(2).Locked = True
        .Columns(2).Caption = "Nomor Part"
        .Columns(2).WrapText = True
        .Columns(3).Width = 1800
        .Columns(3).Locked = True
        .Columns(3).Caption = "Nama"
        .Columns(3).WrapText = True
        .Columns(4).Width = 2200
        .Columns(4).Locked = True
        .Columns(4).Caption = "Tanggal"
        .Columns(4).NumberFormat = "dd MMMM yyyy"
        .Columns(5).Width = 1100
        .Columns(5).Locked = True
        .Columns(5).Caption = "Harga"
        .Columns(5).NumberFormat = "#,##0"
        .Columns(5).Alignment = dbgRight
        .Columns(6).Width = 800
        .Columns(6).Locked = True
        .Columns(6).Caption = "Mata Uang"
    End With
End Sub

Public Property Get Parent() As Boolean
    Parent = blnParent
End Property

Public Property Get VendorOptional() As String
    VendorOptional = strVendor
End Property

Public Property Get GroupOptional() As String
    GroupOptional = strGroup
End Property

Public Property Get CategoryOptional() As String
    CategoryOptional = strCategory
End Property

Public Property Get BrandOptional() As String
    BrandOptional = strBrand
End Property

Public Property Get UnityOptional() As String
    UnityOptional = strUnity
End Property

Public Property Let Parent(ByVal blnEnable As Boolean)
    blnParent = blnEnable
    
    If Not blnEnable Then mdlProcedures.CenterWindows Me
End Property

Public Property Let VendorOptional(ByVal strValue As String)
    strVendor = strValue
End Property

Public Property Let GroupOptional(ByVal strValue As String)
    strGroup = strValue
End Property

Public Property Let CategoryOptional(ByVal strValue As String)
    strCategory = strValue
End Property

Public Property Let BrandOptional(ByVal strValue As String)
    strBrand = strValue
End Property

Public Property Let UnityOptional(ByVal strValue As String)
    strUnity = strValue
End Property
