VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmTDRTRSELL 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7605
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   8910
   Icon            =   "frmTDRTRSELL.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7605
   ScaleWidth      =   8910
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdSave 
      Caption         =   "Simpan"
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
      Left            =   7560
      TabIndex        =   6
      Top             =   7080
      Width           =   1215
   End
   Begin VB.Frame fraSearch 
      Height          =   1335
      Left            =   120
      TabIndex        =   22
      Top             =   120
      Width           =   8655
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
         Left            =   7440
         TabIndex        =   3
         Top             =   360
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
         Left            =   1200
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
         Left            =   1200
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
         Left            =   7440
         TabIndex        =   4
         Top             =   840
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
         Left            =   1200
         TabIndex        =   1
         Top             =   600
         Width           =   4935
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
   Begin VB.Frame fraHeader 
      Height          =   975
      Left            =   120
      TabIndex        =   24
      Top             =   6000
      Width           =   8655
      Begin VB.TextBox txtQty 
         Alignment       =   1  'Right Justify
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
         Left            =   1080
         MaxLength       =   10
         TabIndex        =   5
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lblUnityId 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Satuan"
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
         TabIndex        =   19
         Top             =   600
         Width           =   615
      End
      Begin VB.Label lblQty 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah"
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
         TabIndex        =   18
         Top             =   240
         Width           =   615
      End
      Begin VB.Label txtUnityId 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
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
         Height          =   255
         Left            =   1080
         TabIndex        =   20
         Top             =   600
         Width           =   6975
      End
   End
   Begin VB.Frame fraInfo 
      Height          =   1695
      Left            =   120
      TabIndex        =   23
      Top             =   4320
      Width           =   8655
      Begin VB.Label lblWarehouseId 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Gudang"
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
         TabIndex        =   14
         Top             =   960
         Width           =   675
      End
      Begin VB.Label lblStockQty 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Stok"
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
         TabIndex        =   16
         Top             =   1320
         Width           =   405
      End
      Begin VB.Label txtStockQty 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
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
         Height          =   255
         Left            =   1080
         TabIndex        =   17
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label txtSJQty 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
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
         Height          =   255
         Left            =   1080
         TabIndex        =   13
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label txtItemIdText 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
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
         Height          =   255
         Left            =   1080
         TabIndex        =   11
         Top             =   240
         Width           =   7335
      End
      Begin VB.Label lblSJQty 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Qty"
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
         TabIndex        =   12
         Top             =   600
         Width           =   315
      End
      Begin VB.Label lblItemIdText 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Barang"
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
         Top             =   240
         Width           =   615
      End
      Begin VB.Label txtWarehouseId 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
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
         Height          =   255
         Left            =   1080
         TabIndex        =   15
         Top             =   960
         Width           =   7335
      End
   End
   Begin MSDataGridLib.DataGrid dgdMain 
      Height          =   2655
      Left            =   120
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   1560
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   4683
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
Attribute VB_Name = "frmTDRTRSELL"
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
        If frmTHRTRSELL.Parent Then
            frmTHRTRSELL.Parent = False
        End If
        
        mdlDatabase.CloseRecordset rstMain
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmTDRTRSELL = Nothing
End Sub

Private Sub txtQty_Change()
    Me.txtQty.Text = mdlProcedures.FormatCurrency(Me.txtQty.Text)
    
    Me.txtQty.SelStart = Len(Me.txtQty.Text)
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

Private Sub txtQty_GotFocus()
    mdlProcedures.GotFocus Me.txtQty
End Sub

Private Sub dgdMain_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If rstMain.RecordCount > 0 Then
        FillText
    End If
End Sub

Private Sub cmdOptional_Click()
    If blnParent Then Exit Sub
    
    blnParent = True
    
    mdlProcedures.CornerWindows Me, , False
    
    mdlProcedures.ShowForm frmBRWTMITEMOPT, False, True
End Sub

Private Sub cmdSearch_Click()
    SetGrid False
End Sub

Private Sub cmdSave_Click()
    If Not CheckValidation Then Exit Sub
    
    frmTHRTRSELL.SaveDetail rstMain!ItemId, mdlProcedures.GetCurrency(Me.txtQty.Text)
End Sub

Private Function CheckValidation() As Boolean
    If rstMain.RecordCount > 0 Then
        Dim curQty As Currency
        Dim curStockQty As Currency
        Dim curMaxStock As Currency
        
        curQty = mdlProcedures.GetCurrency(Me.txtQty.Text)
        curStockQty = mdlProcedures.GetCurrency(Me.txtStockQty.Caption)
        curMaxStock = mdlProcedures.GetCurrency(mdlDatabase.GetFieldData(mdlGlobal.conInventory, "MaxStock", mdlTable.CreateTMITEM, "ItemId='" & rstMain!ItemId & "'"))
        
        If curQty > 0 Then
            If curMaxStock > 0 Then
                If (curStockQty + curQty) > curMaxStock Then
                    MsgBox "Stok Maksimum : " & CStr(curMaxStock), vbOKOnly + vbExclamation, Me.Caption
                    
                    Me.txtQty.SetFocus
                    
                    CheckValidation = False
                    
                    Exit Function
                End If
            End If
        Else
            MsgBox "Qty Harap Diisi Terlebih Dahulu", vbOKOnly + vbExclamation, Me.Caption
            
            Me.txtQty.SetFocus
            
            CheckValidation = False
            
            Exit Function
        End If
        
        If Not (mdlProcedures.GetCurrency(Me.txtSJQty.Caption) >= curQty) Then
            MsgBox "Qty Melebihi Qty SJ", vbOKOnly + vbExclamation, Me.Caption
            
            Me.txtQty.SetFocus
            
            CheckValidation = False
            
            Exit Function
        End If
    Else
        Me.txtItemId.SetFocus
        
        CheckValidation = False
        
        Exit Function
    End If
    
    CheckValidation = True
End Function

Private Sub SetInitialization()
    mdlProcedures.CenterWindows Me, False
    
    Me.Caption = mdlText.strTDRTRSELL
    
    blnParent = False
    
    Dim strWarehouseId As String
    
    strWarehouseId = frmTHRTRSELL.WarehouseIdCombo
    
    Me.txtWarehouseId.Caption = strWarehouseId & " | " & mdlDatabase.GetFieldData(mdlGlobal.conInventory, "Name", mdlTable.CreateTMWAREHOUSE, "WarehouseId='" & strWarehouseId & "'")
    
    SetGrid False
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
    
    strTableFirst = mdlTable.CreateTDSJSELL
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
    
    strCriteria = ""
    
    If blnInitialize Then
        strCriteria = strTableFirst & ".ItemId=''"
    Else
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
    End If

    If Not Trim(strCriteria) = "" Then
        strCriteria = strCriteria & " AND "
    End If

    strCriteria = strCriteria & "SJId='" & frmTHRTRSELL.SJIdCombo & "'"
    
    Dim rstTemp As ADODB.Recordset
    
    Set rstTemp = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, strTableFirst & ".ItemId, PartNumber, " & strTableSecond & ".Name, Qty, SJId", strTable, False, strCriteria, strTableFirst & ".ItemId ASC")
    
    SetVirtualRecordset
    
    With rstTemp
        While Not .EOF
            rstMain.AddNew
            
            rstMain!ItemId = !ItemId
            rstMain!PartNumber = !PartNumber
            rstMain!Name = !Name
            rstMain!Qty = mdlProcedures.GetCurrency(!Qty) - mdlTHRTRSELL.GetQtySJFromRTRSELL(!SJId, !ItemId)
            
            rstMain.Update
            
            .MoveNext
        Wend
    End With
    
    mdlDatabase.CloseRecordset rstTemp
    
    If rstMain.RecordCount > 0 Then
        rstMain.MoveFirst
        
        FillText
    Else
        FillText True
    End If
    
    Set Me.dgdMain.DataSource = rstMain
    
    With Me.dgdMain
        .RowHeight = 500
        
        .Columns(0).Width = 1200
        .Columns(0).Locked = True
        .Columns(0).Caption = "Kode"
        .Columns(1).Width = 2500
        .Columns(1).Locked = True
        .Columns(1).Caption = "Nomor Part"
        .Columns(1).WrapText = True
        .Columns(2).Width = 3100
        .Columns(2).Locked = True
        .Columns(2).Caption = "Nama"
        .Columns(2).WrapText = True
        .Columns(3).Width = 1200
        .Columns(3).Locked = True
        .Columns(3).Caption = "Qty. SJ"
        .Columns(3).NumberFormat = "#,##0"
        .Columns(3).Alignment = dbgRight
    End With
End Sub

Private Sub SetVirtualRecordset()
    mdlDatabase.CloseRecordset rstMain
    
    Set rstMain = New ADODB.Recordset
    rstMain.CursorLocation = adUseClient
    
    With rstMain
        .Fields.Append "ItemId", adChar, mdlDatabase.GetColumnSize(mdlGlobal.conInventory, "ItemId", mdlTable.CreateTMITEM)
        .Fields.Append "PartNumber", adChar, mdlDatabase.GetColumnSize(mdlGlobal.conInventory, "PartNumber", mdlTable.CreateTMITEM)
        .Fields.Append "Name", adChar, mdlDatabase.GetColumnSize(mdlGlobal.conInventory, "Name", mdlTable.CreateTMITEM)
        .Fields.Append "Qty", adCurrency
        
        .Open
    End With
End Sub

Private Sub FillText(Optional ByVal blnClear As Boolean = False)
    If blnClear Then
        Me.txtItemIdText.Caption = ""
        Me.txtSJQty.Caption = ""
        Me.txtStockQty.Caption = ""
        Me.txtUnityId.Caption = ""
    Else
        With rstMain
            Me.txtItemIdText.Caption = !ItemId & " | " & !Name
            Me.txtSJQty.Caption = mdlProcedures.FormatCurrency(!Qty)
            Me.txtStockQty.Caption = mdlProcedures.FormatCurrency(mdlTransaction.CheckStock(!ItemId, frmTHRTRSELL.WarehouseIdCombo))
            
            Dim strUnityId As String
            
            strUnityId = mdlDatabase.GetFieldData(mdlGlobal.conInventory, "UnityId", mdlTable.CreateTMITEM, "ItemId='" & !ItemId & "'")
            
            If Trim(strUnityId) = "" Then
                Me.txtUnityId.Caption = strUnityId
            Else
                Me.txtUnityId.Caption = strUnityId & " | " & mdlDatabase.GetFieldData(mdlGlobal.conInventory, "Name", mdlTable.CreateTMUNITY, "UnityId='" & strUnityId & "'")
            End If
        End With
    End If
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
    
    If Not blnEnable Then mdlProcedures.CenterWindows Me, False
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
