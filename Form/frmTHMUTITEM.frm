VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmTHMUTITEM 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7770
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9750
   Icon            =   "frmTHMUTITEM.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7770
   ScaleWidth      =   9750
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtInput 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1320
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   5040
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Frame fraHeader 
      Height          =   1695
      Left            =   120
      TabIndex        =   17
      Top             =   1920
      Width           =   9495
      Begin VB.CommandButton cmdWarehouseId 
         Caption         =   "..."
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
         Left            =   9000
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   720
         Width           =   375
      End
      Begin VB.ComboBox cmbWarehouseTo 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1560
         TabIndex        =   5
         Top             =   1200
         Width           =   7335
      End
      Begin VB.ComboBox cmbWarehouseFrom 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1560
         TabIndex        =   4
         Top             =   720
         Width           =   7335
      End
      Begin MSComCtl2.DTPicker dtpMutDate 
         Height          =   375
         Left            =   1560
         TabIndex        =   3
         Top             =   240
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
         Format          =   51118083
         CurrentDate     =   39330
      End
      Begin VB.Label lblWarehouseTo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Gudang Tujuan"
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
         TabIndex        =   13
         Top             =   1200
         Width           =   1305
      End
      Begin VB.Label lblMutDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal"
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
         TabIndex        =   11
         Top             =   240
         Width           =   675
      End
      Begin VB.Label lblWarehouseFrom 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Gudang Asal"
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
         Top             =   720
         Width           =   1125
      End
   End
   Begin VB.Frame fraDetail 
      Height          =   975
      Left            =   120
      TabIndex        =   18
      Top             =   3600
      Width           =   9495
      Begin VB.TextBox txtNotes 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   1560
         MaxLength       =   150
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Top             =   240
         Width           =   7335
      End
      Begin VB.Label lblNotes 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Keterangan"
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
         Top             =   240
         Width           =   990
      End
   End
   Begin VB.Frame fraSearch 
      Height          =   735
      Left            =   120
      TabIndex        =   16
      Top             =   1200
      Width           =   9495
      Begin VB.TextBox txtYear 
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
         Left            =   5760
         MaxLength       =   4
         TabIndex        =   1
         Top             =   240
         Width           =   615
      End
      Begin VB.ComboBox cmbMonth 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1560
         TabIndex        =   0
         Top             =   240
         Width           =   615
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
         Left            =   8280
         TabIndex        =   2
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label lblYear 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tahun"
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
         Left            =   4560
         TabIndex        =   10
         Top             =   240
         Width           =   525
      End
      Begin VB.Label lblMonth 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bulan"
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
         Top             =   240
         Width           =   495
      End
   End
   Begin MSFlexGridLib.MSFlexGrid flxDetail 
      Height          =   2175
      Left            =   120
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   4680
      Width           =   9480
      _ExtentX        =   16722
      _ExtentY        =   3836
      _Version        =   393216
      ForeColorFixed  =   0
      BackColorBkg    =   -2147483633
      FillStyle       =   1
      SelectionMode   =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbDetail 
      Align           =   2  'Align Bottom
      Height          =   630
      Left            =   0
      TabIndex        =   19
      Top             =   7140
      Width           =   9750
      _ExtentX        =   17198
      _ExtentY        =   1111
      ButtonWidth     =   609
      ButtonHeight    =   953
      Appearance      =   1
      _Version        =   393216
      Begin MSComctlLib.ImageList imlDetail 
         Left            =   9120
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   3
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTHMUTITEM.frx":1F8A
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTHMUTITEM.frx":43DC
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTHMUTITEM.frx":682E
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.Toolbar tlbMain 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Width           =   9750
      _ExtentX        =   17198
      _ExtentY        =   1111
      ButtonWidth     =   609
      ButtonHeight    =   953
      Appearance      =   1
      _Version        =   393216
      Begin MSComctlLib.ImageList imlMain 
         Left            =   9120
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   48
         ImageHeight     =   48
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   7
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTHMUTITEM.frx":8C80
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTHMUTITEM.frx":B0D2
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTHMUTITEM.frx":D524
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTHMUTITEM.frx":F976
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTHMUTITEM.frx":11DC8
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTHMUTITEM.frx":1236A
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTHMUTITEM.frx":147BC
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "frmTHMUTITEM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum ButtonMode
    [AddButton] = 1
    [UpdateButton]
    [DeleteButton]
    [PrintButton]
    [BrowseButton]
    [SaveButton]
    [CancelButton]
End Enum

Private Enum ButtonDetailMode
    [AddDetailButton] = 1
    [UpdateDetailButton]
    [DeleteDetailButton]
End Enum

Private Enum ColumnConstants
    [BlankColumn]
    [ItemIdColumn]
    [NameColumn]
    [QtyColumn]
    [UnityIdColumn]
End Enum

Private rstMain As ADODB.Recordset

Private PrintTransaction As clsPRTTHMUTITEM

Private objMode As FunctionMode

Private strFormCaption As String

Private blnDetailFill As Boolean

Private blnParent As Boolean
Private blnActivate As Boolean
Private blnFrom As Boolean
Private blnTo As Boolean

Private Sub Form_Activate()
    If blnParent Then Exit Sub
    If blnActivate Then Exit Sub
    
    If rstMain.RecordCount > 0 Then
        blnParent = True
        
        mdlProcedures.CornerWindows Me
        
        mdlProcedures.ShowForm frmBRWTHMUTITEM, False, True
    End If
    
    blnActivate = True
End Sub

Private Sub Form_Load()
    SetInitialization
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If blnParent Then
        Cancel = 1
    Else
        If Not PrintTransaction Is Nothing Then
            Set PrintTransaction = Nothing
        End If
        
        mdlDatabase.CloseRecordset rstMain
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmTHMUTITEM = Nothing
End Sub

Private Sub tlbMain_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Me.txtInput.Visible Then Me.txtInput.Visible = False
    
    Select Case Button.Index
        Case AddButton:
            objMode = AddMode
            
            SetMode
        Case UpdateButton:
            objMode = UpdateMode
            
            SetMode
        Case DeleteButton:
            DeleteFunction
        Case PrintButton:
            PrintFunction
        Case BrowseButton:
            If blnParent Then Exit Sub
            
            blnParent = True
            
            mdlProcedures.CornerWindows Me
            
            mdlProcedures.ShowForm frmBRWTHMUTITEM, False, True
        Case SaveButton:
            SaveFunction
        Case CancelButton:
            objMode = ViewMode
            
            SetMode
            
            FillText
    End Select
End Sub

Private Sub tlbDetail_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Me.txtInput.Visible Then Me.txtInput.Visible = False
    
    Select Case Button.Index
        Case AddDetailButton:
            Dim strWarehouseFrom As String
            Dim strWarehouseTo As String
            
            strWarehouseFrom = Me.WarehouseFromCombo
            strWarehouseTo = Me.WarehouseToCombo
            
            If strWarehouseFrom = strWarehouseTo Then
                Me.cmbWarehouseFrom.ListIndex = -1
            End If
            
            If Not mdlProcedures.IsValidComboData(Me.cmbWarehouseFrom) Then
                MsgBox "Gudang Asal Harap Dipilih Terlebih Dahulu", vbOKOnly + vbExclamation, mdlProcedures.SetCaptionMode(strFormCaption, objMode)
                
                Me.cmbWarehouseFrom.SetFocus
                
                Exit Sub
            End If
            
            If Not mdlProcedures.IsValidComboData(Me.cmbWarehouseTo) Then
                MsgBox "Gudang Tujuan Harap Dipilih Terlebih Dahulu", vbOKOnly + vbExclamation, mdlProcedures.SetCaptionMode(strFormCaption, objMode)
                
                Me.cmbWarehouseTo.SetFocus
                
                Exit Sub
            End If
            
            If blnParent Then Exit Sub
            
            blnParent = True
            
            mdlProcedures.CornerWindows Me
            
            mdlProcedures.ShowForm frmTDMUTITEM, False, True
        Case UpdateDetailButton:
            With Me.txtInput
                Me.flxDetail.Col = QtyColumn
                
                .Top = Me.flxDetail.Top + Me.flxDetail.CellTop
                .Left = Me.flxDetail.Left + Me.flxDetail.CellLeft
                .Width = Me.flxDetail.CellWidth
                .Height = Me.flxDetail.CellHeight
                .Text = Me.flxDetail.TextMatrix(Me.flxDetail.Row, QtyColumn)
                
                .Visible = True
                
                .MaxLength = 10
                .Alignment = AlignmentConstants.vbRightJustify
                
                .SetFocus
            End With
        Case DeleteDetailButton:
            If mdlProcedures.SetMsgYesNo("Apakah Anda Yakin ?", mdlProcedures.SetCaptionMode(strFormCaption, objMode)) Then
                If Me.flxDetail.Rows > 2 Then
                    Me.flxDetail.RemoveItem Me.flxDetail.Row
                Else
                    Me.flxDetail.Rows = 1
                End If
                
                SetModeDetail
            End If
    End Select
End Sub

Private Sub flxDetail_Scroll()
    If Me.txtInput.Visible Then
        Me.txtInput.Visible = False
    End If
End Sub

Private Sub cmdSearch_Click()
    If Not objMode = ViewMode Then Exit Sub
    
    Dim strMonth As String
    Dim strYear As String
    
    strMonth = mdlProcedures.GetComboData(Me.cmbMonth)
    strYear = Trim(Me.txtYear.Text)
    
    If CInt(mdlProcedures.FormatDate(Me.dtpMutDate.MinDate, "M")) >= CInt(strMonth) Then
        If CInt(mdlProcedures.FormatDate(Me.dtpMutDate.MinDate, "yyyy")) >= CInt(strYear) Then
            Me.dtpMutDate.MinDate = mdlProcedures.SetDate(strMonth, strYear)
                
            Me.dtpMutDate.MaxDate = mdlProcedures.SetDate(strMonth, strYear, , True)
        ElseIf CInt(mdlProcedures.FormatDate(Me.dtpMutDate.MinDate, "yyyy")) < CInt(strYear) Then
            Me.dtpMutDate.MaxDate = mdlProcedures.SetDate(strMonth, strYear, , True)
                
            Me.dtpMutDate.MinDate = mdlProcedures.SetDate(strMonth, strYear)
        End If
    ElseIf CInt(mdlProcedures.FormatDate(Me.dtpMutDate.MinDate, "M")) < CInt(strMonth) Then
        If CInt(mdlProcedures.FormatDate(Me.dtpMutDate.MinDate, "yyyy")) > CInt(strYear) Then
            Me.dtpMutDate.MinDate = mdlProcedures.SetDate(strMonth, strYear)
            
            Me.dtpMutDate.MaxDate = mdlProcedures.SetDate(strMonth, strYear, , True)
        ElseIf CInt(mdlProcedures.FormatDate(Me.dtpMutDate.MinDate, "yyyy")) >= CInt(strYear) Then
            Me.dtpMutDate.MaxDate = mdlProcedures.SetDate(strMonth, strYear, , True)
                
            Me.dtpMutDate.MinDate = mdlProcedures.SetDate(strMonth, strYear)
        End If
    End If
    
    SetRecordset
End Sub

Private Sub cmbWarehouseFrom_Click()
    Me.flxDetail.Rows = 1
    
    SetModeDetail
End Sub

Private Sub cmbWarehouseTo_Click()
    Me.flxDetail.Rows = 1
    
    SetModeDetail
End Sub

Private Sub txtYear_GotFocus()
    mdlProcedures.GotFocus Me.txtYear
End Sub

Private Sub txtNotes_GotFocus()
    mdlProcedures.GotFocus Me.txtNotes
End Sub

Private Sub txtInput_GotFocus()
    mdlProcedures.GotFocus Me.txtInput
End Sub

Private Sub cmbMonth_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtYear_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub dtpMutDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub cmbWarehouseFrom_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If mdlProcedures.IsValidComboData(Me.cmbWarehouseFrom) Then
            SendKeys "{TAB}"
        Else
            If blnParent Then Exit Sub
            
            blnParent = True
            
            blnFrom = True
            blnTo = False
            
            mdlProcedures.CornerWindows Me
            
            mdlProcedures.ShowForm frmBRWTMWAREHOUSE, False, True
        End If
    End If
End Sub

Private Sub cmbWarehouseTo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If mdlProcedures.IsValidComboData(Me.cmbWarehouseTo) Then
            SendKeys "{TAB}"
        Else
            If blnParent Then Exit Sub
            
            blnParent = True
            
            blnFrom = False
            blnTo = True
            
            mdlProcedures.CornerWindows Me
            
            mdlProcedures.ShowForm frmBRWTMWAREHOUSE, False, True
        End If
    End If
End Sub

Private Sub txtInput_Change()
    Me.txtInput.Text = mdlProcedures.FormatCurrency(Me.txtInput.Text)
    
    Me.txtInput.SelStart = Len(Me.txtInput.Text)
End Sub

Private Sub txtInput_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Dim curQty As Currency
        Dim curMinStock As Currency
        Dim curMaxStock As Currency
        Dim curStockFromQty As Currency
        Dim curStockToQty As Currency
        
        curQty = mdlProcedures.GetCurrency(Me.txtInput.Text)
        curMinStock = mdlProcedures.GetCurrency(mdlDatabase.GetFieldData(mdlGlobal.conInventory, "MinStock", mdlTable.CreateTMITEM, "ItemId='" & Me.flxDetail.TextMatrix(Me.flxDetail.Row, ItemIdColumn) & "'"))
        curMaxStock = mdlProcedures.GetCurrency(mdlDatabase.GetFieldData(mdlGlobal.conInventory, "MaxStock", mdlTable.CreateTMITEM, "ItemId='" & Me.flxDetail.TextMatrix(Me.flxDetail.Row, ItemIdColumn) & "'"))
        curStockFromQty = mdlTransaction.CheckStock(Me.flxDetail.TextMatrix(Me.flxDetail.Row, ItemIdColumn), Me.WarehouseFromCombo)
        curStockToQty = mdlTransaction.CheckStock(Me.flxDetail.TextMatrix(Me.flxDetail.Row, ItemIdColumn), Me.WarehouseToCombo)
        
        If curQty > 0 Then
            If curQty > curStockFromQty Then
                Me.txtInput.Visible = False
                
                Exit Sub
            End If
            
            If curMinStock > 0 Then
                If objMode = AddMode Then
                    If (curStockFromQty - curQty) > curMinStock Then
                        Me.txtInput.Visible = False
                        
                        Exit Sub
                    End If
                ElseIf objMode = UpdateMode Then
                    If (curStockFromQty - (curQty - mdlTHMUTITEM.GetTotalQtyMUTITEM(rstMain!MutId, Me.flxDetail.TextMatrix(Me.flxDetail.Row, ItemIdColumn)))) < curMinStock Then
                        Me.txtInput.Visible = False
                        
                        Exit Sub
                    End If
                End If
            End If
            
            If curMaxStock > 0 Then
                If objMode = AddMode Then
                    If (curStockToQty + curQty) > curMaxStock Then
                        Me.txtInput.Visible = False
                        
                        Exit Sub
                    End If
                ElseIf objMode = UpdateMode Then
                    If (curStockToQty - (curQty - mdlTHMUTITEM.GetTotalQtyMUTITEM(rstMain!MutId, Me.flxDetail.TextMatrix(Me.flxDetail.Row, ItemIdColumn)))) > curMaxStock Then
                        Me.txtInput.Visible = False
                        
                        Exit Sub
                    End If
                End If
            End If
        Else
            Me.txtInput.Visible = False
            
            Exit Sub
        End If
    
        Me.flxDetail.TextMatrix(Me.flxDetail.Row, QtyColumn) = mdlProcedures.FormatCurrency(Me.txtInput.Text)
        
        Me.txtInput.Visible = False
    ElseIf KeyCode = vbKeyEscape Then
        Me.txtInput.Visible = False
    End If
End Sub

Private Sub txtInput_LostFocus()
    Me.txtInput.Visible = False
End Sub

Private Sub cmbMonth_Validate(Cancel As Boolean)
    If Not objMode = ViewMode Then Exit Sub
    
    If Not mdlProcedures.IsValidComboData(Me.cmbMonth) Then Me.cmbMonth.ListIndex = CInt(mdlProcedures.FormatDate(Now, "M")) - 1
End Sub

Private Sub txtYear_Validate(Cancel As Boolean)
    If Not objMode = ViewMode Then Exit Sub
    
    If Not IsNumeric(Me.txtYear.Text) Then Me.txtYear.Text = mdlProcedures.FormatDate(Now, "yyyy")
    
    If CInt(Me.txtYear.Text) < 1601 Then Me.txtYear.Text = mdlProcedures.FormatDate(Now, "yyyy")
End Sub

Private Sub cmbWarehouseFrom_Validate(Cancel As Boolean)
    If Not mdlProcedures.IsValidComboData(Me.cmbWarehouseFrom) Then Exit Sub
    If Not mdlProcedures.IsValidComboData(Me.cmbWarehouseTo) Then Exit Sub
    
    Dim strWarehouseFrom As String
    Dim strWarehouseTo As String
    
    strWarehouseFrom = Me.WarehouseFromCombo
    strWarehouseTo = Me.WarehouseToCombo
    
    If strWarehouseFrom = strWarehouseTo Then
        Me.cmbWarehouseFrom.ListIndex = -1
    End If
End Sub

Private Sub cmbWarehouseTo_Validate(Cancel As Boolean)
    If Not mdlProcedures.IsValidComboData(Me.cmbWarehouseFrom) Then Exit Sub
    If Not mdlProcedures.IsValidComboData(Me.cmbWarehouseTo) Then Exit Sub
    
    Dim strWarehouseFrom As String
    Dim strWarehouseTo As String
    
    strWarehouseFrom = Me.WarehouseFromCombo
    strWarehouseTo = Me.WarehouseToCombo
    
    If strWarehouseFrom = strWarehouseTo Then
        Me.cmbWarehouseTo.ListIndex = -1
    End If
End Sub

Private Sub cmdWarehouseId_Click()
    If Not mdlGlobal.UserAuthority.IsMenuAccess(mdiMain.mnuTMWAREHOUSE.Name) Then Exit Sub
    
    If blnParent Then Exit Sub
    
    blnParent = True
    
    mdlProcedures.CornerWindows Me
    
    mdlProcedures.ShowForm frmTMWAREHOUSE, False
End Sub

Private Sub SetInitialization()
    mdlProcedures.CenterWindows Me

    With Me.tlbMain
        .AllowCustomize = False
        
        .ImageList = Me.imlMain
        
        .Buttons.Add AddButton, , "Tambah", , AddButton
        .Buttons.Add UpdateButton, , "Ubah", , UpdateButton
        .Buttons.Add DeleteButton, , "Hapus", , DeleteButton
        .Buttons.Add PrintButton, , "Cetak", , PrintButton
        .Buttons.Add BrowseButton, , "Daftar", , BrowseButton
        .Buttons.Add SaveButton, , "Simpan", , SaveButton
        .Buttons.Add CancelButton, , "Batal", , CancelButton
    End With
    
    Me.dtpMutDate.CustomFormat = mdlGlobal.strFormatDate
    
    With Me.tlbDetail
        .AllowCustomize = False
        
        .ImageList = Me.imlDetail
        
        .Buttons.Add AddDetailButton, , "Tambah", , AddDetailButton
        .Buttons.Add UpdateDetailButton, , "Ubah", , UpdateDetailButton
        .Buttons.Add DeleteDetailButton, , "Hapus", , DeleteDetailButton
    End With
    
    FillCombo
    
    strFormCaption = mdlText.strTHMUTITEM
    
    Dim strMonth As Integer
    Dim strYear As Integer
    
    strMonth = mdlProcedures.GetComboData(Me.cmbMonth)
    strYear = Trim(Me.txtYear.Text)
    
    Me.dtpMutDate.MinDate = mdlProcedures.SetDate(strMonth, strYear)
    Me.dtpMutDate.MaxDate = mdlProcedures.SetDate(strMonth, strYear, , True)
    
    ArrangeGrid
    
    SetRecordset
    
    blnParent = False
    blnActivate = False
End Sub

Private Sub SetRecordset()
    Set rstMain = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "*", mdlTable.CreateTHMUTITEM, , "MONTH(MutDate)=" & Me.MonthTrans & " AND YEAR(MutDate)=" & Me.YearTrans, "MutId ASC")
    
    objMode = ViewMode
    
    SetMode
    
    FillText
End Sub

Private Sub SetMode()
    Dim blnFront As Boolean
    Dim blnBack As Boolean
    
    If objMode = ViewMode Then
        blnFront = True
        blnBack = False
    Else
        blnFront = False
        blnBack = True
    End If
    
    Me.fraSearch.Enabled = blnFront
    
    With Me.tlbMain
        .Buttons(AddButton).Visible = blnFront
        .Buttons(UpdateButton).Visible = blnFront
        .Buttons(DeleteButton).Visible = blnFront
        .Buttons(PrintButton).Visible = blnFront
        .Buttons(BrowseButton).Visible = blnFront
        .Buttons(SaveButton).Visible = blnBack
        .Buttons(CancelButton).Visible = blnBack
    End With
    
    With Me.tlbDetail
        .Buttons(AddDetailButton).Enabled = blnBack
        .Buttons(UpdateDetailButton).Enabled = blnBack
        .Buttons(DeleteDetailButton).Enabled = blnBack
    End With
    
    Me.fraDetail.Enabled = blnBack
    
    If objMode = AddMode Then
        Me.fraHeader.Enabled = True
    
        mdlProcedures.SetControlMode Me, objMode, , , Me.cmbMonth.Name & " | " & Me.txtYear.Name
        
        Me.flxDetail.Rows = 1
        
        SetModeDetail
        
        Me.dtpMutDate.SetFocus
    ElseIf objMode = UpdateMode Then
        Me.fraHeader.Enabled = False
        
        mdlProcedures.SetControlMode Me, objMode, False, , Me.cmbMonth.Name & " | " & Me.txtYear.Name
        
        SetModeDetail
        
        Me.txtNotes.SetFocus
    Else
        Me.fraHeader.Enabled = False
    End If
    
    If rstMain.RecordCount > 0 Then
        With Me.tlbMain
            .Buttons(UpdateButton).Enabled = True
            .Buttons(DeleteButton).Enabled = True
            .Buttons(PrintButton).Enabled = True
            .Buttons(BrowseButton).Enabled = True
        End With
        
        If objMode = ViewMode Then
            mdlProcedures.SetControlMode Me, objMode, False, , Me.cmbMonth.Name & " | " & Me.txtYear.Name
        End If
    Else
        With Me.tlbMain
            .Buttons(UpdateButton).Enabled = False
            .Buttons(DeleteButton).Enabled = False
            .Buttons(PrintButton).Enabled = False
            .Buttons(BrowseButton).Enabled = False
        End With
        
        If objMode = ViewMode Then
            mdlProcedures.SetControlMode Me, objMode, , , Me.cmbMonth.Name & " | " & Me.txtYear.Name
        End If
    End If
    
    Me.Caption = mdlProcedures.SetCaptionMode(strFormCaption, objMode)
End Sub

Private Sub PrintFunction()
    If Not PrintTransaction Is Nothing Then
        Set PrintTransaction = Nothing
    End If

    Set PrintTransaction = New clsPRTTHMUTITEM
    
    PrintTransaction.ImportToExcel rstMain!MutId
End Sub

Private Sub DeleteFunction()
    objMode = DeleteMode
    
    If mdlProcedures.SetMsgYesNo("Apakah Anda Yakin ?", mdlProcedures.SetCaptionMode(strFormCaption, objMode)) Then
        mdlTHMUTITEM.DeleteTHMUTITEM rstMain
        
        mdlDatabase.DeleteRecordQuery mdlGlobal.conInventory, mdlTable.CreateTDMUTITEM, "MutId='" & rstMain!MutId & "'"
        mdlDatabase.DeleteSingleRecord rstMain
        
        frmMenu.SetRecycle
    End If
    
    objMode = ViewMode
    
    SetMode
    
    FillText
End Sub

Private Sub SaveFunction()
    If Not CheckValidation Then Exit Sub
    
    With rstMain
        Dim strWarehouseFrom As String
        Dim strWarehouseTo As String
        
        strWarehouseFrom = Me.WarehouseFromCombo
        strWarehouseTo = Me.WarehouseToCombo
        
        mdlDatabase.SearchRecordset rstMain, "MutId", strWarehouseFrom & strWarehouseTo & mdlProcedures.FormatDate(Me.dtpMutDate.Value, "ddMMyyyy")
        
        If .EOF Then
            .AddNew
            
            !MutId = strWarehouseFrom & strWarehouseTo & mdlProcedures.FormatDate(Me.dtpMutDate.Value, "ddMMyyyy")
            !MutDate = mdlProcedures.FormatDate(Me.dtpMutDate.Value)
            .Fields("WarehouseFrom").Value = strWarehouseFrom
            !WarehouseTo = strWarehouseTo
            
            !CreateId = Trim(mdlGlobal.UserAuthority.UserId)
            !CreateDate = mdlProcedures.FormatDate(Now)
        End If
        
        !Notes = mdlProcedures.RepDupText(Trim(Me.txtNotes.Text))
        
        !UpdateId = Trim(mdlGlobal.UserAuthority.UserId)
        !UpdateDate = mdlProcedures.FormatDate(Now)
        
        .Update
    End With
    
    SaveDetailFunction
    
    If objMode = AddMode Then
        mdlProcedures.SetControlMode Me, objMode, , , Me.cmbMonth.Name & " | " & Me.txtYear.Name
        
        Me.flxDetail.Rows = 1
        
        SetModeDetail
        
        Me.dtpMutDate.SetFocus
    ElseIf objMode = UpdateMode Then
        objMode = ViewMode
        
        SetMode
    End If
End Sub

Private Sub SaveDetailFunction()
    Dim strMutId As String
    Dim strWarehouseFrom As String
    Dim strWarehouseTo As String
    
    strWarehouseFrom = mdlProcedures.GetComboData(Me.cmbWarehouseFrom)
    strWarehouseTo = mdlProcedures.GetComboData(Me.cmbWarehouseTo)
    
    strMutId = strWarehouseFrom & strWarehouseTo & mdlProcedures.FormatDate(Me.dtpMutDate.Value, "ddMMyyyy")

    mdlDatabase.DeleteRecordQuery mdlGlobal.conInventory, mdlTable.CreateTDMUTITEM, "MutId='" & strMutId & "'"

    If Not Me.flxDetail.Rows > 1 Then Exit Sub
    
    Dim rstTemp As ADODB.Recordset
    
    Set rstTemp = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "*", mdlTable.CreateTDMUTITEM, , "MutId='" & strMutId & "'")
    
    With rstTemp
        If Not .RecordCount > 0 Then
            Dim lngRow As Long
            
            For lngRow = 1 To Me.flxDetail.Rows - 1
                .AddNew
                
                !MutId = strMutId
                !MutDtlId = !MutId & Trim(Me.flxDetail.TextMatrix(lngRow, ItemIdColumn))
                !ItemId = Trim(Me.flxDetail.TextMatrix(lngRow, ItemIdColumn))
                !Qty = mdlProcedures.GetCurrency(Trim(Me.flxDetail.TextMatrix(lngRow, QtyColumn)))
                
                !CreateId = Trim(mdlGlobal.UserAuthority.UserId)
                !CreateDate = mdlProcedures.FormatDate(Now)
                
                !UpdateId = Trim(mdlGlobal.UserAuthority.UserId)
                !UpdateDate = mdlProcedures.FormatDate(Now)
                
                .Update
            Next lngRow
        End If
    End With
    
    mdlDatabase.CloseRecordset rstTemp
End Sub

Private Function CheckValidation() As Boolean
    If Not mdlProcedures.IsValidComboData(Me.cmbWarehouseFrom) Then
        MsgBox "Gudang Asal Harap Dipilih", vbOKOnly + vbExclamation, mdlProcedures.SetCaptionMode(strFormCaption, objMode)
        
        Me.cmbWarehouseFrom.SetFocus
        
        CheckValidation = False
        
        Exit Function
    ElseIf Not mdlProcedures.IsValidComboData(Me.cmbWarehouseTo) Then
        MsgBox "Gudang Tujuan Harap Dipilih", vbOKOnly + vbExclamation, mdlProcedures.SetCaptionMode(strFormCaption, objMode)
        
        Me.cmbWarehouseTo.SetFocus
        
        CheckValidation = False
        
        Exit Function
    ElseIf Not Me.flxDetail.Rows > 1 Then
        MsgBox "Data Barang Harap Diisi", vbOKOnly + vbExclamation, mdlProcedures.SetCaptionMode(strFormCaption, objMode)
        
        Me.txtNotes.SetFocus
        
        CheckValidation = False
        
        Exit Function
    End If
    
    If objMode = AddMode Then
        With rstMain
            mdlDatabase.SearchRecordset rstMain, "MutId", Me.WarehouseFromCombo & Me.WarehouseToCombo & mdlProcedures.FormatDate(Me.dtpMutDate.Value, "ddMMyyyy")
            
            If Not .EOF Then
                MsgBox "Mutasi Sudah Ada", vbOKOnly + vbExclamation, mdlProcedures.SetCaptionMode(strFormCaption, objMode)
                
                Me.dtpMutDate.SetFocus
                
                CheckValidation = False
                
                Exit Function
            End If
        End With
    End If
    
    CheckValidation = True
End Function

Private Sub FillText()
    With rstMain
        If .RecordCount > 0 Then
            Me.dtpMutDate.Value = mdlProcedures.FormatDate(!MutDate, mdlGlobal.strFormatDate)
            
            Me.WarehouseFromCombo = .Fields("WarehouseFrom").Value
            Me.WarehouseToCombo = !WarehouseTo
            
            Me.txtNotes.Text = Trim(!Notes)
            
            FillGrid !MutId
        Else
            FillGrid
        End If
    End With
End Sub

Private Sub FillGrid(Optional ByVal strMutId As String = "")
    With Me.flxDetail
        .Rows = 1
        
        Dim rstTemp As ADODB.Recordset
        
        Set rstTemp = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "*", mdlTable.CreateTDMUTITEM, False, "MutId='" & strMutId & "'", "ItemId ASC")
        
        While Not rstTemp.EOF
            .Rows = .Rows + 1
            
            .TextMatrix(.Rows - 1, ItemIdColumn) = rstTemp!ItemId
            .TextMatrix(.Rows - 1, NameColumn) = mdlDatabase.GetFieldData(mdlGlobal.conInventory, "Name", mdlTable.CreateTMITEM, "ItemId='" & rstTemp!ItemId & "'")
            .TextMatrix(.Rows - 1, QtyColumn) = mdlProcedures.FormatCurrency(CStr(mdlProcedures.GetCurrency(rstTemp!Qty)))
            .TextMatrix(.Rows - 1, UnityIdColumn) = mdlDatabase.GetFieldData(mdlGlobal.conInventory, "UnityId", mdlTable.CreateTMITEM, "ItemId='" & rstTemp!ItemId & "'")
            
            .Row = .Rows - 1
            .Col = ItemIdColumn
            .ColSel = UnityIdColumn
            
            DoEvents
            
            rstTemp.MoveNext
        Wend
        
        If rstTemp.RecordCount > 0 Then .Row = 1
        
        mdlDatabase.CloseRecordset rstTemp
    End With
End Sub

Private Sub FillCombo()
    FillComboSearch
    
    Me.FillComboTMWAREHOUSE
End Sub

Private Sub FillComboSearch()
    mdlProcedures.FillComboMonth Me.cmbMonth, , , mdlProcedures.FormatDate(Now, "M")
    
    Me.txtYear.Text = mdlProcedures.FormatDate(Now, "yyyy")
End Sub

Public Sub FillComboTMWAREHOUSE()
    Dim rstTemp As ADODB.Recordset
    
    Set rstTemp = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "WarehouseId, Name", mdlTable.CreateTMWAREHOUSE, False, , "WarehouseSet DESC")
    
    mdlProcedures.FillComboData Me.cmbWarehouseFrom, rstTemp
    mdlProcedures.FillComboData Me.cmbWarehouseTo, rstTemp
    
    mdlDatabase.CloseRecordset rstTemp
End Sub

Private Sub ArrangeGrid()
    With Me.flxDetail
        .AllowUserResizing = flexResizeColumns
        
        .Rows = 1
        .Cols = UnityIdColumn + 1
        
        .ColWidth(BlankColumn) = 300
        .ColWidth(ItemIdColumn) = 1100
        .ColWidth(NameColumn) = 5000
        .ColWidth(QtyColumn) = 1500
        .ColWidth(UnityIdColumn) = 1200
        
        .ColAlignment(ItemIdColumn) = flexAlignLeftCenter
        .ColAlignment(NameColumn) = flexAlignLeftCenter
        .ColAlignment(QtyColumn) = flexAlignRightCenter
        .ColAlignment(UnityIdColumn) = flexAlignLeftCenter
        
        .TextMatrix(0, ItemIdColumn) = "Kode"
        .TextMatrix(0, NameColumn) = "Nama"
        .TextMatrix(0, QtyColumn) = "Qty"
        .TextMatrix(0, UnityIdColumn) = "Satuan"
    End With
End Sub

Public Function SaveDetail( _
    ByVal strItemId As String, _
    ByVal curQty As Currency) As Boolean
    With Me.flxDetail
        Dim blnExist As Boolean
        
        blnExist = False
    
        If .Rows > 1 Then
            If mdlProcedures.IsDataExistsInFlex(Me.flxDetail, strItemId, , , , , True) Then blnExist = True
        End If
        
        If Not blnExist Then
            .Rows = .Rows + 1
            
            .TextMatrix(.Rows - 1, ItemIdColumn) = strItemId
            .TextMatrix(.Rows - 1, NameColumn) = mdlDatabase.GetFieldData(mdlGlobal.conInventory, "Name", mdlTable.CreateTMITEM, "ItemId='" & strItemId & "'")
            .TextMatrix(.Rows - 1, QtyColumn) = mdlProcedures.FormatCurrency(CStr(curQty))
            .TextMatrix(.Rows - 1, UnityIdColumn) = mdlDatabase.GetFieldData(mdlGlobal.conInventory, "UnityId", mdlTable.CreateTMITEM, "ItemId='" & strItemId & "'")
            
            .Row = .Rows - 1
            .Col = ItemIdColumn
            .ColSel = UnityIdColumn
            
            SetModeDetail
        End If
    End With
End Function

Private Sub SetModeDetail()
    If Me.flxDetail.Rows > 1 Then
        blnDetailFill = True
    Else
        blnDetailFill = False
    End If

    With Me.tlbDetail
        If blnDetailFill Then
            .Buttons(UpdateDetailButton).Enabled = True
            .Buttons(DeleteDetailButton).Enabled = True
        Else
            .Buttons(UpdateDetailButton).Enabled = False
            .Buttons(DeleteDetailButton).Enabled = False
        End If
    End With
End Sub

Public Property Get Parent() As Boolean
    Parent = blnParent
End Property

Public Property Get ToBoolean() As Boolean
    ToBoolean = blnTo
End Property

Public Property Get FromBoolean() As Boolean
    FromBoolean = blnFrom
End Property

Public Property Get MonthTrans() As String
    MonthTrans = mdlProcedures.GetComboData(Me.cmbMonth)
End Property

Public Property Get YearTrans() As String
    YearTrans = Trim(Me.txtYear.Text)
End Property

Public Property Get MutId() As String
    If rstMain Is Nothing Then Exit Property

    If rstMain.RecordCount > 0 Then
        MutId = rstMain!MutId
    End If
End Property

Public Property Get WarehouseFromCombo() As String
    WarehouseFromCombo = mdlProcedures.GetComboData(Me.cmbWarehouseFrom)
End Property

Public Property Get WarehouseToCombo() As String
    WarehouseToCombo = mdlProcedures.GetComboData(Me.cmbWarehouseTo)
End Property

Public Property Let Parent(ByVal blnEnable As Boolean)
    blnParent = blnEnable
    
    If blnParent = False Then mdlProcedures.CenterWindows Me
End Property

Public Property Let MutId(ByVal strMutId As String)
    If rstMain Is Nothing Then Exit Property
    
    mdlDatabase.SearchRecordset rstMain, "MutId", strMutId
    
    If Not rstMain.EOF Then FillText
End Property

Public Property Let WarehouseFromCombo(ByVal strWarehouseId As String)
    mdlProcedures.SetComboData Me.cmbWarehouseFrom, strWarehouseId
End Property

Public Property Let WarehouseToCombo(ByVal strWarehouseId As String)
    mdlProcedures.SetComboData Me.cmbWarehouseTo, strWarehouseId
End Property
