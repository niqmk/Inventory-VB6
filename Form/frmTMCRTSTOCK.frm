VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmTMCRTSTOCK 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7095
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9015
   Icon            =   "frmTMCRTSTOCK.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7095
   ScaleWidth      =   9015
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraSearch 
      Height          =   735
      Left            =   120
      TabIndex        =   11
      Top             =   1200
      Width           =   8775
      Begin VB.ComboBox cmbWarehouseId 
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
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   240
         Width           =   7335
      End
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
         TabIndex        =   5
         Top             =   240
         Width           =   675
      End
   End
   Begin MSDataGridLib.DataGrid dgdMain 
      Height          =   3135
      Left            =   120
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   2040
      Width           =   8700
      _ExtentX        =   15346
      _ExtentY        =   5530
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
   Begin VB.Frame fraDetail 
      Height          =   615
      Left            =   120
      TabIndex        =   13
      Top             =   6360
      Width           =   8775
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
         Left            =   1320
         MaxLength       =   10
         TabIndex        =   2
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lblQty 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah Stok"
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
         Width           =   1080
      End
   End
   Begin VB.Frame fraHeader 
      Height          =   1095
      Left            =   120
      TabIndex        =   12
      Top             =   5160
      Width           =   8775
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
         Left            =   1320
         MaxLength       =   7
         TabIndex        =   1
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdItemId 
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
         Left            =   2520
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   240
         Width           =   375
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
         Left            =   6840
         TabIndex        =   8
         Top             =   240
         Width           =   615
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
         Left            =   5880
         TabIndex        =   7
         Top             =   240
         Width           =   615
      End
      Begin VB.Label txtName 
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
         Left            =   1320
         TabIndex        =   9
         Top             =   720
         Width           =   6135
      End
      Begin VB.Label lblItemId 
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
         TabIndex        =   6
         Top             =   240
         Width           =   615
      End
   End
   Begin MSComctlLib.Toolbar tlbMain 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   1111
      ButtonWidth     =   609
      ButtonHeight    =   953
      Appearance      =   1
      _Version        =   393216
      Begin MSComctlLib.ImageList imlMain 
         Left            =   8400
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   48
         ImageHeight     =   48
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   8
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTMCRTSTOCK.frx":1F8A
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTMCRTSTOCK.frx":43DC
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTMCRTSTOCK.frx":682E
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTMCRTSTOCK.frx":8C80
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTMCRTSTOCK.frx":B0D2
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTMCRTSTOCK.frx":B674
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTMCRTSTOCK.frx":FDAE
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTMCRTSTOCK.frx":12200
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "frmTMCRTSTOCK"
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

Private Enum ColumnConstants
    [BlankColumn]
    [StockIdColumn]
    [ItemIdColumn]
    [NameColumn]
    [QtyColumn]
End Enum

Private rstMain As ADODB.Recordset

Private PrintMaster As clsPRTTMCRTSTOCK

Private objMode As FunctionMode

Private strFormCaption As String

Private blnParent As Boolean

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub Form_Load()
    SetInitialization
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If blnParent Then
        Cancel = 1
    Else
        If Not PrintMaster Is Nothing Then
            Set PrintMaster = Nothing
        End If
        
        mdlDatabase.CloseRecordset rstMain
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmTMCRTSTOCK = Nothing
End Sub

Private Sub tlbMain_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case AddButton:
            If Not mdlGlobal.UserAuthority.IsAdmin Then Exit Sub
            
            If Not mdlProcedures.IsValidComboData(Me.cmbWarehouseId) Then Exit Sub
        
            objMode = AddMode
            
            SetMode
        Case UpdateButton:
            If Not mdlGlobal.UserAuthority.IsAdmin Then Exit Sub
            
            If Not mdlProcedures.IsValidComboData(Me.cmbWarehouseId) Then Exit Sub
            
            objMode = UpdateMode
            
            SetMode
        Case DeleteButton:
            If Not mdlGlobal.UserAuthority.IsAdmin Then Exit Sub
            
            DeleteFunction
        Case PrintButton:
            PrintFunction
        Case BrowseButton:
            If blnParent Then Exit Sub
            
            blnParent = True
            
            mdlProcedures.CornerWindows Me
            
            mdlProcedures.ShowForm frmBRWTMCRTSTOCK, False
        Case SaveButton:
            SaveFunction
        Case CancelButton:
            objMode = ViewMode
            
            SetMode
            
            FillText
    End Select
End Sub

Private Sub cmbWarehouseId_Click()
    If Not objMode = ViewMode Then Exit Sub
    
    SetRecordset
End Sub

Private Sub dgdMain_HeadClick(ByVal ColIndex As Integer)
    If Not objMode = ViewMode Then Exit Sub
    
    rstMain.Sort = rstMain.Fields(ColIndex).Name
    
    If rstMain.RecordCount > 0 Then
        FillText
    End If
End Sub

Private Sub dgdMain_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If Not objMode = ViewMode Then Exit Sub
    
    If rstMain.RecordCount > 0 Then
        FillText
    End If
End Sub

Private Sub txtItemId_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Trim(Me.txtItemId.Text) = "" Then
            If blnParent Then Exit Sub
            
            blnParent = True
            
            mdlProcedures.CornerWindows Me
            
            mdlProcedures.ShowForm frmBRWTMITEM, False, True
        End If
    End If
End Sub

Private Sub txtItemId_GotFocus()
    mdlProcedures.GotFocus Me.txtItemId
End Sub

Private Sub txtQty_GotFocus()
    mdlProcedures.GotFocus Me.txtQty
End Sub

Private Sub txtItemId_Validate(Cancel As Boolean)
    If Trim(Me.txtItemId.Text) = "" Then
        Me.txtName.Caption = ""
        Me.txtUnityId.Caption = ""
    Else
        Me.txtName.Caption = mdlDatabase.GetFieldData(mdlGlobal.conApp, "Name", mdlTable.CreateTMITEM, "ItemId='" & mdlProcedures.RepDupText(Me.txtItemId.Text) & "'")
        Me.txtUnityId.Caption = mdlDatabase.GetFieldData(mdlGlobal.conApp, "UnityId", mdlTable.CreateTMITEM, "ItemId='" & mdlProcedures.RepDupText(Me.txtItemId.Text) & "'")
    End If
End Sub

Private Sub txtQty_Change()
    Me.txtQty.Text = mdlProcedures.FormatCurrency(Me.txtQty.Text)
    
    Me.txtQty.SelStart = Len(Me.txtQty.Text)
End Sub

Private Sub cmdItemId_Click()
    If blnParent Then Exit Sub
    
    blnParent = True
    
    mdlProcedures.CornerWindows Me
    
    mdlProcedures.ShowForm frmBRWTMITEM, False, True
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
    
    FillCombo
    
    strFormCaption = mdlText.strTMCRTSTOCK
    
    blnParent = False
    
    If Me.cmbWarehouseId.ListCount > 0 Then
        Me.cmbWarehouseId.ListIndex = 0
    Else
        SetRecordset
    End If
End Sub

Private Sub FillCombo()
    FillComboTMWAREHOUSE
End Sub

Private Sub FillComboTMWAREHOUSE()
    Dim rstTemp As ADODB.Recordset
    
    Set rstTemp = mdlDatabase.OpenRecordset(mdlGlobal.conApp, "WarehouseId, Name", mdlTable.CreateTMWAREHOUSE)
    
    mdlProcedures.FillComboData Me.cmbWarehouseId, rstTemp
    
    mdlDatabase.CloseRecordset rstTemp
End Sub

Private Sub SetRecordset()
    Set rstMain = mdlDatabase.OpenRecordset(mdlGlobal.conApp, "*", mdlTable.CreateTMCRTSTOCK, , "WarehouseId='" & mdlProcedures.RepDupText(mdlProcedures.GetComboData(Me.cmbWarehouseId)) & "'", "ItemId ASC")
    
    ArrangeGrid
    
    SetMode
    
    objMode = ViewMode
    
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
    Me.dgdMain.Enabled = blnFront
    
    With Me.tlbMain
        .Buttons(AddButton).Visible = blnFront
        .Buttons(UpdateButton).Visible = blnFront
        .Buttons(DeleteButton).Visible = blnFront
        .Buttons(PrintButton).Visible = blnFront
        .Buttons(BrowseButton).Visible = blnFront
        .Buttons(SaveButton).Visible = blnBack
        .Buttons(CancelButton).Visible = blnBack
    End With
    
    Me.fraDetail.Enabled = blnBack
    
    If objMode = AddMode Then
        Me.fraHeader.Enabled = True
        
        Me.txtName.Caption = ""
    
        mdlProcedures.SetControlMode Me, objMode, , , Me.cmbWarehouseId.Name
        
        Me.txtItemId.SetFocus
    ElseIf objMode = UpdateMode Then
        Me.fraHeader.Enabled = False
        
        mdlProcedures.SetControlMode Me, objMode, False, Me.txtItemId.Name, Me.cmbWarehouseId.Name
        
        Me.txtQty.SetFocus
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
            mdlProcedures.SetControlMode Me, objMode, False, , Me.cmbWarehouseId.Name
        End If
    Else
        With Me.tlbMain
            .Buttons(UpdateButton).Enabled = False
            .Buttons(DeleteButton).Enabled = False
            .Buttons(PrintButton).Enabled = False
            .Buttons(BrowseButton).Enabled = False
        End With
        
        If objMode = ViewMode Then
            Me.txtName.Caption = ""
            
            mdlProcedures.SetControlMode Me, objMode, , , Me.cmbWarehouseId.Name
        End If
    End If
    
    Me.Caption = mdlProcedures.SetCaptionMode(strFormCaption, objMode)
End Sub

Private Sub PrintFunction()
    If Not PrintMaster Is Nothing Then
        Set PrintMaster = Nothing
    End If

    Set PrintMaster = New clsPRTTMCRTSTOCK
    
    Dim rstTemp As ADODB.Recordset
    
    Set rstTemp = mdlDatabase.OpenRecordset(mdlGlobal.conApp, "*", mdlTable.CreateTMCRTSTOCK, , "WarehouseId='" & mdlProcedures.RepDupText(mdlProcedures.GetComboData(Me.cmbWarehouseId)) & "'", "ItemId ASC")
    
    PrintMaster.ImportToExcel rstTemp
    
    mdlDatabase.CloseRecordset rstTemp
End Sub

Private Sub DeleteFunction()
    objMode = DeleteMode
    
    If mdlProcedures.SetMsgYesNo("Apakah Anda Yakin ?", mdlProcedures.SetCaptionMode(strFormCaption, objMode)) Then
        mdlDatabase.DeleteSingleRecord rstMain
    End If
    
    objMode = ViewMode
    
    SetMode
    
    FillText
End Sub

Private Sub SaveFunction()
    If Not CheckValidation Then Exit Sub
    
    With rstMain
        Dim mItemId As String
        Dim mWarehouseId As String
        
        mItemId = mdlProcedures.RepDupText(Trim(Me.txtItemId.Text))
        mItemId = _
            mItemId & _
            Space(mdlDatabase.GetColumnSize(mdlGlobal.conApp, "ItemId", mdlTable.CreateTMCRTSTOCK) - Len(mItemId))
        
        mWarehouseId = mdlProcedures.GetComboData(Me.cmbWarehouseId)
    
        mdlDatabase.SearchRecordset rstMain, "StockId", mItemId & mWarehouseId
        
        If .EOF Then
            .AddNew
            
            !StockId = mItemId & mWarehouseId
            
            !ItemId = mItemId
            !WarehouseId = mWarehouseId
            
            !CreateId = Trim(mdlGlobal.UserAuthority.UserId)
            !CreateDate = mdlProcedures.FormatDate(Now)
        End If
        
        !Qty = mdlProcedures.GetCurrency(Me.txtQty.Text)
        
        !UpdateId = Trim(mdlGlobal.UserAuthority.UserId)
        !UpdateDate = mdlProcedures.FormatDate(Now)
        
        .Update
    End With
    
    If objMode = AddMode Then
        Me.txtName.Caption = ""
        
        mdlProcedures.SetControlMode Me, objMode, , , Me.cmbWarehouseId.Name
        
        Me.txtItemId.SetFocus
    ElseIf objMode = UpdateMode Then
        objMode = ViewMode
        
        SetMode
    End If
End Sub

Private Function CheckValidation() As Boolean
    If Not mdlProcedures.IsValidComboData(Me.cmbWarehouseId) Then
        MsgBox "Gudang Harap Diisi", vbOKOnly + vbExclamation, mdlProcedures.SetCaptionMode(strFormCaption, objMode)
        
        CheckValidation = False
        
        Exit Function
    ElseIf mdlProcedures.RepDupText(Trim(Me.txtItemId.Text)) = "" Then
        MsgBox "Kode Harap Diisi", vbOKOnly + vbExclamation, mdlProcedures.SetCaptionMode(strFormCaption, objMode)
        
        Me.txtItemId.SetFocus

        CheckValidation = False

        Exit Function
    ElseIf Not mdlProcedures.GetCurrency(Trim(Me.txtQty.Text)) > 0 Then
        MsgBox "Qty Harap Diisi", vbOKOnly + vbExclamation, mdlProcedures.SetCaptionMode(strFormCaption, objMode)
        
        Me.txtQty.SetFocus
        
        CheckValidation = False
        
        Exit Function
    End If
    
    If objMode = AddMode Then
        Dim mItemId As String
        Dim mWarehouseId As String
        
        mItemId = mdlProcedures.RepDupText(Trim(Me.txtItemId.Text))
        mItemId = mItemId & Space(mdlDatabase.GetColumnSize(mdlGlobal.conApp, "ItemId", mdlTable.CreateTMITEM) - Len(mItemId))
        
        mWarehouseId = mdlProcedures.GetComboData(Me.cmbWarehouseId)
        
        If Not mdlDatabase.IsDataExists(mdlGlobal.conApp, mdlTable.CreateTMITEM, "ItemId='" & mItemId & "'") Then
            MsgBox "Kode Tidak Terdaftar", vbOKOnly + vbExclamation, mdlProcedures.SetCaptionMode(strFormCaption, objMode)
            
            Me.txtItemId.SetFocus
            
            CheckValidation = False
            
            Exit Function
        End If
        
        If mdlDatabase.IsDataExists(mdlGlobal.conApp, mdlTable.CreateTMCRTSTOCK, "StockId='" & mItemId & mWarehouseId & "'") Then
            MsgBox "Kode Sudah Ada", vbOKOnly + vbExclamation, mdlProcedures.SetCaptionMode(strFormCaption, objMode)
            
            Me.txtItemId.SetFocus
            
            CheckValidation = False
            
            Exit Function
        End If
    End If
    
    CheckValidation = True
End Function

Private Sub FillText()
    With rstMain
        If .RecordCount > 0 Then
            Me.txtItemId.Text = Trim(!ItemId)
            Me.txtName.Caption = mdlDatabase.GetFieldData(mdlGlobal.conApp, "Name", mdlTable.CreateTMITEM, "ItemId='" & !ItemId & "'")
            Me.txtUnityId.Caption = mdlDatabase.GetFieldData(mdlGlobal.conApp, "UnityId", mdlTable.CreateTMITEM, "ItemId='" & !ItemId & "'")
            Me.txtQty.Text = mdlProcedures.GetCurrency(CStr(!Qty))
        End If
    End With
End Sub

Private Sub ArrangeGrid()
    Set Me.dgdMain.DataSource = rstMain
    
    With Me.dgdMain
        .Columns(0).Width = 0
        .Columns(0).Locked = True
        .Columns(0).Visible = False
        .Columns(1).Width = 1000
        .Columns(1).Locked = True
        .Columns(1).Caption = "Kode"
        .Columns(2).Width = 0
        .Columns(2).Locked = True
        .Columns(2).Visible = False
        .Columns(3).Width = 1500
        .Columns(3).Locked = True
        .Columns(3).Caption = "Jumlah Stok"
        .Columns(3).Alignment = dbgRight
        .Columns(3).NumberFormat = "#,##0"
        
        Dim intCounter As Integer
        
        For intCounter = 4 To .Columns.Count - 1
            .Columns(intCounter).Width = 0
            .Columns(intCounter).Locked = True
            .Columns(intCounter).Visible = False
        Next intCounter
    End With
End Sub

Public Property Get Parent() As Boolean
    Parent = blnParent
End Property

Public Property Get ItemId() As String
    If rstMain Is Nothing Then Exit Property

    If rstMain.RecordCount > 0 Then
        ItemId = rstMain!ItemId
    End If
End Property

Public Property Get ItemName() As String
    If rstMain Is Nothing Then Exit Property

    If rstMain.RecordCount > 0 Then
        ItemName = mdlDatabase.GetFieldData(mdlGlobal.conApp, "Name", mdlTable.CreateTMITEM, "ItemId='" & rstMain!ItemId & "'")
    End If
End Property

Public Property Get WarehouseIdCombo() As String
    WarehouseIdCombo = mdlProcedures.GetComboData(Me.cmbWarehouseId)
End Property

Public Property Let Parent(ByVal blnEnable As Boolean)
    blnParent = blnEnable
    
    If Not blnEnable Then mdlProcedures.CenterWindows Me
End Property

Public Property Let ItemId(ByVal strItemId As String)
    If rstMain Is Nothing Then Exit Property
    
    mdlDatabase.SearchRecordset rstMain, "ItemId", strItemId
    
    If Not rstMain.EOF Then FillText
End Property

Public Property Let ItemIdText(ByVal strItemId As String)
    Me.txtItemId.Text = Trim(strItemId)
End Property
