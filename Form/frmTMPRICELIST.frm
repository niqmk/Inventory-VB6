VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmTMPRICELIST 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4470
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9510
   Icon            =   "frmTMPRICELIST.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4470
   ScaleWidth      =   9510
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraHeader 
      Height          =   1935
      Left            =   120
      TabIndex        =   14
      Top             =   1200
      Width           =   9255
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
         TabIndex        =   0
         Top             =   240
         Width           =   975
      End
      Begin MSComCtl2.DTPicker dtpPriceListDate 
         Height          =   375
         Left            =   1320
         TabIndex        =   1
         Top             =   1440
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
         Format          =   78446595
         CurrentDate     =   39335
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
      Begin VB.Label lblPriceListDate 
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
         Top             =   1440
         Width           =   675
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
         Height          =   240
         Left            =   1320
         TabIndex        =   10
         Top             =   1080
         Width           =   6060
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
         Top             =   1080
         Width           =   510
      End
      Begin VB.Label txtPartNumber 
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
         TabIndex        =   8
         Top             =   720
         Width           =   4935
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
         TabIndex        =   7
         Top             =   720
         Width           =   990
      End
   End
   Begin VB.Frame fraDetail 
      Height          =   1215
      Left            =   120
      TabIndex        =   15
      Top             =   3120
      Width           =   9255
      Begin VB.TextBox txtPriceListValue 
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
      Begin VB.ComboBox cmbCurrencyId 
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
         TabIndex        =   3
         Top             =   720
         Width           =   7335
      End
      Begin VB.CommandButton cmdCurrencyId 
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
         Left            =   8760
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   720
         Width           =   375
      End
      Begin VB.Label lblPriceListValue 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Harga"
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
         Top             =   240
         Width           =   510
      End
      Begin VB.Label lblCurrencyId 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mata Uang"
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
         Top             =   720
         Width           =   945
      End
   End
   Begin MSComctlLib.Toolbar tlbMain 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   9510
      _ExtentX        =   16775
      _ExtentY        =   1111
      ButtonWidth     =   609
      ButtonHeight    =   953
      Appearance      =   1
      _Version        =   393216
      Begin MSComctlLib.ImageList imlMain 
         Left            =   8880
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
               Picture         =   "frmTMPRICELIST.frx":1F8A
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTMPRICELIST.frx":43DC
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTMPRICELIST.frx":682E
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTMPRICELIST.frx":8C80
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTMPRICELIST.frx":B0D2
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTMPRICELIST.frx":B674
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTMPRICELIST.frx":DAC6
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "frmTMPRICELIST"
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

Private rstMain As ADODB.Recordset

Private PrintMaster As clsPRTTMPRICELIST

Private objMode As FunctionMode

Private strFormCaption As String

Private blnParent As Boolean
Private blnActivate As Boolean

Private Sub Form_Activate()
    If blnParent Then Exit Sub
    If blnActivate Then Exit Sub
    
    If rstMain.RecordCount > 0 Then
        blnParent = True
        
        mdlProcedures.CornerWindows Me
        
        mdlProcedures.ShowForm frmBRWTMPRICELIST, False, True
    End If
    
    blnActivate = True
End Sub

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
    Set frmTMPRICELIST = Nothing
End Sub

Private Sub tlbMain_ButtonClick(ByVal Button As MSComctlLib.Button)
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
            
            mdlProcedures.ShowForm frmBRWTMPRICELIST, False, True
        Case SaveButton:
            SaveFunction
        Case CancelButton:
            objMode = ViewMode
            
            SetMode
            
            FillText
    End Select
End Sub

Private Sub txtItemId_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Trim(Me.txtItemId.Text) = "" Then
            If blnParent Then Exit Sub
            
            blnParent = True
            
            mdlProcedures.CornerWindows Me
            
            mdlProcedures.ShowForm frmBRWTMITEM, False, True
        End If
    End If
End Sub

Private Sub cmbCurrencyId_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Not mdlProcedures.IsValidComboData(Me.cmbCurrencyId) Then
            If blnParent Then Exit Sub
            
            blnParent = True
            
            mdlProcedures.CornerWindows Me
            
            mdlProcedures.ShowForm frmBRWTMCURRENCY, False, True
        End If
    End If
End Sub

Private Sub txtItemId_GotFocus()
    mdlProcedures.GotFocus Me.txtItemId
End Sub

Private Sub txtPriceListValue_GotFocus()
    mdlProcedures.GotFocus Me.txtPriceListValue
End Sub

Private Sub txtItemId_Validate(Cancel As Boolean)
    If Trim(Me.txtItemId.Text) = "" Then
        Me.txtPartNumber.Caption = ""
        Me.txtName.Caption = ""
    Else
        Me.txtPartNumber.Caption = mdlDatabase.GetFieldData(mdlGlobal.conInventory, "PartNumber", mdlTable.CreateTMITEM, "ItemId='" & mdlProcedures.RepDupText(Me.txtItemId.Text) & "'")
        Me.txtName.Caption = mdlDatabase.GetFieldData(mdlGlobal.conInventory, "Name", mdlTable.CreateTMITEM, "ItemId='" & mdlProcedures.RepDupText(Me.txtItemId.Text) & "'")
    End If
End Sub

Private Sub txtPriceListValue_Change()
    Me.txtPriceListValue.Text = mdlProcedures.FormatCurrency(Me.txtPriceListValue.Text)
    
    Me.txtPriceListValue.SelStart = Len(Me.txtPriceListValue.Text)
End Sub

Private Sub cmdItemId_Click()
    If blnParent Then Exit Sub
    
    blnParent = True
    
    mdlProcedures.CornerWindows Me
    
    mdlProcedures.ShowForm frmBRWTMITEM, False, True
    
    Me.txtItemId.SetFocus
End Sub

Private Sub cmdCurrencyId_Click()
    If Not mdlGlobal.UserAuthority.IsMenuAccess(mdiMain.mnuTMCURRENCY.Name) Then Exit Sub

    If blnParent Then Exit Sub
    
    blnParent = True
    
    mdlProcedures.CornerWindows Me
    
    mdlProcedures.ShowForm frmTMCURRENCY, False
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
    
    Me.dtpPriceListDate.CustomFormat = mdlGlobal.strFormatDate
    
    FillCombo
    
    strFormCaption = mdlText.strTMPRICELIST
    
    blnParent = False
    blnActivate = False
    
    Set rstMain = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "*", mdlTable.CreateTMPRICELIST, , , "ItemId ASC")
    
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
    
        mdlProcedures.SetControlMode Me, objMode
        
        Me.txtPartNumber.Caption = ""
        Me.txtName.Caption = ""
        
        Me.txtItemId.SetFocus
    ElseIf objMode = UpdateMode Then
        Me.fraHeader.Enabled = False
        
        mdlProcedures.SetControlMode Me, objMode, False, Me.txtItemId.Name
        
        Me.txtPriceListValue.SetFocus
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
            mdlProcedures.SetControlMode Me, objMode, False
        End If
    Else
        With Me.tlbMain
            .Buttons(UpdateButton).Enabled = False
            .Buttons(DeleteButton).Enabled = False
            .Buttons(PrintButton).Enabled = False
            .Buttons(BrowseButton).Enabled = False
        End With
        
        If objMode = ViewMode Then
            Me.txtPartNumber.Caption = ""
            Me.txtName.Caption = ""
            
            mdlProcedures.SetControlMode Me, objMode
        End If
    End If
    
    Me.Caption = mdlProcedures.SetCaptionMode(strFormCaption, objMode)
End Sub

Private Sub FillCombo()
    Me.FillComboTMCURRENCY
End Sub

Public Sub FillComboTMCURRENCY()
    Dim rstTemp As ADODB.Recordset
    
    Set rstTemp = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "CurrencyId, Name", mdlTable.CreateTMCURRENCY, False)
    
    mdlProcedures.FillComboData Me.cmbCurrencyId, rstTemp
    
    mdlDatabase.CloseRecordset rstTemp
End Sub

Private Sub PrintFunction()
    If Not PrintMaster Is Nothing Then
        Set PrintMaster = Nothing
    End If

    Set PrintMaster = New clsPRTTMPRICELIST
    
    Dim rstTemp As ADODB.Recordset
    
    Set rstTemp = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "*", mdlTable.CreateTMPRICELIST, False, , "ItemId, PriceListDate")
    
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
        Dim strItemId As String
        Dim strPriceListDate As String
        
        strItemId = mdlProcedures.RepDupText(Trim(Me.txtItemId.Text))
        strItemId = _
            strItemId & _
            Space(mdlDatabase.GetColumnSize(mdlGlobal.conInventory, "ItemId", mdlTable.CreateTMPRICELIST) - Len(strItemId))
        
        strPriceListDate = mdlProcedures.FormatDate(Me.dtpPriceListDate, "ddMMyyyy")
        
        mdlDatabase.SearchRecordset rstMain, "PriceListId", strItemId & strPriceListDate
        
        If .EOF Then
            .AddNew
            
            !PriceListId = strItemId & strPriceListDate
            
            !ItemId = strItemId
            !PriceListDate = mdlProcedures.FormatDate(Me.dtpPriceListDate)
            
            !CreateId = Trim(mdlGlobal.UserAuthority.UserId)
            !CreateDate = mdlProcedures.FormatDate(Now)
        End If
        
        !PriceListValue = mdlProcedures.GetCurrency(Me.txtPriceListValue.Text)
        !CurrencyId = Me.CurrencyIdCombo
        
        !UpdateId = Trim(mdlGlobal.UserAuthority.UserId)
        !UpdateDate = mdlProcedures.FormatDate(Now)
        
        .Update
    End With
    
    If objMode = AddMode Then
        Me.txtPartNumber.Caption = ""
        Me.txtName.Caption = ""
        
        mdlProcedures.SetControlMode Me, objMode
        
        Me.txtItemId.SetFocus
    ElseIf objMode = UpdateMode Then
        objMode = ViewMode
        
        SetMode
    End If
End Sub

Private Function CheckValidation() As Boolean
    If mdlProcedures.RepDupText(Trim(Me.txtItemId.Text)) = "" Then
        MsgBox "Kode Barang Harap Diisi", vbOKOnly + vbExclamation, mdlProcedures.SetCaptionMode(strFormCaption, objMode)
        
        Me.txtItemId.SetFocus
        
        CheckValidation = False
        
        Exit Function
    ElseIf Not mdlProcedures.GetCurrency(Trim(Me.txtPriceListValue.Text)) > 0 Then
        MsgBox "Harga Harap Diisi", vbOKOnly + vbExclamation, mdlProcedures.SetCaptionMode(strFormCaption, objMode)
        
        Me.txtPriceListValue.SetFocus
        
        CheckValidation = False
        
        Exit Function
    End If
    
    If objMode = AddMode Then
        Dim strItemId As String
        
        strItemId = mdlProcedures.RepDupText(Trim(Me.txtItemId.Text))
        
        If Not mdlDatabase.IsDataExists(mdlGlobal.conInventory, mdlTable.CreateTMITEM, "ItemId='" & strItemId & "'") Then
            MsgBox "Kode Barang Tidak Terdaftar", vbOKOnly + vbExclamation, mdlProcedures.SetCaptionMode(strFormCaption, objMode)
            
            Me.txtItemId.SetFocus
            
            CheckValidation = False
            
            Exit Function
        End If
        
        Dim strPriceListDate As String
        
        strItemId = _
            strItemId & _
            Space(mdlDatabase.GetColumnSize(mdlGlobal.conInventory, "ItemId", mdlTable.CreateTMPRICELIST) - Len(strItemId))
        
        strPriceListDate = mdlProcedures.FormatDate(Me.dtpPriceListDate, "ddMMyyyy")
        
        With rstMain
            mdlDatabase.SearchRecordset rstMain, "PriceListId", strItemId & strPriceListDate
            
            If Not .EOF Then
                MsgBox "Kode dan Tanggal Sudah Ada", vbOKOnly + vbExclamation, mdlProcedures.SetCaptionMode(strFormCaption, objMode)
                
                Me.txtItemId.SetFocus
                
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
            Me.txtItemId.Text = Trim(!ItemId)
            Me.txtPartNumber.Caption = mdlDatabase.GetFieldData(mdlGlobal.conInventory, "PartNumber", mdlTable.CreateTMITEM, "ItemId='" & !ItemId & "'")
            Me.txtName.Caption = mdlDatabase.GetFieldData(mdlGlobal.conInventory, "Name", mdlTable.CreateTMITEM, "ItemId='" & !ItemId & "'")
            Me.dtpPriceListDate.Value = mdlProcedures.FormatDate(!PriceListDate, mdlGlobal.strFormatDate)
            Me.txtPriceListValue.Text = mdlProcedures.GetCurrency(!PriceListValue)
            
            mdlProcedures.SetComboData Me.cmbCurrencyId, !CurrencyId
        End If
    End With
End Sub

Public Property Get Parent() As Boolean
    Parent = blnParent
End Property

Public Property Get PriceListId() As String
    If rstMain Is Nothing Then Exit Property
    
    If rstMain.RecordCount > 0 Then
        PriceListId = rstMain!PriceListId
    End If
End Property

Public Property Get ItemIdText() As String
    ItemIdText = Me.txtItemId.Text
End Property

Public Property Get ItemId() As String
    If rstMain Is Nothing Then Exit Property

    If rstMain.RecordCount > 0 Then
        ItemId = rstMain!ItemId
    End If
End Property

Public Property Get PartNumberText() As String
    PartNumberText = Me.txtPartNumber.Caption
End Property

Public Property Get ItemNameText() As String
    ItemNameText = Me.txtName.Caption
End Property

Public Property Get PartNumber() As String
    If rstMain Is Nothing Then Exit Property
    
    If rstMain.RecordCount > 0 Then
        PartNumber = mdlDatabase.GetFieldData(mdlGlobal.conInventory, "PartNumber", mdlTable.CreateTMITEM, "ItemId='" & rstMain!ItemId & "'")
    End If
End Property

Public Property Get ItemName() As String
    If rstMain Is Nothing Then Exit Property

    If rstMain.RecordCount > 0 Then
        ItemName = mdlDatabase.GetFieldData(mdlGlobal.conInventory, "Name", mdlTable.CreateTMITEM, "ItemId='" & rstMain!ItemId & "'")
    End If
End Property

Public Property Get PriceListDate() As String
    PriceListDate = mdlProcedures.FormatDate(Me.dtpPriceListDate, "yyyy/MM/dd")
End Property

Public Property Get CurrencyIdCombo() As String
    CurrencyIdCombo = mdlProcedures.GetComboData(Me.cmbCurrencyId)
End Property

Public Property Let Parent(ByVal blnEnable As Boolean)
    blnParent = blnEnable
    
    If Not blnEnable Then mdlProcedures.CenterWindows Me
End Property

Public Property Let PriceListId(ByVal strPriceListId As String)
    If rstMain Is Nothing Then Exit Property
    
    mdlDatabase.SearchRecordset rstMain, "PriceListId", strPriceListId
    
    If Not rstMain.EOF Then FillText
End Property

Public Property Let ItemId(ByVal strItemId As String)
    Me.txtItemId.Text = Trim(strItemId)
End Property

Public Property Let CurrencyIdCombo(ByVal strCurrencyId As String)
    mdlProcedures.SetComboData Me.cmbCurrencyId, strCurrencyId
End Property
