VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmTMCONTACTCUSTOMER 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6525
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8775
   Icon            =   "frmTMCONTACTCUSTOMER.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6525
   ScaleWidth      =   8775
   ShowInTaskbar   =   0   'False
   Begin MSDataGridLib.DataGrid dgdMain 
      Height          =   2535
      Left            =   120
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   3840
      Width           =   8535
      _ExtentX        =   15055
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
   Begin VB.Frame fraHeader 
      Height          =   615
      Left            =   120
      TabIndex        =   12
      Top             =   1800
      Width           =   8535
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
         MaxLength       =   50
         TabIndex        =   0
         Top             =   240
         Width           =   6135
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
         TabIndex        =   7
         Top             =   240
         Width           =   510
      End
   End
   Begin VB.Frame fraDetail 
      Height          =   1335
      Left            =   120
      TabIndex        =   13
      Top             =   2400
      Width           =   8535
      Begin VB.TextBox txtEmail 
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
         MaxLength       =   50
         TabIndex        =   3
         Top             =   960
         Width           =   6135
      End
      Begin VB.TextBox txtHandPhone 
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
         MaxLength       =   50
         TabIndex        =   2
         Top             =   600
         Width           =   6135
      End
      Begin VB.TextBox txtPhone 
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
         MaxLength       =   50
         TabIndex        =   1
         Top             =   240
         Width           =   6135
      End
      Begin VB.Label lblEmail 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Email"
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
         Top             =   960
         Width           =   495
      End
      Begin VB.Label lblHandPhone 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "HP"
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
         Top             =   600
         Width           =   270
      End
      Begin VB.Label lblPhone 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Telepon"
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
         Top             =   240
         Width           =   675
      End
   End
   Begin VB.Frame fraParent 
      Height          =   615
      Left            =   120
      TabIndex        =   11
      Top             =   1200
      Width           =   8535
      Begin VB.Label txtCustomerId 
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
         Left            =   1200
         TabIndex        =   6
         Top             =   240
         Width           =   7215
      End
      Begin VB.Label lblCustomerId 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Customer"
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
         Width           =   840
      End
   End
   Begin MSComctlLib.Toolbar tlbMain 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   1111
      ButtonWidth     =   609
      ButtonHeight    =   953
      Appearance      =   1
      _Version        =   393216
      Begin MSComctlLib.ImageList imlMain 
         Left            =   8160
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   48
         ImageHeight     =   48
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   6
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTMCONTACTCUSTOMER.frx":1F8A
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTMCONTACTCUSTOMER.frx":43DC
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTMCONTACTCUSTOMER.frx":682E
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTMCONTACTCUSTOMER.frx":8C80
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTMCONTACTCUSTOMER.frx":B0D4
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTMCONTACTCUSTOMER.frx":D526
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "frmTMCONTACTCUSTOMER"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum ButtonMode
    [AddButton] = 1
    [UpdateButton]
    [DeleteButton]
    [NotesButton]
    [SaveButton]
    [CancelButton]
End Enum

Private rstMain As ADODB.Recordset

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
        If frmTMCUSTOMER.Parent Then
            frmTMCUSTOMER.Parent = False
        End If
        
        mdlDatabase.CloseRecordset rstMain
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmTMCONTACTCUSTOMER = Nothing
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
        Case NotesButton:
            If blnParent Then Exit Sub
            
            blnParent = True
            
            mdlProcedures.ShowForm frmTMCONTACTNOTES, False, True
        Case SaveButton:
            SaveFunction
        Case CancelButton:
            objMode = ViewMode
            
            SetMode
            
            FillText
    End Select
End Sub

Private Sub txtName_GotFocus()
    mdlProcedures.GotFocus Me.txtName
End Sub

Private Sub txtPhone_GotFocus()
    mdlProcedures.GotFocus Me.txtPhone
End Sub

Private Sub txtHandPhone_GotFocus()
    mdlProcedures.GotFocus Me.txtHandPhone
End Sub

Private Sub txtEmail_GotFocus()
    mdlProcedures.GotFocus Me.txtEmail
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

Private Sub SetInitialization()
    mdlProcedures.CornerWindows Me, False
    
    With Me.tlbMain
        .AllowCustomize = False
        
        .ImageList = Me.imlMain
        
        .Buttons.Add AddButton, , "Tambah", , AddButton
        .Buttons.Add UpdateButton, , "Ubah", , UpdateButton
        .Buttons.Add DeleteButton, , "Hapus", , DeleteButton
        .Buttons.Add NotesButton, , "Catatan", , NotesButton
        .Buttons.Add SaveButton, , "Simpan", , SaveButton
        .Buttons.Add CancelButton, , "Batal", , CancelButton
    End With
    
    Me.txtCustomerId.Caption = frmTMCUSTOMER.CustomerId & " | " & frmTMCUSTOMER.CustomerName
        
    strFormCaption = mdlText.strTMCONTACTCUSTOMER
    
    blnParent = False
    
    Set rstMain = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "*", mdlTable.CreateTMCONTACTCUSTOMER, , "CustomerId='" & frmTMCUSTOMER.CustomerId & "'", "Name ASC")
    
    ArrangeGrid
    
    objMode = ViewMode
    
    SetMode
    
    FillText
End Sub

Private Sub ArrangeGrid()
    Set Me.dgdMain.DataSource = rstMain
    
    With Me.dgdMain
        .RowHeight = 1000
        
        .Columns(2).Width = 3000
        .Columns(2).Locked = True
        .Columns(2).Caption = "Nama"
        .Columns(2).WrapText = True
        .Columns(3).Width = 1600
        .Columns(3).Locked = True
        .Columns(3).Caption = "Telepon"
        .Columns(3).WrapText = True
        .Columns(4).Width = 1600
        .Columns(4).Locked = True
        .Columns(4).Caption = "HP"
        .Columns(4).WrapText = True
        .Columns(5).Width = 1600
        .Columns(5).Locked = True
        .Columns(5).Caption = "Email"
        .Columns(5).WrapText = True
        
        Dim intCounter As Integer
        
        For intCounter = 0 To 1
            .Columns(intCounter).Width = 0
            .Columns(intCounter).Visible = False
            .Columns(intCounter).Locked = True
        Next intCounter
        
        For intCounter = 6 To .Columns.Count - 1
            .Columns(intCounter).Width = 0
            .Columns(intCounter).Visible = False
            .Columns(intCounter).Locked = True
        Next intCounter
    End With
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
    
    Me.dgdMain.Enabled = blnFront
    
    With Me.tlbMain
        .Buttons(AddButton).Visible = blnFront
        .Buttons(UpdateButton).Visible = blnFront
        .Buttons(DeleteButton).Visible = blnFront
        .Buttons(NotesButton).Visible = blnFront
        .Buttons(SaveButton).Visible = blnBack
        .Buttons(CancelButton).Visible = blnBack
    End With
    
    Me.fraDetail.Enabled = blnBack
    
    If objMode = AddMode Then
        Me.fraHeader.Enabled = True
    
        mdlProcedures.SetControlMode Me, objMode
        
        Me.txtName.SetFocus
    ElseIf objMode = UpdateMode Then
        Me.fraHeader.Enabled = False
        
        mdlProcedures.SetControlMode Me, objMode, False, Me.txtName.Name
        
        Me.txtPhone.SetFocus
    Else
        Me.fraHeader.Enabled = False
    End If
    
    If rstMain.RecordCount > 0 Then
        With Me.tlbMain
            .Buttons(UpdateButton).Enabled = True
            .Buttons(DeleteButton).Enabled = True
            .Buttons(NotesButton).Enabled = True
        End With
        
        If objMode = ViewMode Then
            mdlProcedures.SetControlMode Me, objMode, False
        End If
    Else
        With Me.tlbMain
            .Buttons(UpdateButton).Enabled = False
            .Buttons(DeleteButton).Enabled = False
            .Buttons(NotesButton).Enabled = False
        End With
        
        If objMode = ViewMode Then
            mdlProcedures.SetControlMode Me, objMode
        End If
    End If
    
    Me.Caption = mdlProcedures.SetCaptionMode(strFormCaption, objMode)
End Sub

Private Sub DeleteFunction()
    objMode = DeleteMode
    
    If mdlProcedures.SetMsgYesNo("Apakah Anda Yakin ?", mdlProcedures.SetCaptionMode(strFormCaption, objMode)) Then
        mdlDatabase.DeleteRecordQuery mdlGlobal.conInventory, mdlTable.CreateTMCONTACTNOTES, "ContactId='" & rstMain!ContactId & "'"
        mdlDatabase.DeleteSingleRecord rstMain
    End If
    
    objMode = ViewMode
    
    SetMode
    
    FillText
End Sub

Private Sub SaveFunction()
    If Not CheckValidation Then Exit Sub
    
    With rstMain
        mdlDatabase.SearchRecordset rstMain, "ContactId", frmTMCUSTOMER.CustomerId & mdlProcedures.RepDupText(Trim(Me.txtName.Text))
        
        If .EOF Then
            .AddNew
            
            !ContactId = frmTMCUSTOMER.CustomerId & mdlProcedures.RepDupText(Trim(Me.txtName.Text))
            !CustomerId = Trim(frmTMCUSTOMER.CustomerId)
            !Name = mdlProcedures.RepDupText(Trim(Me.txtName.Text))
            
            !CreateId = Trim(mdlGlobal.UserAuthority.UserId)
            !CreateDate = mdlProcedures.FormatDate(Now)
        End If
        
        !Phone = mdlProcedures.RepDupText(Trim(Me.txtPhone.Text))
        !HandPhone = mdlProcedures.RepDupText(Trim(Me.txtHandPhone.Text))
        !Email = mdlProcedures.RepDupText(Trim(Me.txtEmail.Text))
        
        !UpdateId = Trim(mdlGlobal.UserAuthority.UserId)
        !UpdateDate = mdlProcedures.FormatDate(Now)
        
        .Update
    End With
    
    If objMode = AddMode Then
        mdlProcedures.SetControlMode Me, objMode
        
        Me.txtName.SetFocus
    ElseIf objMode = UpdateMode Then
        objMode = ViewMode
        
        SetMode
    End If
End Sub

Private Function CheckValidation() As Boolean
    If mdlProcedures.RepDupText(Trim(Me.txtName.Text)) = "" Then
        MsgBox "Nama Harap Diisi", vbOKOnly + vbExclamation, mdlProcedures.SetCaptionMode(strFormCaption, objMode)
        
        Me.txtName.SetFocus
        
        CheckValidation = False
        
        Exit Function
    End If
    
    If objMode = AddMode Then
        With rstMain
            mdlDatabase.SearchRecordset rstMain, "ContactId", frmTMCUSTOMER.CustomerId & mdlProcedures.RepDupText(Trim(Me.txtName.Text))
            
            If Not .EOF Then
                MsgBox "Nama Sudah Ada", vbOKOnly + vbExclamation, mdlProcedures.SetCaptionMode(strFormCaption, objMode)
                
                Me.txtName.SetFocus
                
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
            Me.txtName.Text = Trim(!Name)
            Me.txtPhone.Text = Trim(!Phone)
            Me.txtHandPhone.Text = Trim(!HandPhone)
            Me.txtEmail.Text = Trim(!Email)
        End If
    End With
End Sub

Public Property Get ContactId() As String
    If rstMain Is Nothing Then Exit Property
    
    If rstMain.RecordCount > 0 Then
        ContactId = rstMain!ContactId
    End If
End Property

Public Property Get ContactName() As String
    If rstMain Is Nothing Then Exit Property
    
    If rstMain.RecordCount > 0 Then
        ContactName = rstMain!Name
    End If
End Property

Public Property Get Parent() As Boolean
    Parent = blnParent
End Property

Public Property Let Parent(ByVal blnEnable As Boolean)
    blnParent = blnEnable
End Property
