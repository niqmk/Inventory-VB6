VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmTMDELIVERYCUSTOMER 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5910
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   12030
   Icon            =   "frmTMDELIVERYCUSTOMER.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5910
   ScaleWidth      =   12030
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraDetail 
      Height          =   3255
      Left            =   4320
      TabIndex        =   17
      Top             =   2520
      Width           =   7575
      Begin VB.TextBox txtAddress 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   885
         Left            =   1320
         MaxLength       =   150
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Top             =   600
         Width           =   6135
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
         Left            =   1320
         MaxLength       =   50
         TabIndex        =   0
         Top             =   240
         Width           =   6135
      End
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
         Height          =   885
         Left            =   1320
         MaxLength       =   150
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   2280
         Width           =   6135
      End
      Begin VB.TextBox txtFax 
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
         MaxLength       =   50
         TabIndex        =   3
         Top             =   1920
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
         Left            =   1320
         MaxLength       =   50
         TabIndex        =   2
         Top             =   1560
         Width           =   6135
      End
      Begin VB.Label lblAddress 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Alamat"
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
         Top             =   600
         Width           =   615
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
         TabIndex        =   10
         Top             =   240
         Width           =   510
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
         Top             =   2280
         Width           =   990
      End
      Begin VB.Label lblFax 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fax"
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
         Top             =   1920
         Width           =   330
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
         TabIndex        =   12
         Top             =   1560
         Width           =   675
      End
   End
   Begin VB.Frame fraParent 
      Height          =   615
      Left            =   120
      TabIndex        =   15
      Top             =   1200
      Width           =   11775
      Begin VB.Label lblCustomer 
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
         TabIndex        =   6
         Top             =   240
         Width           =   840
      End
      Begin VB.Label txtCustomer 
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
         Left            =   1440
         TabIndex        =   7
         Top             =   240
         Width           =   7215
      End
   End
   Begin VB.Frame fraHeader 
      Height          =   615
      Left            =   4320
      TabIndex        =   16
      Top             =   1800
      Width           =   7575
      Begin VB.Label txtNoSeq 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
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
         Height          =   270
         Left            =   1320
         TabIndex        =   9
         Top             =   240
         Width           =   495
      End
      Begin VB.Label lblNoSeq 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nomor"
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
         Width           =   570
      End
   End
   Begin MSDataGridLib.DataGrid dgdMain 
      Height          =   3855
      Left            =   120
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1920
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   6800
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
            LCID            =   1057
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
            LCID            =   1057
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
   Begin MSComctlLib.Toolbar tlbMain 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Width           =   12030
      _ExtentX        =   21220
      _ExtentY        =   1111
      ButtonWidth     =   609
      ButtonHeight    =   953
      Appearance      =   1
      _Version        =   393216
      Begin MSComctlLib.ImageList imlMain 
         Left            =   11400
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   48
         ImageHeight     =   48
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   5
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTMDELIVERYCUSTOMER.frx":1F8A
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTMDELIVERYCUSTOMER.frx":43DC
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTMDELIVERYCUSTOMER.frx":682E
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTMDELIVERYCUSTOMER.frx":8C80
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTMDELIVERYCUSTOMER.frx":B0D2
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "frmTMDELIVERYCUSTOMER"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum ButtonMode
    [AddButton] = 1
    [UpdateButton]
    [DeleteButton]
    [SaveButton]
    [CancelButton]
End Enum

Private rstMain As ADODB.Recordset

Private objMode As FunctionMode

Private strFormCaption As String

Private Sub Form_Load()
    SetInitialization
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If frmTMCUSTOMER.Parent Then
        frmTMCUSTOMER.Parent = False
    End If
    
    If frmTHSJSELL.Parent Then
        frmTHSJSELL.Parent = False
    End If
    
    mdlDatabase.CloseRecordset rstMain
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmTMDELIVERYCUSTOMER = Nothing
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

Private Sub txtAddress_GotFocus()
    mdlProcedures.GotFocus Me.txtAddress
End Sub

Private Sub txtPhone_GotFocus()
    mdlProcedures.GotFocus Me.txtPhone
End Sub

Private Sub txtFax_GotFocus()
    mdlProcedures.GotFocus Me.txtFax
End Sub

Private Sub txtNotes_GotFocus()
    mdlProcedures.GotFocus Me.txtNotes
End Sub

Private Sub txtName_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtPhone_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtFax_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
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
    mdlProcedures.CenterWindows Me, False
    
    With Me.tlbMain
        .AllowCustomize = False
        
        .ImageList = Me.imlMain
        
        .Buttons.Add AddButton, , "Tambah", , AddButton
        .Buttons.Add UpdateButton, , "Ubah", , UpdateButton
        .Buttons.Add DeleteButton, , "Hapus", , DeleteButton
        .Buttons.Add SaveButton, , "Simpan", , SaveButton
        .Buttons.Add CancelButton, , "Batal", , CancelButton
    End With
    
    Dim strCustomerId As String
    
    If frmTMCUSTOMER.Parent Then
        strCustomerId = frmTMCUSTOMER.CustomerId
    ElseIf frmTHSJSELL.Parent Then
        strCustomerId = frmTHSJSELL.CustomerIdCombo
    End If
    
    Me.txtCustomer.Caption = strCustomerId & " | " & mdlDatabase.GetFieldData(mdlGlobal.conInventory, "Name", mdlTable.CreateTMCUSTOMER, "CustomerId='" & strCustomerId & "'")
    
    Set rstMain = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "*", mdlTable.CreateTMDELIVERYCUSTOMER, , "CustomerId='" & strCustomerId & "'", "NoSeq ASC")
    
    ArrangeGrid
    
    strFormCaption = mdlText.strTMDELIVERYCUSTOMER
    
    objMode = ViewMode
    
    SetMode
    
    FillText
End Sub

Private Sub ArrangeGrid()
    Set Me.dgdMain.DataSource = rstMain
    
    With Me.dgdMain
        .Columns(3).Width = 3400
        .Columns(3).Locked = True
        .Columns(3).Caption = "Nama"
        
        Dim intCounter As Integer
        
        For intCounter = 0 To 2
            .Columns(intCounter).Width = 0
            .Columns(intCounter).Visible = False
            .Columns(intCounter).Locked = True
        Next intCounter
        
        For intCounter = 4 To .Columns.Count - 1
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
        .Buttons(SaveButton).Visible = blnBack
        .Buttons(CancelButton).Visible = blnBack
    End With
    
    Me.fraDetail.Enabled = blnBack
    
    If objMode = AddMode Then
        mdlProcedures.SetControlMode Me, objMode
        
        Me.txtNoSeq.Caption = ""
        
        SequentialId
        
        Me.txtName.SetFocus
    ElseIf objMode = UpdateMode Then
        mdlProcedures.SetControlMode Me, objMode, False
        
        Me.txtName.SetFocus
    Else
        Me.fraHeader.Enabled = False
    End If
    
    If rstMain.RecordCount > 0 Then
        With Me.tlbMain
            .Buttons(UpdateButton).Enabled = True
            .Buttons(DeleteButton).Enabled = True
        End With
        
        If objMode = ViewMode Then
            mdlProcedures.SetControlMode Me, objMode, False
        End If
    Else
        With Me.tlbMain
            .Buttons(UpdateButton).Enabled = False
            .Buttons(DeleteButton).Enabled = False
        End With
        
        If objMode = ViewMode Then
            mdlProcedures.SetControlMode Me, objMode
            
            Me.txtNoSeq.Caption = ""
        End If
    End If
    
    Me.Caption = mdlProcedures.SetCaptionMode(strFormCaption, objMode)
End Sub

Private Sub DeleteFunction()
    objMode = DeleteMode
    
    If mdlProcedures.SetMsgYesNo("Apakah Anda Yakin ?", mdlProcedures.SetCaptionMode(strFormCaption, objMode)) Then
        Dim strMessage As String
        
        strMessage = "Tidak Dapat Dihapus" & vbCrLf & "Terdapat Data Yang Masih Terkait"
        
        If mdlDatabase.IsDataExists(mdlGlobal.conInventory, mdlTable.CreateTHSJSELL, "DeliveryId='" & rstMain!DeliveryId & "'") Then
            MsgBox strMessage & mdlText.strTHSJSELL & ")", vbOKOnly + vbCritical, mdlProcedures.SetCaptionMode(strFormCaption, objMode)
        Else
            mdlDatabase.DeleteSingleRecord rstMain
            
            If frmTHSJSELL.Parent Then
                frmTHSJSELL.FillComboTMDELIVERYCUSTOMER
            End If
        End If
    End If
    
    objMode = ViewMode
    
    SetMode
    
    FillText
End Sub

Private Sub SaveFunction()
    If Not CheckValidation Then Exit Sub
    
    With rstMain
        Dim strCustomerId As String
        
        If frmTMCUSTOMER.Parent Then
            strCustomerId = frmTMCUSTOMER.CustomerId
        ElseIf frmTHSJSELL.Parent Then
            strCustomerId = frmTHSJSELL.CustomerIdCombo
        End If
        
        mdlDatabase.SearchRecordset rstMain, "DeliveryId", strCustomerId & Me.txtNoSeq.Caption
        
        If .EOF Then
            .AddNew
            
            !DeliveryId = strCustomerId & Me.txtNoSeq.Caption
            !CustomerId = strCustomerId
            !NoSeq = Me.txtNoSeq.Caption
            
            !CreateId = Trim(mdlGlobal.UserAuthority.UserId)
            !CreateDate = mdlProcedures.FormatDate(Now)
        End If
        
        !Name = mdlProcedures.RepDupText(Trim(Me.txtName.Text))
        !Address = mdlProcedures.RepDupText(Trim(Me.txtAddress.Text))
        !Phone = mdlProcedures.RepDupText(Trim(Me.txtPhone.Text))
        !Fax = mdlProcedures.RepDupText(Trim(Me.txtFax.Text))
        !Notes = mdlProcedures.RepDupText(Trim(Me.txtNotes.Text))
        
        !UpdateId = Trim(mdlGlobal.UserAuthority.UserId)
        !UpdateDate = mdlProcedures.FormatDate(Now)
        
        .Update
    End With
    
    If objMode = AddMode Then
        mdlProcedures.SetControlMode Me, objMode
        
        Me.txtNoSeq.Caption = ""
        
        SequentialId
        
        Me.txtName.SetFocus
    ElseIf objMode = UpdateMode Then
        objMode = ViewMode
        
        SetMode
    End If
    
    If frmTHSJSELL.Parent Then
        frmTHSJSELL.FillComboTMDELIVERYCUSTOMER
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
        Dim strCustomerId As String
        
        If frmTMCUSTOMER.Parent Then
            strCustomerId = frmTMCUSTOMER.CustomerId
        ElseIf frmTHSJSELL.Parent Then
            strCustomerId = frmTHSJSELL.CustomerIdCombo
        End If
        
        With rstMain
            mdlDatabase.SearchRecordset rstMain, "DeliveryId", strCustomerId & Me.txtNoSeq.Caption
            
            If Not .EOF Then
                MsgBox "Alamat Kirim Sudah Ada", vbOKOnly + vbExclamation, mdlProcedures.SetCaptionMode(strFormCaption, objMode)
                
                CheckValidation = False
                
                Exit Function
            End If
        End With
    End If
    
    CheckValidation = True
End Function

Private Sub SequentialId()
    Dim strCustomerId As String
    
    If frmTMCUSTOMER.Parent Then
        strCustomerId = frmTMCUSTOMER.CustomerId
    ElseIf frmTHSJSELL.Parent Then
        strCustomerId = frmTHSJSELL.CustomerIdCombo
    End If
    
    Dim rstTemp As ADODB.Recordset
    
    Set rstTemp = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "NoSeq", mdlTable.CreateTMDELIVERYCUSTOMER, False, "CustomerId='" & strCustomerId & "'", "NoSeq")
    
    With rstTemp
        If .RecordCount > 0 Then
            Dim intCounter As Integer
            
            Dim rstCheck As ADODB.Recordset
            
            Dim intHole As Integer
            
            intHole = 0
            
            For intCounter = 1 To .RecordCount
                Set rstCheck = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "NoSeq", mdlTable.CreateTMDELIVERYCUSTOMER, False, "DeliveryId='" & strCustomerId & mdlProcedures.FormatNumber(intCounter, "000") & "'")
                
                If Not rstCheck.RecordCount > 0 Then
                    intHole = intCounter
                    
                    Exit For
                End If
            Next intCounter
            
            mdlDatabase.CloseRecordset rstCheck
            
            If intHole = 0 Then intHole = .RecordCount + 1
            
            Me.txtNoSeq.Caption = mdlProcedures.FormatNumber(intHole, "000")
        Else
            Me.txtNoSeq = mdlProcedures.FormatNumber(1, "000")
        End If
    End With
    
    mdlDatabase.CloseRecordset rstTemp
End Sub

Private Sub FillText()
    With rstMain
        If .RecordCount > 0 Then
            Me.txtNoSeq.Caption = Trim(!NoSeq)
            Me.txtName.Text = Trim(!Name)
            Me.txtAddress.Text = Trim(!Address)
            Me.txtPhone.Text = Trim(!Phone)
            Me.txtFax.Text = Trim(!Fax)
            Me.txtNotes.Text = Trim(!Notes)
        End If
    End With
End Sub
