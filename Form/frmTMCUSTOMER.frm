VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmTMCUSTOMER 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6045
   ClientLeft      =   45
   ClientTop       =   420
   ClientWidth     =   10335
   Icon            =   "frmTMCUSTOMER.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6045
   ScaleWidth      =   10335
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraDetail 
      Height          =   4095
      Left            =   120
      TabIndex        =   18
      Top             =   1800
      Width           =   10095
      Begin VB.CheckBox chkStatusYN 
         Caption         =   "Status Aktif"
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
         Left            =   6000
         TabIndex        =   2
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox txtNPWP 
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
         MaxLength       =   20
         TabIndex        =   7
         Top             =   2760
         Width           =   2535
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
         TabIndex        =   5
         Top             =   2040
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
         TabIndex        =   6
         Top             =   2400
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
         TabIndex        =   8
         Top             =   3120
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
         TabIndex        =   3
         Top             =   720
         Width           =   6135
      End
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
         TabIndex        =   4
         Top             =   1080
         Width           =   6135
      End
      Begin MSComCtl2.DTPicker dtpCustomerDate 
         Height          =   375
         Left            =   1320
         TabIndex        =   1
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
         Format          =   137625603
         CurrentDate     =   39383
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
         TabIndex        =   13
         Top             =   2040
         Width           =   675
      End
      Begin VB.Label lblNPWP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NPWP"
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
         TabIndex        =   15
         Top             =   2760
         Width           =   600
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
         TabIndex        =   14
         Top             =   2400
         Width           =   330
      End
      Begin VB.Label lblCustomerDate 
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
         TabIndex        =   10
         Top             =   240
         Width           =   675
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
         TabIndex        =   16
         Top             =   3120
         Width           =   990
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
         TabIndex        =   11
         Top             =   720
         Width           =   510
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
         TabIndex        =   12
         Top             =   1080
         Width           =   615
      End
   End
   Begin VB.Frame fraHeader 
      Height          =   615
      Left            =   120
      TabIndex        =   17
      Top             =   1200
      Width           =   10095
      Begin VB.TextBox txtCustomerId 
         BackColor       =   &H8000000F&
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
         MaxLength       =   6
         TabIndex        =   0
         Top             =   240
         Width           =   855
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
         TabIndex        =   9
         Top             =   240
         Width           =   450
      End
   End
   Begin MSComctlLib.Toolbar tlbMain 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   1111
      ButtonWidth     =   609
      ButtonHeight    =   953
      Appearance      =   1
      _Version        =   393216
      Begin MSComctlLib.ImageList imlMain 
         Left            =   9720
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   48
         ImageHeight     =   48
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   11
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTMCUSTOMER.frx":1F8A
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTMCUSTOMER.frx":43DC
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTMCUSTOMER.frx":682E
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTMCUSTOMER.frx":8C80
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTMCUSTOMER.frx":B0D2
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTMCUSTOMER.frx":B674
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTMCUSTOMER.frx":D1C6
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTMCUSTOMER.frx":F618
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTMCUSTOMER.frx":101EA
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTMCUSTOMER.frx":1263E
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTMCUSTOMER.frx":14A90
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "frmTMCUSTOMER"
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
    [ContactButton]
    [ReminderButton]
    [DeliveryButton]
    [NotesButton]
    [SaveButton]
    [CancelButton]
End Enum

Private rstMain As ADODB.Recordset

Private PrintMaster As clsPRTTMCUSTOMER

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
        
        mdlProcedures.ShowForm frmBRWTMCUSTOMER, False, True
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
        mdlDatabase.CloseRecordset rstMain
        
        If frmTHPOSELL.Parent Then
            frmTHPOSELL.Parent = False
        End If
        
        If frmTHSALESSUM.Parent Then
            frmTHSALESSUM.Parent = False
        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmTMCUSTOMER = Nothing
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
            
            mdlProcedures.ShowForm frmBRWTMCUSTOMER, False, True
        Case ContactButton:
            If blnParent Then Exit Sub
            
            blnParent = True
            
            mdlProcedures.CornerWindows Me
            
            mdlProcedures.ShowForm frmTMCONTACTCUSTOMER, False
        Case ReminderButton:
            If blnParent Then Exit Sub
            
            blnParent = True
            
            mdlProcedures.CornerWindows Me
            
            mdlProcedures.ShowForm frmTMREMINDERCUSTOMER, False
        Case DeliveryButton:
            If blnParent Then Exit Sub
            
            blnParent = True
            
            mdlProcedures.CornerWindows Me
            
            mdlProcedures.ShowForm frmTMDELIVERYCUSTOMER, False, True
        Case NotesButton:
            If blnParent Then Exit Sub
            
            blnParent = True
            
            mdlProcedures.CornerWindows Me
            
            mdlProcedures.ShowForm frmTMCUSTOMERNOTES, False, True
        Case SaveButton:
            SaveFunction
        Case CancelButton:
            objMode = ViewMode
            
            SetMode
            
            FillText
    End Select
End Sub

Private Sub txtCustomerId_GotFocus()
    mdlProcedures.GotFocus Me.txtCustomerId
End Sub

Private Sub txtName_GotFocus()
    mdlProcedures.GotFocus Me.txtName
End Sub

Private Sub txtAddress_GotFocus()
    mdlProcedures.GotFocus Me.txtAddress
End Sub

Private Sub txtFax_GotFocus()
    mdlProcedures.GotFocus Me.txtFax
End Sub

Private Sub txtNotes_GotFocus()
    mdlProcedures.GotFocus Me.txtNotes
End Sub

Private Sub txtCustomerId_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub dtpCustomerDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
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

Private Sub txtNPWP_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
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
        .Buttons.Add ContactButton, , "Kontak", , ContactButton
        .Buttons.Add ReminderButton, , "Pengingat", , ReminderButton
        .Buttons.Add DeliveryButton, , "Alamat Kirim", , DeliveryButton
        .Buttons.Add NotesButton, , "Catatan", , NotesButton
        .Buttons.Add SaveButton, , "Simpan", , SaveButton
        .Buttons.Add CancelButton, , "Batal", , CancelButton
    End With
    
    Me.dtpCustomerDate.CustomFormat = mdlGlobal.strFormatDate
    
    Me.txtCustomerId.Locked = True
    
    strFormCaption = mdlText.strTMCUSTOMER
    
    blnParent = False
    blnActivate = False
    
    Set rstMain = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "*", mdlTable.CreateTMCUSTOMER, , , "CustomerId ASC")
    
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
        .Buttons(ContactButton).Visible = blnFront
        .Buttons(ReminderButton).Visible = blnFront
        .Buttons(DeliveryButton).Visible = blnFront
        .Buttons(NotesButton).Visible = blnFront
        .Buttons(SaveButton).Visible = blnBack
        .Buttons(CancelButton).Visible = blnBack
    End With
    
    Me.fraDetail.Enabled = blnBack
    
    If objMode = AddMode Then
        Me.fraHeader.Enabled = True
    
        mdlProcedures.SetControlMode Me, objMode
        
        IncrementId
        
        Me.chkStatusYN.Value = vbChecked
        
        Me.dtpCustomerDate.SetFocus
    ElseIf objMode = UpdateMode Then
        Me.fraHeader.Enabled = False
        
        mdlProcedures.SetControlMode Me, objMode, False, Me.txtCustomerId.Name
        
        Me.dtpCustomerDate.SetFocus
    Else
        Me.fraHeader.Enabled = False
    End If
    
    If rstMain.RecordCount > 0 Then
        With Me.tlbMain
            .Buttons(UpdateButton).Enabled = True
            .Buttons(DeleteButton).Enabled = True
            .Buttons(PrintButton).Enabled = True
            .Buttons(BrowseButton).Enabled = True
            .Buttons(ContactButton).Enabled = True
            .Buttons(ReminderButton).Enabled = True
            .Buttons(DeliveryButton).Enabled = True
            .Buttons(NotesButton).Enabled = True
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
            .Buttons(ContactButton).Enabled = False
            .Buttons(ReminderButton).Enabled = False
            .Buttons(DeliveryButton).Enabled = False
            .Buttons(NotesButton).Enabled = False
        End With
        
        If objMode = ViewMode Then
            mdlProcedures.SetControlMode Me, objMode
        End If
    End If
    
    Me.Caption = mdlProcedures.SetCaptionMode(strFormCaption, objMode)
End Sub

Private Sub PrintFunction()
    If Not PrintMaster Is Nothing Then Set PrintMaster = Nothing

    Set PrintMaster = New clsPRTTMCUSTOMER
    
    Dim rstTemp As ADODB.Recordset
    
    Set rstTemp = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "*", mdlTable.CreateTMCUSTOMER, False, , "CustomerId")
    
    PrintMaster.ImportToExcel rstTemp
    
    mdlDatabase.CloseRecordset rstTemp
End Sub

Private Sub DeleteFunction()
    objMode = DeleteMode
    
    If mdlProcedures.SetMsgYesNo("Apakah Anda Yakin ?", mdlProcedures.SetCaptionMode(strFormCaption, objMode)) Then
        Dim strMessage As String
        
        strMessage = "Tidak Dapat Dihapus" & vbCrLf & "Terdapat Data Yang Masih Terkait ("
        
        If mdlDatabase.IsDataExists(mdlGlobal.conInventory, mdlTable.CreateTHPOSELL, "CustomerId='" & rstMain!CustomerId & "'") Then
            MsgBox strMessage & mdlText.strTHPOSELL & ")", vbCritical + vbOKOnly, mdlProcedures.SetCaptionMode(strFormCaption, objMode)
        ElseIf mdlDatabase.IsDataExists(mdlGlobal.conInventory, mdlTable.CreateTHFKTSELL, "CustomerId='" & rstMain!CustomerId & "'") Then
            MsgBox strMessage & mdlText.strTHFKTSELL & ")", vbCritical + vbOKOnly, mdlProcedures.SetCaptionMode(strFormCaption, objMode)
        Else
            mdlDatabase.DeleteRecordQuery mdlGlobal.conInventory, mdlTable.CreateTMCONTACTCUSTOMER, "CustomerId='" & rstMain!CustomerId & "'"
            mdlDatabase.DeleteRecordQuery mdlGlobal.conInventory, mdlTable.CreateTMCUSTOMERNOTES, "CustomerId='" & rstMain!CustomerId & "'"
            mdlDatabase.DeleteRecordQuery mdlGlobal.conInventory, mdlTable.CreateTMREMINDERCUSTOMER, "CustomerId='" & rstMain!CustomerId & "'"
            
            mdlDatabase.DeleteSingleRecord rstMain
            
            If frmTHPOSELL.Parent Then
                frmTHPOSELL.FillComboTMCUSTOMER
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
        mdlDatabase.SearchRecordset rstMain, "CustomerId", mdlProcedures.RepDupText(Trim(Me.txtCustomerId.Text))
        
        If .EOF Then
            .AddNew
            
            !CustomerId = mdlProcedures.RepDupText(Trim(Me.txtCustomerId.Text))
            
            !CreateId = Trim(mdlGlobal.UserAuthority.UserId)
            !CreateDate = mdlProcedures.FormatDate(Now)
            
            SaveReminder !CustomerId
        End If
        
        !CustomerDate = mdlProcedures.FormatDate(Me.dtpCustomerDate.Value)
        !Name = mdlProcedures.RepDupText(Trim(Me.txtName.Text))
        !Address = mdlProcedures.RepDupText(Trim(Me.txtAddress.Text))
        !Phone = mdlProcedures.RepDupText(Trim(Me.txtPhone.Text))
        !Fax = mdlProcedures.RepDupText(Trim(Me.txtFax.Text))
        !NPWP = mdlProcedures.RepDupText(Trim(Me.txtNPWP.Text))
        
        If Me.chkStatusYN.Value Then
            !StatusYN = mdlGlobal.strYes
        Else
            !StatusYN = mdlGlobal.strNo
        End If
        
        !Notes = mdlProcedures.RepDupText(Trim(Me.txtNotes.Text))
        
        !UpdateId = Trim(mdlGlobal.UserAuthority.UserId)
        !UpdateDate = mdlProcedures.FormatDate(Now)
        
        .Update
    End With
    
    If objMode = AddMode Then
        mdlProcedures.SetControlMode Me, objMode
        
        IncrementId
        
        Me.chkStatusYN.Value = vbChecked
        
        Me.dtpCustomerDate.SetFocus
    ElseIf objMode = UpdateMode Then
        objMode = ViewMode
        
        SetMode
    End If
    
    If frmTHPOSELL.Parent Then
        frmTHPOSELL.FillComboTMCUSTOMER
    End If
End Sub

Private Function CheckValidation() As Boolean
    If mdlProcedures.RepDupText(Trim(Me.txtCustomerId.Text)) = "" Then
        MsgBox "Kode Harap Diisi", vbOKOnly + vbExclamation, mdlProcedures.SetCaptionMode(strFormCaption, objMode)
        
        Me.txtCustomerId.SetFocus
        
        CheckValidation = False
        
        Exit Function
    ElseIf mdlProcedures.RepDupText(Trim(Me.txtName.Text)) = "" Then
        MsgBox "Nama Harap Diisi", vbOKOnly + vbExclamation, mdlProcedures.SetCaptionMode(strFormCaption, objMode)
        
        Me.txtName.SetFocus
        
        CheckValidation = False
        
        Exit Function
    End If
    
    If objMode = AddMode Then
        With rstMain
            mdlDatabase.SearchRecordset rstMain, "CustomerId", mdlProcedures.RepDupText(Trim(Me.txtCustomerId.Text))
            
            If Not .EOF Then
                MsgBox "Kode Sudah Ada", vbOKOnly + vbExclamation, mdlProcedures.SetCaptionMode(strFormCaption, objMode)
                
                Me.txtCustomerId.SetFocus
                
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
            Me.txtCustomerId.Text = Trim(!CustomerId)
            Me.dtpCustomerDate.Value = mdlProcedures.FormatDate(!CustomerDate, mdlGlobal.strFormatDate)
            Me.txtName.Text = Trim(!Name)
            Me.txtAddress.Text = Trim(!Address)
            Me.txtPhone.Text = Trim(!Phone)
            Me.txtFax.Text = Trim(!Fax)
            Me.txtNPWP.Text = Trim(!NPWP)
            
            If Trim(!StatusYN) = mdlGlobal.strYes Then
                Me.chkStatusYN.Value = vbChecked
            Else
                Me.chkStatusYN.Value = vbUnchecked
            End If
            
            Me.txtNotes.Text = Trim(!Notes)
        End If
    End With
End Sub

Private Sub IncrementId()
    Dim rstTemp As ADODB.Recordset
    
    Set rstTemp = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "CustomerId", mdlTable.CreateTMCUSTOMER, False, , "CustomerId")
    
    With rstTemp
        If .RecordCount > 0 Then
            Dim intCounter As Integer
            
            Dim rstCheck As ADODB.Recordset
            
            Dim intHole As Integer
            
            intHole = 0
            
            For intCounter = 1 To .RecordCount
                Set rstCheck = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "CustomerId", mdlTable.CreateTMCUSTOMER, False, "CustomerId='" & mdlText.strCUSTOMERIDINIT & mdlProcedures.FormatNumber(intCounter, "00000") & "'")
                
                If Not rstCheck.RecordCount > 0 Then
                    intHole = intCounter
                    
                    Exit For
                End If
            Next intCounter
            
            mdlDatabase.CloseRecordset rstCheck
            
            If intHole = 0 Then intHole = .RecordCount + 1
            
            Me.txtCustomerId.Text = mdlText.strCUSTOMERIDINIT & mdlProcedures.FormatNumber(intHole, "00000")
        Else
            Me.txtCustomerId.Text = mdlText.strCUSTOMERIDINIT & mdlProcedures.FormatNumber(1, "00000")
        End If
    End With
    
    mdlDatabase.CloseRecordset rstTemp
End Sub

Private Sub SaveReminder(ByVal strCustomerId As String)
    Dim rstTemp As ADODB.Recordset
    
    Set rstTemp = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "*", mdlTable.CreateTMREMINDERCUSTOMER, , "CustomerId='" & strCustomerId & "'")
    
    With rstTemp
        If Not .RecordCount > 0 Then
            .AddNew
            
            !CustomerId = strCustomerId
            
            !ReminderType = ReminderType.FromMaster
            !ReminderDate = mdlProcedures.FormatDate(Me.dtpCustomerDate)
            !ValidateType = ValidateType.OnceMonth
            
            Dim objValidateDate As Date
            
            objValidateDate = mdlProcedures.DateAddFormat(Me.dtpCustomerDate.Value, Me.dtpCustomerDate.Value, "M", 1)
            
            !ValidateDate = mdlProcedures.FormatDate(objValidateDate)
            
            !CreateId = Trim(mdlGlobal.UserAuthority.UserId)
            !CreateDate = mdlProcedures.FormatDate(Now)
            
            !UpdateId = Trim(mdlGlobal.UserAuthority.UserId)
            !UpdateDate = mdlProcedures.FormatDate(Now)
            
            .Update
        End If
    End With
End Sub

Public Property Get Parent() As Boolean
    Parent = blnParent
End Property

Public Property Get CustomerId() As String
    If rstMain Is Nothing Then Exit Property

    If rstMain.RecordCount > 0 Then
        CustomerId = rstMain!CustomerId
    End If
End Property

Public Property Get CustomerName() As String
    If rstMain Is Nothing Then Exit Property

    If rstMain.RecordCount > 0 Then
        CustomerName = rstMain!Name
    End If
End Property

Public Property Let Parent(ByVal blnEnable As Boolean)
    blnParent = blnEnable
    
    If Not blnEnable Then mdlProcedures.CenterWindows Me
End Property

Public Property Let CustomerId(ByVal strCustomerId As String)
    If rstMain Is Nothing Then Exit Property
    
    mdlDatabase.SearchRecordset rstMain, "CustomerId", strCustomerId
    
    If Not rstMain.EOF Then FillText
End Property
