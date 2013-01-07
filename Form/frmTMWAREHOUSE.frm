VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTMWAREHOUSE 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4950
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   9990
   Icon            =   "frmTMWAREHOUSE.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4950
   ScaleWidth      =   9990
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraDetail 
      Height          =   3015
      Left            =   120
      TabIndex        =   13
      Top             =   1800
      Width           =   9735
      Begin VB.CheckBox chkWarehouseSet 
         Caption         =   "Gudang Utama"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   8040
         TabIndex        =   5
         Top             =   600
         Width           =   1455
      End
      Begin VB.ComboBox cmbEmployeeId 
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
         TabIndex        =   3
         Top             =   1560
         Width           =   7575
      End
      Begin VB.CommandButton cmdEmployeeId 
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
         Left            =   9240
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   1560
         Width           =   375
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
         Left            =   1560
         MaxLength       =   150
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   2040
         Width           =   6375
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
         Left            =   1560
         MaxLength       =   50
         TabIndex        =   1
         Top             =   240
         Width           =   6375
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
         Left            =   1560
         MaxLength       =   150
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   600
         Width           =   6375
      End
      Begin VB.Label lblEmployeeId 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kepala Gudang"
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
         Top             =   1560
         Width           =   1335
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
         TabIndex        =   11
         Top             =   2040
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
         TabIndex        =   8
         Top             =   240
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
         TabIndex        =   9
         Top             =   600
         Width           =   615
      End
   End
   Begin VB.Frame fraHeader 
      Height          =   615
      Left            =   120
      TabIndex        =   12
      Top             =   1200
      Width           =   9735
      Begin VB.TextBox txtWarehouseId 
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
         Left            =   1560
         MaxLength       =   5
         TabIndex        =   0
         Top             =   240
         Width           =   735
      End
      Begin VB.Label lblWarehouseId 
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
   End
   Begin MSComctlLib.Toolbar tlbMain 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   9990
      _ExtentX        =   17621
      _ExtentY        =   1164
      ButtonWidth     =   609
      ButtonHeight    =   1005
      Appearance      =   1
      _Version        =   393216
      Begin MSComctlLib.ImageList imlMain 
         Left            =   9360
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
               Picture         =   "frmTMWAREHOUSE.frx":1F8A
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTMWAREHOUSE.frx":43DC
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTMWAREHOUSE.frx":682E
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTMWAREHOUSE.frx":8C80
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTMWAREHOUSE.frx":B0D2
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTMWAREHOUSE.frx":B674
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTMWAREHOUSE.frx":DAC6
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "frmTMWAREHOUSE"
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

Private PrintMaster As clsPRTTMWAREHOUSE

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
        
        mdlProcedures.ShowForm frmBRWTMWAREHOUSE, False, True
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
        If Not PrintMaster Is Nothing Then
            Set PrintMaster = Nothing
        End If
        
        If frmTHITEMIN.Parent Then
            frmTHITEMIN.Parent = False
        End If
        
        If frmTHITEMOUT.Parent Then
            frmTHITEMOUT.Parent = False
        End If
        
        If frmTHMUTITEM.Parent Then
            frmTHMUTITEM.Parent = False
        End If
        
        If frmTHDOBUY.Parent Then
            frmTHDOBUY.Parent = False
        End If
        
        If frmTHRTRBUY.Parent Then
            frmTHRTRBUY.Parent = False
        End If
        
        mdlDatabase.CloseRecordset rstMain
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmTMWAREHOUSE = Nothing
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
            
            mdlProcedures.ShowForm frmBRWTMWAREHOUSE, False, True
        Case SaveButton:
            SaveFunction
        Case CancelButton:
            objMode = ViewMode
            
            SetMode
            
            FillText
    End Select
End Sub

Private Sub txtWarehouseId_GotFocus()
    mdlProcedures.GotFocus Me.txtWarehouseId
End Sub

Private Sub txtName_GotFocus()
    mdlProcedures.GotFocus Me.txtName
End Sub

Private Sub txtAddress_GotFocus()
    mdlProcedures.GotFocus Me.txtAddress
End Sub

Private Sub txtNotes_GotFocus()
    mdlProcedures.GotFocus Me.txtNotes
End Sub

Private Sub txtWarehouseId_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtName_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub cmbEmployeeId_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If mdlProcedures.IsValidComboData(Me.cmbEmployeeId) Then
            SendKeys "{TAB}"
        Else
            If blnParent Then Exit Sub
            
            blnParent = True
            
            mdlProcedures.CornerWindows Me
            
            mdlProcedures.ShowForm frmBRWTMEMPLOYEE, False, True
        End If
    End If
End Sub

Private Sub cmdEmployeeId_Click()
    If Not mdlGlobal.UserAuthority.IsMenuAccess(mdiMain.mnuTMEMPLOYEE.Name) Then Exit Sub
    
    If blnParent Then Exit Sub
    
    blnParent = True
    
    mdlProcedures.CornerWindows Me
    
    mdlProcedures.ShowForm frmTMEMPLOYEE, False
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
    
    strFormCaption = mdlText.strTMWAREHOUSE
    
    blnParent = False
    
    Set rstMain = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "*", mdlTable.CreateTMWAREHOUSE, , , "WarehouseId ASC")
    
    objMode = ViewMode
    
    SetMode
    
    FillText
    
    blnParent = False
    blnActivate = False
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
        
        Me.txtWarehouseId.SetFocus
    ElseIf objMode = UpdateMode Then
        Me.fraHeader.Enabled = False
        
        mdlProcedures.SetControlMode Me, objMode, False, Me.txtWarehouseId.Name
        
        Me.txtName.SetFocus
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
            mdlProcedures.SetControlMode Me, objMode
        End If
    End If
    
    Me.Caption = mdlProcedures.SetCaptionMode(strFormCaption, objMode)
End Sub

Private Sub PrintFunction()
    If Not PrintMaster Is Nothing Then
        Set PrintMaster = Nothing
    End If

    Set PrintMaster = New clsPRTTMWAREHOUSE
    
    Dim rstTemp As ADODB.Recordset
    
    Set rstTemp = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "*", mdlTable.CreateTMWAREHOUSE, False, , "WarehouseSet DESC, WarehouseId ASC")
    
    PrintMaster.ImportToExcel rstTemp
    
    mdlDatabase.CloseRecordset rstTemp
End Sub

Private Sub DeleteFunction()
    If Trim(rstMain!WarehouseSet) = mdlGlobal.strYes Then
        MsgBox "Data Gudang Utama Tidak Dapat Dihapus", vbOKOnly + vbCritical, mdlProcedures.SetCaptionMode(strFormCaption, objMode)
        
        Exit Sub
    End If
    
    objMode = DeleteMode
    
    If mdlProcedures.SetMsgYesNo("Apakah Anda Yakin ?", mdlProcedures.SetCaptionMode(strFormCaption, objMode)) Then
        Dim strMessage As String
        
        strMessage = "Tidak Dapat Dihapus" & vbCrLf & "Terdapat Data Yang Masih Terkait ("
        
        If mdlDatabase.IsDataExists(mdlGlobal.conInventory, mdlTable.CreateTMSTOCKINIT, "WarehouseId='" & rstMain!WarehouseId & "'") Then
            MsgBox strMessage & mdlText.strTMSTOCKINIT & ")", vbOKOnly + vbCritical, mdlProcedures.SetCaptionMode(strFormCaption, objMode)
        ElseIf mdlDatabase.IsDataExists(mdlGlobal.conInventory, mdlTable.CreateTHSTOCK, "WarehouseId='" & rstMain!WarehouseId & "'") Then
            MsgBox strMessage & mdlTable.CreateTHSTOCK & ")", vbOKOnly + vbCritical, mdlProcedures.SetCaptionMode(strFormCaption, objMode)
        ElseIf mdlDatabase.IsDataExists(mdlGlobal.conInventory, mdlTable.CreateTHITEMIN, "WarehouseId='" & rstMain!WarehouseId & "'") Then
            MsgBox strMessage & mdlText.strTHITEMIN & ")", vbOKOnly + vbCritical, mdlProcedures.SetCaptionMode(strFormCaption, objMode)
        ElseIf mdlDatabase.IsDataExists(mdlGlobal.conInventory, mdlTable.CreateTHITEMOUT, "WarehouseId='" & rstMain!WarehouseId & "'") Then
            MsgBox strMessage & mdlText.strTHITEMOUT & ")", vbOKOnly + vbCritical, mdlProcedures.SetCaptionMode(strFormCaption, objMode)
        ElseIf mdlDatabase.IsDataExists(mdlGlobal.conInventory, mdlTable.CreateTHMUTITEM, "WarehouseFrom='" & rstMain!WarehouseId & "' OR WarehouseTo='" & rstMain!WarehouseId & "'") Then
            MsgBox strMessage & mdlText.strTHMUTITEM & ")", vbOKOnly + vbCritical, mdlProcedures.SetCaptionMode(strFormCaption, objMode)
        ElseIf mdlDatabase.IsDataExists(mdlGlobal.conInventory, mdlTable.CreateTHDOBUY, "WarehouseId='" & rstMain!WarehouseId & "'") Then
            MsgBox strMessage & mdlText.strTHDOBUY & ")", vbOKOnly + vbCritical, mdlProcedures.SetCaptionMode(strFormCaption, objMode)
        ElseIf mdlDatabase.IsDataExists(mdlGlobal.conInventory, mdlTable.CreateTDSJSELL, "WarehouseId='" & rstMain!WarehouseId & "'") Then
            MsgBox strMessage & mdlText.strTDSJSELL & ")", vbOKOnly + vbCritical, mdlProcedures.SetCaptionMode(strFormCaption, objMode)
        ElseIf mdlDatabase.IsDataExists(mdlGlobal.conInventory, mdlTable.CreateTHRTRSELL, "WarehouseId='" & rstMain!WarehouseId & "'") Then
            MsgBox strMessage & mdlText.strTHRTRSELL & ")", vbOKOnly + vbCritical, mdlProcedures.SetCaptionMode(strFormCaption, objMode)
        Else
            mdlDatabase.DeleteSingleRecord rstMain
            
            If frmTHITEMIN.Parent Then
                frmTHITEMIN.FillComboTMWAREHOUSE
            End If
            
            If frmTHITEMOUT.Parent Then
                frmTHITEMOUT.FillComboTMWAREHOUSE
            End If
            
            If frmTHMUTITEM.Parent Then
                frmTHMUTITEM.FillComboTMWAREHOUSE
            End If
            
            If frmTHDOBUY.Parent Then
                frmTHDOBUY.FillComboTMWAREHOUSE
            End If
            
            If frmTHRTRSELL.Parent Then
                frmTHRTRSELL.FillComboTMWAREHOUSE
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
        mdlDatabase.SearchRecordset rstMain, "WarehouseId", mdlProcedures.RepDupText(Trim(Me.txtWarehouseId.Text))
        
        If .EOF Then
            .AddNew
            
            !WarehouseId = mdlProcedures.RepDupText(Trim(Me.txtWarehouseId.Text))
            
            !CreateId = Trim(mdlGlobal.UserAuthority.UserId)
            !CreateDate = mdlProcedures.FormatDate(Now)
        End If
        
        !Name = mdlProcedures.RepDupText(Trim(Me.txtName.Text))
        !Address = mdlProcedures.RepDupText(Trim(Me.txtAddress.Text))
        !EmployeeId = Me.EmployeeIdCombo
        !Notes = mdlProcedures.RepDupText(Trim(Me.txtNotes.Text))
        
        If Me.chkWarehouseSet.Value Then
            !WarehouseSet = mdlGlobal.strYes
        Else
            !WarehouseSet = mdlGlobal.strNo
        End If
        
        !UpdateId = Trim(mdlGlobal.UserAuthority.UserId)
        !UpdateDate = mdlProcedures.FormatDate(Now)
        
        .Update
    End With
    
    If objMode = AddMode Then
        mdlProcedures.SetControlMode Me, objMode
        
        Me.txtWarehouseId.SetFocus
    ElseIf objMode = UpdateMode Then
        objMode = ViewMode
        
        SetMode
    End If
    
    If frmTHITEMIN.Parent Then
        frmTHITEMIN.FillComboTMWAREHOUSE
    End If
    
    If frmTHITEMOUT.Parent Then
        frmTHITEMOUT.FillComboTMWAREHOUSE
    End If
    
    If frmTHMUTITEM.Parent Then
        frmTHMUTITEM.FillComboTMWAREHOUSE
    End If
    
    If frmTHDOBUY.Parent Then
        frmTHDOBUY.FillComboTMWAREHOUSE
    End If
    
    If frmTHRTRSELL.Parent Then
        frmTHRTRSELL.FillComboTMWAREHOUSE
    End If
End Sub

Private Function CheckValidation() As Boolean
    If mdlProcedures.RepDupText(Trim(Me.txtWarehouseId.Text)) = "" Then
        MsgBox "Kode Harap Diisi", vbOKOnly + vbExclamation, mdlProcedures.SetCaptionMode(strFormCaption, objMode)
        
        Me.txtWarehouseId.SetFocus
        
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
            mdlDatabase.SearchRecordset rstMain, "WarehouseId", mdlProcedures.RepDupText(Trim(Me.txtWarehouseId.Text))
            
            If Not .EOF Then
                MsgBox "Kode Sudah Ada", vbOKOnly + vbExclamation, mdlProcedures.SetCaptionMode(strFormCaption, objMode)
                
                Me.txtWarehouseId.SetFocus
            
                CheckValidation = False
                
                Exit Function
            End If
        End With
    End If
    
    Dim rstTemp As ADODB.Recordset
    
    Set rstTemp = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "*", mdlTable.CreateTMWAREHOUSE, False, "WarehouseSet='" & mdlGlobal.strYes & "'")
    
    If rstTemp.RecordCount > 0 Then
        If objMode = AddMode Then
            If Me.chkWarehouseSet.Value = vbChecked Then
                If mdlProcedures.SetMsgYesNo("Gudang Utama Sudah Ada" & vbCrLf & "Ganti Gudang Utama ?", mdlProcedures.SetCaptionMode(strFormCaption, objMode)) Then
                    Me.chkWarehouseSet.Value = vbChecked
                    
                    WarehouseSet mdlGlobal.strNo
                Else
                    Me.chkWarehouseSet.SetFocus
                    
                    CheckValidation = False
                    
                    Exit Function
                End If
            End If
        ElseIf objMode = UpdateMode Then
            If Me.chkWarehouseSet.Value Then
                If Not UCase(Trim(rstTemp!WarehouseId)) = UCase(mdlProcedures.RepDupText(Trim(Me.txtWarehouseId.Text))) Then
                    If mdlProcedures.SetMsgYesNo("Gudang Utama Sudah Ada" & vbCrLf & "Ganti Gudang Utama ?", mdlProcedures.SetCaptionMode(strFormCaption, objMode)) Then
                        Me.chkWarehouseSet.Value = vbChecked
                        
                        WarehouseSet mdlGlobal.strNo
                    Else
                        Me.chkWarehouseSet.SetFocus
                        
                        CheckValidation = False
                        
                        Exit Function
                    End If
                End If
            Else
                If UCase(Trim(rstTemp!WarehouseId)) = UCase(mdlProcedures.RepDupText(Trim(Me.txtWarehouseId.Text))) Then
                    MsgBox "Gudang Utama Tidak Dapat Diubah", vbCritical, mdlProcedures.SetCaptionMode(strFormCaption, objMode)
                    
                    Me.chkWarehouseSet.SetFocus
                    
                    CheckValidation = False
                    
                    Exit Function
                End If
                    
            End If
        End If
    Else
        If objMode = AddMode Or objMode = UpdateMode Then
            If Not Me.chkWarehouseSet.Value = vbChecked Then
                If mdlProcedures.SetMsgYesNo("Gudang Utama Belum Terdaftar" & vbCrLf & "Setting Gudang Utama ?", mdlProcedures.SetCaptionMode(strFormCaption, objMode)) Then
                    Me.chkWarehouseSet.Value = vbChecked
                Else
                    Me.chkWarehouseSet.SetFocus
                    
                    CheckValidation = False
                    
                    Exit Function
                End If
            End If
        End If
    End If
    
    mdlDatabase.CloseRecordset rstTemp
    
    CheckValidation = True
End Function

Private Sub WarehouseSet(ByVal strValue As String, Optional ByVal strWarehouseId As String)
    With rstMain
        If .RecordCount > 0 Then
            .MoveFirst
        End If
        
        While Not .EOF
            !WarehouseSet = strValue
            
            .Update
            
            .MoveNext
        Wend
    End With
End Sub

Private Sub FillText()
    With rstMain
        If .RecordCount > 0 Then
            Me.txtWarehouseId.Text = Trim(!WarehouseId)
            Me.txtName.Text = Trim(!Name)
            Me.txtAddress.Text = Trim(!Address)
            
            Me.EmployeeIdCombo = !EmployeeId
            
            Me.txtNotes.Text = Trim(!Notes)
            
            If Trim(!WarehouseSet) = mdlGlobal.strYes Then
                Me.chkWarehouseSet.Value = vbChecked
            Else
                Me.chkWarehouseSet.Value = vbUnchecked
            End If
        End If
    End With
End Sub

Private Sub FillCombo()
    Me.FillComboTMEMPLOYEE
End Sub

Public Sub FillComboTMEMPLOYEE()
    Dim rstTemp As ADODB.Recordset
    
    Set rstTemp = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "EmployeeId, Name", mdlTable.CreateTMEMPLOYEE, False)
    
    mdlProcedures.FillComboData Me.cmbEmployeeId, rstTemp
    
    mdlDatabase.CloseRecordset rstTemp
End Sub

Public Property Get Parent() As Boolean
    Parent = blnParent
End Property

Public Property Get WarehouseId() As String
    If rstMain Is Nothing Then Exit Property

    If rstMain.RecordCount > 0 Then
        WarehouseId = rstMain!WarehouseId
    End If
End Property

Public Property Get WarehouseName() As String
    If rstMain Is Nothing Then Exit Property

    If rstMain.RecordCount > 0 Then
        WarehouseName = rstMain!Name
    End If
End Property

Public Property Get EmployeeIdCombo() As String
    EmployeeIdCombo = mdlProcedures.GetComboData(Me.cmbEmployeeId)
End Property

Public Property Let Parent(ByVal blnEnable As Boolean)
    blnParent = blnEnable
    
    If Not blnEnable Then mdlProcedures.CenterWindows Me
End Property

Public Property Let WarehouseId(ByVal strWarehouseId As String)
    If rstMain Is Nothing Then Exit Property
    
    mdlDatabase.SearchRecordset rstMain, "WarehouseId", strWarehouseId
    
    If Not rstMain.EOF Then FillText
End Property

Public Property Let EmployeeIdCombo(ByVal strEmployeeId As String)
    mdlProcedures.SetComboData Me.cmbEmployeeId, strEmployeeId
End Property
