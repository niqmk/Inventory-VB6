VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmTMCONTACTNOTES 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7965
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   9015
   Icon            =   "frmTMCONTACTNOTES.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7965
   ScaleWidth      =   9015
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraHeader 
      Height          =   735
      Left            =   120
      TabIndex        =   8
      Top             =   2160
      Width           =   8775
      Begin MSComCtl2.DTPicker dtpNotesDate 
         Height          =   375
         Left            =   1440
         TabIndex        =   9
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
         Format          =   91619331
         CurrentDate     =   39408
      End
      Begin MSComCtl2.DTPicker dtpNotesTime 
         Height          =   375
         Left            =   5520
         TabIndex        =   10
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
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
         CustomFormat    =   "hh:mm:ss"
         Format          =   91619331
         CurrentDate     =   39408
      End
      Begin VB.Label lblNotesDate 
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
         TabIndex        =   12
         Top             =   240
         Width           =   675
      End
      Begin VB.Label lblNotesTime 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Jam"
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
         Left            =   4920
         TabIndex        =   11
         Top             =   240
         Width           =   360
      End
   End
   Begin VB.Frame fraDetail 
      Height          =   1215
      Left            =   120
      TabIndex        =   5
      Top             =   2880
      Width           =   8775
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
         Left            =   1440
         MaxLength       =   200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Top             =   240
         Width           =   7215
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
         TabIndex        =   7
         Top             =   240
         Width           =   990
      End
   End
   Begin VB.Frame fraParent 
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   8775
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
         Left            =   1440
         TabIndex        =   4
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
         TabIndex        =   3
         Top             =   240
         Width           =   840
      End
      Begin VB.Label lblContactName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Kontak"
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
         TabIndex        =   2
         Top             =   600
         Width           =   1185
      End
      Begin VB.Label txtContactName 
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
         TabIndex        =   1
         Top             =   600
         Width           =   6135
      End
   End
   Begin MSDataGridLib.DataGrid dgdMain 
      Height          =   3615
      Left            =   120
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   4200
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   6376
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
   Begin MSComctlLib.Toolbar tlbMain 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   1164
      ButtonWidth     =   609
      ButtonHeight    =   1005
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
               Picture         =   "frmTMCONTACTNOTES.frx":1F8A
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTMCONTACTNOTES.frx":43DC
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTMCONTACTNOTES.frx":682E
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTMCONTACTNOTES.frx":8C80
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTMCONTACTNOTES.frx":B0D2
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTMCONTACTNOTES.frx":D524
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "frmTMCONTACTNOTES"
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
    [SaveButton]
    [CancelButton]
End Enum

Private rstMain As ADODB.Recordset

Private PrintMaster As clsPRTTMCONTACTNOTES

Private objMode As FunctionMode

Private strFormCaption As String

Private Sub Form_Load()
    SetInitialization
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Not PrintMaster Is Nothing Then
        Set PrintMaster = Nothing
    End If
    
    If frmTMCONTACTCUSTOMER.Parent Then
        frmTMCONTACTCUSTOMER.Parent = False
    End If
    
    If frmMISTMCUSTOMER.Parent Then
        frmMISTMCUSTOMER.Parent = False
    End If
    
    mdlDatabase.CloseRecordset rstMain
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmTMCONTACTNOTES = Nothing
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
        Case SaveButton:
            SaveFunction
        Case CancelButton:
            objMode = ViewMode
            
            SetMode
            
            FillText
    End Select
End Sub

Private Sub txtNotes_GotFocus()
    mdlProcedures.GotFocus Me.txtNotes
End Sub

Private Sub dtpNotesDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub dtpNotesTime_KeyDown(KeyCode As Integer, Shift As Integer)
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
        .Buttons.Add PrintButton, , "Cetak", , PrintButton
        .Buttons.Add SaveButton, , "Simpan", , SaveButton
        .Buttons.Add CancelButton, , "Batal", , CancelButton
    End With
    
    Me.dtpNotesDate.CustomFormat = mdlGlobal.strFormatDate
    
    Dim strCriteria As String
    
    strCriteria = ""
    
    If frmTMCONTACTCUSTOMER.Parent Then
        Me.txtCustomerId.Caption = frmTMCUSTOMER.CustomerId & " | " & frmTMCUSTOMER.CustomerName
        Me.txtContactName.Caption = frmTMCONTACTCUSTOMER.ContactName
        
        strCriteria = "ContactId='" & frmTMCONTACTCUSTOMER.ContactId & "'"
    ElseIf frmMISTMCUSTOMER.Parent Then
        Me.txtCustomerId.Caption = frmMISTMCUSTOMER.CustomerId & " | " & frmMISTMCUSTOMER.CustomerName
        Me.txtContactName.Caption = frmMISTMCUSTOMER.ContactName
        
        strCriteria = "ContactId='" & frmMISTMCUSTOMER.ContactId & "'"
    End If
    
    strFormCaption = mdlText.strTMCONTACTNOTES
    
    Set rstMain = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "*", mdlTable.CreateTMCONTACTNOTES, , strCriteria, "NotesDate ASC")
    
    ArrangeGrid
    
    objMode = ViewMode
    
    SetMode
    
    FillText
End Sub

Private Sub ArrangeGrid()
    Set Me.dgdMain.DataSource = rstMain
    
    With Me.dgdMain
        .RowHeight = 1000
        
        .Columns(2).Width = 2100
        .Columns(2).Locked = True
        .Columns(2).Caption = "Tanggal"
        .Columns(2).NumberFormat = "dd MMMM yyyy"
        .Columns(3).Width = 6000
        .Columns(3).Locked = True
        .Columns(3).Caption = "Keterangan"
        .Columns(3).WrapText = True
        
        Dim intCounter As Integer
        
        For intCounter = 0 To 1
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
        .Buttons(PrintButton).Visible = blnFront
        .Buttons(SaveButton).Visible = blnBack
        .Buttons(CancelButton).Visible = blnBack
    End With
    
    Me.fraDetail.Enabled = blnBack
    
    If objMode = AddMode Then
        Me.fraHeader.Enabled = True
    
        mdlProcedures.SetControlMode Me, objMode
        
        Me.dtpNotesDate.SetFocus
    ElseIf objMode = UpdateMode Then
        Me.fraHeader.Enabled = False
        
        mdlProcedures.SetControlMode Me, objMode, False
        
        Me.txtNotes.SetFocus
    Else
        Me.fraHeader.Enabled = False
    End If
    
    If rstMain.RecordCount > 0 Then
        With Me.tlbMain
            .Buttons(UpdateButton).Enabled = True
            .Buttons(DeleteButton).Enabled = True
            .Buttons(PrintButton).Enabled = True
        End With
        
        If objMode = ViewMode Then
            mdlProcedures.SetControlMode Me, objMode, False
        End If
    Else
        With Me.tlbMain
            .Buttons(UpdateButton).Enabled = False
            .Buttons(DeleteButton).Enabled = False
            .Buttons(PrintButton).Enabled = False
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

    Set PrintMaster = New clsPRTTMCONTACTNOTES
    
    Dim strCriteria As String
    
    strCriteria = ""
    
    If frmTMCONTACTCUSTOMER.Parent Then
        strCriteria = "ContactId='" & frmTMCONTACTCUSTOMER.ContactId & "'"
    ElseIf frmMISTMCUSTOMER.Parent Then
        strCriteria = "ContactId='" & frmMISTMCUSTOMER.ContactId & "'"
    End If
    
    Dim rstTemp As ADODB.Recordset
    
    Set rstTemp = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "*", mdlTable.CreateTMCONTACTNOTES, , strCriteria, "NotesDate ASC")
    
    PrintMaster.ImportToExcel rstTemp
    
    mdlDatabase.CloseRecordset rstTemp
End Sub

Private Sub DeleteFunction()
    objMode = DeleteMode
    
    If mdlProcedures.SetMsgYesNo("Apakah Anda Yakin ?", mdlProcedures.SetCaptionMode(strFormCaption, objMode)) Then
        mdlDatabase.DeleteSingleRecord rstMain
        
        If frmMISTMCUSTOMER.Parent Then
            frmMISTMCUSTOMER.ContactId = frmMISTMCUSTOMER.ContactId
        End If
    End If
    
    objMode = ViewMode
    
    SetMode
    
    FillText
End Sub

Private Sub SaveFunction()
    If Not CheckValidation Then Exit Sub
    
    With rstMain
        Dim strContactId As String
        
        strContactId = ""
        
        If frmTMCONTACTCUSTOMER.Parent Then
            strContactId = frmTMCONTACTCUSTOMER.ContactId
        ElseIf frmMISTMCUSTOMER.Parent Then
            strContactId = frmMISTMCUSTOMER.ContactId
        End If
        
        mdlDatabase.SearchRecordset rstMain, "NotesId", strContactId & mdlProcedures.FormatDate(Me.dtpNotesDate.Value, "ddMMyyyy") & mdlProcedures.FormatDate(Me.dtpNotesTime.Value, "hhmmss")
        
        If .EOF Then
            .AddNew
            
            !NotesId = strContactId & mdlProcedures.FormatDate(Me.dtpNotesDate.Value, "ddMMyyyy") & mdlProcedures.FormatDate(Me.dtpNotesDate.Value, "hhmmss")
            !ContactId = Trim(strContactId)
            !NotesDate = mdlProcedures.FormatDate(Me.dtpNotesDate.Value, "yyyy/MM/dd hh:mm:ss")
            
            !CreateId = Trim(mdlGlobal.UserAuthority.UserId)
            !CreateDate = mdlProcedures.FormatDate(Now)
        End If
        
        !Notes = mdlProcedures.RepDupText(Trim(Me.txtNotes.Text))
        
        !UpdateId = Trim(mdlGlobal.UserAuthority.UserId)
        !UpdateDate = mdlProcedures.FormatDate(Now)
        
        .Update
    End With
    
    If objMode = AddMode Then
        mdlProcedures.SetControlMode Me, objMode
        
        Me.dtpNotesDate.SetFocus
    ElseIf objMode = UpdateMode Then
        objMode = ViewMode
        
        SetMode
    End If
    
    If frmMISTMCUSTOMER.Parent Then
        frmMISTMCUSTOMER.ContactId = frmMISTMCUSTOMER.ContactId
    End If
End Sub

Private Function CheckValidation() As Boolean
    If mdlProcedures.RepDupText(Trim(Me.txtNotes.Text)) = "" Then
        MsgBox "Keterangan Harap Diisi", vbOKOnly + vbExclamation, mdlProcedures.SetCaptionMode(strFormCaption, objMode)
        
        Me.txtNotes.SetFocus
        
        CheckValidation = False
        
        Exit Function
    End If
    
    If objMode = AddMode Then
        With rstMain
            Dim strContactId As String
            
            strContactId = ""
            
            If frmTMCONTACTCUSTOMER.Parent Then
                strContactId = frmTMCONTACTCUSTOMER.ContactId
            ElseIf frmMISTMCUSTOMER.Parent Then
                strContactId = frmMISTMCUSTOMER.ContactId
            End If
            
            mdlDatabase.SearchRecordset rstMain, "NotesId", strContactId & mdlProcedures.FormatDate(Me.dtpNotesDate.Value, "ddMMyyyy") & mdlProcedures.FormatDate(Me.dtpNotesDate.Value, "hhmmss")
            
            If Not .EOF Then
                MsgBox "Keterangan Sudah Ada", vbOKOnly + vbExclamation, mdlProcedures.SetCaptionMode(strFormCaption, objMode)
                
                Me.dtpNotesDate.SetFocus
                
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
            Me.dtpNotesDate.Value = mdlProcedures.FormatDate(!NotesDate, mdlGlobal.strFormatDate)
            Me.dtpNotesTime.Value = mdlProcedures.FormatDate(!NotesDate, "hh:mm:ss")
            Me.txtNotes.Text = Trim(!Notes)
        End If
    End With
End Sub
