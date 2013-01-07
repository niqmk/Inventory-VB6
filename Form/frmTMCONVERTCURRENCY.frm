VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmTMCONVERTCURRENCY 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3135
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   8775
   Icon            =   "frmTMCONVERTCURRENCY.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   8775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraDetail 
      Height          =   615
      Left            =   2040
      TabIndex        =   9
      Top             =   2400
      Width           =   3015
      Begin VB.TextBox txtConvertValue 
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
         TabIndex        =   4
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lblConvertValue 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nilai Tukar"
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
         Width           =   915
      End
   End
   Begin VB.Frame fraHeader 
      Height          =   1095
      Left            =   2040
      TabIndex        =   8
      Top             =   1200
      Width           =   3015
      Begin MSComCtl2.DTPicker dtpConvertDate 
         Height          =   375
         Left            =   1320
         TabIndex        =   2
         Top             =   600
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
         Format          =   82313219
         CurrentDate     =   39416
      End
      Begin VB.Label Label1 
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
         TabIndex        =   3
         Top             =   600
         Width           =   675
      End
      Begin VB.Label txtCurrencyId 
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
         TabIndex        =   6
         Top             =   240
         Width           =   660
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
         TabIndex        =   5
         Top             =   240
         Width           =   945
      End
   End
   Begin MSDataGridLib.DataGrid dgdMain 
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   3201
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
      Height          =   660
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   8775
      _ExtentX        =   15478
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
            NumListImages   =   5
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTMCONVERTCURRENCY.frx":1F8A
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTMCONVERTCURRENCY.frx":43DC
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTMCONVERTCURRENCY.frx":682E
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTMCONVERTCURRENCY.frx":8C80
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTMCONVERTCURRENCY.frx":B0D2
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin MSDataGridLib.DataGrid dgdHeader 
      Height          =   1815
      Left            =   5160
      TabIndex        =   1
      Top             =   1200
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   3201
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
End
Attribute VB_Name = "frmTMCONVERTCURRENCY"
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
Private rstHeader As ADODB.Recordset

Private objMode As FunctionMode

Private strFormCaption As String

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub Form_Load()
    SetInitialization
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If frmTMCURRENCY.Parent Then
        frmTMCURRENCY.Parent = False
    End If
    
    mdlDatabase.CloseRecordset rstMain
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmTMCONVERTCURRENCY = Nothing
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

Private Sub dgdMain_HeadClick(ByVal ColIndex As Integer)
    If Not objMode = ViewMode Then Exit Sub
    
    rstMain.Sort = rstMain.Fields(ColIndex).Name
    
    If rstMain.RecordCount > 0 Then
        ArrangeHeaderGrid rstMain!CurrencyId
    End If
End Sub

Private Sub dgdHeader_HeadClick(ByVal ColIndex As Integer)
    If Not objMode = ViewMode Then Exit Sub
    
    rstHeader.Sort = rstHeader.Fields(ColIndex).Name
    
    If rstHeader.RecordCount > 0 Then
        FillText
    End If
End Sub

Private Sub dgdMain_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If Not objMode = ViewMode Then Exit Sub
    
    If rstMain.RecordCount > 0 Then
        ArrangeHeaderGrid rstMain!CurrencyId
    End If
End Sub

Private Sub dgdHeader_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If Not objMode = ViewMode Then Exit Sub
    
    If rstMain.RecordCount > 0 Then
        FillText
    End If
End Sub

Private Sub txtConvertValue_GotFocus()
    mdlProcedures.GotFocus Me.txtConvertValue
End Sub

Private Sub txtConvertValue_Validate(Cancel As Boolean)
    If Not IsNumeric(Me.txtConvertValue.Text) Then Me.txtConvertValue.Text = "0"
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
    
    Me.dtpConvertDate.CustomFormat = mdlGlobal.strFormatDate
    
    Me.txtCurrencyId.Caption = frmTMCURRENCY.CurrencyId
    
    strFormCaption = mdlText.strTMCONVERTCURRENCY
    
    Set rstMain = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "CurrencyId", mdlTable.CreateTMCURRENCY, False, "CurrencyId<>'" & frmTMCURRENCY.CurrencyId & "'")
    
    ArrangeGrid
    
    If rstMain.RecordCount > 0 Then
        ArrangeHeaderGrid rstMain!CurrencyId
    Else
        ArrangeHeaderGrid
    End If
    
    objMode = ViewMode
    
    SetMode
    
    FillText
End Sub

Private Sub ArrangeGrid()
    Set Me.dgdMain.DataSource = rstMain
    
    With Me.dgdMain
        .Columns(0).Width = 1200
        .Columns(0).Locked = True
        .Columns(0).Caption = "Mata Uang"
    End With
End Sub

Private Sub ArrangeHeaderGrid(Optional ByVal strCurrencyId As String = "")
    Set rstHeader = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "*", mdlTable.CreateTMCONVERTCURRENCY, , "CurrencyFromId='" & frmTMCURRENCY.CurrencyId & "' AND CurrencyToId='" & strCurrencyId & "'")
    
    If rstHeader.RecordCount > 0 Then
        FillText
    Else
        FillText True
    End If
    
    Set Me.dgdHeader.DataSource = rstHeader
    
    With Me.dgdHeader
        .Columns(0).Width = 0
        .Columns(0).Locked = True
        .Columns(0).Visible = False
        .Columns(1).Width = 1400
        .Columns(1).Locked = True
        .Columns(1).Caption = "Tanggal"
        .Columns(4).Width = 1300
        .Columns(4).Locked = True
        .Columns(4).Caption = "Nilai Tukar"
        
        Dim intCounter As Integer
        
        For intCounter = 2 To 3
            .Columns(intCounter).Width = 0
            .Columns(intCounter).Locked = True
            .Columns(intCounter).Visible = False
        Next intCounter
        
        For intCounter = 5 To .Columns.Count - 1
            .Columns(intCounter).Width = 0
            .Columns(intCounter).Locked = True
            .Columns(intCounter).Visible = False
        Next intCounter
    End With
    
    SetMode
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
    Me.dgdHeader.Enabled = blnFront
    
    With Me.tlbMain
        .Buttons(AddButton).Visible = blnFront
        .Buttons(UpdateButton).Visible = blnFront
        .Buttons(DeleteButton).Visible = blnFront
        .Buttons(SaveButton).Visible = blnBack
        .Buttons(CancelButton).Visible = blnBack
    End With
    
    Me.fraDetail.Enabled = blnBack
    
    If objMode = AddMode Then
        Me.fraHeader.Enabled = True
        
        mdlProcedures.SetControlMode Me, objMode
        
        Me.dtpConvertDate.SetFocus
    ElseIf objMode = UpdateMode Then
        Me.fraHeader.Enabled = False
        
        mdlProcedures.SetControlMode Me, objMode, False
        
        Me.txtConvertValue.SetFocus
    Else
        Me.fraHeader.Enabled = False
    End If
    
    If rstMain.RecordCount > 0 Then
        Me.tlbMain.Buttons(AddButton).Enabled = True
    Else
        Me.tlbMain.Buttons(AddButton).Enabled = False
    End If
    
    If rstHeader.RecordCount > 0 Then
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
        End If
    End If
    
    Me.Caption = mdlProcedures.SetCaptionMode(strFormCaption, objMode)
End Sub

Private Sub DeleteFunction()
    objMode = DeleteMode
    
    If mdlProcedures.SetMsgYesNo("Apakah Anda Yakin ?", mdlProcedures.SetCaptionMode(strFormCaption, objMode)) Then
        mdlDatabase.DeleteSingleRecord rstHeader
        
        mdiMain.CheckConvertCurrency
    End If
    
    objMode = ViewMode
    
    SetMode
    
    FillText
End Sub

Private Sub SaveFunction()
    If Not CheckValidation Then Exit Sub
    
    With rstHeader
        mdlDatabase.SearchRecordset rstHeader, "ConvertId", frmTMCURRENCY.CurrencyId & rstMain!CurrencyId & mdlProcedures.FormatDate(Me.dtpConvertDate.Value, "ddMMyyyy")
        
        If .EOF Then
            .AddNew
            
            !ConvertId = frmTMCURRENCY.CurrencyId & rstMain!CurrencyId & mdlProcedures.FormatDate(Me.dtpConvertDate.Value, "ddMMyyyy")
            !ConvertDate = mdlProcedures.FormatDate(Me.dtpConvertDate.Value, mdlGlobal.strFormatDate)
            !CurrencyFromId = frmTMCURRENCY.CurrencyId
            !CurrencyToId = rstMain!CurrencyId
            
            !CreateId = Trim(mdlGlobal.UserAuthority.UserId)
            !CreateDate = mdlProcedures.FormatDate(Now)
        End If
        
        !ConvertValue = mdlProcedures.GetCurrency(Me.txtConvertValue.Text)
        
        !UpdateId = Trim(mdlGlobal.UserAuthority.UserId)
        !UpdateDate = mdlProcedures.FormatDate(Now)
        
        .Update
    End With
    
    mdiMain.CheckConvertCurrency
    
    If objMode = AddMode Then
        mdlProcedures.SetControlMode Me, objMode
        
        Me.txtConvertValue.SetFocus
    ElseIf objMode = UpdateMode Then
        objMode = ViewMode
        
        SetMode
    End If
End Sub

Private Function CheckValidation() As Boolean
    If Not mdlProcedures.GetCurrency(Me.txtConvertValue.Text) > 0 Then
        MsgBox "Nilai Tukar Harap Diisi", vbOKOnly + vbExclamation, mdlProcedures.SetCaptionMode(strFormCaption, objMode)
        
        Me.txtConvertValue.SetFocus
        
        CheckValidation = False
        
        Exit Function
    End If
    
    With rstHeader
        If objMode = AddMode Then
            mdlDatabase.SearchRecordset rstHeader, "ConvertId", frmTMCURRENCY.CurrencyId & rstMain!CurrencyId & mdlProcedures.FormatDate(Me.dtpConvertDate.Value, "ddMMyyyy")
            
            If Not .EOF Then
                MsgBox "Nilai Tukar Sudah Ada", vbOKOnly + vbExclamation, mdlProcedures.SetCaptionMode(strFormCaption, objMode)
                
                Me.dtpConvertDate.SetFocus
                
                CheckValidation = False
                
                Exit Function
            End If
        End If
    End With
    
    CheckValidation = True
End Function

Private Sub FillText(Optional ByVal blnClear As Boolean = False)
    If blnClear Then
        Me.txtConvertValue.Text = ""
        
        Exit Sub
    End If
    
    With rstHeader
        If .RecordCount > 0 Then
            Me.dtpConvertDate.Value = mdlProcedures.FormatDate(!ConvertDate, mdlGlobal.strFormatDate)
            Me.txtConvertValue.Text = Trim(!ConvertValue)
        End If
    End With
End Sub
