VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmTHDOBUY 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7770
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10110
   Icon            =   "frmTHDOBUY.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7770
   ScaleWidth      =   10110
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
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   5880
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Frame fraSearch 
      Height          =   735
      Left            =   120
      TabIndex        =   25
      Top             =   1200
      Width           =   9855
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
         Left            =   8640
         TabIndex        =   2
         Top             =   240
         Width           =   1095
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
         Left            =   1800
         TabIndex        =   0
         Top             =   240
         Width           =   615
      End
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
         Left            =   6000
         MaxLength       =   4
         TabIndex        =   1
         Top             =   240
         Width           =   615
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
         TabIndex        =   14
         Top             =   240
         Width           =   495
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
         TabIndex        =   15
         Top             =   240
         Width           =   525
      End
   End
   Begin VB.Frame fraDetail 
      Height          =   1335
      Left            =   120
      TabIndex        =   27
      Top             =   4080
      Width           =   9855
      Begin VB.TextBox txtTransportNumber 
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
         Left            =   1800
         MaxLength       =   20
         TabIndex        =   8
         Top             =   240
         Width           =   2535
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
         Height          =   645
         Left            =   1800
         MaxLength       =   150
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         Top             =   600
         Width           =   7455
      End
      Begin VB.Label lblTransportNumber 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nomor Kendaraan"
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
         TabIndex        =   23
         Top             =   240
         Width           =   1560
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
         TabIndex        =   24
         Top             =   720
         Width           =   990
      End
   End
   Begin VB.Frame fraHeader 
      Height          =   2175
      Left            =   120
      TabIndex        =   26
      Top             =   1920
      Width           =   9855
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
         Left            =   9360
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   1680
         Width           =   375
      End
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
         Left            =   1800
         TabIndex        =   7
         Top             =   1680
         Width           =   7455
      End
      Begin VB.ComboBox cmbVendorId 
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
         Left            =   1800
         TabIndex        =   5
         Top             =   720
         Width           =   7455
      End
      Begin VB.CommandButton cmdPOId 
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
         Left            =   4680
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   1200
         Width           =   375
      End
      Begin VB.ComboBox cmbPOId 
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
         Left            =   1800
         TabIndex        =   6
         Top             =   1200
         Width           =   2775
      End
      Begin VB.TextBox txtDOId 
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
         Left            =   1800
         MaxLength       =   20
         TabIndex        =   3
         Top             =   240
         Width           =   2535
      End
      Begin MSComCtl2.DTPicker dtpDODate 
         Height          =   375
         Left            =   6000
         TabIndex        =   4
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
         Format          =   51314691
         CurrentDate     =   39330
      End
      Begin VB.Label txtPODate 
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
         Height          =   270
         Left            =   7320
         TabIndex        =   21
         Top             =   1320
         Width           =   2430
      End
      Begin VB.Label lblPODate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal PO"
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
         Left            =   6000
         TabIndex        =   20
         Top             =   1320
         Width           =   1020
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
         TabIndex        =   22
         Top             =   1680
         Width           =   675
      End
      Begin VB.Label lblVendorId 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pemasok"
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
         Top             =   720
         Width           =   825
      End
      Begin VB.Label lblPOId 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nomor PO"
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
         Top             =   1200
         Width           =   915
      End
      Begin VB.Label lblDODate 
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
         Left            =   4560
         TabIndex        =   17
         Top             =   240
         Width           =   675
      End
      Begin VB.Label lblDOId 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nomor DO"
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
         Top             =   240
         Width           =   915
      End
   End
   Begin MSFlexGridLib.MSFlexGrid flxDetail 
      Height          =   1335
      Left            =   120
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   5520
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   2355
      _Version        =   393216
      ForeColorFixed  =   -2147483640
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
      TabIndex        =   29
      Top             =   7140
      Width           =   10110
      _ExtentX        =   17833
      _ExtentY        =   1111
      ButtonWidth     =   609
      ButtonHeight    =   953
      Appearance      =   1
      _Version        =   393216
      Begin MSComctlLib.ImageList imlDetail 
         Left            =   9480
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
               Picture         =   "frmTHDOBUY.frx":1F8A
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTHDOBUY.frx":43DC
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTHDOBUY.frx":682E
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.Toolbar tlbMain 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   28
      Top             =   0
      Width           =   10110
      _ExtentX        =   17833
      _ExtentY        =   1111
      ButtonWidth     =   609
      ButtonHeight    =   953
      Appearance      =   1
      _Version        =   393216
      Begin MSComctlLib.ImageList imlMain 
         Left            =   9480
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
               Picture         =   "frmTHDOBUY.frx":8C80
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTHDOBUY.frx":B0D2
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTHDOBUY.frx":D524
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTHDOBUY.frx":F976
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTHDOBUY.frx":11DC8
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTHDOBUY.frx":1236A
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTHDOBUY.frx":147BC
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "frmTHDOBUY"
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
    [QtyLastColumn]
    [UnityIdColumn]
End Enum

Private rstMain As ADODB.Recordset

Private PrintTransaction As clsPRTTHDOBUY

Private objMode As FunctionMode

Private strFormCaption As String

Private blnDetailFill As Boolean

Private blnParent As Boolean
Private blnActivate As Boolean

Private Sub Form_Activate()
    If blnParent Then Exit Sub
    If blnActivate Then Exit Sub
    
    If rstMain.RecordCount > 0 Then
        blnParent = True
        
        mdlProcedures.CornerWindows Me
        
        mdlProcedures.ShowForm frmBRWTHDOBUY, False, True
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
        
        If frmTHSJBUY.Parent Then
            frmTHSJBUY.Parent = False
        End If
        
        mdlDatabase.CloseRecordset rstMain
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmTHDOBUY = Nothing
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

            mdlProcedures.ShowForm frmBRWTHDOBUY, False, True
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
            If Not mdlProcedures.IsValidComboData(Me.cmbPOId) Then
                MsgBox "Nomor PO Harap Dipilih Terlebih Dahulu", vbOKOnly + vbExclamation, mdlProcedures.SetCaptionMode(strFormCaption, objMode)
                
                Me.cmbPOId.SetFocus
                
                Exit Sub
            End If
            
            If blnParent Then Exit Sub

            blnParent = True

            mdlProcedures.CornerWindows Me

            mdlProcedures.ShowForm frmTDDOBUY, False, True
        Case UpdateDetailButton:
            If mdlProcedures.GetCurrency(Me.flxDetail.TextMatrix(Me.flxDetail.Row, QtyLastColumn)) > 0 Then
                MsgBox "Tidak Dapat Diubah" & vbCrLf & "Terdapat Data Yang Masih Terkait", vbCritical + vbOKOnly, mdlProcedures.SetCaptionMode(strFormCaption, objMode)

                Exit Sub
            End If

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
                If mdlProcedures.GetCurrency(Me.flxDetail.TextMatrix(Me.flxDetail.Row, QtyLastColumn)) > 0 Then
                    MsgBox "Tidak Dapat Dihapus" & vbCrLf & "Terdapat Data Yang Masih Terkait", vbCritical + vbOKOnly, mdlProcedures.SetCaptionMode(strFormCaption, objMode)

                    Exit Sub
                End If

                If Me.flxDetail.Rows > 2 Then
                    Me.flxDetail.RemoveItem Me.flxDetail.Row
                Else
                    Me.flxDetail.Rows = 1
                End If

                SetModeDetail
            End If
    End Select
End Sub

Private Sub cmdSearch_Click()
    If Not objMode = ViewMode Then Exit Sub
    
    Dim strMonth As String
    Dim strYear As String
    
    strMonth = mdlProcedures.GetComboData(Me.cmbMonth)
    strYear = Trim(Me.txtYear.Text)
    
    If CInt(mdlProcedures.FormatDate(Me.dtpDODate.MinDate, "M")) >= CInt(strMonth) Then
        If CInt(mdlProcedures.FormatDate(Me.dtpDODate.MinDate, "yyyy")) >= CInt(strYear) Then
            Me.dtpDODate.MinDate = mdlProcedures.SetDate(strMonth, strYear)
            
            Me.dtpDODate.MaxDate = mdlProcedures.SetDate(strMonth, strYear, , True)
        ElseIf CInt(mdlProcedures.FormatDate(Me.dtpDODate.MinDate, "yyyy")) < CInt(strYear) Then
            Me.dtpDODate.MaxDate = mdlProcedures.SetDate(strMonth, strYear, , True)
            
            Me.dtpDODate.MinDate = mdlProcedures.SetDate(strMonth, strYear)
        End If
    ElseIf CInt(mdlProcedures.FormatDate(Me.dtpDODate.MinDate, "M")) < CInt(strMonth) Then
        If CInt(mdlProcedures.FormatDate(Me.dtpDODate.MinDate, "yyyy")) > CInt(strYear) Then
            Me.dtpDODate.MinDate = mdlProcedures.SetDate(strMonth, strYear)
            
            Me.dtpDODate.MaxDate = mdlProcedures.SetDate(strMonth, strYear, , True)
        ElseIf CInt(mdlProcedures.FormatDate(Me.dtpDODate.MinDate, "yyyy")) >= CInt(strYear) Then
            Me.dtpDODate.MaxDate = mdlProcedures.SetDate(strMonth, strYear, , True)
            
            Me.dtpDODate.MinDate = mdlProcedures.SetDate(strMonth, strYear)
        End If
    End If

    SetRecordset
End Sub

Private Sub cmbPOId_Click()
    Me.flxDetail.Rows = 1
End Sub

Private Sub txtYear_GotFocus()
    mdlProcedures.GotFocus Me.txtYear
End Sub

Private Sub txtDOId_GotFocus()
    mdlProcedures.GotFocus Me.txtDOId
End Sub

Private Sub txtTransportNumber_GotFocus()
    mdlProcedures.GotFocus Me.txtTransportNumber
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

Private Sub txtDOId_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub dtpDODate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub cmbVendorId_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If mdlProcedures.IsValidComboData(Me.cmbVendorId) Then
            SendKeys "{TAB}"
        Else
            If blnParent Then Exit Sub
            
            blnParent = True
            
            mdlProcedures.CornerWindows Me
            
            mdlProcedures.ShowForm frmBRWTMVENDOR, False, True
        End If
    End If
End Sub

Private Sub cmbPOId_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If mdlProcedures.IsValidComboData(Me.cmbVendorId) Then
            If mdlProcedures.IsValidComboData(Me.cmbPOId) Then
                SendKeys "{TAB}"
            Else
                If blnParent Then Exit Sub
                
                blnParent = True
                
                mdlProcedures.CornerWindows Me
                
                mdlProcedures.ShowForm frmBRWTHPOBUY, False, True
            End If
        Else
            SendKeys "{TAB}"
        End If
    End If
End Sub

Private Sub cmbWarehouseId_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If mdlProcedures.IsValidComboData(Me.cmbWarehouseId) Then
            SendKeys "{TAB}"
        Else
            If blnParent Then Exit Sub
            
            blnParent = True
            
            mdlProcedures.CornerWindows Me
            
            mdlProcedures.ShowForm frmBRWTMWAREHOUSE, False, True
        End If
    End If
End Sub

Private Sub txtTransportNumber_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtInput_Change()
    Me.txtInput.Text = mdlProcedures.FormatCurrency(Me.txtInput.Text)

    Me.txtInput.SelStart = Len(Me.txtInput.Text)
End Sub

Private Sub flxDetail_Scroll()
    If Me.txtInput.Visible Then Me.txtInput.Visible = False
End Sub

Private Sub txtInput_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Dim curQty As Currency
        
        curQty = mdlProcedures.GetCurrency(Me.txtInput.Text)
        
        If curQty > 0 Then
            If curQty > mdlProcedures.GetCurrency(mdlTHPOBUY.GetTotalQtyPOBUY(Me.POIdCombo, Me.flxDetail.TextMatrix(Me.flxDetail.Row, ItemIdColumn))) Then
                Me.txtInput.Visible = False
                
                Exit Sub
            End If
        Else
            Me.txtInput.Visible = False
            
            Exit Sub
        End If
        
        Me.flxDetail.TextMatrix(Me.flxDetail.Row, Me.flxDetail.Col) = mdlProcedures.FormatCurrency(Me.txtInput.Text)
        
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

Private Sub cmbPOId_Validate(Cancel As Boolean)
    FillInfo
End Sub

Private Sub cmbVendorId_Validate(Cancel As Boolean)
    Me.FillComboTHPOBUY
End Sub

Private Sub cmdPOId_Click()
    If Not mdlGlobal.UserAuthority.IsMenuAccess(mdiMain.mnuTHPOBUY.Name) Then Exit Sub

    If blnParent Then Exit Sub

    blnParent = True

    mdlProcedures.CornerWindows Me

    mdlProcedures.ShowForm frmTHPOBUY, False
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
    
    Me.dtpDODate.CustomFormat = mdlGlobal.strFormatDate

    With Me.tlbDetail
        .AllowCustomize = False

        .ImageList = Me.imlDetail

        .Buttons.Add AddDetailButton, , "Tambah", , AddDetailButton
        .Buttons.Add UpdateDetailButton, , "Ubah", , UpdateDetailButton
        .Buttons.Add DeleteDetailButton, , "Hapus", , DeleteDetailButton
    End With

    FillCombo

    strFormCaption = mdlText.strTHDOBUY
    
    Dim strMonth As String
    Dim strYear As String
    
    strMonth = mdlProcedures.GetComboData(Me.cmbMonth)
    strYear = Trim(Me.txtYear.Text)
    
    Me.dtpDODate.MinDate = mdlProcedures.SetDate(strMonth, strYear)

    Me.dtpDODate.MaxDate = mdlProcedures.SetDate(strMonth, strYear, , True)

    ArrangeGrid

    SetRecordset
    
    blnParent = False
    blnActivate = False
End Sub

Private Sub SetRecordset()
    Set rstMain = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "*", mdlTable.CreateTHDOBUY, , "MONTH(DODate)=" & Me.MonthTrans & " AND YEAR(DODate)=" & Me.YearTrans, "DOId ASC")

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
        
        FillInfo True
        
        Me.flxDetail.Rows = 1

        SetModeDetail
        
        IncrementId
        
        Me.cmbPOId.Clear

        Me.txtDOId.SetFocus
    ElseIf objMode = UpdateMode Then
        Me.fraHeader.Enabled = False

        mdlProcedures.SetControlMode Me, objMode, False, Me.txtDOId.Name, Me.cmbMonth.Name & " | " & Me.txtYear.Name

        SetModeDetail

        Me.txtTransportNumber.SetFocus
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
            FillInfo True
            
            mdlProcedures.SetControlMode Me, objMode, , , Me.cmbMonth.Name & " | " & Me.txtYear.Name
        End If
    End If

    Me.Caption = mdlProcedures.SetCaptionMode(strFormCaption, objMode)
End Sub

Private Sub PrintFunction()
    If Not PrintTransaction Is Nothing Then
        Set PrintTransaction = Nothing
    End If

    Set PrintTransaction = New clsPRTTHDOBUY
    
    PrintTransaction.ImportToExcel rstMain!DOId
End Sub

Private Sub DeleteFunction()
    objMode = DeleteMode

    If mdlProcedures.SetMsgYesNo("Apakah Anda Yakin ?", mdlProcedures.SetCaptionMode(strFormCaption, objMode)) Then
        Dim strMessage As String
        
        strMessage = "Tidak Dapat Dihapus" & vbCrLf & "Terdapat Data Yang Masih Terkait ("
        
        If mdlDatabase.IsDataExists(mdlGlobal.conInventory, mdlTable.CreateTHSJBUY, "DOId='" & rstMain!DOId & "'") Then
            MsgBox strMessage & mdlText.strTHSJBUY & ")", vbCritical + vbOKOnly, mdlProcedures.SetCaptionMode(strFormCaption, objMode)
        Else
            mdlTHDOBUY.DeleteTHDOBUY rstMain
            
            mdlDatabase.DeleteRecordQuery mdlGlobal.conInventory, mdlTable.CreateTDDOBUY, "DOId='" & rstMain!DOId & "'"
            mdlDatabase.DeleteSingleRecord rstMain

            If frmTHSJBUY.Parent Then
                frmTHSJBUY.FillComboTHDOBUY
            End If
            
            frmMenu.SetRecycle
        End If
    End If

    objMode = ViewMode

    SetMode

    FillText
End Sub

Private Sub SaveFunction()
    If Not CheckValidation Then Exit Sub

    With rstMain
        mdlDatabase.SearchRecordset rstMain, "DOId", mdlProcedures.RepDupText(Trim(Me.txtDOId.Text))

        If .EOF Then
            .AddNew

            !DOId = mdlProcedures.RepDupText(Trim(Me.txtDOId.Text))
            !DODate = mdlProcedures.FormatDate(Me.dtpDODate.Value)
            !POId = Me.POIdCombo
            !WarehouseId = Me.WarehouseIdCombo

            !CreateId = Trim(mdlGlobal.UserAuthority.UserId)
            !CreateDate = mdlProcedures.FormatDate(Now)
        End If

        !TransportNumber = mdlProcedures.RepDupText(Trim(Me.txtTransportNumber.Text))
        !Notes = mdlProcedures.RepDupText(Trim(Me.txtNotes.Text))

        !UpdateId = Trim(mdlGlobal.UserAuthority.UserId)
        !UpdateDate = mdlProcedures.FormatDate(Now)

        .Update
    End With

    SaveDetailFunction

    If objMode = AddMode Then
        mdlProcedures.SetControlMode Me, objMode, , , Me.cmbMonth.Name & " | " & Me.txtYear.Name
        
        FillInfo True
        
        IncrementId
        
        Me.cmbPOId.Clear
        
        Me.txtDOId.SetFocus
    ElseIf objMode = UpdateMode Then
        objMode = ViewMode

        SetMode
    End If
    
    If frmTHSJBUY.Parent Then
        frmTHSJBUY.FillComboTHDOBUY
    End If
End Sub

Private Sub SaveDetailFunction()
    Dim strDOId As String

    strDOId = mdlProcedures.RepDupText(Trim(Me.txtDOId.Text))

    mdlDatabase.DeleteRecordQuery mdlGlobal.conInventory, mdlTable.CreateTDDOBUY, "DOId='" & strDOId & "'"

    If Not Me.flxDetail.Rows > 1 Then Exit Sub

    Dim rstTemp As ADODB.Recordset

    Set rstTemp = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "*", mdlTable.CreateTDDOBUY, , "DOId='" & strDOId & "'")

    With rstTemp
        If Not .RecordCount > 0 Then
            Dim lngRow As Long

            For lngRow = 1 To Me.flxDetail.Rows - 1
                .AddNew

                !DOId = strDOId
                !DODtlId = !DOId & Trim(Me.flxDetail.TextMatrix(lngRow, ItemIdColumn))
                !ItemId = Trim(Me.flxDetail.TextMatrix(lngRow, ItemIdColumn))
                !Qty = mdlProcedures.GetCurrency(Me.flxDetail.TextMatrix(lngRow, QtyColumn))

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
    If mdlProcedures.RepDupText(Trim(Me.txtDOId.Text)) = "" Then
        MsgBox "Nomor DO Harap Diisi", vbOKOnly + vbExclamation, mdlProcedures.SetCaptionMode(strFormCaption, objMode)
        
        Me.txtDOId.SetFocus

        CheckValidation = False

        Exit Function
    ElseIf Not mdlProcedures.IsValidComboData(Me.cmbPOId) Then
        MsgBox "Nomor PO Harap Dipilih", vbOKOnly + vbExclamation, mdlProcedures.SetCaptionMode(strFormCaption, objMode)
        
        Me.cmbPOId.SetFocus

        CheckValidation = False

        Exit Function
    ElseIf Not mdlProcedures.IsValidComboData(Me.cmbWarehouseId) Then
        MsgBox "Gudang Harap Dipilih", vbOKOnly + vbExclamation, mdlProcedures.SetCaptionMode(strFormCaption, objMode)
        
        Me.cmbWarehouseId.SetFocus

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
            mdlDatabase.SearchRecordset rstMain, "DOId", mdlProcedures.RepDupText(Trim(Me.txtDOId.Text))

            If Not .EOF Then
                MsgBox "Nomor DO Sudah Ada", vbOKOnly + vbExclamation, mdlProcedures.SetCaptionMode(strFormCaption, objMode)
                
                Me.txtDOId.SetFocus
                
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
            Me.txtDOId.Text = Trim(!DOId)
            Me.dtpDODate.Value = mdlProcedures.FormatDate(!DODate, mdlGlobal.strFormatDate)
            
            mdlProcedures.SetComboData _
                Me.cmbVendorId, _
                mdlDatabase.GetFieldData(mdlGlobal.conInventory, "VendorId", mdlTable.CreateTHPOBUY, "POId='" & !POId & "'")
            
            Me.FillComboTHPOBUY
            
            mdlProcedures.SetComboData Me.cmbPOId, !POId
            mdlProcedures.SetComboData Me.cmbWarehouseId, !WarehouseId
            
            FillInfo
            
            Me.txtTransportNumber.Text = Trim(!TransportNumber)
            Me.txtNotes.Text = Trim(!Notes)

            FillGrid !DOId
        Else
            FillGrid
        End If
    End With
End Sub

Private Sub FillGrid(Optional ByVal strDOId As String = "")
    With Me.flxDetail
        .Rows = 1

        Dim rstTemp As ADODB.Recordset

        Set rstTemp = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "*", mdlTable.CreateTDDOBUY, False, "DOId='" & strDOId & "'")

        While Not rstTemp.EOF
            .Rows = .Rows + 1

            .TextMatrix(.Rows - 1, ItemIdColumn) = rstTemp!ItemId
            .TextMatrix(.Rows - 1, NameColumn) = mdlDatabase.GetFieldData(mdlGlobal.conInventory, "Name", mdlTable.CreateTMITEM, "ItemId='" & rstTemp!ItemId & "'")
            .TextMatrix(.Rows - 1, QtyColumn) = mdlProcedures.FormatCurrency(rstTemp!Qty)
            .TextMatrix(.Rows - 1, QtyLastColumn) = mdlProcedures.FormatCurrency(CStr(mdlTHSJBUY.GetQtyDOFromSJBUY(strDOId, rstTemp!ItemId)))
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

Private Sub FillInfo(Optional ByVal blnClear As Boolean = False)
    If blnClear Then
        Me.txtPODate.Caption = ""
        
        Exit Sub
    End If
    
    Dim strPOId As String
    Dim strPODate As String
    
    strPOId = Me.POIdCombo
    strPODate = mdlDatabase.GetFieldData(mdlGlobal.conInventory, "PODate", mdlTable.CreateTHPOBUY, "POId='" & strPOId & "'")
    
    If Trim(strPODate) = "" Then
        Me.txtPODate.Caption = strPODate
    Else
        Me.txtPODate.Caption = mdlProcedures.FormatDate(strPODate, "dd MMMM yyyy")
    End If
End Sub

Private Sub FillCombo()
    FillComboSearch
    
    Me.FillComboTMVENDOR
    Me.FillComboTMWAREHOUSE
End Sub

Private Sub FillComboSearch()
    mdlProcedures.FillComboMonth Me.cmbMonth, , , mdlProcedures.FormatDate(Now, "M")

    Me.txtYear.Text = mdlProcedures.FormatDate(Now, "yyyy")
End Sub

Public Sub FillComboTHPOBUY()
    Dim rstTemp As ADODB.Recordset
    
    Set rstTemp = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "POId", mdlTable.CreateTHPOBUY, False, "VendorId='" & mdlProcedures.GetComboData(Me.cmbVendorId) & "'")
    
    mdlProcedures.FillComboData Me.cmbPOId, rstTemp

    mdlDatabase.CloseRecordset rstTemp
End Sub

Public Sub FillComboTMVENDOR()
    Dim rstTemp As ADODB.Recordset

    Set rstTemp = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "VendorId, Name", mdlTable.CreateTMVENDOR, False)

    mdlProcedures.FillComboData Me.cmbVendorId, rstTemp

    mdlDatabase.CloseRecordset rstTemp
End Sub

Public Sub FillComboTMWAREHOUSE()
    Dim rstTemp As ADODB.Recordset

    Set rstTemp = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "WarehouseId, Name", mdlTable.CreateTMWAREHOUSE, False, , "WarehouseSet DESC")

    mdlProcedures.FillComboData Me.cmbWarehouseId, rstTemp

    mdlDatabase.CloseRecordset rstTemp
End Sub

Private Sub ArrangeGrid()
    With Me.flxDetail
        .AllowUserResizing = flexResizeColumns
        
        .Rows = 1
        .Cols = UnityIdColumn + 1

        .ColWidth(BlankColumn) = 300
        .ColWidth(ItemIdColumn) = 1000
        .ColWidth(NameColumn) = 3600
        .ColWidth(QtyColumn) = 1800
        .ColWidth(QtyLastColumn) = 1800
        .ColWidth(UnityIdColumn) = 1000
        
        .ColAlignment(ItemIdColumn) = flexAlignLeftCenter
        .ColAlignment(NameColumn) = flexAlignLeftCenter
        .ColAlignment(QtyColumn) = flexAlignRightCenter
        .ColAlignment(QtyLastColumn) = flexAlignRightCenter
        .ColAlignment(UnityIdColumn) = flexAlignLeftCenter

        .TextMatrix(0, ItemIdColumn) = "Kode"
        .TextMatrix(0, NameColumn) = "Nama"
        .TextMatrix(0, QtyColumn) = "Qty"
        .TextMatrix(0, QtyLastColumn) = "Qty (Terima)"
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
            .TextMatrix(.Rows - 1, QtyLastColumn) = mdlProcedures.FormatCurrency("0")
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

Private Sub IncrementId()
    Dim intCounter As Integer
    
    intCounter = 0
    
    Dim rstTemp As ADODB.Recordset
    
    Set rstTemp = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "DOId", mdlTable.CreateTHDOBUY, , "MONTH(DODate)=" & Me.MonthTrans & " AND YEAR(DODate)=" & Me.YearTrans)
    
    With rstTemp
        If .RecordCount > 0 Then
            Dim intNoSeq As Integer
            
            While Not .EOF
                intNoSeq = ExtractSequential(Trim(!DOId))
                
                If Not intNoSeq < 0 Then
                    If intCounter < intNoSeq Then
                        intCounter = intNoSeq
                    End If
                End If
                
                .MoveNext
            Wend
        End If
    End With
    
    Dim strReferencesNumber As String
    
    strReferencesNumber = mdlText.strDOIDINIT & "/" & mdlText.strBUYINIT
    
    Set rstTemp = _
        mdlDatabase.OpenRecordset( _
            mdlGlobal.conInventory, _
            "ReferencesNumber", _
            mdlTable.CreateTHRECYCLE, _
            False, _
            "LEFT(ReferencesNumber, " & Len(strReferencesNumber) & ")='" & strReferencesNumber & "'")
            
    While Not rstTemp.EOF
        intNoSeq = ExtractSequential(Trim(rstTemp!ReferencesNumber))
        
        If Not intNoSeq < 0 Then
            If intCounter < intNoSeq Then
                intCounter = intNoSeq
            End If
        End If
        
        rstTemp.MoveNext
    Wend
    
    intCounter = intCounter + 1
    
    Me.txtDOId.Text = ZipId(intCounter)
    
    mdlDatabase.CloseRecordset rstTemp
End Sub

Private Function ExtractSequential(ByVal strDOId As String) As Integer
    Dim strExtract() As String
    
    strExtract = Split(strDOId, "/")
    
    If Not UBound(strExtract) = 4 Then
        ExtractSequential = 0
        
        Exit Function
    End If
    
    If Not IsNumeric(strExtract(4)) Then
        ExtractSequential = 0
        
        Exit Function
    End If
    
    ExtractSequential = CInt(strExtract(4))
End Function

Private Function ZipId(ByVal intNoSeq As Integer) As String
    ZipId = _
        mdlText.strDOIDINIT & _
        "/" & _
        mdlText.strBUYINIT & _
        "/" & _
        mdlProcedures.FormatDate(Me.dtpDODate.Value, "yyyy") & _
        "/" & _
        mdlProcedures.FormatDate(Me.dtpDODate.Value, "MM") & _
        "/" & _
        mdlProcedures.FormatNumber(intNoSeq, "0000")
End Function

Public Property Get Parent() As Boolean
    Parent = blnParent
End Property

Public Property Get MonthTrans() As String
    MonthTrans = mdlProcedures.GetComboData(Me.cmbMonth)
End Property

Public Property Get YearTrans() As String
    YearTrans = Trim(Me.txtYear.Text)
End Property

Public Property Get DOId() As String
    If rstMain Is Nothing Then Exit Property

    If rstMain.RecordCount > 0 Then
        DOId = rstMain!DOId
    End If
End Property

Public Property Get POId() As String
    If rstMain Is Nothing Then Exit Property

    If rstMain.RecordCount > 0 Then
        POId = rstMain!POId
    End If
End Property

Public Property Get VendorIdCombo() As String
    VendorIdCombo = mdlProcedures.GetComboData(Me.cmbVendorId)
End Property

Public Property Get POIdCombo() As String
    POIdCombo = mdlProcedures.GetComboData(Me.cmbPOId)
End Property

Public Property Get WarehouseIdCombo() As String
    WarehouseIdCombo = mdlProcedures.GetComboData(Me.cmbWarehouseId)
End Property

Public Property Let Parent(ByVal blnEnable As Boolean)
    blnParent = blnEnable

    If Not blnEnable Then mdlProcedures.CenterWindows Me
End Property

Public Property Let DOId(ByVal strDOId As String)
    If rstMain Is Nothing Then Exit Property

    mdlDatabase.SearchRecordset rstMain, "DOId", strDOId

    If Not rstMain.EOF Then FillText
End Property

Public Property Let VendorIdCombo(ByVal strVendorId As String)
    mdlProcedures.SetComboData Me.cmbVendorId, strVendorId
End Property

Public Property Let POIdCombo(ByVal strPOId As String)
    mdlProcedures.SetComboData Me.cmbPOId, strPOId
End Property

Public Property Let WarehouseIdCombo(ByVal strWarehouseId As String)
    mdlProcedures.SetComboData Me.cmbWarehouseId, strWarehouseId
End Property
