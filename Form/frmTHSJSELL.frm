VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmTHSJSELL 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7770
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   9990
   Icon            =   "frmTHSJSELL.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7770
   ScaleWidth      =   9990
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
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   5880
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Frame fraHeader 
      Height          =   2055
      Left            =   120
      TabIndex        =   29
      Top             =   1920
      Width           =   9735
      Begin VB.CommandButton cmdReferencesNumber 
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
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   1200
         Width           =   375
      End
      Begin VB.TextBox txtReferencesNumber 
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
         Left            =   6360
         MaxLength       =   20
         TabIndex        =   7
         Top             =   1200
         Width           =   2535
      End
      Begin VB.ComboBox cmbSOId 
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
         TabIndex        =   6
         Top             =   1200
         Width           =   2775
      End
      Begin VB.CommandButton cmdSOId 
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
         Left            =   4320
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   1200
         Width           =   375
      End
      Begin VB.ComboBox cmbCustomerId 
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
         TabIndex        =   5
         Top             =   720
         Width           =   7455
      End
      Begin VB.TextBox txtSJId 
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
         TabIndex        =   3
         Top             =   240
         Width           =   2535
      End
      Begin MSComCtl2.DTPicker dtpSJDate 
         Height          =   375
         Left            =   6360
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
         Format          =   81002499
         CurrentDate     =   39330
      End
      Begin VB.Label lblReferencesNumber 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nomor Referensi"
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
         Left            =   4800
         TabIndex        =   20
         Top             =   1200
         Width           =   1440
      End
      Begin VB.Label txtDateLine 
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
         Left            =   6360
         TabIndex        =   24
         Top             =   1680
         Width           =   2430
      End
      Begin VB.Label lblDateLine 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal Jatuh Tempo"
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
         Left            =   4320
         TabIndex        =   23
         Top             =   1680
         Width           =   1845
      End
      Begin VB.Label txtPOId 
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
         Left            =   1320
         TabIndex        =   22
         Top             =   1680
         Width           =   2430
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
         TabIndex        =   21
         Top             =   1680
         Width           =   915
      End
      Begin VB.Label lblSOId 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nomor SO"
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
         TabIndex        =   18
         Top             =   720
         Width           =   840
      End
      Begin VB.Label lblSJId 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nomor SJ"
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
         Width           =   855
      End
      Begin VB.Label lblSJDate 
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
         Left            =   4320
         TabIndex        =   17
         Top             =   240
         Width           =   675
      End
   End
   Begin VB.Frame fraDetail 
      Height          =   1455
      Left            =   120
      TabIndex        =   30
      Top             =   3960
      Width           =   9735
      Begin VB.CommandButton cmdDeliveryId 
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
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   240
         Width           =   375
      End
      Begin VB.ComboBox cmbDeliveryId 
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
         TabIndex        =   8
         Top             =   240
         Width           =   7815
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
         Left            =   1320
         MaxLength       =   150
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         Top             =   720
         Width           =   7815
      End
      Begin VB.Label lblDeliveryId 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Alamat Kirim"
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
         TabIndex        =   25
         Top             =   240
         Width           =   1125
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
         TabIndex        =   26
         Top             =   720
         Width           =   990
      End
   End
   Begin VB.Frame fraSearch 
      Height          =   735
      Left            =   120
      TabIndex        =   28
      Top             =   1200
      Width           =   9735
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
         Left            =   6360
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
         Left            =   1320
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
         Left            =   8520
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
         Left            =   4320
         TabIndex        =   15
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
         TabIndex        =   14
         Top             =   240
         Width           =   495
      End
   End
   Begin MSFlexGridLib.MSFlexGrid flxDetail 
      Height          =   1335
      Left            =   120
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   5520
      Width           =   9735
      _ExtentX        =   17171
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
      Height          =   660
      Left            =   0
      TabIndex        =   31
      Top             =   7110
      Width           =   9990
      _ExtentX        =   17621
      _ExtentY        =   1164
      ButtonWidth     =   609
      ButtonHeight    =   1005
      Appearance      =   1
      _Version        =   393216
      Begin MSComctlLib.ImageList imlDetail 
         Left            =   9240
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
               Picture         =   "frmTHSJSELL.frx":1F8A
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTHSJSELL.frx":43DC
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTHSJSELL.frx":682E
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.Toolbar tlbMain 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   32
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
            NumListImages   =   8
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTHSJSELL.frx":8C80
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTHSJSELL.frx":B0D2
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTHSJSELL.frx":D524
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTHSJSELL.frx":F976
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTHSJSELL.frx":11DC8
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTHSJSELL.frx":1391A
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTHSJSELL.frx":13EBC
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTHSJSELL.frx":1630E
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "frmTHSJSELL"
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
    [FormButton]
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
    [WarehouseIdColumn]
End Enum

Private rstMain As ADODB.Recordset

Private PrintTransaction As clsPRTTHSJSELL
Private FormTransaction As clsFRMTHSJSELL

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
        
        mdlProcedures.ShowForm frmBRWTHSJSELL, False, True
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
    Set frmTHSJSELL = Nothing
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
        Case FormButton:
            FormFunction
        Case BrowseButton:
            If blnParent Then Exit Sub
            
            blnParent = True
            
            mdlProcedures.CornerWindows Me
            
            mdlProcedures.ShowForm frmBRWTHSJSELL, False, True
        Case SaveButton:
            SaveFunction
        Case CancelButton:
            objMode = ViewMode
            
            SetMode
            
            FillText
    End Select
End Sub

Private Sub tlbDetail_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case AddDetailButton:
            If Not mdlProcedures.IsValidComboData(Me.cmbSOId) Then
                MsgBox "Nomor SO Harap Dipilih Terlebih Dahulu", vbOKOnly + vbExclamation, mdlProcedures.SetCaptionMode(strFormCaption, objMode)
                
                Me.cmbSOId.SetFocus
                
                Exit Sub
            End If
            
            If blnParent Then Exit Sub
            
            blnParent = True
            
            mdlProcedures.CornerWindows Me
            
            mdlProcedures.ShowForm frmTDSJSELL, False, True
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
    
    If CInt(mdlProcedures.FormatDate(Me.dtpSJDate.MinDate, "M")) >= CInt(strMonth) Then
        If CInt(mdlProcedures.FormatDate(Me.dtpSJDate.MinDate, "yyyy")) >= CInt(strYear) Then
            Me.dtpSJDate.MinDate = mdlProcedures.SetDate(strMonth, strYear)
            
            Me.dtpSJDate.MaxDate = mdlProcedures.SetDate(strMonth, strYear, , True)
        ElseIf CInt(mdlProcedures.FormatDate(Me.dtpSJDate.MinDate, "yyyy")) < CInt(strYear) Then
            Me.dtpSJDate.MaxDate = mdlProcedures.SetDate(strMonth, strYear, , True)
            
            Me.dtpSJDate.MinDate = mdlProcedures.SetDate(strMonth, strYear)
        End If
    ElseIf CInt(mdlProcedures.FormatDate(Me.dtpSJDate.MinDate, "M")) < CInt(strMonth) Then
        If CInt(mdlProcedures.FormatDate(Me.dtpSJDate.MinDate, "yyyy")) > CInt(strYear) Then
            Me.dtpSJDate.MinDate = mdlProcedures.SetDate(strMonth, strYear)
            
            Me.dtpSJDate.MaxDate = mdlProcedures.SetDate(strMonth, strYear, , True)
        ElseIf CInt(mdlProcedures.FormatDate(Me.dtpSJDate.MinDate, "yyyy")) >= CInt(strYear) Then
            Me.dtpSJDate.MaxDate = mdlProcedures.SetDate(strMonth, strYear, , True)
            
            Me.dtpSJDate.MinDate = mdlProcedures.SetDate(strMonth, strYear)
        End If
    End If
    
    SetRecordset
End Sub

Private Sub cmbSOId_Click()
    Me.flxDetail.Rows = 1
End Sub

Private Sub txtYear_GotFocus()
    mdlProcedures.GotFocus Me.txtYear
End Sub

Private Sub txtSJId_GotFocus()
    mdlProcedures.GotFocus Me.txtSJId
End Sub

Private Sub txtReferencesNumber_GotFocus()
    mdlProcedures.GotFocus Me.txtReferencesNumber
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

Private Sub txtSJId_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub dtpSJDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub cmbCustomerId_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If mdlProcedures.IsValidComboData(Me.cmbCustomerId) Then
            SendKeys "{TAB}"
        Else
            If blnParent Then Exit Sub
            
            blnParent = True
            
            mdlProcedures.CornerWindows Me
            
            mdlProcedures.ShowForm frmBRWTMCUSTOMER
        End If
    End If
End Sub

Private Sub cmbSOId_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If mdlProcedures.IsValidComboData(Me.cmbCustomerId) Then
            If mdlProcedures.IsValidComboData(Me.cmbSOId) Then
                SendKeys "{TAB}"
            Else
                If blnParent Then Exit Sub
                
                blnParent = True
                
                mdlProcedures.CornerWindows Me
                
                mdlProcedures.ShowForm frmBRWTHSOSELL, False, True
            End If
        Else
            SendKeys "{TAB}"
        End If
    End If
End Sub

Private Sub txtReferencesNumber_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub cmbDeliveryId_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys "{TAB}"
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
        Dim curStockQty As Currency
        
        curQty = mdlProcedures.GetCurrency(Me.txtInput.Text)
        curMinStock = mdlProcedures.GetCurrency(mdlDatabase.GetFieldData(mdlGlobal.conInventory, "MinStock", mdlTable.CreateTMITEM, "ItemId='" & Me.flxDetail.TextMatrix(Me.flxDetail.Row, ItemIdColumn) & "'"))
        curStockQty = mdlTransaction.CheckStock(Me.flxDetail.TextMatrix(Me.flxDetail.Row, ItemIdColumn))
        
        If curQty > 0 Then
            If curQty > curStockQty Then
                Me.txtInput.Visible = False
                
                Exit Sub
            Else
                If curMinStock > 0 Then
                    If objMode = AddMode Then
                        If (curStockQty - curQty) < curMinStock Then
                            Me.txtInput.Visible = False
                            
                            Exit Sub
                        End If
                    ElseIf objMode = UpdateMode Then
                        If (curStockQty - (curQty - mdlTHSJSELL.GetTotalQtySJSELL(rstMain!SJId, Me.flxDetail.TextMatrix(Me.flxDetail.Row, ItemIdColumn)))) < curMinStock Then
                            Me.txtInput.Visible = False
                            
                            Exit Sub
                        End If
                    End If
                End If
            End If
            
            Dim curTotalQtySO As Currency
            
            curTotalQtySO = mdlTHSOSELL.GetTotalQtySOSELL(Me.SOIdCombo, Me.flxDetail.TextMatrix(Me.flxDetail.Row, ItemIdColumn))

            Dim curQtySO As Currency
            
            curQtySO = mdlTHSJSELL.GetQtySOFromSJSELL(Me.SOIdCombo, Me.flxDetail.TextMatrix(Me.flxDetail.Row, ItemIdColumn))
            
            If objMode = UpdateMode Then
                curQtySO = curQtySO - mdlTHSJSELL.GetTotalQtySJSELL(rstMain!SJId, _
                    Me.flxDetail.TextMatrix(Me.flxDetail.Row, ItemIdColumn))
            End If
            
            If (curQty + curQtySO) > curTotalQtySO Then
                Me.txtInput.Visible = False
                
                Exit Sub
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

Private Sub cmbCustomerId_Validate(Cancel As Boolean)
    Me.FillComboTHSOSELL
    Me.FillComboTMDELIVERYCUSTOMER
    
    FillInfo
End Sub

Private Sub cmbSOId_Validate(Cancel As Boolean)
    FillInfo
End Sub

Private Sub cmdSOId_Click()
    If Not mdlGlobal.UserAuthority.IsMenuAccess(mdiMain.mnuTHSOSELL.Name) Then Exit Sub

    If blnParent Then Exit Sub
    
    blnParent = True
    
    mdlProcedures.CornerWindows Me
    
    mdlProcedures.ShowForm frmTHSOSELL, False
End Sub

Private Sub cmdReferencesNumber_Click()
    Dim rstTemp As ADODB.Recordset
    
    Set rstTemp = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "ReferencesNumber", mdlTable.CreateTHSJSELL, , "MONTH(SJDate)=" & Me.MonthTrans & " AND YEAR(SJDate)=" & Me.YearTrans, "SJId ASC")
    
    Dim intCounter As Integer
    
    intCounter = 0
    
    With rstTemp
        If .RecordCount > 0 Then
            Dim intNoSeq As Integer
            
            While Not .EOF
                intNoSeq = mdlTransaction.ExtractSequential(Trim(!ReferencesNumber), 3)
                
                If Not intNoSeq < 0 Then
                    If intCounter < intNoSeq Then
                        intCounter = intNoSeq
                    End If
                End If
                
                .MoveNext
            Wend
        End If
    End With
    
    intCounter = intCounter + 1
    
    Me.txtReferencesNumber.Text = ZipReferencesNumber(intCounter)
    
    mdlDatabase.CloseRecordset rstTemp
End Sub

Private Sub cmdDeliveryId_Click()
    If Not mdlGlobal.UserAuthority.IsMenuAccess(mdiMain.mnuTMCUSTOMER.Name) Then Exit Sub
    
    If Not mdlProcedures.IsValidComboData(Me.cmbCustomerId) Then Exit Sub
    
    If blnParent Then Exit Sub
    
    blnParent = True
    
    mdlProcedures.CornerWindows Me
    
    mdlProcedures.ShowForm frmTMDELIVERYCUSTOMER, False, True
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
        .Buttons.Add FormButton, , "Form", , FormButton
        .Buttons.Add BrowseButton, , "Daftar", , BrowseButton
        .Buttons.Add SaveButton, , "Simpan", , SaveButton
        .Buttons.Add CancelButton, , "Batal", , CancelButton
    End With
    
    Me.dtpSJDate.CustomFormat = mdlGlobal.strFormatDate
    
    With Me.tlbDetail
        .AllowCustomize = False
        
        .ImageList = Me.imlDetail
        
        .Buttons.Add AddDetailButton, , "Tambah", , AddDetailButton
        .Buttons.Add UpdateDetailButton, , "Ubah", , UpdateDetailButton
        .Buttons.Add DeleteDetailButton, , "Hapus", , DeleteDetailButton
    End With
    
    FillCombo
    
    strFormCaption = mdlText.strTHSJSELL
    
    Dim strMonth As String
    Dim strYear As String
    
    strMonth = mdlProcedures.GetComboData(Me.cmbMonth)
    strYear = Trim(Me.txtYear.Text)
    
    Me.dtpSJDate.MinDate = mdlProcedures.SetDate(strMonth, strYear)
    Me.dtpSJDate.MaxDate = mdlProcedures.SetDate(strMonth, strYear, , True)
    
    ArrangeGrid
    
    SetRecordset
    
    blnParent = False
    blnActivate = False
End Sub

Private Sub SetRecordset()
    Set rstMain = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "*", mdlTable.CreateTHSJSELL, , "MONTH(SJDate)=" & Me.MonthTrans & " AND YEAR(SJDate)=" & Me.YearTrans, "SJId ASC")
    
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
        .Buttons(FormButton).Visible = blnFront
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
        
        Me.FillComboTHSOSELL
        Me.FillComboTMDELIVERYCUSTOMER
        
        IncrementId
        
        Me.txtSJId.SetFocus
    ElseIf objMode = UpdateMode Then
        Me.fraHeader.Enabled = False
        
        mdlProcedures.SetControlMode Me, objMode, False, Me.txtSJId.Name, Me.cmbMonth.Name & " | " & Me.txtYear.Name
        
        SetModeDetail
        
        Me.cmbDeliveryId.SetFocus
    Else
        Me.fraHeader.Enabled = False
    End If
    
    If rstMain.RecordCount > 0 Then
        With Me.tlbMain
            .Buttons(UpdateButton).Enabled = True
            .Buttons(DeleteButton).Enabled = True
            .Buttons(PrintButton).Enabled = True
            .Buttons(FormButton).Enabled = True
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
            .Buttons(FormButton).Enabled = False
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

    Set PrintTransaction = New clsPRTTHSJSELL
    
    PrintTransaction.ImportToExcel rstMain!SJId
End Sub

Private Sub FormFunction()
    If Not FormTransaction Is Nothing Then
        Set FormTransaction = Nothing
    End If

    Set FormTransaction = New clsFRMTHSJSELL
    
    FormTransaction.ImportToExcel rstMain
End Sub

Private Sub DeleteFunction()
    objMode = DeleteMode
    
    If mdlProcedures.SetMsgYesNo("Apakah Anda Yakin ?", mdlProcedures.SetCaptionMode(strFormCaption, objMode)) Then
        Dim rstTemp As ADODB.Recordset
        
        Set rstTemp = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "ItemId, WarehouseId", mdlTable.CreateTDSJSELL, False, "SJId='" & rstMain!SJId & "'")
        
        With rstTemp
            While Not .EOF
                mdlTransaction.UpdateStock !ItemId, _
                    !WarehouseId, _
                    rstMain!SJId, _
                    Me.dtpSJDate.Value
                
                .MoveNext
            Wend
        End With
        
        mdlDatabase.CloseRecordset rstTemp
        
        mdlTHSJSELL.DeleteTHSJSELL rstMain
        
        mdlDatabase.DeleteRecordQuery mdlGlobal.conInventory, mdlTable.CreateTDSJSELL, "SJId='" & rstMain!SJId & "'"
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
        mdlDatabase.SearchRecordset rstMain, "SJId", mdlProcedures.RepDupText(Trim(Me.txtSJId.Text))
        
        If .EOF Then
            .AddNew
            
            !SJId = mdlProcedures.RepDupText(Trim(Me.txtSJId.Text))
            !SJDate = mdlProcedures.FormatDate(Me.dtpSJDate.Value)
            
            !SOId = Me.SOIdCombo
            !ReferencesNumber = mdlProcedures.RepDupText(Trim(Me.txtReferencesNumber.Text))
            
            !CreateId = Trim(mdlGlobal.UserAuthority.UserId)
            !CreateDate = mdlProcedures.FormatDate(Now)
        End If
        
        !DeliveryId = mdlProcedures.GetComboData(Me.cmbDeliveryId)
        !Notes = mdlProcedures.RepDupText(Trim(Me.txtNotes.Text))
        
        !UpdateId = Trim(mdlGlobal.UserAuthority.UserId)
        !UpdateDate = mdlProcedures.FormatDate(Now)
        
        .Update
    End With
    
    SaveDetailFunction
    
    If objMode = AddMode Then
        mdlProcedures.SetControlMode Me, objMode, , , Me.cmbMonth.Name & " | " & Me.txtYear.Name
        
        FillInfo True
        
        Me.flxDetail.Rows = 1
        
        SetModeDetail
        
        IncrementId
        
        Me.txtSJId.SetFocus
    ElseIf objMode = UpdateMode Then
        objMode = ViewMode
        
        SetMode
    End If
End Sub

Private Sub SaveDetailFunction()
    Dim strSJId As String
    
    strSJId = mdlProcedures.RepDupText(Trim(Me.txtSJId.Text))
    
    mdlDatabase.DeleteRecordQuery mdlGlobal.conInventory, mdlTable.CreateTHSTOCK, "ReferencesNumber='" & strSJId & "'"
    mdlDatabase.DeleteRecordQuery mdlGlobal.conInventory, mdlTable.CreateTDSJSELL, "SJId='" & strSJId & "'"

    If Not Me.flxDetail.Rows > 1 Then Exit Sub
    
    Dim rstTemp As ADODB.Recordset
    
    Set rstTemp = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "*", mdlTable.CreateTDSJSELL, , "SJId='" & strSJId & "'")
    
    With rstTemp
        If Not .RecordCount > 0 Then
            Dim lngRow As Long
            
            For lngRow = 1 To Me.flxDetail.Rows - 1
                .AddNew
                
                !SJId = strSJId
                !SJDtlId = !SJId & Trim(Me.flxDetail.TextMatrix(lngRow, ItemIdColumn))
                !ItemId = Trim(Me.flxDetail.TextMatrix(lngRow, ItemIdColumn))
                !WarehouseId = Trim(Me.flxDetail.TextMatrix(lngRow, WarehouseIdColumn))
                !Qty = mdlProcedures.GetCurrency(Trim(Me.flxDetail.TextMatrix(lngRow, QtyColumn)))
                
                !CreateId = Trim(mdlGlobal.UserAuthority.UserId)
                !CreateDate = mdlProcedures.FormatDate(Now)
                
                !UpdateId = Trim(mdlGlobal.UserAuthority.UserId)
                !UpdateDate = mdlProcedures.FormatDate(Now)
                
                .Update
                
                mdlTransaction.UpdateStock _
                    Trim(Me.flxDetail.TextMatrix(lngRow, ItemIdColumn)), _
                    Trim(Me.flxDetail.TextMatrix(lngRow, WarehouseIdColumn)), _
                    strSJId, _
                    Me.dtpSJDate.Value, , _
                    mdlProcedures.GetCurrency(Trim(Me.flxDetail.TextMatrix(lngRow, QtyColumn)))
            Next lngRow
        End If
    End With
    
    mdlDatabase.CloseRecordset rstTemp
End Sub

Private Function CheckValidation() As Boolean
    If Trim(Me.txtSJId.Text) = "" Then
        MsgBox "Nomor SJ Harap Diisi", vbOKOnly + vbExclamation, mdlProcedures.SetCaptionMode(strFormCaption, objMode)
        
        Me.txtSJId.SetFocus
        
        CheckValidation = False
        
        Exit Function
    ElseIf Not mdlProcedures.IsValidComboData(Me.cmbSOId) Then
        MsgBox "Nomor SO Harap Dipilih", vbOKOnly + vbExclamation, mdlProcedures.SetCaptionMode(strFormCaption, objMode)
        
        Me.cmbSOId.SetFocus
        
        CheckValidation = False
        
        Exit Function
    ElseIf Not mdlProcedures.IsValidComboData(Me.cmbDeliveryId) Then
        MsgBox "Alamat Kirim Harap Dipilih", vbOKOnly + vbExclamation, mdlProcedures.SetCaptionMode(strFormCaption, objMode)
        
        Me.cmbDeliveryId.SetFocus
        
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
            mdlDatabase.SearchRecordset rstMain, "SJId", mdlProcedures.RepDupText(Trim(Me.txtSJId.Text))
            
            If Not .EOF Then
                MsgBox "Nomor SJ Sudah Ada", vbOKOnly + vbExclamation, mdlProcedures.SetCaptionMode(strFormCaption, objMode)
                
                Me.txtSJId.SetFocus
                
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
            Me.txtSJId.Text = Trim(!SJId)
            Me.dtpSJDate.Value = mdlProcedures.FormatDate(!SJDate, mdlGlobal.strFormatDate)
            
            Me.CustomerIdCombo = mdlDatabase.GetFieldData(mdlGlobal.conInventory, "CustomerId", mdlTable.CreateTHPOSELL, "POId='" & _
                mdlDatabase.GetFieldData(mdlGlobal.conInventory, "POId", mdlTable.CreateTHSOSELL, "SOId='" & !SOId & "'") & "'")
            
            Me.FillComboTHSOSELL
            
            Me.SOIdCombo = !SOId
            
            Me.txtReferencesNumber.Text = Trim(!ReferencesNumber)
            
            FillInfo
            
            Me.FillComboTMDELIVERYCUSTOMER
            
            mdlProcedures.SetComboData Me.cmbDeliveryId, !DeliveryId
            
            Me.txtNotes.Text = Trim(!Notes)
            
            FillGrid !SJId
        Else
            FillGrid
        End If
    End With
End Sub

Private Sub FillGrid(Optional ByVal strSJId As String = "")
    With Me.flxDetail
        .Rows = 1
        
        Dim rstTemp As ADODB.Recordset
        
        Set rstTemp = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "*", mdlTable.CreateTDSJSELL, False, "SJId='" & strSJId & "'")
        
        While Not rstTemp.EOF
            .Rows = .Rows + 1
            
            .TextMatrix(.Rows - 1, ItemIdColumn) = rstTemp!ItemId
            .TextMatrix(.Rows - 1, NameColumn) = mdlDatabase.GetFieldData(mdlGlobal.conInventory, "Name", mdlTable.CreateTMITEM, "ItemId='" & rstTemp!ItemId & "'")
            .TextMatrix(.Rows - 1, QtyColumn) = mdlProcedures.FormatCurrency(rstTemp!Qty)
            .TextMatrix(.Rows - 1, QtyLastColumn) = mdlProcedures.FormatCurrency(mdlTHSOSELL.GetTotalQtySOSELL(Me.SOIdCombo, rstTemp!ItemId))
            .TextMatrix(.Rows - 1, UnityIdColumn) = mdlDatabase.GetFieldData(mdlGlobal.conInventory, "UnityId", mdlTable.CreateTMITEM, "ItemId='" & rstTemp!ItemId & "'")
            .TextMatrix(.Rows - 1, WarehouseIdColumn) = rstTemp!WarehouseId
            
            .Row = .Rows - 1
            .Col = ItemIdColumn
            .ColSel = WarehouseIdColumn
            
            DoEvents
            
            rstTemp.MoveNext
        Wend
        
        If rstTemp.RecordCount > 0 Then .Row = 1
        
        mdlDatabase.CloseRecordset rstTemp
    End With
End Sub

Private Sub FillCombo()
    FillComboSearch
    
    Me.FillComboTMCUSTOMER
End Sub

Private Sub FillComboSearch()
    mdlProcedures.FillComboMonth Me.cmbMonth, , , mdlProcedures.FormatDate(Now, "M")
    
    Me.txtYear.Text = mdlProcedures.FormatDate(Now, "yyyy")
End Sub

Public Sub FillComboTMCUSTOMER()
    Dim rstTemp As ADODB.Recordset
    
    Set rstTemp = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "CustomerId, Name", mdlTable.CreateTMCUSTOMER, False, "StatusYN='" & mdlGlobal.strYes & "'")
    
    mdlProcedures.FillComboData Me.cmbCustomerId, rstTemp
    
    mdlDatabase.CloseRecordset rstTemp
End Sub

Public Sub FillComboTMDELIVERYCUSTOMER()
    Dim rstTemp As ADODB.Recordset
    
    Set rstTemp = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "DeliveryId, Name", mdlTable.CreateTMDELIVERYCUSTOMER, False, "CustomerId='" & Me.CustomerIdCombo & "'")
    
    mdlProcedures.FillComboData Me.cmbDeliveryId, rstTemp
    
    If Not rstTemp.RecordCount > 0 Then
        If mdlProcedures.IsValidComboData(Me.cmbCustomerId) Then
            Set rstTemp = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "CustomerId, Name", mdlTable.CreateTMCUSTOMER, False, "CustomerId='" & Me.CustomerIdCombo & "'")
            
            mdlProcedures.FillComboData Me.cmbDeliveryId, rstTemp
        End If
    End If
    
    mdlDatabase.CloseRecordset rstTemp
End Sub

Public Sub FillComboTHSOSELL()
    Dim strTableFirst As String
    Dim strTableSecond As String
    Dim strTable As String
    
    strTableFirst = mdlTable.CreateTHSOSELL
    strTableSecond = mdlTable.CreateTHPOSELL
    
    strTable = strTableFirst & " LEFT JOIN " & strTableSecond & " ON " & _
        strTableFirst & ".POId=" & strTableSecond & ".POId"
        
    Dim strCriteria As String
    
    strCriteria = strTableSecond & ".CustomerId='" & Me.CustomerIdCombo & "'"

    Dim rstTemp As ADODB.Recordset
    
    Set rstTemp = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "SOId", strTable, False, strCriteria)
    
    mdlProcedures.FillComboData Me.cmbSOId, rstTemp
    
    mdlDatabase.CloseRecordset rstTemp
End Sub

Private Sub ArrangeGrid()
    With Me.flxDetail
        .AllowUserResizing = flexResizeColumns
        
        .Rows = 1
        .Cols = WarehouseIdColumn + 1
        
        .ColWidth(BlankColumn) = 300
        .ColWidth(ItemIdColumn) = 1200
        .ColWidth(NameColumn) = 3800
        .ColWidth(QtyColumn) = 1500
        .ColWidth(QtyLastColumn) = 1500
        .ColWidth(UnityIdColumn) = 1000
        .ColWidth(WarehouseIdColumn) = 0
        
        .ColAlignment(ItemIdColumn) = flexAlignLeftCenter
        .ColAlignment(NameColumn) = flexAlignLeftCenter
        .ColAlignment(QtyColumn) = flexAlignRightCenter
        .ColAlignment(QtyLastColumn) = flexAlignRightCenter
        .ColAlignment(UnityIdColumn) = flexAlignLeftCenter
        
        .TextMatrix(0, ItemIdColumn) = "Kode"
        .TextMatrix(0, NameColumn) = "Nama"
        .TextMatrix(0, QtyColumn) = "Qty"
        .TextMatrix(0, QtyLastColumn) = "Qty. SO"
        .TextMatrix(0, UnityIdColumn) = "Satuan"
    End With
End Sub

Public Function SaveDetail( _
    ByVal strItemId As String, _
    ByVal curQty As Currency, _
    ByVal strWarehouseId As String) As Boolean
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
            .TextMatrix(.Rows - 1, QtyLastColumn) = mdlProcedures.FormatCurrency(mdlTHSOSELL.GetTotalQtySOSELL(Me.SOIdCombo, strItemId))
            .TextMatrix(.Rows - 1, UnityIdColumn) = mdlDatabase.GetFieldData(mdlGlobal.conInventory, "UnityId", mdlTable.CreateTMITEM, "ItemId='" & strItemId & "'")
            .TextMatrix(.Rows - 1, WarehouseIdColumn) = strWarehouseId
            
            .Row = .Rows - 1
            .Col = ItemIdColumn
            .ColSel = WarehouseIdColumn
            
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
    
    Set rstTemp = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "SJId", mdlTable.CreateTHSJSELL, , "MONTH(SJDate)=" & Me.MonthTrans & " AND YEAR(SJDate)=" & Me.YearTrans)
    
    With rstTemp
        If .RecordCount > 0 Then
            Dim intNoSeq As Integer
            
            While Not .EOF
                intNoSeq = mdlTransaction.ExtractSequential(Trim(!SJId), 4)
                
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
    
    strReferencesNumber = mdlText.strSJIDINIT & "/" & mdlText.strSELLINIT
    
    Set rstTemp = _
        mdlDatabase.OpenRecordset( _
            mdlGlobal.conInventory, _
            "ReferencesNumber", _
            mdlTable.CreateTHRECYCLE, _
            False, _
            "LEFT(ReferencesNumber, " & Len(strReferencesNumber) & ")='" & strReferencesNumber & "'")
            
    While Not rstTemp.EOF
        intNoSeq = mdlTransaction.ExtractSequential(Trim(rstTemp!ReferencesNumber), 4)
        
        If Not intNoSeq < 0 Then
            If intCounter < intNoSeq Then
                intCounter = intNoSeq
            End If
        End If
        
        rstTemp.MoveNext
    Wend
    
    intCounter = intCounter + 1
    
    Me.txtSJId.Text = ZipId(intCounter)
    
    mdlDatabase.CloseRecordset rstTemp
End Sub

Private Function ZipId(ByVal intNoSeq As Integer) As String
    ZipId = _
        mdlText.strSJIDINIT & _
        "/" & _
        mdlText.strSELLINIT & _
        "/" & _
        mdlProcedures.FormatDate(Me.dtpSJDate.Value, "yyyy") & _
        "/" & _
        mdlProcedures.FormatDate(Me.dtpSJDate.Value, "MM") & _
        "/" & _
        mdlProcedures.FormatNumber(intNoSeq, "0000")
End Function

Private Function ZipReferencesNumber(ByVal intNoSeq As Integer) As String
    Dim dteNow As Date
    
    dteNow = mdlProcedures.SetDate(mdlProcedures.GetComboData(Me.cmbMonth), Me.txtYear.Text)
    
    ZipReferencesNumber = _
        mdlText.strTownInit & _
        "/" & _
        mdlProcedures.FormatDate(dteNow, "yy") & _
        "/" & _
        UCase(mdlProcedures.FormatDate(dteNow, "MMM")) & _
        "/" & _
        mdlProcedures.FormatNumber(intNoSeq)
End Function

Private Sub FillInfo(Optional ByVal blnClear As Boolean = False)
    If blnClear Then
        Me.txtPOId.Caption = ""
        Me.txtDateLine.Caption = ""
        
        Exit Sub
    End If
    
    Dim strPOId As String
    Dim strDateLine As String
    
    strPOId = mdlDatabase.GetFieldData(mdlGlobal.conInventory, "POId", mdlTable.CreateTHSOSELL, "SOId='" & Me.SOIdCombo & "'")
    
    Me.txtPOId.Caption = strPOId
    
    strDateLine = mdlDatabase.GetFieldData(mdlGlobal.conInventory, "DateLine", mdlTable.CreateTHPOSELL, "POId='" & strPOId & "'")
    
    If Trim(strDateLine) = "" Then
        Me.txtDateLine.Caption = strDateLine
    Else
        Me.txtDateLine.Caption = mdlProcedures.FormatDate(strDateLine, "dd MMMM yyyy")
    End If
End Sub

Public Property Get Parent() As Boolean
    Parent = blnParent
End Property

Public Property Get MonthTrans() As String
    MonthTrans = mdlProcedures.GetComboData(Me.cmbMonth)
End Property

Public Property Get YearTrans() As String
    YearTrans = Trim(Me.txtYear.Text)
End Property

Public Property Get SJId() As String
    If rstMain Is Nothing Then Exit Property

    If rstMain.RecordCount > 0 Then
        SJId = rstMain!SJId
    End If
End Property

Public Property Get CustomerIdCombo() As String
    CustomerIdCombo = mdlProcedures.GetComboData(Me.cmbCustomerId)
End Property

Public Property Get SOIdCombo() As String
    SOIdCombo = mdlProcedures.GetComboData(Me.cmbSOId)
End Property

Public Property Let Parent(ByVal blnEnable As Boolean)
    blnParent = blnEnable
    
    If Not blnEnable Then mdlProcedures.CenterWindows Me
End Property

Public Property Let SJId(ByVal strSJId As String)
    If rstMain Is Nothing Then Exit Property
    
    mdlDatabase.SearchRecordset rstMain, "SJId", strSJId
    
    If Not rstMain.EOF Then FillText
End Property

Public Property Let CustomerIdCombo(ByVal strCustomerId As String)
    mdlProcedures.SetComboData Me.cmbCustomerId, strCustomerId
End Property

Public Property Let SOIdCombo(ByVal strSOId As String)
    mdlProcedures.SetComboData Me.cmbSOId, strSOId
End Property
