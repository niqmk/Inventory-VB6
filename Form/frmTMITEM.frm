VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmTMITEM 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7950
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10245
   Icon            =   "frmTMITEM.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7950
   ScaleWidth      =   10245
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraHeader 
      Height          =   735
      Left            =   120
      TabIndex        =   29
      Top             =   1200
      Width           =   9975
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
         Left            =   1800
         MaxLength       =   7
         TabIndex        =   0
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lblItemId 
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
         TabIndex        =   12
         Top             =   240
         Width           =   450
      End
   End
   Begin VB.Frame fraDetail 
      Height          =   5895
      Left            =   120
      TabIndex        =   30
      Top             =   1920
      Width           =   9975
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
         TabIndex        =   4
         Top             =   1440
         Width           =   7455
      End
      Begin VB.CommandButton cmdVendorId 
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
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   1440
         Width           =   375
      End
      Begin VB.TextBox txtPartNumber 
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
         MaxLength       =   40
         TabIndex        =   2
         Top             =   720
         Width           =   4935
      End
      Begin VB.CommandButton cmdBrandId 
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
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   2880
         Width           =   375
      End
      Begin VB.CommandButton cmdCategoryId 
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
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   2400
         Width           =   375
      End
      Begin VB.CommandButton cmdGroupId 
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
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   1920
         Width           =   375
      End
      Begin MSComCtl2.DTPicker dtpItemDate 
         Height          =   375
         Left            =   1800
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
         Format          =   47448067
         CurrentDate     =   39330
      End
      Begin VB.Frame fraStock 
         Caption         =   "Informasi Stok"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Left            =   120
         TabIndex        =   31
         Top             =   3240
         Width           =   9735
         Begin VB.CommandButton cmdUnityId 
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
            TabIndex        =   25
            TabStop         =   0   'False
            Top             =   360
            Width           =   375
         End
         Begin VB.TextBox txtMaxStock 
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
            Left            =   1680
            MaxLength       =   10
            TabIndex        =   10
            Top             =   1200
            Width           =   1335
         End
         Begin VB.TextBox txtMinStock 
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
            Left            =   1680
            MaxLength       =   10
            TabIndex        =   9
            Top             =   840
            Width           =   1335
         End
         Begin VB.ComboBox cmbUnityId 
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
            Left            =   1680
            TabIndex        =   8
            Top             =   360
            Width           =   7215
         End
         Begin VB.Label lblMaxStock 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Max. Stok"
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
            TabIndex        =   27
            Top             =   1200
            Width           =   900
         End
         Begin VB.Label lblMinStock 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Min. Stok"
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
            Top             =   840
            Width           =   840
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
            Left            =   120
            TabIndex        =   24
            Top             =   360
            Width           =   615
         End
      End
      Begin VB.ComboBox cmbCategoryId 
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
         Top             =   2400
         Width           =   7215
      End
      Begin VB.ComboBox cmbBrandId 
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
         Top             =   2880
         Width           =   7215
      End
      Begin VB.ComboBox cmbGroupId 
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
         Top             =   1920
         Width           =   7215
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
         Left            =   1800
         MaxLength       =   50
         TabIndex        =   3
         Top             =   1080
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
         Left            =   1800
         MaxLength       =   150
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   11
         Top             =   4920
         Width           =   7215
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
         TabIndex        =   16
         Top             =   1440
         Width           =   825
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
         TabIndex        =   14
         Top             =   720
         Width           =   990
      End
      Begin VB.Label lblItemDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal Masuk"
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
         Top             =   240
         Width           =   1320
      End
      Begin VB.Label lblCategoryId 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Jenis"
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
         TabIndex        =   20
         Top             =   2400
         Width           =   450
      End
      Begin VB.Label lblBrandId 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Merk"
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
         Top             =   2880
         Width           =   435
      End
      Begin VB.Label lblGroupId 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Grup"
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
         TabIndex        =   17
         Top             =   1920
         Width           =   420
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
         TabIndex        =   15
         Top             =   1080
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
         TabIndex        =   28
         Top             =   4920
         Width           =   990
      End
   End
   Begin MSComctlLib.Toolbar tlbMain 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   32
      Top             =   0
      Width           =   10245
      _ExtentX        =   18071
      _ExtentY        =   1111
      ButtonWidth     =   609
      ButtonHeight    =   953
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
               Picture         =   "frmTMITEM.frx":1F8A
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTMITEM.frx":43DC
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTMITEM.frx":682E
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTMITEM.frx":8C80
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTMITEM.frx":B0D2
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTMITEM.frx":B674
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTMITEM.frx":DAC6
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "frmTMITEM"
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

Private PrintMaster As clsPRTTMITEM

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
        
        mdlProcedures.ShowForm frmBRWTMITEM, False, True
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
        
        mdlDatabase.CloseRecordset rstMain
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmTMITEM = Nothing
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
            
            mdlProcedures.ShowForm frmBRWTMITEM, False, True
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
        SendKeys "{TAB}"
    End If
End Sub

Private Sub dtpItemDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtPartNumber_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtName_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub cmbCategoryId_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If mdlProcedures.IsValidComboData(Me.cmbCategoryId) Then
            SendKeys "{TAB}"
        Else
            If blnParent Then Exit Sub
            
            blnParent = True
            
            mdlProcedures.CornerWindows Me
            
            mdlProcedures.ShowForm frmBRWTMCATEGORY, False, True
        End If
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

Private Sub cmbGroupId_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If mdlProcedures.IsValidComboData(Me.cmbGroupId) Then
            SendKeys "{TAB}"
        Else
            If blnParent Then Exit Sub
            
            blnParent = True
            
            mdlProcedures.CornerWindows Me
            
            mdlProcedures.ShowForm frmBRWTMGROUP, False, True
        End If
    End If
End Sub

Private Sub cmbBrandId_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If mdlProcedures.IsValidComboData(Me.cmbBrandId) Then
            SendKeys "{TAB}"
        Else
            If blnParent Then Exit Sub
            
            blnParent = True
            
            mdlProcedures.CornerWindows Me
            
            mdlProcedures.ShowForm frmBRWTMBRAND, False, True
        End If
    End If
End Sub

Private Sub cmbUnityId_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If mdlProcedures.IsValidComboData(Me.cmbUnityId) Then
            SendKeys "{TAB}"
        Else
            If blnParent Then Exit Sub
            
            blnParent = True
            
            mdlProcedures.CornerWindows Me
            
            mdlProcedures.ShowForm frmBRWTMUNITY, False, True
        End If
    End If
End Sub

Private Sub txtMinStock_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtMaxStock_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtItemId_GotFocus()
    mdlProcedures.GotFocus Me.txtItemId
End Sub

Private Sub txtPartNumber_GotFocus()
    mdlProcedures.GotFocus Me.txtPartNumber
End Sub

Private Sub txtName_GotFocus()
    mdlProcedures.GotFocus Me.txtName
End Sub

Private Sub txtMinStock_GotFocus()
    mdlProcedures.GotFocus Me.txtMinStock
End Sub

Private Sub txtMaxStock_GotFocus()
    mdlProcedures.GotFocus Me.txtMaxStock
End Sub

Private Sub txtNotes_GotFocus()
    mdlProcedures.GotFocus Me.txtNotes
End Sub

Private Sub txtName_Validate(Cancel As Boolean)
    If Not objMode = AddMode Then Exit Sub

    If Trim(Me.txtName.Text) = "" Then Exit Sub
    
    If Not Trim(Me.txtItemId.Text) = "" Then
        If Not mdlProcedures.SetMsgYesNo("Apakah Anda Ingin Mengganti Kode Barang ?", Me.Caption) Then Exit Sub
    End If
    
    IncrementId
End Sub

Private Sub txtMinStock_Change()
    Me.txtMinStock.Text = mdlProcedures.FormatCurrency(Me.txtMinStock.Text)
    
    Me.txtMinStock.SelStart = Len(Me.txtMinStock.Text)
End Sub

Private Sub txtMaxStock_Change()
    Me.txtMaxStock.Text = mdlProcedures.FormatCurrency(Me.txtMaxStock.Text)
    
    Me.txtMaxStock.SelStart = Len(Me.txtMaxStock.Text)
End Sub

Private Sub cmdVendorId_Click()
    If Not mdlGlobal.UserAuthority.IsMenuAccess(mdiMain.mnuTMVENDOR.Name) Then Exit Sub

    If blnParent Then Exit Sub
    
    blnParent = True
    
    mdlProcedures.CornerWindows Me
    
    mdlProcedures.ShowForm frmTMVENDOR, False
End Sub

Private Sub cmdGroupId_Click()
    If Not mdlGlobal.UserAuthority.IsMenuAccess(mdiMain.mnuTMGROUP.Name) Then Exit Sub

    If blnParent Then Exit Sub
    
    blnParent = True
    
    mdlProcedures.CornerWindows Me
    
    mdlProcedures.ShowForm frmTMGROUP, False
End Sub

Private Sub cmdCategoryId_Click()
    If Not mdlGlobal.UserAuthority.IsMenuAccess(mdiMain.mnuTMCATEGORY.Name) Then Exit Sub
    
    If blnParent Then Exit Sub
    
    blnParent = True
    
    mdlProcedures.CornerWindows Me
    
    mdlProcedures.ShowForm frmTMCATEGORY, False
End Sub

Private Sub cmdBrandId_Click()
    If Not mdlGlobal.UserAuthority.IsMenuAccess(mdiMain.mnuTMBRAND.Name) Then Exit Sub
    
    If blnParent Then Exit Sub
    
    blnParent = True
    
    mdlProcedures.CornerWindows Me
    
    mdlProcedures.ShowForm frmTMBRAND, False
End Sub

Private Sub cmdUnityId_Click()
    If Not mdlGlobal.UserAuthority.IsMenuAccess(mdiMain.mnuTMUNITY.Name) Then Exit Sub
    
    If blnParent Then Exit Sub
    
    blnParent = True
    
    mdlProcedures.CornerWindows Me
    
    mdlProcedures.ShowForm frmTMUNITY, False
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
    
    Me.dtpItemDate.CustomFormat = mdlGlobal.strFormatDate
    
    FillCombo
    
    strFormCaption = mdlText.strTMITEM
    
    blnParent = False
    blnActivate = False
    
    Set rstMain = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "*", mdlTable.CreateTMITEM, , , "ItemId ASC")
    
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
        
        Me.txtItemId.SetFocus
    ElseIf objMode = UpdateMode Then
        Me.fraHeader.Enabled = False
        
        mdlProcedures.SetControlMode Me, objMode, False, Me.txtItemId.Name
        
        Me.dtpItemDate.SetFocus
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

Private Sub DeleteFunction()
    objMode = DeleteMode
    
    If mdlProcedures.SetMsgYesNo("Apakah Anda Yakin ?", mdlProcedures.SetCaptionMode(strFormCaption, objMode)) Then
        Dim strMessage As String
        
        strMessage = "Tidak Dapat Dihapus" & vbCrLf & "Terdapat Data Yang Masih Terkait ("
        
        If mdlDatabase.IsDataExists(mdlGlobal.conInventory, mdlTable.CreateTMITEMPRICE, "ItemId='" & rstMain!ItemId & "'") Then
            MsgBox strMessage & mdlText.strTMITEMPRICE & ")", vbOKOnly + vbCritical, mdlProcedures.SetCaptionMode(strFormCaption, objMode)
        ElseIf mdlDatabase.IsDataExists(mdlGlobal.conInventory, mdlTable.CreateTMPRICELIST, "ItemId='" & rstMain!ItemId & "'") Then
            MsgBox strMessage & mdlText.strTMPRICELIST & ")", vbOKOnly + vbCritical, mdlProcedures.SetCaptionMode(strFormCaption, objMode)
        ElseIf mdlDatabase.IsDataExists(mdlGlobal.conInventory, mdlTable.CreateTMCONVERTPRICE, "ItemId='" & rstMain!ItemId & "'") Then
            MsgBox strMessage & mdlText.strTMCONVERTPRICE & ")", vbOKOnly + vbCritical, mdlProcedures.SetCaptionMode(strFormCaption, objMode)
        ElseIf mdlDatabase.IsDataExists(mdlGlobal.conInventory, mdlTable.CreateTMSTOCKINIT, "ItemId='" & rstMain!ItemId & "'") Then
            MsgBox strMessage & mdlText.strTMSTOCKINIT & ")", vbOKOnly + vbCritical, mdlProcedures.SetCaptionMode(strFormCaption, objMode)
        ElseIf mdlDatabase.IsDataExists(mdlGlobal.conInventory, mdlTable.CreateTHSTOCK, "ItemId='" & rstMain!ItemId & "'") Then
            MsgBox strMessage & mdlTable.CreateTHSTOCK & ")", vbOKOnly + vbCritical, mdlProcedures.SetCaptionMode(strFormCaption, objMode)
        ElseIf mdlDatabase.IsDataExists(mdlGlobal.conInventory, mdlTable.CreateTDITEMIN, "ItemId='" & rstMain!ItemId & "'") Then
            MsgBox strMessage & mdlText.strTDITEMIN & ")", vbOKOnly + vbCritical, mdlProcedures.SetCaptionMode(strFormCaption, objMode)
        ElseIf mdlDatabase.IsDataExists(mdlGlobal.conInventory, mdlTable.CreateTDITEMOUT, "ItemId='" & rstMain!ItemId & "'") Then
            MsgBox strMessage & mdlText.strTDITEMOUT & ")", vbOKOnly + vbCritical, mdlProcedures.SetCaptionMode(strFormCaption, objMode)
        ElseIf mdlDatabase.IsDataExists(mdlGlobal.conInventory, mdlTable.CreateTDMUTITEM, "ItemId='" & rstMain!ItemId & "'") Then
            MsgBox strMessage & mdlText.strTDMUTITEM & ")", vbOKOnly + vbCritical, mdlProcedures.SetCaptionMode(strFormCaption, objMode)
        ElseIf mdlDatabase.IsDataExists(mdlGlobal.conInventory, mdlTable.CreateTDPOBUY, "ItemId='" & rstMain!ItemId & "'") Then
            MsgBox strMessage & mdlText.strTDPOBUY & ")", vbOKOnly + vbCritical, mdlProcedures.SetCaptionMode(strFormCaption, objMode)
        ElseIf mdlDatabase.IsDataExists(mdlGlobal.conInventory, mdlTable.CreateTDSJBUY, "ItemId='" & rstMain!ItemId & "'") Then
            MsgBox strMessage & mdlText.strTDSJBUY & ")", vbOKOnly + vbCritical, mdlProcedures.SetCaptionMode(strFormCaption, objMode)
        ElseIf mdlDatabase.IsDataExists(mdlGlobal.conInventory, mdlTable.CreateTDRTRBUY, "ItemId='" & rstMain!ItemId & "'") Then
            MsgBox strMessage & mdlText.strTDRTRBUY & ")", vbOKOnly + vbCritical, mdlProcedures.SetCaptionMode(strFormCaption, objMode)
        ElseIf mdlDatabase.IsDataExists(mdlGlobal.conInventory, mdlTable.CreateTDPOSELL, "ItemId='" & rstMain!ItemId & "'") Then
            MsgBox strMessage & mdlText.strTDPOSELL & ")", vbOKOnly + vbCritical, mdlProcedures.SetCaptionMode(strFormCaption, objMode)
        ElseIf mdlDatabase.IsDataExists(mdlGlobal.conInventory, mdlTable.CreateTDSOSELL, "ItemId='" & rstMain!ItemId & "'") Then
            MsgBox strMessage & mdlText.strTDSOSELL & ")", vbOKOnly + vbCritical, mdlProcedures.SetCaptionMode(strFormCaption, objMode)
        ElseIf mdlDatabase.IsDataExists(mdlGlobal.conInventory, mdlTable.CreateTDSJSELL, "ItemId='" & rstMain!ItemId & "'") Then
            MsgBox strMessage & mdlText.strTDSJSELL & ")", vbOKOnly + vbCritical, mdlProcedures.SetCaptionMode(strFormCaption, objMode)
        Else
            mdlDatabase.DeleteSingleRecord rstMain
        End If
    End If
    
    objMode = ViewMode
    
    SetMode
    
    FillText
End Sub

Private Sub PrintFunction()
    If Not PrintMaster Is Nothing Then
        Set PrintMaster = Nothing
    End If

    Set PrintMaster = New clsPRTTMITEM
    
    Dim rstTemp As ADODB.Recordset
    
    Set rstTemp = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "*", mdlTable.CreateTMITEM, False, , "ItemId ASC")
    
    PrintMaster.ImportToExcel rstTemp
    
    mdlDatabase.CloseRecordset rstTemp
End Sub

Private Sub SaveFunction()
    If Not CheckValidation Then Exit Sub
    
    With rstMain
        mdlDatabase.SearchRecordset rstMain, "ItemId", mdlProcedures.RepDupText(Trim(Me.txtItemId.Text))
        
        If .EOF Then
            .AddNew
            
            !ItemId = mdlProcedures.RepDupText(Trim(Me.txtItemId.Text))
            
            !CreateId = Trim(mdlGlobal.UserAuthority.UserId)
            !CreateDate = mdlProcedures.FormatDate(Now)
        End If
        
        !ItemDate = mdlProcedures.FormatDate(Me.dtpItemDate.Value)
        !PartNumber = mdlProcedures.RepDupText(Trim(Me.txtPartNumber.Text))
        !Name = mdlProcedures.RepDupText(Trim(Me.txtName.Text))
        !VendorId = Me.VendorIdCombo
        !GroupId = Me.GroupIdCombo
        !CategoryId = Me.CategoryIdCombo
        !BrandId = Me.BrandIdCombo
        !UnityId = Me.UnityIdCombo
        !MinStock = mdlProcedures.GetCurrency(Me.txtMinStock.Text)
        !MaxStock = mdlProcedures.GetCurrency(Me.txtMaxStock.Text)
        !Notes = mdlProcedures.RepDupText(Trim(Me.txtNotes.Text))
        
        !UpdateId = Trim(mdlGlobal.UserAuthority.UserId)
        !UpdateDate = mdlProcedures.FormatDate(Now)
        
        .Update
    End With
    
    If objMode = AddMode Then
        mdlProcedures.SetControlMode Me, objMode
        
        Me.txtItemId.SetFocus
    ElseIf objMode = UpdateMode Then
        objMode = ViewMode
        
        SetMode
    End If
End Sub

Private Function CheckValidation() As Boolean
    If mdlProcedures.RepDupText(Trim(Me.txtItemId.Text)) = "" Then
        If objMode = AddMode Then
            IncrementId
            
            Exit Function
        Else
            MsgBox "Kode Harap Diisi", vbOKOnly + vbExclamation, mdlProcedures.SetCaptionMode(strFormCaption, objMode)
            
            Me.txtItemId.SetFocus
            
            CheckValidation = False
            
            Exit Function
        End If
    ElseIf mdlProcedures.RepDupText(Trim(Me.txtPartNumber.Text)) = "" Then
        MsgBox "Nomor Part Harap Diisi", vbOKOnly + vbExclamation, mdlProcedures.SetCaptionMode(strFormCaption, objMode)
        
        Me.txtPartNumber.SetFocus
        
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
            mdlDatabase.SearchRecordset rstMain, "ItemId", mdlProcedures.RepDupText(Trim(Me.txtItemId.Text))
            
            If Not .EOF Then
                MsgBox "Kode Sudah Ada", vbOKOnly + vbExclamation, mdlProcedures.SetCaptionMode(strFormCaption, objMode)
                
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
            Me.dtpItemDate.Value = mdlProcedures.FormatDate(!ItemDate, mdlGlobal.strFormatDate)
            Me.txtPartNumber.Text = Trim(!PartNumber)
            Me.txtName.Text = Trim(!Name)
            
            Me.VendorIdCombo = !GroupId
            Me.GroupIdCombo = !GroupId
            Me.CategoryIdCombo = !CategoryId
            Me.BrandIdCombo = !BrandId
            Me.UnityIdCombo = !UnityId
            
            Me.txtMinStock.Text = mdlProcedures.GetCurrency(!MinStock)
            Me.txtMaxStock.Text = mdlProcedures.GetCurrency(!MaxStock)
            Me.txtNotes.Text = Trim(!Notes)
        End If
    End With
End Sub

Private Sub IncrementId()
    Dim rstTemp As ADODB.Recordset
    
    Set rstTemp = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "ItemId", mdlTable.CreateTMITEM, False, "LEFT(ItemId, 1)='" & Left(Trim(Me.txtName.Text), 1) & "'", "ItemId")
    
    With rstTemp
        If .RecordCount > 0 Then
            Dim intCounter As Integer
        
            Dim rstCheck As ADODB.Recordset
            
            Dim intHole As Integer
            
            intHole = 0
            
            For intCounter = 1 To rstTemp.RecordCount
                Set rstCheck = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "ItemId", mdlTable.CreateTMITEM, False, "ItemId='" & Left(Trim(Me.txtName.Text), 1) & mdlProcedures.FormatNumber(intCounter, "000000") & "'")
                
                If Not rstCheck.RecordCount > 0 Then
                    intHole = intCounter
                    
                    Exit For
                End If
            Next intCounter
            
            mdlDatabase.CloseRecordset rstCheck
            
            If intHole = 0 Then intHole = .RecordCount + 1
            
            Me.txtItemId.Text = Left(Trim(Me.txtName.Text), 1) & mdlProcedures.FormatNumber(intHole, "000000")
        Else
            Me.txtItemId.Text = Left(Trim(Me.txtName.Text), 1) & mdlProcedures.FormatNumber(1, "000000")
        End If
    End With
    
    mdlDatabase.CloseRecordset rstTemp
End Sub

Private Sub FillCombo()
    Me.FillComboTMVENDOR
    Me.FillComboTMGROUP
    Me.FillComboTMCATEGORY
    Me.FillComboTMBRAND
    Me.FillComboTMUNITY
End Sub

Public Sub FillComboTMVENDOR()
    Dim rstTemp As ADODB.Recordset
    
    Set rstTemp = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "VendorId, Name", mdlTable.CreateTMVENDOR, False)
    
    mdlProcedures.FillComboData Me.cmbVendorId, rstTemp
    
    mdlDatabase.CloseRecordset rstTemp
End Sub

Public Sub FillComboTMGROUP()
    Dim rstTemp As ADODB.Recordset
    
    Set rstTemp = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "GroupId, Name", mdlTable.CreateTMGROUP, False)
    
    mdlProcedures.FillComboData Me.cmbGroupId, rstTemp
    
    mdlDatabase.CloseRecordset rstTemp
End Sub

Public Sub FillComboTMCATEGORY()
    Dim rstTemp As ADODB.Recordset
    
    Set rstTemp = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "CategoryId, Name", mdlTable.CreateTMCATEGORY, False)
    
    mdlProcedures.FillComboData Me.cmbCategoryId, rstTemp
    
    mdlDatabase.CloseRecordset rstTemp
End Sub

Public Sub FillComboTMBRAND()
    Dim rstTemp As ADODB.Recordset
    
    Set rstTemp = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "BrandId, Name", mdlTable.CreateTMBRAND, False)
    
    mdlProcedures.FillComboData Me.cmbBrandId, rstTemp
    
    mdlDatabase.CloseRecordset rstTemp
End Sub

Public Sub FillComboTMUNITY()
    Dim rstTemp As ADODB.Recordset
    
    Set rstTemp = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "UnityId, Name", mdlTable.CreateTMUNITY, False)
    
    mdlProcedures.FillComboData Me.cmbUnityId, rstTemp
    
    mdlDatabase.CloseRecordset rstTemp
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

Public Property Get PartNumber() As String
    If rstMain Is Nothing Then Exit Property
    
    If rstMain.RecordCount > 0 Then
        PartNumber = rstMain!PartNumber
    End If
End Property

Public Property Get ItemName() As String
    If rstMain Is Nothing Then Exit Property

    If rstMain.RecordCount > 0 Then
        ItemName = rstMain!Name
    End If
End Property

Public Property Get VendorIdCombo() As String
    VendorIdCombo = mdlProcedures.GetComboData(Me.cmbVendorId)
End Property

Public Property Get GroupIdCombo() As String
    GroupIdCombo = mdlProcedures.GetComboData(Me.cmbGroupId)
End Property

Public Property Get CategoryIdCombo() As String
    CategoryIdCombo = mdlProcedures.GetComboData(Me.cmbCategoryId)
End Property

Public Property Get BrandIdCombo() As String
    BrandIdCombo = mdlProcedures.GetComboData(Me.cmbBrandId)
End Property

Public Property Get UnityIdCombo() As String
    UnityIdCombo = mdlProcedures.GetComboData(Me.cmbUnityId)
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

Public Property Let VendorIdCombo(ByVal strVendorId As String)
    mdlProcedures.SetComboData Me.cmbVendorId, strVendorId
End Property

Public Property Let GroupIdCombo(ByVal strGroupId As String)
    mdlProcedures.SetComboData Me.cmbGroupId, strGroupId
End Property

Public Property Let CategoryIdCombo(ByVal strCategoryId As String)
    mdlProcedures.SetComboData Me.cmbCategoryId, strCategoryId
End Property

Public Property Let BrandIdCombo(ByVal strBrandId As String)
    mdlProcedures.SetComboData Me.cmbBrandId, strBrandId
End Property

Public Property Let UnityIdCombo(ByVal strUnityId As String)
    mdlProcedures.SetComboData Me.cmbUnityId, strUnityId
End Property
