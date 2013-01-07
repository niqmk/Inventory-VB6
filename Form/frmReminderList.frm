VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmReminderList 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6375
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   7950
   Icon            =   "frmReminderList.frx":0000
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6375
   ScaleWidth      =   7950
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraMain 
      Height          =   5655
      Left            =   2400
      TabIndex        =   27
      Top             =   120
      Width           =   5415
      Begin VB.Frame fraPOSELL 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5415
         Left            =   240
         TabIndex        =   29
         Top             =   120
         Width           =   4935
         Begin VB.Label lblQtySJ 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Qty SJ"
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
            Left            =   2760
            TabIndex        =   23
            Top             =   2640
            Width           =   600
         End
         Begin VB.Label txtQtySJ 
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
            Left            =   2760
            TabIndex        =   24
            Top             =   2880
            Width           =   2055
         End
         Begin VB.Label txtQtyPO 
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
            Left            =   120
            TabIndex        =   22
            Top             =   2880
            Width           =   2055
         End
         Begin VB.Label lblQtyPO 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Qty PO"
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
            Top             =   2640
            Width           =   660
         End
         Begin VB.Label txtPONotes 
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
            Height          =   1335
            Left            =   120
            TabIndex        =   26
            Top             =   3480
            Width           =   4695
         End
         Begin VB.Label lblPONotes 
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
            TabIndex        =   25
            Top             =   3240
            Width           =   990
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
            ForeColor       =   &H000000C0&
            Height          =   255
            Left            =   120
            TabIndex        =   20
            Top             =   2280
            Width           =   3255
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
            Left            =   120
            TabIndex        =   19
            Top             =   2040
            Width           =   1845
         End
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
            Left            =   120
            TabIndex        =   18
            Top             =   1680
            Width           =   4695
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
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   1080
            Width           =   3255
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
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   480
            Width           =   2535
         End
         Begin VB.Label lblCustomerId 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nama Customer"
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
            Top             =   1440
            Width           =   1410
         End
         Begin VB.Label lblPODate 
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
            TabIndex        =   15
            Top             =   840
            Width           =   675
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
            TabIndex        =   13
            Top             =   240
            Width           =   915
         End
      End
      Begin VB.Frame fraCustomer 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5415
         Left            =   240
         TabIndex        =   28
         Top             =   120
         Width           =   4935
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
            TabIndex        =   3
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
            TabIndex        =   5
            Top             =   840
            Width           =   615
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
            TabIndex        =   7
            Top             =   2520
            Width           =   330
         End
         Begin VB.Label lblCustomerNotes 
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
            Top             =   3720
            Width           =   990
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
            Height          =   255
            Left            =   120
            TabIndex        =   4
            Top             =   480
            Width           =   4695
         End
         Begin VB.Label txtAddress 
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
            Height          =   1335
            Left            =   120
            TabIndex        =   6
            Top             =   1080
            Width           =   4695
         End
         Begin VB.Label txtFax 
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
            Left            =   120
            TabIndex        =   8
            Top             =   2760
            Width           =   4695
         End
         Begin VB.Label txtCustomerNotes 
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
            Height          =   1335
            Left            =   120
            TabIndex        =   12
            Top             =   3960
            Width           =   4695
         End
         Begin VB.Label txtPhone 
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
            Left            =   120
            TabIndex        =   10
            Top             =   3360
            Width           =   4695
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
            TabIndex        =   9
            Top             =   3120
            Width           =   675
         End
      End
   End
   Begin VB.CommandButton cmdViewTransaction 
      Caption         =   "Lihat Transaksi"
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
      Left            =   6240
      TabIndex        =   2
      Top             =   5880
      Width           =   1575
   End
   Begin MSComctlLib.ImageList imlMain 
      Left            =   120
      Top             =   5760
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
            Picture         =   "frmReminderList.frx":1F8A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReminderList.frx":3ADC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReminderList.frx":562E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdView 
      Caption         =   "Lihat"
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
      Left            =   5040
      TabIndex        =   1
      Top             =   5880
      Width           =   1095
   End
   Begin MSComctlLib.TreeView trvMain 
      Height          =   5655
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   9975
      _Version        =   393217
      Style           =   7
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmReminderList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum TreeNodeList
    CustomerNode = 1
    POSELLNode
End Enum

Private strCustomerId As String
Private strCustomerName As String
Private strPOId As String
Private strPOCustomer As String

Private dtePODate As Date

Private blnParent As Boolean

Private Sub Form_Load()
    SetInitialization
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If blnParent Then
        Cancel = 1
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmReminderList = Nothing
End Sub

Private Sub trvMain_NodeClick(ByVal Node As MSComctlLib.Node)
    If Not Node.Parent Is Nothing Then
        If Node.Parent = mdlText.strCUSTOMERREMINDER Then
            strCustomerId = Node.Text
            
            FillInfoCustomer Node.Text
            
            Me.fraCustomer.Visible = True
            Me.fraPOSELL.Visible = False
        ElseIf Node.Parent = mdlText.strPOSELLREMINDER Then
            strPOId = Node.Text
                
            FillInfoPOSELL Node.Text
            
            Me.fraCustomer.Visible = False
            Me.fraPOSELL.Visible = True
        End If
    Else
        If Node.FullPath = mdlText.strCUSTOMERREMINDER Then
            strCustomerId = ""
            
            FillInfoCustomer
        ElseIf Node.FullPath = mdlText.strPOSELLREMINDER Then
            strPOId = ""
            
            FillInfoPOSELL
        End If
    End If
End Sub

Private Sub cmdView_Click()
    If blnParent Then Exit Sub
    If Trim(strCustomerId) = "" And Trim(strPOId) = "" Then Exit Sub
    
    blnParent = True
    
    mdlProcedures.CornerWindows Me
    
    If Not Trim(strCustomerId) = "" Then
        mdlProcedures.ShowForm frmMISTMCUSTOMER, False
    ElseIf Not Trim(strPOId) = "" Then
        mdlProcedures.ShowForm frmMISTHPOSELL, False
    End If
End Sub

Private Sub cmdViewTransaction_Click()
    If blnParent Then Exit Sub
    If Trim(strCustomerId) = "" Then Exit Sub
    
    blnParent = True
    
    mdlProcedures.CornerWindows Me
    
    mdlProcedures.ShowForm frmMISTMCUSTOMERTRANS, False
End Sub

Private Sub SetInitialization()
    mdlProcedures.CenterWindows Me
    
    Me.Caption = mdlText.strReminderList
    
    With Me.trvMain
        .FullRowSelect = True
        .LabelEdit = tvwManual
        .Style = tvwTreelinesPlusMinusPictureText
        
        .ImageList = Me.imlMain
    End With
    
    Me.fraCustomer.Visible = False
    Me.fraPOSELL.Visible = False
    
    SetTree
End Sub

Private Sub SetTree()
    Dim treNode As Node
    
    With Me.trvMain
        Set treNode = .Nodes.Add(, , , mdlText.strCUSTOMERREMINDER, CustomerNode)
        FillCustomer treNode
        
        Set treNode = .Nodes.Add(, , , mdlText.strPOSELLREMINDER, POSELLNode)
        FillPOSELL treNode
    End With
End Sub

Private Sub FillCustomer(ByRef treNode As Node)
    Dim rstTemp As ADODB.Recordset
    
    If mdlGlobal.objDatabaseInit = SQLSERVER7 Or _
        mdlGlobal.objDatabaseInit = SQLSERVER2000 Or _
        mdlGlobal.objDatabaseInit = SQLEXPRESS Then
        Set rstTemp = _
            mdlDatabase.OpenRecordset( _
                mdlGlobal.conInventory, _
                "*", _
                mdlTable.CreateTMREMINDERCUSTOMER, _
                True, _
                "ValidateType<>'" & ValidateType.NoneValidate & "' AND ValidateDate<='" & mdlProcedures.FormatDate(Now) & "'")
    ElseIf mdlGlobal.objDatabaseInit = MSACCESS Then
        Set rstTemp = _
            mdlDatabase.OpenRecordset( _
                mdlGlobal.conInventory, _
                "*", _
                mdlTable.CreateTMREMINDERCUSTOMER, _
                True, _
                "ValidateType<>'" & ValidateType.NoneValidate & "' AND ValidateDate<=#" & mdlProcedures.FormatDate(Now) & "#")
    ElseIf mdlGlobal.objDatabaseInit = MYSQL Then
        Set rstTemp = _
            mdlDatabase.OpenRecordset( _
                mdlGlobal.conInventory, _
                "*", _
                mdlTable.CreateTMREMINDERCUSTOMER, _
                True, _
                "ValidateType<>'" & ValidateType.NoneValidate & "' AND ValidateDate<='" & mdlProcedures.FormatDate(Now) & "'")
    End If
    
    Dim intDifferent As Integer
    
    With rstTemp
        While Not .EOF
            Me.trvMain.Nodes.Add treNode, tvwChild, , Trim(!CustomerId), 3
            
            Select Case CInt(Trim(!ValidateType))
                Case ValidateType.OnceMonth:
                    !ValidateDate = _
                        mdlProcedures.FormatDate(mdlProcedures.DateAddFormat( _
                            mdlProcedures.FormatDate(Now), !ValidateDate, "M", 1))
                    
                    !ReminderDate = _
                        mdlProcedures.FormatDate(mdlProcedures.DateAddFormat( _
                            mdlProcedures.FormatDate(Now), !ValidateDate, "M", -1, False))
                Case ValidateType.MonthSequence:
                    intDifferent = DateDiff("m", !ReminderDate, !ValidateDate)
                    
                    If intDifferent = 0 Then intDifferent = 1
                
                    !ValidateDate = _
                        mdlProcedures.FormatDate(mdlProcedures.DateAddFormat( _
                            mdlProcedures.FormatDate(Now), !ValidateDate, "M", intDifferent))
                    
                    !ReminderDate = _
                        mdlProcedures.FormatDate(mdlProcedures.DateAddFormat( _
                            mdlProcedures.FormatDate(Now), !ValidateDate, "M", -intDifferent, False))
                Case ValidateType.DaySequence:
                    intDifferent = DateDiff("d", !ReminderDate, !ValidateDate)
                    
                    If intDifferent <= 0 Then intDifferent = 1
                    
                    !ValidateDate = _
                        mdlProcedures.FormatDate(mdlProcedures.DateAddFormat( _
                            mdlProcedures.FormatDate(Now), !ValidateDate, "d", intDifferent))
                    
                    !ReminderDate = _
                        mdlProcedures.FormatDate(mdlProcedures.DateAddFormat( _
                            mdlProcedures.FormatDate(Now), !ValidateDate, "d", -intDifferent, False))
            End Select
            
            .Update
            
            .MoveNext
        Wend
    End With
    
    mdlDatabase.CloseRecordset rstTemp
    
    mdiMain.CheckReminder
End Sub

Private Sub FillPOSELL(ByRef treNode As Node)
    Dim rstTemp As ADODB.Recordset
    
    If mdlGlobal.objDatabaseInit = SQLSERVER7 Or _
        mdlGlobal.objDatabaseInit = SQLSERVER2000 Or _
        mdlGlobal.objDatabaseInit = SQLEXPRESS Then
        Set rstTemp = _
            mdlDatabase.OpenRecordset( _
                mdlGlobal.conInventory, _
                "*", _
                mdlTable.CreateTHPOSELL, _
                False, _
                "DateLine<='" & mdlProcedures.FormatDate(Now) & "'")
    ElseIf mdlGlobal.objDatabaseInit = MSACCESS Then
        Set rstTemp = _
            mdlDatabase.OpenRecordset( _
                mdlGlobal.conInventory, _
                "*", _
                mdlTable.CreateTHPOSELL, _
                False, _
                "DateLine<=#" & mdlProcedures.FormatDate(Now) & "#")
    ElseIf mdlGlobal.objDatabaseInit = MYSQL Then
        Set rstTemp = _
            mdlDatabase.OpenRecordset( _
                mdlGlobal.conInventory, _
                "*", _
                mdlTable.CreateTHPOSELL, _
                False, _
                "DateLine<='" & mdlProcedures.FormatDate(Now) & "'")
    End If
    
    Dim curQtyPO As Currency
    Dim curQtySJ As Currency
    
    With rstTemp
        While Not .EOF
            curQtyPO = mdlTHPOSELL.GetTotalQtyPOSELL(!POId)
            curQtySJ = mdlTHSJSELL.GetQtyPOFromSJSELL(!POId)
            
            If curQtySJ < curQtyPO Then
                Me.trvMain.Nodes.Add treNode, tvwChild, , Trim(!POId), 3
            End If
            
            .MoveNext
        Wend
    End With
    
    mdlDatabase.CloseRecordset rstTemp
End Sub

Private Sub FillInfoCustomer(Optional ByVal mCustomerId As String = "")
    Dim rstTemp As ADODB.Recordset
    
    Set rstTemp = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "*", mdlTable.CreateTMCUSTOMER, False, "CustomerId='" & mCustomerId & "'")
    
    With rstTemp
        If .RecordCount > 0 Then
            Me.txtName.Caption = !Name
            Me.txtAddress.Caption = !Address
            Me.txtFax.Caption = !Fax
            Me.txtPhone.Caption = !Phone
            Me.txtCustomerNotes.Caption = !Notes
        Else
            Me.txtName.Caption = ""
            Me.txtAddress.Caption = ""
            Me.txtFax.Caption = ""
            Me.txtPhone.Caption = ""
            Me.txtCustomerNotes.Caption = ""
        End If
    End With
    
    strCustomerName = Trim(Me.txtName.Caption)
    
    mdlDatabase.CloseRecordset rstTemp
End Sub

Private Sub FillInfoPOSELL(Optional ByVal mPOId As String = "")
    Dim rstTemp As ADODB.Recordset
    
    Set rstTemp = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "*", mdlTable.CreateTHPOSELL, False, "POId='" & mPOId & "'")
    
    With rstTemp
        If .RecordCount > 0 Then
            Me.txtPOId.Caption = !POId
            Me.txtPODate.Caption = mdlProcedures.FormatDate(!PODate, "dd-MMMM-yyyy")
            Me.txtCustomerId.Caption = !CustomerId
            Me.txtDateLine.Caption = mdlProcedures.FormatDate(!DateLine, "dd-MMMM-yyyy")
            Me.txtQtyPO.Caption = mdlProcedures.FormatCurrency(mdlTHPOSELL.GetTotalQtyPOSELL(!POId))
            Me.txtQtySJ.Caption = mdlProcedures.FormatCurrency(mdlTHSJSELL.GetQtyPOFromSJSELL(!POId))
            Me.txtPONotes.Caption = !Notes
            
            dtePODate = mdlProcedures.FormatDate(!PODate, mdlGlobal.strFormatDate)
        Else
            Me.txtPOId.Caption = ""
            Me.txtPODate.Caption = ""
            Me.txtCustomerId.Caption = ""
            Me.txtDateLine.Caption = ""
            Me.txtQtyPO.Caption = ""
            Me.txtQtySJ.Caption = ""
            Me.txtPONotes.Caption = ""
            
            dtePODate = mdlProcedures.FormatDate(Now, mdlGlobal.strFormatDate)
        End If
    End With
    
    strPOCustomer = Trim(Me.txtCustomerId.Caption)
    
    mdlDatabase.CloseRecordset rstTemp
End Sub

Public Property Get Parent() As Boolean
    Parent = blnParent
End Property

Public Property Get CustomerId() As String
    CustomerId = strCustomerId
End Property

Public Property Get CustomerName() As String
    CustomerName = strCustomerName
End Property

Public Property Get POId() As String
    POId = strPOId
End Property

Public Property Get PODate() As Date
    PODate = dtePODate
End Property

Public Property Get POCustomer() As String
    POCustomer = strPOCustomer
End Property

Public Property Let Parent(ByVal blnEnable As Boolean)
    blnParent = blnEnable
    
    If Not blnEnable Then mdlProcedures.CenterWindows Me
End Property
