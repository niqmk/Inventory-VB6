VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmTMREMINDERCUSTOMER 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6510
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   10830
   Icon            =   "frmTMREMINDERCUSTOMER.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6510
   ScaleWidth      =   10830
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdSave 
      Caption         =   "Simpan"
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
      Left            =   9600
      TabIndex        =   5
      Top             =   6000
      Width           =   1095
   End
   Begin VB.Frame fraMain 
      Height          =   5775
      Left            =   7440
      TabIndex        =   25
      Top             =   120
      Width           =   3255
      Begin VB.Frame fraValidate 
         Caption         =   "Waktu Pengingat"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   120
         TabIndex        =   24
         Top             =   3960
         Width           =   3015
         Begin VB.TextBox txtValidate 
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
            Index           =   1
            Left            =   1320
            MaxLength       =   2
            TabIndex        =   4
            Top             =   1320
            Width           =   375
         End
         Begin VB.OptionButton optValidate 
            Caption         =   "Setiap"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   2
            Left            =   240
            TabIndex        =   20
            TabStop         =   0   'False
            Top             =   1320
            Width           =   975
         End
         Begin VB.TextBox txtValidate 
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
            Index           =   0
            Left            =   1320
            MaxLength       =   2
            TabIndex        =   3
            Top             =   840
            Width           =   375
         End
         Begin VB.OptionButton optValidate 
            Caption         =   "Setiap"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   1
            Left            =   240
            TabIndex        =   19
            TabStop         =   0   'False
            Top             =   840
            Width           =   975
         End
         Begin VB.OptionButton optValidate 
            Caption         =   "Setiap Sebulan Sekali"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   0
            Left            =   240
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   360
            Width           =   2535
         End
         Begin VB.Label lblValidate 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Hari"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   1
            Left            =   1920
            TabIndex        =   9
            Top             =   1320
            Width           =   420
         End
         Begin VB.Label lblValidate 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Bulan"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   0
            Left            =   1920
            TabIndex        =   8
            Top             =   840
            Width           =   525
         End
      End
      Begin VB.Frame fraReminderType 
         Caption         =   "Jenis Pengingat"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3615
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Width           =   3015
         Begin VB.OptionButton optReminderType 
            Caption         =   "Dari Transaksi Customer Terakhir"
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
            Index           =   4
            Left            =   240
            TabIndex        =   17
            TabStop         =   0   'False
            Top             =   3000
            Width           =   2295
         End
         Begin VB.OptionButton optReminderType 
            Caption         =   "Dari Tanggal Pertama"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   3
            Left            =   240
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   2520
            Width           =   2295
         End
         Begin MSComCtl2.DTPicker dtpReminderFromDate 
            Height          =   375
            Left            =   480
            TabIndex        =   21
            TabStop         =   0   'False
            Top             =   1920
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
            Format          =   47448065
            CurrentDate     =   39387
         End
         Begin VB.OptionButton optReminderType 
            Caption         =   "Dari Tanggal"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   2
            Left            =   240
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   1560
            Width           =   1455
         End
         Begin VB.OptionButton optReminderType 
            Caption         =   "Dari Tanggal Customer"
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
            Index           =   1
            Left            =   240
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   960
            Width           =   2295
         End
         Begin VB.OptionButton optReminderType 
            Caption         =   "Tidak Terpilih"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   0
            Left            =   240
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   480
            Width           =   1455
         End
      End
   End
   Begin VB.Frame fraSearch 
      Height          =   1455
      Left            =   120
      TabIndex        =   22
      Top             =   120
      Width           =   7215
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
         Left            =   6000
         TabIndex        =   2
         Top             =   960
         Width           =   1095
      End
      Begin VB.CheckBox chkReminder 
         Caption         =   "Semua"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   2400
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   960
         Width           =   975
      End
      Begin VB.CheckBox chkReminder 
         Caption         =   "Terdapat Pengingat"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   960
         Width           =   2055
      End
      Begin VB.TextBox txtCustomerId 
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
         Left            =   960
         TabIndex        =   0
         Top             =   240
         Width           =   855
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
         Left            =   960
         TabIndex        =   1
         Top             =   600
         Width           =   6135
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
         TabIndex        =   6
         Top             =   240
         Width           =   450
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
         Top             =   600
         Width           =   510
      End
   End
   Begin MSDataGridLib.DataGrid dgdMain 
      Height          =   4215
      Left            =   120
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   1680
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   7435
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
End
Attribute VB_Name = "frmTMREMINDERCUSTOMER"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private rstMain As ADODB.Recordset

Private blnLoad As Boolean

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub Form_Load()
    SetInitialization
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If frmTMCUSTOMER.Parent Then
        frmTMCUSTOMER.Parent = False
    End If
    
    mdlDatabase.CloseRecordset rstMain
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmTMREMINDERCUSTOMER = Nothing
End Sub

Private Sub dgdMain_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If rstMain.RecordCount > 0 Then
        FillDetail rstMain!CustomerId
    End If
End Sub

Private Sub optReminderType_Click(Index As Integer)
    Me.dtpReminderFromDate.Value = Now
    
    Select Case Index
        Case ReminderType.NoneType:
            Me.dtpReminderFromDate.Enabled = False
            
            SetValidateOption
            
            Me.fraValidate.Enabled = False
        Case FromMaster:
            Me.dtpReminderFromDate.Enabled = False
            
            Me.fraValidate.Enabled = True
        Case ReminderType.FromDate:
            Me.dtpReminderFromDate.Enabled = True
            
            Me.fraValidate.Enabled = True
        Case ReminderType.FromFirstDate:
            Me.dtpReminderFromDate.Enabled = False
            
            Me.fraValidate.Enabled = True
        Case ReminderType.FromTransaction:
            Me.dtpReminderFromDate.Enabled = False
            
            Me.fraValidate.Enabled = True
    End Select
End Sub

Private Sub cmdSearch_Click()
    SetGrid
End Sub

Private Sub optValidate_Click(Index As Integer)
    Select Case Index
        Case ValidateType.OnceMonth - 1:
            Me.txtValidate(0).Text = ""
            Me.txtValidate(1).Text = ""
            
            Me.txtValidate(0).Enabled = False
            Me.txtValidate(1).Enabled = False
        Case ValidateType.MonthSequence - 1:
            Me.txtValidate(0).Enabled = True
            Me.txtValidate(1).Enabled = False
            
            Me.txtValidate(1).Text = ""
            
            If Trim(Me.txtValidate(0).Text) = "" Then Me.txtValidate(0).Text = "1"
            
            If Not blnLoad Then Me.txtValidate(0).SetFocus
        Case ValidateType.DaySequence - 1:
            Me.txtValidate(0).Enabled = False
            Me.txtValidate(1).Enabled = True
            
            Me.txtValidate(0).Text = ""
            
            If Trim(Me.txtValidate(1).Text) = "" Then Me.txtValidate(1).Text = "1"
            
            If Not blnLoad Then Me.txtValidate(1).SetFocus
    End Select
End Sub

Private Sub txtCustomerId_GotFocus()
    mdlProcedures.GotFocus Me.txtCustomerId
End Sub

Private Sub txtName_GotFocus()
    mdlProcedures.GotFocus Me.txtName
End Sub

Private Sub txtValidate_GotFocus(Index As Integer)
    mdlProcedures.GotFocus Me.txtValidate(Index)
End Sub

Private Sub txtValidate_Validate(Index As Integer, Cancel As Boolean)
    If Not IsNumeric(Me.txtValidate(Index).Text) Then Me.txtValidate(Index).Text = "1"
    
    If CInt(Me.txtValidate(Index).Text) <= 0 Then Me.txtValidate(Index).Text = "1"
End Sub

Private Sub cmdSave_Click()
    If Not CheckValidation Then Exit Sub
    
    SaveFunction
    
    MsgBox "Data Sudah Disimpan", vbOKOnly + vbInformation, Me.Caption
End Sub

Private Sub SetInitialization()
    blnLoad = True
    
    mdlProcedures.CenterWindows Me
    
    Me.Caption = mdlText.strTMREMINDERCUSTOMER
    
    If frmTMCUSTOMER.Parent Then
        Me.txtCustomerId.Text = Trim(frmTMCUSTOMER.CustomerId)
        Me.txtName.Text = Trim(frmTMCUSTOMER.CustomerName)
    End If
    
    Me.chkReminder(1).Value = vbChecked
    
    SetReminderInitialize
    
    SetGrid
    
    blnLoad = False
End Sub

Private Sub SetReminderInitialize()
    Dim rstTemp As ADODB.Recordset
    
    Set rstTemp = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "*", mdlTable.CreateTMCUSTOMER, False)
    
    With rstTemp
        Dim rstReminder As ADODB.Recordset
        
        While Not .EOF
            Set rstReminder = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "*", mdlTable.CreateTMREMINDERCUSTOMER, , "CustomerId='" & rstTemp!CustomerId & "'")
            
            If Not rstReminder.RecordCount > 0 Then
                rstReminder.AddNew
                
                rstReminder!CustomerId = !CustomerId
                rstReminder!ReminderType = ReminderType.NoneType
                rstReminder!ReminderDate = mdlProcedures.FormatDate(Now)
                rstReminder!ValidateType = ValidateType.NoneValidate
                rstReminder!ValidateDate = mdlProcedures.FormatDate(Now)
                rstReminder!CreateId = Trim(mdlGlobal.UserAuthority.UserId)
                rstReminder!CreateDate = mdlProcedures.FormatDate(Now)
                rstReminder!UpdateId = Trim(mdlGlobal.UserAuthority.UserId)
                rstReminder!UpdateDate = mdlProcedures.FormatDate(Now)
                
                rstReminder.Update
            End If
        
            .MoveNext
        Wend
        
        mdlDatabase.CloseRecordset rstReminder
    End With
    
    mdlDatabase.CloseRecordset rstTemp
End Sub

Private Sub SetGrid()
    Dim strTableFirst As String
    Dim strTableSecond As String
    Dim strTable As String
    
    strTableFirst = mdlTable.CreateTMCUSTOMER
    strTableSecond = mdlTable.CreateTMREMINDERCUSTOMER
    
    strTable = strTableFirst & " LEFT JOIN " & strTableSecond & " ON " & _
        strTableFirst & ".CustomerId=" & strTableSecond & ".CustomerId"
    
    Dim strCriteria As String
    
    strCriteria = ""
    
    If Not Trim(Me.txtCustomerId.Text) = "" Then
        strCriteria = mdlProcedures.QueryLikeCriteria(strTableFirst & ".CustomerId", mdlProcedures.RepDupText(Me.txtCustomerId.Text))
    End If
    
    If Not Trim(Me.txtName.Text) = "" Then
        If Not Trim(strCriteria) = "" Then
            strCriteria = strCriteria & " AND "
        End If
        
        strCriteria = strCriteria & mdlProcedures.QueryLikeCriteria("Name", mdlProcedures.RepDupText(Me.txtName.Text))
    End If
    
    If Not Me.chkReminder(1).Value = vbChecked Then
        If Not Trim(strCriteria) = "" Then
            strCriteria = strCriteria & " AND "
        End If
        
        If Me.chkReminder(0).Value = vbChecked Then
            strCriteria = strCriteria & "ReminderType<>'" & ReminderType.NoneType & "'"
        Else
            strCriteria = strCriteria & "ReminderType='" & ReminderType.NoneType & "'"
        End If
    End If
    
    Set rstMain = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, strTableFirst & ".CustomerId, Name", strTable, False, strCriteria, strTableFirst & ".CustomerId ASC")
    
    If rstMain.RecordCount > 0 Then
        Me.fraMain.Enabled = True
        
        FillDetail rstMain!CustomerId
    Else
        Me.fraMain.Enabled = False
    End If
    
    Set Me.dgdMain.DataSource = rstMain
    
    With Me.dgdMain
        .RowHeight = 500
        
        .Columns(0).Width = 800
        .Columns(0).Locked = True
        .Columns(0).Caption = "Kode"
        .Columns(1).Width = 5700
        .Columns(1).Locked = True
        .Columns(1).Caption = "Nama"
    End With
End Sub

Private Sub FillDetail(Optional ByVal strCustomerId As String = "")
    Dim rstTemp As ADODB.Recordset
    
    Set rstTemp = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "*", mdlTable.CreateTMREMINDERCUSTOMER, False, "CustomerId='" & strCustomerId & "'")
    
    With rstTemp
        If .RecordCount > 0 Then
            If IsNumeric(Trim(!ReminderType)) Then
                If Not CInt(Trim(!ReminderType)) >= ReminderType.FromTransaction And Not CInt(Trim(!ReminderType)) <= ReminderType.NoneType Then
                    Me.optReminderType(CInt(Trim(!ReminderType))).Value = True
                    
                    If CInt(Trim(!ReminderType)) = ReminderType.FromDate Then
                        Me.dtpReminderFromDate.Value = mdlProcedures.FormatDate(!ReminderDate, "MM/dd/yyyy")
                    End If
                Else
                    Me.optReminderType(0).Value = True
                End If
            Else
                Me.optReminderType(0).Value = True
            End If
            
            If Not Me.optReminderType(0).Value Then
                If IsNumeric(Trim(!ValidateType)) Then
                    If CInt(Trim(!ValidateType)) <= ValidateType.DaySequence And Not CInt(Trim(!ValidateType)) <= ValidateType.OnceMonth Then
                        Me.optValidate(CInt(Trim(!ValidateType)) - 1).Value = True
                        
                        Select Case CInt(Trim(!ValidateType))
                            Case ValidateType.MonthSequence:
                                Me.txtValidate(0).Text = DateDiff("M", !ReminderDate, !ValidateDate)
                            Case ValidateType.DaySequence:
                                Me.txtValidate(1).Text = DateDiff("d", !ReminderDate, !ValidateDate)
                        End Select
                    Else
                        Me.optValidate(0).Value = True
                    End If
                End If
            End If
        Else
            Me.optReminderType(0).Value = True
        End If
    End With
    
    mdlDatabase.CloseRecordset rstTemp
End Sub

Private Function CheckValidation() As Boolean
    If Not CheckReminderType Then
        CheckValidation = False
        
        Exit Function
    End If
    
    If Not Me.optReminderType(ReminderType.NoneType).Value Then
        If CheckValidateType = NoneValidate Then
            MsgBox "Waktu Pengingat Belum Terisi", vbCritical + vbOKOnly, Me.Caption
            
            CheckValidation = False
            
            Exit Function
        ElseIf Me.optValidate(ValidateType.MonthSequence - 1).Value Then
            If Not IsNumeric(Me.txtValidate(0).Text) Then
                Me.txtValidate(0).SetFocus
                
                CheckValidation = False
                
                Exit Function
            End If
            
            If CInt(Me.txtValidate(0).Text) <= 0 Then Me.txtValidate(0).Text = "1"
        ElseIf Me.optValidate(ValidateType.DaySequence - 1).Value Then
            If Not IsNumeric(Me.txtValidate(1).Text) Then
                Me.txtValidate(1).SetFocus
                
                CheckValidation = False
                
                Exit Function
            End If
            
            If CInt(Me.txtValidate(1).Text) <= 0 Then Me.txtValidate(1).Text = "1"
        End If
    End If
    
    CheckValidation = True
End Function

Private Sub SetValidateOption(Optional ByVal objValidate As ValidateType = ValidateType.NoneValidate)
    Dim intCounter As Integer
    
    For intCounter = 0 To Me.optValidate.Count - 1
        If intCounter + 1 = objValidate Then
            Me.optValidate(intCounter).Value = True
        Else
            Me.optValidate(intCounter).Value = False
            
            If intCounter = ValidateType.MonthSequence - 1 Then
                Me.txtValidate(0).Text = ""
            ElseIf intCounter = ValidateType.DaySequence - 1 Then
                Me.txtValidate(1).Text = ""
            End If
        End If
    Next intCounter
End Sub

Private Sub SaveFunction()
    Dim rstTemp As ADODB.Recordset
    
    Set rstTemp = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "*", mdlTable.CreateTMREMINDERCUSTOMER, , "CustomerId='" & rstMain!CustomerId & "'")
    
    With rstTemp
        If Not .RecordCount > 0 Then
            .AddNew
            
            !CustomerId = rstMain!CustomerId
            
            !CreateId = Trim(mdlGlobal.UserAuthority.UserId)
            !CreateDate = mdlProcedures.FormatDate(Now)
        End If
        
        If Me.optReminderType(ReminderType.NoneType).Value Then
            !ReminderType = ReminderType.NoneType
            
            If IsNull(!ReminderDate) Then !ReminderDate = mdlProcedures.FormatDate(Now)
            
            !ValidateType = ValidateType.NoneValidate
            
            If IsNull(!ValidateDate) Then !ValidateDate = mdlProcedures.FormatDate(Now)
        ElseIf Me.optReminderType(ReminderType.FromMaster).Value Then
            !ReminderType = ReminderType.FromMaster
            
            !ReminderDate = mdlProcedures.FormatDate(mdlDatabase.GetFieldData(mdlGlobal.conInventory, "CustomerDate", mdlTable.CreateTMCUSTOMER, "CustomerId='" & rstMain!CustomerId & "'"))
            
            !ValidateType = CheckValidateType
            
            !ValidateDate = mdlProcedures.FormatDate(CheckValidateDate(!ReminderDate, mdlDatabase.GetFieldData(mdlGlobal.conInventory, "CustomerDate", mdlTable.CreateTMCUSTOMER, "CustomerId='" & rstMain!CustomerId & "'")))
        ElseIf Me.optReminderType(ReminderType.FromDate).Value Then
            !ReminderType = ReminderType.FromDate
            
            !ValidateType = CheckValidateType
            
            !ValidateDate = mdlProcedures.FormatDate(CheckValidateDate(Me.dtpReminderFromDate.Value, Me.dtpReminderFromDate.Value))
            
            If CheckValidateType = OnceMonth Then
                !ReminderDate = _
                    mdlProcedures.FormatDate(mdlProcedures.DateAddFormat( _
                        mdlProcedures.FormatDate(Now), !ValidateDate, "M", -1, False))
            ElseIf CheckValidateType = MonthSequence Then
                !ReminderDate = _
                    mdlProcedures.FormatDate(mdlProcedures.DateAddFormat( _
                        mdlProcedures.FormatDate(Now), !ValidateDate, "M", -CDbl(Me.txtValidate(0).Text), False))
            ElseIf CheckValidateType = DaySequence Then
                !ReminderDate = _
                    mdlProcedures.FormatDate(mdlProcedures.DateAddFormat( _
                        mdlProcedures.FormatDate(Now), !ValidateDate, "d", -CDbl(Me.txtValidate(1).Text), False))
            End If
        ElseIf Me.optReminderType(ReminderType.FromFirstDate).Value Then
            !ReminderType = ReminderType.FromMaster
            
            !ReminderDate = mdlProcedures.FormatDate(mdlProcedures.FormatDate(Now, "yyyy/MM") & "/01")
            
            !ValidateType = CheckValidateType
            
            !ValidateDate = mdlProcedures.FormatDate(CheckValidateDate(!ReminderDate, !ReminderDate))
        ElseIf Me.optReminderType(ReminderType.FromTransaction).Value Then
            Dim dteTransaction As Date
            
            If mdlDatabase.IsDataExists(mdlGlobal.conInventory, mdlTable.CreateTHPOSELL, "CustomerId='" & rstMain!CustomerId & "'") Then
                dteTransaction = mdlProcedures.FormatDate(mdlDatabase.GetFieldData(mdlGlobal.conInventory, "PODate", mdlTable.CreateTHPOSELL, "CustomerId='" & rstMain!CustomerId & "'", "PODate DESC"))
            Else
                dteTransaction = mdlProcedures.FormatDate(Now)
            End If
            
            !ReminderType = ReminderType.FromTransaction
            
            !ReminderDate = mdlProcedures.FormatDate(dteTransaction)
            
            !ValidateDate = mdlProcedures.FormatDate(CheckValidateDate(dteTransaction, dteTransaction))
        End If
        
        !UpdateId = Trim(mdlGlobal.UserAuthority.UserId)
        !UpdateDate = mdlProcedures.FormatDate(Now)
        
        .Update
    End With
    
    mdlDatabase.CloseRecordset rstTemp
End Sub

Private Function CheckReminderType() As Boolean
    CheckReminderType = False
    
    Dim intCounter As Integer
    
    For intCounter = 0 To Me.optReminderType.Count - 1
        If Me.optReminderType(intCounter).Value Then
            CheckReminderType = True
            
            Exit For
        End If
    Next intCounter
End Function

Private Function CheckValidateType() As ValidateType
    Dim intCounter As Integer
    
    Dim objValidateType As ValidateType
    
    objValidateType = NoneValidate
    
    For intCounter = 0 To Me.optValidate.Count - 1
        If Me.optValidate(intCounter).Value Then
            objValidateType = intCounter + 1
        End If
    Next intCounter
    
    CheckValidateType = objValidateType
End Function

Private Function CheckValidateDate(ByVal objCompareDate As Date, ByVal objDate As Date) As Date
    If Me.optValidate(ValidateType.OnceMonth - 1).Value Then
        CheckValidateDate = mdlProcedures.DateAddFormat(objCompareDate, objDate, "M", 1)
    ElseIf Me.optValidate(ValidateType.MonthSequence - 1).Value Then
        CheckValidateDate = mdlProcedures.DateAddFormat(objCompareDate, objDate, "M", CDbl(Me.txtValidate(0)))
    ElseIf Me.optValidate(ValidateType.DaySequence - 1).Value Then
        CheckValidateDate = mdlProcedures.DateAddFormat(objCompareDate, objDate, "d", CDbl(Me.txtValidate(1).Text))
    End If
End Function
