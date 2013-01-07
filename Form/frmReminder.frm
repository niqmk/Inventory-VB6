VERSION 5.00
Begin VB.Form frmReminder 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3255
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   3390
   Icon            =   "frmReminder.frx":0000
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   3390
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
      Left            =   2040
      TabIndex        =   4
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Frame fraMain 
      Height          =   2535
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   3135
      Begin VB.OptionButton optReminder 
         Caption         =   "Setiap Membuka Transaksi"
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
         Index           =   3
         Left            =   120
         TabIndex        =   3
         Top             =   2040
         Width           =   2895
      End
      Begin VB.TextBox txtMinute 
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
         Left            =   960
         MaxLength       =   2
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   1680
         Width           =   375
      End
      Begin VB.TextBox txtHour 
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
         Left            =   480
         MaxLength       =   2
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   1680
         Width           =   375
      End
      Begin VB.OptionButton optReminder 
         Caption         =   "Setiap Jam"
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
         Index           =   2
         Left            =   120
         TabIndex        =   2
         Top             =   1200
         Width           =   2895
      End
      Begin VB.OptionButton optReminder 
         Caption         =   "Setiap Membuka Master"
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
         Index           =   1
         Left            =   120
         TabIndex        =   1
         Top             =   720
         Width           =   2895
      End
      Begin VB.OptionButton optReminder 
         Caption         =   "Setiap Program Startup"
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
         Index           =   0
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   2895
      End
   End
End
Attribute VB_Name = "frmReminder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private objReminderLoadTypeMain As ReminderLoadType

Private Sub Form_Load()
    SetInitialization
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmReminder = Nothing
End Sub

Private Sub optReminder_Click(Index As Integer)
    objReminderLoadTypeMain = Index
End Sub

Private Sub txtHour_Change()
    If Not Me.optReminder(ReminderLoadType.ClockType).Value Then
        Me.optReminder(ReminderLoadType.ClockType).Value = True
    End If
End Sub

Private Sub txtMinute_Change()
    If Not Me.optReminder(ReminderLoadType.ClockType).Value Then
        Me.optReminder(ReminderLoadType.ClockType).Value = True
    End If
End Sub

Private Sub txtHour_GotFocus()
    mdlProcedures.GotFocus Me.txtHour
End Sub

Private Sub txtMinute_GotFocus()
    mdlProcedures.GotFocus Me.txtMinute
End Sub

Private Sub txtHour_Validate(Cancel As Boolean)
    If Not IsNumeric(Me.txtHour.Text) Then Me.txtHour.Text = "0"
    
    If CInt(Me.txtHour.Text) > 24 Or CInt(Me.txtHour.Text) < 1 Then
        Me.txtHour.Text = "0"
    End If
    
    Me.txtHour.Text = mdlProcedures.FormatNumber(CInt(Me.txtHour.Text), "00")
End Sub

Private Sub txtMinute_Validate(Cancel As Boolean)
    If Not IsNumeric(Me.txtMinute.Text) Then Me.txtMinute.Text = "1"
    
    If CInt(Me.txtMinute.Text) > 59 Or CInt(Me.txtMinute.Text) < 0 Then
        Me.txtMinute.Text = "1"
    End If
    
    Me.txtMinute.Text = mdlProcedures.FormatNumber(CInt(Me.txtMinute.Text), "00")
End Sub

Private Sub cmdSave_Click()
    SaveFunction
    
    If mdiMain.Reminder Then mdiMain.Reminder = False
End Sub

Private Sub SetInitialization()
    mdlProcedures.CenterWindows Me
    
    Me.Caption = mdlText.strReminder
    
    Me.optReminder(mdlGlobal.objReminderLoadType).Value = True
    
    If mdlGlobal.objReminderLoadType = ClockType Then
        Dim strValue() As String
        
        strValue = Split(mdlGlobal.strReminderTimeText, ":")
        
        Me.txtHour.Text = strValue(0)
        Me.txtMinute.Text = strValue(1)
    End If
End Sub

Private Sub SaveFunction()
    If Not CheckValidation Then Exit Sub
    
    mdlGlobal.objReminderLoadType = objReminderLoadTypeMain
    mdlGlobal.strReminderTimeText = Me.txtHour.Text & ":" & Me.txtMinute.Text
    
    SetRegistry
    
    Unload Me
End Sub

Private Function CheckValidation() As Boolean
    If Not CheckReminderLoadType Then
        MsgBox "Jadwal Pengingat Kosong", vbCritical + vbOKOnly, Me.Caption
        
        CheckValidation = False
        
        Exit Function
    End If
    
    If Me.optReminder(ReminderLoadType.ClockType).Value Then
        If Trim(Me.txtHour.Text) = "" Then
            MsgBox "Jam Kosong", vbCritical + vbOKOnly, Me.Caption
            
            Me.txtHour.SetFocus
            
            CheckValidation = False
            
            Exit Function
        ElseIf Trim(Me.txtMinute.Text) = "" Then
            MsgBox "Menit Kosong", vbCritical + vbOKOnly, Me.Caption
            
            Me.txtMinute.SetFocus
            
            CheckValidation = False
            
            Exit Function
        End If
    End If
    
    CheckValidation = True
End Function

Private Function CheckReminderLoadType() As Boolean
    CheckReminderLoadType = False

    Dim intCounter As Integer
    
    For intCounter = 0 To Me.optReminder.Count - 1
        If Me.optReminder(intCounter).Value Then
            CheckReminderLoadType = True
            
            Exit For
        End If
    Next intCounter
End Function

Private Sub SetRegistry()
    Dim lngRegistry As Long
    Dim lngRegKey As Long
    Dim lngType As Long
    Dim lngSize As Long
    
    On Local Error GoTo ErrHandler

    lngRegistry = _
        mdlRegistry.OpenRegistry(mdlRegistry.HKEY_CURRENT_USER, mdlRegistry.KEYS_SYS_INFO, lngRegKey)
        
    If Not lngRegistry = 0 Then
        lngRegistry = mdlRegistry.WriteToRegistry(mdlRegistry.HKEY_CURRENT_USER, mdlRegistry.KEYS_SYS_INFO)
    End If
    
    lngRegistry = _
            mdlRegistry.WriteValueRegistry(mdlRegistry.HKEY_CURRENT_USER, mdlRegistry.KEYS_SYS_INFO, mdlGlobal.REMINDER_REGISTRY, CStr(mdlGlobal.objReminderLoadType))
            
    lngRegistry = _
        mdlRegistry.WriteValueRegistry(mdlRegistry.HKEY_CURRENT_USER, mdlRegistry.KEYS_SYS_INFO, mdlGlobal.REMINDER_VALUE_REGISTRY, Me.txtHour.Text & ":" & Me.txtMinute.Text)

    lngRegistry = mdlRegistry.CloseRegistry(lngRegKey)
    
    Exit Sub

ErrHandler:
    MsgBox "Jadwal Pengingat Tidak Dapat Disimpan", vbCritical + vbOKOnly, Me.Caption
End Sub
