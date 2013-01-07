VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmSyncFinance 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5775
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   9045
   Icon            =   "frmSyncFinance.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5775
   ScaleWidth      =   9045
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
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
      Left            =   7800
      TabIndex        =   6
      Top             =   5280
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
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
      Left            =   6600
      TabIndex        =   4
      Top             =   5280
      Width           =   1095
   End
   Begin VB.CommandButton cmdImport 
      Caption         =   "<-"
      Height          =   375
      Left            =   4320
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   2400
      Width           =   375
   End
   Begin VB.CommandButton cmdExport 
      Caption         =   "->"
      Height          =   375
      Left            =   4320
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1800
      Width           =   375
   End
   Begin MSComctlLib.ListView lsvInventory 
      Height          =   3855
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   6800
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Frame fraMain 
      Height          =   735
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   8775
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
         Left            =   7560
         TabIndex        =   1
         Top             =   240
         Width           =   1095
      End
      Begin VB.ComboBox cmbTable 
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
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   240
         Width           =   5775
      End
      Begin VB.Label lblTable 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Tabel"
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
         Width           =   1035
      End
   End
   Begin MSComctlLib.ListView lsvFinance 
      Height          =   3855
      Left            =   4800
      TabIndex        =   3
      Top             =   1320
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   6800
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Label lblFinance 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Finance"
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
      TabIndex        =   8
      Top             =   960
      Width           =   690
   End
   Begin VB.Label lblInventory 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Inventory"
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
      Top             =   960
      Width           =   765
   End
End
Attribute VB_Name = "frmSyncFinance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum TableMode
    EmployeeMode
End Enum

Private objTableMode As TableMode

Private clsFinance As clsFinance

Private Sub Form_Load()
    SetInitialization
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If mdlGlobal.blnFill Then
        Cancel = 1
    Else
        Set clsFinance = Nothing
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmSyncFinance = Nothing
End Sub

Private Sub cmdView_Click()
    If mdlGlobal.blnFill Then Exit Sub
    
    mdlGlobal.blnFill = True
    
    ViewInventory
    ViewFinance
    
    mdlGlobal.blnFill = False
End Sub

Private Sub cmdExport_Click()
    If mdlGlobal.blnFill Then Exit Sub
    
    mdlGlobal.blnFill = True
    
    Dim lstItem As ListItem
    
    Dim intCounter As Integer
    
    Dim strValue As String
    Dim strSubValue As String
    
    For intCounter = 1 To Me.lsvInventory.ListItems.Count
        If Me.lsvInventory.ListItems(intCounter).Selected Then
            strValue = Me.lsvInventory.ListItems(intCounter).Text
            strSubValue = Me.lsvInventory.ListItems(intCounter).ListSubItems(1).Text
            
            If Not mdlProcedures.IsDataExistsInListView(Me.lsvFinance, strValue) Then
                Set lstItem = Me.lsvFinance.ListItems.Add(, , strValue)
                lstItem.ListSubItems.Add , , strSubValue
            End If
        End If
    Next intCounter
    
    mdlGlobal.blnFill = False
End Sub

Private Sub cmdImport_Click()
    If mdlGlobal.blnFill Then Exit Sub
    
    mdlGlobal.blnFill = True
    
    Dim lstItem As ListItem
    
    Dim intCounter As Integer
    
    Dim strValue As String
    Dim strSubValue As String
    
    For intCounter = 1 To Me.lsvFinance.ListItems.Count
        If Me.lsvFinance.ListItems(intCounter).Selected Then
            strValue = Me.lsvFinance.ListItems(intCounter).Text
            strSubValue = Me.lsvFinance.ListItems(intCounter).ListSubItems(1).Text
            
            If Not mdlProcedures.IsDataExistsInListView(Me.lsvInventory, strValue) Then
                Set lstItem = Me.lsvInventory.ListItems.Add(, , strValue)
                lstItem.ListSubItems.Add , , strSubValue
            End If
        End If
    Next intCounter
    
    mdlGlobal.blnFill = False
End Sub

Private Sub cmdApply_Click()
    If mdlGlobal.blnFill Then Exit Sub
    
    mdlGlobal.blnFill = True
    
    SaveInventory
    SaveFinance
    
    mdlGlobal.blnFill = False
End Sub

Private Sub cmdOK_Click()
    If mdlGlobal.blnFill Then Exit Sub
    
    mdlGlobal.blnFill = True
    
    SaveInventory
    SaveFinance
    
    mdlGlobal.blnFill = False
    
    Unload Me
End Sub

Private Sub SetInitialization()
    mdlProcedures.CenterWindows Me
    
    Me.Caption = mdlText.strSyncFinance
    
    With Me.lsvInventory
        .LabelEdit = lvwManual
        .View = lvwReport
        .MultiSelect = True
        .FullRowSelect = True
        
        .ColumnHeaders.Add
        .ColumnHeaders.Add
    End With
    
    With Me.lsvFinance
        .LabelEdit = lvwManual
        .View = lvwReport
        .MultiSelect = True
        .FullRowSelect = True
        
        .ColumnHeaders.Add
        .ColumnHeaders.Add
    End With
    
    FillCombo
    
    mdlGlobal.blnFill = True
    
    ViewInventory
    ViewFinance
    
    mdlGlobal.blnFill = False
End Sub

Private Sub FillCombo()
    Set clsFinance = New clsFinance

    With Me.cmbTable
        .AddItem clsFinance.CreateTMEMPLOYEE(mdlGlobal.conFinance)
    
        If .ListCount > 0 Then
            .ListIndex = 0
            
            objTableMode = EmployeeMode
        End If
    End With
End Sub

Private Sub ViewInventory()
    If Not mdlProcedures.IsValidComboData(Me.cmbTable) Then Exit Sub
    
    Dim rstTemp As ADODB.Recordset
    
    Select Case mdlProcedures.GetComboData(Me.cmbTable)
        Case mdlTable.CreateTMEMPLOYEE:
            Set rstTemp = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "*", mdlProcedures.GetComboData(Me.cmbTable), False, , "EmployeeId ASC")
            
            FillViewInventoryTMEMPLOYEE rstTemp
            
            objTableMode = EmployeeMode
    End Select
    
    mdlDatabase.CloseRecordset rstTemp
End Sub

Private Sub SaveInventory()
    Select Case objTableMode
        Case EmployeeMode:
            SaveInventoryTMEMPLOYEE
    End Select
End Sub

Private Sub ViewFinance()
    If Not mdlProcedures.IsValidComboData(Me.cmbTable) Then Exit Sub
    
    Dim rstTemp As ADODB.Recordset
    
    Select Case mdlProcedures.GetComboData(Me.cmbTable)
        Case clsFinance.CreateTMEMPLOYEE(mdlGlobal.conFinance):
            Set rstTemp = mdlDatabase.OpenRecordset(mdlGlobal.conFinance, "*", mdlProcedures.GetComboData(Me.cmbTable), False, , "EmployeeId ASC")
            
            FillViewFinanceTMEMPLOYEE rstTemp
    End Select
    
    mdlDatabase.CloseRecordset rstTemp
End Sub

Private Sub SaveFinance()
    Select Case objTableMode
        Case EmployeeMode:
            SaveFinanceTMEMPLOYEE
    End Select
End Sub

Private Sub FillViewInventoryTMEMPLOYEE(ByRef rstTemp As ADODB.Recordset)
    With Me.lsvInventory
        .ListItems.Clear
        .ColumnHeaders.Clear
        
        .ColumnHeaders.Add , , "Kode", 900
        .ColumnHeaders.Add , , "Nama", 2600
    End With
    
    Dim lstItem As ListItem
    
    With rstTemp
        While Not .EOF
            Set lstItem = Me.lsvInventory.ListItems.Add(, , rstTemp!EmployeeId)
            lstItem.ListSubItems.Add , , rstTemp!Name
            
            DoEvents
        
            .MoveNext
        Wend
        
        If .RecordCount > 0 Then
            Me.lsvInventory.ListItems(1).Selected = True
        End If
    End With
End Sub

Private Sub SaveInventoryTMEMPLOYEE()
    Dim rstTemp As ADODB.Recordset
    
    Dim intCounter As Integer
    
    Dim strEmployeeId As String
    
    If Me.lsvInventory.ListItems.Count > 0 Then
        For intCounter = 1 To Me.lsvInventory.ListItems.Count
            strEmployeeId = Me.lsvInventory.ListItems(intCounter).Text
            
            Set rstTemp = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "*", mdlTable.CreateTMEMPLOYEE, , "EmployeeId='" & strEmployeeId & "'")
            
            With rstTemp
                If Not .RecordCount > 0 Then
                    .AddNew
                    
                    !EmployeeId = strEmployeeId
                    !EmployeeDate = mdlProcedures.FormatDate(mdlDatabase.GetFieldData(mdlGlobal.conFinance, "EmployeeDate", clsFinance.CreateTMEMPLOYEE(mdlGlobal.conFinance), "EmployeeId='" & strEmployeeId & "'"))
                    !Name = mdlDatabase.GetFieldData(mdlGlobal.conFinance, "Name", clsFinance.CreateTMEMPLOYEE(mdlGlobal.conFinance), "EmployeeId='" & strEmployeeId & "'")
                    !Address = mdlDatabase.GetFieldData(mdlGlobal.conFinance, "Address", clsFinance.CreateTMEMPLOYEE(mdlGlobal.conFinance), "EmployeeId='" & strEmployeeId & "'")
                    !Phone = mdlDatabase.GetFieldData(mdlGlobal.conFinance, "Phone", clsFinance.CreateTMEMPLOYEE(mdlGlobal.conFinance), "EmployeeId='" & strEmployeeId & "'")
                    !HandPhone = mdlDatabase.GetFieldData(mdlGlobal.conFinance, "HandPhone", clsFinance.CreateTMEMPLOYEE(mdlGlobal.conFinance), "EmployeeId='" & strEmployeeId & "'")
                    !Fax = mdlDatabase.GetFieldData(mdlGlobal.conFinance, "Fax", clsFinance.CreateTMEMPLOYEE(mdlGlobal.conFinance), "EmployeeId='" & strEmployeeId & "'")
                    !Email = mdlDatabase.GetFieldData(mdlGlobal.conFinance, "Email", clsFinance.CreateTMEMPLOYEE(mdlGlobal.conFinance), "EmployeeId='" & strEmployeeId & "'")
                    !Notes = mdlDatabase.GetFieldData(mdlGlobal.conFinance, "Notes", clsFinance.CreateTMEMPLOYEE(mdlGlobal.conFinance), "EmployeeId='" & strEmployeeId & "'")
                    !CreateId = mdlGlobal.UserAuthority.UserId
                    !CreateDate = mdlProcedures.FormatDate(Now)
                    !UpdateId = mdlGlobal.UserAuthority.UserId
                    !UpdateDate = mdlProcedures.FormatDate(Now)
                    
                    .Update
                End If
            End With
        Next intCounter
    End If
    
    mdlDatabase.CloseRecordset rstTemp
End Sub

Private Sub FillViewFinanceTMEMPLOYEE(ByRef rstTemp As ADODB.Recordset)
    With Me.lsvFinance
        .ListItems.Clear
        .ColumnHeaders.Clear
        
        .ColumnHeaders.Add , , "Kode", 900
        .ColumnHeaders.Add , , "Nama", 2600
    End With
    
    Dim lstItem As ListItem
    
    With rstTemp
        While Not .EOF
            Set lstItem = Me.lsvFinance.ListItems.Add(, , !EmployeeId)
            lstItem.ListSubItems.Add , , !Name
            
            DoEvents
        
            .MoveNext
        Wend
        
        If .RecordCount > 0 Then
            Me.lsvInventory.ListItems(1).Selected = True
        End If
    End With
End Sub

Private Sub SaveFinanceTMEMPLOYEE()
    Dim rstTemp As ADODB.Recordset
    
    Dim intCounter As Integer
    
    Dim strEmployeeId As String
    
    If Me.lsvFinance.ListItems.Count > 0 Then
        For intCounter = 1 To Me.lsvFinance.ListItems.Count
            strEmployeeId = Me.lsvFinance.ListItems(intCounter).Text
            
            Set rstTemp = mdlDatabase.OpenRecordset(mdlGlobal.conFinance, "*", clsFinance.CreateTMEMPLOYEE(mdlGlobal.conFinance), , "EmployeeId='" & strEmployeeId & "'")
            
            With rstTemp
                If Not .RecordCount > 0 Then
                    .AddNew
                    
                    !EmployeeId = strEmployeeId
                    !EmployeeDate = mdlProcedures.FormatDate(mdlDatabase.GetFieldData(mdlGlobal.conInventory, "EmployeeDate", mdlTable.CreateTMEMPLOYEE, "EmployeeId='" & strEmployeeId & "'"))
                    !Name = mdlDatabase.GetFieldData(mdlGlobal.conInventory, "Name", mdlTable.CreateTMEMPLOYEE, "EmployeeId='" & strEmployeeId & "'")
                    !Address = mdlDatabase.GetFieldData(mdlGlobal.conInventory, "Address", mdlTable.CreateTMEMPLOYEE, "EmployeeId='" & strEmployeeId & "'")
                    !Phone = mdlDatabase.GetFieldData(mdlGlobal.conInventory, "Phone", mdlTable.CreateTMEMPLOYEE, "EmployeeId='" & strEmployeeId & "'")
                    !HandPhone = mdlDatabase.GetFieldData(mdlGlobal.conInventory, "HandPhone", mdlTable.CreateTMEMPLOYEE, "EmployeeId='" & strEmployeeId & "'")
                    !Fax = mdlDatabase.GetFieldData(mdlGlobal.conInventory, "Fax", mdlTable.CreateTMEMPLOYEE, "EmployeeId='" & strEmployeeId & "'")
                    !Email = mdlDatabase.GetFieldData(mdlGlobal.conInventory, "Email", mdlTable.CreateTMEMPLOYEE, "EmployeeId='" & strEmployeeId & "'")
                    !Notes = mdlDatabase.GetFieldData(mdlGlobal.conInventory, "Notes", mdlTable.CreateTMEMPLOYEE, "EmployeeId='" & strEmployeeId & "'")
                    !CreateId = mdlGlobal.UserAuthority.UserId
                    !CreateDate = mdlProcedures.FormatDate(Now)
                    !UpdateId = mdlGlobal.UserAuthority.UserId
                    !UpdateDate = mdlProcedures.FormatDate(Now)
                    
                    .Update
                End If
            End With
        Next intCounter
    End If
    
    mdlDatabase.CloseRecordset rstTemp
End Sub
