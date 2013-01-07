VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmSyncAccounting 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5775
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   9045
   Icon            =   "frmSyncAccounting.frx":0000
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
   Begin MSComctlLib.ListView lsvAccounting 
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
   Begin VB.Label lblAccounting 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Accounting"
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
      Width           =   975
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
Attribute VB_Name = "frmSyncAccounting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum TableMode
    ItemMode
    CustomerMode
    VendorMode
End Enum

Private objTableMode As TableMode

Private clsAccounting As clsAccounting

Private Sub Form_Load()
    SetInitialization
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If mdlGlobal.blnFill Then
        Cancel = 1
    Else
        Set clsAccounting = Nothing
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmSyncAccounting = Nothing
End Sub

Private Sub cmdView_Click()
    If mdlGlobal.blnFill Then Exit Sub
    
    mdlGlobal.blnFill = True
    
    ViewInventory
    ViewAccounting
    
    mdlGlobal.blnFill = False
End Sub

Private Sub lsvInventory_DblClick()
    ExportFunction
End Sub

Private Sub lsvAccounting_DblClick()
    ImportFunction
End Sub

Private Sub cmdExport_Click()
    ExportFunction
End Sub

Private Sub cmdImport_Click()
    ImportFunction
End Sub

Private Sub cmdApply_Click()
    If mdlGlobal.blnFill Then Exit Sub
    
    mdlGlobal.blnFill = True
    
    SaveInventory
    SaveAccounting
    
    mdlGlobal.blnFill = False
End Sub

Private Sub cmdOK_Click()
    If mdlGlobal.blnFill Then Exit Sub
    
    mdlGlobal.blnFill = True
    
    SaveInventory
    SaveAccounting
    
    mdlGlobal.blnFill = False
    
    Unload Me
End Sub

Private Sub SetInitialization()
    mdlProcedures.CenterWindows Me
    
    Me.Caption = mdlText.strSyncAccounting
    
    With Me.lsvInventory
        .LabelEdit = lvwManual
        .View = lvwReport
        .MultiSelect = True
        .FullRowSelect = True
        
        .ColumnHeaders.Add
        .ColumnHeaders.Add
    End With
    
    With Me.lsvAccounting
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
    ViewAccounting
    
    mdlGlobal.blnFill = False
End Sub

Private Sub FillCombo()
    Set clsAccounting = New clsAccounting

    With Me.cmbTable
        .AddItem clsAccounting.CreateTMITEM(mdlGlobal.conAccounting)
        .AddItem clsAccounting.CreateTMCUSTOMER(mdlGlobal.conAccounting)
        .AddItem clsAccounting.CreateTMVENDOR(mdlGlobal.conAccounting)
    
        If .ListCount > 0 Then
            .ListIndex = 0
            
            objTableMode = ItemMode
        End If
    End With
End Sub

Private Sub ExportFunction()
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
            
            If Not mdlProcedures.IsDataExistsInListView(Me.lsvAccounting, strValue) Then
                Set lstItem = Me.lsvAccounting.ListItems.Add(, , strValue)
                lstItem.ListSubItems.Add , , strSubValue
            End If
        End If
    Next intCounter
    
    mdlGlobal.blnFill = False
End Sub

Private Sub ImportFunction()
    If mdlGlobal.blnFill Then Exit Sub
    
    mdlGlobal.blnFill = True
    
    Dim lstItem As ListItem
    
    Dim intCounter As Integer
    
    Dim strValue As String
    Dim strSubValue As String
    
    For intCounter = 1 To Me.lsvAccounting.ListItems.Count
        If Me.lsvAccounting.ListItems(intCounter).Selected Then
            strValue = Me.lsvAccounting.ListItems(intCounter).Text
            strSubValue = Me.lsvAccounting.ListItems(intCounter).ListSubItems(1).Text
            
            If Not mdlProcedures.IsDataExistsInListView(Me.lsvInventory, strValue) Then
                Set lstItem = Me.lsvInventory.ListItems.Add(, , strValue)
                lstItem.ListSubItems.Add , , strSubValue
            End If
        End If
    Next intCounter
    
    mdlGlobal.blnFill = False
End Sub

Private Sub ViewInventory()
    If Not mdlProcedures.IsValidComboData(Me.cmbTable) Then Exit Sub
    
    Dim rstTemp As ADODB.Recordset
    
    Select Case mdlProcedures.GetComboData(Me.cmbTable)
        Case mdlTable.CreateTMITEM:
            Set rstTemp = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "*", mdlProcedures.GetComboData(Me.cmbTable), False, , "ItemId ASC")
            
            FillViewInventoryTMITEM rstTemp
            
            objTableMode = ItemMode
        Case mdlTable.CreateTMCUSTOMER:
            Set rstTemp = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "*", mdlProcedures.GetComboData(Me.cmbTable), False, , "CustomerId ASC")
            
            FillViewInventoryTMCUSTOMER rstTemp
            
            objTableMode = CustomerMode
        Case mdlTable.CreateTMVENDOR:
            Set rstTemp = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "*", mdlProcedures.GetComboData(Me.cmbTable), False, , "VendorId ASC")
            
            FillViewInventoryTMVENDOR rstTemp
            
            objTableMode = VendorMode
    End Select
    
    mdlDatabase.CloseRecordset rstTemp
End Sub

Private Sub SaveInventory()
    Select Case objTableMode
        Case ItemMode:
            SaveInventoryTMITEM
        Case CustomerMode:
            SaveInventoryTMCUSTOMER
        Case VendorMode:
            SaveInventoryTMVENDOR
    End Select
End Sub

Private Sub ViewAccounting()
    If Not mdlProcedures.IsValidComboData(Me.cmbTable) Then Exit Sub
    
    Dim rstTemp As ADODB.Recordset
    
    Select Case mdlProcedures.GetComboData(Me.cmbTable)
        Case clsAccounting.CreateTMITEM(mdlGlobal.conAccounting):
            Set rstTemp = mdlDatabase.OpenRecordset(mdlGlobal.conAccounting, "*", mdlProcedures.GetComboData(Me.cmbTable), False, , "ItemId ASC")
            
            FillViewAccountingTMITEM rstTemp
        Case clsAccounting.CreateTMCUSTOMER(mdlGlobal.conAccounting):
            Set rstTemp = mdlDatabase.OpenRecordset(mdlGlobal.conAccounting, "*", mdlProcedures.GetComboData(Me.cmbTable), False, , "CustomerId ASC")
            
            FillViewAccountingTMCUSTOMER rstTemp
        Case clsAccounting.CreateTMVENDOR(mdlGlobal.conAccounting):
            Set rstTemp = mdlDatabase.OpenRecordset(mdlGlobal.conAccounting, "*", mdlProcedures.GetComboData(Me.cmbTable), False, , "VendorId ASC")
            
            FillViewAccountingTMVENDOR rstTemp
    End Select
    
    mdlDatabase.CloseRecordset rstTemp
End Sub

Private Sub SaveAccounting()
    Select Case objTableMode
        Case ItemMode:
            SaveAccountingTMITEM
        Case CustomerMode:
            SaveAccountingTMCUSTOMER
        Case VendorMode:
            SaveAccountingTMVENDOR
    End Select
End Sub

Private Sub FillViewInventoryTMITEM(ByRef rstTemp As ADODB.Recordset)
    With Me.lsvInventory
        .ListItems.Clear
        .ColumnHeaders.Clear
        
        .ColumnHeaders.Add , , "Kode", 900
        .ColumnHeaders.Add , , "Nama", 2600
    End With
    
    Dim lstItem As ListItem
    
    With rstTemp
        While Not .EOF
            Set lstItem = Me.lsvInventory.ListItems.Add(, , !ItemId)
            lstItem.ListSubItems.Add , , !Name
            
            DoEvents
        
            .MoveNext
        Wend
        
        If .RecordCount > 0 Then
            Me.lsvInventory.ListItems(1).Selected = True
        End If
    End With
End Sub

Private Sub FillViewInventoryTMCUSTOMER(ByRef rstTemp As ADODB.Recordset)
    With Me.lsvInventory
        .ListItems.Clear
        .ColumnHeaders.Clear
        
        .ColumnHeaders.Add , , "Kode", 900
        .ColumnHeaders.Add , , "Nama", 2600
    End With
    
    Dim lstItem As ListItem
    
    With rstTemp
        While Not .EOF
            Set lstItem = Me.lsvInventory.ListItems.Add(, , !CustomerId)
            lstItem.ListSubItems.Add , , !Name
            
            DoEvents
        
            .MoveNext
        Wend
        
        If .RecordCount > 0 Then
            Me.lsvInventory.ListItems(1).Selected = True
        End If
    End With
End Sub

Private Sub FillViewInventoryTMVENDOR(ByRef rstTemp As ADODB.Recordset)
    With Me.lsvInventory
        .ListItems.Clear
        .ColumnHeaders.Clear
        
        .ColumnHeaders.Add , , "Kode", 900
        .ColumnHeaders.Add , , "Nama", 2600
    End With
    
    Dim lstItem As ListItem
    
    With rstTemp
        While Not .EOF
            Set lstItem = Me.lsvInventory.ListItems.Add(, , !VendorId)
            lstItem.ListSubItems.Add , , !Name
            
            DoEvents
        
            .MoveNext
        Wend
        
        If .RecordCount > 0 Then
            Me.lsvInventory.ListItems(1).Selected = True
        End If
    End With
End Sub

Private Sub SaveInventoryTMITEM()
    Dim rstTemp As ADODB.Recordset
    
    Dim intCounter As Integer
    
    Dim strItemId As String
    
    If Me.lsvInventory.ListItems.Count > 0 Then
        For intCounter = 1 To Me.lsvInventory.ListItems.Count
            strItemId = Me.lsvInventory.ListItems(intCounter).Text
            
            Set rstTemp = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "*", mdlTable.CreateTMITEM, , "ItemId='" & strItemId & "'")
            
            With rstTemp
                If Not .RecordCount > 0 Then
                    .AddNew
                    
                    !ItemId = strItemId
                    !ItemDate = mdlProcedures.FormatDate(mdlDatabase.GetFieldData(mdlGlobal.conAccounting, "ItemDate", clsAccounting.CreateTMITEM(mdlGlobal.conAccounting), "ItemId='" & strItemId & "'"))
                    !PartNumber = mdlDatabase.GetFieldData(mdlGlobal.conAccounting, "PartNumber", clsAccounting.CreateTMITEM(mdlGlobal.conAccounting), "ItemId='" & strItemId & "'")
                    !Name = mdlDatabase.GetFieldData(mdlGlobal.conAccounting, "Name", clsAccounting.CreateTMITEM(mdlGlobal.conAccounting), "ItemId='" & strItemId & "'")
                    !GroupId = mdlDatabase.GetFieldData(mdlGlobal.conAccounting, "GroupId", clsAccounting.CreateTMITEM(mdlGlobal.conAccounting), "ItemId='" & strItemId & "'")
                    !CategoryId = mdlDatabase.GetFieldData(mdlGlobal.conAccounting, "CategoryId", clsAccounting.CreateTMITEM(mdlGlobal.conAccounting), "ItemId='" & strItemId & "'")
                    !BrandId = mdlDatabase.GetFieldData(mdlGlobal.conAccounting, "BrandId", clsAccounting.CreateTMITEM(mdlGlobal.conAccounting), "ItemId='" & strItemId & "'")
                    !UnityId = mdlDatabase.GetFieldData(mdlGlobal.conAccounting, "UnityId", clsAccounting.CreateTMITEM(mdlGlobal.conAccounting), "ItemId='" & strItemId & "'")
                    !MinStock = mdlProcedures.GetCurrency("0")
                    !MaxStock = mdlProcedures.GetCurrency("0")
                    !Notes = mdlDatabase.GetFieldData(mdlGlobal.conAccounting, "Notes", clsAccounting.CreateTMITEM(mdlGlobal.conAccounting), "ItemId='" & strItemId & "'")
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

Private Sub SaveInventoryTMCUSTOMER()
    Dim rstTemp As ADODB.Recordset
    
    Dim intCounter As Integer
    
    Dim strCustomerId As String
    
    If Me.lsvInventory.ListItems.Count > 0 Then
        For intCounter = 1 To Me.lsvInventory.ListItems.Count
            strCustomerId = Me.lsvInventory.ListItems(intCounter).Text
            
            Set rstTemp = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "*", mdlTable.CreateTMCUSTOMER, , "CustomerId='" & strCustomerId & "'")
            
            With rstTemp
                If Not .RecordCount > 0 Then
                    .AddNew
                    
                    !CustomerId = strCustomerId
                    !Name = mdlDatabase.GetFieldData(mdlGlobal.conAccounting, "Name", clsAccounting.CreateTMCUSTOMER(mdlGlobal.conAccounting), "CustomerId='" & strCustomerId & "'")
                    !CustomerDate = mdlProcedures.FormatDate(mdlDatabase.GetFieldData(mdlGlobal.conAccounting, "CustomerDate", clsAccounting.CreateTMCUSTOMER(mdlGlobal.conAccounting), "CustomerId='" & strCustomerId & "'"))
                    !Address = mdlDatabase.GetFieldData(mdlGlobal.conAccounting, "Address", clsAccounting.CreateTMCUSTOMER(mdlGlobal.conAccounting), "CustomerId='" & strCustomerId & "'")
                    !Phone = mdlDatabase.GetFieldData(mdlGlobal.conAccounting, "Phone", clsAccounting.CreateTMCUSTOMER(mdlGlobal.conAccounting), "CustomerId='" & strCustomerId & "'")
                    !Fax = mdlDatabase.GetFieldData(mdlGlobal.conAccounting, "Fax", clsAccounting.CreateTMCUSTOMER(mdlGlobal.conAccounting), "CustomerId='" & strCustomerId & "'")
                    !NPWP = mdlDatabase.GetFieldData(mdlGlobal.conAccounting, "NPWP", clsAccounting.CreateTMCUSTOMER(mdlGlobal.conAccounting), "CustomerId='" & strCustomerId & "'")
                    !Notes = mdlDatabase.GetFieldData(mdlGlobal.conAccounting, "Notes", clsAccounting.CreateTMCUSTOMER(mdlGlobal.conAccounting), "CustomerId='" & strCustomerId & "'")
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

Private Sub SaveInventoryTMVENDOR()
    Dim rstTemp As ADODB.Recordset
    
    Dim intCounter As Integer
    
    Dim strVendorId As String
    
    If Me.lsvInventory.ListItems.Count > 0 Then
        For intCounter = 1 To Me.lsvInventory.ListItems.Count
            strVendorId = Me.lsvInventory.ListItems(intCounter).Text
            
            Set rstTemp = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "*", mdlTable.CreateTMVENDOR, , "VendorId='" & strVendorId & "'")
            
            With rstTemp
                If Not .RecordCount > 0 Then
                    .AddNew
                    
                    !VendorId = strVendorId
                    !Name = mdlDatabase.GetFieldData(mdlGlobal.conAccounting, "Name", clsAccounting.CreateTMVENDOR(mdlGlobal.conAccounting), "VendorId='" & strVendorId & "'")
                    !Address = mdlDatabase.GetFieldData(mdlGlobal.conAccounting, "Address", clsAccounting.CreateTMVENDOR(mdlGlobal.conAccounting), "VendorId='" & strVendorId & "'")
                    !Website = mdlDatabase.GetFieldData(mdlGlobal.conAccounting, "Website", clsAccounting.CreateTMVENDOR(mdlGlobal.conAccounting), "VendorId='" & strVendorId & "'")
                    !Email = mdlDatabase.GetFieldData(mdlGlobal.conAccounting, "Email", clsAccounting.CreateTMVENDOR(mdlGlobal.conAccounting), "VendorId='" & strVendorId & "'")
                    !Phone = mdlDatabase.GetFieldData(mdlGlobal.conAccounting, "Phone", clsAccounting.CreateTMVENDOR(mdlGlobal.conAccounting), "VendorId='" & strVendorId & "'")
                    !Fax = mdlDatabase.GetFieldData(mdlGlobal.conAccounting, "Fax", clsAccounting.CreateTMVENDOR(mdlGlobal.conAccounting), "VendorId='" & strVendorId & "'")
                    !Notes = mdlDatabase.GetFieldData(mdlGlobal.conAccounting, "Notes", clsAccounting.CreateTMVENDOR(mdlGlobal.conAccounting), "VendorId='" & strVendorId & "'")
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

Private Sub FillViewAccountingTMITEM(ByRef rstTemp As ADODB.Recordset)
    With Me.lsvAccounting
        .ListItems.Clear
        .ColumnHeaders.Clear
        
        .ColumnHeaders.Add , , "Kode", 900
        .ColumnHeaders.Add , , "Nama", 2600
    End With
    
    Dim lstItem As ListItem
    
    With rstTemp
        While Not .EOF
            Set lstItem = Me.lsvAccounting.ListItems.Add(, , !ItemId)
            lstItem.ListSubItems.Add , , !Name
            
            DoEvents
        
            .MoveNext
        Wend
        
        If .RecordCount > 0 Then
            Me.lsvInventory.ListItems(1).Selected = True
        End If
    End With
End Sub

Private Sub FillViewAccountingTMCUSTOMER(ByRef rstTemp As ADODB.Recordset)
    With Me.lsvAccounting
        .ListItems.Clear
        .ColumnHeaders.Clear
        
        .ColumnHeaders.Add , , "Kode", 900
        .ColumnHeaders.Add , , "Nama", 2600
    End With
    
    Dim lstItem As ListItem
    
    With rstTemp
        While Not .EOF
            Set lstItem = Me.lsvAccounting.ListItems.Add(, , !CustomerId)
            lstItem.ListSubItems.Add , , !Name
            
            DoEvents
        
            .MoveNext
        Wend
        
        If .RecordCount > 0 Then
            Me.lsvInventory.ListItems(1).Selected = True
        End If
    End With
End Sub

Private Sub FillViewAccountingTMVENDOR(ByRef rstTemp As ADODB.Recordset)
    With Me.lsvAccounting
        .ListItems.Clear
        .ColumnHeaders.Clear
        
        .ColumnHeaders.Add , , "Kode", 900
        .ColumnHeaders.Add , , "Nama", 2600
    End With
    
    Dim lstItem As ListItem
    
    With rstTemp
        While Not .EOF
            Set lstItem = Me.lsvAccounting.ListItems.Add(, , !VendorId)
            lstItem.ListSubItems.Add , , !Name
            
            DoEvents
        
            .MoveNext
        Wend
        
        If .RecordCount > 0 Then
            Me.lsvInventory.ListItems(1).Selected = True
        End If
    End With
End Sub

Private Sub SaveAccountingTMITEM()
    Dim rstTemp As ADODB.Recordset
    
    Dim intCounter As Integer
    
    Dim strItemId As String
    
    If Me.lsvAccounting.ListItems.Count > 0 Then
        For intCounter = 1 To Me.lsvAccounting.ListItems.Count
            strItemId = Me.lsvAccounting.ListItems(intCounter).Text
            
            Set rstTemp = mdlDatabase.OpenRecordset(mdlGlobal.conAccounting, "*", clsAccounting.CreateTMITEM(mdlGlobal.conAccounting), , "ItemId='" & strItemId & "'")
            
            With rstTemp
                If Not .RecordCount > 0 Then
                    .AddNew
                    
                    !ItemId = strItemId
                    !ItemDate = mdlProcedures.FormatDate(mdlDatabase.GetFieldData(mdlGlobal.conInventory, "ItemDate", mdlTable.CreateTMITEM, "ItemId='" & strItemId & "'"))
                    !PartNumber = mdlDatabase.GetFieldData(mdlGlobal.conInventory, "PartNumber", mdlTable.CreateTMITEM, "ItemId='" & strItemId & "'")
                    !Name = mdlDatabase.GetFieldData(mdlGlobal.conInventory, "Name", mdlTable.CreateTMITEM, "ItemId='" & strItemId & "'")
                    !GroupId = mdlDatabase.GetFieldData(mdlGlobal.conInventory, "GroupId", mdlTable.CreateTMITEM, "ItemId='" & strItemId & "'")
                    !CategoryId = mdlDatabase.GetFieldData(mdlGlobal.conInventory, "CategoryId", mdlTable.CreateTMITEM, "ItemId='" & strItemId & "'")
                    !BrandId = mdlDatabase.GetFieldData(mdlGlobal.conInventory, "BrandId", mdlTable.CreateTMITEM, "ItemId='" & strItemId & "'")
                    !UnityId = mdlDatabase.GetFieldData(mdlGlobal.conInventory, "UnityId", mdlTable.CreateTMITEM, "ItemId='" & strItemId & "'")
                    !Notes = mdlDatabase.GetFieldData(mdlGlobal.conInventory, "Notes", mdlTable.CreateTMITEM, "ItemId='" & strItemId & "'")
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

Private Sub SaveAccountingTMCUSTOMER()
    Dim rstTemp As ADODB.Recordset
    
    Dim intCounter As Integer
    
    Dim strCustomerId As String
    
    If Me.lsvAccounting.ListItems.Count > 0 Then
        For intCounter = 1 To Me.lsvAccounting.ListItems.Count
            strCustomerId = Me.lsvAccounting.ListItems(intCounter).Text
            
            Set rstTemp = mdlDatabase.OpenRecordset(mdlGlobal.conAccounting, "*", clsAccounting.CreateTMCUSTOMER(mdlGlobal.conAccounting), , "CustomerId='" & strCustomerId & "'")
            
            With rstTemp
                If Not .RecordCount > 0 Then
                    .AddNew
                    
                    !CustomerId = strCustomerId
                    !Name = mdlDatabase.GetFieldData(mdlGlobal.conInventory, "Name", mdlTable.CreateTMCUSTOMER, "CustomerId='" & strCustomerId & "'")
                    !CustomerDate = mdlProcedures.FormatDate(mdlDatabase.GetFieldData(mdlGlobal.conInventory, "CustomerDate", mdlTable.CreateTMCUSTOMER, "CustomerId='" & strCustomerId & "'"))
                    !Address = mdlDatabase.GetFieldData(mdlGlobal.conInventory, "Address", mdlTable.CreateTMCUSTOMER, "CustomerId='" & strCustomerId & "'")
                    !Phone = mdlDatabase.GetFieldData(mdlGlobal.conInventory, "Phone", mdlTable.CreateTMCUSTOMER, "CustomerId='" & strCustomerId & "'")
                    !Fax = mdlDatabase.GetFieldData(mdlGlobal.conInventory, "Fax", mdlTable.CreateTMCUSTOMER, "CustomerId='" & strCustomerId & "'")
                    !NPWP = mdlDatabase.GetFieldData(mdlGlobal.conInventory, "NPWP", mdlTable.CreateTMCUSTOMER, "CustomerId='" & strCustomerId & "'")
                    !Notes = mdlDatabase.GetFieldData(mdlGlobal.conInventory, "Notes", mdlTable.CreateTMCUSTOMER, "CustomerId='" & strCustomerId & "'")
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

Private Sub SaveAccountingTMVENDOR()
    Dim rstTemp As ADODB.Recordset
    
    Dim intCounter As Integer
    
    Dim strVendorId As String
    
    If Me.lsvAccounting.ListItems.Count > 0 Then
        For intCounter = 1 To Me.lsvAccounting.ListItems.Count
            strVendorId = Me.lsvAccounting.ListItems(intCounter).Text
            
            Set rstTemp = mdlDatabase.OpenRecordset(mdlGlobal.conAccounting, "*", clsAccounting.CreateTMVENDOR(mdlGlobal.conAccounting), , "VendorId='" & strVendorId & "'")
            
            With rstTemp
                If Not .RecordCount > 0 Then
                    .AddNew
                    
                    !VendorId = strVendorId
                    !Name = mdlDatabase.GetFieldData(mdlGlobal.conInventory, "Name", mdlTable.CreateTMVENDOR, "VendorId='" & strVendorId & "'")
                    !Address = mdlDatabase.GetFieldData(mdlGlobal.conInventory, "Address", mdlTable.CreateTMVENDOR, "VendorId='" & strVendorId & "'")
                    !Website = mdlDatabase.GetFieldData(mdlGlobal.conInventory, "Website", mdlTable.CreateTMVENDOR, "VendorId='" & strVendorId & "'")
                    !Email = mdlDatabase.GetFieldData(mdlGlobal.conInventory, "Email", mdlTable.CreateTMVENDOR, "VendorId='" & strVendorId & "'")
                    !Phone = mdlDatabase.GetFieldData(mdlGlobal.conInventory, "Phone", mdlTable.CreateTMVENDOR, "VendorId='" & strVendorId & "'")
                    !Fax = mdlDatabase.GetFieldData(mdlGlobal.conInventory, "Fax", mdlTable.CreateTMVENDOR, "VendorId='" & strVendorId & "'")
                    !Notes = mdlDatabase.GetFieldData(mdlGlobal.conInventory, "Notes", mdlTable.CreateTMVENDOR, "VendorId='" & strVendorId & "'")
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
