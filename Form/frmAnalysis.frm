VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form frmAnalysis 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6630
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   7095
   Icon            =   "frmAnalysis.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6630
   ScaleWidth      =   7095
   ShowInTaskbar   =   0   'False
   Begin MSChart20Lib.MSChart mscMain 
      Height          =   2895
      Left            =   120
      OleObjectBlob   =   "frmAnalysis.frx":1F8A
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   3600
      Width           =   6855
   End
   Begin MSComctlLib.ListView lsvMain 
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   5953
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
End
Attribute VB_Name = "frmAnalysis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    SetInitialization
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmAnalysis = Nothing
End Sub

Private Sub lsvMain_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If Me.lsvMain.ListItems.Count > 0 Then
        FillGraph Item.Text
    End If
End Sub

Private Sub SetInitialization()
    mdlProcedures.CenterWindows Me
    
    Me.Caption = mdlText.strAnalysis
    
    With Me.lsvMain
        .FullRowSelect = True
        .LabelEdit = lvwManual
        .View = lvwReport
        
        .ColumnHeaders.Add , , "Kode", 900
        .ColumnHeaders.Add , , "Nama", 2500
        .ColumnHeaders.Add , , "Telepon"
        .ColumnHeaders.Add , , "Fax"
    End With
    
    FillInfo
End Sub

Private Sub FillInfo()
    Dim strYear As String
    Dim strMonth As String
    
    strYear = mdlProcedures.FormatDate(Now, "yyyy")
    strMonth = mdlProcedures.FormatDate(Now, "MM")
    
    Dim dteStart As Date
    Dim dteFinish As Date
    
    dteStart = mdlProcedures.SetDate(strMonth, CStr(CInt(strYear) - 1))
    dteFinish = mdlProcedures.SetDate(strMonth, strYear, , True)

    Dim strCriteria As String
    
    If mdlGlobal.objDatabaseInit = SQLSERVER7 Or mdlGlobal.objDatabaseInit = SQLSERVER2000 Or mdlGlobal.objDatabaseInit = SQLEXPRESS Then
        strCriteria = "PODate>='" & mdlProcedures.FormatDate(dteStart) & "' AND PODate<='" & mdlProcedures.FormatDate(dteFinish) & "'"
    ElseIf mdlGlobal.objDatabaseInit = MSACCESS Then
        strCriteria = "PODate>=#" & mdlProcedures.FormatDate(dteStart) & "# AND PODate<=#" & mdlProcedures.FormatDate(dteFinish) & "#"
    ElseIf mdlGlobal.objDatabaseInit = MYSQL Then
        strCriteria = "PODate>='" & mdlProcedures.FormatDate(dteStart) & "' AND PODate<='" & mdlProcedures.FormatDate(dteFinish) & "'"
    End If

    Dim rstTemp As ADODB.Recordset
    
    Set rstTemp = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "CustomerId, COUNT(*)", mdlTable.CreateTHPOSELL, False, strCriteria, , "CustomerId")
    
    With rstTemp
        If .RecordCount > 0 Then
            .Sort = rstTemp.Fields(1).Name & " DESC"
            
            FillGraph !CustomerId
        Else
            FillGraph
        End If
        
        Dim strCustomerId As String
        Dim lstItem As ListItem
        
        strCustomerId = ""
        
        While Not .EOF
            If Not UCase(Trim(strCustomerId)) = UCase(Trim(!CustomerId)) Then
                strCustomerId = !CustomerId
            End If
            
            Set lstItem = Me.lsvMain.ListItems.Add(, , strCustomerId)
            lstItem.ListSubItems.Add , , mdlDatabase.GetFieldData(mdlGlobal.conInventory, "Name", mdlTable.CreateTMCUSTOMER, "CustomerId='" & strCustomerId & "'")
            lstItem.ListSubItems.Add , , mdlDatabase.GetFieldData(mdlGlobal.conInventory, "Phone", mdlTable.CreateTMCUSTOMER, "CustomerId='" & strCustomerId & "'")
            lstItem.ListSubItems.Add , , mdlDatabase.GetFieldData(mdlGlobal.conInventory, "Fax", mdlTable.CreateTMCUSTOMER, "CustomerId='" & strCustomerId & "'")
            
            .MoveNext
        Wend
    End With
    
    mdlDatabase.CloseRecordset rstTemp
End Sub

Private Sub FillGraph(Optional ByVal strCustomerId As String = "")
    Dim intCounter As Integer
    
    Dim strYear As String
    Dim strMonth As String
    
    strYear = mdlProcedures.FormatDate(Now, "yyyy")
    strMonth = mdlProcedures.FormatDate(Now, "MM")
    
    Dim dteStart As Date
    Dim dteTemp As Date
    
    dteStart = mdlProcedures.SetDate(strMonth, CStr(CInt(strYear) - 1))
    
    With Me.mscMain
        Dim strCriteria As String
        
        .RowCount = 13
        
        For intCounter = 1 To 13
            .Row = intCounter
            
            dteTemp = DateAdd("M", CDbl(intCounter - 1), dteStart)
            
            .RowLabel = mdlProcedures.FormatDate(dteTemp, "MM/yy")
            
            strCriteria = "CustomerId='" & strCustomerId & "'"
            strCriteria = strCriteria & " AND MONTH(PODate)=" & mdlProcedures.FormatDate(dteTemp, "MM") & " AND YEAR(PODate)=" & mdlProcedures.FormatDate(dteTemp, "yyyy")
            
            .Data = mdlDatabase.GetFieldData(mdlGlobal.conInventory, "COUNT(CustomerId)", mdlTable.CreateTHPOSELL, strCriteria)
        Next intCounter
    End With
End Sub
