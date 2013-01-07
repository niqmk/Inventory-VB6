VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPRTTMGROUP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const GROUPID_COL As Integer = 1
Private Const NAME_COL As Integer = 2

Private xlSheet As Excel.Worksheet

Private intRow As Integer

Private Sub Class_Initialize()
    mdlExcel.OpenExcel
End Sub

Public Sub ImportToExcel(ByRef rstMain As ADODB.Recordset)
    If rstMain Is Nothing Then Exit Sub
    
    mdlExcel.OpenWorkSheet xlSheet
    
    SetWidth
    
    FillTitle
    
    FillText rstMain
End Sub

Private Sub SetWidth()
    intRow = 1
    
    With xlSheet
        .Columns("A").ColumnWidth = 8
        .Columns("B").ColumnWidth = 70
        
        .Rows(intRow).RowHeight = 125
    End With
End Sub

Private Sub FillTitle()
    On Local Error GoTo ErrHandler
    
    With xlSheet
        .Shapes.AddPicture mdlGlobal.strLogoImageText, True, True, 0, 0, 100, 100
        
        intRow = intRow + 1
        
        mdlExcel.FillEdge xlSheet, "A" & intRow & ":B" & intRow + 1
        
        .Range("A" & intRow & ":B" & intRow).Font.Bold = True
        
        .Cells(intRow, GROUPID_COL).Value = "Kode"
        .Cells(intRow, NAME_COL).Value = "Nama"
    End With
    
ErrHandler:
End Sub

Private Sub FillText(ByRef rstMain As ADODB.Recordset)
    On Local Error GoTo ErrHandler
    
    With rstMain
        Dim intTemp As Integer

        intTemp = intRow + 1
    
        While Not .EOF
            intRow = intRow + 1
            
            xlSheet.Range("A" & intRow & ":B" & intRow).VerticalAlignment = xlTop
            xlSheet.Range("A" & intRow & ":B" & intRow).HorizontalAlignment = xlLeft
            xlSheet.Range("A" & intRow & ":B" & intRow).WrapText = True
            
            xlSheet.Cells(intRow, GROUPID_COL).Value = "'" & Trim(.Fields("GroupId").Value)
            xlSheet.Cells(intRow, NAME_COL).Value = "'" & Trim(.Fields("Name").Value)
            
            xlSheet.Range("A" & intRow & ":B" & intRow).Select
            
            .MoveNext
        Wend
        
        If .RecordCount > 0 Then
            mdlExcel.FillEdge xlSheet, "A" & intTemp & ":B" & intRow, False
            
            xlSheet.Range("A" & intTemp & ":B" & intTemp).Select
        End If
    End With
    
ErrHandler:
End Sub

Private Sub Class_Terminate()
    mdlExcel.CloseExcel
End Sub
