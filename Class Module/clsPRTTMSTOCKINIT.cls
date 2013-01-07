VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPRTTMSTOCKINIT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const ITEMID_COL As Integer = 1
Private Const PARTNUMBER_COL As Integer = 2
Private Const NAME_COL As Integer = 3
Private Const QTY_COL As Integer = 4
Private Const UNITYNAME_COL As Integer = 5

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
        .Columns("B").ColumnWidth = 20
        .Columns("C").ColumnWidth = 20
        .Columns("D").ColumnWidth = 10
        .Columns("E").ColumnWidth = 25
        
        .Rows(intRow).RowHeight = 125
    End With
End Sub

Private Sub FillTitle()
    On Local Error GoTo ErrHandler
    
    With xlSheet
        .Shapes.AddPicture mdlGlobal.strLogoImageText, True, True, 0, 0, 100, 100
        
        intRow = intRow + 1
        
        mdlExcel.FillEdge xlSheet, "A" & intRow & ":E" & intRow + 1
        
        .Range("A" & intRow & ":E" & intRow).Font.Bold = True
        
        .Cells(intRow, ITEMID_COL).Value = "Kode"
        .Cells(intRow, PARTNUMBER_COL).Value = "Nomor Part"
        .Cells(intRow, NAME_COL).Value = "Nama"
        .Cells(intRow, QTY_COL).Value = "Qty"
        .Cells(intRow, UNITYNAME_COL).Value = "Satuan"
    End With
    
ErrHandler:
End Sub

Private Sub FillText(ByRef rstMain As ADODB.Recordset)
    On Local Error GoTo ErrHandler
    
    With rstMain
        Dim intTemp As Integer
        
        Dim strItemId As String

        intTemp = intRow + 1
    
        While Not .EOF
            intRow = intRow + 1
            
            xlSheet.Range("A" & intRow & ":E" & intRow).VerticalAlignment = xlTop
            xlSheet.Range("A" & intRow & ":C" & intRow).HorizontalAlignment = xlLeft
            xlSheet.Range("D" & intRow & ":D" & intRow).HorizontalAlignment = xlRight
            xlSheet.Range("E" & intRow & ":E" & intRow).HorizontalAlignment = xlLeft
            xlSheet.Range("A" & intRow & ":E" & intRow).WrapText = True
            
            strItemId = Trim(.Fields("ItemId").Value)
            
            xlSheet.Cells(intRow, ITEMID_COL).Value = "'" & strItemId
            xlSheet.Cells(intRow, PARTNUMBER_COL).Value = "'" & Trim(mdlDatabase.GetFieldData(mdlGlobal.conInventory, "PartNumber", mdlTable.CreateTMITEM, "ItemId='" & strItemId & "'"))
            xlSheet.Cells(intRow, NAME_COL).Value = "'" & Trim(mdlDatabase.GetFieldData(mdlGlobal.conInventory, "Name", mdlTable.CreateTMITEM, "ItemId='" & strItemId & "'"))
            xlSheet.Cells(intRow, QTY_COL).Value = CStr(.Fields("Qty").Value)
            xlSheet.Cells(intRow, QTY_COL).NumberFormat = "#,##0"
            xlSheet.Cells(intRow, UNITYNAME_COL).Value = "'" & Trim(mdlDatabase.GetFieldData( _
                mdlGlobal.conInventory, _
                "Name", _
                mdlTable.CreateTMUNITY, _
                "UnityId='" & _
                    mdlDatabase.GetFieldData( _
                        mdlGlobal.conInventory, _
                        "UnityId", _
                        mdlTable.CreateTMITEM, _
                        "ItemId='" & strItemId & "'") & "'"))
            
            xlSheet.Range("A" & intRow & ":E" & intRow).Select
            
            .MoveNext
        Wend
        
        If .RecordCount > 0 Then
            mdlExcel.FillEdge xlSheet, "A" & intTemp & ":E" & intRow, False
            
            xlSheet.Range("A" & intTemp & ":E" & intTemp).Select
        End If
    End With
    
ErrHandler:
End Sub

Private Sub Class_Terminate()
    mdlExcel.CloseExcel
End Sub
