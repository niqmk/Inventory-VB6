VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPRTTHDOBUY"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const DOID_COL As Integer = 1
Private Const DATE_COL As Integer = 1
Private Const POID_COL As Integer = 1
Private Const VENDORNAME_COL As Integer = 1
Private Const WAREHOUSENAME_COL As Integer = 1

Private Const ITEMID_COL As Integer = 1
Private Const NAME_COL As Integer = 2
Private Const QTY_COL As Integer = 5
Private Const UNITYNAME_COL As Integer = 6

Private xlSheet As Excel.Worksheet

Private intRow As Integer

Private Sub Class_Initialize()
    mdlExcel.OpenExcel
End Sub

Public Sub ImportToExcel(ByVal strDOId As String)
    If Trim(strDOId) = "" Then Exit Sub
    
    mdlExcel.OpenWorkSheet xlSheet
    
    SetWidth
    
    FillTitle strDOId
    
    FillDetail strDOId
End Sub

Private Sub SetWidth()
    intRow = 1
    
    With xlSheet
        .PageSetup.LeftMargin = 25
        .PageSetup.RightMargin = 25
        
        .Columns("A").ColumnWidth = 10
        .Columns("B").ColumnWidth = 1
        .Columns("C").ColumnWidth = 25
        .Columns("D").ColumnWidth = 22
        .Columns("E").ColumnWidth = 11
        .Columns("F").ColumnWidth = 25
        
        .Rows(intRow).RowHeight = 125
    End With
End Sub

Private Sub FillTitle(ByVal strDOId As String)
    On Local Error GoTo ErrHandler
    
    With xlSheet
        .Shapes.AddPicture mdlGlobal.strLogoImageText, True, True, 0, 0, 100, 100
        
        intRow = intRow + 1
        
        .Range("A" & intRow & ":C" & intRow + 4).Font.Name = "Courier New"
        
        .Cells(intRow, DOID_COL).Value = "NO. DO"
        .Cells(intRow, DOID_COL + 1).Value = ":"
        .Cells(intRow, DOID_COL + 2).Value = strDOId
        
        intRow = intRow + 1
        
        .Cells(intRow, DATE_COL).Value = "TGL."
        .Cells(intRow, DATE_COL + 1).Value = ":"
        .Cells(intRow, DATE_COL + 2).HorizontalAlignment = xlHAlignLeft
        .Cells(intRow, DATE_COL + 2).Value = mdlProcedures.FormatDate(mdlDatabase.GetFieldData(mdlGlobal.conInventory, "DODate", mdlTable.CreateTHDOBUY, "DOId='" & strDOId & "'"))
        .Cells(intRow, DATE_COL + 2).NumberFormat = "dd MMMM yyyy"
        
        intRow = intRow + 1
        
        Dim strPOId As String
        
        strPOId = mdlDatabase.GetFieldData(mdlGlobal.conInventory, "POId", mdlTable.CreateTHDOBUY, "DOId='" & strDOId & "'")
        
        .Cells(intRow, POID_COL).Value = "NO. PO"
        .Cells(intRow, POID_COL + 1).Value = ":"
        .Cells(intRow, POID_COL + 2).Value = strPOId
        
        intRow = intRow + 1
        
        .Cells(intRow, VENDORNAME_COL).Value = "PEMASOK"
        .Cells(intRow, VENDORNAME_COL + 1).Value = ":"
        .Cells(intRow, VENDORNAME_COL + 2).Value = _
            "'" & mdlDatabase.GetFieldData(mdlGlobal.conInventory, "Name", mdlTable.CreateTMVENDOR, "VendorId='" & _
                mdlDatabase.GetFieldData(mdlGlobal.conInventory, "VendorId", mdlTable.CreateTHPOBUY, "POId='" & strPOId & "'") & "'")
        
        intRow = intRow + 1
        
        Dim strWarehouseId As String
        
        strWarehouseId = mdlDatabase.GetFieldData(mdlGlobal.conInventory, "WarehouseId", mdlTable.CreateTHDOBUY, "DOId='" & strDOId & "'")
        
        .Cells(intRow, WAREHOUSENAME_COL).Value = "GUDANG"
        .Cells(intRow, WAREHOUSENAME_COL + 1).Value = ":"
        .Cells(intRow, WAREHOUSENAME_COL + 2).Value = _
            mdlDatabase.GetFieldData(mdlGlobal.conInventory, "Name", mdlTable.CreateTMWAREHOUSE, "WarehouseId='" & strWarehouseId & "'")
        
        intRow = intRow + 2
        
        .Range("A" & intRow & ":F" & intRow).Font.Bold = True
        .Range("B" & intRow & ":D" & intRow).MergeCells = True
        .Range("A" & intRow & ":F" & intRow).HorizontalAlignment = xlHAlignCenter
        
        mdlExcel.FillEdge xlSheet, "A" & intRow & ":F" & intRow + 1
        
        .Cells(intRow, ITEMID_COL).Value = "Kode"
        .Cells(intRow, NAME_COL).Value = "Nama"
        .Cells(intRow, QTY_COL).Value = "Qty"
        .Cells(intRow, UNITYNAME_COL).Value = "Satuan"
        
        intRow = intRow + 1
        
        .Rows(intRow).RowHeight = 3
        
        .Range("A" & intRow & ":F" & intRow).MergeCells = True
    End With
    
ErrHandler:
End Sub

Private Sub FillDetail(Optional ByVal strDOId As String = "")
    On Local Error GoTo ErrHandler
    
    Dim rstDetail As ADODB.Recordset
    
    Set rstDetail = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "*", mdlTable.CreateTDDOBUY, False, "DOId='" & strDOId & "'", "ItemId ASC")
    
    With rstDetail
        Dim intTemp As Integer
        
        Dim strItemId As String
        
        If .RecordCount > 0 Then
            intTemp = intRow + 1
        End If
        
        While Not .EOF
            intRow = intRow + 1
            
            xlSheet.Range("A" & intRow & ":D" & intRow).HorizontalAlignment = xlHAlignLeft
            xlSheet.Range("E" & intRow & ":E" & intRow).HorizontalAlignment = xlHAlignRight
            xlSheet.Range("F" & intRow & ":F" & intRow).HorizontalAlignment = xlHAlignLeft
            xlSheet.Range("A" & intRow & ":F" & intRow).VerticalAlignment = xlVAlignTop
            
            xlSheet.Range("B" & intRow & ":D" & intRow).MergeCells = True
            xlSheet.Range("F" & intRow & ":F" & intRow).WrapText = True
            
            strItemId = Trim(.Fields("ItemId").Value)
            
            xlSheet.Cells(intRow, ITEMID_COL).Value = "'" & strItemId
            xlSheet.Cells(intRow, NAME_COL).Value = "'" & Trim(mdlDatabase.GetFieldData(mdlGlobal.conInventory, "Name", mdlTable.CreateTMITEM, "ItemId='" & strItemId & "'"))
            xlSheet.Cells(intRow, QTY_COL).Value = CStr(.Fields("Qty").Value)
            xlSheet.Cells(intRow, QTY_COL).NumberFormat = "#,##0"
            xlSheet.Cells(intRow, UNITYNAME_COL).Value = "'" & _
                Trim(mdlDatabase.GetFieldData(mdlGlobal.conInventory, "Name", mdlTable.CreateTMUNITY, "UnityId='" & _
                    mdlDatabase.GetFieldData(mdlGlobal.conInventory, "UnityId", mdlTable.CreateTMITEM, "ItemId='" & strItemId & "'") & "'"))
            
            xlSheet.Range("A" & intTemp & ":F" & intTemp).Select
            
            .MoveNext
        Wend
        
        If .RecordCount > 0 Then
            mdlExcel.FillEdge xlSheet, "A" & intTemp & ":F" & intRow, False, , False
            
            xlSheet.Range("A" & intTemp & ":F" & intTemp).Select
        End If
    End With
    
ErrHandler:
End Sub

Private Sub Class_Terminate()
    mdlExcel.CloseExcel
End Sub
