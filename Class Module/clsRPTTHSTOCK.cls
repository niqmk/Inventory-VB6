VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRPTTHSTOCK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const ITEMID_COL As Integer = 1
Private Const PARTNUMBER_COL As Integer = 1
Private Const NAME_COL As Integer = 1
Private Const QTYLAST_COL As Integer = 1

Private Const STOCKDATE_COL As Integer = 2
Private Const QTYIN_COL As Integer = 3
Private Const QTYOUT_COL As Integer = 4
Private Const REFERENCESNUMBER_COL As Integer = 5

Private xlSheet As Excel.Worksheet

Private intRow As Integer

Private Sub Class_Initialize()
    mdlExcel.OpenExcel
End Sub

Public Sub ImportToExcel( _
    ByRef rstMain As ADODB.Recordset, _
    ByVal dteStartDate As Date, _
    ByVal dteFinishDate As Date, _
    ByVal strWarehouseId As String, _
    Optional ByVal blnStockEmpty As Boolean = False, _
    Optional ByVal blnDetail As Boolean = False, _
    Optional ByVal blnPrintAll As Boolean = True)
    If rstMain Is Nothing Then Exit Sub
    
    mdlExcel.OpenWorkSheet xlSheet
    
    SetWidth
    
    FillTitle
    
    FillHeader _
        rstMain, _
        dteStartDate, _
        dteFinishDate, _
        strWarehouseId, _
        blnStockEmpty, _
        blnDetail, _
        blnPrintAll
End Sub

Private Sub SetWidth()
    intRow = 1
    
    With xlSheet
        .PageSetup.LeftMargin = 30
        .PageSetup.RightMargin = 30
        
        .Columns("A").ColumnWidth = 5
        .Columns("B").ColumnWidth = 18
        .Columns("C").ColumnWidth = 18
        .Columns("D").ColumnWidth = 18
        .Columns("E").ColumnWidth = 25
        .Columns("F").ColumnWidth = 5
        
        .Rows(intRow).RowHeight = 125
    End With
End Sub

Private Sub FillTitle()
    On Local Error GoTo ErrHandler
    
    With xlSheet
        .Shapes.AddPicture mdlGlobal.strLogoImageText, True, True, 0, 0, 100, 100
        
        intRow = intRow + 2
    End With
    
ErrHandler:
End Sub

Private Sub FillHeader( _
    ByRef rstHeader As ADODB.Recordset, _
    ByVal dteStartDate As Date, _
    ByVal dteFinishDate As Date, _
    ByVal strWarehouseId As String, _
    Optional ByVal blnStockEmpty As Boolean = False, _
    Optional ByVal blnDetail As Boolean = False, _
    Optional ByVal blnPrintAll As Boolean = True)
    On Local Error GoTo ErrHandler
    
    With rstHeader
        Dim intFirst As Integer
        Dim intTemp As Integer
        
        intFirst = intRow
    
        While Not .EOF
            If blnStockEmpty Then
                If mdlTransaction.CheckStock(.Fields("ItemId").Value, strWarehouseId) > 0 Then
                    GoTo NextRecordset
                End If
            Else
                If Not blnPrintAll Then
                    If Not mdlTransaction.CheckStock(.Fields("ItemId").Value, strWarehouseId) > 0 Then
                        GoTo NextRecordset
                    End If
                End If
            End If
            
            If Not intRow = intFirst Then
                intRow = intRow + 2
                
                xlSheet.Rows(intRow - 1).RowHeight = 5
            End If
            
            intTemp = intRow
            
            xlSheet.Range("A" & intRow & ":F" & intRow + 3).Font.Name = "Courier New"
            xlSheet.Range("A" & intRow & ":F" & intRow + 3).Font.Size = 10
            xlSheet.Range("A" & intRow & ":F" & intRow + 3).HorizontalAlignment = xlLeft
            
            xlSheet.Cells(intRow, ITEMID_COL).Value = "Kode       : " & Trim(.Fields("ItemId").Value)
            
            intRow = intRow + 1
            
            xlSheet.Cells(intRow, PARTNUMBER_COL).Value = "Nomor Part : " & Trim(.Fields("PartNumber").Value)
            
            intRow = intRow + 1
            
            xlSheet.Cells(intRow, NAME_COL).Value = "Nama       : " & Trim(.Fields("Name").Value)
            
            intRow = intRow + 1
            
            xlSheet.Cells(intRow, QTYLAST_COL).Value = "Sisa Stok  : " & mdlProcedures.FormatCurrency(mdlTransaction.CheckStock(.Fields("ItemId").Value, strWarehouseId))
            
            xlSheet.Range("A" & intRow & ":D" & intRow).Select
            
            If blnDetail Then
                FillDetail .Fields("ItemId").Value, dteStartDate, dteFinishDate, strWarehouseId
            End If
            
            mdlExcel.FillEdge xlSheet, "A" & intTemp & ":F" & intRow, , , False, False

NextRecordset:
            .MoveNext
        Wend
        
        If .RecordCount > 0 Then
            xlSheet.Range("A" & intFirst & ":F" & intFirst).Select
        End If
    End With
    
ErrHandler:
End Sub

Private Sub FillDetail( _
    ByVal strItemId As String, _
    ByVal dteStartDate As Date, _
    ByVal dteFinishDate As Date, _
    ByVal strWarehouseId As String)
    On Local Error GoTo ErrHandler
    
    Dim strCriteria As String
    
    strCriteria = "ItemId='" & strItemId & "'"
    strCriteria = strCriteria & " AND WarehouseId='" & strWarehouseId & "'"
    
    If mdlGlobal.objDatabaseInit = SQLSERVER7 Or mdlGlobal.objDatabaseInit = SQLSERVER2000 Or mdlGlobal.objDatabaseInit = SQLEXPRESS Then
        strCriteria = strCriteria & "AND StockDate>='" & mdlProcedures.FormatDate(dteStartDate) & "' AND StockDate<='" & mdlProcedures.FormatDate(dteFinishDate) & "'"
    ElseIf mdlGlobal.objDatabaseInit = MSACCESS Then
        strCriteria = strCriteria & "AND StockDate>=#" & mdlProcedures.FormatDate(dteStartDate) & "# AND StockDate<=#" & mdlProcedures.FormatDate(dteFinishDate) & "#"
    ElseIf mdlGlobal.objDatabaseInit = MYSQL Then
        strCriteria = strCriteria & "AND StockDate>='" & mdlProcedures.FormatDate(dteStartDate) & "' AND StockDate<='" & mdlProcedures.FormatDate(dteFinishDate) & "'"
    End If

    Dim rstDetail As ADODB.Recordset

    Set rstDetail = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "*", mdlTable.CreateTHSTOCK, False, strCriteria, "StockDate ASC")

    With rstDetail
        If .RecordCount > 0 Then
            intRow = intRow + 1
                
            xlSheet.Rows(intRow).RowHeight = 10
            
            intRow = intRow + 1
            
            xlSheet.Range("B" & intRow & ":E" & intRow).HorizontalAlignment = xlHAlignCenter
            xlSheet.Range("B" & intRow & ":E" & intRow).Font.Bold = True
            
            xlSheet.Cells(intRow, STOCKDATE_COL).Value = "Tanggal"
            xlSheet.Cells(intRow, QTYIN_COL).Value = "Masuk"
            xlSheet.Cells(intRow, QTYOUT_COL).Value = "Keluar"
            xlSheet.Cells(intRow, REFERENCESNUMBER_COL).Value = "Nomor Surat"
            
            mdlExcel.FillEdge xlSheet, "B" & intRow & ":E" & intRow + .RecordCount
        End If

        While Not .EOF
            intRow = intRow + 1

            xlSheet.Range("B" & intRow & ":B" & intRow).HorizontalAlignment = xlLeft
            xlSheet.Range("C" & intRow & ":D" & intRow).HorizontalAlignment = xlRight
            xlSheet.Range("E" & intRow & ":E" & intRow).HorizontalAlignment = xlLeft
            
            xlSheet.Cells(intRow, STOCKDATE_COL).Value = "'" & mdlProcedures.FormatDate(.Fields("StockDate").Value, "dd MMMM yyyy")
            xlSheet.Cells(intRow, QTYIN_COL).Value = .Fields("QtyIn").Value
            xlSheet.Cells(intRow, QTYIN_COL).NumberFormat = "#,##0"
            xlSheet.Cells(intRow, QTYOUT_COL).Value = .Fields("QtyOut").Value
            xlSheet.Cells(intRow, QTYOUT_COL).NumberFormat = "#,##0"
            xlSheet.Cells(intRow, REFERENCESNUMBER_COL).Value = "'" & Trim(.Fields("ReferencesNumber").Value)

            .MoveNext
        Wend
        
        If .RecordCount > 0 Then
            intRow = intRow + 1
            
            xlSheet.Rows(intRow).RowHeight = 10
        End If
    End With

ErrHandler:
End Sub

Private Sub Class_Terminate()
    mdlExcel.CloseExcel
End Sub
