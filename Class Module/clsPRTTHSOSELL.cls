VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPRTTHSOSELL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const SOID_COL As Integer = 1
Private Const DATE_COL As Integer = 1
Private Const POID_COL As Integer = 1
Private Const CUSTOMERNAME_COL As Integer = 1
Private Const CURRENCYNAME_COL As Integer = 1

Private Const ITEMID_COL As Integer = 1
Private Const NAME_COL As Integer = 2
Private Const QTY_COL As Integer = 4
Private Const ITEMPRICE_COL As Integer = 5
Private Const TOTAL_COL As Integer = 6

Private Const DISC_COL As Integer = 6
Private Const TAX_COL As Integer = 6

Private xlSheet As Excel.Worksheet

Private intRow As Integer

Private Sub Class_Initialize()
    mdlExcel.OpenExcel
End Sub

Public Sub ImportToExcel(ByVal strSOId As String)
    If Trim(strSOId) = "" Then Exit Sub
    
    mdlExcel.OpenWorkSheet xlSheet
    
    SetWidth
    
    FillTitle strSOId
    
    FillDetail strSOId
End Sub

Private Sub SetWidth()
    intRow = 1
    
    With xlSheet
        .PageSetup.LeftMargin = 25
        .PageSetup.RightMargin = 25
        
        .Columns("A").ColumnWidth = 11
        .Columns("B").ColumnWidth = 1
        .Columns("C").ColumnWidth = 45
        .Columns("D").ColumnWidth = 10
        .Columns("E").ColumnWidth = 12
        .Columns("F").ColumnWidth = 15
        
        .Rows(intRow).RowHeight = 125
    End With
End Sub

Private Sub FillTitle(ByVal strSOId As String)
    On Local Error GoTo ErrHandler
    
    With xlSheet
        .Shapes.AddPicture mdlGlobal.strLogoImageText, True, True, 0, 0, 100, 100
        
        intRow = intRow + 1
        
        .Range("A" & intRow & ":C" & intRow + 5).Font.Name = "Courier New"
        
        .Cells(intRow, SOID_COL).Value = "NO. SO"
        .Cells(intRow, SOID_COL + 1).Value = ":"
        .Cells(intRow, SOID_COL + 2).Value = strSOId
        
        intRow = intRow + 1
        
        .Cells(intRow, DATE_COL).Value = "TGL."
        .Cells(intRow, DATE_COL + 1).Value = ":"
        .Cells(intRow, DATE_COL + 2).HorizontalAlignment = xlHAlignLeft
        .Cells(intRow, DATE_COL + 2).Value = mdlProcedures.FormatDate(mdlDatabase.GetFieldData(mdlGlobal.conInventory, "SODate", mdlTable.CreateTHSOSELL, "SOId='" & strSOId & "'"))
        .Cells(intRow, DATE_COL + 2).NumberFormat = "dd MMMM yyyy"
        
        intRow = intRow + 1
        
        Dim strPOId As String
        
        strPOId = mdlDatabase.GetFieldData(mdlGlobal.conInventory, "POId", mdlTable.CreateTHSOSELL, "SOId='" & strSOId & "'")
        
        .Cells(intRow, POID_COL).Value = "NO. PO"
        .Cells(intRow, POID_COL + 1).Value = ":"
        .Cells(intRow, POID_COL + 2).Value = strPOId
        
        intRow = intRow + 1
        
        .Cells(intRow, CUSTOMERNAME_COL).Value = "CUSTOMER"
        .Cells(intRow, CUSTOMERNAME_COL + 1).Value = ":"
        .Cells(intRow, CUSTOMERNAME_COL + 2).Value = "'" & _
            mdlDatabase.GetFieldData(mdlGlobal.conInventory, "Name", mdlTable.CreateTMCUSTOMER, "CustomerId='" & _
                mdlDatabase.GetFieldData(mdlGlobal.conInventory, "CustomerId", mdlTable.CreateTHPOSELL, "POId='" & strPOId & "'") & "'")
        
        intRow = intRow + 2
        
        .Cells(intRow, CURRENCYNAME_COL).Value = "MATA UANG"
        .Cells(intRow, CURRENCYNAME_COL + 1).Value = ":"
        .Cells(intRow, CURRENCYNAME_COL + 2).Value = "'" & _
            mdlDatabase.GetFieldData(mdlGlobal.conInventory, "Name", mdlTable.CreateTMCURRENCY, "CurrencyId='" & _
                mdlDatabase.GetFieldData(mdlGlobal.conInventory, "CurrencyId", mdlTable.CreateTHSOSELL, "SOId='" & strSOId & "'") & "'")
        
        intRow = intRow + 1
        
        .Rows(intRow + 1).RowHeight = 3
        
        .Range("A" & intRow & ":F" & intRow).Font.Bold = True
        .Range("B" & intRow & ":C" & intRow).MergeCells = True
        .Range("A" & intRow & ":F" & intRow).HorizontalAlignment = xlHAlignCenter
        
        mdlExcel.FillEdge xlSheet, "A" & intRow & ":F" & intRow + 1
        
        .Cells(intRow, ITEMID_COL).Value = "Kode"
        .Cells(intRow, NAME_COL).Value = "Nama"
        .Cells(intRow, QTY_COL).Value = "Qty"
        .Cells(intRow, ITEMPRICE_COL).Value = "Harga"
        .Cells(intRow, TOTAL_COL).Value = "Total"
        
        intRow = intRow + 1
        
        .Rows(intRow).RowHeight = 3
        
        .Range("A" & intRow & ":F" & intRow).MergeCells = True
    End With
    
ErrHandler:
End Sub

Private Sub FillDetail(Optional ByVal strSOId As String = "")
    On Local Error GoTo ErrHandler
    
    Dim rstDetail As ADODB.Recordset
    
    Set rstDetail = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "*", mdlTable.CreateTDSOSELL, False, "SOId='" & strSOId & "'", "ItemId ASC")
    
    With rstDetail
        Dim intTemp As Integer
        
        If .RecordCount > 0 Then
            intTemp = intRow + 1
        End If
        
        While Not .EOF
            intRow = intRow + 1
            
            xlSheet.Range("A" & intRow & ":B" & intRow).HorizontalAlignment = xlHAlignLeft
            xlSheet.Range("D" & intRow & ":F" & intRow).HorizontalAlignment = xlHAlignRight
            
            xlSheet.Range("B" & intRow & ":C" & intRow).MergeCells = True
            
            xlSheet.Cells(intRow, ITEMID_COL).Value = "'" & Trim(.Fields("ItemId").Value)
            xlSheet.Cells(intRow, NAME_COL).Value = "'" & Trim(mdlDatabase.GetFieldData(mdlGlobal.conInventory, "Name", mdlTable.CreateTMITEM, "ItemId='" & .Fields("ItemId").Value & "'"))
            xlSheet.Cells(intRow, QTY_COL).Value = CStr(.Fields("Qty").Value)
            xlSheet.Cells(intRow, QTY_COL).NumberFormat = "#,##0"
            xlSheet.Cells(intRow, ITEMPRICE_COL).Value = CStr(mdlTransaction.ConvertCurrency( _
                mdlDatabase.GetFieldData(mdlGlobal.conInventory, "CurrencyId", mdlTable.CreateTMITEMPRICE, "PriceId='" & .Fields("PriceId").Value & "'"), _
                mdlDatabase.GetFieldData(mdlGlobal.conInventory, "CurrencyId", mdlTable.CreateTHSOSELL, "SOId='" & strSOId & "'"), _
                mdlDatabase.GetFieldData(mdlGlobal.conInventory, "ItemPrice", mdlTable.CreateTMITEMPRICE, "PriceId='" & .Fields("PriceId") & "'")))
            xlSheet.Cells(intRow, ITEMPRICE_COL).NumberFormat = "#,##0"
            xlSheet.Cells(intRow, TOTAL_COL).Value = "=D" & intRow & "*E" & intRow
            xlSheet.Cells(intRow, TOTAL_COL).NumberFormat = "#,##0"
            
            .MoveNext
        Wend
        
        If .RecordCount > 0 Then
            xlSheet.Range("B" & intRow + 1 & ":C" & intRow + 1).MergeCells = True
            
            intRow = intRow + 2
            
            mdlExcel.FillEdge xlSheet, "A" & intRow & ":F" & intRow, , False, False, False, False, False, xlDouble
            
            xlSheet.Range("D" & intRow & ":E" & intRow).Font.Bold = True
            xlSheet.Range("D" & intRow & ":F" & intRow).HorizontalAlignment = xlHAlignRight
            
            xlSheet.Range("B" & intRow & ":C" & intRow).MergeCells = True

            xlSheet.Cells(intRow, ITEMPRICE_COL).Value = "SUBTOTAL"

            xlSheet.Cells(intRow, TOTAL_COL).Value = "=SUM(F" & intTemp & ":F" & intRow - 2 & ")"
            xlSheet.Cells(intRow, TOTAL_COL).NumberFormat = "#,##0"

            intRow = intRow + 1
            
            Dim curDisc As Currency
            
            curDisc = mdlProcedures.GetCurrency(mdlDatabase.GetFieldData(mdlGlobal.conInventory, "Disc", mdlTable.CreateTHSOSELL, "SOId='" & strSOId & "'"))

            xlSheet.Range("D" & intRow & ":E" & intRow).Font.Bold = True
            xlSheet.Range("D" & intRow & ":F" & intRow).HorizontalAlignment = xlHAlignRight
            
            xlSheet.Range("B" & intRow & ":C" & intRow).MergeCells = True

            xlSheet.Cells(intRow, ITEMPRICE_COL).Value = "DISC"

            xlSheet.Cells(intRow, TOTAL_COL).Value = "=(F" & intRow - 1 & "*" & curDisc & ")/100"
            xlSheet.Cells(intRow, TOTAL_COL).NumberFormat = "#,##0"

            intRow = intRow + 1
            
            mdlExcel.FillEdge xlSheet, "F" & intRow & ":F" & intRow, , False, False, False, False, False, xlDouble

            xlSheet.Range("D" & intRow & ":F" & intRow).HorizontalAlignment = xlHAlignRight
            
            xlSheet.Range("B" & intRow & ":C" & intRow).MergeCells = True

            xlSheet.Cells(intRow, TOTAL_COL).Value = "=F" & intRow - 2 & "-F" & intRow - 1
            xlSheet.Cells(intRow, TOTAL_COL).NumberFormat = "#,##0"

            intRow = intRow + 1
            
            Dim curTax As Currency
            
            curTax = mdlProcedures.GetCurrency(mdlDatabase.GetFieldData(mdlGlobal.conInventory, "Tax", mdlTable.CreateTHSOSELL, "SOId='" & strSOId & "'"))

            xlSheet.Range("D" & intRow & ":E" & intRow).Font.Bold = True
            xlSheet.Range("D" & intRow & ":E" & intRow).HorizontalAlignment = xlHAlignRight
            
            xlSheet.Range("B" & intRow & ":C" & intRow).MergeCells = True

            xlSheet.Cells(intRow, ITEMPRICE_COL).Value = "PPN"

            xlSheet.Cells(intRow, TOTAL_COL).Value = "=(F" & intRow - 1 & "*" & curTax & ")/100"
            xlSheet.Cells(intRow, TOTAL_COL).NumberFormat = "#,##0"

            intRow = intRow + 1

            mdlExcel.FillEdge xlSheet, "F" & intRow & ":F" & intRow, , False, False, False, False, False, xlDouble
            
            xlSheet.Range("D" & intRow & ":F" & intRow).Font.Bold = True
            xlSheet.Range("D" & intRow & ":F" & intRow).HorizontalAlignment = xlRight
            
            xlSheet.Range("B" & intRow & ":C" & intRow).MergeCells = True

            xlSheet.Cells(intRow, ITEMPRICE_COL).Value = "TOTAL"

            xlSheet.Cells(intRow, TOTAL_COL).Value = "=F" & intRow - 2 & "+F" & intRow - 1
            xlSheet.Cells(intRow, TOTAL_COL).NumberFormat = "#,##0"
            
            mdlExcel.FillEdge xlSheet, "A" & intTemp & ":F" & intRow, False, , False
            
            xlSheet.Range("A" & intTemp & ":F" & intTemp).Select
        End If
    End With
    
ErrHandler:
End Sub

Private Sub Class_Terminate()
    mdlExcel.CloseExcel
End Sub
