VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRPTTHSOSELL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const SOID_COL As Integer = 1
Private Const DATE_COL As Integer = 2
Private Const POID_COL As Integer = 3
Private Const DISC_COL As Integer = 4
Private Const TAX_COL As Integer = 5
Private Const CURRENCYNAME_COL As Integer = 6

Private Const ITEMID_COL As Integer = 2
Private Const NAME_COL As Integer = 3
Private Const QTY_COL As Integer = 4
Private Const ITEMPRICE_COL As Integer = 5
Private Const TOTAL_COL As Integer = 6

Private xlSheet As Excel.Worksheet

Private intRow As Integer

Private Sub Class_Initialize()
    mdlExcel.OpenExcel
End Sub

Public Sub ImportToExcel(ByRef rstHeader As ADODB.Recordset)
    If rstHeader Is Nothing Then Exit Sub
    
    mdlExcel.OpenWorkSheet xlSheet
    
    SetWidth
    
    FillTitle
    
    FillHeader rstHeader
End Sub

Private Sub SetWidth()
    intRow = 1
    
    With xlSheet
        .PageSetup.LeftMargin = 25
        .PageSetup.RightMargin = 25
        
        .Columns("A").ColumnWidth = 20
        .Columns("B").ColumnWidth = 18
        .Columns("C").ColumnWidth = 20
        .Columns("D").ColumnWidth = 10
        .Columns("E").ColumnWidth = 10
        .Columns("F").ColumnWidth = 15
        
        .Rows(intRow).RowHeight = 125
    End With
End Sub

Private Sub FillTitle()
    On Local Error GoTo ErrHandler
    
    With xlSheet
        .Shapes.AddPicture mdlGlobal.strLogoImageText, True, True, 0, 0, 100, 100
        
        intRow = intRow + 2
    
        .Range("A" & intRow & ":A" & intRow + 1).MergeCells = True
        .Range("A" & intRow & ":F" & intRow + 1).HorizontalAlignment = xlLeft
        .Range("A" & intRow & ":F" & intRow + 1).VerticalAlignment = xlTop
        .Range("A" & intRow & ":F" & intRow + 1).Font.Bold = True
        
        mdlExcel.FillEdge xlSheet, "A" & intRow & ":F" & intRow + 1
    
        .Cells(intRow, SOID_COL).Value = "Nomor SO"
        .Cells(intRow, DATE_COL).Value = "Tanggal"
        .Cells(intRow, POID_COL).Value = "Nomor PO"
        .Cells(intRow, DISC_COL).Value = "Disc (%)"
        .Cells(intRow, TAX_COL).Value = "PPN (%)"
        .Cells(intRow, CURRENCYNAME_COL).Value = "Mata Uang"
        
        intRow = intRow + 1
        
        .Cells(intRow, ITEMID_COL).Value = "Kode"
        .Cells(intRow, NAME_COL).Value = "Nama"
        .Cells(intRow, QTY_COL).Value = "Qty"
        .Cells(intRow, ITEMPRICE_COL).Value = "Harga"
        .Cells(intRow, TOTAL_COL).Value = "Total"
    End With
    
ErrHandler:
End Sub

Private Sub FillHeader(ByRef rstHeader As ADODB.Recordset)
    On Local Error GoTo ErrHandler
    
    With rstHeader
        Dim intFirst As Integer
        Dim intTemp As Integer
        
        intFirst = intRow + 1
    
        While Not .EOF
            If Not intRow = intFirst - 1 Then
                intRow = intRow + 1
                
                xlSheet.Rows(intRow).RowHeight = 5
                mdlExcel.FillEdge xlSheet, "A" & intRow & ":F" & intRow, False, , False, False, False, False
            End If
            
            intRow = intRow + 1
            
            intTemp = intRow
            
            xlSheet.Range("A" & intRow & ":C" & intRow).HorizontalAlignment = xlHAlignLeft
            xlSheet.Range("D" & intRow & ":E" & intRow).HorizontalAlignment = xlHAlignRight
            xlSheet.Range("F" & intRow & ":F" & intRow).HorizontalAlignment = xlHAlignLeft
            xlSheet.Range("A" & intRow & ":F" & intRow).VerticalAlignment = xlVAlignTop
            
            xlSheet.Cells(intRow, SOID_COL).Value = "'" & Trim(.Fields("SOId").Value)
            xlSheet.Cells(intRow, DATE_COL).Value = mdlProcedures.FormatDate(.Fields("SODate").Value)
            xlSheet.Cells(intRow, DATE_COL).NumberFormat = "dd MMMM yyyy"
            xlSheet.Cells(intRow, POID_COL).Value = "'" & Trim(.Fields("POId").Value)
            xlSheet.Cells(intRow, DISC_COL).Value = CStr(.Fields("Disc").Value)
            xlSheet.Cells(intRow, TAX_COL).Value = CStr(.Fields("Tax").Value)
            xlSheet.Cells(intRow, CURRENCYNAME_COL).Value = "'" & Trim(mdlDatabase.GetFieldData(mdlGlobal.conInventory, "Name", mdlTable.CreateTMCURRENCY, "CurrencyId='" & .Fields("CurrencyId").Value & "'"))
            
            xlSheet.Range("A" & intRow & ":F" & intRow).Select
            
            FillDetail .Fields("SOId").Value
            
            xlSheet.Range("A" & intTemp & ":F" & intRow).VerticalAlignment = xlTop
            xlSheet.Range("A" & intTemp & ":F" & intRow).WrapText = True
            xlSheet.Range("A" & intTemp & ":A" & intRow).MergeCells = True
            
            mdlExcel.FillEdge xlSheet, "A" & intTemp & ":F" & intRow, False, , False
            
            .MoveNext
        Wend
        
        If .RecordCount > 0 Then
            xlSheet.Range("A" & intFirst & ":F" & intFirst).Select
        End If
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
            
            xlSheet.Range("B" & intRow & ":F" & intRow).Borders(xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        End If
        
        While Not .EOF
            intRow = intRow + 1
            
            xlSheet.Range("B" & intRow & ":C" & intRow).HorizontalAlignment = xlLeft
            xlSheet.Range("D" & intRow & ":F" & intRow).HorizontalAlignment = xlRight
            
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
            intRow = intRow + 1
            
            xlSheet.Range("B" & intRow & ":F" & intRow).Borders(xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
            xlSheet.Range("B" & intRow & ":E" & intRow).MergeCells = True
            xlSheet.Range("B" & intRow).Font.Bold = True
            xlSheet.Range("B" & intRow & ":F" & intRow).HorizontalAlignment = xlRight
            
            xlSheet.Cells(intRow, ITEMID_COL).Value = "SUBTOTAL"
            
            xlSheet.Cells(intRow, TOTAL_COL).Value = "=SUM(F" & intTemp & ":F" & intRow - 1 & ")"
            xlSheet.Cells(intRow, TOTAL_COL).NumberFormat = "#,##0"
            
            intRow = intRow + 1
            
            xlSheet.Range("B" & intRow & ":F" & intRow).Borders(xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
            xlSheet.Range("B" & intRow & ":E" & intRow).MergeCells = True
            xlSheet.Range("B" & intRow).Font.Bold = True
            xlSheet.Range("B" & intRow & ":F" & intRow).HorizontalAlignment = xlRight
            
            xlSheet.Cells(intRow, ITEMID_COL).Value = "DISC"
            
            xlSheet.Cells(intRow, TOTAL_COL).Value = "=(F" & intRow - 1 & "*D" & intTemp - 1 & ")/100"
            xlSheet.Cells(intRow, TOTAL_COL).NumberFormat = "#,##0"
            
            intRow = intRow + 1
            
            xlSheet.Range("B" & intRow & ":F" & intRow).Borders(xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
            xlSheet.Range("B" & intRow & ":E" & intRow).MergeCells = True
            xlSheet.Range("B" & intRow & ":F" & intRow).HorizontalAlignment = xlRight
            
            xlSheet.Cells(intRow, TOTAL_COL).Value = "=F" & intRow - 2 & "-F" & intRow - 1
            xlSheet.Cells(intRow, TOTAL_COL).NumberFormat = "#,##0"
            
            intRow = intRow + 1
            
            xlSheet.Range("B" & intRow & ":F" & intRow).Borders(xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
            xlSheet.Range("B" & intRow & ":E" & intRow).MergeCells = True
            xlSheet.Range("B" & intRow).Font.Bold = True
            xlSheet.Range("B" & intRow & ":F" & intRow).HorizontalAlignment = xlRight
            
            xlSheet.Cells(intRow, ITEMID_COL).Value = "PPN"
            
            xlSheet.Cells(intRow, TOTAL_COL).Value = "=(F" & intRow - 1 & "*E" & intTemp - 1 & ")/100"
            xlSheet.Cells(intRow, TOTAL_COL).NumberFormat = "#,##0"
            
            intRow = intRow + 1
            
            xlSheet.Range("B" & intRow & ":F" & intRow).Borders(xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
            xlSheet.Range("B" & intRow & ":E" & intRow).MergeCells = True
            xlSheet.Range("B" & intRow).Font.Bold = True
            xlSheet.Range("B" & intRow & ":F" & intRow).HorizontalAlignment = xlRight
            
            xlSheet.Cells(intRow, ITEMID_COL).Value = "TOTAL"
            
            xlSheet.Cells(intRow, TOTAL_COL).Value = "=F" & intRow - 2 & "+F" & intRow - 1
            xlSheet.Cells(intRow, TOTAL_COL).NumberFormat = "#,##0"
        End If
    End With
    
ErrHandler:
End Sub

Private Sub Class_Terminate()
    mdlExcel.CloseExcel
End Sub
