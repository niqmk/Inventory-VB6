VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRPTTHFKTSELL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const FKTID_COL As Integer = 1
Private Const DATE_COL As Integer = 2
Private Const CUSTOMERNAME_COL As Integer = 3

Private Const SJID_COL As Integer = 2
Private Const SJDATE_COL As Integer = 3
Private Const QTY_COL As Integer = 4

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
        .Columns("B").ColumnWidth = 20
        .Columns("C").ColumnWidth = 40
        .Columns("D").ColumnWidth = 15
        
        .Rows(intRow).RowHeight = 125
    End With
End Sub

Private Sub FillTitle()
    On Local Error GoTo ErrHandler
    
    With xlSheet
        .Shapes.AddPicture mdlGlobal.strLogoImageText, True, True, 0, 0, 100, 100
        
        intRow = intRow + 2
        
        .Range("A" & intRow & ":A" & intRow + 1).MergeCells = True
        .Range("C" & intRow & ":D" & intRow).MergeCells = True
        .Range("A" & intRow & ":D" & intRow + 1).HorizontalAlignment = xlLeft
        .Range("A" & intRow & ":D" & intRow + 1).VerticalAlignment = xlTop
        .Range("B" & intRow + 1 & ":D" & intRow + 1).WrapText = True
        .Range("A" & intRow & ":D" & intRow + 1).Font.Bold = True
        
        mdlExcel.FillEdge xlSheet, "A" & intRow & ":D" & intRow + 1
        
        .Cells(intRow, FKTID_COL).Value = "Nomor Faktur"
        .Cells(intRow, DATE_COL).Value = "Tanggal"
        .Cells(intRow, CUSTOMERNAME_COL).Value = "Pemasok"
        
        intRow = intRow + 1
        
        .Cells(intRow, SJID_COL).Value = "Nomor SJ"
        .Cells(intRow, SJDATE_COL).Value = "Tanggal SJ"
        .Cells(intRow, QTY_COL).Value = "Qty"
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
                mdlExcel.FillEdge xlSheet, "A" & intRow & ":D" & intRow, False, , False, False, False, False
            End If
            
            intRow = intRow + 1
            
            intTemp = intRow
            
            xlSheet.Range("C" & intRow & ":D" & intRow).MergeCells = True
            xlSheet.Range("A" & intRow & ":D" & intRow).HorizontalAlignment = xlHAlignLeft
            xlSheet.Range("A" & intRow & ":D" & intRow).VerticalAlignment = xlVAlignTop
            
            xlSheet.Cells(intRow, FKTID_COL).Value = "'" & Trim(.Fields("FktId").Value)
            xlSheet.Cells(intRow, DATE_COL).Value = mdlProcedures.FormatDate(.Fields("FktDate").Value)
            xlSheet.Cells(intRow, DATE_COL).NumberFormat = "dd MMMM yyyy"
            xlSheet.Cells(intRow, CUSTOMERNAME_COL).Value = "'" & Trim(mdlDatabase.GetFieldData(mdlGlobal.conInventory, "Name", mdlTable.CreateTMCUSTOMER, "CustomerId='" & .Fields("CustomerId").Value & "'"))
            
            xlSheet.Range("A" & intRow & ":D" & intRow).Select
            
            FillDetail .Fields("FktId").Value
            
            xlSheet.Range("A" & intTemp & ":A" & intRow).MergeCells = True
            
            mdlExcel.FillEdge xlSheet, "A" & intTemp & ":D" & intRow, False, , False
            
            .MoveNext
        Wend
        
        If .RecordCount > 0 Then
            xlSheet.Range("A" & intFirst & ":D" & intFirst).Select
        End If
    End With
    
ErrHandler:
End Sub

Private Sub FillDetail(Optional ByVal strFktId As String = "")
    On Local Error GoTo ErrHandler
    
    Dim rstDetail As ADODB.Recordset
    
    Set rstDetail = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "*", mdlTable.CreateTDFKTSELL, False, "FktId='" & strFktId & "'", "SJId ASC")
    
    With rstDetail
        If .RecordCount > 0 Then
            xlSheet.Range("B" & intRow & ":D" & intRow).Borders(xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        End If
        
        While Not .EOF
            intRow = intRow + 1
            
            xlSheet.Range("B" & intRow & ":D" & intRow).WrapText = True
            xlSheet.Range("B" & intRow & ":C" & intRow).HorizontalAlignment = xlHAlignLeft
            xlSheet.Range("D" & intRow & ":D" & intRow).HorizontalAlignment = xlHAlignRight
            
            xlSheet.Cells(intRow, SJID_COL).Value = Trim(.Fields("SJId").Value)
            xlSheet.Cells(intRow, SJDATE_COL).Value = mdlProcedures.FormatDate(mdlDatabase.GetFieldData(mdlGlobal.conInventory, "SJDate", mdlTable.CreateTHSJSELL, "SJId='" & .Fields("SJId").Value & "'"))
            xlSheet.Cells(intRow, SJDATE_COL).NumberFormat = "dd MMMM yyyy"
            xlSheet.Cells(intRow, QTY_COL).Value = CStr(mdlTHSJSELL.GetTotalQtySJSELL(!SJId))
            xlSheet.Cells(intRow, QTY_COL).NumberFormat = "#,##0"
            
            .MoveNext
        Wend
    End With
    
ErrHandler:
End Sub

Private Sub Class_Terminate()
    mdlExcel.CloseExcel
End Sub
