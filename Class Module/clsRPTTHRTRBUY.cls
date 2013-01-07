VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRPTTHRTRBUY"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const RTRID_COL As Integer = 1
Private Const DATE_COL As Integer = 2
Private Const FKTID_COL As Integer = 3
Private Const FKTDATE_COL As Integer = 4

Private Const ITEMID_COL As Integer = 2
Private Const NAME_COL As Integer = 3
Private Const QTY_COL As Integer = 4
Private Const UNITYNAME_COL As Integer = 5

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
        .Columns("C").ColumnWidth = 24
        .Columns("D").ColumnWidth = 10
        .Columns("E").ColumnWidth = 23
        
        .Rows(intRow).RowHeight = 125
    End With
End Sub

Private Sub FillTitle()
    On Local Error GoTo ErrHandler
    
    With xlSheet
        .Shapes.AddPicture mdlGlobal.strLogoImageText, True, True, 0, 0, 100, 100
        
        intRow = intRow + 2
    
        .Range("A" & intRow & ":A" & intRow + 1).MergeCells = True
        .Range("D" & intRow & ":E" & intRow).MergeCells = True
        .Range("A" & intRow & ":E" & intRow + 1).HorizontalAlignment = xlLeft
        .Range("A" & intRow & ":E" & intRow + 1).VerticalAlignment = xlTop
        .Range("A" & intRow & ":E" & intRow + 1).Font.Bold = True
        
        mdlExcel.FillEdge xlSheet, "A" & intRow & ":E" & intRow + 1
    
        .Cells(intRow, RTRID_COL).Value = "Nomor Retur"
        .Cells(intRow, DATE_COL).Value = "Tanggal"
        .Cells(intRow, FKTID_COL).Value = "Nomor Faktur"
        .Cells(intRow, FKTDATE_COL).Value = "Tanggal Faktur"
        
        intRow = intRow + 1
        
        .Cells(intRow, ITEMID_COL).Value = "Kode"
        .Cells(intRow, NAME_COL).Value = "Nama"
        .Cells(intRow, QTY_COL).Value = "Qty"
        .Cells(intRow, UNITYNAME_COL).Value = "Satuan"
    End With
    
ErrHandler:
End Sub

Private Sub FillHeader(ByRef rstHeader As ADODB.Recordset)
    On Local Error GoTo ErrHandler
    
    With rstHeader
        Dim intFirst As Integer
        Dim intTemp As Integer
        
        Dim strFktId As String
        
        intFirst = intRow + 1
    
        While Not .EOF
            If Not intRow = intFirst - 1 Then
                intRow = intRow + 1
                
                xlSheet.Rows(intRow).RowHeight = 5
                mdlExcel.FillEdge xlSheet, "A" & intRow & ":E" & intRow, False, , False, False, False, False
            End If
            
            intRow = intRow + 1
            
            intTemp = intRow
            
            xlSheet.Range("D" & intRow & ":E" & intRow).MergeCells = True
            xlSheet.Range("A" & intRow & ":D" & intRow).HorizontalAlignment = xlHAlignLeft
            xlSheet.Range("A" & intRow & ":D" & intRow).VerticalAlignment = xlVAlignTop
            
            xlSheet.Cells(intRow, RTRID_COL).Value = "'" & Trim(.Fields("RtrId").Value)
            xlSheet.Cells(intRow, DATE_COL).Value = mdlProcedures.FormatDate(.Fields("RtrDate").Value)
            xlSheet.Cells(intRow, DATE_COL).NumberFormat = "dd MMMM yyyy"
            
            strFktId = mdlDatabase.GetFieldData(mdlGlobal.conInventory, "FktId", mdlTable.CreateTDFKTBUY, "SJId='" & .Fields("SJId").Value & "'")
            
            xlSheet.Cells(intRow, FKTID_COL).Value = strFktId
            xlSheet.Cells(intRow, FKTDATE_COL).Value = mdlProcedures.FormatDate(mdlDatabase.GetFieldData(mdlGlobal.conInventory, "FktDate", mdlTable.CreateTHFKTBUY, "FktId='" & strFktId & "'"))
            xlSheet.Cells(intRow, FKTDATE_COL).NumberFormat = "dd MMMM yyyy"
            
            xlSheet.Range("A" & intRow & ":E" & intRow).Select
            
            FillDetail .Fields("RtrId").Value
            
            xlSheet.Range("A" & intTemp & ":E" & intRow).WrapText = True
            xlSheet.Range("A" & intTemp & ":A" & intRow).MergeCells = True
            
            mdlExcel.FillEdge xlSheet, "A" & intTemp & ":E" & intRow, False, , False
            
            .MoveNext
        Wend
        
        If .RecordCount > 0 Then
            xlSheet.Range("A" & intFirst & ":E" & intFirst).Select
        End If
    End With
    
ErrHandler:
End Sub

Private Sub FillDetail(Optional ByVal strRtrId As String = "")
    On Local Error GoTo ErrHandler
    
    Dim rstDetail As ADODB.Recordset
    
    Set rstDetail = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "*", mdlTable.CreateTDRTRBUY, False, "RtrId='" & strRtrId & "'", "ItemId ASC")
    
    With rstDetail
        If .RecordCount > 0 Then
            xlSheet.Range("B" & intRow & ":E" & intRow).Borders(xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        End If
        
        While Not .EOF
            intRow = intRow + 1
            
            xlSheet.Range("B" & intRow & ":C" & intRow).HorizontalAlignment = xlHAlignLeft
            xlSheet.Range("D" & intRow & ":D" & intRow).HorizontalAlignment = xlHAlignRight
            xlSheet.Range("E" & intRow & ":E" & intRow).HorizontalAlignment = xlHAlignLeft
            
            xlSheet.Cells(intRow, ITEMID_COL).Value = "'" & Trim(.Fields("ItemId").Value)
            xlSheet.Cells(intRow, NAME_COL).Value = "'" & Trim(mdlDatabase.GetFieldData(mdlGlobal.conInventory, "Name", mdlTable.CreateTMITEM, "ItemId='" & .Fields("ItemId").Value & "'"))
            xlSheet.Cells(intRow, QTY_COL).Value = CStr(.Fields("Qty").Value)
            xlSheet.Cells(intRow, QTY_COL).NumberFormat = "#,##0"
            xlSheet.Cells(intRow, UNITYNAME_COL).Value = "'" & Trim( _
                mdlDatabase.GetFieldData(mdlGlobal.conInventory, "Name", mdlTable.CreateTMUNITY, "UnityId='" & _
                    mdlDatabase.GetFieldData(mdlGlobal.conInventory, "UnityId", mdlTable.CreateTMITEM, "ItemId='" & .Fields("ItemId").Value & "'") & "'"))
            
            .MoveNext
        Wend
    End With
    
ErrHandler:
End Sub

Private Sub Class_Terminate()
    mdlExcel.CloseExcel
End Sub
