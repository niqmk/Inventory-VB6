VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPRTTHFKTBUY"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const FKTID_COL As Integer = 1
Private Const DATE_COL As Integer = 1
Private Const VENDORNAME_COL As Integer = 1

Private Const SJID_COL As Integer = 1
Private Const SJDATE_COL As Integer = 2
Private Const QTY_COL As Integer = 4

Private xlSheet As Excel.Worksheet

Private intRow As Integer

Private Sub Class_Initialize()
    mdlExcel.OpenExcel
End Sub

Public Sub ImportToExcel(ByVal strFktId As String)
    If Trim(strFktId) = "" Then Exit Sub
    
    mdlExcel.OpenWorkSheet xlSheet
    
    SetWidth
    
    FillTitle strFktId
    
    FillDetail strFktId
End Sub

Private Sub SetWidth()
    intRow = 1
    
    With xlSheet
        .PageSetup.LeftMargin = 25
        .PageSetup.RightMargin = 25
        
        .Columns("A").ColumnWidth = 20
        .Columns("B").ColumnWidth = 1
        .Columns("C").ColumnWidth = 25
        .Columns("D").ColumnWidth = 15
        
        .Rows(intRow).RowHeight = 125
    End With
End Sub

Private Sub FillTitle(ByVal strFktId As String)
    On Local Error GoTo ErrHandler
    
    With xlSheet
        .Shapes.AddPicture mdlGlobal.strLogoImageText, True, True, 0, 0, 100, 100
        
        intRow = intRow + 1
        
        .Range("A" & intRow & ":C" & intRow + 2).Font.Name = "Courier New"
        
        .Cells(intRow, FKTID_COL).Value = "NO. FAKTUR"
        .Cells(intRow, FKTID_COL + 1).Value = ":"
        .Cells(intRow, FKTID_COL + 2).Value = strFktId
        
        intRow = intRow + 1
        
        .Cells(intRow, DATE_COL).Value = "TGL."
        .Cells(intRow, DATE_COL + 1).Value = ":"
        .Cells(intRow, DATE_COL + 2).HorizontalAlignment = xlHAlignLeft
        .Cells(intRow, DATE_COL + 2).Value = mdlProcedures.FormatDate(mdlDatabase.GetFieldData(mdlGlobal.conInventory, "FktDate", mdlTable.CreateTHFKTBUY, "FktId='" & strFktId & "'"))
        .Cells(intRow, DATE_COL + 2).NumberFormat = "dd MMMM yyyy"
        
        intRow = intRow + 1
        
        .Cells(intRow, VENDORNAME_COL).Value = "PEMASOK"
        .Cells(intRow, VENDORNAME_COL + 1).Value = ":"
        .Cells(intRow, VENDORNAME_COL + 2).Value = "'" & _
            mdlDatabase.GetFieldData(mdlGlobal.conInventory, "Name", mdlTable.CreateTMVENDOR, "VendorId='" & _
                mdlDatabase.GetFieldData(mdlGlobal.conInventory, "VendorId", mdlTable.CreateTHFKTBUY, "FktId='" & strFktId & "'") & "'")
        
        intRow = intRow + 2
        
        .Range("A" & intRow & ":D" & intRow).Font.Bold = True
        .Range("A" & intRow & ":D" & intRow).HorizontalAlignment = xlHAlignCenter
        
        .Range("B" & intRow & ":C" & intRow).MergeCells = True
        
        mdlExcel.FillEdge xlSheet, "A" & intRow & ":D" & intRow + 1
        
        .Cells(intRow, SJID_COL).Value = "Nomor SJ"
        .Cells(intRow, SJDATE_COL).Value = "Tanggal"
        .Cells(intRow, QTY_COL).Value = "Qty"
        
        intRow = intRow + 1
        
        .Rows(intRow).RowHeight = 3
        
        .Range("A" & intRow & ":D" & intRow).MergeCells = True
    End With
    
ErrHandler:
End Sub

Private Sub FillDetail(Optional ByVal strFktId As String = "")
    On Local Error GoTo ErrHandler
    
    Dim rstDetail As ADODB.Recordset
    
    Set rstDetail = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "*", mdlTable.CreateTDFKTBUY, False, "FktId='" & strFktId & "'", "SJId ASC")
    
    With rstDetail
        Dim intTemp As Integer
        
        Dim strSJId As String
        
        If .RecordCount > 0 Then
            intTemp = intRow + 1
        End If
        
        While Not .EOF
            intRow = intRow + 1
            
            xlSheet.Range("A" & intRow & ":C" & intRow).HorizontalAlignment = xlHAlignLeft
            xlSheet.Range("D" & intRow & ":D" & intRow).HorizontalAlignment = xlHAlignRight
            xlSheet.Range("A" & intRow & ":D" & intRow).VerticalAlignment = xlVAlignTop
            
            xlSheet.Range("B" & intRow & ":C" & intRow).MergeCells = True
            
            strSJId = Trim(.Fields("SJId").Value)
            
            xlSheet.Cells(intRow, SJID_COL).Value = "'" & strSJId
            xlSheet.Cells(intRow, SJDATE_COL).HorizontalAlignment = xlHAlignLeft
            xlSheet.Cells(intRow, SJDATE_COL).Value = mdlProcedures.FormatDate(mdlDatabase.GetFieldData(mdlGlobal.conInventory, "SJDate", mdlTable.CreateTHSJBUY, "SJId='" & strSJId & "'"))
            xlSheet.Cells(intRow, SJDATE_COL).NumberFormat = "dd MMMM yyyy"
            xlSheet.Cells(intRow, QTY_COL).Value = CStr(mdlTHSJBUY.GetTotalQtySJBUY(strSJId))
            xlSheet.Cells(intRow, QTY_COL).NumberFormat = "#,##0"
            
            xlSheet.Range("A" & intTemp & ":D" & intTemp).Select
            
            .MoveNext
        Wend
        
        If .RecordCount > 0 Then
            mdlExcel.FillEdge xlSheet, "A" & intTemp & ":D" & intRow, False, , False
            
            xlSheet.Range("A" & intTemp & ":D" & intTemp).Select
        End If
    End With
    
ErrHandler:
End Sub

Private Sub Class_Terminate()
    mdlExcel.CloseExcel
End Sub
