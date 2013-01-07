VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPRTTMVENDOR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const VENDORID_COL As Integer = 1
Private Const NAME_COL As Integer = 2
Private Const ADDRESS_COL As Integer = 3
Private Const FAX_COL As Integer = 5

Private Const CONTNAME_COL As Integer = 2
Private Const CONTPHONE_COL As Integer = 3
Private Const CONTHANDPHONE_COL As Integer = 4
Private Const CONTEMAIL_COL As Integer = 5

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
        .Columns("A").ColumnWidth = 12
        .Columns("B").ColumnWidth = 15
        .Columns("C").ColumnWidth = 15
        .Columns("D").ColumnWidth = 15
        .Columns("E").ColumnWidth = 15
        
        .Rows(intRow).RowHeight = 125
    End With
End Sub

Private Sub FillTitle()
    On Local Error GoTo ErrHandler
    
    With xlSheet
        .Shapes.AddPicture mdlGlobal.strLogoImageText, True, True, 0, 0, 100, 100
        
        intRow = intRow + 1
    
        .Range("A" & intRow & ":A" & intRow + 1).MergeCells = True
        .Range("C" & intRow & ":D" & intRow).MergeCells = True
        .Range("A" & intRow).VerticalAlignment = xlTop
        
        mdlExcel.FillEdge xlSheet, "A" & intRow & ":E" & intRow + 1
        
        .Range("A" & intRow & ":E" & intRow).Font.Bold = True
    
        .Cells(intRow, VENDORID_COL).Value = "Kode"
        .Cells(intRow, NAME_COL).Value = "Nama"
        .Cells(intRow, ADDRESS_COL).Value = "Alamat"
        .Cells(intRow, FAX_COL).Value = "Fax"
        
        intRow = intRow + 1
        
        .Range("A" & intRow & ":E" & intRow).Font.Bold = True
        
        .Cells(intRow, CONTNAME_COL).Value = "Nama Kontak"
        .Cells(intRow, CONTPHONE_COL).Value = "Telepon"
        .Cells(intRow, CONTHANDPHONE_COL).Value = "HandPhone"
        .Cells(intRow, CONTEMAIL_COL).Value = "Email"
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
            intRow = intRow + 1
            
            intTemp = intRow
            
            xlSheet.Range("C" & intRow & ":D" & intRow).MergeCells = True
            
            xlSheet.Rows(intRow).RowHeight = 52
            
            xlSheet.Cells(intRow, VENDORID_COL).Value = "'" & Trim(.Fields("VendorId").Value)
            xlSheet.Cells(intRow, NAME_COL).Value = "'" & Trim(.Fields("Name").Value)
            xlSheet.Cells(intRow, ADDRESS_COL).Value = "'" & Trim(.Fields("Address").Value)
            xlSheet.Cells(intRow, FAX_COL).Value = "'" & Trim(.Fields("Fax").Value)
            
            xlSheet.Range("A" & intRow & ":E" & intRow).Select
            
            FillDetail .Fields("VendorId").Value
            
            xlSheet.Range("A" & intTemp & ":E" & intRow).VerticalAlignment = xlTop
            xlSheet.Range("A" & intTemp & ":E" & intRow).HorizontalAlignment = xlLeft
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

Private Sub FillDetail(Optional ByVal strVendorId As String = "")
    On Local Error GoTo ErrHandler
    
    Dim rstDetail As ADODB.Recordset
    
    Set rstDetail = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "*", mdlTable.CreateTMCONTACTVENDOR, False, "VendorId='" & strVendorId & "'", "Name ASC")
    
    With rstDetail
        If .RecordCount > 0 Then
            xlSheet.Range("B" & intRow & ":E" & intRow).Borders(xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        End If
        
        While Not .EOF
            intRow = intRow + 1
            
            xlSheet.Range("B" & intRow & ":E" & intRow).VerticalAlignment = xlTop
            xlSheet.Range("B" & intRow & ":E" & intRow).HorizontalAlignment = xlLeft
            
            xlSheet.Cells(intRow, CONTNAME_COL).Value = "'" & Trim(.Fields("Name").Value)
            xlSheet.Cells(intRow, CONTPHONE_COL).Value = "'" & Trim(.Fields("Phone").Value)
            xlSheet.Cells(intRow, CONTHANDPHONE_COL).Value = "'" & Trim(.Fields("HandPhone").Value)
            xlSheet.Cells(intRow, CONTEMAIL_COL).Value = "'" & Trim(.Fields("Email").Value)
        
            .MoveNext
        Wend
    End With
    
ErrHandler:
End Sub

Private Sub Class_Terminate()
    mdlExcel.CloseExcel
End Sub
