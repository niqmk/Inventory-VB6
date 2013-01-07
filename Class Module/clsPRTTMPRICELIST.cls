VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPRTTMPRICELIST"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const ITEMID_COL As Integer = 1
Private Const NAME_COL As Integer = 2
Private Const PRICELISTDATE_COL As Integer = 3
Private Const PRICELISTVALUE_COL As Integer = 4
Private Const CURRENCYNAME_COL As Integer = 5

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
        .Columns("B").ColumnWidth = 25
        .Columns("C").ColumnWidth = 17
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
        
        mdlExcel.FillEdge xlSheet, "A" & intRow & ":E" & intRow + 1
        
        .Range("A" & intRow & ":E" & intRow).Font.Bold = True
        
        .Cells(intRow, ITEMID_COL).Value = "Kode"
        .Cells(intRow, NAME_COL).Value = "Nama"
        .Cells(intRow, PRICELISTDATE_COL).Value = "Tanggal"
        .Cells(intRow, PRICELISTVALUE_COL).Value = "Harga"
        .Cells(intRow, CURRENCYNAME_COL).Value = "Mata Uang"
    End With
    
ErrHandler:
End Sub

Private Sub FillText(ByRef rstMain As ADODB.Recordset)
    On Local Error GoTo ErrHandler
    
    With rstMain
        Dim intFirst As Integer
        Dim intTemp As Integer
        
        Dim strItemId As String
        
        intFirst = intRow + 1
        intTemp = 0
        
        While Not .EOF
            intRow = intRow + 1
            
            xlSheet.Range("A" & intRow & ":E" & intRow).VerticalAlignment = xlTop
            xlSheet.Range("A" & intRow & ":C" & intRow).HorizontalAlignment = xlLeft
            xlSheet.Range("D" & intRow & ":D" & intRow).HorizontalAlignment = xlRight
            xlSheet.Range("E" & intRow & ":E" & intRow).HorizontalAlignment = xlLeft
            xlSheet.Range("A" & intRow & ":E" & intRow).WrapText = True
            
            If Trim(strItemId) = Trim(.Fields("ItemId").Value) Then
                xlSheet.Range("A" & intTemp & ":A" & intRow).MergeCells = True
                xlSheet.Range("B" & intTemp & ":B" & intRow).MergeCells = True
            Else
                If Not intTemp = 0 Then
                    mdlExcel.FillEdge xlSheet, "A" & intTemp & ":E" & intRow
                End If
                
                strItemId = Trim(.Fields("ItemId").Value)
                
                xlSheet.Cells(intRow, ITEMID_COL).Value = "'" & strItemId
                xlSheet.Cells(intRow, NAME_COL).Value = "'" & Trim(mdlDatabase.GetFieldData(mdlGlobal.conInventory, "Name", mdlTable.CreateTMITEM, "ItemId='" & strItemId & "'"))
                
                intTemp = intRow
            End If
            
            xlSheet.Cells(intRow, PRICELISTDATE_COL).Value = "'" & mdlProcedures.FormatDate(!PriceListDate, "dd-MMMM-yyyy")
            xlSheet.Cells(intRow, PRICELISTVALUE_COL).Value = rstMain.Fields("PriceListValue").Value
            xlSheet.Cells(intRow, PRICELISTVALUE_COL).NumberFormat = "#,##0"
            xlSheet.Cells(intRow, CURRENCYNAME_COL).Value = "'" & Trim(mdlDatabase.GetFieldData(mdlGlobal.conInventory, "Name", mdlTable.CreateTMCURRENCY, "CurrencyId='" & .Fields("CurrencyId").Value & "'"))
            
            xlSheet.Range("C" & intRow & ":E" & intRow).Select
            
            .MoveNext
        Wend
        
        If .RecordCount > 0 Then
            mdlExcel.FillEdge xlSheet, "A" & intTemp & ":E" & intRow, False
            
            xlSheet.Range("A" & intFirst & ":E" & intFirst).Select
        End If
    End With
    
ErrHandler:
End Sub

Private Sub Class_Terminate()
    mdlExcel.CloseExcel
End Sub
