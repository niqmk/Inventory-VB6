VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPRTTHSALESSUM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const POID_COL As Integer = 1
Private Const DATE_COL As Integer = 2
Private Const POCUSTOMERID_COL As Integer = 3
Private Const CUSTOMERNAME_COL As Integer = 4
Private Const PRICEVALUE_COL As Integer = 5
Private Const CURRENCYID_COL As Integer = 6

Private xlSheet As Excel.Worksheet

Private intRow As Integer

Private Sub Class_Initialize()
    mdlExcel.OpenExcel
End Sub

Public Sub ImportToExcel(ByVal strPOId As String)
    If Trim(strPOId) = "" Then Exit Sub
    
    mdlExcel.OpenWorkSheet xlSheet
    
    SetWidth
    
    FillTitle strPOId
End Sub

Private Sub SetWidth()
    intRow = 1
    
    With xlSheet
        .PageSetup.Orientation = xlLandscape
        
        .PageSetup.LeftMargin = 25
        .PageSetup.RightMargin = 25
        
        .Columns("A").ColumnWidth = 20
        .Columns("B").ColumnWidth = 18
        .Columns("C").ColumnWidth = 21
        .Columns("D").ColumnWidth = 35
        .Columns("E").ColumnWidth = 10
        .Columns("F").ColumnWidth = 15
        
        .Rows(intRow).RowHeight = 125
    End With
End Sub

Private Sub FillTitle(ByVal strPOId As String)
    On Local Error GoTo ErrHandler
    
    With xlSheet
        .Shapes.AddPicture mdlGlobal.strLogoImageText, True, True, 0, 0, 100, 100
        
        intRow = intRow + 1
        
        .Rows(intRow).RowHeight = 30
        
        intRow = intRow + 1
        
        .Cells(intRow, POID_COL).Value = "NO. PO"
        .Cells(intRow, DATE_COL).Value = "TGL."
        .Cells(intRow, POCUSTOMERID_COL).Value = "CUSTOMER PO"
        .Cells(intRow, CUSTOMERNAME_COL).Value = "CUSTOMER"
        .Cells(intRow, PRICEVALUE_COL).Value = "TOTAL"
        .Cells(intRow, CURRENCYID_COL).Value = "MATA UANG"
        
        mdlExcel.FillEdge xlSheet, "A" & intRow & ":F" & intRow + 1
        .Range("A" & intRow & ":F" & intRow).Font.Bold = True
        
        intRow = intRow + 1
        
        .Cells(intRow, POID_COL).Value = strPOId
        .Cells(intRow, DATE_COL).HorizontalAlignment = xlHAlignLeft
        .Cells(intRow, DATE_COL).Value = mdlProcedures.FormatDate(mdlDatabase.GetFieldData(mdlGlobal.conInventory, "PODate", mdlTable.CreateTHSALESSUM, "POId='" & strPOId & "'"))
        .Cells(intRow, DATE_COL).NumberFormat = "dd MMMM yyyy"
        .Cells(intRow, POCUSTOMERID_COL).Value = "'" & mdlDatabase.GetFieldData(mdlGlobal.conInventory, "POCustomerId", mdlTable.CreateTHSALESSUM, "POId='" & strPOId & "'")
        .Cells(intRow, CUSTOMERNAME_COL).Value = "'" & _
            mdlDatabase.GetFieldData(mdlGlobal.conInventory, "Name", mdlTable.CreateTMCUSTOMER, "CustomerId='" & _
            mdlDatabase.GetFieldData(mdlGlobal.conInventory, "CustomerId", mdlTable.CreateTHSALESSUM, "POId='" & strPOId & "'") & "'")
        .Cells(intRow, PRICEVALUE_COL).Value = mdlDatabase.GetFieldData(mdlGlobal.conInventory, "PriceValue", mdlTable.CreateTHSALESSUM, "POId='" & strPOId & "'")
        .Cells(intRow, PRICEVALUE_COL).NumberFormat = "#,##0"
        .Cells(intRow, CURRENCYID_COL).Value = mdlDatabase.GetFieldData(mdlGlobal.conInventory, "CurrencyId", mdlTable.CreateTHSALESSUM, "POId='" & strPOId & "'")
        
        mdlExcel.FillEdge xlSheet, "A" & intRow & ":F" & intRow
    End With
    
ErrHandler:
End Sub

Private Sub Class_Terminate()
    mdlExcel.CloseExcel
End Sub
