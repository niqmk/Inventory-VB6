VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPRTTHSJSELL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const SJID_COL As Integer = 1
Private Const DATE_COL As Integer = 1
Private Const POID_COL As Integer = 1
Private Const POCUSTOMERID_COL As Integer = 1
Private Const CUSTOMERNAME_COL As Integer = 1
Private Const DELIVERYADDRESS_COL As Integer = 1

Private Const ITEMID_COL As Integer = 1
Private Const NAME_COL As Integer = 2
Private Const QTY_COL As Integer = 5
Private Const UNITYNAME_COL As Integer = 6

Private xlSheet As Excel.Worksheet

Private intRow As Integer

Private Sub Class_Initialize()
    mdlExcel.OpenExcel
End Sub

Public Sub ImportToExcel(ByVal strSJId As String)
    If Trim(strSJId) = "" Then Exit Sub
    
    mdlExcel.OpenWorkSheet xlSheet
    
    SetWidth
    
    FillTitle strSJId
    
    FillDetail strSJId
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

Private Sub FillTitle(ByVal strSJId As String)
    On Local Error GoTo ErrHandler
    
    With xlSheet
        .Shapes.AddPicture mdlGlobal.strLogoImageText, True, True, 0, 0, 100, 100
        
        intRow = intRow + 1
        
        .Range("A" & intRow & ":C" & intRow + 5).Font.Name = "Courier New"
        
        .Cells(intRow, SJID_COL).Value = "NO. SJ"
        .Cells(intRow, SJID_COL + 1).Value = ":"
        .Cells(intRow, SJID_COL + 2).Value = strSJId
        
        intRow = intRow + 1
        
        .Cells(intRow, DATE_COL).Value = "TGL."
        .Cells(intRow, DATE_COL + 1).Value = ":"
        .Cells(intRow, DATE_COL + 2).HorizontalAlignment = xlHAlignLeft
        .Cells(intRow, DATE_COL + 2).Value = mdlProcedures.FormatDate(mdlDatabase.GetFieldData(mdlGlobal.conInventory, "SJDate", mdlTable.CreateTHSJSELL, "SJId='" & strSJId & "'"))
        .Cells(intRow, DATE_COL + 2).NumberFormat = "dd MMMM yyyy"
        
        intRow = intRow + 1
        
        Dim strPOId As String
        
        strPOId = mdlDatabase.GetFieldData(mdlGlobal.conInventory, "POId", mdlTable.CreateTHSOSELL, "SOId='" & _
            mdlDatabase.GetFieldData(mdlGlobal.conInventory, "SOId", mdlTable.CreateTHSJSELL, "SJId='" & strSJId & "'") & "'")
        
        .Cells(intRow, POID_COL).Value = "NO. PO"
        .Cells(intRow, POID_COL + 1).Value = ":"
        .Cells(intRow, POID_COL + 2).Value = strPOId
        
        intRow = intRow + 1
        
        .Cells(intRow, POCUSTOMERID_COL).Value = "CUSTOMER PO"
        .Cells(intRow, POCUSTOMERID_COL + 1).Value = ":"
        .Cells(intRow, POCUSTOMERID_COL + 2).Value = "'" & _
            mdlDatabase.GetFieldData(mdlGlobal.conInventory, "POCustomerId", mdlTable.CreateTHPOSELL, "POId='" & strPOId & "'")
        
        Dim strCustomerId As String
        
        strCustomerId = mdlDatabase.GetFieldData(mdlGlobal.conInventory, "CustomerId", mdlTable.CreateTHPOSELL, "POId='" & strPOId & "'")
        
        .Cells(intRow, CUSTOMERNAME_COL).Value = "CUSTOMER"
        .Cells(intRow, CUSTOMERNAME_COL + 1).Value = ":"
        .Cells(intRow, CUSTOMERNAME_COL + 2).Value = "'" & _
            mdlDatabase.GetFieldData(mdlGlobal.conInventory, "Name", mdlTable.CreateTMCUSTOMER, "CustomerId='" & strCustomerId & "'")
        
        intRow = intRow + 1
        
        Dim strDeliveryId As String
        
        strDeliveryId = mdlDatabase.GetFieldData(mdlGlobal.conInventory, "DeliveryId", mdlTable.CreateTHSJSELL, "SJId='" & strSJId & "'")
        
        xlSheet.Range("A" & intRow & ":C" & intRow).VerticalAlignment = xlVAlignTop
        
        .Cells(intRow, DELIVERYADDRESS_COL).Value = "KIRIM KE"
        .Cells(intRow, DELIVERYADDRESS_COL + 1).Value = ":"
        
        If mdlDatabase.IsDataExists(mdlGlobal.conInventory, mdlTable.CreateTMDELIVERYCUSTOMER, "DeliveryId='" & strDeliveryId & "'") Then
            .Cells(intRow, DELIVERYADDRESS_COL + 2).Value = "'" & _
                mdlDatabase.GetFieldData(mdlGlobal.conInventory, "Address", mdlTable.CreateTMDELIVERYCUSTOMER, "DeliveryId='" & strDeliveryId & "'")
        Else
            .Cells(intRow, DELIVERYADDRESS_COL + 2).Value = "'" & _
                mdlDatabase.GetFieldData(mdlGlobal.conInventory, "Address", mdlTable.CreateTMCUSTOMER, "CustomerId='" & strDeliveryId & "'")
        End If
        
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

Private Sub FillDetail(Optional ByVal strSJId As String = "")
    On Local Error GoTo ErrHandler
    
    Dim rstDetail As ADODB.Recordset
    
    Set rstDetail = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "*", mdlTable.CreateTDSJSELL, False, "SJId='" & strSJId & "'", "ItemId ASC")
    
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
