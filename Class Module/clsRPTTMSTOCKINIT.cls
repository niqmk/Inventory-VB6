VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRPTTMSTOCKINIT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const WAREHOUSENAME_COL As Integer = 1

Private Const ITEMID_COL As Integer = 1
Private Const PARTNUMBER_COL As Integer = 2
Private Const NAME_COL As Integer = 3
Private Const QTY_COL As Integer = 4
Private Const UNITYNAME_COL As Integer = 5

Private xlSheet As Excel.Worksheet

Private intRow As Integer

Private Sub Class_Initialize()
    mdlExcel.OpenExcel
End Sub

Public Sub ImportToExcel( _
    Optional ByVal strItemId As String = "", _
    Optional ByVal strPartNumber As String = "", _
    Optional ByVal strName As String = "", _
    Optional ByVal strWarehouseId As String = "", _
    Optional ByVal strVendor As String = "", _
    Optional ByVal strGroup As String = "", _
    Optional ByVal strCategory As String = "", _
    Optional ByVal strBrand As String = "", _
    Optional ByVal strUnity As String = "")
    mdlExcel.OpenWorkSheet xlSheet
    
    SetWidth
    
    FillTitle strItemId, strPartNumber, strName, strWarehouseId, strGroup, strCategory, strBrand, strUnity
End Sub

Private Sub SetWidth()
    intRow = 1
    
    With xlSheet
        .Columns("A").ColumnWidth = 8
        .Columns("B").ColumnWidth = 20
        .Columns("C").ColumnWidth = 20
        .Columns("D").ColumnWidth = 10
        .Columns("E").ColumnWidth = 25
        
        .Rows(intRow).RowHeight = 125
    End With
End Sub

Private Sub FillTitle( _
    Optional ByVal strItemId As String = "", _
    Optional ByVal strPartNumber As String = "", _
    Optional ByVal strName As String = "", _
    Optional ByVal strWarehouseId As String = "", _
    Optional ByVal strVendor As String = "", _
    Optional ByVal strGroup As String = "", _
    Optional ByVal strCategory As String = "", _
    Optional ByVal strBrand As String = "", _
    Optional ByVal strUnity As String = "")
    On Local Error GoTo ErrHandler
    
    With xlSheet
        .Shapes.AddPicture mdlGlobal.strLogoImageText, True, True, 0, 0, 100, 100
        
        Dim strCriteria As String
        
        strCriteria = ""
        
        If Not Trim(strWarehouseId) = "" Then
            strCriteria = "WarehouseId='" & strWarehouseId & "'"
        End If
        
        Dim rstTemp As ADODB.Recordset
        
        Set rstTemp = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "WarehouseId, Name", mdlTable.CreateTMWAREHOUSE, False, strCriteria)
        
        While Not rstTemp.EOF
            intRow = intRow + 1
            
            .Range("A" & intRow & ":E" & intRow).Font.Name = "Courier New"
            
            .Cells(intRow, WAREHOUSENAME_COL).Value = "Gudang : " & rstTemp!WarehouseId & " | " & rstTemp!Name
            
            intRow = intRow + 2
            
            mdlExcel.FillEdge xlSheet, "A" & intRow & ":E" & intRow + 1
            
            .Range("A" & intRow & ":E" & intRow).Font.Bold = True
            
            .Cells(intRow, ITEMID_COL).Value = "Kode"
            .Cells(intRow, PARTNUMBER_COL).Value = "Nomor Part"
            .Cells(intRow, NAME_COL).Value = "Nama"
            .Cells(intRow, QTY_COL).Value = "Qty"
            .Cells(intRow, UNITYNAME_COL).Value = "Satuan"
            
            FillText strItemId, strPartNumber, strName, rstTemp!WarehouseId, strVendor, strGroup, strCategory, strBrand, strUnity
            
            intRow = intRow + 1
            
            rstTemp.MoveNext
        Wend
        
        mdlDatabase.CloseRecordset rstTemp
    End With
    
ErrHandler:
End Sub

Private Sub FillText( _
    Optional ByVal strItemId As String = "", _
    Optional ByVal strPartNumber As String = "", _
    Optional ByVal strName As String = "", _
    Optional ByVal strWarehouseId As String = "", _
    Optional ByVal strVendor As String = "", _
    Optional ByVal strGroup As String = "", _
    Optional ByVal strCategory As String = "", _
    Optional ByVal strBrand As String = "", _
    Optional ByVal strUnity As String = "")
    On Local Error GoTo ErrHandler
    
    Dim strTableFirst As String
    Dim strTableSecond As String
    Dim strTableThird As String
    Dim strTableFourth As String
    Dim strTableFifth As String
    Dim strTableSixth As String
    Dim strTableSeventh As String
    Dim strTable As String
    
    strTableFirst = mdlTable.CreateTMSTOCKINIT
    strTableSecond = mdlTable.CreateTMITEM
    strTableThird = mdlTable.CreateTMVENDOR
    strTableFourth = mdlTable.CreateTMGROUP
    strTableFifth = mdlTable.CreateTMCATEGORY
    strTableSixth = mdlTable.CreateTMBRAND
    strTableSeventh = mdlTable.CreateTMUNITY
    
    strTable = "((((((" & strTableFirst & " LEFT JOIN " & strTableSecond & _
        " ON " & strTableFirst & ".ItemId=" & strTableSecond & ".ItemId) LEFT JOIN " & strTableThird & _
        " ON " & strTableSecond & ".VendorId=" & strTableThird & ".VendorId) LEFT JOIN " & strTableFourth & _
        " ON " & strTableSecond & ".GroupId=" & strTableFourth & ".GroupId) LEFT JOIN " & strTableFifth & _
        " ON " & strTableSecond & ".CategoryId=" & strTableFifth & ".CategoryId) LEFT JOIN " & strTableSixth & _
        " ON " & strTableSecond & ".BrandId=" & strTableSixth & ".BrandId) LEFT JOIN " & strTableSeventh & _
        " ON " & strTableSecond & ".UnityId=" & strTableSeventh & ".UnityId)"
        
    Dim strCriteria As String
    
    strCriteria = ""
    
    If Not Trim(strItemId) = "" Then
        strCriteria = mdlProcedures.QueryLikeCriteria(strTableFirst & ".ItemId", mdlProcedures.RepDupText(strItemId))
    End If
    
    If Not Trim(strPartNumber) = "" Then
        If Not Trim(strCriteria) = "" Then
            strCriteria = strCriteria & " AND "
        End If
        
        strCriteria = strCriteria & mdlProcedures.QueryLikeCriteria("PartNumber", mdlProcedures.RepDupText(strPartNumber))
    End If
    
    If Not Trim(strName) = "" Then
        If Not Trim(strCriteria) = "" Then
            strCriteria = strCriteria & " AND "
        End If
        
        strCriteria = strCriteria & mdlProcedures.QueryLikeCriteria(strTableSecond & ".Name", mdlProcedures.RepDupText(strName))
    End If
    
    Dim strOptional(1) As String
    
    If Not Trim(strVendor) = "" Then
        strOptional(0) = mdlProcedures.SplitData(strVendor)
        strOptional(1) = mdlProcedures.SplitData(strVendor, 1)
        
        If Not Trim(strOptional(0)) = "" Then
            If Not Trim(strCriteria) = "" Then
                strCriteria = strCriteria & " AND "
            End If
            
            strCriteria = strCriteria & mdlProcedures.QueryLikeCriteria(strTableSecond & ".VendorId", mdlProcedures.RepDupText(strOptional(0)))
        End If
        
        If Not Trim(strOptional(1)) = "" Then
            If Not Trim(strCriteria) = "" Then
                strCriteria = strCriteria & " AND "
            End If
            
            strCriteria = strCriteria & mdlProcedures.QueryLikeCriteria(strTableThird & ".Name", mdlProcedures.RepDupText(strOptional(1)))
        End If
    End If
    
    If Not Trim(strGroup) = "" Then
        strOptional(0) = mdlProcedures.SplitData(strGroup)
        strOptional(1) = mdlProcedures.SplitData(strGroup, 1)
        
        If Not Trim(strOptional(0)) = "" Then
            If Not Trim(strCriteria) = "" Then
                strCriteria = strCriteria & " AND "
            End If
            
            strCriteria = strCriteria & mdlProcedures.QueryLikeCriteria(strTableSecond & ".GroupId", mdlProcedures.RepDupText(strOptional(0)))
        End If
        
        If Not Trim(strOptional(1)) = "" Then
            If Not Trim(strCriteria) = "" Then
                strCriteria = strCriteria & " AND "
            End If
            
            strCriteria = strCriteria & mdlProcedures.QueryLikeCriteria(strTableFourth & ".Name", mdlProcedures.RepDupText(strOptional(1)))
        End If
    End If
    
    If Not Trim(strCategory) = "" Then
        strOptional(0) = mdlProcedures.SplitData(strCategory)
        strOptional(1) = mdlProcedures.SplitData(strCategory, 1)
        
        If Not Trim(strOptional(0)) = "" Then
            If Not Trim(strCriteria) = "" Then
                strCriteria = strCriteria & " AND "
            End If
            
            strCriteria = strCriteria & mdlProcedures.QueryLikeCriteria(strTableSecond & ".CategoryId", mdlProcedures.RepDupText(strOptional(0)))
        End If
        
        If Not Trim(strOptional(1)) = "" Then
            If Not Trim(strCriteria) = "" Then
                strCriteria = strCriteria & " AND "
            End If
            
            strCriteria = strCriteria & mdlProcedures.QueryLikeCriteria(strTableFifth & ".Name", mdlProcedures.RepDupText(strOptional(1)))
        End If
    End If
    
    If Not Trim(strBrand) = "" Then
        strOptional(0) = mdlProcedures.SplitData(strBrand)
        strOptional(1) = mdlProcedures.SplitData(strBrand, 1)
        
        If Not Trim(strOptional(0)) = "" Then
            If Not Trim(strCriteria) = "" Then
                strCriteria = strCriteria & " AND "
            End If
            
            strCriteria = strCriteria & mdlProcedures.QueryLikeCriteria(strTableSecond & ".BrandId", mdlProcedures.RepDupText(strOptional(0)))
        End If
        
        If Not Trim(strOptional(1)) = "" Then
            If Not Trim(strCriteria) = "" Then
                strCriteria = strCriteria & " AND "
            End If
            
            strCriteria = strCriteria & mdlProcedures.QueryLikeCriteria(strTableSixth & ".Name", mdlProcedures.RepDupText(strOptional(1)))
        End If
    End If
    
    If Not Trim(strUnity) = "" Then
        strOptional(0) = mdlProcedures.SplitData(strUnity)
        strOptional(1) = mdlProcedures.SplitData(strUnity, 1)
        
        If Not Trim(strOptional(0)) = "" Then
            If Not Trim(strCriteria) = "" Then
                strCriteria = strCriteria & " AND "
            End If
            
            strCriteria = strCriteria & mdlProcedures.QueryLikeCriteria(strTableSecond & ".UnityId", mdlProcedures.RepDupText(strOptional(0)))
        End If
        
        If Not Trim(strOptional(1)) = "" Then
            If Not Trim(strCriteria) = "" Then
                strCriteria = strCriteria & " AND "
            End If
            
            strCriteria = strCriteria & mdlProcedures.QueryLikeCriteria(strTableSeventh & ".Name", mdlProcedures.RepDupText(strOptional(1)))
        End If
    End If
    
    If Not Trim(strCriteria) = "" Then
        strCriteria = strCriteria & " AND "
    End If
    
    strCriteria = strCriteria & "WarehouseId='" & strWarehouseId & "'"
    
    Dim rstTemp As ADODB.Recordset
    
    Set rstTemp = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, strTableFirst & ".ItemId, PartNumber, " & strTableSecond & ".Name, Qty, " & strTableThird & ".Name", strTable, False, strCriteria, strTableFirst & ".ItemId ASC")
    
    With rstTemp
        Dim intTemp As Integer
        
        Dim mItemId As String

        intTemp = intRow + 1
    
        While Not .EOF
            intRow = intRow + 1
            
            xlSheet.Range("A" & intRow & ":E" & intRow).VerticalAlignment = xlTop
            xlSheet.Range("A" & intRow & ":C" & intRow).HorizontalAlignment = xlLeft
            xlSheet.Range("D" & intRow & ":D" & intRow).HorizontalAlignment = xlRight
            xlSheet.Range("E" & intRow & ":E" & intRow).HorizontalAlignment = xlLeft
            xlSheet.Range("A" & intRow & ":E" & intRow).WrapText = True
            
            mItemId = Trim(.Fields("ItemId").Value)
            
            xlSheet.Cells(intRow, ITEMID_COL).Value = "'" & mItemId
            xlSheet.Cells(intRow, PARTNUMBER_COL).Value = "'" & Trim(mdlDatabase.GetFieldData(mdlGlobal.conInventory, "PartNumber", mdlTable.CreateTMITEM, "ItemId='" & mItemId & "'"))
            xlSheet.Cells(intRow, NAME_COL).Value = "'" & Trim(mdlDatabase.GetFieldData(mdlGlobal.conInventory, "Name", mdlTable.CreateTMITEM, "ItemId='" & mItemId & "'"))
            xlSheet.Cells(intRow, QTY_COL).Value = CStr(.Fields("Qty").Value)
            xlSheet.Cells(intRow, QTY_COL).NumberFormat = "#,##0"
            xlSheet.Cells(intRow, UNITYNAME_COL).Value = "'" & Trim(mdlDatabase.GetFieldData( _
                mdlGlobal.conInventory, _
                "Name", _
                mdlTable.CreateTMUNITY, _
                "UnityId='" & _
                    mdlDatabase.GetFieldData( _
                        mdlGlobal.conInventory, _
                        "UnityId", _
                        mdlTable.CreateTMITEM, _
                        "ItemId='" & mItemId & "'") & "'"))
            
            xlSheet.Range("A" & intRow & ":E" & intRow).Select
            
            .MoveNext
        Wend
        
        If .RecordCount > 0 Then
            mdlExcel.FillEdge xlSheet, "A" & intTemp & ":E" & intRow, False
            
            xlSheet.Range("A" & intTemp & ":E" & intTemp).Select
        End If
    End With
    
ErrHandler:
End Sub

Private Sub Class_Terminate()
    mdlExcel.CloseExcel
End Sub
