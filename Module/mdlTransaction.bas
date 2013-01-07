Attribute VB_Name = "mdlTransaction"
Option Explicit

Public Function CheckStock(ByVal strItemId As String, Optional ByVal strWarehouseId As String = "") As Currency
    Dim strCriteria As String
    
    strCriteria = "ItemId='" & strItemId & "'"
    
    If Trim(strWarehouseId) = "" Then
        strWarehouseId = mdlTransaction.GetWarehouseIdSet
    End If
    
    If Not Trim(strWarehouseId) = "" Then
        strCriteria = strCriteria & " AND WarehouseId='" & strWarehouseId & "'"
    End If
    
    Dim rstTemp As ADODB.Recordset
    
    Set rstTemp = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "QtyIn, QtyOut", mdlTable.CreateTHSTOCK, False, strCriteria)
    
    Dim curQty As Currency
    
    curQty = mdlProcedures.GetCurrency(mdlDatabase.GetFieldData(mdlGlobal.conInventory, "Qty", mdlTable.CreateTMSTOCKINIT, strCriteria))
    
    With rstTemp
        While Not .EOF
            curQty = curQty + (mdlProcedures.GetCurrency(!QtyIn) - mdlProcedures.GetCurrency(!QtyOut))
            
            .MoveNext
        Wend
    End With
    
    mdlDatabase.CloseRecordset rstTemp
    
    curQty = curQty - mdlTHMUTITEM.GetQtyMUTITEM(strItemId, strWarehouseId)
    curQty = curQty + mdlTHMUTITEM.GetQtyMUTITEM(strItemId, , strWarehouseId)
    
    CheckStock = curQty
End Function

Public Sub UpdateStock( _
    ByVal strItemId As String, _
    ByVal strWarehouseId As String, _
    ByVal strReferencesNumber As String, _
    ByVal dteStock As Date, _
    Optional ByVal curQtyIn As Currency = 0, _
    Optional ByVal curQtyOut As Currency = 0)
    Dim strStockId As String
    
    strStockId = strItemId & Space(mdlDatabase.GetColumnSize(mdlGlobal.conInventory, "ItemId", mdlTable.CreateTHSTOCK) - Len(strItemId))
    strStockId = strStockId & strWarehouseId & Space(mdlDatabase.GetColumnSize(mdlGlobal.conInventory, "WarehouseId", mdlTable.CreateTHSTOCK) - Len(strWarehouseId))
    strStockId = strStockId & strReferencesNumber & Space(mdlDatabase.GetColumnSize(mdlGlobal.conInventory, "ReferencesNumber", mdlTable.CreateTHSTOCK) - Len(strReferencesNumber))
    strStockId = strStockId & mdlProcedures.FormatDate(dteStock, "ddMMyyyy")
    
    Dim strCriteria As String
    
    strCriteria = "StockId='" & strStockId & "'"
    
    If curQtyIn = 0 And curQtyOut = 0 Then
        mdlDatabase.DeleteRecordQuery mdlGlobal.conInventory, mdlTable.CreateTHSTOCK, strCriteria
        
        Exit Sub
    End If
    
    Dim rstTemp As ADODB.Recordset
    
    Set rstTemp = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "*", mdlTable.CreateTHSTOCK, , strCriteria)
    
    With rstTemp
        If Not .RecordCount > 0 Then
            .AddNew
            
            !StockId = strStockId
            !ItemId = strItemId
            !WarehouseId = strWarehouseId
            !ReferencesNumber = strReferencesNumber
            !StockDate = mdlProcedures.FormatDate(dteStock)
            
            !CreateId = Trim(mdlGlobal.UserAuthority.UserId)
            !CreateDate = mdlProcedures.FormatDate(Now)
        End If
        
        If IsNull(!QtyIn) Then !QtyIn = mdlProcedures.GetCurrency("0")
        !QtyIn = mdlProcedures.GetCurrency(!QtyIn) + (curQtyIn)
        
        If IsNull(!QtyOut) Then !QtyOut = mdlProcedures.GetCurrency("0")
        !QtyOut = mdlProcedures.GetCurrency(!QtyOut) + (curQtyOut)
        
        !UpdateId = Trim(mdlGlobal.UserAuthority.UserId)
        !UpdateDate = mdlProcedures.FormatDate(Now)
            
        .Update
    End With
    
    mdlDatabase.CloseRecordset rstTemp
End Sub

Public Function GetWarehouseIdSet() As String
    Dim rstTemp As ADODB.Recordset
    
    Set rstTemp = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "WarehouseId", mdlTable.CreateTMWAREHOUSE, False, "WarehouseSet='" & mdlGlobal.strYes & "'")
    
    If rstTemp.RecordCount > 0 Then
        GetWarehouseIdSet = rstTemp!WarehouseId
    Else
        GetWarehouseIdSet = ""
    End If
    
    mdlDatabase.CloseRecordset rstTemp
End Function

Public Function ConvertCurrency(ByVal strCurrencyFromId As String, ByVal strCurrencyToId As String, Optional ByVal curValue As Currency) As Currency
    If Trim(strCurrencyFromId) = "" Then
        ConvertCurrency = 0
        
        Exit Function
    End If
    
    If Trim(strCurrencyToId) = "" Then
        ConvertCurrency = curValue
        
        Exit Function
    End If
    
    If Trim(strCurrencyFromId) = Trim(strCurrencyToId) Then
        ConvertCurrency = curValue
        
        Exit Function
    End If
    
    Dim strCriteria As String
    
    strCriteria = "CurrencyFromId='" & strCurrencyFromId & "'"
    strCriteria = strCriteria & " AND CurrencyToId='" & strCurrencyToId & "'"
    
    Dim rstTemp As ADODB.Recordset
    
    Set rstTemp = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "ConvertValue", mdlTable.CreateTMCONVERTCURRENCY, False, strCriteria, "ConvertDate DESC")
    
    If rstTemp.RecordCount > 0 Then
        ConvertCurrency = mdlProcedures.GetCurrency(rstTemp!ConvertValue) * curValue
    Else
        ConvertCurrency = 0
    End If
    
    mdlDatabase.CloseRecordset rstTemp
End Function

Public Function IsRecycleExists() As Boolean
    IsRecycleExists = mdlDatabase.IsDataExists(mdlGlobal.conInventory, mdlTable.CreateTHRECYCLE)
End Function

Public Function ExtractSequential( _
    ByVal strReferencesNumber As String, _
    Optional ByVal intIndex As Integer = 0, _
    Optional ByVal blnIsUpperIndex As Boolean = True) As Integer
    Dim strExtract() As String
    
    strExtract = Split(strReferencesNumber, "/")
    
    If intIndex > UBound(strExtract) Then
        ExtractSequential = 0
        
        Exit Function
    End If
    
    If blnIsUpperIndex Then
        If Not intIndex = UBound(strExtract) Then
            ExtractSequential = 0
            
            Exit Function
        End If
    End If
    
    If Not IsNumeric(strExtract(intIndex)) Then
        ExtractSequential = 0
        
        Exit Function
    End If
    
    ExtractSequential = CInt(strExtract(intIndex))
End Function
