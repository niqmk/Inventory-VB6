Attribute VB_Name = "mdlTHDOBUY"
Option Explicit

Public Function GetQtyPOFromDOBUY(ByVal strPOId As String, Optional ByVal strItemId As String = "") As Currency
    Dim strCriteria As String
    
    strCriteria = "POId='" & strPOId & "'"
    
    Dim rstTemp As ADODB.Recordset
    
    Set rstTemp = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "DOId", mdlTable.CreateTHDOBUY, False, strCriteria)
    
    Dim curQty As Currency
    
    curQty = 0
    
    With rstTemp
        While Not .EOF
            curQty = curQty + mdlTHDOBUY.GetTotalQtyDOBUY(!DOId, strItemId)
            
            .MoveNext
        Wend
    End With
    
    mdlDatabase.CloseRecordset rstTemp
    
    GetQtyPOFromDOBUY = curQty
End Function

Public Function GetTotalQtyDOBUY(ByVal strDOId As String, Optional ByVal strItemId As String = "") As Currency
    Dim strCriteria As String
    
    strCriteria = "DOId='" & strDOId & "'"
    
    If Not Trim(strItemId) = "" Then
        strCriteria = strCriteria & " AND ItemId='" & strItemId & "'"
    End If
    
    Dim rstTemp As ADODB.Recordset
    
    Set rstTemp = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "Qty", mdlTable.CreateTDDOBUY, False, strCriteria)
    
    Dim curQty As Currency
    
    curQty = 0
    
    With rstTemp
        While Not .EOF
            curQty = curQty + mdlProcedures.GetCurrency(!Qty)
            
            .MoveNext
        Wend
    End With
    
    mdlDatabase.CloseRecordset rstTemp
    
    GetTotalQtyDOBUY = curQty
End Function

Public Sub DeleteTHDOBUY(ByRef rstMain As ADODB.Recordset)
    Dim strRecycleId As String
    
    strRecycleId = rstMain!DOId & Space(mdlDatabase.GetColumnSize(mdlGlobal.conInventory, "ReferencesNumber", mdlTable.CreateTHRECYCLE) - Len(rstMain!DOId))
    strRecycleId = strRecycleId & mdlProcedures.FormatDate(Now, "ddMMyyyy") & mdlProcedures.FormatDate(rstMain!DODate, "ddMMyyyy")
    
    Dim rstTemp As ADODB.Recordset
    
    Set rstTemp = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "*", mdlTable.CreateTHRECYCLE, , "RecycleId='" & strRecycleId & "'")
    
    With rstTemp
        If Not .RecordCount > 0 Then
            .AddNew
            
            !RecycleId = strRecycleId
            !ReferencesNumber = rstMain!DOId
            !RecycleDate = mdlProcedures.FormatDate(Now)
            !ReferencesDate = mdlProcedures.FormatDate(rstMain!DODate)
            
            !CreateId = mdlGlobal.UserAuthority.UserId
            !CreateDate = mdlProcedures.FormatDate(Now)
        End If
        
        !OptInfoFirst = rstMain!POId
        !OptInfoSecond = rstMain!WarehouseId
        !OptInfoThird = rstMain!TransportNumber
        !OptInfoFourth = rstMain!Notes
        !OptInfoFifth = ""
        !OptInfoSixth = ""
        !OptInfoSeventh = ""
        !OptInfoEight = ""
        !OptInfoNineth = ""
        !OptInfoTenth = ""
        
        !UpdateId = mdlGlobal.UserAuthority.UserId
        !UpdateDate = mdlProcedures.FormatDate(Now)
        
        .Update
    End With
    
    Set rstTemp = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "*", mdlTable.CreateTDRECYCLE, , "RecycleId='" & strRecycleId & "'")
    
    With rstTemp
        Dim rstDetail As ADODB.Recordset
    
        Set rstDetail = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "*", mdlTable.CreateTDDOBUY, False, "DOId='" & rstMain!DOId & "'")
        
        While Not rstDetail.EOF
            If Not .RecordCount > 0 Then
                .AddNew
                
                !RecycleDtlId = strRecycleId & rstDetail!ItemId
                !RecycleId = strRecycleId
                !ReferencesNumber = rstDetail!ItemId
                
                !CreateId = mdlGlobal.UserAuthority.UserId
                !CreateDate = mdlProcedures.FormatDate(Now)
            End If
            
            !OptInfoFirst = rstDetail!Qty
            !OptInfoSecond = ""
            !OptInfoThird = ""
            !OptInfoFourth = ""
            !OptInfoFifth = ""
            
            !UpdateId = mdlGlobal.UserAuthority.UserId
            !UpdateDate = mdlProcedures.FormatDate(Now)
            
            .Update
        
            rstDetail.MoveNext
        Wend
    End With
    
    mdlDatabase.CloseRecordset rstTemp
End Sub

Public Function RestoreTHDOBUY(ByVal strRecycleId As String) As Boolean
    If Not mdlGlobal.UserAuthority.IsMenuAccess(mdiMain.mnuTHDOBUY.Name) Then
        RestoreTHDOBUY = False
        
        Exit Function
    End If
    
    Dim blnValid As Boolean
    
    blnValid = True
    
    Dim rstTemp As ADODB.Recordset
    
    Set rstTemp = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "*", mdlTable.CreateTHRECYCLE, False, "RecycleId='" & strRecycleId & "'")
    
    If Not mdlDatabase.IsDataExists(mdlGlobal.conInventory, mdlTable.CreateTHPOBUY, "POId='" & Trim(rstTemp!OptInfoFirst) & "'") Then
        blnValid = False
    End If
    
    Dim strDOId As String
    
    If blnValid Then
        With rstTemp
            Dim rstHeader As ADODB.Recordset
            
            Set rstHeader = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "*", mdlTable.CreateTHDOBUY, , "DOId='" & Trim(!ReferencesNumber) & "'")
            
            If rstHeader.RecordCount > 0 Then
                blnValid = False
            Else
                rstHeader.AddNew
                
                rstHeader!DOId = Trim(!ReferencesNumber)
                
                strDOId = rstHeader!DOId
                
                rstHeader!DODate = mdlProcedures.FormatDate(!ReferencesDate)
                rstHeader!POId = Trim(!OptInfoFirst)
                rstHeader!Notes = Trim(!OptInfoSecond)
                
                rstHeader!CreateId = mdlGlobal.UserAuthority.UserId
                rstHeader!CreateDate = mdlProcedures.FormatDate(Now)
                rstHeader!UpdateId = mdlGlobal.UserAuthority.UserId
                rstHeader!UpdateDate = mdlProcedures.FormatDate(Now)
                
                rstHeader.Update
            End If
            
            mdlDatabase.CloseRecordset rstHeader
        End With
    End If
    
    If blnValid Then
        Set rstTemp = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "*", mdlTable.CreateTDRECYCLE, False, "RecycleId='" & strRecycleId & "'")
        
        With rstTemp
            Dim strDODtlId As String
            
            Dim rstDetail As ADODB.Recordset
            
            While Not .EOF
                strDODtlId = strDOId & Trim(!ReferencesNumber)
                
                Set rstDetail = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "*", mdlTable.CreateTDDOBUY, , "DODtlId='" & strDODtlId & "'")
                
                If Not rstDetail.RecordCount > 0 Then
                    rstDetail.AddNew
                    
                    rstDetail!DODtlId = strDODtlId
                    rstDetail!DOId = strDOId
                    rstDetail!ItemId = Trim(!ReferencesNumber)
                    
                    rstDetail!CreateId = mdlGlobal.UserAuthority.UserId
                    rstDetail!CreateDate = mdlProcedures.FormatDate(Now)
                End If
                
                rstDetail!Qty = mdlProcedures.GetCurrency(Trim(!OptInfoFirst))
                rstDetail!UpdateId = mdlGlobal.UserAuthority.UserId
                rstDetail!UpdateDate = mdlProcedures.FormatDate(Now)
                
                rstDetail.Update
                
                .MoveNext
            Wend
            
            mdlDatabase.CloseRecordset rstDetail
        End With
    End If
    
    mdlDatabase.CloseRecordset rstTemp
    
    RestoreTHDOBUY = blnValid
End Function
