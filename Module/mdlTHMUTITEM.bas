Attribute VB_Name = "mdlTHMUTITEM"
Option Explicit

Public Function GetQtyMUTITEM( _
    Optional ByVal strItemId As String = "", _
    Optional ByVal strWarehouseFrom As String = "", _
    Optional ByVal strWarehouseTo As String = "") As Currency
    Dim strCriteria As String
    
    strCriteria = ""
    
    If Not Trim(strWarehouseFrom) = "" Then
        strCriteria = strCriteria & "WarehouseFrom='" & strWarehouseFrom & "'"
    End If
    
    If Not Trim(strWarehouseTo) = "" Then
        If Not Trim(strCriteria) = "" Then
            strCriteria = strCriteria & " AND "
        End If
        
        strCriteria = strCriteria & "WarehouseTo='" & strWarehouseTo & "'"
    End If
    
    Dim rstTemp As ADODB.Recordset
    
    Set rstTemp = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "MutId", mdlTable.CreateTHMUTITEM, False, strCriteria)
    
    Dim curQty As Currency
    
    curQty = 0
    
    With rstTemp
        While Not .EOF
            curQty = curQty + mdlTHMUTITEM.GetTotalQtyMUTITEM(!MutId, strItemId)
            
            .MoveNext
        Wend
    End With
    
    mdlDatabase.CloseRecordset rstTemp
    
    GetQtyMUTITEM = curQty
End Function

Public Function GetTotalQtyMUTITEM(ByVal strMutId As String, Optional ByVal strItemId As String = "") As Currency
    Dim strCriteria As String
    
    strCriteria = "MutId='" & strMutId & "'"
    
    If Not Trim(strItemId) = "" Then
        If Not Trim(strCriteria) = "" Then
            strCriteria = strCriteria & " AND "
        End If
        
        strCriteria = strCriteria & "ItemId='" & strItemId & "'"
    End If
    
    Dim rstTemp As ADODB.Recordset
    
    Set rstTemp = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "Qty", mdlTable.CreateTDMUTITEM, False, strCriteria)
    
    Dim curQty As Currency
    
    curQty = 0
    
    With rstTemp
        While Not .EOF
            curQty = curQty + mdlProcedures.GetCurrency(!Qty)
            
            .MoveNext
        Wend
    End With
    
    mdlDatabase.CloseRecordset rstTemp
    
    GetTotalQtyMUTITEM = curQty
End Function

Public Sub DeleteTHMUTITEM(ByRef rstMain As ADODB.Recordset)
    Dim strRecycleId As String
    
    strRecycleId = rstMain!MutId & Space(mdlDatabase.GetColumnSize(mdlGlobal.conInventory, "ReferencesNumber", mdlTable.CreateTHRECYCLE) - Len(rstMain!MutId))
    strRecycleId = strRecycleId & mdlProcedures.FormatDate(Now, "ddMMyyyy") & mdlProcedures.FormatDate(rstMain!MutDate, "ddMMyyyy")
    
    Dim rstTemp As ADODB.Recordset
    
    Set rstTemp = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "*", mdlTable.CreateTHRECYCLE, , "RecycleId='" & strRecycleId & "'")
    
    With rstTemp
        If Not .RecordCount > 0 Then
            .AddNew
            
            !RecycleId = strRecycleId
            !ReferencesNumber = rstMain!MutId
            !RecycleDate = mdlProcedures.FormatDate(Now)
            !ReferencesDate = mdlProcedures.FormatDate(rstMain!MutDate)
            
            !CreateId = mdlGlobal.UserAuthority.UserId
            !CreateDate = mdlProcedures.FormatDate(Now)
        End If
        
        !OptInfoFirst = rstMain!WarehouseFrom
        !OptInfoSecond = rstMain!WarehouseTo
        !OptInfoThird = rstMain!Notes
        !OptInfoFourth = mdlText.strMUTITEMIDINIT
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
    
        Set rstDetail = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "*", mdlTable.CreateTDMUTITEM, False, "MutId='" & rstMain!MutId & "'")
        
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

Public Function RestoreTHMUTITEM(ByVal strRecycleId As String) As Boolean
    If Not mdlGlobal.UserAuthority.IsMenuAccess(mdiMain.mnuTHMUTITEM.Name) Then
        RestoreTHMUTITEM = False
        
        Exit Function
    End If
    
    If Not mdlDatabase.IsDataCorrect(mdlGlobal.conInventory, "OptInfoFourth", mdlTable.CreateTHRECYCLE, "RecycleId='" & strRecycleId & "'", mdlText.strMUTITEMIDINIT, True) Then
        RestoreTHMUTITEM = False
        
        Exit Function
    End If
    
    Dim blnValid As Boolean
    
    blnValid = True
    
    Dim rstTemp As ADODB.Recordset
    
    Set rstTemp = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "*", mdlTable.CreateTHRECYCLE, False, "RecycleId='" & strRecycleId & "'")
    
    If mdlDatabase.IsDataExists(mdlGlobal.conInventory, mdlTable.CreateTHMUTITEM, "MutId='" & Trim(rstTemp!ReferencesNumber) & "'") Then
        blnValid = False
    End If
    
    Dim strMutId As String
    
    If blnValid Then
        With rstTemp
            Dim rstHeader As ADODB.Recordset
            
            Set rstHeader = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "*", mdlTable.CreateTHMUTITEM, , "MutId='" & Trim(!ReferencesNumber) & "'")
            
            If rstHeader.RecordCount > 0 Then
                blnValid = False
            Else
                rstHeader.AddNew
                
                rstHeader!MutId = Trim(!ReferencesNumber)
                
                strMutId = rstHeader!MutId
                
                rstHeader!MutDate = mdlProcedures.FormatDate(!ReferencesDate)
                rstHeader!WarehouseFrom = Trim(!OptInfoFirst)
                rstHeader!WarehouseTo = Trim(!OptInfoSecond)
                rstHeader!Notes = Trim(!OptInfoThird)
                
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
            Dim strMutDtlId As String
            
            Dim rstDetail As ADODB.Recordset
            
            While Not .EOF
                strMutDtlId = strMutId & Trim(!ReferencesNumber)
                
                Set rstDetail = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "*", mdlTable.CreateTDMUTITEM, , "MutDtlId='" & strMutDtlId & "'")
                
                If Not rstDetail.RecordCount > 0 Then
                    rstDetail.AddNew
                    
                    rstDetail!MutDtlId = strMutDtlId
                    rstDetail!MutId = strMutId
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
    
    RestoreTHMUTITEM = blnValid
End Function
