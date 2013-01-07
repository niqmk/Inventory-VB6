Attribute VB_Name = "mdlTHRTRBUY"
Option Explicit

Public Function GetQtySJFromRTRBUY(ByVal strSJId As String, Optional ByVal strItemId As String = "") As Currency
    Dim strCriteria As String
    
    strCriteria = "SJId='" & strSJId & "'"
    
    Dim rstTemp As ADODB.Recordset
    
    Set rstTemp = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "RtrId", mdlTable.CreateTHRTRBUY, False, strCriteria)
    
    Dim curQty As Currency
    
    curQty = 0
    
    With rstTemp
        While Not .EOF
            curQty = curQty + mdlTHRTRBUY.GetTotalQtyRTRBUY(!RtrId, strItemId)
            
            .MoveNext
        Wend
    End With
    
    mdlDatabase.CloseRecordset rstTemp
    
    GetQtySJFromRTRBUY = curQty
End Function

Public Function GetTotalQtyRTRBUY(ByVal strRtrId As String, Optional ByVal strItemId As String = "") As Currency
    Dim strCriteria As String
    
    strCriteria = "RtrId='" & strRtrId & "'"
    
    If Not Trim(strItemId) = "" Then
        strCriteria = strCriteria & " AND ItemId='" & strItemId & "'"
    End If
    
    Dim rstTemp As ADODB.Recordset
    
    Set rstTemp = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "Qty", mdlTable.CreateTDRTRBUY, False, strCriteria)
    
    Dim curQty As Currency
    
    curQty = 0
    
    With rstTemp
        While Not .EOF
            curQty = curQty + mdlProcedures.GetCurrency(!Qty)
            
            .MoveNext
        Wend
    End With
    
    mdlDatabase.CloseRecordset rstTemp
    
    GetTotalQtyRTRBUY = curQty
End Function

Public Sub DeleteTHRTRBUY(ByRef rstMain As ADODB.Recordset)
    Dim strRecycleId As String
    
    strRecycleId = rstMain!RtrId & Space(mdlDatabase.GetColumnSize(mdlGlobal.conInventory, "ReferencesNumber", mdlTable.CreateTHRECYCLE) - Len(rstMain!RtrId))
    strRecycleId = strRecycleId & mdlProcedures.FormatDate(Now, "ddMMyyyy") & mdlProcedures.FormatDate(rstMain!RtrDate, "ddMMyyyy")
    
    Dim rstTemp As ADODB.Recordset
    
    Set rstTemp = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "*", mdlTable.CreateTHRECYCLE, , "RecycleId='" & strRecycleId & "'")
    
    With rstTemp
        If Not .RecordCount > 0 Then
            .AddNew
            
            !RecycleId = strRecycleId
            !ReferencesNumber = rstMain!RtrId
            !RecycleDate = mdlProcedures.FormatDate(Now)
            !ReferencesDate = mdlProcedures.FormatDate(rstMain!RtrDate)
            
            !CreateId = mdlGlobal.UserAuthority.UserId
            !CreateDate = mdlProcedures.FormatDate(Now)
        End If
        
        !OptInfoFirst = rstMain!SJId
        !OptInfoSecond = rstMain!Notes
        !OptInfoThird = ""
        !OptInfoFourth = ""
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
    
        Set rstDetail = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "*", mdlTable.CreateTDRTRBUY, False, "RtrId='" & rstMain!RtrId & "'")
        
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

Public Function RestoreTHRTRBUY(ByVal strRecycleId As String) As Boolean
    If Not mdlGlobal.UserAuthority.IsMenuAccess(mdiMain.mnuTHRTRBUY.Name) Then
        RestoreTHRTRBUY = False
        
        Exit Function
    End If
    
    Dim blnValid As Boolean
    
    blnValid = True
    
    Dim rstTemp As ADODB.Recordset
    
    Set rstTemp = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "*", mdlTable.CreateTHRECYCLE, False, "RecycleId='" & strRecycleId & "'")
    
    If Not mdlDatabase.IsDataExists(mdlGlobal.conInventory, mdlTable.CreateTHSJBUY, "SJId='" & Trim(rstTemp!OptInfoFirst) & "'") Then
        blnValid = False
    End If
    
    Dim strRtrId As String
    Dim strSJId As String
    Dim dteRtrDate As Date
    
    If blnValid Then
        With rstTemp
            Dim rstHeader As ADODB.Recordset
            
            Set rstHeader = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "*", mdlTable.CreateTHRTRBUY, , "RtrId='" & Trim(!ReferencesNumber) & "'")
            
            If rstHeader.RecordCount > 0 Then
                blnValid = False
            Else
                rstHeader.AddNew
                
                rstHeader!RtrId = Trim(!ReferencesNumber)
                
                strRtrId = rstHeader!RtrId
                strSJId = Trim(!OptInfoFirst)
                dteRtrDate = mdlProcedures.FormatDate(!ReferencesDate)
                
                rstHeader!RtrDate = mdlProcedures.FormatDate(!ReferencesDate)
                rstHeader!SJId = Trim(!OptInfoFirst)
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
            Dim strRtrDtlId As String
            
            Dim rstDetail As ADODB.Recordset
            
            While Not .EOF
                strRtrDtlId = strRtrId & Trim(!ReferencesNumber)
                
                Set rstDetail = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "*", mdlTable.CreateTDRTRBUY, , "RtrDtlId='" & strRtrDtlId & "'")
                
                If Not rstDetail.RecordCount > 0 Then
                    rstDetail.AddNew
                    
                    rstDetail!RtrDtlId = strRtrDtlId
                    rstDetail!RtrId = strRtrId
                    rstDetail!ItemId = Trim(!ReferencesNumber)
                    
                    rstDetail!CreateId = mdlGlobal.UserAuthority.UserId
                    rstDetail!CreateDate = mdlProcedures.FormatDate(Now)
                End If
                
                mdlTransaction.UpdateStock _
                    Trim(!ReferencesNumber), _
                    mdlDatabase.GetFieldData(mdlGlobal.conInventory, "WarehouseId", mdlTable.CreateTHDOBUY, "DOId='" & _
                        mdlDatabase.GetFieldData(mdlGlobal.conInventory, "DOId", mdlTable.CreateTHSJBUY, "SJId='" & strSJId & "'") & "'"), _
                    strRtrId, _
                    dteRtrDate, , _
                    mdlProcedures.GetCurrency(Trim(!OptInfoFirst))

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
    
    RestoreTHRTRBUY = blnValid
End Function
