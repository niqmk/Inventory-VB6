Attribute VB_Name = "mdlTHPOBUY"
Option Explicit

Public Function GetTotalQtyPOBUY(ByVal strPOId As String, Optional ByVal strItemId As String = "") As Currency
    Dim strCriteria As String
    
    strCriteria = "POId='" & strPOId & "'"
    
    If Not Trim(strItemId) = "" Then
        strCriteria = strCriteria & " AND ItemId='" & strItemId & "'"
    End If
    
    Dim rstTemp As ADODB.Recordset
    
    Set rstTemp = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "Qty", mdlTable.CreateTDPOBUY, False, strCriteria)
    
    Dim curQty As Currency
    
    curQty = 0
    
    With rstTemp
        While Not .EOF
            curQty = curQty + mdlProcedures.GetCurrency(!Qty)
            
            .MoveNext
        Wend
    End With
    
    mdlDatabase.CloseRecordset rstTemp
    
    GetTotalQtyPOBUY = curQty
End Function

Public Sub DeleteTHPOBUY(ByRef rstMain As ADODB.Recordset)
    Dim strRecycleId As String
    
    strRecycleId = rstMain!POId & Space(mdlDatabase.GetColumnSize(mdlGlobal.conInventory, "ReferencesNumber", mdlTable.CreateTHRECYCLE) - Len(rstMain!POId))
    strRecycleId = strRecycleId & mdlProcedures.FormatDate(Now, "ddMMyyyy") & mdlProcedures.FormatDate(rstMain!SODate, "ddMMyyyy")
    
    Dim rstTemp As ADODB.Recordset
    
    Set rstTemp = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "*", mdlTable.CreateTHRECYCLE, , "RecycleId='" & strRecycleId & "'")
    
    With rstTemp
        If Not .RecordCount > 0 Then
            .AddNew
            
            !RecycleId = strRecycleId
            !ReferencesNumber = rstMain!POId
            !RecycleDate = mdlProcedures.FormatDate(Now)
            !ReferencesDate = mdlProcedures.FormatDate(rstMain!PODate)
            
            !CreateId = mdlGlobal.UserAuthority.UserId
            !CreateDate = mdlProcedures.FormatDate(Now)
        End If
        
        !OptInfoFirst = rstMain!DateLine
        !OptInfoSecond = rstMain!VendorId
        !OptInfoThird = rstMain!EmployeeBy
        !OptInfoFourth = rstMain!EmployeeAgree
        !OptInfoFifth = rstMain!Notes
        !OptInfoSixth = rstMain!Disc
        !OptInfoSeventh = rstMain!Tax
        !OptInfoEight = rstMain!CurrencyId
        !OptInfoNineth = ""
        !OptInfoTenth = ""
        
        !UpdateId = mdlGlobal.UserAuthority.UserId
        !UpdateDate = mdlProcedures.FormatDate(Now)
        
        .Update
    End With
    
    Set rstTemp = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "*", mdlTable.CreateTDRECYCLE, , "RecycleId='" & strRecycleId & "'")
    
    With rstTemp
        Dim rstDetail As ADODB.Recordset
    
        Set rstDetail = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "*", mdlTable.CreateTDSOSELL, False, "SOId='" & rstMain!SOId & "'")
        
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
            !OptInfoSecond = rstDetail!PriceId
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

Public Function RestoreTHPOBUY(ByVal strRecycleId As String) As Boolean
    If Not mdlGlobal.UserAuthority.IsMenuAccess(mdiMain.mnuTHPOBUY.Name) Then
        RestoreTHPOBUY = False
        
        Exit Function
    End If
    
    Dim blnValid As Boolean
    
    blnValid = True
    
    Dim rstTemp As ADODB.Recordset
    
    Set rstTemp = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "*", mdlTable.CreateTHRECYCLE, False, "RecycleId='" & strRecycleId & "'")
    
    If Not mdlDatabase.IsDataExists(mdlGlobal.conInventory, mdlTable.CreateTMVENDOR, "VendorId='" & Trim(rstTemp!OptInfoSecond) & "'") Then
        blnValid = False
    End If
    
    If Not mdlDatabase.IsDataExists(mdlGlobal.conInventory, mdlTable.CreateTMEMPLOYEE, "EmployeeId='" & Trim(rstTemp!OptInfoThird) & "'") Then
        blnValid = False
    End If
    
    If Not mdlDatabase.IsDataExists(mdlGlobal.conInventory, mdlTable.CreateTMEMPLOYEE, "EmployeeId='" & Trim(rstTemp!OptInfoFourth) & "'") Then
        blnValid = False
    End If
    
    Dim strPOId As String
    
    If blnValid Then
        With rstTemp
            Dim rstHeader As ADODB.Recordset
            
            Set rstHeader = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "*", mdlTable.CreateTHPOBUY, , "POId='" & Trim(!ReferencesNumber) & "'")
            
            If rstHeader.RecordCount > 0 Then
                blnValid = False
            Else
                rstHeader.AddNew
                
                rstHeader!POId = Trim(!ReferencesNumber)
                
                strPOId = rstHeader!POId
                
                rstHeader!PODate = mdlProcedures.FormatDate(!ReferencesDate)
                rstHeader!DateLine = Trim(!OptInfoFirst)
                rstHeader!VendorId = Trim(!OptInfoSecond)
                rstHeader!EmployeeBy = Trim(!OptInfoThird)
                rstHeader!EmployeeAgree = Trim(!OptInfoFourth)
                rstHeader!Notes = Trim(!OptInfoFifth)
                rstHeader!Disc = mdlProcedures.GetCurrency(!OptInfoSixth)
                rstHeader!Tax = mdlProcedures.GetCurrency(!OptInfoSeventh)
                rstHeader!CurrencyId = Trim(!OptInfoEight)
                
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
            Dim strPODtlId As String
            
            Dim rstDetail As ADODB.Recordset
            
            While Not .EOF
                strPODtlId = strPOId & Trim(!ReferencesNumber)
                
                Set rstDetail = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "*", mdlTable.CreateTDPOBUY, , "PODtlId='" & strPODtlId & "'")
                
                If Not rstDetail.RecordCount > 0 Then
                    rstDetail.AddNew
                    
                    rstDetail!PODtlId = strPODtlId
                    rstDetail!POId = strPOId
                    rstDetail!ItemId = Trim(!ReferencesNumber)
                    
                    rstDetail!CreateId = mdlGlobal.UserAuthority.UserId
                    rstDetail!CreateDate = mdlProcedures.FormatDate(Now)
                End If
                
                rstDetail!Qty = mdlProcedures.GetCurrency(Trim(!OptInfoFirst))
                rstDetail!PriceId = Trim(!OptInfoSecond)
                
                rstDetail!UpdateId = mdlGlobal.UserAuthority.UserId
                rstDetail!UpdateDate = mdlProcedures.FormatDate(Now)
                
                rstDetail.Update
                
                .MoveNext
            Wend
            
            mdlDatabase.CloseRecordset rstDetail
        End With
    End If
    
    mdlDatabase.CloseRecordset rstTemp
    
    RestoreTHPOBUY = blnValid
End Function
