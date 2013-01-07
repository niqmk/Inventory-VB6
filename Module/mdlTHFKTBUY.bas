Attribute VB_Name = "mdlTHFKTBUY"
Option Explicit

Public Sub DeleteTHFKTBUY(ByRef rstMain As ADODB.Recordset)
    Dim strRecycleId As String
    
    strRecycleId = rstMain!FktId & Space(mdlDatabase.GetColumnSize(mdlGlobal.conInventory, "ReferencesNumber", mdlTable.CreateTHRECYCLE) - Len(rstMain!FktId))
    strRecycleId = strRecycleId & mdlProcedures.FormatDate(Now, "ddMMyyyy") & mdlProcedures.FormatDate(rstMain!FktDate, "ddMMyyyy")
    
    Dim rstTemp As ADODB.Recordset
    
    Set rstTemp = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "*", mdlTable.CreateTHRECYCLE, , "RecycleId='" & strRecycleId & "'")
    
    With rstTemp
        If Not .RecordCount > 0 Then
            .AddNew
            
            !RecycleId = strRecycleId
            !ReferencesNumber = rstMain!FktId
            !RecycleDate = mdlProcedures.FormatDate(Now)
            !ReferencesDate = mdlProcedures.FormatDate(rstMain!FktDate)
            
            !CreateId = mdlGlobal.UserAuthority.UserId
            !CreateDate = mdlProcedures.FormatDate(Now)
        End If
        
        !OptInfoFirst = rstMain!VendorId
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
    
        Set rstDetail = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "*", mdlTable.CreateTDFKTBUY, False, "FktId='" & rstMain!FktId & "'")
        
        While Not rstDetail.EOF
            If Not .RecordCount > 0 Then
                .AddNew
                
                !RecycleDtlId = strRecycleId & rstDetail!SJId
                !RecycleId = strRecycleId
                !ReferencesNumber = rstDetail!SJId
                
                !CreateId = mdlGlobal.UserAuthority.UserId
                !CreateDate = mdlProcedures.FormatDate(Now)
            End If
            
            !OptInfoFirst = ""
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

Public Function RestoreTHFKTBUY(ByVal strRecycleId As String) As Boolean
    If Not mdlGlobal.UserAuthority.IsMenuAccess(mdiMain.mnuTHFKTBUY.Name) Then
        RestoreTHFKTBUY = False
        
        Exit Function
    End If
    
    Dim blnValid As Boolean
    
    blnValid = True
    
    Dim rstTemp As ADODB.Recordset
    
    Set rstTemp = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "*", mdlTable.CreateTHRECYCLE, False, "RecycleId='" & strRecycleId & "'")
    
    If Not mdlDatabase.IsDataExists(mdlGlobal.conInventory, mdlTable.CreateTMVENDOR, "VendorId='" & Trim(rstTemp!OptInfoFirst) & "'") Then
        blnValid = False
    End If
    
    Dim strFktId As String
    
    If blnValid Then
        With rstTemp
            Dim rstHeader As ADODB.Recordset
            
            Set rstHeader = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "*", mdlTable.CreateTHFKTBUY, , "FktId='" & Trim(!ReferencesNumber) & "'")
            
            If rstHeader.RecordCount > 0 Then
                blnValid = False
            Else
                rstHeader.AddNew
                
                rstHeader!FktId = Trim(!ReferencesNumber)
                
                strFktId = rstHeader!FktId
                
                rstHeader!FktDate = mdlProcedures.FormatDate(!ReferencesDate)
                rstHeader!VendorId = Trim(!OptInfoFirst)
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
            Dim strFktDtlId As String
            
            Dim rstDetail As ADODB.Recordset
            
            While Not .EOF
                strFktDtlId = strFktId & Trim(!ReferencesNumber)
                
                Set rstDetail = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "*", mdlTable.CreateTDFKTBUY, , "FktDtlId='" & strFktDtlId & "'")
                
                If Not rstDetail.RecordCount > 0 Then
                    rstDetail.AddNew
                    
                    rstDetail!FktDtlId = strFktDtlId
                    rstDetail!FktId = strFktId
                    rstDetail!SJId = Trim(!ReferencesNumber)
                    
                    rstDetail!CreateId = mdlGlobal.UserAuthority.UserId
                    rstDetail!CreateDate = mdlProcedures.FormatDate(Now)
                End If
                
                rstDetail!UpdateId = mdlGlobal.UserAuthority.UserId
                rstDetail!UpdateDate = mdlProcedures.FormatDate(Now)
                
                rstDetail.Update
                
                .MoveNext
            Wend
            
            mdlDatabase.CloseRecordset rstDetail
        End With
    End If
    
    mdlDatabase.CloseRecordset rstTemp
    
    RestoreTHFKTBUY = blnValid
End Function
