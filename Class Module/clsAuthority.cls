VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAuthority"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private strUserId As String
Private strUserName As String

Public Function IsAdmin(Optional ByVal mUserId As String = "") As Boolean
    If Trim(strUserId) = "" Then
        IsAdmin = False
        
        Exit Function
    End If
    
    Dim rstSQL As ADODB.Recordset
    
    If Trim(mUserId) = "" Then
        Set rstSQL = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "UserId, UserType", mdlTable.CreateTMUSER, False, "UserId='" & strUserId & "'")
    Else
        Set rstSQL = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "UserId, UserType", mdlTable.CreateTMUSER, False, "UserId='" & mUserId & "'")
    End If
    
    If rstSQL.RecordCount > 0 Then
        If Me.GetType(mUserId) = strAdministrator Then
            IsAdmin = True
        Else
            IsAdmin = False
        End If
    End If
    
    mdlDatabase.CloseRecordset rstSQL
End Function

Public Function IsMenuAccess(ByVal strMenuId As String, Optional ByVal mUserId As String = "") As Boolean
    If Trim(strMenuId) = "" Then
        IsMenuAccess = False
        
        Exit Function
    End If
    
    If Trim(mUserId) = "" Then
        If Trim(strUserId) = "" Then
            IsMenuAccess = False
        
            Exit Function
        End If
        
        If Me.IsAdmin Then
            IsMenuAccess = True
            
            Exit Function
        End If
    Else
        If Me.IsAdmin(mUserId) Then
            IsMenuAccess = True
            
            Exit Function
        End If
    End If
    
    Dim rstSQL As ADODB.Recordset
    
    If Trim(mUserId) = "" Then
        Set rstSQL = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "AuthorityId, UserId, MenuId, AccessYN", mdlTable.CreateTMMENUAUTHORITY, False, "UserId='" & strUserId & "' AND MenuId='" & strMenuId & "'")
    Else
        Set rstSQL = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "AuthorityId, UserId, MenuId, AccessYN", mdlTable.CreateTMMENUAUTHORITY, False, "UserId='" & mUserId & "' AND MenuId='" & strMenuId & "'")
    End If
    
    If rstSQL.RecordCount > 0 Then
        If Trim(mdlSecurity.DecryptText(Trim(rstSQL!AccessYN), rstSQL!UserId)) = mdlGlobal.strYes Then
            IsMenuAccess = True
        Else
            IsMenuAccess = False
        End If
    Else
        IsMenuAccess = False
    End If
    
    mdlDatabase.CloseRecordset rstSQL
End Function

Public Function IsMenuRoot(ByRef strMenuId() As String) As Boolean
    If Trim(strUserId) = "" Then
        IsMenuRoot = False
        
        Exit Function
    End If
    
    Dim blnFound As Boolean
    
    blnFound = False
    
    Dim intCounter As Integer
    
    For intCounter = 0 To UBound(strMenuId)
        If Me.IsMenuAccess(strMenuId(intCounter)) Then
            blnFound = True
            
            Exit For
        End If
    Next intCounter
    
    If blnFound Then
        IsMenuRoot = True
    Else
        IsMenuRoot = False
    End If
End Function

Public Function GetType(Optional ByVal mUserId As String = "") As String
    Dim rstSQL As ADODB.Recordset
    
    If Trim(mUserId) = "" Then
        Set rstSQL = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "UserId, UserType", mdlTable.CreateTMUSER, False, "UserId='" & strUserId & "'")
    Else
        Set rstSQL = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "UserId, UserType", mdlTable.CreateTMUSER, False, "UserId='" & mUserId & "'")
    End If
    
    If rstSQL.RecordCount > 0 Then
        GetType = Trim(mdlSecurity.DecryptText(Trim(rstSQL!UserType), rstSQL!UserId))
    Else
        GetType = ""
    End If
    
    mdlDatabase.CloseRecordset rstSQL
End Function

Public Sub SetLogin(ByVal mUserId As String, Optional ByVal strUserIP As String = "")
    Dim rstSQL As ADODB.Recordset
    
    Set rstSQL = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "*", mdlTable.CreateTMUSERLOGIN(True), , "UserId='" & mUserId & "'")
    
    With rstSQL
        If Not .RecordCount > 0 Then
            .AddNew
            
            !UserId = mUserId
            
            !LogoutDate = mdlProcedures.FormatDate(Now)
            
            !CreateId = Trim(mUserId)
            !CreateDate = mdlProcedures.FormatDate(Now)
        End If
        
        !LoginDate = mdlProcedures.FormatDate(Now)
        !UserIP = strUserIP
        !LogYN = mdlGlobal.strYes
        
        !UpdateId = Trim(mUserId)
        !UpdateDate = mdlProcedures.FormatDate(Now)
        
        .Update
    End With
    
    mdlDatabase.CloseRecordset rstSQL
End Sub

Public Sub SetLogout(ByVal mUserId As String, Optional ByVal strUserIP As String = "")
    Dim rstSQL As ADODB.Recordset
    
    Set rstSQL = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "*", mdlTable.CreateTMUSERLOGIN, , "UserId='" & mUserId & "'")
    
    With rstSQL
        If Not .RecordCount > 0 Then
            .AddNew
            
            !UserId = mUserId
            
            !LoginDate = mdlProcedures.FormatDate(Now)
            !UserIP = strUserIP
            
            !CreateId = Trim(mUserId)
            !CreateDate = mdlProcedures.FormatDate(Now)
        End If
        
        !LogoutDate = mdlProcedures.FormatDate(Now)
        !LogYN = mdlGlobal.strNo
        
        !UpdateId = Trim(mUserId)
        !UpdateDate = mdlProcedures.FormatDate(Now)
        
        .Update
    End With
    
    mdlDatabase.CloseRecordset rstSQL
End Sub

Public Property Get UserId() As String
    UserId = strUserId
End Property

Public Property Let UserId(ByVal mUserId As String)
    strUserId = mUserId
End Property
