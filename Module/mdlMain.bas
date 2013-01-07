Attribute VB_Name = "mdlMain"
Option Explicit

Private Enum RegistryMode
    [Notesxists]
    [Exists]
    [ErrorReg]
End Enum

Public Sub Main()
    SetVariable
    
    Dim objMode As RegistryMode
    
    objMode = ServerRegistryInitialize
    
    If objMode = Notesxists Then
        frmServer.Show vbModal
        
        objMode = ServerRegistryInitialize
        
        If objMode = Notesxists Then Exit Sub
    End If
    
    If objDatabaseInit = MSACCESS Then
        If Not mdlGlobal.fso.FileExists(mdlGlobal.strServerInit) Then
            Exit Sub
        End If
    End If
    
    If Trim(mdlGlobal.strUserIdInit) = "" Then
        If Not mdlDatabase.CreateDatabase(mdlGlobal.strInventory, mdlGlobal.strServerInit, , mdlGlobal.strPasswordInit, mdlGlobal.objDatabaseInit) Then Exit Sub
        
        Set mdlGlobal.conInventory = mdlDatabase.OpenConnection(mdlGlobal.strInventory, mdlGlobal.strServerInit, , mdlGlobal.strPasswordInit, mdlGlobal.objDatabaseInit)
    Else
        If Not mdlDatabase.CreateDatabase(mdlGlobal.strInventory, mdlGlobal.strServerInit, mdlGlobal.strUserIdInit, mdlGlobal.strPasswordInit, mdlGlobal.objDatabaseInit) Then Exit Sub
        
        Set mdlGlobal.conInventory = mdlDatabase.OpenConnection(mdlGlobal.strInventory, mdlGlobal.strServerInit, mdlGlobal.strUserIdInit, mdlGlobal.strPasswordInit, mdlGlobal.objDatabaseInit)
    End If
    
    Dim clsFinance As New clsFinance
    Set mdlGlobal.conFinance = clsFinance.SetConnection
    Set clsFinance = Nothing
    
    Dim clsAccounting As New clsAccounting
    Set mdlGlobal.conAccounting = clsAccounting.SetConnection
    Set clsAccounting = Nothing
    
    objMode = RegistryInitialize
    
    If objMode = Exists Then
        frmLogin.Show
    ElseIf objMode = Notesxists Then
        frmProfile.Show vbModal
        
        If RegistryInitialize = Exists Then
            Set mdlGlobal.conInventory = mdlDatabase.OpenConnection(mdlGlobal.strInventory, mdlGlobal.strServerInit, , , objDatabaseInit)
            
            frmLogin.Show
        Else
            Set mdlGlobal.fso = Nothing
        End If
    End If
End Sub

Private Sub SetVariable()
    If Right(App.Path, 1) = "\" Then
        mdlGlobal.strPath = App.Path
    Else
        mdlGlobal.strPath = App.Path & "\"
    End If
    
    If mdlProcedures.IsValidDateRegion Then
        mdlGlobal.strFormatDate = "MM/dd/yyyy"
    Else
        mdlGlobal.strFormatDate = "dd/MM/yyyy"
    End If
    
    Set mdlGlobal.fso = New FileSystemObject
End Sub

Private Function ServerRegistryInitialize() As RegistryMode
    Dim lngRegistry As Long
    Dim lngRegKey As Long
    Dim lngType As Long
    Dim lngSize As Long
    
    On Local Error GoTo ErrHandler

    lngRegistry = _
        mdlRegistry.OpenRegistry(mdlRegistry.HKEY_CURRENT_USER, mdlRegistry.KEYS_SYS_INFO_SERVER1, lngRegKey)

    If lngRegistry = 0 Then
        Dim strDatabaseInitialize As String
        
        lngRegistry = _
            mdlRegistry.ReadValueRegistry(lngRegKey, mdlGlobal.SERVER_REGISTRY, lngType, mdlGlobal.strServerInit, lngSize)
        lngRegistry = _
            mdlRegistry.ReadValueRegistry(lngRegKey, mdlGlobal.USERID_REGISTRY, lngType, mdlGlobal.strUserIdInit, lngSize)
        lngRegistry = _
            mdlRegistry.ReadValueRegistry(lngRegKey, mdlGlobal.PASSWORD_REGISTRY, lngType, mdlGlobal.strPasswordInit, lngSize)
        lngRegistry = _
            mdlRegistry.ReadValueRegistry(lngRegKey, mdlGlobal.DATABASE_REGISTRY, lngType, strDatabaseInitialize, lngSize)
            
        mdlGlobal.strServerInit = mdlProcedures.RepRegistryUnknown(Trim(CStr(mdlGlobal.strServerInit)))
        mdlGlobal.strUserIdInit = mdlProcedures.RepRegistryUnknown(Trim(CStr(mdlGlobal.strUserIdInit)))
        mdlGlobal.strPasswordInit = mdlProcedures.RepRegistryUnknown(Trim(CStr(mdlGlobal.strPasswordInit)))
        
        If Not Trim(mdlGlobal.strUserIdInit) = "" Then
            mdlGlobal.strUserIdInit = Trim(mdlSecurity.DecryptText(mdlGlobal.strUserIdInit, mdlGlobal.PUBLIC_KEY))
        End If
        
        If Not Trim(mdlGlobal.strPasswordInit) = "" Then
            mdlGlobal.strPasswordInit = Trim(mdlSecurity.DecryptText(mdlGlobal.strPasswordInit, mdlGlobal.PUBLIC_KEY))
        End If
        
        mdlGlobal.objDatabaseInit = CInt(mdlProcedures.RepRegistryUnknown(Trim(CStr(strDatabaseInitialize))))
        
        ServerRegistryInitialize = Exists
    Else
        ServerRegistryInitialize = Notesxists
    End If

    lngRegistry = mdlRegistry.CloseRegistry(lngRegKey)
    
    Exit Function

ErrHandler:
    ServerRegistryInitialize = ErrorReg
End Function

Private Function RegistryInitialize() As RegistryMode
    Dim lngRegistry As Long
    Dim lngRegKey As Long
    Dim lngType As Long
    Dim lngSize As Long
    
    On Local Error GoTo ErrHandler

    lngRegistry = _
        mdlRegistry.OpenRegistry(mdlRegistry.HKEY_CURRENT_USER, mdlRegistry.KEYS_SYS_INFO, lngRegKey)

    If lngRegistry = 0 Then
        lngRegistry = _
            mdlRegistry.ReadValueRegistry(lngRegKey, mdlGlobal.COMPANY_REGISTRY, lngType, mdlGlobal.strCompanyText, lngSize)
        lngRegistry = _
            mdlRegistry.ReadValueRegistry(lngRegKey, mdlGlobal.ADDRESS_REGISTRY, lngType, mdlGlobal.strAddressText, lngSize)
        lngRegistry = _
            mdlRegistry.ReadValueRegistry(lngRegKey, mdlGlobal.WEBSITE_REGISTRY, lngType, mdlGlobal.strWebsiteText, lngSize)
        lngRegistry = _
            mdlRegistry.ReadValueRegistry(lngRegKey, mdlGlobal.EMAIL_REGISTRY, lngType, mdlGlobal.strEmailText, lngSize)
        lngRegistry = _
            mdlRegistry.ReadValueRegistry(lngRegKey, mdlGlobal.PHONE_REGISTRY, lngType, mdlGlobal.strPhoneText, lngSize)
        lngRegistry = _
            mdlRegistry.ReadValueRegistry(lngRegKey, mdlGlobal.FAX_REGISTRY, lngType, mdlGlobal.strFaxText, lngSize)
        lngRegistry = _
            mdlRegistry.ReadValueRegistry(lngRegKey, mdlGlobal.NPWP_REGISTRY, lngType, mdlGlobal.strNPWPText, lngSize)
            
        mdlGlobal.strCompanyText = mdlProcedures.RepRegistryUnknown(Trim(CStr(mdlGlobal.strCompanyText)))
        mdlGlobal.strAddressText = mdlProcedures.RepRegistryUnknown(Trim(CStr(mdlGlobal.strAddressText)))
        mdlGlobal.strWebsiteText = mdlProcedures.RepRegistryUnknown(Trim(CStr(mdlGlobal.strWebsiteText)))
        mdlGlobal.strEmailText = mdlProcedures.RepRegistryUnknown(Trim(CStr(mdlGlobal.strEmailText)))
        mdlGlobal.strPhoneText = mdlProcedures.RepRegistryUnknown(Trim(CStr(mdlGlobal.strPhoneText)))
        mdlGlobal.strFaxText = mdlProcedures.RepRegistryUnknown(Trim(CStr(mdlGlobal.strFaxText)))
        mdlGlobal.strNPWPText = mdlProcedures.RepRegistryUnknown(Trim(CStr(mdlGlobal.strNPWPText)))
        
        RegistryInitialize = Exists
    Else
        RegistryInitialize = Notesxists
    End If

    lngRegistry = mdlRegistry.CloseRegistry(lngRegKey)
    
    Exit Function

ErrHandler:
    RegistryInitialize = ErrorReg
End Function
