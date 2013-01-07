VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFinance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private objDatabase As SQLDATABASE

Private strServer As String
Private strUserId As String
Private strPassword As String

Public Function SetConnection() As ADODB.Connection
    On Local Error GoTo ErrHandler
    
    Dim conSQL As ADODB.Connection
    
    ServerInitialize
    
    If Trim(strServer) = "" Then
        Set SetConnection = Nothing
        
        Exit Function
    End If
    
    If Trim(strUserId) = "" Then
        Set SetConnection = mdlDatabase.OpenConnection(mdlGlobal.strFinance, strServer, , strPassword, objDatabase)
    Else
        Set SetConnection = mdlDatabase.OpenConnection(mdlGlobal.strFinance, strServer, strUserId, strPassword, objDatabase)
    End If
    
ErrHandler:
End Function

Private Sub ServerInitialize()
    Dim lngRegistry As Long
    Dim lngRegKey As Long
    Dim lngType As Long
    Dim lngSize As Long
    
    On Local Error GoTo ErrHandler

    lngRegistry = _
        mdlRegistry.OpenRegistry(mdlRegistry.HKEY_CURRENT_USER, mdlRegistry.KEYS_SYS_INFO_SERVER2, lngRegKey)

    If lngRegistry = 0 Then
        Dim strDatabaseInitialize As String
        
        lngRegistry = _
            mdlRegistry.ReadValueRegistry(lngRegKey, mdlGlobal.SERVER_REGISTRY, lngType, strServer, lngSize)
        lngRegistry = _
            mdlRegistry.ReadValueRegistry(lngRegKey, mdlGlobal.USERID_REGISTRY, lngType, strUserId, lngSize)
        lngRegistry = _
            mdlRegistry.ReadValueRegistry(lngRegKey, mdlGlobal.PASSWORD_REGISTRY, lngType, strPassword, lngSize)
        lngRegistry = _
            mdlRegistry.ReadValueRegistry(lngRegKey, mdlGlobal.DATABASE_REGISTRY, lngType, strDatabaseInitialize, lngSize)
            
        strServer = mdlProcedures.RepRegistryUnknown(Trim(CStr(strServer)))
        strUserId = mdlProcedures.RepRegistryUnknown(Trim(CStr(strUserId)))
        strPassword = mdlProcedures.RepRegistryUnknown(Trim(CStr(strPassword)))
        
        If Not Trim(strUserId) = "" Then
            strUserId = Trim(mdlSecurity.DecryptText(strUserId, mdlGlobal.PUBLIC_KEY))
        End If
        
        If Not Trim(strPassword) = "" Then
            strPassword = Trim(mdlSecurity.DecryptText(strPassword, mdlGlobal.PUBLIC_KEY))
        End If
        
        objDatabase = CInt(mdlProcedures.RepRegistryUnknown(Trim(CStr(strDatabaseInitialize))))
    End If

    lngRegistry = mdlRegistry.CloseRegistry(lngRegKey)
    
    Exit Sub

ErrHandler:
End Sub

Public Function CreateTMEMPLOYEE(ByRef conSQL As ADODB.Connection) As String
    ServerInitialize
    
    Dim strTable As String
    
    strTable = "TMEMPLOYEE"
    
    CreateTMEMPLOYEE = strTable

    Dim strQuery As String
    
    If objDatabase = SQLSERVER7 Or _
        objDatabase = SQLSERVER2000 Or _
        objDatabase = SQLEXPRESS Then
        strQuery = "CREATE TABLE " & strTable & " (" & _
            "EmployeeId char(7) PRIMARY KEY, " & _
            "EmployeeDate datetime NULL DEFAULT GETDATE(), " & _
            "Name char(50) NULL DEFAULT '', " & _
            "JobTypeId char(4) NULL DEFAULT '', " & _
            "Address varchar(150) NULL DEFAULT '', " & _
            "Phone char(50) NULL DEFAULT '', " & _
            "HandPhone char(50) NULL DEFAULT '', " & _
            "Fax char(50) NULL DEFAULT '', " & _
            "Email char(50) NULL DEFAULT '', " & _
            "Notes varchar(150) NULL DEFAULT '', " & _
            "CreateId char(20) NOT NULL, " & _
            "CreateDate datetime NULL DEFAULT GETDATE(), " & _
            "UpdateId char(20) NOT NULL, " & _
            "UpdateDate datetime NULL DEFAULT GETDATE())"
    ElseIf mdlGlobal.objDatabaseInit = MSACCESS Then
        strQuery = "CREATE TABLE " & strTable & " (" & _
            "EmployeeId char(7) PRIMARY KEY, " & _
            "EmployeeDate datetime NOT NULL, " & _
            "Name char(50) NOT NULL, " & _
            "JobTypeId char(4) NOT NULL, " & _
            "Address varchar(150) NOT NULL, " & _
            "Phone char(50) NOT NULL, " & _
            "HandPhone char(50) NOT NULL, " & _
            "Fax char(50) NOT NULL, " & _
            "Email char(50) NOT NULL, " & _
            "Notes varchar(150) NOT NULL, " & _
            "CreateId char(20) NOT NULL, " & _
            "CreateDate datetime NOT NULL, " & _
            "UpdateId char(20) NOT NULL, " & _
            "UpdateDate datetime NOT NULL)"
    ElseIf mdlGlobal.objDatabaseInit = MYSQL Then
        strQuery = "CREATE TABLE `" & strTable & "` (" & _
            strQuery = "CREATE TABLE `" & strTable & "` (" & _
            "`EmployeeId` char(7) NOT NULL, " & _
            "`EmployeeDate` datetime NOT NULL, " & _
            "`Name` char(50) NOT NULL, " & _
            "`JobTypeId` char(4) NOT NULL, " & _
            "`Address` varchar(150) NOT NULL, " & _
            "`Phone` char(50) NOT NULL, " & _
            "`HandPhone` char(50) NOT NULL, " & _
            "`Fax` char(50) NOT NULL, " & _
            "`Email` char(50) NOT NULL, " & _
            "`Notes` varchar(150) NOT NULL, " & _
            "`CreateId` char(8) NOT NULL, " & _
            "`CreateDate` datetime NOT NULL, " & _
            "`UpdateId` char(8) NOT NULL, " & _
            "`UpdateDate` datetime NOT NULL, " & _
            "PRIMARY KEY(`EmployeeId`)) TYPE=MyISAM;"
    End If
        
    mdlDatabase.CreateTable conSQL, strQuery, strTable, objDatabase
End Function
