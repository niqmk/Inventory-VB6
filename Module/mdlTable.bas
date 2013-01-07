Attribute VB_Name = "mdlTable"
Option Explicit

Public Function CreateTMUSER(Optional ByVal blnCreate As Boolean = False) As String
    Dim strTable As String
    
    strTable = "TMUSER"
    
    CreateTMUSER = strTable
    
    If blnCreate Then
        Dim strQuery As String
        
        If mdlGlobal.objDatabaseInit = SQLSERVER7 Or _
            mdlGlobal.objDatabaseInit = SQLSERVER2000 Or _
            mdlGlobal.objDatabaseInit = SQLEXPRESS Then
            strQuery = "CREATE TABLE " & strTable & " (" & _
                "UserId char(8) PRIMARY KEY, " & _
                "UserName char(50) NOT NULL, " & _
                "UserPwd char(24) NOT NULL, " & _
                "UserType char(16) NOT NULL, " & _
                "CreateId char(8) NOT NULL, " & _
                "CreateDate datetime NULL DEFAULT GETDATE(), " & _
                "UpdateId char(8) NOT NULL, " & _
                "UpdateDate datetime NULL DEFAULT GETDATE())"
        ElseIf mdlGlobal.objDatabaseInit = MSACCESS Then
            strQuery = "CREATE TABLE " & strTable & " (" & _
                "UserId char(8) PRIMARY KEY, " & _
                "UserName char(50) NOT NULL, " & _
                "UserPwd char(24) NOT NULL, " & _
                "UserType char(16) NOT NULL, " & _
                "CreateId char(8) NOT NULL, " & _
                "CreateDate datetime NOT NULL, " & _
                "UpdateId char(8) NOT NULL, " & _
                "UpdateDate datetime NOT NULL)"
        ElseIf mdlGlobal.objDatabaseInit = MYSQL Then
            strQuery = "CREATE TABLE `" & strTable & "` (" & _
                "`UserId` char(8) NOT NULL, " & _
                "`UserName` char(50) NOT NULL, " & _
                "`UserPwd` char(24) NOT NULL, " & _
                "`UserType` char(16) NOT NULL, " & _
                "`CreateId` char(8) NOT NULL, " & _
                "`CreateDate` datetime NOT NULL, " & _
                "`UpdateId` char(8) NOT NULL, " & _
                "`UpdateDate` datetime NOT NULL, " & _
                "PRIMARY KEY(`UserId`)) TYPE=MyISAM;"
        End If
            
        mdlDatabase.CreateTable mdlGlobal.conInventory, strQuery, strTable, mdlGlobal.objDatabaseInit
    End If
End Function

Public Function CreateTMUSERLOGIN(Optional ByVal blnCreate As Boolean = False) As String
    Dim strTable As String
    
    strTable = "TMUSERLOGIN"
    
    CreateTMUSERLOGIN = strTable
    
    If blnCreate Then
        Dim strQuery As String
        
        If mdlGlobal.objDatabaseInit = SQLSERVER7 Or _
            mdlGlobal.objDatabaseInit = SQLSERVER2000 Or _
            mdlGlobal.objDatabaseInit = SQLEXPRESS Then
            strQuery = "CREATE TABLE " & strTable & " (" & _
                "UserId char(8) PRIMARY KEY, " & _
                "LoginDate datetime NULL DEFAULT GETDATE(), " & _
                "LogoutDate datetime NULL DEFAULT GETDATE(), " & _
                "UserIP char(15) NOT NULL, " & _
                "LogYN char(1) NOT NULL, " & _
                "CreateId char(8) NOT NULL, " & _
                "CreateDate datetime NULL DEFAULT GETDATE(), " & _
                "UpdateId char(8) NOT NULL, " & _
                "UpdateDate datetime NULL DEFAULT GETDATE())"
        ElseIf mdlGlobal.objDatabaseInit = MSACCESS Then
            strQuery = "CREATE TABLE " & strTable & " (" & _
                "UserId char(8) PRIMARY KEY, " & _
                "LoginDate datetime NOT NULL, " & _
                "LogoutDate datetime NOT NULL, " & _
                "UserIP char(15) NOT NULL, " & _
                "LogYN char(1) NOT NULL, " & _
                "CreateId char(8) NOT NULL, " & _
                "CreateDate datetime NOT NULL, " & _
                "UpdateId char(8) NOT NULL, " & _
                "UpdateDate datetime NOT NULL)"
        ElseIf mdlGlobal.objDatabaseInit = MYSQL Then
            strQuery = "CREATE TABLE `" & strTable & "` (" & _
                "`UserId` char(8) NOT NULL, " & _
                "`LoginDate` datetime NOT NULL, " & _
                "`LogoutDate` datetime NOT NULL, " & _
                "`UserIP` char(15) NOT NULL, " & _
                "`LogYN` char(1) NOT NULL, " & _
                "`CreateId` char(8) NOT NULL, " & _
                "`CreateDate` datetime NOT NULL, " & _
                "`UpdateId` char(8) NOT NULL, " & _
                "`UpdateDate` datetime NOT NULL, " & _
                "PRIMARY KEY(`UserId`)) TYPE=MyISAM;"
        End If
            
        mdlDatabase.CreateTable mdlGlobal.conInventory, strQuery, strTable, mdlGlobal.objDatabaseInit
    End If
End Function

Public Function CreateTMMENU(Optional ByVal blnCreate As Boolean = False) As String
    Dim strTable As String
    
    strTable = "TMMENU"
    
    CreateTMMENU = strTable
    
    If mdlGlobal.UserAuthority.IsAdmin Or blnCreate Then
        Dim strQuery As String
        
        If mdlGlobal.objDatabaseInit = SQLSERVER7 Or _
            mdlGlobal.objDatabaseInit = SQLSERVER2000 Or _
            mdlGlobal.objDatabaseInit = SQLEXPRESS Then
            strQuery = "CREATE TABLE " & strTable & " (" & _
                "MenuId char(30) PRIMARY KEY, " & _
                "MenuName char(200) NOT NULL, " & _
                "CreateId char(8) NOT NULL, " & _
                "CreateDate datetime NULL DEFAULT GETDATE(), " & _
                "UpdateId char(8) NOT NULL, " & _
                "UpdateDate datetime NULL DEFAULT GETDATE())"
        ElseIf mdlGlobal.objDatabaseInit = MSACCESS Then
            strQuery = "CREATE TABLE " & strTable & " (" & _
                "MenuId char(30) PRIMARY KEY, " & _
                "MenuName char(200) NOT NULL, " & _
                "CreateId char(8) NOT NULL, " & _
                "CreateDate datetime NOT NULL, " & _
                "UpdateId char(8) NOT NULL, " & _
                "UpdateDate datetime NOT NULL)"
        ElseIf mdlGlobal.objDatabaseInit = MYSQL Then
            strQuery = "CREATE TABLE `" & strTable & "` (" & _
                "`MenuId` char(30) NOT NULL, " & _
                "`MenuName` char(200) NOT NULL, " & _
                "`CreateId` char(8) NOT NULL, " & _
                "`CreateDate` datetime NOT NULL, " & _
                "`UpdateId` char(8) NOT NULL, " & _
                "`UpdateDate` datetime NOT NULL, " & _
                "PRIMARY KEY(`MenuId`)) TYPE=MyISAM;"
        End If
            
        mdlDatabase.CreateTable mdlGlobal.conInventory, strQuery, strTable, mdlGlobal.objDatabaseInit
    End If
End Function

Public Function CreateTMMENUAUTHORITY(Optional ByVal blnCreate As Boolean = False) As String
    Dim strTable As String
    
    strTable = "TMMENUAUTHORITY"
    
    CreateTMMENUAUTHORITY = strTable
    
    If mdlGlobal.UserAuthority.IsAdmin Or blnCreate Then
        Dim strQuery As String
        
        If mdlGlobal.objDatabaseInit = SQLSERVER7 Or _
            mdlGlobal.objDatabaseInit = SQLSERVER2000 Or _
            mdlGlobal.objDatabaseInit = SQLEXPRESS Then
            strQuery = "CREATE TABLE " & strTable & " (" & _
                "AuthorityId char(50) PRIMARY KEY, " & _
                "UserId char(8) NOT NULL, " & _
                "MenuId char(30) NOT NULL, " & _
                "AccessYN char(55) NOT NULL, " & _
                "CreateId char(8) NOT NULL, " & _
                "CreateDate datetime NULL DEFAULT GETDATE(), " & _
                "UpdateId char(8) NOT NULL, " & _
                "UpdateDate datetime NULL DEFAULT GETDATE())"
        ElseIf mdlGlobal.objDatabaseInit = MSACCESS Then
            strQuery = "CREATE TABLE " & strTable & " (" & _
                "AuthorityId char(50) PRIMARY KEY, " & _
                "UserId char(8) NOT NULL, " & _
                "MenuId char(30) NOT NULL, " & _
                "AccessYN char(55) NOT NULL, " & _
                "CreateId char(8) NOT NULL, " & _
                "CreateDate datetime NOT NULL, " & _
                "UpdateId char(8) NOT NULL, " & _
                "UpdateDate datetime NOT NULL)"
        ElseIf mdlGlobal.objDatabaseInit = MYSQL Then
            strQuery = "CREATE TABLE `" & strTable & "` (" & _
                "`AuthorityId` char(50) NOT NULL, " & _
                "`UserId` char(8) NOT NULL, " & _
                "`MenuId` char(30) NOT NULL, " & _
                "`AccessYN` char(55) NOT NULL, " & _
                "`CreateId` char(8) NOT NULL, " & _
                "`CreateDate` datetime NOT NULL, " & _
                "`UpdateId` char(8) NOT NULL, " & _
                "`UpdateDate` datetime NOT NULL, " & _
                "PRIMARY KEY(`AuthorityId`)) TYPE=MyISAM;"
        End If
            
        mdlDatabase.CreateTable mdlGlobal.conInventory, strQuery, strTable, mdlGlobal.objDatabaseInit
    End If
End Function

Public Function CreateTMJOBTYPE(Optional ByVal blnCreate As Boolean = False) As String
    Dim strTable As String
    
    strTable = "TMJOBTYPE"
    
    CreateTMJOBTYPE = strTable
    
    If mdlGlobal.UserAuthority.IsAdmin Or blnCreate Then
        Dim strQuery As String
        
        If mdlGlobal.objDatabaseInit = SQLSERVER7 Or _
            mdlGlobal.objDatabaseInit = SQLSERVER2000 Or _
            mdlGlobal.objDatabaseInit = SQLEXPRESS Then
            strQuery = "CREATE TABLE " & strTable & " (" & _
                "JobTypeId char(4) PRIMARY KEY, " & _
                "Name char(50) NULL DEFAULT '', " & _
                "DivisionId char(4) NULL DEFAULT '', " & _
                "Notes varchar(150) NULL DEFAULT '', " & _
                "CreateId char(8) NOT NULL, " & _
                "CreateDate datetime NULL DEFAULT GETDATE(), " & _
                "UpdateId char(8) NOT NULL, " & _
                "UpdateDate datetime NULL DEFAULT GETDATE())"
        ElseIf mdlGlobal.objDatabaseInit = MSACCESS Then
            strQuery = "CREATE TABLE " & strTable & " (" & _
                "JobTypeId char(4) PRIMARY KEY, " & _
                "Name char(50) NOT NULL, " & _
                "DivisionId char(4) NOT NULL, " & _
                "Notes varchar(150) NOT NULL, " & _
                "CreateId char(8) NOT NULL, " & _
                "CreateDate datetime NOT NULL, " & _
                "UpdateId char(8) NOT NULL, " & _
                "UpdateDate datetime NOT NULL)"
        ElseIf mdlGlobal.objDatabaseInit = MYSQL Then
            strQuery = "CREATE TABLE `" & strTable & "` (" & _
                "`JobTypeId` char(4) NOT NULL, " & _
                "`Name` char(50) NOT NULL, " & _
                "`DivisionId` char(4) NOT NULL, " & _
                "`Notes` varchar(150) NOT NULL, " & _
                "`CreateId` char(8) NOT NULL, " & _
                "`CreateDate` datetime NOT NULL, " & _
                "`UpdateId` char(8) NOT NULL, " & _
                "`UpdateDate` datetime NOT NULL, " & _
                "PRIMARY KEY(`JobTypeId`)) TYPE=MyISAM;"
        End If
            
        mdlDatabase.CreateTable mdlGlobal.conInventory, strQuery, strTable, mdlGlobal.objDatabaseInit
    End If
End Function

Public Function CreateTMDIVISION(Optional ByVal blnCreate As Boolean = False) As String
    Dim strTable As String
    
    strTable = "TMDIVISION"
    
    CreateTMDIVISION = strTable
    
    If mdlGlobal.UserAuthority.IsAdmin Or blnCreate Then
        Dim strQuery As String
        
        If mdlGlobal.objDatabaseInit = SQLSERVER7 Or _
            mdlGlobal.objDatabaseInit = SQLSERVER2000 Or _
            mdlGlobal.objDatabaseInit = SQLEXPRESS Then
            strQuery = "CREATE TABLE " & strTable & " (" & _
                "DivisionId char(4) PRIMARY KEY, " & _
                "Name char(50) NULL DEFAULT '', " & _
                "Notes varchar(150) NULL DEFAULT '', " & _
                "CreateId char(8) NOT NULL, " & _
                "CreateDate datetime NULL DEFAULT GETDATE(), " & _
                "UpdateId char(8) NOT NULL, " & _
                "UpdateDate datetime NULL DEFAULT GETDATE())"
        ElseIf mdlGlobal.objDatabaseInit = MSACCESS Then
            strQuery = "CREATE TABLE " & strTable & " (" & _
                "DivisionId char(4) PRIMARY KEY, " & _
                "Name char(50) NOT NULL, " & _
                "Notes varchar(150) NOT NULL, " & _
                "CreateId char(8) NOT NULL, " & _
                "CreateDate datetime NOT NULL, " & _
                "UpdateId char(8) NOT NULL, " & _
                "UpdateDate datetime NOT NULL)"
        ElseIf mdlGlobal.objDatabaseInit = MYSQL Then
            strQuery = "CREATE TABLE `" & strTable & "` (" & _
                "`DivisionId` char(4) NOT NULL, " & _
                "`Name` char(50) NOT NULL, " & _
                "`Notes` varchar(150) NOT NULL, " & _
                "`CreateId` char(8) NOT NULL, " & _
                "`CreateDate` datetime NOT NULL, " & _
                "`UpdateId` char(8) NOT NULL, " & _
                "`UpdateDate` datetime NOT NULL, " & _
                "PRIMARY KEY(`DivisionId`)) TYPE=MyISAM;"
        End If
            
        mdlDatabase.CreateTable mdlGlobal.conInventory, strQuery, strTable, mdlGlobal.objDatabaseInit
    End If
End Function

Public Function CreateTMEMPLOYEE(Optional ByVal blnCreate As Boolean = False) As String
    Dim strTable As String
    
    strTable = "TMEMPLOYEE"
    
    CreateTMEMPLOYEE = strTable
    
    If mdlGlobal.UserAuthority.IsAdmin Or blnCreate Then
        Dim strQuery As String
        
        If mdlGlobal.objDatabaseInit = SQLSERVER7 Or _
            mdlGlobal.objDatabaseInit = SQLSERVER2000 Or _
            mdlGlobal.objDatabaseInit = SQLEXPRESS Then
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
                "CreateId char(8) NOT NULL, " & _
                "CreateDate datetime NULL DEFAULT GETDATE(), " & _
                "UpdateId char(8) NOT NULL, " & _
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
                "CreateId char(8) NOT NULL, " & _
                "CreateDate datetime NOT NULL, " & _
                "UpdateId char(8) NOT NULL, " & _
                "UpdateDate datetime NOT NULL)"
        ElseIf mdlGlobal.objDatabaseInit = MYSQL Then
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
            
        mdlDatabase.CreateTable mdlGlobal.conInventory, strQuery, strTable, mdlGlobal.objDatabaseInit
    End If
End Function

Public Function CreateTMVENDOR(Optional ByVal blnCreate As Boolean = False) As String
    Dim strTable As String
    
    strTable = "TMVENDOR"
    
    CreateTMVENDOR = strTable
    
    If mdlGlobal.UserAuthority.IsAdmin Or blnCreate Then
        Dim strQuery As String
        
        If mdlGlobal.objDatabaseInit = SQLSERVER7 Or _
            mdlGlobal.objDatabaseInit = SQLSERVER2000 _
            Or mdlGlobal.objDatabaseInit = SQLEXPRESS Then
            strQuery = "CREATE TABLE " & strTable & " (" & _
                "VendorId char(6) PRIMARY KEY, " & _
                "Name char(50) NULL DEFAULT '', " & _
                "Address varchar(150) NULL DEFAULT '', " & _
                "Website char(50) NULL DEFAULT '', " & _
                "Email char(50) NULL DEFAULT '', " & _
                "Phone char(50) NULL DEFAULT '', " & _
                "Fax char(50) NULL DEFAULT '', " & _
                "Notes varchar(150) NULL DEFAULT '', " & _
                "CreateId char(8) NOT NULL, " & _
                "CreateDate datetime NULL DEFAULT GETDATE(), " & _
                "UpdateId char(8) NOT NULL, " & _
                "UpdateDate datetime NULL DEFAULT GETDATE())"
        ElseIf mdlGlobal.objDatabaseInit = MSACCESS Then
            strQuery = "CREATE TABLE " & strTable & " (" & _
                "VendorId char(6) PRIMARY KEY, " & _
                "Name char(50) NOT NULL, " & _
                "Address varchar(150) NOT NULL, " & _
                "Website char(50) NOT NULL, " & _
                "Email char(50) NOT NULL, " & _
                "Phone char(50) NOT NULL, " & _
                "Fax char(50) NOT NULL, " & _
                "Notes varchar(150) NOT NULL, " & _
                "CreateId char(8) NOT NULL, " & _
                "CreateDate datetime NOT NULL, " & _
                "UpdateId char(8) NOT NULL, " & _
                "UpdateDate datetime NOT NULL)"
        ElseIf mdlGlobal.objDatabaseInit = MYSQL Then
            strQuery = "CREATE TABLE `" & strTable & "` (" & _
                "`VendorId` char(6) NOT NULL, " & _
                "`Name` char(50) NOT NULL, " & _
                "`Address` varchar(150) NOT NULL, " & _
                "`Website` char(50) NOT NULL, " & _
                "`Email` char(50) NOT NULL, " & _
                "`Phone` char(50) NOT NULL, " & _
                "`Fax` char(50) NOT NULL, " & _
                "`Notes` varchar(150) NOT NULL, " & _
                "`CreateId` char(8) NOT NULL, " & _
                "`CreateDate` datetime NOT NULL, " & _
                "`UpdateId` char(8) NOT NULL, " & _
                "`UpdateDate` datetime NOT NULL, " & _
                "PRIMARY KEY(`VendorId`)) TYPE=MyISAM;"
        End If
            
        mdlDatabase.CreateTable mdlGlobal.conInventory, strQuery, strTable, mdlGlobal.objDatabaseInit
    End If
End Function

Public Function CreateTMWAREHOUSE(Optional ByVal blnCreate As Boolean = False) As String
    Dim strTable As String
    
    strTable = "TMWAREHOUSE"
    
    CreateTMWAREHOUSE = strTable
    
    If mdlGlobal.UserAuthority.IsAdmin Or blnCreate Then
        Dim strQuery As String
        
        If mdlGlobal.objDatabaseInit = SQLSERVER7 Or _
            mdlGlobal.objDatabaseInit = SQLSERVER2000 _
            Or mdlGlobal.objDatabaseInit = SQLEXPRESS Then
            strQuery = "CREATE TABLE " & strTable & " (" & _
                "WarehouseId char(5) PRIMARY KEY, " & _
                "Name char(50) NULL DEFAULT '', " & _
                "Address varchar(150) NULL DEFAULT '', " & _
                "EmployeeId char(7) NULL DEFAULT '', " & _
                "Notes varchar(150) NULL DEFAULT '', " & _
                "WarehouseSet char(1) NULL DEFAULT '" & mdlGlobal.strNo & "', " & _
                "CreateId char(8) NOT NULL, " & _
                "CreateDate datetime NULL DEFAULT GETDATE(), " & _
                "UpdateId char(8) NOT NULL, " & _
                "UpdateDate datetime NULL DEFAULT GETDATE())"
        ElseIf mdlGlobal.objDatabaseInit = MSACCESS Then
            strQuery = "CREATE TABLE " & strTable & " (" & _
                "WarehouseId char(5) PRIMARY KEY, " & _
                "Name char(50) NOT NULL, " & _
                "Address varchar(150) NOT NULL, " & _
                "EmployeeId char(7) NOT NULL, " & _
                "Notes varchar(150) NOT NULL, " & _
                "WarehouseSet char(1) NOT NULL, " & _
                "CreateId char(8) NOT NULL, " & _
                "CreateDate datetime NOT NULL, " & _
                "UpdateId char(8) NOT NULL, " & _
                "UpdateDate datetime NOT NULL)"
        ElseIf mdlGlobal.objDatabaseInit = MYSQL Then
            strQuery = "CREATE TABLE `" & strTable & "` (" & _
                "`WarehouseId` char(5) NOT NULL, " & _
                "`Name` char(50) NOT NULL, " & _
                "`Address` varchar(150) NOT NULL, " & _
                "`EmployeeId` char(7) NOT NULL, " & _
                "`Notes` varchar(150) NOT NULL, " & _
                "`WarehouseSet` char(1) NOT NULL, " & _
                "`CreateId` char(8) NOT NULL, " & _
                "`CreateDate` datetime NOT NULL, " & _
                "`UpdateId` char(8) NOT NULL, " & _
                "`UpdateDate` datetime NOT NULL, " & _
                "PRIMARY KEY(`WarehouseId`)) TYPE=MyISAM;"
        End If
            
        mdlDatabase.CreateTable mdlGlobal.conInventory, strQuery, strTable, mdlGlobal.objDatabaseInit
    End If
End Function

Public Function CreateTMCUSTOMER(Optional ByVal blnCreate As Boolean = False) As String
    Dim strTable As String
    
    strTable = "TMCUSTOMER"
    
    CreateTMCUSTOMER = strTable
    
    If mdlGlobal.UserAuthority.IsAdmin Or blnCreate Then
        Dim strQuery As String
        
        If mdlGlobal.objDatabaseInit = SQLSERVER7 Or _
            mdlGlobal.objDatabaseInit = SQLSERVER2000 Or _
            mdlGlobal.objDatabaseInit = SQLEXPRESS Then
            strQuery = "CREATE TABLE " & strTable & " (" & _
                "CustomerId char(6) PRIMARY KEY, " & _
                "CustomerDate datetime NULL DEFAULT GETDATE(), " & _
                "Name char(50) NULL DEFAULT '', " & _
                "Address varchar(150) NULL DEFAULT '', " & _
                "Phone char(50) NULL DEFAULT '', " & _
                "Fax char(50) NULL DEFAULT '', " & _
                "NPWP char(20) NULL DEFAULT '', " & _
                "StatusYN char(1) NOT NULL, " & _
                "Notes varchar(150) NULL DEFAULT '', " & _
                "CreateId char(8) NOT NULL, " & _
                "CreateDate datetime NULL DEFAULT GETDATE(), " & _
                "UpdateId char(8) NOT NULL, " & _
                "UpdateDate datetime NULL DEFAULT GETDATE())"
        ElseIf mdlGlobal.objDatabaseInit = MSACCESS Then
            strQuery = "CREATE TABLE " & strTable & " (" & _
                "CustomerId char(6) PRIMARY KEY, " & _
                "CustomerDate datetime NOT NULL, " & _
                "Name char(50) NOT NULL, " & _
                "Address varchar(150) NOT NULL, " & _
                "Phone char(50) NOT NULL, " & _
                "Fax char(50) NOT NULL, " & _
                "NPWP char(20) NOT NULL, " & _
                "StatusYN char(1) NOT NULL, " & _
                "Notes varchar(150) NOT NULL, " & _
                "CreateId char(8) NOT NULL, " & _
                "CreateDate datetime NOT NULL, " & _
                "UpdateId char(8) NOT NULL, " & _
                "UpdateDate datetime NOT NULL)"
        ElseIf mdlGlobal.objDatabaseInit = MYSQL Then
            strQuery = "CREATE TABLE `" & strTable & "` (" & _
                "`CustomerId` char(6) NOT NULL, " & _
                "`CustomerDate` datetime NOT NULL, " & _
                "`Name` char(50) NOT NULL, " & _
                "`Address` varchar(150) NOT NULL, " & _
                "`Phone` char(50) NOT NULL, " & _
                "`Fax` char(50) NOT NULL, " & _
                "`NPWP` char(20) NOT NULL, " & _
                "`StatusYN` char(1) NOT NULL, " & _
                "`Notes` varchar(150) NOT NULL, " & _
                "`CreateId` char(8) NOT NULL, " & _
                "`CreateDate` datetime NOT NULL, " & _
                "`UpdateId` char(8) NOT NULL, " & _
                "`UpdateDate` datetime NOT NULL, " & _
                "PRIMARY KEY(`CustomerId`)) TYPE=MyISAM;"
        End If
            
        mdlDatabase.CreateTable mdlGlobal.conInventory, strQuery, strTable, mdlGlobal.objDatabaseInit
    End If
End Function

Public Function CreateTMCONTACTVENDOR(Optional ByVal blnCreate As Boolean = False) As String
    Dim strTable As String
    
    strTable = "TMCONTACTVENDOR"
    
    CreateTMCONTACTVENDOR = strTable
    
    If mdlGlobal.UserAuthority.IsAdmin Or blnCreate Then
        Dim strQuery As String
        
        If mdlGlobal.objDatabaseInit = SQLSERVER7 Or _
            mdlGlobal.objDatabaseInit = SQLSERVER2000 Or _
            mdlGlobal.objDatabaseInit = SQLEXPRESS Then
            strQuery = "CREATE TABLE " & strTable & " (" & _
                "ContactId char(56) PRIMARY KEY, " & _
                "VendorId char(6) NOT NULL, " & _
                "Name char(50) NULL DEFAULT '', " & _
                "Phone char(50) NULL DEFAULT '', " & _
                "HandPhone char(50) NULL DEFAULT '', " & _
                "Email char(50) NULL DEFAULT '', " & _
                "CreateId char(8) NOT NULL, " & _
                "CreateDate datetime NULL DEFAULT GETDATE(), " & _
                "UpdateId char(8) NOT NULL, " & _
                "UpdateDate datetime NULL DEFAULT GETDATE())"
        ElseIf mdlGlobal.objDatabaseInit = MSACCESS Then
            strQuery = "CREATE TABLE " & strTable & " (" & _
                "ContactId char(56) PRIMARY KEY, " & _
                "VendorId char(6) NOT NULL, " & _
                "Name char(50) NOT NULL, " & _
                "Phone char(50) NOT NULL, " & _
                "HandPhone char(50) NOT NULL, " & _
                "Email char(50) NOT NULL, " & _
                "CreateId char(8) NOT NULL, " & _
                "CreateDate datetime NOT NULL, " & _
                "UpdateId char(8) NOT NULL, " & _
                "UpdateDate datetime NOT NULL)"
        ElseIf mdlGlobal.objDatabaseInit = MYSQL Then
            strQuery = "CREATE TABLE `" & strTable & "` (" & _
                "`ContactId` char(56) NOT NULL, " & _
                "`VendorId` char(6) NOT NULL, " & _
                "`Name` char(50) NOT NULL, " & _
                "`Phone` char(50) NOT NULL, " & _
                "`HandPhone` char(50) NOT NULL, " & _
                "`Email` char(50) NOT NULL, " & _
                "`CreateId` char(8) NOT NULL, " & _
                "`CreateDate` datetime NOT NULL, " & _
                "`UpdateId` char(8) NOT NULL, " & _
                "`UpdateDate` datetime NOT NULL, " & _
                "PRIMARY KEY(`ContactId`)) TYPE=MyISAM;"
        End If
            
        mdlDatabase.CreateTable mdlGlobal.conInventory, strQuery, strTable, mdlGlobal.objDatabaseInit
    End If
End Function

Public Function CreateTMDELIVERYCUSTOMER(Optional ByVal blnCreate As Boolean = False) As String
    Dim strTable As String
    
    strTable = "TMDELIVERYCUSTOMER"
    
    CreateTMDELIVERYCUSTOMER = strTable
    
    If mdlGlobal.UserAuthority.IsAdmin Or blnCreate Then
        Dim strQuery As String
        
        If mdlGlobal.objDatabaseInit = SQLSERVER7 Or _
            mdlGlobal.objDatabaseInit = SQLSERVER2000 Or _
            mdlGlobal.objDatabaseInit = SQLEXPRESS Then
            strQuery = "CREATE TABLE " & strTable & " (" & _
                "DeliveryId char(9) PRIMARY KEY, " & _
                "CustomerId char(6) NOT NULL, " & _
                "NoSeq char(3) NOT NULL, " & _
                "Name char(50) NULL DEFAULT '', " & _
                "Address varchar(150) NULL DEFAULT '', " & _
                "Phone char(50) NULL DEFAULT '', " & _
                "Fax char(50) NULL DEFAULT '', " & _
                "Notes varchar(150) NULL DEFAULT '', " & _
                "CreateId char(8) NOT NULL, " & _
                "CreateDate datetime NULL DEFAULT GETDATE(), " & _
                "UpdateId char(8) NOT NULL, " & _
                "UpdateDate datetime NULL DEFAULT GETDATE())"
        ElseIf mdlGlobal.objDatabaseInit = MSACCESS Then
            strQuery = "CREATE TABLE " & strTable & " (" & _
                "DeliveryId char(9) PRIMARY KEY, " & _
                "CustomerId char(6) NOT NULL, " & _
                "NoSeq char(3) NOT NULL, " & _
                "Name char(50) NOT NULL, " & _
                "Address varchar(150) NOT NULL, " & _
                "Phone char(50) NOT NULL, " & _
                "Fax char(50) NOT NULL, " & _
                "Notes varchar(150) NOT NULL, " & _
                "CreateId char(8) NOT NULL, " & _
                "CreateDate datetime NOT NULL, " & _
                "UpdateId char(8) NOT NULL, " & _
                "UpdateDate datetime NOT NULL)"
        ElseIf mdlGlobal.objDatabaseInit = MYSQL Then
            strQuery = "CREATE TABLE `" & strTable & "` (" & _
                "`DeliveryId` char(9) NOT NULL, " & _
                "`CustomerId` char(6) NOT NULL, " & _
                "`NoSeq` char(3) NOT NULL, " & _
                "`Name` char(50) NOT NULL, " & _
                "`Address` varchar(150) NOT NULL, " & _
                "`Phone` char(50) NOT NULL, " & _
                "`Fax` char(50) NOT NULL, " & _
                "`Notes` varchar(150) NOT NULL, " & _
                "`CreateId` char(8) NOT NULL, " & _
                "`CreateDate` datetime NOT NULL, " & _
                "`UpdateId` char(8) NOT NULL, " & _
                "`UpdateDate` datetime NOT NULL, " & _
                "PRIMARY KEY(`DeliveryId`)) TYPE=MyISAM;"
        End If
            
        mdlDatabase.CreateTable mdlGlobal.conInventory, strQuery, strTable, mdlGlobal.objDatabaseInit
    End If
End Function

Public Function CreateTMCONTACTCUSTOMER(Optional ByVal blnCreate As Boolean = False) As String
    Dim strTable As String
    
    strTable = "TMCONTACTCUSTOMER"
    
    CreateTMCONTACTCUSTOMER = strTable
    
    If mdlGlobal.UserAuthority.IsAdmin Or blnCreate Then
        Dim strQuery As String
        
        If mdlGlobal.objDatabaseInit = SQLSERVER7 Or _
            mdlGlobal.objDatabaseInit = SQLSERVER2000 Or _
            mdlGlobal.objDatabaseInit = SQLEXPRESS Then
            strQuery = "CREATE TABLE " & strTable & " (" & _
                "ContactId char(56) PRIMARY KEY, " & _
                "CustomerId char(6) NOT NULL, " & _
                "Name char(50) NULL DEFAULT '', " & _
                "Phone char(50) NULL DEFAULT '', " & _
                "HandPhone char(50) NULL DEFAULT '', " & _
                "Email char(50) NULL DEFAULT '', " & _
                "CreateId char(8) NOT NULL, " & _
                "CreateDate datetime NULL DEFAULT GETDATE(), " & _
                "UpdateId char(8) NOT NULL, " & _
                "UpdateDate datetime NULL DEFAULT GETDATE())"
        ElseIf mdlGlobal.objDatabaseInit = MSACCESS Then
            strQuery = "CREATE TABLE " & strTable & " (" & _
                "ContactId char(56) PRIMARY KEY, " & _
                "CustomerId char(6) NOT NULL, " & _
                "Name char(50) NOT NULL, " & _
                "Phone char(50) NOT NULL, " & _
                "HandPhone char(50) NOT NULL, " & _
                "Email char(50) NOT NULL, " & _
                "CreateId char(8) NOT NULL, " & _
                "CreateDate datetime NOT NULL, " & _
                "UpdateId char(8) NOT NULL, " & _
                "UpdateDate datetime NOT NULL)"
        ElseIf mdlGlobal.objDatabaseInit = MYSQL Then
            strQuery = "CREATE TABLE `" & strTable & "` (" & _
                "`ContactId` char(56) NOT NULL, " & _
                "`CustomerId` char(6) NOT NULL, " & _
                "`Name` char(50) NOT NULL, " & _
                "`Phone` char(50) NOT NULL, " & _
                "`HandPhone` char(50) NOT NULL, " & _
                "`Email` char(50) NOT NULL, " & _
                "`CreateId` char(8) NOT NULL, " & _
                "`CreateDate` datetime NOT NULL, " & _
                "`UpdateId` char(8) NOT NULL, " & _
                "`UpdateDate` datetime NOT NULL, " & _
                "PRIMARY KEY(`ContactId`)) TYPE=MyISAM;"
        End If
            
        mdlDatabase.CreateTable mdlGlobal.conInventory, strQuery, strTable, mdlGlobal.objDatabaseInit
    End If
End Function

Public Function CreateTMCUSTOMERNOTES(Optional ByVal blnCreate As Boolean = False) As String
    Dim strTable As String
    
    strTable = "TMCUSTOMERNOTES"
    
    CreateTMCUSTOMERNOTES = strTable
    
    If mdlGlobal.UserAuthority.IsAdmin Or blnCreate Then
        Dim strQuery As String
        
        If mdlGlobal.objDatabaseInit = SQLSERVER7 Or _
            mdlGlobal.objDatabaseInit = SQLSERVER2000 Or _
            mdlGlobal.objDatabaseInit = SQLEXPRESS Then
            strQuery = "CREATE TABLE " & strTable & " (" & _
                "NotesId char(20) PRIMARY KEY, " & _
                "CustomerId char(6) NOT NULL, " & _
                "NotesDate datetime NULL DEFAULT GETDATE(), " & _
                "Notes varchar(200) NULL DEFAULT '', " & _
                "CreateId char(8) NOT NULL, " & _
                "CreateDate datetime NULL DEFAULT GETDATE(), " & _
                "UpdateId char(8) NOT NULL, " & _
                "UpdateDate datetime NULL DEFAULT GETDATE())"
        ElseIf mdlGlobal.objDatabaseInit = MSACCESS Then
            strQuery = "CREATE TABLE " & strTable & " (" & _
                "NotesId char(20) PRIMARY KEY, " & _
                "CustomerId char(6) NOT NULL, " & _
                "NotesDate datetime NOT NULL, " & _
                "Notes varchar(200) NOT NULL, " & _
                "CreateId char(8) NOT NULL, " & _
                "CreateDate datetime NOT NULL, " & _
                "UpdateId char(8) NOT NULL, " & _
                "UpdateDate datetime NOT NULL)"
        ElseIf mdlGlobal.objDatabaseInit = MYSQL Then
            strQuery = "CREATE TABLE `" & strTable & "` (" & _
                "`NotesId` char(20) NOT NULL, " & _
                "`CustomerId` char(6) NOT NULL, " & _
                "`NotesDate` datetime NOT NULL, " & _
                "`Notes` varchar(200) NOT NULL, " & _
                "`CreateId` char(8) NOT NULL, " & _
                "`CreateDate` datetime NOT NULL, " & _
                "`UpdateId` char(8) NOT NULL, " & _
                "`UpdateDate` datetime NOT NULL, " & _
                "PRIMARY KEY(`NotesId`)) TYPE=MyISAM;"
        End If
            
        mdlDatabase.CreateTable mdlGlobal.conInventory, strQuery, strTable, mdlGlobal.objDatabaseInit
    End If
End Function

Public Function CreateTMCONTACTNOTES(Optional ByVal blnCreate As Boolean = False) As String
    Dim strTable As String
    
    strTable = "TMCONTACTNOTES"
    
    CreateTMCONTACTNOTES = strTable
    
    If mdlGlobal.UserAuthority.IsAdmin Or blnCreate Then
        Dim strQuery As String
        
        If mdlGlobal.objDatabaseInit = SQLSERVER7 Or _
            mdlGlobal.objDatabaseInit = SQLSERVER2000 Or _
            mdlGlobal.objDatabaseInit = SQLEXPRESS Then
            strQuery = "CREATE TABLE " & strTable & " (" & _
                "NotesId char(70) PRIMARY KEY, " & _
                "ContactId char(56) NOT NULL, " & _
                "NotesDate datetime NULL DEFAULT GETDATE(), " & _
                "Notes varchar(200) NULL DEFAULT '', " & _
                "CreateId char(8) NOT NULL, " & _
                "CreateDate datetime NULL DEFAULT GETDATE(), " & _
                "UpdateId char(8) NOT NULL, " & _
                "UpdateDate datetime NULL DEFAULT GETDATE())"
        ElseIf mdlGlobal.objDatabaseInit = MSACCESS Then
            strQuery = "CREATE TABLE " & strTable & " (" & _
                "NotesId char(70) PRIMARY KEY, " & _
                "ContactId char(56) NOT NULL, " & _
                "NotesDate datetime NOT NULL, " & _
                "Notes varchar(200) NOT NULL, " & _
                "CreateId char(8) NOT NULL, " & _
                "CreateDate datetime NOT NULL, " & _
                "UpdateId char(8) NOT NULL, " & _
                "UpdateDate datetime NOT NULL)"
        ElseIf mdlGlobal.objDatabaseInit = MYSQL Then
            strQuery = "CREATE TABLE `" & strTable & "` (" & _
                "`NotesId` char(70) NOT NULL, " & _
                "`ContactId` char(56) NOT NULL, " & _
                "`NotesDate` datetime NOT NULL, " & _
                "`Notes` varchar(200) NOT NULL, " & _
                "`CreateId` char(8) NOT NULL, " & _
                "`CreateDate` datetime NOT NULL, " & _
                "`UpdateId` char(8) NOT NULL, " & _
                "`UpdateDate` datetime NOT NULL, " & _
                "PRIMARY KEY(`NotesId`)) TYPE=MyISAM;"
        End If
            
        mdlDatabase.CreateTable mdlGlobal.conInventory, strQuery, strTable, mdlGlobal.objDatabaseInit
    End If
End Function

Public Function CreateTMREMINDERCUSTOMER(Optional ByVal blnCreate As Boolean = False) As String
    Dim strTable As String
    
    strTable = "TMREMINDERCUSTOMER"
    
    CreateTMREMINDERCUSTOMER = strTable
    
    If mdlGlobal.UserAuthority.IsAdmin Or blnCreate Then
        Dim strQuery As String
        
        If mdlGlobal.objDatabaseInit = SQLSERVER7 Or _
            mdlGlobal.objDatabaseInit = SQLSERVER2000 Or _
            mdlGlobal.objDatabaseInit = SQLEXPRESS Then
            strQuery = "CREATE TABLE " & strTable & " (" & _
                "CustomerId char(6) PRIMARY KEY, " & _
                "ReminderType char(1) NULL DEFAULT '', " & _
                "ReminderDate datetime NULL DEFAULT GETDATE(), " & _
                "ValidateType char(1) NULL DEFAULT '', " & _
                "ValidateDate datetime NULL DEFAULT GETDATE(), " & _
                "CreateId char(8) NOT NULL, " & _
                "CreateDate datetime NULL DEFAULT GETDATE(), " & _
                "UpdateId char(8) NOT NULL, " & _
                "UpdateDate datetime NULL DEFAULT GETDATE())"
        ElseIf mdlGlobal.objDatabaseInit = MSACCESS Then
            strQuery = "CREATE TABLE " & strTable & " (" & _
                "CustomerId char(6) PRIMARY KEY, " & _
                "ReminderType char(1) NOT NULL, " & _
                "ReminderDate datetime NOT NULL, " & _
                "ValidateType char(1) NOT NULL, " & _
                "ValidateDate datetime NOT NULL, " & _
                "CreateId char(8) NOT NULL, " & _
                "CreateDate datetime NOT NULL, " & _
                "UpdateId char(8) NOT NULL, " & _
                "UpdateDate datetime NOT NULL)"
        ElseIf mdlGlobal.objDatabaseInit = MYSQL Then
            strQuery = "CREATE TABLE `" & strTable & "` (" & _
                "`CustomerId` char(6) NOT NULL, " & _
                "`ReminderType` char(1) NOT NULL, " & _
                "`ReminderDate` datetime NOT NULL, " & _
                "`ValidateType` char(1) NOT NULL, " & _
                "`ValidateDate` datetime NOT NULL, " & _
                "`CreateId` char(8) NOT NULL, " & _
                "`CreateDate` datetime NOT NULL, " & _
                "`UpdateId` char(8) NOT NULL, " & _
                "`UpdateDate` datetime NOT NULL, " & _
                "PRIMARY KEY(`CustomerId`)) TYPE=MyISAM;"
        End If
            
        mdlDatabase.CreateTable mdlGlobal.conInventory, strQuery, strTable, mdlGlobal.objDatabaseInit
    End If
End Function

Public Function CreateTMITEM(Optional ByVal blnCreate As Boolean = False) As String
    Dim strTable As String
    
    strTable = "TMITEM"
    
    CreateTMITEM = strTable
    
    If mdlGlobal.UserAuthority.IsAdmin Or blnCreate Then
        Dim strQuery As String
        
        If mdlGlobal.objDatabaseInit = SQLSERVER7 Or _
            mdlGlobal.objDatabaseInit = SQLSERVER2000 Or _
            mdlGlobal.objDatabaseInit = SQLEXPRESS Then
            strQuery = "CREATE TABLE " & strTable & " (" & _
                "ItemId char(7) PRIMARY KEY, " & _
                "ItemDate datetime NULL DEFAULT GETDATE(), " & _
                "PartNumber char(40) NULL DEFAULT '', " & _
                "Name char(50) NULL DEFAULT '', " & _
                "VendorId char(6) NULL DEFAULT '', " & _
                "GroupId char(4) NULL DEFAULT '', " & _
                "CategoryId char(4) NULL DEFAULT '', " & _
                "BrandId char(4) NULL DEFAULT '', " & _
                "UnityId char(4) NULL DEFAULT '', " & _
                "MinStock money NULL DEFAULT 0, " & _
                "MaxStock money NULL DEFAULT 0, " & _
                "Notes varchar(150) NULL DEFAULT '', " & _
                "CreateId char(8) NOT NULL, " & _
                "CreateDate datetime NULL DEFAULT GETDATE(), " & _
                "UpdateId char(8) NOT NULL, " & _
                "UpdateDate datetime NULL DEFAULT GETDATE())"
        ElseIf mdlGlobal.objDatabaseInit = MSACCESS Then
            strQuery = "CREATE TABLE " & strTable & " (" & _
                "ItemId char(7) PRIMARY KEY, " & _
                "ItemDate datetime NOT NULL, " & _
                "PartNumber char(40) NOT NULL, " & _
                "Name char(50) NOT NULL, " & _
                "VendorId char(6) NOT NULL, " & _
                "GroupId char(4) NOT NULL, " & _
                "CategoryId char(4) NOT NULL, " & _
                "BrandId char(4) NOT NULL, " & _
                "UnityId char(4) NOT NULL, " & _
                "MinStock money NOT NULL, " & _
                "MaxStock money NOT NULL, " & _
                "Notes varchar(150) NOT NULL, " & _
                "CreateId char(8) NOT NULL, " & _
                "CreateDate datetime NOT NULL, " & _
                "UpdateId char(8) NOT NULL, " & _
                "UpdateDate datetime NOT NULL)"
        ElseIf mdlGlobal.objDatabaseInit = MYSQL Then
            strQuery = "CREATE TABLE `" & strTable & "` (" & _
                "`ItemId` char(7) NOT NULL, " & _
                "`ItemDate` datetime NOT NULL, " & _
                "`PartNumber` char(40) NOT NULL, " & _
                "`Name` char(50) NOT NULL, " & _
                "`VendorId` char(6) NOT NULL, " & _
                "`GroupId` char(4) NOT NULL, " & _
                "`CategoryId` char(4) NOT NULL, " & _
                "`BrandId` char(4) NOT NULL, " & _
                "`UnityId` char(4) NOT NULL, " & _
                "`MinStock` double NOT NULL, " & _
                "`MaxStock` double NOT NULL, " & _
                "`Notes` varchar(50) NOT NULL, " & _
                "`CreateId` char(8) NOT NULL, " & _
                "`CreateDate` datetime NOT NULL, " & _
                "`UpdateId` char(8) NOT NULL, " & _
                "`UpdateDate` datetime NOT NULL, " & _
                "PRIMARY KEY(`ItemId`)) TYPE=MyISAM;"
        End If
            
        mdlDatabase.CreateTable mdlGlobal.conInventory, strQuery, strTable, mdlGlobal.objDatabaseInit
    End If
End Function

Public Function CreateTMPRICELIST(Optional ByVal blnCreate As Boolean = False) As String
    Dim strTable As String

    strTable = "TMPRICELIST"

    CreateTMPRICELIST = strTable
    
    If mdlGlobal.UserAuthority.IsAdmin Or blnCreate Then
        Dim strQuery As String
    
        If mdlGlobal.objDatabaseInit = SQLSERVER7 Or _
            mdlGlobal.objDatabaseInit = SQLSERVER2000 Or _
            mdlGlobal.objDatabaseInit = SQLEXPRESS Then
            strQuery = "CREATE TABLE " & strTable & " (" & _
                "PriceListId char(15) PRIMARY KEY, " & _
                "PriceListDate datetime NULL DEFAULT GETDATE(), " & _
                "ItemId char(7) NOT NULL, " & _
                "PriceListValue money NULL DEFAULT 0, " & _
                "CurrencyId char(5) NOT NULL, " & _
                "CreateId char(8) NOT NULL, " & _
                "CreateDate datetime NULL DEFAULT GETDATE(), " & _
                "UpdateId char(8) NOT NULL, " & _
                "UpdateDate datetime NULL DEFAULT GETDATE())"
        ElseIf mdlGlobal.objDatabaseInit = MSACCESS Then
            strQuery = "CREATE TABLE " & strTable & " (" & _
                "PriceListId char(15) PRIMARY KEY, " & _
                "PriceListDate datetime NOT NULL, " & _
                "ItemId char(7) NOT NULL, " & _
                "PriceListValue money NOT NULL, " & _
                "CurrencyId char(5) NOT NULL, " & _
                "CreateId char(8) NOT NULL, " & _
                "CreateDate datetime NOT NULL, " & _
                "UpdateId char(8) NOT NULL, " & _
                "UpdateDate datetime NOT NULL)"
        ElseIf mdlGlobal.objDatabaseInit = MYSQL Then
            strQuery = "CREATE TABLE `" & strTable & "` (" & _
                "`PriceListId` char(15) NOT NULL, " & _
                "`PriceListDate` datetime NOT NULL, " & _
                "`ItemId` char(7) NOT NULL, " & _
                "`PriceListValue` double NOT NULL, " & _
                "`CurrencyId` char(5) NOT NULL, " & _
                "`CreateId` char(8) NOT NULL, " & _
                "`CreateDate` datetime NOT NULL, " & _
                "`UpdateId` char(8) NOT NULL, " & _
                "`UpdateDate` datetime NOT NULL, " & _
                "PRIMARY KEY(`PriceListId`)) TYPE=MyISAM;"
        End If
    
        mdlDatabase.CreateTable mdlGlobal.conInventory, strQuery, strTable, mdlGlobal.objDatabaseInit
    End If
End Function

Public Function CreateTMITEMPRICE(Optional ByVal blnCreate As Boolean = False) As String
    Dim strTable As String
    
    strTable = "TMITEMPRICE"
    
    CreateTMITEMPRICE = strTable
    
    If mdlGlobal.UserAuthority.IsAdmin Or blnCreate Then
        Dim strQuery As String
        
        If mdlGlobal.objDatabaseInit = SQLSERVER7 Or _
            mdlGlobal.objDatabaseInit = SQLSERVER2000 Or _
            mdlGlobal.objDatabaseInit = SQLEXPRESS Then
            strQuery = "CREATE TABLE " & strTable & " (" & _
                "PriceId char(15) PRIMARY KEY, " & _
                "ItemId char(7) NOT NULL, " & _
                "PriceDate datetime NULL DEFAULT GETDATE(), " & _
                "ItemPrice money NULL DEFAULT 0, " & _
                "CurrencyId char(5) NULL DEFAULT '', " & _
                "CreateId char(8) NOT NULL, " & _
                "CreateDate datetime NULL DEFAULT GETDATE(), " & _
                "UpdateId char(8) NOT NULL, " & _
                "UpdateDate datetime NULL DEFAULT GETDATE())"
        ElseIf mdlGlobal.objDatabaseInit = MSACCESS Then
            strQuery = "CREATE TABLE " & strTable & " (" & _
                "PriceId char(15) PRIMARY KEY, " & _
                "ItemId char(7) NOT NULL, " & _
                "PriceDate datetime NOT NULL, " & _
                "ItemPrice money NOT NULL, " & _
                "CurrencyId char(5) NOT NULL, " & _
                "CreateId char(8) NOT NULL, " & _
                "CreateDate datetime NOT NULL, " & _
                "UpdateId char(8) NOT NULL, " & _
                "UpdateDate datetime NOT NULL)"
        ElseIf mdlGlobal.objDatabaseInit = MYSQL Then
            strQuery = "CREATE TABLE `" & strTable & "` (" & _
                "`PriceId` char(15) NOT NULL, " & _
                "`ItemId` char(7) NOT NULL, " & _
                "`PriceDate` datetime NOT NULL, " & _
                "`ItemPrice` double NOT NULL, " & _
                "`CurrencyId` char(5) NOT NULL, " & _
                "`CreateId` char(8) NOT NULL, " & _
                "`CreateDate` datetime NOT NULL, " & _
                "`UpdateId` char(8) NOT NULL, " & _
                "`UpdateDate` datetime NOT NULL, " & _
                "PRIMARY KEY(`PriceId`)) TYPE=MyISAM;"
        End If
            
        mdlDatabase.CreateTable mdlGlobal.conInventory, strQuery, strTable, mdlGlobal.objDatabaseInit
    End If
End Function

Public Function CreateTMPRICEBUY(Optional ByVal blnCreate As Boolean = False) As String
    Dim strTable As String
    
    strTable = "TMPRICEBUY"
    
    CreateTMPRICEBUY = strTable
    
    If mdlGlobal.UserAuthority.IsAdmin Or blnCreate Then
        Dim strQuery As String
        
        If mdlGlobal.objDatabaseInit = SQLSERVER7 Or _
            mdlGlobal.objDatabaseInit = SQLSERVER2000 Or _
            mdlGlobal.objDatabaseInit = SQLEXPRESS Then
            strQuery = "CREATE TABLE " & strTable & " (" & _
                "PriceId char(15) PRIMARY KEY, " & _
                "ItemId char(7) NOT NULL, " & _
                "PriceDate datetime NULL DEFAULT GETDATE(), " & _
                "ItemPrice money NULL DEFAULT 0, " & _
                "CurrencyId char(5) NULL DEFAULT '', " & _
                "CreateId char(8) NOT NULL, " & _
                "CreateDate datetime NULL DEFAULT GETDATE(), " & _
                "UpdateId char(8) NOT NULL, " & _
                "UpdateDate datetime NULL DEFAULT GETDATE())"
        ElseIf mdlGlobal.objDatabaseInit = MSACCESS Then
            strQuery = "CREATE TABLE " & strTable & " (" & _
                "PriceId char(15) PRIMARY KEY, " & _
                "ItemId char(7) NOT NULL, " & _
                "PriceDate datetime NOT NULL, " & _
                "ItemPrice money NOT NULL, " & _
                "CurrencyId char(5) NOT NULL, " & _
                "CreateId char(8) NOT NULL, " & _
                "CreateDate datetime NOT NULL, " & _
                "UpdateId char(8) NOT NULL, " & _
                "UpdateDate datetime NOT NULL)"
        ElseIf mdlGlobal.objDatabaseInit = MYSQL Then
            strQuery = "CREATE TABLE `" & strTable & "` (" & _
                "`PriceId` char(15) NOT NULL, " & _
                "`ItemId` char(7) NOT NULL, " & _
                "`PriceDate` datetime NOT NULL, " & _
                "`ItemPrice` double NOT NULL, " & _
                "`CurrencyId` char(5) NOT NULL, " & _
                "`CreateId` char(8) NOT NULL, " & _
                "`CreateDate` datetime NOT NULL, " & _
                "`UpdateId` char(8) NOT NULL, " & _
                "`UpdateDate` datetime NOT NULL, " & _
                "PRIMARY KEY(`PriceId`)) TYPE=MyISAM;"
        End If
            
        mdlDatabase.CreateTable mdlGlobal.conInventory, strQuery, strTable, mdlGlobal.objDatabaseInit
    End If
End Function

Public Function CreateTMCONVERTPRICE(Optional ByVal blnCreate As Boolean = False) As String
    Dim strTable As String
    
    strTable = "TMCONVERTPRICE"
    
    CreateTMCONVERTPRICE = strTable
    
    If mdlGlobal.UserAuthority.IsAdmin Or blnCreate Then
        Dim strQuery As String
        
        If mdlGlobal.objDatabaseInit = SQLSERVER7 Or _
            mdlGlobal.objDatabaseInit = SQLSERVER2000 Or _
            mdlGlobal.objDatabaseInit = SQLEXPRESS Then
            strQuery = "CREATE TABLE " & strTable & " (" & _
                "ConvertId char(30) PRIMARY KEY, " & _
                "ConvertDate datetime NULL DEFAULT GETDATE(), " & _
                "ItemId char(7) NOT NULL, " & _
                "PriceListId char(15) NOT NULL, " & _
                "Weight money NULL DEFAULT 0, " & _
                "Typemf char(2) NOT NULL, " & _
                "TypemfValue money NULL DEFAULT 0, " & _
                "Freight money NULL DEFAULT 0, " & _
                "CreateId char(8) NOT NULL, " & _
                "CreateDate datetime NULL DEFAULT GETDATE(), " & _
                "UpdateId char(8) NOT NULL, " & _
                "UpdateDate datetime NULL DEFAULT GETDATE())"
        ElseIf mdlGlobal.objDatabaseInit = MSACCESS Then
            strQuery = "CREATE TABLE " & strTable & " (" & _
                "ConvertId char(30) PRIMARY KEY, " & _
                "ConvertDate datetime NOT NULL, " & _
                "ItemId char(7) NOT NULL, " & _
                "PriceListId char(15) NOT NULL, " & _
                "Weight money NOT NULL, " & _
                "Typemf char(2) NOT NULL, " & _
                "TypemfValue money NOT NULL, " & _
                "Freight money NOT NULL, " & _
                "CreateId char(8) NOT NULL, " & _
                "CreateDate datetime NOT NULL, " & _
                "UpdateId char(8) NOT NULL, " & _
                "UpdateDate datetime NOT NULL)"
        ElseIf mdlGlobal.objDatabaseInit = MYSQL Then
            strQuery = "CREATE TABLE `" & strTable & "` (" & _
                "`ConvertId` char(30) NOT NULL, " & _
                "`ConvertDate` datetime NOT NULL, " & _
                "`ItemId` char(7) NOT NULL, " & _
                "`PriceListId` char(15) NOT NULL, " & _
                "`Weight` double NOT NULL, " & _
                "`Typemf` char(2) NOT NULL, " & _
                "`TypemfValue` double NOT NULL, " & _
                "`Freight` double NOT NULL, " & _
                "`CreateId` char(8) NOT NULL, " & _
                "`CreateDate` datetime NOT NULL, " & _
                "`UpdateId` char(8) NOT NULL, " & _
                "`UpdateDate` datetime NOT NULL, " & _
                "PRIMARY KEY(`ConvertId`)) TYPE=MyISAM;"
        End If
            
        mdlDatabase.CreateTable mdlGlobal.conInventory, strQuery, strTable, mdlGlobal.objDatabaseInit
    End If
End Function

Public Function CreateTMUNITY(Optional ByVal blnCreate As Boolean = False) As String
    Dim strTable As String
    
    strTable = "TMUNITY"
    
    CreateTMUNITY = strTable
    
    If mdlGlobal.UserAuthority.IsAdmin Or blnCreate Then
        Dim strQuery As String
        
        If mdlGlobal.objDatabaseInit = SQLSERVER7 Or _
            mdlGlobal.objDatabaseInit = SQLSERVER2000 Or _
            mdlGlobal.objDatabaseInit = SQLEXPRESS Then
            strQuery = "CREATE TABLE " & strTable & " (" & _
                "UnityId char(4) PRIMARY KEY, " & _
                "Name char(50) NULL DEFAULT '', " & _
                "Notes varchar(150) NULL DEFAULT '', " & _
                "CreateId char(8) NOT NULL, " & _
                "CreateDate datetime NULL DEFAULT GETDATE(), " & _
                "UpdateId char(8) NOT NULL, " & _
                "UpdateDate datetime NULL DEFAULT GETDATE())"
        ElseIf mdlGlobal.objDatabaseInit = MSACCESS Then
            strQuery = "CREATE TABLE " & strTable & " (" & _
                "UnityId char(4) PRIMARY KEY, " & _
                "Name char(50) NOT NULL, " & _
                "Notes varchar(150) NOT NULL, " & _
                "CreateId char(8) NOT NULL, " & _
                "CreateDate datetime NOT NULL, " & _
                "UpdateId char(8) NOT NULL, " & _
                "UpdateDate datetime NOT NULL)"
        ElseIf mdlGlobal.objDatabaseInit = MYSQL Then
            strQuery = "CREATE TABLE `" & strTable & "` (" & _
                "`UnityId` char(4) NOT NULL, " & _
                "`Name` char(50) NOT NULL, " & _
                "`Notes` varchar(150) NOT NULL, " & _
                "`CreateId` char(8) NOT NULL, " & _
                "`CreateDate` datetime NOT NULL, " & _
                "`UpdateId` char(8) NOT NULL, " & _
                "`UpdateDate` datetime NOT NULL, " & _
                "PRIMARY KEY(`UnityId`)) TYPE=MyISAM;"
        End If
            
        mdlDatabase.CreateTable mdlGlobal.conInventory, strQuery, strTable, mdlGlobal.objDatabaseInit
    End If
End Function

Public Function CreateTMBRAND(Optional ByVal blnCreate As Boolean = False) As String
    Dim strTable As String
    
    strTable = "TMBRAND"
    
    CreateTMBRAND = strTable
    
    If mdlGlobal.UserAuthority.IsAdmin Or blnCreate Then
        Dim strQuery As String
        
        If mdlGlobal.objDatabaseInit = SQLSERVER7 Or _
            mdlGlobal.objDatabaseInit = SQLSERVER2000 Or _
            mdlGlobal.objDatabaseInit = SQLEXPRESS Then
            strQuery = "CREATE TABLE " & strTable & " (" & _
                "BrandId char(4) PRIMARY KEY, " & _
                "Name char(50) NULL DEFAULT '', " & _
                "Notes varchar(150) NULL DEFAULT '', " & _
                "CreateId char(8) NOT NULL, " & _
                "CreateDate datetime NULL DEFAULT GETDATE(), " & _
                "UpdateId char(8) NOT NULL, " & _
                "UpdateDate datetime NULL DEFAULT GETDATE())"
        ElseIf mdlGlobal.objDatabaseInit = MSACCESS Then
            strQuery = "CREATE TABLE " & strTable & " (" & _
                "BrandId char(4) PRIMARY KEY, " & _
                "Name char(50) NOT NULL, " & _
                "Notes varchar(150) NOT NULL, " & _
                "CreateId char(8) NOT NULL, " & _
                "CreateDate datetime NOT NULL, " & _
                "UpdateId char(8) NOT NULL, " & _
                "UpdateDate datetime NOT NULL)"
        ElseIf mdlGlobal.objDatabaseInit = MYSQL Then
            strQuery = "CREATE TABLE `" & strTable & "` (" & _
                "`BrandId` char(4) NOT NULL, " & _
                "`Name` char(50) NOT NULL, " & _
                "`Notes` varchar(150) NOT NULL, " & _
                "`CreateId` char(8) NOT NULL, " & _
                "`CreateDate` datetime NOT NULL, " & _
                "`UpdateId` char(8) NOT NULL, " & _
                "`UpdateDate` datetime NOT NULL, " & _
                "PRIMARY KEY(`BrandId`)) TYPE=MyISAM;"
        End If
            
        mdlDatabase.CreateTable mdlGlobal.conInventory, strQuery, strTable, mdlGlobal.objDatabaseInit
    End If
End Function

Public Function CreateTMGROUP(Optional ByVal blnCreate As Boolean = False) As String
    Dim strTable As String
    
    strTable = "TMGROUP"
    
    CreateTMGROUP = strTable
    
    If mdlGlobal.UserAuthority.IsAdmin Or blnCreate Then
        Dim strQuery As String
        
        If mdlGlobal.objDatabaseInit = SQLSERVER7 Or _
            mdlGlobal.objDatabaseInit = SQLSERVER2000 Or _
            mdlGlobal.objDatabaseInit = SQLEXPRESS Then
            strQuery = "CREATE TABLE " & strTable & " (" & _
                "GroupId char(4) PRIMARY KEY, " & _
                "Name char(50) NULL DEFAULT '', " & _
                "Notes varchar(150) NULL DEFAULT '', " & _
                "CreateId char(8) NOT NULL, " & _
                "CreateDate datetime NULL DEFAULT GETDATE(), " & _
                "UpdateId char(8) NOT NULL, " & _
                "UpdateDate datetime NULL DEFAULT GETDATE())"
        ElseIf mdlGlobal.objDatabaseInit = MSACCESS Then
            strQuery = "CREATE TABLE " & strTable & " (" & _
                "GroupId char(4) PRIMARY KEY, " & _
                "Name char(50) NOT NULL, " & _
                "Notes varchar(150) NOT NULL, " & _
                "CreateId char(8) NOT NULL, " & _
                "CreateDate datetime NOT NULL, " & _
                "UpdateId char(8) NOT NULL, " & _
                "UpdateDate datetime NOT NULL)"
        ElseIf mdlGlobal.objDatabaseInit = MYSQL Then
            strQuery = "CREATE TABLE `" & strTable & "` (" & _
                "`GroupId` char(4) NOT NULL, " & _
                "`Name` char(50) NOT NULL, " & _
                "`Notes` varchar(150) NOT NULL, " & _
                "`CreateId` char(8) NOT NULL, " & _
                "`CreateDate` datetime NOT NULL, " & _
                "`UpdateId` char(8) NOT NULL, " & _
                "`UpdateDate` datetime NOT NULL, " & _
                "PRIMARY KEY(`GroupId`)) TYPE=MyISAM;"
        End If
            
        mdlDatabase.CreateTable mdlGlobal.conInventory, strQuery, strTable, mdlGlobal.objDatabaseInit
    End If
End Function

Public Function CreateTMCATEGORY(Optional ByVal blnCreate As Boolean = False) As String
    Dim strTable As String
    
    strTable = "TMCATEGORY"
    
    CreateTMCATEGORY = strTable
    
    If mdlGlobal.UserAuthority.IsAdmin Or blnCreate Then
        Dim strQuery As String
        
        If mdlGlobal.objDatabaseInit = SQLSERVER7 Or _
            mdlGlobal.objDatabaseInit = SQLSERVER2000 Or _
            mdlGlobal.objDatabaseInit = SQLEXPRESS Then
            strQuery = "CREATE TABLE " & strTable & " (" & _
                "CategoryId char(4) PRIMARY KEY, " & _
                "Name char(50) NULL DEFAULT '', " & _
                "Notes varchar(150) NULL DEFAULT '', " & _
                "CreateId char(8) NOT NULL, " & _
                "CreateDate datetime NULL DEFAULT GETDATE(), " & _
                "UpdateId char(8) NOT NULL, " & _
                "UpdateDate datetime NULL DEFAULT GETDATE())"
        ElseIf mdlGlobal.objDatabaseInit = MSACCESS Then
            strQuery = "CREATE TABLE " & strTable & " (" & _
                "CategoryId char(4) PRIMARY KEY, " & _
                "Name char(50) NOT NULL, " & _
                "Notes varchar(150) NOT NULL, " & _
                "CreateId char(8) NOT NULL, " & _
                "CreateDate datetime NOT NULL, " & _
                "UpdateId char(8) NOT NULL, " & _
                "UpdateDate datetime NOT NULL)"
        ElseIf mdlGlobal.objDatabaseInit = MYSQL Then
            strQuery = "CREATE TABLE `" & strTable & "` (" & _
                "`CategoryId` char(4) NOT NULL, " & _
                "`Name` char(50) NOT NULL, " & _
                "`Notes` varchar(150) NOT NULL, " & _
                "`CreateId` char(8) NOT NULL, " & _
                "`CreateDate` datetime NOT NULL, " & _
                "`UpdateId` char(8) NOT NULL, " & _
                "`UpdateDate` datetime NOT NULL, " & _
                "PRIMARY KEY(`CategoryId`)) TYPE=MyISAM;"
        End If
            
        mdlDatabase.CreateTable mdlGlobal.conInventory, strQuery, strTable, mdlGlobal.objDatabaseInit
    End If
End Function

Public Function CreateTMCURRENCY(Optional ByVal blnCreate As Boolean = False) As String
    Dim strTable As String
    
    strTable = "TMCURRENCY"
    
    CreateTMCURRENCY = strTable
    
    If mdlGlobal.UserAuthority.IsAdmin Or blnCreate Then
        Dim strQuery As String
        
        If mdlGlobal.objDatabaseInit = SQLSERVER7 Or _
            mdlGlobal.objDatabaseInit = SQLSERVER2000 Or _
            mdlGlobal.objDatabaseInit = SQLEXPRESS Then
            strQuery = "CREATE TABLE " & strTable & " (" & _
                "CurrencyId char(5) PRIMARY KEY, " & _
                "Name char(50) NULL DEFAULT '', " & _
                "Notes varchar(150) NULL DEFAULT '', " & _
                "CreateId char(8) NOT NULL, " & _
                "CreateDate datetime NULL DEFAULT GETDATE(), " & _
                "UpdateId char(8) NOT NULL, " & _
                "UpdateDate datetime NULL DEFAULT GETDATE())"
        ElseIf mdlGlobal.objDatabaseInit = MSACCESS Then
            strQuery = "CREATE TABLE " & strTable & " (" & _
                "CurrencyId char(5) PRIMARY KEY, " & _
                "Name char(50) NOT NULL, " & _
                "Notes varchar(150) NOT NULL, " & _
                "CreateId char(8) NOT NULL, " & _
                "CreateDate datetime NOT NULL, " & _
                "UpdateId char(8) NOT NULL, " & _
                "UpdateDate datetime NOT NULL)"
        ElseIf mdlGlobal.objDatabaseInit = MYSQL Then
            strQuery = "CREATE TABLE `" & strTable & "` (" & _
                "`CurrencyId` char(5) NOT NULL, " & _
                "`Name` char(50) NOT NULL, " & _
                "`Notes` varchar(150) NOT NULL, " & _
                "`CreateId` char(8) NOT NULL, " & _
                "`CreateDate` datetime NOT NULL, " & _
                "`UpdateId` char(8) NOT NULL, " & _
                "`UpdateDate` datetime NOT NULL, " & _
                "PRIMARY KEY(`CurrencyId`)) TYPE=MyISAM;"
        End If
            
        mdlDatabase.CreateTable mdlGlobal.conInventory, strQuery, strTable, mdlGlobal.objDatabaseInit
    End If
End Function

Public Function CreateTMCONVERTCURRENCY(Optional ByVal blnCreate As Boolean = False) As String
    Dim strTable As String
    
    strTable = "TMCONVERTCURRENCY"
    
    CreateTMCONVERTCURRENCY = strTable
    
    If mdlGlobal.UserAuthority.IsAdmin Or blnCreate Then
        Dim strQuery As String
        
        If mdlGlobal.objDatabaseInit = SQLSERVER7 Or _
            mdlGlobal.objDatabaseInit = SQLSERVER2000 Or _
            mdlGlobal.objDatabaseInit = SQLEXPRESS Then
            strQuery = "CREATE TABLE " & strTable & " (" & _
                "ConvertId char(18) PRIMARY KEY, " & _
                "ConvertDate datetime NULL DEFAULT GETDATE(), " & _
                "CurrencyFromId char(5) NOT NULL, " & _
                "CurrencyToId char(5) NOT NULL, " & _
                "ConvertValue money NULL DEFAULT 0, " & _
                "CreateId char(8) NOT NULL, " & _
                "CreateDate datetime NULL DEFAULT GETDATE(), " & _
                "UpdateId char(8) NOT NULL, " & _
                "UpdateDate datetime NULL DEFAULT GETDATE())"
        ElseIf mdlGlobal.objDatabaseInit = MSACCESS Then
            strQuery = "CREATE TABLE " & strTable & " (" & _
                "ConvertId char(18) PRIMARY KEY, " & _
                "ConvertDate datetime NOT NULL, " & _
                "CurrencyFromId char(5) NOT NULL, " & _
                "CurrencyToId char(5) NOT NULL, " & _
                "ConvertValue money NOT NULL, " & _
                "CreateId char(8) NOT NULL, " & _
                "CreateDate datetime NOT NULL, " & _
                "UpdateId char(8) NOT NULL, " & _
                "UpdateDate datetime NOT NULL)"
        ElseIf mdlGlobal.objDatabaseInit = MYSQL Then
            strQuery = "CREATE TABLE `" & strTable & "` (" & _
                "`ConvertId` char(18) NOT NULL, " & _
                "`ConvertDate` datetime NOT NULL, " & _
                "`CurrencyFromId` char(5) NOT NULL, " & _
                "`CurrencyToId` char(5) NOT NULL, " & _
                "`ConvertValue` double NOT NULL, " & _
                "`CreateId` char(8) NOT NULL, " & _
                "`CreateDate` datetime NOT NULL, " & _
                "`UpdateId` char(8) NOT NULL, " & _
                "`UpdateDate` datetime NOT NULL, " & _
                "PRIMARY KEY(`ConvertId`)) TYPE=MyISAM;"
        End If
            
        mdlDatabase.CreateTable mdlGlobal.conInventory, strQuery, strTable, mdlGlobal.objDatabaseInit
    End If
End Function

Public Function CreateTMSTOCKINIT(Optional ByVal blnCreate As Boolean = False) As String
    Dim strTable As String
    
    strTable = "TMSTOCKINIT"
    
    CreateTMSTOCKINIT = strTable
    
    If mdlGlobal.UserAuthority.IsAdmin Or blnCreate Then
        Dim strQuery As String
        
        If mdlGlobal.objDatabaseInit = SQLSERVER7 Or _
            mdlGlobal.objDatabaseInit = SQLSERVER2000 Or _
            mdlGlobal.objDatabaseInit = SQLEXPRESS Then
            strQuery = "CREATE TABLE " & strTable & " (" & _
                "StockInitId char(12) PRIMARY KEY, " & _
                "ItemId char(7) NOT NULL, " & _
                "WarehouseId char(5) NOT NULL, " & _
                "Qty money NULL DEFAULT 0, " & _
                "CreateId char(8) NOT NULL, " & _
                "CreateDate datetime NULL DEFAULT GETDATE(), " & _
                "UpdateId char(8) NOT NULL, " & _
                "UpdateDate datetime NULL DEFAULT GETDATE())"
        ElseIf mdlGlobal.objDatabaseInit = MSACCESS Then
            strQuery = "CREATE TABLE " & strTable & " (" & _
                "StockInitId char(12) PRIMARY KEY, " & _
                "ItemId char(7) NOT NULL, " & _
                "WarehouseId char(5) NOT NULL, " & _
                "Qty money NOT NULL, " & _
                "CreateId char(8) NOT NULL, " & _
                "CreateDate datetime NOT NULL, " & _
                "UpdateId char(8) NOT NULL, " & _
                "UpdateDate datetime NOT NULL)"
        ElseIf mdlGlobal.objDatabaseInit = MYSQL Then
            strQuery = "CREATE TABLE `" & strTable & "` (" & _
                "`StockInitId` char(12) NOT NULL, " & _
                "`ItemId` char(7) NOT NULL, " & _
                "`WarehouseId` char(5) NOT NULL, " & _
                "`Qty` double NOT NULL, " & _
                "`CreateId` char(8) NOT NULL, " & _
                "`CreateDate` datetime NOT NULL, " & _
                "`UpdateId` char(8) NOT NULL, " & _
                "`UpdateDate` datetime NOT NULL, " & _
                "PRIMARY KEY(`StockInitId`)) TYPE=MyISAM;"
        End If
            
        mdlDatabase.CreateTable mdlGlobal.conInventory, strQuery, strTable, mdlGlobal.objDatabaseInit
    End If
End Function

Public Function CreateTHSTOCK(Optional ByVal blnCreate As Boolean = False) As String
    Dim strTable As String
    
    strTable = "THSTOCK"
    
    CreateTHSTOCK = strTable
    
    If mdlGlobal.UserAuthority.IsAdmin Or blnCreate Then
        Dim strQuery As String
        
        If mdlGlobal.objDatabaseInit = SQLSERVER7 Or _
            mdlGlobal.objDatabaseInit = SQLSERVER2000 Or _
            mdlGlobal.objDatabaseInit = SQLEXPRESS Then
            strQuery = "CREATE TABLE " & strTable & " (" & _
                "StockId char(50) PRIMARY KEY, " & _
                "StockDate datetime NULL DEFAULT GETDATE(), " & _
                "ItemId char(7) NOT NULL, " & _
                "WarehouseId char(5) NOT NULL, " & _
                "ReferencesNumber char(30) NOT NULL, " & _
                "QtyIn money NULL DEFAULT 0, " & _
                "QtyOut money NULL DEFAULT 0, " & _
                "CreateId char(8) NOT NULL, " & _
                "CreateDate datetime NULL DEFAULT GETDATE(), " & _
                "UpdateId char(8) NOT NULL, " & _
                "UpdateDate datetime NULL DEFAULT GETDATE())"
        ElseIf mdlGlobal.objDatabaseInit = MSACCESS Then
            strQuery = "CREATE TABLE " & strTable & " (" & _
                "StockId char(50) PRIMARY KEY, " & _
                "StockDate datetime NOT NULL, " & _
                "ItemId char(7) NOT NULL, " & _
                "WarehouseId char(5) NOT NULL, " & _
                "ReferencesNumber char(30) NOT NULL, " & _
                "QtyIn money NOT NULL, " & _
                "QtyOut money NOT NULL, " & _
                "CreateId char(8) NOT NULL, " & _
                "CreateDate datetime NOT NULL, " & _
                "UpdateId char(8) NOT NULL, " & _
                "UpdateDate datetime NOT NULL)"
        ElseIf mdlGlobal.objDatabaseInit = MYSQL Then
            strQuery = "CREATE TABLE `" & strTable & "` (" & _
                "`StockId` char(50) NOT NULL, " & _
                "`StockDate` datetime NOT NULL, " & _
                "`ItemId` char(7) NOT NULL, " & _
                "`WarehouseId` char(5) NOT NULL, " & _
                "`ReferencesNumber` char(30) NOT NULL, " & _
                "`QtyIn` double NOT NULL, " & _
                "`QtyOut` double NOT NULL, " & _
                "`CreateId` char(8) NOT NULL, " & _
                "`CreateDate` datetime NOT NULL, " & _
                "`UpdateId` char(8) NOT NULL, " & _
                "`UpdateDate` datetime NOT NULL, " & _
                "PRIMARY KEY(`StockId`)) TYPE=MyISAM;"
        End If
            
        mdlDatabase.CreateTable mdlGlobal.conInventory, strQuery, strTable, mdlGlobal.objDatabaseInit
    End If
End Function

Public Function CreateTHITEMIN(Optional ByVal blnCreate As Boolean = False) As String
    Dim strTable As String
    
    strTable = "THITEMIN"
    
    CreateTHITEMIN = strTable
    
    If mdlGlobal.UserAuthority.IsAdmin Or blnCreate Then
        Dim strQuery As String
        
        If mdlGlobal.objDatabaseInit = SQLSERVER7 Or _
            mdlGlobal.objDatabaseInit = SQLSERVER2000 Or _
            mdlGlobal.objDatabaseInit = SQLEXPRESS Then
            strQuery = "CREATE TABLE " & strTable & " (" & _
                "ItemInId char(18) PRIMARY KEY, " & _
                "ItemInDate datetime NULL DEFAULT GETDATE(), " & _
                "WarehouseId char(5) NULL DEFAULT '', " & _
                "Notes varchar(150) NULL DEFAULT '', " & _
                "CreateId char(8) NOT NULL, " & _
                "CreateDate datetime NULL DEFAULT GETDATE(), " & _
                "UpdateId char(8) NOT NULL, " & _
                "UpdateDate datetime NULL DEFAULT GETDATE())"
        ElseIf mdlGlobal.objDatabaseInit = MSACCESS Then
            strQuery = "CREATE TABLE " & strTable & " (" & _
                "ItemInId char(18) PRIMARY KEY, " & _
                "ItemInDate datetime NOT NULL, " & _
                "WarehouseId char(5) NOT NULL, " & _
                "Notes varchar(150) NOT NULL, " & _
                "CreateId char(8) NOT NULL, " & _
                "CreateDate datetime NOT NULL, " & _
                "UpdateId char(8) NOT NULL, " & _
                "UpdateDate datetime NOT NULL)"
        ElseIf mdlGlobal.objDatabaseInit = MYSQL Then
            strQuery = "CREATE TABLE `" & strTable & "` (" & _
                "`ItemInId` char(18) NOT NULL, " & _
                "`ItemInDate` datetime NOT NULL, " & _
                "`WarehouseId` char(5) NOT NULL, " & _
                "`Notes` varchar(150) NOT NULL, " & _
                "`CreateId` char(8) NOT NULL, " & _
                "`CreateDate` datetime NOT NULL, " & _
                "`UpdateId` char(8) NOT NULL, " & _
                "`UpdateDate` datetime NOT NULL, " & _
                "PRIMARY KEY(`ItemInId`)) TYPE=MyISAM;"
        End If
            
        mdlDatabase.CreateTable mdlGlobal.conInventory, strQuery, strTable, mdlGlobal.objDatabaseInit
    End If
End Function

Public Function CreateTDITEMIN(Optional ByVal blnCreate As Boolean = False) As String
    Dim strTable As String
    
    strTable = "TDITEMIN"
    
    CreateTDITEMIN = strTable
    
    If mdlGlobal.UserAuthority.IsAdmin Or blnCreate Then
        Dim strQuery As String
        
        If mdlGlobal.objDatabaseInit = SQLSERVER7 Or _
            mdlGlobal.objDatabaseInit = SQLSERVER2000 Or _
            mdlGlobal.objDatabaseInit = SQLEXPRESS Then
            strQuery = "CREATE TABLE " & strTable & " (" & _
                "ItemInDtlId char(25) PRIMARY KEY, " & _
                "ItemInId char(18) NOT NULL, " & _
                "ItemId char(7) NOT NULL, " & _
                "Qty money NULL DEFAULT 0, " & _
                "CreateId char(8) NOT NULL, " & _
                "CreateDate datetime NULL DEFAULT GETDATE(), " & _
                "UpdateId char(8) NOT NULL, " & _
                "UpdateDate datetime NULL DEFAULT GETDATE())"
        ElseIf mdlGlobal.objDatabaseInit = MSACCESS Then
            strQuery = "CREATE TABLE " & strTable & " (" & _
                "ItemInDtlId char(25) PRIMARY KEY, " & _
                "ItemInId char(18) NOT NULL, " & _
                "ItemId char(7) NOT NULL, " & _
                "Qty money NOT NULL, " & _
                "CreateId char(8) NOT NULL, " & _
                "CreateDate datetime NOT NULL, " & _
                "UpdateId char(8) NOT NULL, " & _
                "UpdateDate datetime NOT NULL)"
        ElseIf mdlGlobal.objDatabaseInit = MYSQL Then
            strQuery = "CREATE TABLE `" & strTable & "` (" & _
                "`ItemInDtlId` char(25) NOT NULL, " & _
                "`ItemInId` char(18) NOT NULL, " & _
                "`ItemId` char(7) NOT NULL, " & _
                "`Qty` double NOT NULL, " & _
                "`CreateId` char(8) NOT NULL, " & _
                "`CreateDate` datetime NOT NULL, " & _
                "`UpdateId` char(8) NOT NULL, " & _
                "`UpdateDate` datetime NOT NULL, " & _
                "PRIMARY KEY(`ItemInDtlId`)) TYPE=MyISAM;"
        End If
            
        mdlDatabase.CreateTable mdlGlobal.conInventory, strQuery, strTable, mdlGlobal.objDatabaseInit
    End If
End Function

Public Function CreateTHITEMOUT(Optional ByVal blnCreate As Boolean = False) As String
    Dim strTable As String
    
    strTable = "THITEMOUT"
    
    CreateTHITEMOUT = strTable
    
    If mdlGlobal.UserAuthority.IsAdmin Or blnCreate Then
        Dim strQuery As String
        
        If mdlGlobal.objDatabaseInit = SQLSERVER7 Or _
            mdlGlobal.objDatabaseInit = SQLSERVER2000 Or _
            mdlGlobal.objDatabaseInit = SQLEXPRESS Then
            strQuery = "CREATE TABLE " & strTable & " (" & _
                "ItemOutId char(18) PRIMARY KEY, " & _
                "ItemOutDate datetime NULL DEFAULT GETDATE(), " & _
                "WarehouseId char(5) NULL DEFAULT '', " & _
                "Notes varchar(150) NULL DEFAULT '', " & _
                "CreateId char(8) NOT NULL, " & _
                "CreateDate datetime NULL DEFAULT GETDATE(), " & _
                "UpdateId char(8) NOT NULL, " & _
                "UpdateDate datetime NULL DEFAULT GETDATE())"
        ElseIf mdlGlobal.objDatabaseInit = MSACCESS Then
            strQuery = "CREATE TABLE " & strTable & " (" & _
                "ItemOutId char(18) PRIMARY KEY, " & _
                "ItemOutDate datetime NOT NULL, " & _
                "WarehouseId char(5) NOT NULL, " & _
                "Notes varchar(150) NOT NULL, " & _
                "CreateId char(8) NOT NULL, " & _
                "CreateDate datetime NOT NULL, " & _
                "UpdateId char(8) NOT NULL, " & _
                "UpdateDate datetime NOT NULL)"
        ElseIf mdlGlobal.objDatabaseInit = MYSQL Then
            strQuery = "CREATE TABLE `" & strTable & "` (" & _
                "`ItemOutId` char(18) NOT NULL, " & _
                "`ItemOutDate` datetime NOT NULL, " & _
                "`WarehouseId` char(5) NOT NULL, " & _
                "`Notes` varchar(150) NOT NULL, " & _
                "`CreateId` char(8) NOT NULL, " & _
                "`CreateDate` datetime NOT NULL, " & _
                "`UpdateId` char(8) NOT NULL, " & _
                "`UpdateDate` datetime NOT NULL, " & _
                "PRIMARY KEY(`ItemOutId`)) TYPE=MyISAM;"
        End If
            
        mdlDatabase.CreateTable mdlGlobal.conInventory, strQuery, strTable, mdlGlobal.objDatabaseInit
    End If
End Function

Public Function CreateTDITEMOUT(Optional ByVal blnCreate As Boolean = False) As String
    Dim strTable As String
    
    strTable = "TDITEMOUT"
    
    CreateTDITEMOUT = strTable
    
    If mdlGlobal.UserAuthority.IsAdmin Or blnCreate Then
        Dim strQuery As String
        
        If mdlGlobal.objDatabaseInit = SQLSERVER7 Or _
            mdlGlobal.objDatabaseInit = SQLSERVER2000 Or _
            mdlGlobal.objDatabaseInit = SQLEXPRESS Then
            strQuery = "CREATE TABLE " & strTable & " (" & _
                "ItemOutDtlId char(25) PRIMARY KEY, " & _
                "ItemOutId char(18) NOT NULL, " & _
                "ItemId char(7) NOT NULL, " & _
                "Qty money NULL DEFAULT 0, " & _
                "CreateId char(8) NOT NULL, " & _
                "CreateDate datetime NULL DEFAULT GETDATE(), " & _
                "UpdateId char(8) NOT NULL, " & _
                "UpdateDate datetime NULL DEFAULT GETDATE())"
        ElseIf mdlGlobal.objDatabaseInit = MSACCESS Then
            strQuery = "CREATE TABLE " & strTable & " (" & _
                "ItemOutDtlId char(25) PRIMARY KEY, " & _
                "ItemOutId char(18) NOT NULL, " & _
                "ItemId char(7) NOT NULL, " & _
                "Qty money NOT NULL, " & _
                "CreateId char(8) NOT NULL, " & _
                "CreateDate datetime NOT NULL, " & _
                "UpdateId char(8) NOT NULL, " & _
                "UpdateDate datetime NOT NULL)"
        ElseIf mdlGlobal.objDatabaseInit = MYSQL Then
            strQuery = "CREATE TABLE `" & strTable & "` (" & _
                "`ItemOutDtlId` char(25) NOT NULL, " & _
                "`ItemOutId` char(18) NOT NULL, " & _
                "`ItemId` char(7) NOT NULL, " & _
                "`Qty` double NOT NULL, " & _
                "`CreateId` char(8) NOT NULL, " & _
                "`CreateDate` datetime NOT NULL, " & _
                "`UpdateId` char(8) NOT NULL, " & _
                "`UpdateDate` datetime NOT NULL, " & _
                "PRIMARY KEY(`ItemOutDtlId`)) TYPE=MyISAM;"
        End If
            
        mdlDatabase.CreateTable mdlGlobal.conInventory, strQuery, strTable, mdlGlobal.objDatabaseInit
    End If
End Function

Public Function CreateTHMUTITEM(Optional ByVal blnCreate As Boolean = False) As String
    Dim strTable As String
    
    strTable = "THMUTITEM"
    
    CreateTHMUTITEM = strTable
    
    If mdlGlobal.UserAuthority.IsAdmin Or blnCreate Then
        Dim strQuery As String
        
        If mdlGlobal.objDatabaseInit = SQLSERVER7 Or _
            mdlGlobal.objDatabaseInit = SQLSERVER2000 Or _
            mdlGlobal.objDatabaseInit = SQLEXPRESS Then
            strQuery = "CREATE TABLE " & strTable & " (" & _
                "MutId char(18) PRIMARY KEY, " & _
                "MutDate datetime NULL DEFAULT GETDATE(), " & _
                "WarehouseFrom char(5) NOT NULL, " & _
                "WarehouseTo char(5) NOT NULL, " & _
                "Notes varchar(150) NULL DEFAULT '', " & _
                "CreateId char(8) NOT NULL, " & _
                "CreateDate datetime NULL DEFAULT GETDATE(), " & _
                "UpdateId char(8) NOT NULL, " & _
                "UpdateDate datetime NULL DEFAULT GETDATE())"
        ElseIf mdlGlobal.objDatabaseInit = MSACCESS Then
            strQuery = "CREATE TABLE " & strTable & " (" & _
                "MutId char(18) PRIMARY KEY, " & _
                "MutDate datetime NOT NULL, " & _
                "WarehouseFrom char(5) NOT NULL, " & _
                "WarehouseTo char(5) NOT NULL, " & _
                "Notes varchar(150) NOT NULL, " & _
                "CreateId char(8) NOT NULL, " & _
                "CreateDate datetime NOT NULL, " & _
                "UpdateId char(8) NOT NULL, " & _
                "UpdateDate datetime NOT NULL)"
        ElseIf mdlGlobal.objDatabaseInit = MYSQL Then
            strQuery = "CREATE TABLE `" & strTable & "` (" & _
                "`MutId` char(18) NOT NULL, " & _
                "`MutDate` datetime NOT NULL, " & _
                "`WarehouseFrom` char(5) NOT NULL, " & _
                "`WarehouseTo` char(5) NOT NULL, " & _
                "`Notes` varchar(150) NOT NULL, " & _
                "`CreateId` char(8) NOT NULL, " & _
                "`CreateDate` datetime NOT NULL, " & _
                "`UpdateId` char(8) NOT NULL, " & _
                "`UpdateDate` datetime NOT NULL, " & _
                "PRIMARY KEY(`MutId`)) TYPE=MyISAM;"
        End If
            
        mdlDatabase.CreateTable mdlGlobal.conInventory, strQuery, strTable, mdlGlobal.objDatabaseInit
    End If
End Function

Public Function CreateTDMUTITEM(Optional ByVal blnCreate As Boolean = False) As String
    Dim strTable As String
    
    strTable = "TDMUTITEM"
    
    CreateTDMUTITEM = strTable
    
    If mdlGlobal.UserAuthority.IsAdmin Or blnCreate Then
        Dim strQuery As String
        
        If mdlGlobal.objDatabaseInit = SQLSERVER7 Or _
            mdlGlobal.objDatabaseInit = SQLSERVER2000 Or _
            mdlGlobal.objDatabaseInit = SQLEXPRESS Then
            strQuery = "CREATE TABLE " & strTable & " (" & _
                "MutDtlId char(25) PRIMARY KEY, " & _
                "MutId char(18) NOT NULL, " & _
                "ItemId char(7) NOT NULL, " & _
                "Qty money NULL DEFAULT 0, " & _
                "CreateId char(8) NOT NULL, " & _
                "CreateDate datetime NULL DEFAULT GETDATE(), " & _
                "UpdateId char(8) NOT NULL, " & _
                "UpdateDate datetime NULL DEFAULT GETDATE())"
        ElseIf mdlGlobal.objDatabaseInit = MSACCESS Then
            strQuery = "CREATE TABLE " & strTable & " (" & _
                "MutDtlId char(25) PRIMARY KEY, " & _
                "MutId char(18) NOT NULL, " & _
                "ItemId char(7) NOT NULL, " & _
                "Qty money NOT NULL, " & _
                "CreateId char(8) NOT NULL, " & _
                "CreateDate datetime NOT NULL, " & _
                "UpdateId char(8) NOT NULL, " & _
                "UpdateDate datetime NOT NULL)"
        ElseIf mdlGlobal.objDatabaseInit = MYSQL Then
            strQuery = "CREATE TABLE `" & strTable & "` (" & _
                "`MutDtlId` char(25) NOT NULL, " & _
                "`MutId` char(18) NOT NULL, " & _
                "`ItemId` char(7) NOT NULL, " & _
                "`Qty` double NOT NULL, " & _
                "`CreateId` char(8) NOT NULL, " & _
                "`CreateDate` datetime NOT NULL, " & _
                "`UpdateId` char(8) NOT NULL, " & _
                "`UpdateDate` datetime NOT NULL, " & _
                "PRIMARY KEY(`MutDtlId`)) TYPE=MyISAM;"
        End If
            
        mdlDatabase.CreateTable mdlGlobal.conInventory, strQuery, strTable, mdlGlobal.objDatabaseInit
    End If
End Function

Public Function CreateTHPOBUY(Optional ByVal blnCreate As Boolean = False) As String
    Dim strTable As String
    
    strTable = "THPOBUY"
    
    CreateTHPOBUY = strTable
    
    If mdlGlobal.UserAuthority.IsAdmin Or blnCreate Then
        Dim strQuery As String
        
        If mdlGlobal.objDatabaseInit = SQLSERVER7 Or _
            mdlGlobal.objDatabaseInit = SQLSERVER2000 Or _
            mdlGlobal.objDatabaseInit = SQLEXPRESS Then
            strQuery = "CREATE TABLE " & strTable & " (" & _
                "POId char(20) PRIMARY KEY, " & _
                "PODate datetime NULL DEFAULT GETDATE(), " & _
                "DateLine datetime NULL DEFAULT GETDATE(), " & _
                "VendorId char(6) NULL DEFAULT '', " & _
                "EmployeeBy char(7) NULL DEFAULT '', " & _
                "EmployeeAgree char(7) NULL DEFAULT '', " & _
                "CurrencyId char(5) NULL DEFAULT '', " & _
                "Disc money NULL DEFAULT 0, " & _
                "Tax money NULL DEFAULT 0, " & _
                "Notes varchar(150) NULL DEFAULT '', " & _
                "CreateId char(8) NOT NULL, " & _
                "CreateDate datetime NULL DEFAULT GETDATE(), " & _
                "UpdateId char(8) NOT NULL, " & _
                "UpdateDate datetime NULL DEFAULT GETDATE())"
        ElseIf mdlGlobal.objDatabaseInit = MSACCESS Then
            strQuery = "CREATE TABLE " & strTable & " (" & _
                "POId char(20) PRIMARY KEY, " & _
                "PODate datetime NOT NULL, " & _
                "DateLine datetime NOT NULL, " & _
                "CurrencyId char(5) NOT NULL, " & _
                "EmployeeBy char(7) NOT NULL, " & _
                "EmployeeAgree char(7) NOT NULL, " & _
                "VendorId char(6) NOT NULL, " & _
                "Disc money NOT NULL, " & _
                "Tax money NOT NULL, " & _
                "Notes varchar(150) NOT NULL, " & _
                "CreateId char(8) NOT NULL, " & _
                "CreateDate datetime NOT NULL, " & _
                "UpdateId char(8) NOT NULL, " & _
                "UpdateDate datetime NOT NULL)"
        ElseIf mdlGlobal.objDatabaseInit = MYSQL Then
            strQuery = "CREATE TABLE `" & strTable & "` (" & _
                "`POId` char(20) NOT NULL, " & _
                "`PODate` datetime NOT NULL, " & _
                "`DateLine` datetime NOT NULL, " & _
                "`CurrencyId` char(5) NOT NULL, " & _
                "`EmployeeBy` char(7) NOT NULL, " & _
                "`EmployeeAgree` char(7) NOT NULL, " & _
                "`VendorId` char(6) NOT NULL, " & _
                "`Disc` double NOT NULL, " & _
                "`Tax` double NOT NULL, " & _
                "`Notes` varchar(150) NOT NULL, " & _
                "`CreateId` char(8) NOT NULL, " & _
                "`CreateDate` datetime NOT NULL, " & _
                "`UpdateId` char(8) NOT NULL, " & _
                "`UpdateDate` datetime NOT NULL, " & _
                "PRIMARY KEY(`POId`)) TYPE=MyISAM;"
        End If
            
        mdlDatabase.CreateTable mdlGlobal.conInventory, strQuery, strTable, mdlGlobal.objDatabaseInit
    End If
End Function

Public Function CreateTDPOBUY(Optional ByVal blnCreate As Boolean = False) As String
    Dim strTable As String
    
    strTable = "TDPOBUY"
    
    CreateTDPOBUY = strTable
    
    If mdlGlobal.UserAuthority.IsAdmin Or blnCreate Then
        Dim strQuery As String
        
        If mdlGlobal.objDatabaseInit = SQLSERVER7 Or _
            mdlGlobal.objDatabaseInit = SQLSERVER2000 Or _
            mdlGlobal.objDatabaseInit = SQLEXPRESS Then
            strQuery = "CREATE TABLE " & strTable & " (" & _
                "PODtlId char(27) PRIMARY KEY, " & _
                "POId char(20) NOT NULL, " & _
                "ItemId char(7) NOT NULL, " & _
                "Qty money NULL DEFAULT 0, " & _
                "PriceId char(15) NULL DEFAULT '', " & _
                "CreateId char(8) NOT NULL, " & _
                "CreateDate datetime NULL DEFAULT GETDATE(), " & _
                "UpdateId char(8) NOT NULL, " & _
                "UpdateDate datetime NULL DEFAULT GETDATE())"
        ElseIf mdlGlobal.objDatabaseInit = MSACCESS Then
            strQuery = "CREATE TABLE " & strTable & " (" & _
                "PODtlId char(27) PRIMARY KEY, " & _
                "POId char(20) NOT NULL, " & _
                "ItemId char(7) NOT NULL, " & _
                "Qty money NOT NULL, " & _
                "PriceId char(15) NOT NULL, " & _
                "CreateId char(8) NOT NULL, " & _
                "CreateDate datetime NOT NULL, " & _
                "UpdateId char(8) NOT NULL, " & _
                "UpdateDate datetime NOT NULL)"
        ElseIf mdlGlobal.objDatabaseInit = MYSQL Then
            strQuery = "CREATE TABLE `" & strTable & "` (" & _
                "`PODtlId` char(27) NOT NULL, " & _
                "`POId` char(20) NOT NULL, " & _
                "`ItemId` char(7) NOT NULL, " & _
                "`Qty` double NOT NULL, " & _
                "`PriceId` char(15) NOT NULL, " & _
                "`CreateId` char(8) NOT NULL, " & _
                "`CreateDate` datetime NOT NULL, " & _
                "`UpdateId` char(8) NOT NULL, " & _
                "`UpdateDate` datetime NOT NULL, " & _
                "PRIMARY KEY(`PODtlId`)) TYPE=MyISAM;"
        End If
            
        mdlDatabase.CreateTable mdlGlobal.conInventory, strQuery, strTable, mdlGlobal.objDatabaseInit
    End If
End Function

Public Function CreateTHDOBUY(Optional ByVal blnCreate As Boolean = False) As String
    Dim strTable As String
    
    strTable = "THDOBUY"
    
    CreateTHDOBUY = strTable
    
    If mdlGlobal.UserAuthority.IsAdmin Or blnCreate Then
        Dim strQuery As String
        
        If mdlGlobal.objDatabaseInit = SQLSERVER7 Or _
            mdlGlobal.objDatabaseInit = SQLSERVER2000 Or _
            mdlGlobal.objDatabaseInit = SQLEXPRESS Then
            strQuery = "CREATE TABLE " & strTable & " (" & _
                "DOId char(20) PRIMARY KEY, " & _
                "DODate datetime NULL DEFAULT GETDATE(), " & _
                "POId char(20) NULL DEFAULT '', " & _
                "WarehouseId char(5) NULL DEFAULT '', " & _
                "TransportNumber char(20) NULL DEFAULT '', " & _
                "Notes varchar(150) NULL DEFAULT '', " & _
                "CreateId char(8) NOT NULL, " & _
                "CreateDate datetime NULL DEFAULT GETDATE(), " & _
                "UpdateId char(8) NOT NULL, " & _
                "UpdateDate datetime NULL DEFAULT GETDATE())"
        ElseIf mdlGlobal.objDatabaseInit = MSACCESS Then
            strQuery = "CREATE TABLE " & strTable & " (" & _
                "DOId char(20) PRIMARY KEY, " & _
                "DODate datetime NOT NULL, " & _
                "POId char(20) NOT NULL, " & _
                "WarehouseId char(5) NOT NULL, " & _
                "TransportNumber char(20) NOT NULL, " & _
                "Notes varchar(150) NOT NULL, " & _
                "CreateId char(8) NOT NULL, " & _
                "CreateDate datetime NOT NULL, " & _
                "UpdateId char(8) NOT NULL, " & _
                "UpdateDate datetime NOT NULL)"
        ElseIf mdlGlobal.objDatabaseInit = MYSQL Then
            strQuery = "CREATE TABLE `" & strTable & "` (" & _
                "`DOId` char(20) NOT NULL, " & _
                "`DODate` datetime NOT NULL, " & _
                "`POId` char(20) NOT NULL, " & _
                "`WarehouseId` char(5) NOT NULL, " & _
                "`TransportNumber` char(20) NOT NULL, " & _
                "`Notes` varchar(150) NOT NULL, " & _
                "`CreateId` char(8) NOT NULL, " & _
                "`CreateDate` datetime NOT NULL, " & _
                "`UpdateId` char(8) NOT NULL, " & _
                "`UpdateDate` datetime NOT NULL, " & _
                "PRIMARY KEY(`DOId`)) TYPE=MyISAM;"
        End If
        
        mdlDatabase.CreateTable mdlGlobal.conInventory, strQuery, strTable, mdlGlobal.objDatabaseInit
    End If
End Function

Public Function CreateTDDOBUY(Optional ByVal blnCreate As Boolean = False) As String
    Dim strTable As String
    
    strTable = "TDDOBUY"
    
    CreateTDDOBUY = strTable
    
    If mdlGlobal.UserAuthority.IsAdmin Or blnCreate Then
        Dim strQuery As String
        
        If mdlGlobal.objDatabaseInit = SQLSERVER7 Or _
            mdlGlobal.objDatabaseInit = SQLSERVER2000 Or _
            mdlGlobal.objDatabaseInit = SQLEXPRESS Then
            strQuery = "CREATE TABLE " & strTable & " (" & _
                "DODtlId char(27) PRIMARY KEY, " & _
                "DOId char(20) NOT NULL, " & _
                "ItemId char(7) NOT NULL, " & _
                "Qty money NULL DEFAULT 0, " & _
                "CreateId char(8) NOT NULL, " & _
                "CreateDate datetime NULL DEFAULT GETDATE(), " & _
                "UpdateId char(8) NOT NULL, " & _
                "UpdateDate datetime NULL DEFAULT GETDATE())"
        ElseIf mdlGlobal.objDatabaseInit = MSACCESS Then
            strQuery = "CREATE TABLE " & strTable & " (" & _
                "DODtlId char(27) PRIMARY KEY, " & _
                "DOId char(20) NOT NULL, " & _
                "ItemId char(7) NOT NULL, " & _
                "Qty money NOT NULL, " & _
                "CreateId char(8) NOT NULL, " & _
                "CreateDate datetime NOT NULL, " & _
                "UpdateId char(8) NOT NULL, " & _
                "UpdateDate datetime NOT NULL)"
        ElseIf mdlGlobal.objDatabaseInit = MYSQL Then
            strQuery = "CREATE TABLE `" & strTable & "` (" & _
                "`DODtlId` char(27) NOT NULL, " & _
                "`DOId` char(20) NOT NULL, " & _
                "`ItemId` char(7) NOT NULL, " & _
                "`Qty` double NOT NULL, " & _
                "`PriceId` char(15) NOT NULL, " & _
                "`CreateId` char(8) NOT NULL, " & _
                "`CreateDate` datetime NOT NULL, " & _
                "`UpdateId` char(8) NOT NULL, " & _
                "`UpdateDate` datetime NOT NULL, " & _
                "PRIMARY KEY(`PODtlId`)) TYPE=MyISAM;"
        End If
            
        mdlDatabase.CreateTable mdlGlobal.conInventory, strQuery, strTable, mdlGlobal.objDatabaseInit
    End If
End Function

Public Function CreateTHSJBUY(Optional ByVal blnCreate As Boolean = False) As String
    Dim strTable As String
    
    strTable = "THSJBUY"
    
    CreateTHSJBUY = strTable
    
    If mdlGlobal.UserAuthority.IsAdmin Or blnCreate Then
        Dim strQuery As String
        
        If mdlGlobal.objDatabaseInit = SQLSERVER7 Or _
            mdlGlobal.objDatabaseInit = SQLSERVER2000 Or _
            mdlGlobal.objDatabaseInit = SQLEXPRESS Then
            strQuery = "CREATE TABLE " & strTable & " (" & _
                "SJId char(20) PRIMARY KEY, " & _
                "SJDate datetime NULL DEFAULT GETDATE(), " & _
                "DOId char(20) NULL DEFAULT '', " & _
                "Notes varchar(150) NULL DEFAULT '', " & _
                "CreateId char(8) NOT NULL, " & _
                "CreateDate datetime NULL DEFAULT GETDATE(), " & _
                "UpdateId char(8) NOT NULL, " & _
                "UpdateDate datetime NULL DEFAULT GETDATE())"
        ElseIf mdlGlobal.objDatabaseInit = MSACCESS Then
            strQuery = "CREATE TABLE " & strTable & " (" & _
                "SJId char(20) PRIMARY KEY, " & _
                "SJDate datetime NOT NULL, " & _
                "DOId char(20) NOT NULL, " & _
                "Notes varchar(150) NOT NULL, " & _
                "CreateId char(8) NOT NULL, " & _
                "CreateDate datetime NOT NULL, " & _
                "UpdateId char(8) NOT NULL, " & _
                "UpdateDate datetime NOT NULL)"
        ElseIf mdlGlobal.objDatabaseInit = MYSQL Then
            strQuery = "CREATE TABLE `" & strTable & "` (" & _
                "`SJId` char(20) NOT NULL, " & _
                "`SJDate` datetime NOT NULL, " & _
                "`DOId` char(20) NOT NULL, " & _
                "`Notes` varchar(150) NOT NULL, " & _
                "`CreateId` char(8) NOT NULL, " & _
                "`CreateDate` datetime NOT NULL, " & _
                "`UpdateId` char(8) NOT NULL, " & _
                "`UpdateDate` datetime NOT NULL, " & _
                "PRIMARY KEY(`SJId`)) TYPE=MyISAM;"
        End If
        
        mdlDatabase.CreateTable mdlGlobal.conInventory, strQuery, strTable, mdlGlobal.objDatabaseInit
    End If
End Function

Public Function CreateTDSJBUY(Optional ByVal blnCreate As Boolean = False) As String
    Dim strTable As String
    
    strTable = "TDSJBUY"
    
    CreateTDSJBUY = strTable
    
    If mdlGlobal.UserAuthority.IsAdmin Or blnCreate Then
        Dim strQuery As String
        
        If mdlGlobal.objDatabaseInit = SQLSERVER7 Or _
            mdlGlobal.objDatabaseInit = SQLSERVER2000 Or _
            mdlGlobal.objDatabaseInit = SQLEXPRESS Then
            strQuery = "CREATE TABLE " & strTable & " (" & _
                "SJDtlId char(27) PRIMARY KEY, " & _
                "SJId char(20) NOT NULL, " & _
                "ItemId char(7) NOT NULL, " & _
                "Qty money NULL DEFAULT 0, " & _
                "CreateId char(8) NOT NULL, " & _
                "CreateDate datetime NULL DEFAULT GETDATE(), " & _
                "UpdateId char(8) NOT NULL, " & _
                "UpdateDate datetime NULL DEFAULT GETDATE())"
        ElseIf mdlGlobal.objDatabaseInit = MSACCESS Then
            strQuery = "CREATE TABLE " & strTable & " (" & _
                "SJDtlId char(27) PRIMARY KEY, " & _
                "SJId char(20) NOT NULL, " & _
                "ItemId char(7) NOT NULL, " & _
                "Qty money NOT NULL, " & _
                "CreateId char(8) NOT NULL, " & _
                "CreateDate datetime NOT NULL, " & _
                "UpdateId char(8) NOT NULL, " & _
                "UpdateDate datetime NOT NULL)"
        ElseIf mdlGlobal.objDatabaseInit = MYSQL Then
            strQuery = "CREATE TABLE `" & strTable & "` (" & _
                "`SJDtlId` char(27) NOT NULL, " & _
                "`SJId` char(20) NOT NULL, " & _
                "`ItemId` char(7) NOT NULL, " & _
                "`Qty` double NOT NULL, " & _
                "`CreateId` char(8) NOT NULL, " & _
                "`CreateDate` datetime NOT NULL, " & _
                "`UpdateId` char(8) NOT NULL, " & _
                "`UpdateDate` datetime NOT NULL, " & _
                "PRIMARY KEY(`SJDtlId`)) TYPE=MyISAM;"
        End If
            
        mdlDatabase.CreateTable mdlGlobal.conInventory, strQuery, strTable, mdlGlobal.objDatabaseInit
    End If
End Function

Public Function CreateTHFKTBUY(Optional ByVal blnCreate As Boolean = False) As String
    Dim strTable As String
    
    strTable = "THFKTBUY"
    
    CreateTHFKTBUY = strTable
    
    If mdlGlobal.UserAuthority.IsAdmin Or blnCreate Then
        Dim strQuery As String
        
        If mdlGlobal.objDatabaseInit = SQLSERVER7 Or _
            mdlGlobal.objDatabaseInit = SQLSERVER2000 Or _
            mdlGlobal.objDatabaseInit = SQLEXPRESS Then
            strQuery = "CREATE TABLE " & strTable & " (" & _
                "FktId char(20) PRIMARY KEY, " & _
                "FktDate datetime NULL DEFAULT GETDATE(), " & _
                "VendorId char(6) NULL DEFAULT '', " & _
                "Notes varchar(150) NULL DEFAULT '', " & _
                "CreateId char(8) NOT NULL, " & _
                "CreateDate datetime NULL DEFAULT GETDATE(), " & _
                "UpdateId char(8) NOT NULL, " & _
                "UpdateDate datetime NULL DEFAULT GETDATE())"
        ElseIf mdlGlobal.objDatabaseInit = MSACCESS Then
            strQuery = "CREATE TABLE " & strTable & " (" & _
                "FktId char(20) PRIMARY KEY, " & _
                "FktDate datetime NOT NULL, " & _
                "VendorId char(6) NOT NULL, " & _
                "Notes varchar(150) NOT NULL, " & _
                "CreateId char(8) NOT NULL, " & _
                "CreateDate datetime NOT NULL, " & _
                "UpdateId char(8) NOT NULL, " & _
                "UpdateDate datetime NOT NULL)"
        ElseIf mdlGlobal.objDatabaseInit = MYSQL Then
            strQuery = "CREATE TABLE `" & strTable & "` (" & _
                "`FktId` char(20) NOT NULL, " & _
                "`FktDate` datetime NOT NULL, " & _
                "`VendorId` char(6) NOT NULL, " & _
                "`Notes` varchar(150) NOT NULL, " & _
                "`CreateId` char(8) NOT NULL, " & _
                "`CreateDate` datetime NOT NULL, " & _
                "`UpdateId` char(8) NOT NULL, " & _
                "`UpdateDate` datetime NOT NULL, " & _
                "PRIMARY KEY(`FktId`)) TYPE=MyISAM;"
        End If
            
        mdlDatabase.CreateTable mdlGlobal.conInventory, strQuery, strTable, mdlGlobal.objDatabaseInit
    End If
End Function

Public Function CreateTDFKTBUY(Optional ByVal blnCreate As Boolean = False) As String
    Dim strTable As String
    
    strTable = "TDFKTBUY"
    
    CreateTDFKTBUY = strTable
    
    If mdlGlobal.UserAuthority.IsAdmin Or blnCreate Then
        Dim strQuery As String
        
        If mdlGlobal.objDatabaseInit = SQLSERVER7 Or _
            mdlGlobal.objDatabaseInit = SQLSERVER2000 Or _
            mdlGlobal.objDatabaseInit = SQLEXPRESS Then
            strQuery = "CREATE TABLE " & strTable & " (" & _
                "FktDtlId char(40) PRIMARY KEY, " & _
                "FktId char(20) NOT NULL, " & _
                "SJId char(20) NOT NULL, " & _
                "CreateId char(8) NOT NULL, " & _
                "CreateDate datetime NULL DEFAULT GETDATE(), " & _
                "UpdateId char(8) NOT NULL, " & _
                "UpdateDate datetime NULL DEFAULT GETDATE())"
        ElseIf mdlGlobal.objDatabaseInit = MSACCESS Then
            strQuery = "CREATE TABLE " & strTable & " (" & _
                "FktDtlId char(40) PRIMARY KEY, " & _
                "FktId char(20) NOT NULL, " & _
                "SJId char(20) NOT NULL, " & _
                "CreateId char(8) NOT NULL, " & _
                "CreateDate datetime NOT NULL, " & _
                "UpdateId char(8) NOT NULL, " & _
                "UpdateDate datetime NOT NULL)"
        ElseIf mdlGlobal.objDatabaseInit = MYSQL Then
            strQuery = "CREATE TABLE `" & strTable & "` (" & _
                "`FktDtlId` char(40) NOT NULL, " & _
                "`FktId` char(20) NOT NULL, " & _
                "`SJId` char(20) NOT NULL, " & _
                "`CreateId` char(8) NOT NULL, " & _
                "`CreateDate` datetime NOT NULL, " & _
                "`UpdateId` char(8) NOT NULL, " & _
                "`UpdateDate` datetime NOT NULL, " & _
                "PRIMARY KEY(`FktDtlId`)) TYPE=MyISAM;"
        End If
            
        mdlDatabase.CreateTable mdlGlobal.conInventory, strQuery, strTable, mdlGlobal.objDatabaseInit
    End If
End Function

Public Function CreateTHRTRBUY(Optional ByVal blnCreate As Boolean = False) As String
    Dim strTable As String
    
    strTable = "THRTRBUY"
    
    CreateTHRTRBUY = strTable
    
    If mdlGlobal.UserAuthority.IsAdmin Or blnCreate Then
        Dim strQuery As String
        
        If mdlGlobal.objDatabaseInit = SQLSERVER7 Or _
            mdlGlobal.objDatabaseInit = SQLSERVER2000 Or _
            mdlGlobal.objDatabaseInit = SQLEXPRESS Then
            strQuery = "CREATE TABLE " & strTable & " (" & _
                "RtrId char(20) PRIMARY KEY, " & _
                "RtrDate datetime NULL DEFAULT GETDATE(), " & _
                "SJId char(20) NULL DEFAULT '', " & _
                "Notes varchar(150) NULL DEFAULT '', " & _
                "CreateId char(8) NOT NULL, " & _
                "CreateDate datetime NULL DEFAULT GETDATE(), " & _
                "UpdateId char(8) NOT NULL, " & _
                "UpdateDate datetime NULL DEFAULT GETDATE())"
        ElseIf mdlGlobal.objDatabaseInit = MSACCESS Then
            strQuery = "CREATE TABLE " & strTable & " (" & _
                "RtrId char(20) PRIMARY KEY, " & _
                "RtrDate datetime NOT NULL, " & _
                "SJId char(20) NOT NULL, " & _
                "Notes varchar(150) NOT NULL, " & _
                "CreateId char(8) NOT NULL, " & _
                "CreateDate datetime NOT NULL, " & _
                "UpdateId char(8) NOT NULL, " & _
                "UpdateDate datetime NOT NULL)"
        ElseIf mdlGlobal.objDatabaseInit = MYSQL Then
            strQuery = "CREATE TABLE `" & strTable & "` (" & _
                "`RtrId` char(20) NOT NULL, " & _
                "`RtrDate` datetime NOT NULL, " & _
                "`SJId` char(20) NOT NULL, " & _
                "`Notes` varchar(150) NOT NULL, " & _
                "`CreateId` char(8) NOT NULL, " & _
                "`CreateDate` datetime NOT NULL, " & _
                "`UpdateId` char(8) NOT NULL, " & _
                "`UpdateDate` datetime NOT NULL, " & _
                "PRIMARY KEY(`RtrId`)) TYPE=MyISAM;"
        End If
            
        mdlDatabase.CreateTable mdlGlobal.conInventory, strQuery, strTable, mdlGlobal.objDatabaseInit
    End If
End Function

Public Function CreateTDRTRBUY(Optional ByVal blnCreate As Boolean = False) As String
    Dim strTable As String
    
    strTable = "TDRTRBUY"
    
    CreateTDRTRBUY = strTable
    
    If mdlGlobal.UserAuthority.IsAdmin Or blnCreate Then
        Dim strQuery As String
        
        If mdlGlobal.objDatabaseInit = SQLSERVER7 Or _
            mdlGlobal.objDatabaseInit = SQLSERVER2000 Or _
            mdlGlobal.objDatabaseInit = SQLEXPRESS Then
            strQuery = "CREATE TABLE " & strTable & " (" & _
                "RtrDtlId char(27) PRIMARY KEY, " & _
                "RtrId char(20) NOT NULL, " & _
                "ItemId char(7) NOT NULL, " & _
                "Qty money NULL DEFAULT '', " & _
                "CreateId char(8) NOT NULL, " & _
                "CreateDate datetime NULL DEFAULT GETDATE(), " & _
                "UpdateId char(8) NOT NULL, " & _
                "UpdateDate datetime NULL DEFAULT GETDATE())"
        ElseIf mdlGlobal.objDatabaseInit = MSACCESS Then
            strQuery = "CREATE TABLE " & strTable & " (" & _
                "RtrDtlId char(27) PRIMARY KEY, " & _
                "RtrId char(20) NOT NULL, " & _
                "ItemId char(7) NOT NULL, " & _
                "Qty money NOT NULL, " & _
                "CreateId char(8) NOT NULL, " & _
                "CreateDate datetime NOT NULL, " & _
                "UpdateId char(8) NOT NULL, " & _
                "UpdateDate datetime NOT NULL)"
        ElseIf mdlGlobal.objDatabaseInit = MYSQL Then
            strQuery = "CREATE TABLE `" & strTable & "` (" & _
                "`RtrDtlId` char(27) NOT NULL, " & _
                "`RtrId` char(20) NOT NULL, " & _
                "`ItemId` char(7) NOT NULL, " & _
                "`Qty` double NOT NULL, " & _
                "`CreateId` char(8) NOT NULL, " & _
                "`CreateDate` datetime NOT NULL, " & _
                "`UpdateId` char(8) NOT NULL, " & _
                "`UpdateDate` datetime NOT NULL, " & _
                "PRIMARY KEY(`RtrDtlId`)) TYPE=MyISAM;"
        End If
            
        mdlDatabase.CreateTable mdlGlobal.conInventory, strQuery, strTable, mdlGlobal.objDatabaseInit
    End If
End Function

Public Function CreateTHSALESSUM(Optional ByVal blnCreate As Boolean = False) As String
    Dim strTable As String
    
    strTable = "THSALESSUM"
    
    CreateTHSALESSUM = strTable
    
    If mdlGlobal.UserAuthority.IsAdmin Or blnCreate Then
        Dim strQuery As String
        
        If mdlGlobal.objDatabaseInit = SQLSERVER7 Or _
            mdlGlobal.objDatabaseInit = SQLSERVER2000 Or _
            mdlGlobal.objDatabaseInit = SQLEXPRESS Then
            strQuery = "CREATE TABLE " & strTable & " (" & _
                "POId char(20) PRIMARY KEY, " & _
                "PODate datetime NULL DEFAULT GETDATE(), " & _
                "CustomerId char(6) NULL DEFAULT '', " & _
                "POCustomerId char(20) NULL DEFAULT '', " & _
                "PriceValue money NULL DEFAULT 0, " & _
                "CurrencyId char(5) NULL DEFAULT '', " & _
                "Notes varchar(150) NULL DEFAULT '', " & _
                "CreateId char(8) NOT NULL, " & _
                "CreateDate datetime NULL DEFAULT GETDATE(), " & _
                "UpdateId char(8) NOT NULL, " & _
                "UpdateDate datetime NULL DEFAULT GETDATE())"
        ElseIf mdlGlobal.objDatabaseInit = MSACCESS Then
            strQuery = "CREATE TABLE " & strTable & " (" & _
                "POId char(20) PRIMARY KEY, " & _
                "PODate datetime NOT NULL, " & _
                "CustomerId char(6) NOT NULL, " & _
                "POCustomerId char(20) NOT NULL, " & _
                "PriceValue money NOT NULL, " & _
                "CurrencyId char(5) NOT NULL , " & _
                "Notes varchar(150) NOT NULL, " & _
                "CreateId char(8) NOT NULL, " & _
                "CreateDate datetime NOT NULL, " & _
                "UpdateId char(8) NOT NULL, " & _
                "UpdateDate datetime NOT NULL)"
        ElseIf mdlGlobal.objDatabaseInit = MYSQL Then
            strQuery = "CREATE TABLE `" & strTable & "` (" & _
                "`POId` char(20) NOT NULL, " & _
                "`PODate` datetime NOT NULL, " & _
                "`CustomerId` char(6) NOT NULL, " & _
                "`POCustomerId` char(20) NOT NULL, " & _
                "`PriceValue` double NOT NULL, " & _
                "`CurrencyId` char(5) NOT NULL, " & _
                "`Notes` varchar(150) NOT NULL, " & _
                "`CreateId` char(8) NOT NULL, " & _
                "`CreateDate` datetime NOT NULL, " & _
                "`UpdateId` char(8) NOT NULL, " & _
                "`UpdateDate` datetime NOT NULL, " & _
                "PRIMARY KEY(`POId`)) TYPE=MyISAM;"
        End If
            
        mdlDatabase.CreateTable mdlGlobal.conInventory, strQuery, strTable, mdlGlobal.objDatabaseInit
    End If
End Function

Public Function CreateTHPOSELL(Optional ByVal blnCreate As Boolean = False) As String
    Dim strTable As String
    
    strTable = "THPOSELL"
    
    CreateTHPOSELL = strTable
    
    If mdlGlobal.UserAuthority.IsAdmin Or blnCreate Then
        Dim strQuery As String
        
        If mdlGlobal.objDatabaseInit = SQLSERVER7 Or _
            mdlGlobal.objDatabaseInit = SQLSERVER2000 Or _
            mdlGlobal.objDatabaseInit = SQLEXPRESS Then
            strQuery = "CREATE TABLE " & strTable & " (" & _
                "POId char(20) PRIMARY KEY, " & _
                "PODate datetime NULL DEFAULT GETDATE(), " & _
                "CustomerId char(6) NULL DEFAULT '', " & _
                "POCustomerId char(20) NULL DEFAULT '', " & _
                "DateLine datetime NULL DEFAULT GETDATE(), " & _
                "Notes varchar(150) NULL DEFAULT '', " & _
                "CreateId char(8) NOT NULL, " & _
                "CreateDate datetime NULL DEFAULT GETDATE(), " & _
                "UpdateId char(8) NOT NULL, " & _
                "UpdateDate datetime NULL DEFAULT GETDATE())"
        ElseIf mdlGlobal.objDatabaseInit = MSACCESS Then
            strQuery = "CREATE TABLE " & strTable & " (" & _
                "POId char(20) PRIMARY KEY, " & _
                "PODate datetime NOT NULL, " & _
                "CustomerId char(6) NOT NULL, " & _
                "POCustomerId char(20) NOT NULL, " & _
                "DateLine datetime NOT NULL, " & _
                "Notes varchar(150) NOT NULL, " & _
                "CreateId char(8) NOT NULL, " & _
                "CreateDate datetime NOT NULL, " & _
                "UpdateId char(8) NOT NULL, " & _
                "UpdateDate datetime NOT NULL)"
        ElseIf mdlGlobal.objDatabaseInit = MYSQL Then
            strQuery = "CREATE TABLE `" & strTable & "` (" & _
                "`POId` char(20) NOT NULL, " & _
                "`PODate` datetime NOT NULL, " & _
                "`CustomerId` char(6) NOT NULL, " & _
                "`POCustomerId` char(20) NOT NULL, " & _
                "`DateLine` datetime NOT NULL, " & _
                "`Notes` varchar(150) NOT NULL, " & _
                "`CreateId` char(8) NOT NULL, " & _
                "`CreateDate` datetime NOT NULL, " & _
                "`UpdateId` char(8) NOT NULL, " & _
                "`UpdateDate` datetime NOT NULL, " & _
                "PRIMARY KEY(`POId`)) TYPE=MyISAM;"
        End If
            
        mdlDatabase.CreateTable mdlGlobal.conInventory, strQuery, strTable, mdlGlobal.objDatabaseInit
    End If
End Function

Public Function CreateTDPOSELL(Optional ByVal blnCreate As Boolean = False) As String
    Dim strTable As String
    
    strTable = "TDPOSELL"
    
    CreateTDPOSELL = strTable
    
    If mdlGlobal.UserAuthority.IsAdmin Or blnCreate Then
        Dim strQuery As String
        
        If mdlGlobal.objDatabaseInit = SQLSERVER7 Or _
            mdlGlobal.objDatabaseInit = SQLSERVER2000 Or _
            mdlGlobal.objDatabaseInit = SQLEXPRESS Then
            strQuery = "CREATE TABLE " & strTable & " (" & _
                "PODtlId char(27) PRIMARY KEY, " & _
                "POId char(20) NOT NULL, " & _
                "ItemId char(7) NOT NULL, " & _
                "Qty money NULL DEFAULT 0, " & _
                "CreateId char(8) NOT NULL, " & _
                "CreateDate datetime NULL DEFAULT GETDATE(), " & _
                "UpdateId char(8) NOT NULL, " & _
                "UpdateDate datetime NULL DEFAULT GETDATE())"
        ElseIf mdlGlobal.objDatabaseInit = MSACCESS Then
            strQuery = "CREATE TABLE " & strTable & " (" & _
                "PODtlId char(27) PRIMARY KEY, " & _
                "POId char(20) NOT NULL, " & _
                "ItemId char(7) NOT NULL, " & _
                "Qty money NOT NULL, " & _
                "CreateId char(8) NOT NULL, " & _
                "CreateDate datetime NOT NULL, " & _
                "UpdateId char(8) NOT NULL, " & _
                "UpdateDate datetime NOT NULL)"
        ElseIf mdlGlobal.objDatabaseInit = MYSQL Then
            strQuery = "CREATE TABLE `" & strTable & "` (" & _
                "`PODtlId` char(27) NOT NULL, " & _
                "`POId` char(20) NOT NULL, " & _
                "`ItemId` char(7) NOT NULL, " & _
                "`Qty` double NOT NULL, " & _
                "`CreateId` char(8) NOT NULL, " & _
                "`CreateDate` datetime NOT NULL, " & _
                "`UpdateId` char(8) NOT NULL, " & _
                "`UpdateDate` datetime NOT NULL, " & _
                "PRIMARY KEY(`PODtlId`)) TYPE=MyISAM;"
        End If
            
        mdlDatabase.CreateTable mdlGlobal.conInventory, strQuery, strTable, mdlGlobal.objDatabaseInit
    End If
End Function

Public Function CreateTHSOSELL(Optional ByVal blnCreate As Boolean = False) As String
    Dim strTable As String
    
    strTable = "THSOSELL"
    
    CreateTHSOSELL = strTable
    
    If mdlGlobal.UserAuthority.IsAdmin Or blnCreate Then
        Dim strQuery As String
        
        If mdlGlobal.objDatabaseInit = SQLSERVER7 Or _
            mdlGlobal.objDatabaseInit = SQLSERVER2000 Or _
            mdlGlobal.objDatabaseInit = SQLEXPRESS Then
            strQuery = "CREATE TABLE " & strTable & " (" & _
                "SOId char(20) PRIMARY KEY, " & _
                "SODate datetime NULL DEFAULT GETDATE(), " & _
                "POId char(20) NULL DEFAULT '', " & _
                "Tax money NULL DEFAULT 0, " & _
                "Disc money NULL DEFAULT 0, " & _
                "CurrencyId char(5) NULL DEFAULT '', " & _
                "Notes varchar(150) NULL DEFAULT '', " & _
                "CreateId char(8) NOT NULL, " & _
                "CreateDate datetime NULL DEFAULT GETDATE(), " & _
                "UpdateId char(8) NOT NULL, " & _
                "UpdateDate datetime NULL DEFAULT GETDATE())"
        ElseIf mdlGlobal.objDatabaseInit = MSACCESS Then
            strQuery = "CREATE TABLE " & strTable & " (" & _
                "SOId char(20) PRIMARY KEY, " & _
                "SODate datetime NOT NULL, " & _
                "POId char(20) NOT NULL, " & _
                "Tax money NOT NULL, " & _
                "Disc money NOT NULL, " & _
                "CurrencyId char(5) NOT NULL, " & _
                "Notes varchar(150) NOT NULL, " & _
                "CreateId char(8) NOT NULL, " & _
                "CreateDate datetime NOT NULL, " & _
                "UpdateId char(8) NOT NULL, " & _
                "UpdateDate datetime NOT NULL)"
        ElseIf mdlGlobal.objDatabaseInit = MYSQL Then
            strQuery = "CREATE TABLE `" & strTable & "` (" & _
                "`SOId` char(20) NOT NULL, " & _
                "`SODate` datetime NOT NULL, " & _
                "`POId` char(20) NOT NULL, " & _
                "`Tax` double NOT NULL, " & _
                "`Disc` double NOT NULL, " & _
                "`CurrencyId` char(5) NOT NULL, " & _
                "`Notes` varchar(150) NOT NULL, " & _
                "`CreateId` char(8) NOT NULL, " & _
                "`CreateDate` datetime NOT NULL, " & _
                "`UpdateId` char(8) NOT NULL, " & _
                "`UpdateDate` datetime NOT NULL, " & _
                "PRIMARY KEY(`SOId`)) TYPE=MyISAM;"
        End If
            
        mdlDatabase.CreateTable mdlGlobal.conInventory, strQuery, strTable, mdlGlobal.objDatabaseInit
    End If
End Function

Public Function CreateTDSOSELL(Optional ByVal blnCreate As Boolean = False) As String
    Dim strTable As String
    
    strTable = "TDSOSELL"
    
    CreateTDSOSELL = strTable
    
    If mdlGlobal.UserAuthority.IsAdmin Or blnCreate Then
        Dim strQuery As String
        
        If mdlGlobal.objDatabaseInit = SQLSERVER7 Or _
            mdlGlobal.objDatabaseInit = SQLSERVER2000 Or _
            mdlGlobal.objDatabaseInit = SQLEXPRESS Then
            strQuery = "CREATE TABLE " & strTable & " (" & _
                "SODtlId char(27) PRIMARY KEY, " & _
                "SOId char(20) NOT NULL, " & _
                "ItemId char(7) NOT NULL, " & _
                "Qty money NULL DEFAULT 0, " & _
                "PriceId char(15) NULL DEFAULT '', " & _
                "CreateId char(8) NOT NULL, " & _
                "CreateDate datetime NULL DEFAULT GETDATE(), " & _
                "UpdateId char(8) NOT NULL, " & _
                "UpdateDate datetime NULL DEFAULT GETDATE())"
        ElseIf mdlGlobal.objDatabaseInit = MSACCESS Then
            strQuery = "CREATE TABLE " & strTable & " (" & _
                "SODtlId char(27) PRIMARY KEY, " & _
                "SOId char(20) NOT NULL, " & _
                "ItemId char(7) NOT NULL, " & _
                "Qty money NOT NULL, " & _
                "PriceId char(15) NOT NULL, " & _
                "CreateId char(8) NOT NULL, " & _
                "CreateDate datetime NOT NULL, " & _
                "UpdateId char(8) NOT NULL, " & _
                "UpdateDate datetime NOT NULL)"
        ElseIf mdlGlobal.objDatabaseInit = MYSQL Then
            strQuery = "CREATE TABLE `" & strTable & "` (" & _
                "`SODtlId` char(27) NOT NULL, " & _
                "`SOId` char(20) NOT NULL, " & _
                "`ItemId` char(7) NOT NULL, " & _
                "`Qty` double NOT NULL, " & _
                "`PriceId` char(15) NOT NULL, " & _
                "`CreateId` char(8) NOT NULL, " & _
                "`CreateDate` datetime NOT NULL, " & _
                "`UpdateId` char(8) NOT NULL, " & _
                "`UpdateDate` datetime NOT NULL, " & _
                "PRIMARY KEY(`SODtlId`)) TYPE=MyISAM;"
        End If
            
        mdlDatabase.CreateTable mdlGlobal.conInventory, strQuery, strTable, mdlGlobal.objDatabaseInit
    End If
End Function

Public Function CreateTHSJSELL(Optional ByVal blnCreate As Boolean = False) As String
    Dim strTable As String
    
    strTable = "THSJSELL"
    
    CreateTHSJSELL = strTable
    
    If mdlGlobal.UserAuthority.IsAdmin Or blnCreate Then
        Dim strQuery As String
        
        If mdlGlobal.objDatabaseInit = SQLSERVER7 Or _
            mdlGlobal.objDatabaseInit = SQLSERVER2000 Or _
            mdlGlobal.objDatabaseInit = SQLEXPRESS Then
            strQuery = "CREATE TABLE " & strTable & " (" & _
                "SJId char(20) PRIMARY KEY, " & _
                "SJDate datetime NULL DEFAULT GETDATE(), " & _
                "SOId char(20) NULL DEFAULT '', " & _
                "ReferencesNumber char(20) NULL DEFAULT '', " & _
                "DeliveryId char(9) NULL DEFAULT '', " & _
                "Notes varchar(150) NULL DEFAULT '', " & _
                "CreateId char(8) NOT NULL, " & _
                "CreateDate datetime NULL DEFAULT GETDATE(), " & _
                "UpdateId char(8) NOT NULL, " & _
                "UpdateDate datetime NULL DEFAULT GETDATE())"
        ElseIf mdlGlobal.objDatabaseInit = MSACCESS Then
            strQuery = "CREATE TABLE " & strTable & " (" & _
                "SJId char(20) PRIMARY KEY, " & _
                "SJDate datetime NOT NULL, " & _
                "SOId char(20) NOT NULL, " & _
                "ReferencesNumber char(20) NOT NULL, " & _
                "DeliveryId char(9) NOT NULL, " & _
                "Notes varchar(150) NOT NULL, " & _
                "CreateId char(8) NOT NULL, " & _
                "CreateDate datetime NOT NULL, " & _
                "UpdateId char(8) NOT NULL, " & _
                "UpdateDate datetime NOT NULL)"
        ElseIf mdlGlobal.objDatabaseInit = MYSQL Then
            strQuery = "CREATE TABLE `" & strTable & "` (" & _
                "`SJId` char(20) NOT NULL, " & _
                "`SJDate` datetime NOT NULL, " & _
                "`SOId` char(20) NOT NULL, " & _
                "`ReferencesNumber` char(20) NOT NULL, " & _
                "`DeliveryId` char(9) NOT NULL, " & _
                "`Notes` varchar(150) NOT NULL, " & _
                "`CreateId` char(8) NOT NULL, " & _
                "`CreateDate` datetime NOT NULL, " & _
                "`UpdateId` char(8) NOT NULL, " & _
                "`UpdateDate` datetime NOT NULL, " & _
                "PRIMARY KEY(`SJId`)) TYPE=MyISAM;"
        End If
            
        mdlDatabase.CreateTable mdlGlobal.conInventory, strQuery, strTable, mdlGlobal.objDatabaseInit
    End If
End Function

Public Function CreateTDSJSELL(Optional ByVal blnCreate As Boolean = False) As String
    Dim strTable As String
    
    strTable = "TDSJSELL"
    
    CreateTDSJSELL = strTable
    
    If mdlGlobal.UserAuthority.IsAdmin Or blnCreate Then
        Dim strQuery As String
        
        If mdlGlobal.objDatabaseInit = SQLSERVER7 Or _
            mdlGlobal.objDatabaseInit = SQLSERVER2000 Or _
            mdlGlobal.objDatabaseInit = SQLEXPRESS Then
            strQuery = "CREATE TABLE " & strTable & " (" & _
                "SJDtlId char(27) PRIMARY KEY, " & _
                "SJId char(20) NOT NULL, " & _
                "ItemId char(7) NOT NULL, " & _
                "WarehouseId char(5) NULL DEFAULT '', " & _
                "Qty money NULL DEFAULT 0, " & _
                "CreateId char(8) NOT NULL, " & _
                "CreateDate datetime NULL DEFAULT GETDATE(), " & _
                "UpdateId char(8) NOT NULL, " & _
                "UpdateDate datetime NULL DEFAULT GETDATE())"
        ElseIf mdlGlobal.objDatabaseInit = MSACCESS Then
            strQuery = "CREATE TABLE " & strTable & " (" & _
                "SJDtlId char(27) PRIMARY KEY, " & _
                "SJId char(20) NOT NULL, " & _
                "ItemId char(7) NOT NULL, " & _
                "WarehouseId char(5) NOT NULL, " & _
                "Qty money NOT NULL, " & _
                "CreateId char(8) NOT NULL, " & _
                "CreateDate datetime NOT NULL, " & _
                "UpdateId char(8) NOT NULL, " & _
                "UpdateDate datetime NOT NULL)"
        ElseIf mdlGlobal.objDatabaseInit = MYSQL Then
            strQuery = "CREATE TABLE `" & strTable & "` (" & _
                "`SJDtlId` char(27) NOT NULL, " & _
                "`SJId` char(20) NOT NULL, " & _
                "`ItemId` char(7) NOT NULL, " & _
                "`WarehouseId` char(5) NOT NULL, " & _
                "`Qty` double NOT NULL, " & _
                "`CreateId` char(8) NOT NULL, " & _
                "`CreateDate` datetime NOT NULL, " & _
                "`UpdateId` char(8) NOT NULL, " & _
                "`UpdateDate` datetime NOT NULL, " & _
                "PRIMARY KEY(`SJDtlId`)) TYPE=MyISAM;"
        End If
            
        mdlDatabase.CreateTable mdlGlobal.conInventory, strQuery, strTable, mdlGlobal.objDatabaseInit
    End If
End Function

Public Function CreateTHFKTSELL(Optional ByVal blnCreate As Boolean = False) As String
    Dim strTable As String
    
    strTable = "THFKTSELL"
    
    CreateTHFKTSELL = strTable
    
    If mdlGlobal.UserAuthority.IsAdmin Or blnCreate Then
        Dim strQuery As String
        
        If mdlGlobal.objDatabaseInit = SQLSERVER7 Or _
            mdlGlobal.objDatabaseInit = SQLSERVER2000 Or _
            mdlGlobal.objDatabaseInit = SQLEXPRESS Then
            strQuery = "CREATE TABLE " & strTable & " (" & _
                "FktId char(20) PRIMARY KEY, " & _
                "FktDate datetime NULL DEFAULT GETDATE(), " & _
                "CustomerId char(6) NULL DEFAULT '', " & _
                "Notes varchar(150) NULL DEFAULT '', " & _
                "CreateId char(8) NOT NULL, " & _
                "CreateDate datetime NULL DEFAULT GETDATE(), " & _
                "UpdateId char(8) NOT NULL, " & _
                "UpdateDate datetime NULL DEFAULT GETDATE())"
        ElseIf mdlGlobal.objDatabaseInit = MSACCESS Then
            strQuery = "CREATE TABLE " & strTable & " (" & _
                "FktId char(20) PRIMARY KEY, " & _
                "FktDate datetime NOT NULL, " & _
                "CustomerId char(6) NOT NULL, " & _
                "Notes varchar(150) NOT NULL, " & _
                "CreateId char(8) NOT NULL, " & _
                "CreateDate datetime NOT NULL, " & _
                "UpdateId char(8) NOT NULL, " & _
                "UpdateDate datetime NOT NULL)"
        ElseIf mdlGlobal.objDatabaseInit = MYSQL Then
            strQuery = "CREATE TABLE `" & strTable & "` (" & _
                "`FktId` char(20) NOT NULL, " & _
                "`FktDate` datetime NOT NULL, " & _
                "`CustomerId` char(6) NOT NULL, " & _
                "`Notes` varchar(150) NOT NULL, " & _
                "`CreateId` char(8) NOT NULL, " & _
                "`CreateDate` datetime NOT NULL, " & _
                "`UpdateId` char(8) NOT NULL, " & _
                "`UpdateDate` datetime NOT NULL, " & _
                "PRIMARY KEY(`FktId`)) TYPE=MyISAM;"
        End If
            
        mdlDatabase.CreateTable mdlGlobal.conInventory, strQuery, strTable, mdlGlobal.objDatabaseInit
    End If
End Function

Public Function CreateTDFKTSELL(Optional ByVal blnCreate As Boolean = False) As String
    Dim strTable As String
    
    strTable = "TDFKTSELL"
    
    CreateTDFKTSELL = strTable
    
    If mdlGlobal.UserAuthority.IsAdmin Or blnCreate Then
        Dim strQuery As String
        
        If mdlGlobal.objDatabaseInit = SQLSERVER7 Or _
            mdlGlobal.objDatabaseInit = SQLSERVER2000 Or _
            mdlGlobal.objDatabaseInit = SQLEXPRESS Then
            strQuery = "CREATE TABLE " & strTable & " (" & _
                "FktDtlId char(40) PRIMARY KEY, " & _
                "FktId char(20) NOT NULL, " & _
                "SJId char(20) NOT NULL, " & _
                "CreateId char(8) NOT NULL, " & _
                "CreateDate datetime NULL DEFAULT GETDATE(), " & _
                "UpdateId char(8) NOT NULL, " & _
                "UpdateDate datetime NULL DEFAULT GETDATE())"
        ElseIf mdlGlobal.objDatabaseInit = MSACCESS Then
            strQuery = "CREATE TABLE " & strTable & " (" & _
                "FktDtlId char(40) PRIMARY KEY, " & _
                "FktId char(20) NOT NULL, " & _
                "SJId char(20) NOT NULL, " & _
                "CreateId char(8) NOT NULL, " & _
                "CreateDate datetime NOT NULL, " & _
                "UpdateId char(8) NOT NULL, " & _
                "UpdateDate datetime NOT NULL)"
        ElseIf mdlGlobal.objDatabaseInit = MYSQL Then
            strQuery = "CREATE TABLE `" & strTable & "` (" & _
                "`FktDtlId` char(40) NOT NULL, " & _
                "`FktId` char(20) NOT NULL, " & _
                "`SJId` char(20) NOT NULL, " & _
                "`CreateId` char(8) NOT NULL, " & _
                "`CreateDate` datetime NOT NULL, " & _
                "`UpdateId` char(8) NOT NULL, " & _
                "`UpdateDate` datetime NOT NULL, " & _
                "PRIMARY KEY(`FktDtlId`)) TYPE=MyISAM;"
        End If
            
        mdlDatabase.CreateTable mdlGlobal.conInventory, strQuery, strTable, mdlGlobal.objDatabaseInit
    End If
End Function

Public Function CreateTHRTRSELL(Optional ByVal blnCreate As Boolean = False) As String
    Dim strTable As String
    
    strTable = "THRTRSELL"
    
    CreateTHRTRSELL = strTable
    
    If mdlGlobal.UserAuthority.IsAdmin Or blnCreate Then
        Dim strQuery As String
        
        If mdlGlobal.objDatabaseInit = SQLSERVER7 Or _
            mdlGlobal.objDatabaseInit = SQLSERVER2000 Or _
            mdlGlobal.objDatabaseInit = SQLEXPRESS Then
            strQuery = "CREATE TABLE " & strTable & " (" & _
                "RtrId char(20) PRIMARY KEY, " & _
                "RtrDate datetime NULL DEFAULT GETDATE(), " & _
                "SJId char(20) NULL DEFAULT '', " & _
                "WarehouseId char(5) NULL DEFAULT '', " & _
                "Notes varchar(150) NULL DEFAULT '', " & _
                "CreateId char(8) NOT NULL, " & _
                "CreateDate datetime NULL DEFAULT GETDATE(), " & _
                "UpdateId char(8) NOT NULL, " & _
                "UpdateDate datetime NULL DEFAULT GETDATE())"
        ElseIf mdlGlobal.objDatabaseInit = MSACCESS Then
            strQuery = "CREATE TABLE " & strTable & " (" & _
                "RtrId char(20) PRIMARY KEY, " & _
                "RtrDate datetime NOT NULL, " & _
                "SJId char(20) NOT NULL, " & _
                "WarehouseId char(5) NOT NULL, " & _
                "Notes varchar(150) NOT NULL, " & _
                "CreateId char(8) NOT NULL, " & _
                "CreateDate datetime NOT NULL, " & _
                "UpdateId char(8) NOT NULL, " & _
                "UpdateDate datetime NOT NULL)"
        ElseIf mdlGlobal.objDatabaseInit = MYSQL Then
            strQuery = "CREATE TABLE `" & strTable & "` (" & _
                "`RtrId` char(20) NOT NULL, " & _
                "`RtrDate` datetime NOT NULL, " & _
                "`SJId` char(20) NOT NULL, " & _
                "`WarehouseId` char(5) NOT NULL, " & _
                "`Notes` char(150) NOT NULL, " & _
                "`CreateId` char(8) NOT NULL, " & _
                "`CreateDate` datetime NOT NULL, " & _
                "`UpdateId` char(8) NOT NULL, " & _
                "`UpdateDate` datetime NOT NULL, " & _
                "PRIMARY KEY(`RtrId`)) TYPE=MyISAM;"
        End If
            
        mdlDatabase.CreateTable mdlGlobal.conInventory, strQuery, strTable, mdlGlobal.objDatabaseInit
    End If
End Function

Public Function CreateTDRTRSELL(Optional ByVal blnCreate As Boolean = False) As String
    Dim strTable As String
    
    strTable = "TDRTRSELL"
    
    CreateTDRTRSELL = strTable
    
    If mdlGlobal.UserAuthority.IsAdmin Or blnCreate Then
        Dim strQuery As String
        
        If mdlGlobal.objDatabaseInit = SQLSERVER7 Or _
            mdlGlobal.objDatabaseInit = SQLSERVER2000 Or _
            mdlGlobal.objDatabaseInit = SQLEXPRESS Then
            strQuery = "CREATE TABLE " & strTable & " (" & _
                "RtrDtlId char(27) PRIMARY KEY, " & _
                "RtrId char(20) NOT NULL, " & _
                "ItemId char(7) NOT NULL, " & _
                "Qty money NULL DEFAULT '', " & _
                "CreateId char(8) NOT NULL, " & _
                "CreateDate datetime NULL DEFAULT GETDATE(), " & _
                "UpdateId char(8) NOT NULL, " & _
                "UpdateDate datetime NULL DEFAULT GETDATE())"
        ElseIf mdlGlobal.objDatabaseInit = MSACCESS Then
            strQuery = "CREATE TABLE " & strTable & " (" & _
                "RtrDtlId char(27) PRIMARY KEY, " & _
                "RtrId char(20) NOT NULL, " & _
                "ItemId char(7) NOT NULL, " & _
                "Qty money NOT NULL, " & _
                "CreateId char(8) NOT NULL, " & _
                "CreateDate datetime NOT NULL, " & _
                "UpdateId char(8) NOT NULL, " & _
                "UpdateDate datetime NOT NULL)"
        ElseIf mdlGlobal.objDatabaseInit = MYSQL Then
            strQuery = "CREATE TABLE `" & strTable & "` (" & _
                "`RtrDtlId` char(27) NOT NULL, " & _
                "`RtrId` char(20) NOT NULL, " & _
                "`ItemId` char(7) NOT NULL, " & _
                "`Qty` double NOT NULL, " & _
                "`CreateId` char(8) NOT NULL, " & _
                "`CreateDate` datetime NOT NULL, " & _
                "`UpdateId` char(8) NOT NULL, " & _
                "`UpdateDate` datetime NOT NULL, " & _
                "PRIMARY KEY(`RtrDtlId`)) TYPE=MyISAM;"
        End If
            
        mdlDatabase.CreateTable mdlGlobal.conInventory, strQuery, strTable, mdlGlobal.objDatabaseInit
    End If
End Function

Public Function CreateTHRECYCLE(Optional ByVal blnCreate As Boolean = False) As String
    Dim strTable As String
    
    strTable = "THRECYCLE"
    
    CreateTHRECYCLE = strTable
    
    If mdlGlobal.UserAuthority.IsAdmin Or blnCreate Then
        Dim strQuery As String
        
        If mdlGlobal.objDatabaseInit = SQLSERVER7 Or _
            mdlGlobal.objDatabaseInit = SQLSERVER2000 Or _
            mdlGlobal.objDatabaseInit = SQLEXPRESS Then
            strQuery = "CREATE TABLE " & strTable & " (" & _
                "RecycleId char(46) PRIMARY KEY, " & _
                "ReferencesNumber char(30) NOT NULL, " & _
                "RecycleDate datetime NULL DEFAULT GETDATE(), " & _
                "ReferencesDate datetime NULL DEFAULT GETDATE(), " & _
                "OptInfoFirst varchar(150) NULL DEFAULT '', " & _
                "OptInfoSecond varchar(150) NULL DEFAULT '', " & _
                "OptInfoThird varchar(150) NULL DEFAULT '', " & _
                "OptInfoFourth varchar(150) NULL DEFAULT '', " & _
                "OptInfoFifth varchar(150) NULL DEFAULT '', " & _
                "OptInfoSixth varchar(150) NULL DEFAULT '', " & _
                "OptInfoSeventh varchar(150) NULL DEFAULT '', " & _
                "OptInfoEight varchar(150) NULL DEFAULT '', " & _
                "OptInfoNineth varchar(150) NULL DEFAULT '', " & _
                "OptInfoTenth varchar(150) NULL DEFAULT '', " & _
                "CreateId char(8) NOT NULL, " & _
                "CreateDate datetime NULL DEFAULT GETDATE(), " & _
                "UpdateId char(8) NOT NULL, " & _
                "UpdateDate datetime NULL DEFAULT GETDATE())"
        ElseIf mdlGlobal.objDatabaseInit = MSACCESS Then
            strQuery = "CREATE TABLE " & strTable & " (" & _
                "RecycleId char(46) PRIMARY KEY, " & _
                "ReferencesNumber char(30) NOT NULL, " & _
                "RecycleDate datetime NOT NULL, " & _
                "ReferencesDate datetime NOT NULL, " & _
                "OptInfoFirst varchar(150) NOT NULL, " & _
                "OptInfoSecond varchar(150) NOT NULL, " & _
                "OptInfoThird varchar(150) NOT NULL, " & _
                "OptInfoFourth varchar(150) NOT NULL, " & _
                "OptInfoFifth varchar(150) NOT NULL, " & _
                "OptInfoSixth varchar(150) NOT NULL, " & _
                "OptInfoSeventh varchar(150) NOT NULL, " & _
                "OptInfoEight varchar(150) NOT NULL, " & _
                "OptInfoNineth varchar(150) NOT NULL, " & _
                "OptInfoTenth varchar(150) NOT NULL, " & _
                "CreateId char(8) NOT NULL, " & _
                "CreateDate datetime NOT NULL, " & _
                "UpdateId char(8) NOT NULL, " & _
                "UpdateDate datetime NOT NULL)"
        ElseIf mdlGlobal.objDatabaseInit = MYSQL Then
            strQuery = "CREATE TABLE `" & strTable & "` (" & _
                "`RecycleId` char(46) NOT NULL, " & _
                "`ReferencesNumber` char(30) NOT NULL, " & _
                "`RecycleDate` datetime NOT NULL, " & _
                "`ReferencesDate` varchar(150) NOT NULL, " & _
                "`OptInfoFirst` varchar(150) NOT NULL, " & _
                "`OptInfoSecond` varchar(150) NOT NULL, " & _
                "`OptInfoThird` varchar(150) NOT NULL, " & _
                "`OptInfoFourth` varchar(150) NOT NULL, " & _
                "`OptInfoFifth` varchar(150) NOT NULL, " & _
                "`OptInfoSixth` varchar(150) NOT NULL, " & _
                "`OptInfoSeventh` varchar(150) NOT NULL, " & _
                "`OptInfoEight` varchar(150) NOT NULL, " & _
                "`OptInfoNineth` varchar(150) NOT NULL, " & _
                "`OptInfoTenth` varchar(150) NOT NULL, " & _
                "`CreateId` char(8) NOT NULL, " & _
                "`CreateDate` datetime NOT NULL, " & _
                "`UpdateId` char(8) NOT NULL, " & _
                "`UpdateDate` datetime NOT NULL, " & _
                "PRIMARY KEY(`RecycleId`)) TYPE=MyISAM;"
        End If
            
        mdlDatabase.CreateTable mdlGlobal.conInventory, strQuery, strTable, mdlGlobal.objDatabaseInit
    End If
End Function

Public Function CreateTDRECYCLE(Optional ByVal blnCreate As Boolean = False) As String
    Dim strTable As String
    
    strTable = "TDRECYCLE"
    
    CreateTDRECYCLE = strTable
    
    If mdlGlobal.UserAuthority.IsAdmin Or blnCreate Then
        Dim strQuery As String
        
        If mdlGlobal.objDatabaseInit = SQLSERVER7 Or _
            mdlGlobal.objDatabaseInit = SQLSERVER2000 Or _
            mdlGlobal.objDatabaseInit = SQLEXPRESS Then
            strQuery = "CREATE TABLE " & strTable & " (" & _
                "RecycleDtlId char(76) PRIMARY KEY, " & _
                "RecycleId char(46) NOT NULL, " & _
                "ReferencesNumber char(30) NOT NULL, " & _
                "OptInfoFirst varchar(150) NULL DEFAULT '', " & _
                "OptInfoSecond varchar(150) NULL DEFAULT '', " & _
                "OptInfoThird varchar(150) NULL DEFAULT '', " & _
                "OptInfoFourth varchar(150) NULL DEFAULT '', " & _
                "OptInfoFifth varchar(150) NULL DEFAULT '', " & _
                "CreateId char(8) NOT NULL, " & _
                "CreateDate datetime NULL DEFAULT GETDATE(), " & _
                "UpdateId char(8) NOT NULL, " & _
                "UpdateDate datetime NULL DEFAULT GETDATE())"
        ElseIf mdlGlobal.objDatabaseInit = MSACCESS Then
            strQuery = "CREATE TABLE " & strTable & " (" & _
                "RecycleDtlId char(76) PRIMARY KEY, " & _
                "RecycleId char(46) NOT NULL, " & _
                "ReferencesNumber char(30) NOT NULL, " & _
                "OptInfoFirst varchar(150) NOT NULL, " & _
                "OptInfoSecond varchar(150) NOT NULL, " & _
                "OptInfoThird varchar(150) NOT NULL, " & _
                "OptInfoFourth varchar(150) NOT NULL, " & _
                "OptInfoFifth varchar(150) NOT NULL, " & _
                "CreateId char(8) NOT NULL, " & _
                "CreateDate datetime NOT NULL, " & _
                "UpdateId char(8) NOT NULL, " & _
                "UpdateDate datetime NOT NULL)"
        ElseIf mdlGlobal.objDatabaseInit = MYSQL Then
            strQuery = "CREATE TABLE `" & strTable & "` (" & _
                "`RecycleDtlId` char(76) NOT NULL, " & _
                "`RecycleId` char(46) NOT NULL, " & _
                "`ReferencesNumber` char(30) NOT NULL, " & _
                "`OptInfoFirst` varchar(150) NOT NULL, " & _
                "`OptInfoSecond` varchar(150) NOT NULL, " & _
                "`OptInfoThird` varchar(150) NOT NULL, " & _
                "`OptInfoFourth` varchar(150) NOT NULL, " & _
                "`OptInfoFifth` varchar(150) NOT NULL, " & _
                "`CreateId` char(8) NOT NULL, " & _
                "`CreateDate` datetime NOT NULL, " & _
                "`UpdateId` char(8) NOT NULL, " & _
                "`UpdateDate` datetime NOT NULL, " & _
                "PRIMARY KEY(`RecycleDtlId`)) TYPE=MyISAM;"
        End If
            
        mdlDatabase.CreateTable mdlGlobal.conInventory, strQuery, strTable, mdlGlobal.objDatabaseInit
    End If
End Function
