Attribute VB_Name = "mdlDatabase"
Option Explicit

Public Enum SQLDATABASE
    SQLSERVER7
    SQLSERVER2000
    SQLEXPRESS
    MSACCESS
    MYSQL
End Enum

Public Function CreateDatabase( _
    ByVal strDatabase As String, _
    Optional ByVal strServer As String = "localhost", _
    Optional ByVal strUserId As String = "sa", _
    Optional ByVal strPassword As String = "", _
    Optional objDatabase As SQLDATABASE = SQLSERVER7) As Boolean
    Dim conSQL As New ADODB.Connection
    
    On Local Error GoTo ErrHandler
    
    Dim strQuery As String
    
    strQuery = "IF NOT EXISTS (SELECT * FROM master..sysdatabases WHERE Name='" & _
        strDatabase & "') CREATE DATABASE " & strDatabase
    
    If objDatabase = SQLSERVER7 Or objDatabase = SQLSERVER2000 Then
        conSQL.Open "Provider=SQLOLEDB.1" & _
            ";Persist Security Info=False" & _
            ";Data Source=" & strServer & _
            ";User ID=" & strUserId & _
            ";Password=" & strPassword & _
            ";Integrated Security=SSPI"
            
        conSQL.Execute strQuery
        
        mdlDatabase.CloseConnection conSQL
        
        CreateDatabase = True
    ElseIf objDatabase = SQLEXPRESS Then
        If Trim(strUserId) = "" Then
            conSQL.Open "Provider=SQLOLEDB.1" & _
                ";Data Source=" & strServer & "\SQLEXPRESS" & _
                ";Integrated Security=SSPI"
        Else
            conSQL.Open "Provider=SQLOLEDB.1" & _
                ";Data Source=" & strServer & "\SQLEXPRESS" & _
                ";User ID=" & strUserId & _
                ";Password=" & strPassword & _
                ";Integrated Security=SSPI"
        End If
            
        conSQL.Execute strQuery
            
        CreateDatabase = True
    ElseIf objDatabase = MSACCESS Then
        If mdlGlobal.fso.FileExists(mdlGlobal.strPath & strDatabase & ".mdb") Then
            CreateDatabase = True
        Else
            CreateDatabase = False
        End If
    ElseIf objDatabase = MYSQL Then
        strQuery = "CREATE DATABASE IF NOT EXISTS " & strDatabase
    
        conSQL.CursorLocation = adUseClient
        
        conSQL.Open "Provider=MSDASQL.1" & _
            ";Persist Security Info=False" & _
            ";Data Source=MySQL ODBC 5.1 Driver"
        conSQL.Execute strQuery
            
        CreateDatabase = True
    End If
    
    Exit Function
    
ErrHandler:
    MsgBox Err.Number & " : " & Err.Description, vbExclamation, ""
    
    CreateDatabase = False
End Function

Public Sub CreateTable( _
    ByVal conSQL As ADODB.Connection, _
    ByVal strFields As String, _
    ByVal strTable As String, _
    Optional objDatabase As SQLDATABASE = SQLSERVER7)
    On Local Error GoTo ErrHandler
    
    Dim strQuery As String
    
    If objDatabase = SQLSERVER7 Or objDatabase = SQLSERVER2000 Or objDatabase = SQLEXPRESS Then
        strQuery = "IF NOT EXISTS (SELECT * FROM " & conSQL.DefaultDatabase & ".dbo.sysobjects WHERE Name='" & _
            strTable & "') BEGIN " & strFields & " END"
    ElseIf objDatabase = MSACCESS Then
        strQuery = strFields
    ElseIf objDatabase = MYSQL Then
        strQuery = strFields
    End If
    
    conSQL.Execute strQuery
    
    Exit Sub
    
ErrHandler:
End Sub

Public Function IsTableExist(ByRef conSQL As ADODB.Connection, ByVal strTable As String) As Boolean
    Dim rstSQL As ADODB.Recordset
    
    Set rstSQL = conSQL.OpenSchema(adSchemaTables, Array(Empty, Empty, Empty, "TABLE"))
    
    While Not rstSQL.EOF
        If rstSQL!TABLE_NAME = strTable Then
            IsTableExist = True
            
            mdlDatabase.CloseRecordset rstSQL
            
            Exit Function
        End If
        
        rstSQL.MoveNext
    Wend
    
    IsTableExist = False
    
    mdlDatabase.CloseRecordset rstSQL
End Function

Public Function OpenConnection( _
    ByVal strDatabase As String, _
    Optional ByVal strServer As String = "localhost", _
    Optional ByVal strUserId As String = "sa", _
    Optional ByVal strPassword As String = "", _
    Optional ByVal objDatabase As SQLDATABASE = SQLSERVER7) As ADODB.Connection
    Set OpenConnection = New ADODB.Connection
    
    If objDatabase = SQLSERVER7 Or objDatabase = SQLSERVER2000 Then
        OpenConnection.Open "Provider=SQLOLEDB.1" & _
            ";Persist Security Info=False" & _
            ";Initial Catalog=" & strDatabase & _
            ";Data Source=" & strServer & _
            ";User ID=" & strUserId & _
            ";Password=" & strPassword & _
            ";Integrated Security=SSPI"
    ElseIf objDatabase = SQLEXPRESS Then
        OpenConnection.Open "Provider=SQLOLEDB.1" & _
            ";Data Source=" & strServer & "\SQLEXPRESS" & _
            ";Database=" & strDatabase & _
            ";Integrated Security=SSPI"
    ElseIf objDatabase = MSACCESS Then
        OpenConnection.Provider = "Microsoft.Jet.OLEDB.4.0"
        OpenConnection.Open strServer
    ElseIf objDatabase = MYSQL Then
        OpenConnection.Open "Provider=MSDASQL.1" & _
            ";Persist Security Info=False" & _
            ";Data Source=MySQL ODBC 5.1 Driver" & _
            ";Database=" & strDatabase
    End If
End Function

Public Function OpenRecordset( _
    ByRef conSQL As ADODB.Connection, _
    ByVal strFields As String, _
    ByVal strTable As String, _
    Optional ByVal blnActive As Boolean = True, _
    Optional ByVal strCriteria As String = "", _
    Optional ByVal strOrderBy As String = "", _
    Optional ByVal strGroupBy As String = "") As ADODB.Recordset
    On Local Error GoTo ErrHandler
    
    mdlDatabase.CloseRecordset OpenRecordset
    
    Set OpenRecordset = New ADODB.Recordset
    OpenRecordset.CursorLocation = adUseClient
    
    Dim strQuery As String
    
    strQuery = "SELECT " & strFields & " FROM " & strTable
    
    If Not Trim(strCriteria) = "" Then strQuery = strQuery & " WHERE " & strCriteria
    If Not Trim(strOrderBy) = "" Then strQuery = strQuery & " ORDER BY " & strOrderBy
    If Not Trim(strGroupBy) = "" Then strQuery = strQuery & " GROUP BY " & strGroupBy
    
    If mdlGlobal.objDatabaseInit = MYSQL Then
        OpenRecordset.Open strQuery, conSQL, adOpenStatic, adLockOptimistic
    Else
        OpenRecordset.Open strQuery, conSQL, adOpenDynamic, adLockOptimistic
    End If
    
    If Not blnActive Then Set OpenRecordset.ActiveConnection = Nothing
    
    If OpenRecordset.RecordCount > 0 Then OpenRecordset.MoveFirst
    
    Exit Function
    
ErrHandler:
    MsgBox "Terdapat Masalah Dalam Query" & vbCrLf & strQuery, vbCritical + vbOKOnly, "Masalah"
    
    End
End Function

Public Function GetFieldData( _
    ByRef conSQL As ADODB.Connection, _
    ByVal strField As String, _
    ByVal strTable As String, _
    ByVal strCriteria As String, _
    Optional ByVal strOrderBy As String = "") As String
    Dim rstSQL As ADODB.Recordset
    
    Set rstSQL = mdlDatabase.OpenRecordset(conSQL, strField, strTable, False, strCriteria, strOrderBy)
    
    If rstSQL.RecordCount > 0 Then
        rstSQL.MoveFirst
        
        GetFieldData = rstSQL.Fields(0).Value
    Else
        GetFieldData = ""
    End If
    
    mdlDatabase.CloseRecordset rstSQL
End Function

Public Sub SearchRecordset( _
    ByRef rstSQL As ADODB.Recordset, _
    ByVal strKey As String, _
    ByVal strValue As String)
    With rstSQL
        If .RecordCount > 0 Then .MoveFirst
        
        .Find strKey & "='" & strValue & "'"
    End With
End Sub

Public Function IsDataExists( _
    ByRef conSQL As ADODB.Connection, _
    ByVal strTable As String, _
    Optional ByVal strCriteria As String = "") As Boolean
    Dim rstSQL As ADODB.Recordset
    
    Set rstSQL = mdlDatabase.OpenRecordset(conSQL, "*", strTable, False, strCriteria)
    
    If rstSQL.RecordCount > 0 Then
        IsDataExists = True
    Else
        IsDataExists = False
    End If
    
    mdlDatabase.CloseRecordset rstSQL
End Function

Public Function IsDataCorrect( _
    ByRef conSQL As ADODB.Connection, _
    ByVal strField As String, _
    ByVal strTable As String, _
    ByVal strCriteria As String, _
    ByVal strCompare As String, _
    Optional ByVal blnCaseSensitive As Boolean = False) As Boolean
    Dim rstSQL As ADODB.Recordset
    
    Set rstSQL = mdlDatabase.OpenRecordset(conSQL, strField, strTable, False, strCriteria)
    
    If rstSQL.RecordCount > 0 Then
        If blnCaseSensitive Then
            If Trim(rstSQL.Fields(strField).Value) = strCompare Then
                IsDataCorrect = True
            Else
                IsDataCorrect = False
            End If
        Else
            If UCase(Trim(rstSQL.Fields(strField).Value)) = UCase(strCompare) Then
                IsDataCorrect = True
            Else
                IsDataCorrect = False
            End If
        End If
    Else
        IsDataCorrect = False
    End If
    
    mdlDatabase.CloseRecordset rstSQL
End Function

Public Function BackupDatabase( _
    ByRef conSQL As ADODB.Connection, _
    ByVal strDatabase As String, _
    ByVal strFileName As String, _
    Optional objDatabase As SQLDATABASE = SQLSERVER7) As Boolean
    On Local Error GoTo ErrHandler
    
    If mdlGlobal.fso.FileExists(strFileName) Then mdlGlobal.fso.DeleteFile strFileName, True
    
    If objDatabase = SQLSERVER7 Or objDatabase = SQLSERVER2000 Then
        Dim strQuery As String
        
        strQuery = "BACKUP DATABASE " & strDatabase
        strQuery = strQuery & " TO DISK='" & strFileName & "'"
    
        conSQL.Execute strQuery
    ElseIf objDatabase = MSACCESS Then
        mdlGlobal.fso.CopyFile mdlGlobal.strPath & strDatabase & ".mdb", strFileName, True
    End If
    
    BackupDatabase = True
    
    Exit Function
    
ErrHandler:
    BackupDatabase = False
End Function

Public Function GetColumnSize(ByRef conSQL As ADODB.Connection, ByVal strField As String, ByVal strTable As String) As Integer
    Dim rstSQL As ADODB.Recordset
    
    Set rstSQL = mdlDatabase.OpenRecordset(conSQL, strField, strTable)
    
    GetColumnSize = rstSQL.Fields(strField).DefinedSize
    
    mdlDatabase.CloseRecordset rstSQL
End Function

Public Sub DeleteSingleRecord(ByRef rstSQL As ADODB.Recordset)
    If rstSQL.RecordCount > 0 Then
        rstSQL.Delete
        
        rstSQL.MoveNext
        
        If rstSQL.EOF Then rstSQL.MovePrevious
    End If
End Sub

Public Sub DeleteRecordQuery(ByRef conSQL As ADODB.Connection, ByVal strTable As String, ByVal strCriteria As String)
    conSQL.Execute "DELETE FROM " & strTable & " WHERE " & strCriteria
End Sub

Public Sub TruncateTable( _
    ByRef conSQL As ADODB.Connection, _
    ByVal strTable As String, _
    Optional objDatabase As SQLDATABASE = SQLSERVER7)
    If objDatabase = SQLSERVER7 Or _
        objDatabase = SQLSERVER2000 Or _
        objDatabase = SQLEXPRESS Then
        conSQL.Execute "TRUNCATE TABLE " & strTable
    ElseIf objDatabase = MSACCESS Then
        conSQL.Execute "DROP TABLE " & strTable
        
        mdlTable.CreateTHRECYCLE
    End If
End Sub

Public Sub CloseConnection(ByRef conSQL As ADODB.Connection)
    If conSQL Is Nothing Then Exit Sub
    
    conSQL.Close
    
    Set conSQL = Nothing
End Sub

Public Sub CloseRecordset(ByRef rstSQL As ADODB.Recordset)
    If rstSQL Is Nothing Then Exit Sub
    
    If rstSQL.State = ObjectStateEnum.adStateOpen Then rstSQL.Close
    
    Set rstSQL = Nothing
End Sub
