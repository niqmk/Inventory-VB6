Attribute VB_Name = "mdlExcel"
Option Explicit

Private xlApp As Excel.Application
Private xlBook As Excel.Workbook

Public Sub OpenExcel(Optional ByVal strFileName As String = "", Optional ByVal blnVisible As Boolean = True)
    On Local Error GoTo ErrHandler
    
    Set xlApp = New Excel.Application
    
    If Trim(strFileName) = "" Then
        Set xlBook = xlApp.Workbooks.Add
    Else
        Set xlBook = xlApp.Workbooks.Open(strFileName)
    End If
    
    xlApp.Visible = blnVisible
    
ErrHandler:
End Sub

Public Sub CloseExcel()
    On Local Error GoTo ErrHandler
    
    If xlApp Is Nothing Then Exit Sub
    
    xlBook.Close
    
ErrHandler:
    Set xlBook = Nothing
    
    Set xlApp = Nothing
End Sub

Public Sub OpenWorkSheet(ByRef xlSheet As Excel.Worksheet, Optional ByVal blnAddSheet As Boolean = True)
    On Local Error GoTo ErrHandler
    
    If xlBook Is Nothing Then Exit Sub
    
    mdlExcel.CloseWorkSheet xlSheet
    
    If blnAddSheet Then
        Set xlSheet = xlBook.Worksheets.Add
    Else
        Set xlSheet = xlBook.Worksheets(1)
    End If
    
ErrHandler:
End Sub

Public Sub CloseWorkSheet(ByRef xlSheet As Excel.Worksheet)
    Set xlSheet = Nothing
End Sub

Public Sub FillEdge( _
    ByRef xlSheet As Excel.Worksheet, _
    ByVal strRange As String, _
    Optional ByVal blnTop As Boolean = True, _
    Optional ByVal blnBottom As Boolean = True, _
    Optional ByVal blnInsideHorizontal As Boolean = True, _
    Optional ByVal blnInsideVertical As Boolean = True, _
    Optional ByVal blnLeft As Boolean = True, _
    Optional ByVal blnRight As Boolean = True, _
    Optional ByVal lneStyle As XlLineStyle = XlLineStyle.xlContinuous)
    If blnTop Then xlSheet.Range(strRange).Borders(xlEdgeTop).LineStyle = lneStyle
    If blnBottom Then xlSheet.Range(strRange).Borders(xlEdgeBottom).LineStyle = lneStyle
    If blnInsideHorizontal Then xlSheet.Range(strRange).Borders(xlInsideHorizontal).LineStyle = lneStyle
    If blnInsideVertical Then xlSheet.Range(strRange).Borders(xlInsideVertical).LineStyle = lneStyle
    If blnLeft Then xlSheet.Range(strRange).Borders(xlEdgeLeft).LineStyle = lneStyle
    If blnRight Then xlSheet.Range(strRange).Borders(xlEdgeRight).LineStyle = lneStyle
End Sub

Public Function GetNumberColumn(ByVal strAlphabet As String) As Integer
    Dim strAlpha(1 To 26) As String
    
    strAlpha(1) = "A"
    strAlpha(2) = "B"
    strAlpha(3) = "C"
    strAlpha(4) = "D"
    strAlpha(5) = "E"
    strAlpha(6) = "F"
    strAlpha(7) = "G"
    strAlpha(8) = "H"
    strAlpha(9) = "I"
    strAlpha(10) = "J"
    strAlpha(11) = "K"
    strAlpha(12) = "L"
    strAlpha(13) = "M"
    strAlpha(14) = "N"
    strAlpha(15) = "O"
    strAlpha(16) = "P"
    strAlpha(17) = "Q"
    strAlpha(18) = "R"
    strAlpha(19) = "S"
    strAlpha(20) = "T"
    strAlpha(21) = "U"
    strAlpha(22) = "V"
    strAlpha(23) = "W"
    strAlpha(24) = "X"
    strAlpha(25) = "Y"
    strAlpha(26) = "Z"
    
    GetNumberColumn = 0
    
    Dim intCounter As Integer
    
    For intCounter = 1 To 26
        If strAlpha(intCounter) = UCase(strAlphabet) Then
            GetNumberColumn = intCounter
            
            Exit For
        End If
    Next intCounter
End Function

Public Function GetAlphabetColumn(ByVal intNumber As Integer) As String
    Dim intMod As Integer
    Dim intDivision As Integer
    
    intMod = intNumber Mod 26
    intDivision = intNumber \ 26
    
    If intMod = 0 Then
        GetAlphabetColumn = GetAlphabet(intMod)
    Else
        GetAlphabetColumn = GetAlphabet(intMod) & GetAlphabet(intDivision, True)
    End If
End Function

Private Function GetAlphabet(ByVal intNumber As Integer, Optional ByVal blnAlphabetNull As Boolean = False) As String
    If blnAlphabetNull Then
        If intNumber = 0 Then
            GetAlphabet = ""
            
            Exit Function
        End If
    End If
    
    GetAlphabet = Mid("ZABCDEFGHIJKLMNOPQRSTUVWXYZ", intNumber + 1, 1)
End Function
