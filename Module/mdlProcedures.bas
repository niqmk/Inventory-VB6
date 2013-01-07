Attribute VB_Name = "mdlProcedures"
Option Explicit

Public Sub ShowForm(ByRef frmShow As Form, Optional ByVal blnCloseAll As Boolean = True, Optional blnModal As Boolean = False, Optional ByVal strException As String = "")
    If blnCloseAll Then
        mdlProcedures.CloseAllForms frmShow, strException
    End If
    
    If blnModal Then
        frmShow.Show 1
    Else
        frmShow.Show
    End If
End Sub

Public Sub CloseAllForms(ByRef frmShow As Form, Optional ByVal strException As String = "")
    If mdlGlobal.blnFill Then Exit Sub
    
    On Local Error GoTo ErrHandler
    
    Dim intCounter As Integer
        
    For intCounter = 1 To Forms.Count - 1
        If Not Forms(intCounter).Name = strException Then
            Unload Forms(intCounter)
        End If
    Next intCounter
        
ErrHandler:
End Sub

Public Sub CenterWindows(ByRef frmCenter As Form, Optional ByVal blnTop As Boolean = True, Optional ByVal blnLeft As Boolean = False)
    If blnLeft Then
        frmCenter.Move 0, (Screen.Height - frmCenter.ScaleHeight) / 2
    Else
        frmCenter.Move (Screen.Width - frmCenter.ScaleWidth) / 2, (Screen.Height - frmCenter.ScaleHeight) / 2
    End If
    
    If blnTop Then frmCenter.Move frmCenter.Left, 0
End Sub

Public Sub CornerWindows(ByRef frmCorner As Form, Optional ByVal blnLeft As Boolean = True, Optional ByVal blnTop As Boolean = True)
    If blnLeft Then
        If blnTop Then
            frmCorner.Move 0, 0
        Else
            frmCorner.Move 0, (Screen.Height - frmCorner.Height)
        End If
    Else
        If blnTop Then
            frmCorner.Move (Screen.Width - frmCorner.Width) - 70, 0
        Else
            frmCorner.Move (Screen.Width - frmCorner.Width) - 70, (Screen.Height - frmCorner.Height)
        End If
    End If
End Sub

Public Sub GotFocus(ByRef txtFocus As TextBox)
    With txtFocus
        txtFocus.SelStart = 0
        txtFocus.SelLength = Len(txtFocus)
    End With
End Sub

Public Sub FillComboMonth( _
    ByRef cmbMonth As ComboBox, _
    Optional ByVal intStartMonth As Integer = 1, _
    Optional ByVal intFinishMonth As Integer = 12, _
    Optional ByVal intMonth As Integer = 1)
    cmbMonth.Clear
    
    Dim intCounter As Integer
    
    For intCounter = intStartMonth To intFinishMonth
        cmbMonth.AddItem Format(intCounter, "00")
        
        If intCounter = intMonth Then cmbMonth.ListIndex = intCounter - intStartMonth
    Next intCounter
End Sub

Public Sub FillComboData( _
    ByRef cmbData As ComboBox, _
    ByVal rstSQL As ADODB.Recordset, _
    Optional ByVal strSeparator As String = " | ", _
    Optional ByVal blnCheckDateFields As Boolean = False)
    cmbData.Clear
    
    If rstSQL.RecordCount > 0 Then rstSQL.MoveFirst
    
    Dim intCounter As Integer
    
    Dim strValue As String
    Dim strTemp As String
    
    While Not rstSQL.EOF
        strValue = ""
        
        For intCounter = 0 To rstSQL.Fields.Count - 1
            If intCounter > 0 Then strValue = strValue & strSeparator
            
            If blnCheckDateFields Then
                If IsDate(rstSQL.Fields(intCounter).Value) Then
                    strValue = strValue & mdlProcedures.FormatDate(rstSQL.Fields(intCounter).Value, mdlGlobal.strFormatDate)
                Else
                    strTemp = rstSQL.Fields(intCounter).Value
                    strTemp = strTemp & _
                        Space(rstSQL.Fields(intCounter).DefinedSize - Len(strTemp))
                    
                    strValue = strValue & strTemp
                End If
            Else
                strTemp = rstSQL.Fields(intCounter).Value
                strTemp = strTemp & _
                    Space(rstSQL.Fields(intCounter).DefinedSize - Len(strTemp))
                
                strValue = strValue & strTemp
            End If
        Next intCounter
        
        cmbData.AddItem strValue
        
        rstSQL.MoveNext
    Wend
End Sub

Public Sub SetComboData( _
    ByRef cmbData As ComboBox, _
    ByVal strSearch As String, _
    Optional ByVal intData As Integer = 0, _
    Optional ByVal strSeparator As String = " | ")
    Dim blnFound As Boolean
    
    blnFound = False
    
    Dim intCounter As Integer
    
    Dim strValue() As String
    
    For intCounter = 0 To cmbData.ListCount - 1
        strValue = Split(cmbData.List(intCounter), strSeparator)
        
        If UCase(Trim(strValue(intData))) = UCase(Trim(strSearch)) Then
            blnFound = True
            
            Exit For
        End If
    Next intCounter
    
    If blnFound Then
        cmbData.ListIndex = intCounter
    Else
        cmbData.ListIndex = -1
    End If
End Sub

Public Function GetComboData( _
    ByRef cmbData As ComboBox, _
    Optional ByVal intIndex As Integer = -1, _
    Optional ByVal intPosition As Integer = 0, _
    Optional ByVal strSeparator As String = " | ") As String
    Dim strData() As String
    
    If intIndex = -1 Then
        If cmbData.ListCount > 0 Then
            strData = Split(cmbData.List(cmbData.ListIndex), strSeparator)
            
            If intPosition > UBound(strData) Then
                GetComboData = ""
            Else
                GetComboData = strData(intPosition)
            End If
        Else
            GetComboData = ""
        End If
    Else
        If intIndex > cmbData.ListCount - 1 Then
            GetComboData = ""
        Else
            strData = Split(cmbData.List(cmbData.ListIndex), strSeparator)
            
            If intPosition > UBound(strData) Then
                GetComboData = ""
            Else
                GetComboData = strData(intPosition)
            End If
        End If
    End If
End Function

Public Function SplitData( _
    ByVal strValue As String, _
    Optional ByVal intPosition As Integer = 0, _
    Optional ByVal strSeparator As String = " | ") As String
    Dim strData() As String
    
    strData = Split(strValue, strSeparator)
    
    If intPosition > UBound(strData) Then
        SplitData = ""
    Else
        SplitData = strData(intPosition)
    End If
End Function

Public Function IsValidComboData(ByRef cmbData As ComboBox) As Boolean
    If cmbData.ListCount > 0 Then
        If cmbData.ListIndex > -1 Then
            IsValidComboData = True
        Else
            IsValidComboData = False
        End If
    Else
        IsValidComboData = False
    End If
End Function

Public Function RepDupText(ByVal strText As String) As String
    strText = Replace(strText, "'", "")
    strText = Replace(strText, "%", "")
    strText = Replace(strText, "|", "")
    strText = Replace(strText, "$", "")
    
    RepDupText = strText
End Function

Public Function RepRegistryUnknown(ByVal strText As String) As String
    Dim intPosition As String
    
    intPosition = InStr(strText, Chr(0))
    
    If Not intPosition = 0 Then
        strText = Mid(strText, 1, intPosition - 1)
    Else
        strText = ""
    End If
    
    RepRegistryUnknown = Replace(strText, " | ", vbCrLf)
End Function

Public Sub SetControlMode( _
    ByVal frmControl As Form, _
    Optional ByVal objMode As FunctionMode = ViewMode, _
    Optional ByVal blnClear As Boolean = True, _
    Optional ByVal strPrimary As String = "", _
    Optional ByVal strSearch As String = "", _
    Optional ByVal strSeparator As String = " | ")
    Dim cntl As Control
    
    Dim strSearchControl() As String
    
    strSearchControl = Split(strSearch, strSeparator)
    
    Dim intCounter As Integer
    
    Dim blnSearch As Boolean

    For Each cntl In frmControl.Controls
        If (TypeOf cntl Is TextBox) Then
            blnSearch = False
            
            If objMode = ViewMode Then
                For intCounter = 0 To UBound(strSearchControl)
                    If cntl.Name = strSearchControl(intCounter) Then
                        cntl.Appearance = cc3D
                        
                        blnSearch = True
                        
                        Exit For
                    End If
                Next intCounter
            Else
                For intCounter = 0 To UBound(strSearchControl)
                    If cntl.Name = strSearchControl(intCounter) Then
                        cntl.Appearance = ccFlat
                        
                        blnSearch = True
                        
                        Exit For
                    End If
                Next intCounter
            End If
        
            If Not blnSearch Then
                If Not cntl.Name = strPrimary Then
                    If objMode = ViewMode Then
                        cntl.Appearance = ccFlat
                    Else
                        If cntl.Name = strPrimary Then
                            cntl.Appearance = ccFlat
                        Else
                            cntl.Appearance = cc3D
                        End If
                    End If
                    
                    If blnClear Then
                        If Not cntl.Name = strPrimary Then
                            cntl.Text = ""
                        End If
                    End If
                End If
            End If
        ElseIf (TypeOf cntl Is DTPicker) Then
            If blnClear Then
                If (mdlProcedures.IsDateDiff(mdlProcedures.FormatDate(cntl.MinDate, mdlGlobal.strFormatDate), mdlProcedures.FormatDate(Now, mdlGlobal.strFormatDate)) And _
                    mdlProcedures.IsDateDiff(mdlProcedures.FormatDate(Now, mdlGlobal.strFormatDate), mdlProcedures.FormatDate(cntl.MaxDate, mdlGlobal.strFormatDate))) Then
                    cntl.Value = mdlProcedures.FormatDate(Now, mdlGlobal.strFormatDate) & " " & mdlProcedures.FormatDate(Now, "hh:mm:ss")
                Else
                    cntl.Value = cntl.MinDate & " " & mdlProcedures.FormatDate(Now, "hh:mm:ss")
                End If
            End If
        ElseIf (TypeOf cntl Is ComboBox) Then
            If blnClear Then
                blnSearch = False

                For intCounter = 0 To UBound(strSearchControl)
                    If cntl.Name = strSearchControl(intCounter) Then
                        blnSearch = True
                        
                        Exit For
                    End If
                Next intCounter

                If Not blnSearch Then cntl.ListIndex = -1
            End If
        ElseIf (TypeOf cntl Is CheckBox) Then
            If blnClear Then
                cntl.Value = vbUnchecked
            End If
        End If
    Next cntl
End Sub

Public Function SetCaptionMode( _
    ByVal strText As String, _
    Optional ByVal objMode As FunctionMode = ViewMode) As String
    Dim strCaption As String
    
    strCaption = ""
    
    Select Case objMode
        Case ViewMode:
            strCaption = " (MODE LIHAT)"
        Case AddMode:
            strCaption = " (MODE TAMBAH)"
        Case UpdateMode:
            strCaption = " (MODE UBAH)"
        Case DeleteMode:
            strCaption = " (MODE HAPUS)"
        Case PrintMode:
            strCaption = " (MODE CETAK)"
    End Select
    
    SetCaptionMode = strText & strCaption
End Function

Public Function SetMsgYesNo( _
    Optional ByVal strMessage As String = "", _
    Optional ByVal strTitle As String = "") As Boolean
    If MsgBox(strMessage, vbYesNo + vbQuestion, strTitle) = vbYes Then
        SetMsgYesNo = True
    Else
        SetMsgYesNo = False
    End If
End Function

Public Function FormatDate(ByVal objDate As Date, Optional ByVal strFormat As String = "yyyy/MM/dd")
    FormatDate = Format(objDate, strFormat)
End Function

Public Function GetAscii(ByVal strValue As String) As Integer
    If Len(strValue) > 1 Then
        GetAscii = 0
    Else
        GetAscii = Asc(strValue)
    End If
End Function

Public Function GetCurrency(ByVal strValue As String) As Currency
    If IsNumeric(strValue) Then
        GetCurrency = CCur(strValue)
    Else
        GetCurrency = 0
    End If
End Function

Public Function GetDouble(ByVal strValue As String) As Double
    If IsNumeric(strValue) Then
        GetDouble = CDbl(strValue)
    Else
        GetDouble = 0#
    End If
End Function

Public Function GetNumber( _
    ByVal strValue As String, _
    Optional intLimit As Integer = 32767, _
    Optional ByVal blnLimitNegative As Boolean = True) As Integer
    If IsNumeric(strValue) Then
        Dim intLimitNegative As Integer
        
        If blnLimitNegative Then
            intLimitNegative = -(intLimit)
        Else
            intLimitNegative = 0
        End If
        
        If strValue >= intLimitNegative And strValue <= intLimit Then
            GetNumber = CInt(strValue)
        Else
            GetNumber = 0
        End If
    Else
        GetNumber = 0
    End If
End Function

Public Function FormatCurrency(ByVal strValue As String, Optional ByVal strFormat As String = "#,##0") As String
    If IsNumeric(strValue) Then
        FormatCurrency = Format(CCur(strValue), strFormat)
    Else
        FormatCurrency = "0"
    End If
End Function

Public Function FormatNumber(ByVal intValue As Integer, Optional ByVal strFormat As String = "000") As String
    FormatNumber = Format(intValue, strFormat)
End Function

Public Function GetMaxDate(ByVal intMonth As Integer, ByVal intYear As Integer) As Integer
    Dim blnMonthFull As Boolean
    
    blnMonthFull = False

    If (intYear Mod 4) = 0 Then blnMonthFull = True
    
    Select Case intMonth
        Case 1:
            GetMaxDate = 31
        Case 2:
            If blnMonthFull Then
                GetMaxDate = 29
            Else
                GetMaxDate = 28
            End If
        Case 3:
            GetMaxDate = 31
        Case 4:
            GetMaxDate = 30
        Case 5:
            GetMaxDate = 31
        Case 6:
            GetMaxDate = 30
        Case 7:
            GetMaxDate = 31
        Case 8:
            GetMaxDate = 31
        Case 9:
            GetMaxDate = 30
        Case 10:
            GetMaxDate = 31
        Case 11:
            GetMaxDate = 30
        Case 12:
            GetMaxDate = 31
        Case Else:
            GetMaxDate = 0
    End Select
End Function

Public Function IsDateDiff(ByVal dteFrom As Date, ByVal dteTo As Date) As Boolean
    Dim intDifferent As Integer
    
    Dim intYearFrom As Integer
    Dim intYearTo As Integer
    Dim intMonthFrom As Integer
    Dim intMonthTo As Integer
    
    intYearFrom = CInt(mdlProcedures.FormatDate(dteFrom, "yyyy"))
    intYearTo = CInt(mdlProcedures.FormatDate(dteTo, "yyyy"))
    intMonthFrom = CInt(mdlProcedures.FormatDate(dteFrom, "MM"))
    intMonthTo = CInt(mdlProcedures.FormatDate(dteTo, "MM"))
    
    If intYearFrom > intYearTo Then
        intDifferent = -1
    ElseIf intYearFrom < intYearTo Then
        intDifferent = 0
    ElseIf (intYearFrom = intYearTo) And (intMonthFrom > intMonthTo) Then
        intDifferent = -1
    Else
        intDifferent = DateDiff("d", dteFrom, dteTo)
    End If
    
    If intDifferent >= 0 Then
        IsDateDiff = True
    Else
        IsDateDiff = False
    End If
End Function

Public Function IsDataExistsInFlex( _
    ByRef flxGrid As MSFlexGrid, _
    ByVal strValue As String, _
    Optional ByVal lngColumn As Long = 1, _
    Optional ByVal lngStartRow As Long = 1, _
    Optional ByVal lngFinishRow As Long = -1, _
    Optional ByVal blnUpperCase As Boolean = True, _
    Optional ByVal blnFocus As Boolean = False) As Boolean
    Dim blnFound As Boolean
    
    Dim lngCounter As Long
    
    blnFound = False
    
    With flxGrid
        Dim lngFinish As Long
        
        If lngFinishRow = -1 Then
            lngFinish = .Rows - 1
        Else
            lngFinish = lngFinishRow
        End If
        
        For lngCounter = lngStartRow To lngFinish
            If blnUpperCase Then
                If UCase(Trim(.TextMatrix(lngCounter, lngColumn))) = UCase(Trim(strValue)) Then
                    blnFound = True
                    
                    If blnFocus Then .Row = lngCounter
                    
                    Exit For
                End If
            Else
                If Trim(.TextMatrix(lngCounter, lngColumn)) = Trim(strValue) Then
                    blnFound = True
                    
                    If blnFocus Then .Row = lngCounter
                    
                    Exit For
                End If
            End If
        Next lngCounter
    End With
    
    IsDataExistsInFlex = blnFound
End Function

Public Function IsDataExistsInListView( _
    ByRef lsvView As ListView, _
    ByVal strValue As String, _
    Optional ByVal intStartRow As Integer = 1, _
    Optional ByVal blnUpperCase As Boolean = True, _
    Optional ByVal blnFocus As Boolean = True) As Boolean
    Dim blnFound As Boolean
    
    Dim intCounter As Long
    
    blnFound = False
    
    With lsvView
        For intCounter = intStartRow To .ListItems.Count
            If blnUpperCase Then
                If UCase(Trim(.ListItems(intCounter).Text)) = UCase(Trim(strValue)) Then
                    blnFound = True
                    
                    If blnFocus Then
                        .ListItems(intCounter).Selected = True
                        .SetFocus
                    End If
                    
                    Exit For
                End If
            Else
                If Trim(.ListItems(intCounter).Text) = Trim(strValue) Then
                    blnFound = True
                    
                    If blnFocus Then
                        .ListItems(intCounter).Selected = True
                        
                        .SetFocus
                    End If
                    
                    Exit For
                End If
            End If
        Next intCounter
    End With
    
    IsDataExistsInListView = blnFound
End Function

Public Function DateAddFormat( _
    ByVal objCompareDate As Date, _
    ByVal objDate As Date, _
    ByVal strInterval As String, _
    ByVal dblValue As Double, _
    Optional ByVal blnTillNow As Boolean = True) As Date
    
    If blnTillNow Then
        While mdlProcedures.FormatDate(objCompareDate) >= mdlProcedures.FormatDate(objDate)
            objDate = DateAdd(strInterval, dblValue, objDate)
        Wend
        
        While mdlProcedures.FormatDate(Now) >= mdlProcedures.FormatDate(objDate)
            objDate = DateAdd(strInterval, dblValue, objDate)
        Wend
    Else
        objDate = DateAdd(strInterval, dblValue, objDate)
    End If
    
    DateAddFormat = objDate
End Function

Public Function IsValidDateRegion() As Boolean
    Dim dteValue As Date
    
    dteValue = "01/31/1990"
    
    Dim strValue As String
    
    strValue = CStr(dteValue)
    
    If Left(strValue, 2) = "31" Then
        IsValidDateRegion = False
    Else
        IsValidDateRegion = True
    End If
End Function

Public Function SetDate( _
    ByVal strMonth As String, _
    ByVal strYear As String, _
    Optional ByVal strDate As String = "01", _
    Optional ByVal blnMaxDate As Boolean = False) As Date
    If blnMaxDate Then
        strDate = mdlProcedures.GetMaxDate(CInt(strMonth), CInt(strYear))
    End If
    
    Dim dteValue As Date
    
    If mdlProcedures.IsValidDateRegion Then
        dteValue = _
            strMonth & _
            "/" & _
            strDate & _
            "/" & _
            strYear
    Else
        dteValue = _
            strDate & _
            "/" & _
            strMonth & _
            "/" & _
            strYear
    End If
    
    SetDate = dteValue
End Function

Public Function QueryLikeCriteria(ByVal strField As String, ByVal strValue As String) As String
    Dim strCriteria As String
    
    strCriteria = "(" & strField & " LIKE '" & strValue & "%'"
    strCriteria = strCriteria & " OR " & strField & " LIKE '%" & strValue & "%'"
    strCriteria = strCriteria & " OR " & strField & " LIKE '%" & strValue & "')"
    
    QueryLikeCriteria = strCriteria
End Function
