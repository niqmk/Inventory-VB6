Attribute VB_Name = "mdlMath"
Option Explicit

Public Function Shl(ByVal bytValue As Byte, ByVal intValue As Integer) As Integer
    Dim strBinary As String
    Dim strTemp As String

    strBinary = ByteToBinary(bytValue)
    strTemp = ""
    
    While Not Len(strTemp) = intValue
        strTemp = strTemp & "0"
    Wend
    
    strBinary = Mid(strBinary, intValue) & strTemp
    
    Shl = BinaryToInteger(strBinary)
End Function

Public Function Shr(ByVal bytValue As Byte, ByVal intValue As Integer) As Integer
    Dim strBinary As String
    Dim strTemp As String
    
    strBinary = ByteToBinary(bytValue)
    strTemp = ""
    
    While Not Len(strTemp) = intValue
        strTemp = strTemp & "0"
    Wend
    
    strBinary = strTemp & Mid(strBinary, 1, Len(strBinary) - Len(strTemp) - 1)
    
    Shr = BinaryToInteger(strBinary)
End Function

Public Function StringToByte(ByVal strValue As String) As Byte()
    Dim bytResult() As Byte
    
    ReDim bytResult(Len(strValue) - 1) As Byte
    
    Dim intCounter As Integer
    
    For intCounter = 1 To Len(strValue)
        bytResult(intCounter - 1) = Conversion.CByte(Asc(Mid(strValue, intCounter, 1)))
    Next intCounter

    StringToByte = bytResult
End Function

Public Function ByteToString(ByRef bytValue() As Byte) As String
    Dim strResult As String
    
    strResult = ""
    
    Dim intCounter As Integer
    
    For intCounter = 0 To UBound(bytValue)
        strResult = strResult & Chr(bytValue(intCounter))
    Next intCounter
    
    ByteToString = strResult
End Function

Public Function ByteToBinary(ByVal bytValue As Byte) As String
    Dim intTemp As Integer

    Dim strResult As String
    
    strResult = ""
    
    Do
        intTemp = bytValue Mod 2
        
        strResult = Conversion.CStr(intTemp) & strResult
        
        bytValue = Conversion.CByte(bytValue \ 2)
    Loop Until bytValue = 0
    
    While Len(strResult) < 32
        strResult = "0" & strResult
    Wend
    
    ByteToBinary = strResult
End Function

Public Function BinaryToInteger(ByVal strBinary As String) As Integer
    Dim intTarget As Integer
    
    intTarget = 0
    
    Dim intCounter As Integer
    
    For intCounter = Len(strBinary) To 1 Step -1
        intTarget = _
            intTarget + _
            EachBinaryToInteger(Mid(strBinary, intCounter, 1), Len(strBinary) - intCounter)
    Next intCounter
    
    BinaryToInteger = intTarget
End Function

Private Function EachBinaryToInteger(ByVal strEachBinary As String, ByVal intCounter As Integer) As Integer
    Select Case strEachBinary
        Case "0"
            EachBinaryToInteger = 0
        Case "1"
            EachBinaryToInteger = CInt(2 ^ intCounter)
    End Select
End Function
