Attribute VB_Name = "mdlPrinter"
Option Explicit

Private Const HWND_BROADCAST = &HFFFF
Private Const WM_WININICHANGE = &H1A

Private Const CCHDEVICENAME = 32
Private Const CCHFORMNAME = 32

Private Const STANDARD_RIGHTS_REQUIRED = &HF0000
Private Const PRINTER_ACCESS_ADMINISTER = &H4
Private Const PRINTER_ACCESS_USE = &H8
Private Const PRINTER_ALL_ACCESS = (STANDARD_RIGHTS_REQUIRED Or PRINTER_ACCESS_ADMINISTER Or PRINTER_ACCESS_USE)

Private Const PRINTER_ATTRIBUTE_DEFAULT = 4

Private Const VER_PLATFORM_WIN32_WINDOWS As Long = 1

Public Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type

Public Type DEVMODE
     dmDeviceName As String * CCHDEVICENAME
     dmSpecVersion As Integer
     dmDriverVersion As Integer
     dmSize As Integer
     dmDriverExtra As Integer
     dmFields As Long
     dmOrientation As Integer
     dmPaperSize As Integer
     dmPaperLength As Integer
     dmPaperWidth As Integer
     dmScale As Integer
     dmCopies As Integer
     dmDefaultSource As Integer
     dmPrintQuality As Integer
     dmColor As Integer
     dmDuplex As Integer
     dmYResolution As Integer
     dmTTOption As Integer
     dmCollate As Integer
     dmFormName As String * CCHFORMNAME
     dmLogPixels As Integer
     dmBitsPerPel As Long
     dmPelsWidth As Long
     dmPelsHeight As Long
     dmDisplayFlags As Long
     dmDisplayFrequency As Long
     dmICMMethod As Long
     dmICMIntent As Long
     dmMediaType As Long
     dmDitherType As Long
     dmReserved1 As Long
     dmReserved2 As Long
End Type

Public Type PRINTER_INFO_5
     pPrinterName As String
     pPortName As String
     Attributes As Long
     DeviceNotSelectedTimeout As Long
     TransmissionRetryTimeout As Long
End Type

Public Type PRINTER_DEFAULTS
     pDatatype As Long
     pDevMode As Long
     DesiredAccess As Long
End Type

Private Declare Function GetProfileString Lib "kernel32" Alias "GetProfileStringA" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long) As Long

Private Declare Function WriteProfileString Lib "kernel32" Alias "WriteProfileStringA" (ByVal lpszSection As String, ByVal lpszKeyName As String, ByVal lpszString As String) As Long

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lparam As String) As Long

Private Declare Function GetVersionExA Lib "kernel32" (lpVersionInformation As OSVERSIONINFO) As Integer

Private Declare Function OpenPrinter Lib "winspool.drv" Alias "OpenPrinterA" (ByVal pPrinterName As String, phPrinter As Long, pDefault As PRINTER_DEFAULTS) As Long

Private Declare Function SetPrinter Lib "winspool.drv" Alias "SetPrinterA" (ByVal hPrinter As Long, ByVal Level As Long, pPrinter As Any, ByVal Command As Long) As Long

Private Declare Function GetPrinter Lib "winspool.drv" Alias "GetPrinterA" (ByVal hPrinter As Long, ByVal Level As Long, pPrinter As Any, ByVal cbBuf As Long, pcbNeeded As Long) As Long

Private Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As String, ByVal lpString2 As Any) As Long

Private Declare Function ClosePrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long

Public Sub FillComboPrinter(ByRef cmbData As ComboBox)
    Dim lngResult As Long
    Dim strBuffer As String
    
    strBuffer = Space(8192)
    lngResult = GetProfileString("PrinterPorts", vbNullString, "", strBuffer, Len(strBuffer))
    
    cmbData.Clear
    
    Dim intPosition As Integer
    Dim strParse As String
    
    Do
        intPosition = InStr(strBuffer, Chr(0))
        
        If intPosition > 0 Then
            strParse = Left(strBuffer, intPosition - 1)
            
            If Len(Trim(strParse)) Then cmbData.AddItem strParse
            
            strBuffer = Mid(strBuffer, intPosition + 1)
        Else
            If Len(Trim(strBuffer)) Then cmbData.AddItem strBuffer
            
            strBuffer = ""
        End If
    Loop While intPosition > 0
    
    mdlProcedures.SetComboData cmbData, Printer.DeviceName
End Sub

Public Sub SetPrinterText(ByVal strPrinterText As String)
    Dim typOSInfo As mdlPrinter.OSVERSIONINFO
    Dim intRetValue As Integer
    
    typOSInfo.dwOSVersionInfoSize = 148
    typOSInfo.szCSDVersion = Space$(128)
    intRetValue = GetVersionExA(typOSInfo)

    If typOSInfo.dwPlatformId = VER_PLATFORM_WIN32_WINDOWS Then
        Win95SetDefaultPrinter strPrinterText
    Else
        WinNTSetDefaultPrinter strPrinterText
    End If
End Sub

Private Sub Win95SetDefaultPrinter(ByVal strPrinterText As String)
    Dim typPD As PRINTER_DEFAULTS
    Dim typPI5 As PRINTER_INFO_5
    
    Dim lngHandle As Long
    Dim lngResult As Long
    Dim lngNeed As Long
    
    Dim LastError As Long
    
    If Trim(strPrinterText) = "" Then
        Exit Sub
    End If
    
    typPD.pDatatype = 0&
    typPD.DesiredAccess = PRINTER_ALL_ACCESS Or typPD.DesiredAccess

    lngResult = OpenPrinter(strPrinterText, lngHandle, typPD)
    
    If lngResult = 0 Then
        Exit Sub
    End If
    
    lngResult = GetPrinter(lngHandle, 5, ByVal 0&, 0, lngNeed)
    
    Dim lngTemp() As Long
    
    ReDim lngTemp((lngNeed \ 4)) As Long
    
    lngResult = GetPrinter(lngHandle, 5, lngTemp(0), lngNeed, lngNeed)
    
    If lngResult = 0 Then
        Exit Sub
    End If
    
    typPI5.pPrinterName = PtrCtoVbString(lngTemp(0))
    typPI5.pPortName = PtrCtoVbString(lngTemp(1))
    typPI5.Attributes = lngTemp(2)
    typPI5.DeviceNotSelectedTimeout = lngTemp(3)
    typPI5.TransmissionRetryTimeout = lngTemp(4)
    typPI5.Attributes = PRINTER_ATTRIBUTE_DEFAULT
    
    lngResult = SetPrinter(lngHandle, 5, typPI5, 0)
    
    If lngResult = 0 Then
        MsgBox "SetPrinter Failed. Error code: " & Err.LastDllError
        
        Exit Sub
    Else
        If Not Printer.DeviceName = strPrinterText Then
            SelectPrinter strPrinterText
        End If
    End If
    
    ClosePrinter lngHandle
End Sub

Public Sub SelectPrinter(ByVal strPrinterText As String)
    Dim objPrinter As Printer
    
    For Each objPrinter In Printers
        If objPrinter.DeviceName = strPrinterText Then
            Set Printer = objPrinter
            
            Exit For
        End If
    Next
End Sub

Private Sub WinNTSetDefaultPrinter(ByVal strPrinterText As String)
    Dim lngResult As Long
    
    Dim strBuffer As String
    Dim strDeviceName As String
    Dim strDriverName As String
    Dim strPrinterPort As String
    
    If Trim(strPrinterText) = "" Then Exit Sub
    
    strBuffer = Space(1024)
    
    lngResult = GetProfileString("PrinterPorts", strPrinterText, "", strBuffer, Len(strBuffer))
    
    GetDriverAndPort strBuffer, strDriverName, strPrinterPort
    
    If (Not Trim(strDriverName) = "") And (Not strPrinterPort = "") Then
        SetDefaultPrinter strPrinterText, strDriverName, strPrinterPort
        
        If Not Printer.DeviceName = strPrinterText Then
            SelectPrinter strPrinterText
        End If
    End If
End Sub

Private Function PtrCtoVbString(ByVal lngAdd As Long) As String
    Dim lngResult As Long
    
    Dim strTemp As String * 512

    lngResult = lstrcpy(strTemp, lngAdd)
    
    If (InStr(1, strTemp, Chr(0)) = 0) Then
         PtrCtoVbString = ""
    Else
         PtrCtoVbString = Left(strTemp, InStr(1, strTemp, Chr(0)) - 1)
    End If
End Function

Private Sub SetDefaultPrinter(ByVal strPrinterName As String, ByVal strDriverName As String, ByVal strPrinterPort As String)
    Dim lngResult As Long
    Dim lngSend As Long
    
    Dim strDeviceLine As String
    
    strDeviceLine = strPrinterName & "," & strDriverName & "," & strPrinterPort
    
    lngResult = WriteProfileString("windows", "Device", strDeviceLine)
    
    lngSend = SendMessage(HWND_BROADCAST, WM_WININICHANGE, 0, "windows")
End Sub

Private Sub GetDriverAndPort(ByVal strBuffer As String, ByRef strDriverName As String, ByRef strPrinterPort As String)
    Dim intDriver As Integer
    Dim intPort As Integer
    
    strDriverName = ""
    strPrinterPort = ""
    
    intDriver = InStr(strBuffer, ",")
    
    If intDriver > 0 Then
        strDriverName = Left(strBuffer, intDriver - 1)
        
        intPort = InStr(intDriver + 1, strBuffer, ",")
        
        If intPort > 0 Then
            strPrinterPort = Mid(strBuffer, intDriver + 1, intPort - intDriver - 1)
        End If
    End If
End Sub
