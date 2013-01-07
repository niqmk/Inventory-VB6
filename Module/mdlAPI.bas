Attribute VB_Name = "mdlAPI"
Option Explicit

Public Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    
Public Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
    
Public Declare Function SetLayeredWindowAttributes Lib "user32.dll" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
    
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Long) As Long

Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Declare Function CreateRoundRectRgn Lib "gdi32.dll" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
        
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Declare Function Beep Lib "kernel32" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long

Private Const GWL_EXSTYLE As Integer = (-20)

Private Const WS_EX_LAYERED As Long = &H80000

Public Sub SmoothForm(frmSmooth As Form, Optional ByVal dblCurvature As Double = 25)
    Dim lngRgn As Long
    
    Dim X1 As Long
    Dim Y1 As Long
    
    X1 = frmSmooth.Width / Screen.TwipsPerPixelX
    Y1 = frmSmooth.Height / Screen.TwipsPerPixelY
    
    lngRgn = CreateRoundRectRgn(0, 0, X1, Y1, dblCurvature, dblCurvature)
    
    SetWindowRgn frmSmooth.hwnd, lngRgn, True
    DeleteObject lngRgn
End Sub

Public Sub SetLayeredWindow(ByVal hwnd As Long, ByVal bIsLayered As Boolean)
    Dim lngWinInfo As Long

    lngWinInfo = mdlAPI.GetWindowLong(hwnd, GWL_EXSTYLE)
    
    lngWinInfo = lngWinInfo Or WS_EX_LAYERED
    
    SetWindowLong hwnd, GWL_EXSTYLE, lngWinInfo
End Sub
