VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmWallpaper 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7230
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8670
   Icon            =   "frmWallpaper.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7230
   ScaleWidth      =   8670
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7200
      TabIndex        =   2
      Top             =   6600
      Width           =   1335
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "Gambar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5520
      TabIndex        =   1
      Top             =   6600
      Width           =   1575
   End
   Begin VB.PictureBox picWallpaper 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   6375
      Left            =   120
      ScaleHeight     =   6345
      ScaleWidth      =   8385
      TabIndex        =   0
      Top             =   120
      Width           =   8415
   End
   Begin MSComDlg.CommonDialog cdlPicture 
      Left            =   120
      Top             =   6600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmWallpaper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private strWallpaperFileName As String

Private Sub Form_Load()
    SetInitialization
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmWallpaper = Nothing
End Sub

Private Sub cmdBrowse_Click()
    On Local Error GoTo ErrHandler
    
    With Me.cdlPicture
        .ShowOpen
        
        If Not Trim(.FileName) = "" Then
            strWallpaperFileName = Trim(.FileName)
            
            Set Me.picWallpaper.Picture = LoadPicture(strWallpaperFileName)
        End If
    End With
    
ErrHandler:
End Sub

Private Sub cmdOK_Click()
    If Not Trim(strWallpaperFileName) = "" Then
        If mdlGlobal.fso.FileExists(strWallpaperFileName) Then
            'Dim strExtension() As String
            
            'strExtension = Split(strWallpaperFileName, ".")
            
            'mdlGlobal.strWallpaperImageText = mdlGlobal.strPath & mdlGlobal.WALLPAPER_IMAGE_FILE & "." & strExtension(UBound(strExtension))
            
            mdlGlobal.strWallpaperImageText = Trim(strWallpaperFileName)
            
            'mdlGlobal.fso.CopyFile strWallpaperFileName, mdlGlobal.strWallpaperImageText, True
            
            SetRegistry
            
            Set mdiMain.Picture = LoadPicture(mdlGlobal.strWallpaperImageText)
        End If
    End If
    
    Unload Me
End Sub

Private Sub SetInitialization()
    On Local Error GoTo ErrHandler
    
    mdlProcedures.CenterWindows Me, False
    
    Me.Caption = mdlText.strWall
    
    With Me.cdlPicture
        .CancelError = True
        .Filter = "All Images |*.JPG;*.BMP"
    End With
    
    If mdlGlobal.fso.FileExists(mdlGlobal.strWallpaperImageText) Then
        Me.picWallpaper.Picture = LoadPicture(mdlGlobal.strWallpaperImageText)
    End If
    
ErrHandler:
End Sub

Private Sub SetRegistry()
    Dim lngRegistry As Long
    Dim lngRegKey As Long
    Dim lngType As Long
    Dim lngSize As Long
    
    lngRegistry = _
        mdlRegistry.OpenRegistry(mdlRegistry.HKEY_CURRENT_USER, mdlRegistry.KEYS_SYS_INFO, lngRegKey)
        
    If Not lngRegistry = 0 Then
        lngRegistry = mdlRegistry.WriteToRegistry(mdlRegistry.HKEY_CURRENT_USER, mdlRegistry.KEYS_SYS_INFO)
    End If
    
    lngRegistry = _
            mdlRegistry.WriteValueRegistry(mdlRegistry.HKEY_CURRENT_USER, mdlRegistry.KEYS_SYS_INFO, mdlGlobal.WALLPAPER_REGISTRY, mdlGlobal.strWallpaperImageText)

    lngRegistry = mdlRegistry.CloseRegistry(lngRegKey)
End Sub
