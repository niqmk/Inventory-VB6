VERSION 5.00
Begin VB.Form frmPrinter 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1470
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6735
   Icon            =   "frmPrinter.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1470
   ScaleWidth      =   6735
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cmbPrinter 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   480
      Width           =   6495
   End
   Begin VB.CommandButton cmdPrinter 
      Caption         =   "Set Default"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5280
      TabIndex        =   1
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label lblPrinter 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Setting Default Printer"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1905
   End
End
Attribute VB_Name = "frmPrinter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    SetInitialization
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmPrinter = Nothing
End Sub

Private Sub cmdPrinter_Click()
    If Not mdlProcedures.IsValidComboData(Me.cmbPrinter) Then Exit Sub
    
    Dim strTemp As String
    strTemp = Me.lblPrinter.Caption
    
    Me.lblPrinter.Caption = "Silahkan Tunggu"
    
    DoEvents
    
    mdlPrinter.SetPrinterText mdlProcedures.GetComboData(Me.cmbPrinter)
    
    Me.lblPrinter.Caption = strTemp
End Sub

Private Sub SetInitialization()
    mdlProcedures.CenterWindows Me
    
    Me.Caption = mdlText.strPrinter
    
    mdlPrinter.FillComboPrinter Me.cmbPrinter
End Sub
