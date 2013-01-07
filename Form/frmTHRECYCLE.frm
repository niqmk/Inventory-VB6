VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTHRECYCLE 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6495
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   8295
   Icon            =   "frmTHRECYCLE.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6495
   ScaleWidth      =   8295
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.Toolbar tlbMain 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   1164
      ButtonWidth     =   609
      ButtonHeight    =   1005
      Appearance      =   1
      _Version        =   393216
      Begin MSComctlLib.ImageList imlMain 
         Left            =   7680
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   3
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTHRECYCLE.frx":1F8A
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTHRECYCLE.frx":3ADC
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTHRECYCLE.frx":562E
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.ListView lsvMain 
      Height          =   5295
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   9340
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin MSComctlLib.ImageList imlRecycle 
      Left            =   120
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTHRECYCLE.frx":7180
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTHRECYCLE.frx":8CD2
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmTHRECYCLE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum ButtonMode
    [DeleteAllButton] = 1
    [RestoreButton]
    [RefreshButton]
End Enum

Private Sub Form_Load()
    SetInitialization
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmTHRECYCLE = Nothing
End Sub

Private Sub tlbMain_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case DeleteAllButton:
            If mdlProcedures.SetMsgYesNo("Anda Ingin Hapus Semua ?", Me.Caption) Then
                mdlDatabase.TruncateTable mdlGlobal.conInventory, mdlTable.CreateTHRECYCLE, mdlGlobal.objDatabaseInit
                
                SetRecordset
            End If
        Case RestoreButton:
            Dim strIdInit As String
            
            strIdInit = Me.lsvMain.ListItems(Me.lsvMain.SelectedItem.Index).ToolTipText
            
            If InStr(strIdInit, "/") > 0 Then
                strIdInit = Left(strIdInit, InStr(InStr(strIdInit, "/") + 1, strIdInit, "/") - 1)
            End If
            
            Dim blnValid As Boolean
            
            blnValid = False
            
            Select Case strIdInit
                Case mdlText.strPOIDINIT & "/" & mdlText.strSELLINIT:
                    blnValid = mdlTHPOSELL.RestoreTHPOSELL(Me.lsvMain.ListItems(Me.lsvMain.SelectedItem.Index).ToolTipText)
                    
                    If Not blnValid Then
                        blnValid = mdlTHSALESSUM.RestoreTHSALESSUM(Me.lsvMain.ListItems(Me.lsvMain.SelectedItem.Index).ToolTipText)
                    End If
                Case mdlText.strSOIDINIT & "/" & mdlText.strSELLINIT:
                    blnValid = mdlTHSOSELL.RestoreTHSOSELL(Me.lsvMain.ListItems(Me.lsvMain.SelectedItem.Index).ToolTipText)
                Case mdlText.strSJIDINIT & "/" & mdlText.strSELLINIT:
                    blnValid = mdlTHSJSELL.RestoreTHSJSELL(Me.lsvMain.ListItems(Me.lsvMain.SelectedItem.Index).ToolTipText)
                Case mdlText.strFKTIDINIT & "/" & mdlText.strSELLINIT:
                    blnValid = mdlTHFKTSELL.RestoreTHFKTSELL(Me.lsvMain.ListItems(Me.lsvMain.SelectedItem.Index).ToolTipText)
                Case mdlText.strRTRIDINIT & "/" & mdlText.strSELLINIT:
                    blnValid = mdlTHRTRSELL.RestoreTHRTRSELL(Me.lsvMain.ListItems(Me.lsvMain.SelectedItem.Index).ToolTipText)
                Case mdlText.strDOIDINIT & "/" & mdlText.strBUYINIT:
                    blnValid = mdlTHDOBUY.RestoreTHDOBUY(Me.lsvMain.ListItems(Me.lsvMain.SelectedItem.Index).ToolTipText)
                Case mdlText.strSJIDINIT & "/" & mdlText.strBUYINIT:
                    blnValid = mdlTHSJBUY.RestoreTHSJBUY(Me.lsvMain.ListItems(Me.lsvMain.SelectedItem.Index).ToolTipText)
                Case mdlText.strFKTIDINIT & "/" & mdlText.strBUYINIT:
                    blnValid = mdlTHFKTBUY.RestoreTHFKTBUY(Me.lsvMain.ListItems(Me.lsvMain.SelectedItem.Index).ToolTipText)
                Case mdlText.strRTRIDINIT & "/" & mdlText.strBUYINIT:
                    blnValid = mdlTHRTRBUY.RestoreTHRTRBUY(Me.lsvMain.ListItems(Me.lsvMain.SelectedItem.Index).ToolTipText)
                Case Else:
                    If InStr(strIdInit, "/") > 0 Then
                        If UCase(Left(strIdInit, Len(mdlText.strITEMOUTIDINIT))) = UCase(mdlText.strITEMOUTIDINIT) Then
                            blnValid = mdlTHITEMOUT.RestoreTHITEMOUT(Me.lsvMain.ListItems(Me.lsvMain.SelectedItem.Index).ToolTipText)
                        ElseIf UCase(Left(strIdInit, Len(mdlText.strITEMINIDINIT))) = UCase(mdlText.strITEMINIDINIT) Then
                            blnValid = mdlTHITEMIN.RestoreTHITEMIN(Me.lsvMain.ListItems(Me.lsvMain.SelectedItem.Index).ToolTipText)
                        End If
                    Else
                        blnValid = mdlTHMUTITEM.RestoreTHMUTITEM(Me.lsvMain.ListItems(Me.lsvMain.SelectedItem.Index).ToolTipText)
                    End If
            End Select
            
            If blnValid Then
                mdlDatabase.DeleteRecordQuery mdlGlobal.conInventory, mdlTable.CreateTHRECYCLE, "RecycleId='" & Me.lsvMain.ListItems(Me.lsvMain.SelectedItem.Index).ToolTipText & "'"
                mdlDatabase.DeleteRecordQuery mdlGlobal.conInventory, mdlTable.CreateTDRECYCLE, "RecycleId='" & Me.lsvMain.ListItems(Me.lsvMain.SelectedItem.Index).ToolTipText & "'"
                
                SetRecordset
                
                frmMenu.SetRecycle
            Else
                MsgBox "Konflik Data", vbOKOnly + vbCritical, Me.Caption
            End If
        Case RefreshButton:
            SetRecordset
    End Select
    
    frmMenu.SetRecycle
End Sub

Private Sub lsvMain_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        If mdlProcedures.SetMsgYesNo("Anda Ingin Hapus " & Me.lsvMain.ListItems(Me.lsvMain.SelectedItem.Index).Text & " ?", Me.Caption) Then
            mdlDatabase.DeleteRecordQuery mdlGlobal.conInventory, mdlTable.CreateTHRECYCLE, "ReferencesNumber='" & Me.lsvMain.ListItems(Me.lsvMain.SelectedItem.Index).Text & "'"
            
            Me.lsvMain.ListItems.Remove Me.lsvMain.SelectedItem.Index
            
            frmMenu.SetRecycle
        End If
    End If
End Sub

Private Sub SetInitialization()
    mdlProcedures.CenterWindows Me
    
    Me.Caption = mdlText.strTHRECYCLE
    
    With Me.tlbMain
        .AllowCustomize = False
        .ImageList = Me.imlMain
        
        .Buttons.Add DeleteAllButton, , "Hapus Semua", , DeleteAllButton
        .Buttons.Add RestoreButton, , "Restore", , RestoreButton
        .Buttons.Add RefreshButton, , "Refresh", , RefreshButton
    End With
    
    With Me.lsvMain
        .Icons = Me.imlRecycle
    End With
    
    SetRecordset
End Sub

Private Sub SetRecordset()
    Me.lsvMain.ListItems.Clear
    
    Dim rstTemp As ADODB.Recordset
    
    Set rstTemp = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "*", mdlTable.CreateTHRECYCLE, False)
    
    With rstTemp
        If .RecordCount > 0 Then
            Me.tlbMain.Buttons(DeleteAllButton).Enabled = True
            Me.tlbMain.Buttons(RestoreButton).Enabled = True
        Else
            Me.tlbMain.Buttons(DeleteAllButton).Enabled = False
            Me.tlbMain.Buttons(RestoreButton).Enabled = False
        End If
        
        While Not .EOF
            If InStr(Trim(!RecycleId), "/") > 0 Then
                Me.lsvMain.ListItems.Add , , Trim(!ReferencesNumber), 1
            Else
                Me.lsvMain.ListItems.Add , , Trim(!ReferencesNumber), 2
            End If
            
            Me.lsvMain.ListItems(Me.lsvMain.ListItems.Count).ToolTipText = Trim(!RecycleId)
            
            .MoveNext
        Wend
    End With
    
    mdlDatabase.CloseRecordset rstTemp
End Sub
