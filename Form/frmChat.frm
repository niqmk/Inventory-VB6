VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmChat 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6015
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   7365
   Icon            =   "frmChat.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6015
   ScaleWidth      =   7365
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ImageList imlMain 
      Left            =   1920
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":1F8A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":3ADC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSWinsockLib.Winsock wskChat 
      Left            =   120
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox rtbChat 
      Height          =   5295
      Left            =   2640
      TabIndex        =   3
      Top             =   120
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   9340
      _Version        =   393217
      ScrollBars      =   2
      TextRTF         =   $"frmChat.frx":562E
   End
   Begin VB.CommandButton cmdChat 
      Caption         =   "Kirim"
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
      Left            =   6120
      TabIndex        =   1
      Top             =   5520
      Width           =   1095
   End
   Begin VB.TextBox txtChat 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   5640
      Width           =   5895
   End
   Begin MSComctlLib.ListView lsvMain 
      Height          =   5415
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   9551
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
End
Attribute VB_Name = "frmChat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    SetInitialization
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    mdlGlobal.blnChat = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmChat = Nothing
End Sub

Private Sub lsvMain_DblClick()
    If Me.lsvMain.ListItems.Count > 0 Then
        mdlProcedures.CornerWindows Me
        
        mdlProcedures.ShowForm frmChatDetail, False, True
    End If
End Sub

Private Sub wskChat_Connect()
    Me.wskChat.SendData mdlGlobal.UserAuthority.UserId & " | " & Me.txtChat.Text
End Sub

Private Sub wskChat_SendComplete()
    Me.wskChat.Close
End Sub

Private Sub cmdChat_Click()
    SendChatText
End Sub

Private Sub SetInitialization()
    mdlProcedures.CenterWindows Me
    
    Me.Caption = mdlText.strChat
    
    With Me.lsvMain
        .FullRowSelect = True
        .LabelEdit = lvwManual
        .View = lvwReport
        .SmallIcons = Me.imlMain
        
        .ColumnHeaders.Add , , "List Pemakai"
        .ColumnHeaders.Add , , "IP"
        
        .ColumnHeaders(1).Width = 2000
    End With
    
    Me.rtbChat.Locked = True
    
    mdlGlobal.blnChat = True
    
    FillUserChat
End Sub

Private Sub SendChatText()
    Dim rstTemp As ADODB.Recordset
    
    Set rstTemp = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "UserId, LogYN, UserIP", mdlTable.CreateTMUSERLOGIN, False)
    
    With rstTemp
        While Not .EOF
            If Not Trim(!UserId) = mdlGlobal.UserAuthority.UserId Then
                If Trim(!LogYN) = mdlGlobal.strYes Then
                    SetConnection Trim(!UserIP)
                End If
            End If
            
            .MoveNext
        Wend
    End With
    
    mdlDatabase.CloseRecordset rstTemp
    
    Me.ReceiveChatText mdlGlobal.UserAuthority.UserId, Me.txtChat.Text
End Sub

Private Sub SetConnection(ByVal strIP As String)
    If Not Me.wskChat.State = sckClosed Then
        Me.wskChat.Close
    End If
    
    Me.wskChat.Connect strIP, mdlGlobal.intPort
End Sub

Private Sub FillUserChat()
    Dim lstItem As ListItem
    
    Dim rstTemp As ADODB.Recordset
    
    Set rstTemp = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "UserId, LogYN, UserIP", mdlTable.CreateTMUSERLOGIN, False)
    
    With rstTemp
        While Not .EOF
            If Not Trim(!UserId) = mdlGlobal.UserAuthority.UserId Then
                If Trim(!LogYN) = mdlGlobal.strYes Then
                    Set lstItem = Me.lsvMain.ListItems.Add(, , Trim(!UserId), , 1)
                Else
                    Set lstItem = Me.lsvMain.ListItems.Add(, , Trim(!UserId), , 2)
                End If
                
                lstItem.ListSubItems.Add , , Trim(!UserIP)
            End If
            
            .MoveNext
        Wend
    End With
    
    mdlDatabase.CloseRecordset rstTemp
End Sub

Public Sub ReceiveChatText(ByVal strUserId As String, ByVal strText As String)
    Me.rtbChat.Text = Me.rtbChat.Text & strUserId & " : " & strText & vbCrLf
End Sub

Public Property Get UserIP() As String
    If Me.lsvMain.ListItems.Count > 0 Then
        If Me.lsvMain.SelectedItem.Selected Then
            UserIP = Me.lsvMain.SelectedItem.ListSubItems(1).Text
        End If
    End If
End Property
