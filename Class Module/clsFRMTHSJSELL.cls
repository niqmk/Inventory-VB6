VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFRMTHSJSELL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const COMPANY_COL As Integer = 2
Private Const CUSTOMER_COL As Integer = 5

Private Const NOSEQ_COL As Integer = 1
Private Const QTY_COL As Integer = 2
Private Const NAME_COL As Integer = 3

Private xlSheet As Excel.Worksheet

Private intRow As Integer

Private Sub Class_Initialize()
    If mdlGlobal.fso.FileExists(mdlGlobal.strPath & mdlGlobal.strTemplateFolder & "\" & mdlTable.CreateTHSJSELL & ".xls") Then
        mdlExcel.OpenExcel mdlGlobal.strPath & mdlGlobal.strTemplateFolder & "\" & mdlTable.CreateTHSJSELL & ".xls"
    Else
        mdlExcel.OpenExcel
    End If
End Sub

Public Sub ImportToExcel(ByRef rstMain As ADODB.Recordset)
    If rstMain Is Nothing Then Exit Sub
    
    mdlExcel.OpenWorkSheet xlSheet, False

    FillTitle rstMain

    FillDetail rstMain!SJId
End Sub

Private Sub FillTitle(ByRef rstMain As ADODB.Recordset)
    On Local Error GoTo ErrHandler

    With xlSheet
        .Range("A1:E28").HorizontalAlignment = xlHAlignLeft
        .Range("A1:E28").VerticalAlignment = xlVAlignTop
        .Range("A1:E28").Font.Bold = True
        
        Dim strPOId As String
        Dim strCustomerId As String
        Dim strDeliveryAddress As String
        
        strPOId = mdlDatabase.GetFieldData(mdlGlobal.conInventory, "POId", mdlTable.CreateTHSOSELL, "SOId='" & rstMain!SOId & "'")
        strCustomerId = mdlDatabase.GetFieldData(mdlGlobal.conInventory, "CustomerId", mdlTable.CreateTHPOSELL, "POId='" & strPOId & "'")
        strDeliveryAddress = mdlDatabase.GetFieldData(mdlGlobal.conInventory, "Address", mdlTable.CreateTMDELIVERYCUSTOMER, "DeliveryId='" & rstMain!DeliveryId & "'")
        
        If Trim(strDeliveryAddress) = "" Then
            strDeliveryAddress = mdlDatabase.GetFieldData(mdlGlobal.conInventory, "Address", mdlTable.CreateTMCUSTOMER, "CustomerId='" & rstMain!DeliveryId & "'")
        End If
        
        Dim strAddress() As String
        
        strAddress = Split(strDeliveryAddress, vbCrLf)
        
        .Cells(1, CUSTOMER_COL).Value = "'" & mdlDatabase.GetFieldData(mdlGlobal.conInventory, "Name", mdlTable.CreateTMCUSTOMER, "CustomerId='" & strCustomerId & "'")
        .Cells(2, CUSTOMER_COL).Value = "'" & strAddress(0)
        
        If UBound(strAddress) > 0 Then
            .Cells(3, CUSTOMER_COL).Value = "'" & strAddress(1)
        End If
        
        .Cells(5, CUSTOMER_COL).Value = "P.O. No. : " & mdlDatabase.GetFieldData(mdlGlobal.conInventory, "POCustomerId", mdlTable.CreateTHPOSELL, "POId='" & strPOId & "'")
        .Cells(8, CUSTOMER_COL).Value = "'" & rstMain!ReferencesNumber
        
        strAddress = Split(mdlGlobal.strAddressText, vbCrLf)
        
        .Cells(2, COMPANY_COL).Value = "'" & strAddress(0)
        
        If UBound(strAddress) > 0 Then
            .Cells(3, COMPANY_COL).Value = "'" & strAddress(1)
        End If
        
        .Cells(4, COMPANY_COL).Value = "Telp : " & mdlGlobal.strPhoneText
        .Cells(5, COMPANY_COL).Value = "Fax  : " & mdlGlobal.strFaxText
        .Cells(8, COMPANY_COL).NumberFormat = "dd-MM-yyyy"
        .Cells(8, COMPANY_COL).Value = mdlProcedures.FormatDate(rstMain!SJDate, "dd-MM-yyyy")
    End With

ErrHandler:
End Sub

Private Sub FillDetail(Optional ByVal strSJId As String = "")
    On Local Error GoTo ErrHandler

    Dim rstDetail As ADODB.Recordset

    Set rstDetail = mdlDatabase.OpenRecordset(mdlGlobal.conInventory, "*", mdlTable.CreateTDSJSELL, False, "SJId='" & strSJId & "'", "ItemId ASC")

    With rstDetail
        Dim intNoSeq As Integer
        Dim intTemp As Integer

        Dim strItemId As String

        If .RecordCount > 0 Then
            intRow = 11
            intNoSeq = 0
            intTemp = intRow + 1
        End If

        While Not .EOF
            intNoSeq = intNoSeq + 1
            intRow = intRow + 1
            
            xlSheet.Range("B" & intRow & ":B" & intRow).HorizontalAlignment = xlHAlignCenter
            
            strItemId = Trim(.Fields("ItemId").Value)
            
            xlSheet.Cells(intRow, NOSEQ_COL).Value = CStr(intNoSeq)
            xlSheet.Cells(intRow, QTY_COL).Value = "'" & mdlProcedures.FormatCurrency(.Fields("Qty").Value) & " " & Trim(mdlDatabase.GetFieldData(mdlGlobal.conInventory, "UnityId", mdlTable.CreateTMITEM, "ItemId='" & strItemId & "'"))
            xlSheet.Cells(intRow, NAME_COL).Value = "'" & Trim(mdlDatabase.GetFieldData(mdlGlobal.conInventory, "Name", mdlTable.CreateTMITEM, "ItemId='" & strItemId & "'"))

            xlSheet.Range("A" & intTemp & ":E" & intTemp).Select

            .MoveNext
        Wend

        If .RecordCount > 0 Then
            xlSheet.Range("A" & intTemp & ":E" & intTemp).Select
        End If
    End With
    
    xlSheet.Cells(28, 4).Value = "Mengetahui :"

ErrHandler:
End Sub

Private Sub Class_Terminate()
    mdlExcel.CloseExcel
End Sub
