VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPRTTMEMPLOYEE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const EMPLOYEEID_COL As Integer = 1
Private Const DATE_COL As Integer = 2
Private Const NAME_COL As Integer = 3
Private Const JOBTYPENAME_COL As Integer = 4
Private Const ADDRESS_COL As Integer = 5
Private Const PHONE_COL As Integer = 6
Private Const HANDPHONE_COL As Integer = 7
Private Const FAX_COL As Integer = 8
Private Const EMAIL_COL As Integer = 9

Private xlSheet As Excel.Worksheet

Private intRow As Integer

Private Sub Class_Initialize()
    mdlExcel.OpenExcel
End Sub

Public Sub ImportToExcel(ByRef rstMain As ADODB.Recordset)
    If rstMain Is Nothing Then Exit Sub
    
    mdlExcel.OpenWorkSheet xlSheet
    
    SetWidth
    
    FillTitle
    
    FillText rstMain
End Sub

Private Sub SetWidth()
    intRow = 1
    
    With xlSheet
        .PageSetup.TopMargin = 20
        .PageSetup.LeftMargin = 20
        .PageSetup.RightMargin = 20
        .PageSetup.Orientation = xlLandscape
        
        .Columns("A").ColumnWidth = 8
        .Columns("B").ColumnWidth = 18
        .Columns("C").ColumnWidth = 20
        .Columns("D").ColumnWidth = 18
        .Columns("E").ColumnWidth = 18
        .Columns("F").ColumnWidth = 18
        .Columns("G").ColumnWidth = 10
        .Columns("H").ColumnWidth = 10
        .Columns("I").ColumnWidth = 18
        
        .Rows(intRow).RowHeight = 125
    End With
End Sub

Private Sub FillTitle()
    On Local Error GoTo ErrHandler
    
    With xlSheet
        .Shapes.AddPicture mdlGlobal.strLogoImageText, True, True, 0, 0, 100, 100
        
        intRow = intRow + 1
        
        mdlExcel.FillEdge xlSheet, "A" & intRow & ":I" & intRow + 1
        
        .Range("A" & intRow & ":I" & intRow).Font.Bold = True
        
        .Cells(intRow, EMPLOYEEID_COL).Value = "Kode"
        .Cells(intRow, DATE_COL).Value = "Tanggal"
        .Cells(intRow, NAME_COL).Value = "Nama"
        .Cells(intRow, JOBTYPENAME_COL).Value = "Jabatan"
        .Cells(intRow, ADDRESS_COL).Value = "Alamat"
        .Cells(intRow, PHONE_COL).Value = "Telepon"
        .Cells(intRow, HANDPHONE_COL).Value = "Handphone"
        .Cells(intRow, FAX_COL).Value = "Fax"
        .Cells(intRow, EMAIL_COL).Value = "Email"
    End With
    
ErrHandler:
End Sub

Private Sub FillText(ByRef rstMain As ADODB.Recordset)
    On Local Error GoTo ErrHandler
    
    With rstMain
        Dim intTemp As Integer

        intTemp = intRow + 1
    
        While Not .EOF
            intRow = intRow + 1
            
            xlSheet.Range("A" & intRow & ":I" & intRow).VerticalAlignment = xlTop
            xlSheet.Range("A" & intRow & ":I" & intRow).HorizontalAlignment = xlLeft
            xlSheet.Range("A" & intRow & ":I" & intRow).WrapText = True
            
            xlSheet.Cells(intRow, EMPLOYEEID_COL).Value = "'" & Trim(.Fields("EmployeeId").Value)
            xlSheet.Cells(intRow, DATE_COL).Value = "'" & mdlProcedures.FormatDate(.Fields("EmployeeDate").Value, "dd-MMMM-yyyy")
            xlSheet.Cells(intRow, NAME_COL).Value = "'" & Trim(.Fields("Name").Value)
            xlSheet.Cells(intRow, ADDRESS_COL).Value = "'" & Trim(.Fields("Address").Value)
            xlSheet.Cells(intRow, PHONE_COL).Value = "'" & Trim(.Fields("Phone").Value)
            xlSheet.Cells(intRow, HANDPHONE_COL).Value = "'" & Trim(.Fields("HandPhone").Value)
            xlSheet.Cells(intRow, FAX_COL).Value = "'" & Trim(.Fields("Fax").Value)
            xlSheet.Cells(intRow, EMAIL_COL).Value = "'" & Trim(.Fields("Email").Value)
            
            xlSheet.Range("A" & intRow & ":I" & intRow).Select
            
            .MoveNext
        Wend
        
        If .RecordCount > 0 Then
            mdlExcel.FillEdge xlSheet, "A" & intTemp & ":I" & intRow, False
            
            xlSheet.Range("A" & intTemp & ":I" & intTemp).Select
        End If
    End With
    
ErrHandler:
End Sub

Private Sub Class_Terminate()
    mdlExcel.CloseExcel
End Sub
