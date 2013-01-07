VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPRTTMITEM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const ITEMID_COL As Integer = 1
Private Const PARTNUMBER_COL As Integer = 2
Private Const NAME_COL As Integer = 3
Private Const GROUPNAME_COL As Integer = 4
Private Const CATEGORYNAME_COL As Integer = 5
Private Const BRANDNAME_COL As Integer = 6
Private Const MINSTOCK_COL As Integer = 7
Private Const MAXSTOCK_COL As Integer = 8
Private Const UNITYNAME_COL As Integer = 9

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
        .Columns("B").ColumnWidth = 12
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
        
        .Cells(intRow, ITEMID_COL).Value = "Kode"
        .Cells(intRow, PARTNUMBER_COL).Value = "Nomor Part"
        .Cells(intRow, NAME_COL).Value = "Nama"
        .Cells(intRow, GROUPNAME_COL).Value = "Grup"
        .Cells(intRow, CATEGORYNAME_COL).Value = "Jenis"
        .Cells(intRow, BRANDNAME_COL).Value = "Merk"
        .Cells(intRow, MINSTOCK_COL).Value = "Min. Stok"
        .Cells(intRow, MAXSTOCK_COL).Value = "Max. Stok"
        .Cells(intRow, UNITYNAME_COL).Value = "Satuan"
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
            xlSheet.Range("A" & intRow & ":F" & intRow).HorizontalAlignment = xlLeft
            xlSheet.Range("G" & intRow & ":H" & intRow).HorizontalAlignment = xlRight
            xlSheet.Range("I" & intRow & ":I" & intRow).HorizontalAlignment = xlLeft
            xlSheet.Range("A" & intRow & ":I" & intRow).WrapText = True
            
            xlSheet.Cells(intRow, ITEMID_COL).Value = "'" & Trim(.Fields("ItemId").Value)
            xlSheet.Cells(intRow, PARTNUMBER_COL).Value = "'" & Trim(.Fields("PartNumber").Value)
            xlSheet.Cells(intRow, NAME_COL).Value = "'" & Trim(.Fields("Name").Value)
            xlSheet.Cells(intRow, GROUPNAME_COL).Value = "'" & Trim(mdlDatabase.GetFieldData(mdlGlobal.conInventory, "Name", mdlTable.CreateTMGROUP, "GroupId='" & .Fields("GroupId").Value & "'"))
            xlSheet.Cells(intRow, CATEGORYNAME_COL).Value = "'" & Trim(mdlDatabase.GetFieldData(mdlGlobal.conInventory, "Name", mdlTable.CreateTMCATEGORY, "CategoryId='" & .Fields("CategoryId").Value & "'"))
            xlSheet.Cells(intRow, BRANDNAME_COL).Value = "'" & Trim(mdlDatabase.GetFieldData(mdlGlobal.conInventory, "Name", mdlTable.CreateTMBRAND, "BrandId='" & .Fields("BrandId").Value & "'"))
            xlSheet.Cells(intRow, MINSTOCK_COL).Value = rstMain.Fields("MinStock").Value
            xlSheet.Cells(intRow, MINSTOCK_COL).NumberFormat = "#,##0"
            xlSheet.Cells(intRow, MAXSTOCK_COL).Value = rstMain.Fields("MaxStock").Value
            xlSheet.Cells(intRow, MAXSTOCK_COL).NumberFormat = "#,##0"
            xlSheet.Cells(intRow, UNITYNAME_COL).Value = "'" & Trim(mdlDatabase.GetFieldData(mdlGlobal.conInventory, "Name", mdlTable.CreateTMUNITY, "UnityId='" & .Fields("UnityId").Value & "'"))
            
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
