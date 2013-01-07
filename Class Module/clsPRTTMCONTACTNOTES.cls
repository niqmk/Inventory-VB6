VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPRTTMCONTACTNOTES"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const NAME_COL As Integer = 1
Private Const DATE_COL As Integer = 2
Private Const TIME_COL As Integer = 3
Private Const NOTES_COL As Integer = 4

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
        .Columns("A").ColumnWidth = 28
        .Columns("B").ColumnWidth = 18
        .Columns("C").ColumnWidth = 10
        .Columns("D").ColumnWidth = 25
        
        .Rows(intRow).RowHeight = 125
    End With
End Sub

Private Sub FillTitle()
    On Local Error GoTo ErrHandler
    
    With xlSheet
        .Shapes.AddPicture mdlGlobal.strLogoImageText, True, True, 0, 0, 100, 100
        
        intRow = intRow + 1
        
        mdlExcel.FillEdge xlSheet, "A" & intRow & ":D" & intRow + 1
        
        .Range("A" & intRow & ":D" & intRow).Font.Bold = True
        
        .Cells(intRow, NAME_COL).Value = "Nama"
        .Cells(intRow, DATE_COL).Value = "Tanggal"
        .Cells(intRow, TIME_COL).Value = "Jam"
        .Cells(intRow, NOTES_COL).Value = "Keterangan"
    End With
    
ErrHandler:
End Sub

Private Sub FillText(ByRef rstMain As ADODB.Recordset)
    On Local Error GoTo ErrHandler
    
    With rstMain
        Dim intFirst As Integer
        Dim intTemp As Integer
        
        Dim mContactId As String

        intFirst = intRow + 1
        intTemp = 0
        
        While Not .EOF
            intRow = intRow + 1
            
            xlSheet.Range("A" & intRow & ":D" & intRow).VerticalAlignment = xlTop
            xlSheet.Range("A" & intRow & ":D" & intRow).HorizontalAlignment = xlLeft
            xlSheet.Range("A" & intRow & ":D" & intRow).WrapText = True
            
            If Trim(mContactId) = Trim(.Fields("ContactId").Value) Then
                xlSheet.Range("A" & intTemp & ":A" & intRow).MergeCells = True
            Else
                If Not intTemp = 0 Then
                    mdlExcel.FillEdge xlSheet, "A" & intTemp & ":D" & intRow
                End If
                
                mContactId = Trim(.Fields("ContactId").Value)
                
                xlSheet.Cells(intRow, NAME_COL).Value = "'" & Trim(mdlDatabase.GetFieldData(mdlGlobal.conInventory, "Name", mdlTable.CreateTMCONTACTCUSTOMER, "ContactId='" & mContactId & "'"))
                
                intTemp = intRow
            End If
            
            xlSheet.Cells(intRow, DATE_COL).Value = "'" & mdlProcedures.FormatDate(.Fields("NotesDate").Value, "dd-MMMM-yyyy")
            xlSheet.Cells(intRow, TIME_COL).Value = "'" & mdlProcedures.FormatDate(.Fields("NotesDate").Value, "hh:mm:ss")
            xlSheet.Cells(intRow, NOTES_COL).Value = "'" & Trim(.Fields("Notes").Value)
            
            xlSheet.Range("C" & intRow & ":D" & intRow).Select
            
            .MoveNext
        Wend
        
        If .RecordCount > 0 Then
            mdlExcel.FillEdge xlSheet, "A" & intTemp & ":D" & intRow, False
            
            xlSheet.Range("A" & intFirst & ":D" & intFirst).Select
        End If
    End With
    
ErrHandler:
End Sub

Private Sub Class_Terminate()
    mdlExcel.CloseExcel
End Sub
