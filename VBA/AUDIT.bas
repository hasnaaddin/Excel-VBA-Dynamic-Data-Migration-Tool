Attribute VB_Name = "AUDIT"
'--- Updated Validation Sub with sums and sumValidationStatus
Sub AppendValidationWithSum(SourceWSName As String, targetWSName As String, SourceColName As String, SourceCount As Long, SourceSum As Variant, targetColName As String, TargetCount As Long, TargetSum As Variant, CountStatus As String, SumStatus As String)
    Dim wsValidation As Worksheet
    Dim rowOut As Long
    On Error Resume Next
    Set wsValidation = ThisWorkbook.Sheets("Validation")
    On Error GoTo 0
    If wsValidation Is Nothing Then
        Set wsValidation = ThisWorkbook.Sheets.Add
        wsValidation.Name = "Validation"
        wsValidation.Range("A2:J2").Value = Array("SourceWS", "TargetWS", "SourceColName", "SourceCount", "SourceSum", "TargetColName", "TargetCount", "TargetSum", "CountValidation", "SumValidation")
        wsValidation.Range("A2:J2").Font.Bold = True
        wsValidation.Range("A2:J2").Interior.Color = RGB(255, 204, 0)
    End If
    rowOut = wsValidation.Cells(wsValidation.rows.Count, "A").End(xlUp).Row + 1
    wsValidation.Cells(rowOut, 1).Value = SourceWSName
    wsValidation.Cells(rowOut, 2).Value = targetWSName
    wsValidation.Cells(rowOut, 3).Value = SourceColName
    wsValidation.Cells(rowOut, 4).Value = SourceCount
    wsValidation.Cells(rowOut, 5).Value = SourceSum
    wsValidation.Cells(rowOut, 6).Value = targetColName
    wsValidation.Cells(rowOut, 7).Value = TargetCount
    wsValidation.Cells(rowOut, 8).Value = TargetSum
    wsValidation.Cells(rowOut, 9).Value = CountStatus
    wsValidation.Cells(rowOut, 10).Value = SumStatus
End Sub
