Attribute VB_Name = "DEFAULT_VALUE"
Sub ApplyDefaultValue(targetWS As Worksheet, targetColIndex As Long, TargetDataRow As Long, _
                      RowCount As Long, defaultValue As Variant, DataType As String, FormatStr As String)

    Dim r As Long
    Dim TgtRange As Range
    Set TgtRange = targetWS.Range(targetWS.Cells(TargetDataRow, targetColIndex), _
                                  targetWS.Cells(TargetDataRow + RowCount - 1, targetColIndex))

    ' Apply number/date formatting same as normal logic
    If Len(DataType) > 0 Then
        Select Case UCase(DataType)
            Case "TEXT":
                TgtRange.NumberFormat = "@"
            Case "NUMERIC":
                If FormatStr <> "" Then
                    TgtRange.NumberFormat = "0." & String(Val(Replace(FormatStr, " decimal", "")), "0")
                Else
                    TgtRange.NumberFormat = "General"
                End If
            Case "DATE":
                If FormatStr <> "" Then
                    TgtRange.NumberFormat = FormatStr
                Else
                    TgtRange.NumberFormat = FormatStr
                End If
        End Select
    End If
    
    ' Fill with default value
    For r = 1 To RowCount
        TgtRange.Cells(r, 1).Value = defaultValue
    Next r

End Sub


