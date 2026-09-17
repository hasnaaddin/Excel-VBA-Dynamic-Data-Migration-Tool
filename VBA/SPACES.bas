Attribute VB_Name = "SPACES"
Sub CleanTargetAndReportWhitespace(wbTarget As Workbook, TargetHeaderRow As Long, TargetDataRowStart As Long)
    Dim wsValidation As Worksheet
    Dim wsTarget As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim targetWSName As String, targetColName As String
    Dim headerRow As Long, dataStartRow As Long
    Dim targetColIndex As Long
    Dim cellValue As String
    Dim reportRow As Long
    
    ' Get Validation sheet
    On Error Resume Next
    Set wsValidation = ThisWorkbook.Sheets("Validation")
    On Error GoTo 0
    
    If wsValidation Is Nothing Then
        MsgBox "Validation sheet not found!", vbCritical
        Exit Sub
    End If
    
    ' Set header and data start rows from parameters
    headerRow = TargetHeaderRow
    dataStartRow = TargetDataRowStart
    
    ' Find last row in Validation column B
    lastRow = wsValidation.Cells(wsValidation.rows.Count, "B").End(xlUp).Row
    
    ' Start reporting from row 3 in Column O/P/Q
    reportRow = 3
    
    ' Loop through each TargetWS in Validation
    For i = 3 To lastRow
        targetWSName = Trim(wsValidation.Cells(i, "B").Value)
        targetColName = Trim(wsValidation.Cells(i, "F").Value)
        
        If Len(targetWSName) > 0 And Len(targetColName) > 0 Then
            On Error Resume Next
            Set wsTarget = wbTarget.Sheets(targetWSName)
            On Error GoTo 0
            
            If Not wsTarget Is Nothing Then
                ' Find column index for TargetColName in header row
                targetColIndex = 0
                Dim col As Long
                For col = 1 To wsTarget.Cells(headerRow, wsTarget.Columns.Count).End(xlToLeft).Column
                    If Trim(CStr(wsTarget.Cells(headerRow, col).Value)) = targetColName Then
                        targetColIndex = col
                        Exit For
                    End If
                Next col
                
                ' If column found, check data for whitespace issues
                If targetColIndex > 0 Then
                    Dim lastDataRow As Long
                    lastDataRow = wsTarget.Cells(wsTarget.rows.Count, targetColIndex).End(xlUp).Row
                    
                    Dim r As Long
                    For r = dataStartRow To lastDataRow
                        cellValue = CStr(wsTarget.Cells(r, targetColIndex).Value)
                        
                        If Len(cellValue) > 0 Then
                            ' Check for leading/trailing spaces or double spaces inside
                            If Left(cellValue, 1) = " " Or Right(cellValue, 1) = " " Or InStr(cellValue, "  ") > 0 Then
                                wsValidation.Cells(reportRow, "O").Value = targetWSName
                                wsValidation.Cells(reportRow, "P").Value = targetColName
                                wsValidation.Cells(reportRow, "Q").Value = cellValue
                                reportRow = reportRow + 1
                            End If
                        End If
                    Next r
                End If
            End If
            
            Set wsTarget = Nothing
        End If
    Next i
    
    ' -------------------------------
    ' NEW SECTION: CLEAN REPORTED DATA
    ' -------------------------------
    Dim cleanLastRow As Long
    cleanLastRow = wsValidation.Cells(wsValidation.rows.Count, "O").End(xlUp).Row
    
    For i = 3 To cleanLastRow
        targetWSName = Trim(wsValidation.Cells(i, "O").Value)
        targetColName = Trim(wsValidation.Cells(i, "P").Value)
        cellValue = CStr(wsValidation.Cells(i, "Q").Value)
        
        If Len(targetWSName) > 0 And Len(targetColName) > 0 And Len(cellValue) > 0 Then
            On Error Resume Next
            Set wsTarget = wbTarget.Sheets(targetWSName)
            On Error GoTo 0
            
            If Not wsTarget Is Nothing Then
                ' Find column index again
                targetColIndex = 0
                Dim col2 As Long
                For col2 = 1 To wsTarget.Cells(headerRow, wsTarget.Columns.Count).End(xlToLeft).Column
                    If Trim(CStr(wsTarget.Cells(headerRow, col2).Value)) = targetColName Then
                        targetColIndex = col2
                        Exit For
                    End If
                Next col2
                
                If targetColIndex > 0 Then
                    Dim lastDataRow2 As Long
                    lastDataRow2 = wsTarget.Cells(wsTarget.rows.Count, targetColIndex).End(xlUp).Row
                    
                    Dim r2 As Long
                    For r2 = dataStartRow To lastDataRow2
                        If CStr(wsTarget.Cells(r2, targetColIndex).Value) = cellValue Then
                            ' Clean the value
                            wsTarget.Cells(r2, targetColIndex).Value = WorksheetFunction.Trim(Replace(cellValue, "  ", " "))
                        End If
                    Next r2
                End If
            End If
            
            Set wsTarget = Nothing
        End If
    Next i
    
    'MsgBox "Whitespace cleaning completed based on Validation report.", vbInformation
End Sub
