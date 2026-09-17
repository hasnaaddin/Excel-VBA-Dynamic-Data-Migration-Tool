Attribute VB_Name = "MIGRATION"
Option Explicit

Sub RunDataMigration()
    Dim wbControl As Workbook, wsControl As Worksheet
    Dim SourceFile As String, TargetFile As String
    Dim wbSource As Workbook, wbTarget As Workbook
    Dim SourceWS As Worksheet, targetWS As Worksheet
    Dim SourceHeaderRow As Long, SourceHeaderCol As Long, SourceDataRow As Long
    Dim TargetHeaderRow As Long, TargetHeaderCol As Long, TargetDataRow As Long
    Dim lastRow As Long, i As Long
    Dim SourceColName As String, targetColName As String, DataType As String, FormatStr As String
    Dim SourceColIndex As Long, targetColIndex As Long
    Dim SourceHeaders As Variant, TargetHeaders As Variant
    Dim SourceData As Variant, TargetData As Variant
    Dim MissingColCell As Range
    Dim SavePath As String, NewFileName As String
    Dim LogRow As Long
    Dim r As Long ' <-- Declared once here at top
    
    On Error GoTo ErrHandler
    
    '--- Initialize
    Set wbControl = ThisWorkbook
    Set wsControl = wbControl.Sheets("Control")
    
    '--- Prepare Validation sheet
    Dim wsValidation As Worksheet
    On Error Resume Next
    Set wsValidation = wbControl.Sheets("Validation")
    On Error GoTo 0
    If wsValidation Is Nothing Then
        Set wsValidation = wbControl.Sheets.Add
        wsValidation.Name = "Validation"
        wsValidation.Range("A2:J2").Value = Array("SourceWS", "TargetWS", "SourceColName", "SourceCount", "SourceSum", "TargetColName", "TargetCount", "TargetSum", "CountValidation", "SumValidation")
        wsValidation.Range("A2:J2").Font.Bold = True
        wsValidation.Range("A2:J2").Interior.Color = RGB(255, 204, 0)
    Else
        wsValidation.rows("3:" & wsValidation.rows.Count).ClearContents
    End If
    
    '--- Prompt user for Source and Target files
    SourceFile = GetFilePath("Select Source File")
    TargetFile = GetFilePath("Select Target File")
    If SourceFile = "" Or TargetFile = "" Then
        MsgBox "File selection cancelled.", vbExclamation
        Exit Sub
    End If
    
    '--- Clear MissingColumn data
    Call ClearColumnDataByTitle(wsControl, "SourceFilename")
    Call ClearColumnDataByTitle(wsControl, "TargetFilename")
    Call ClearColumnDataByTitle(wsControl, "MissingColumn")
    
    '--- Write file paths in Control sheet
    Dim sourceCellAddress As String
    Dim targetCellAddress As String
    sourceCellAddress = GetNextEmptyCellAddress(wsControl, "SourceFilename")
    targetCellAddress = GetNextEmptyCellAddress(wsControl, "TargetFilename")
    wsControl.Range(sourceCellAddress).Value = SourceFile
    wsControl.Range(targetCellAddress).Value = TargetFile
    
    '--- Open Source and Target workbooks
    Set wbSource = Workbooks.Open(SourceFile)
    Set wbTarget = Workbooks.Open(TargetFile)
    
    '--- Read control parameters
    SourceHeaderRow = GetColumnDataByTitle(wsControl, "SourceHeaderRowStart")
    SourceHeaderCol = GetColumnDataByTitle(wsControl, "SourceHeaderColStart")
    SourceDataRow = GetColumnDataByTitle(wsControl, "SourceDataRowStart")
    TargetHeaderRow = GetColumnDataByTitle(wsControl, "TargetHeaderRowStart")
    TargetHeaderCol = GetColumnDataByTitle(wsControl, "TargetHeaderColStart")
    TargetDataRow = GetColumnDataByTitle(wsControl, "TargetDataRowStart")
    Set SourceWS = wbSource.Sheets(GetColumnDataByTitle(wsControl, "SourceWS"))
    
    '--- Load Source headers
    SourceHeaders = Application.Transpose(SourceWS.Range(SourceWS.Cells(SourceHeaderRow, SourceHeaderCol), _
        SourceWS.Cells(SourceHeaderRow, SourceWS.Cells(SourceHeaderRow, Columns.Count).End(xlToLeft).Column)))
    
    '--- Determine last mapping row
    lastRow = GetLastRowByTitle(wsControl, "TargetWS")
    
    '--- Prepare MissingColumn logging
    Set MissingColCell = wsControl.Range(GetNextEmptyCellAddress(wsControl, "MissingColumn"))
    LogRow = MissingColCell.Row
    
    '*** NEW LOGIC STARTS HERE ***
    '--- Calculate MaxRowCount for default values
    Dim MaxRowCount As Long
    MaxRowCount = 1
    Dim j As Long
    For j = 6 To lastRow
        Dim tempSourceColName As String
        tempSourceColName = Trim(wsControl.Cells(j, GetColumn(wsControl, "SourceColName")).Value)
        
        If tempSourceColName <> "" Then
            Dim tempSourceColIndex As Long
            tempSourceColIndex = GetColumnIndex(SourceHeaders, tempSourceColName)
            If tempSourceColIndex > 0 Then
                Dim tempLastRow As Long
                tempLastRow = SourceWS.Cells(SourceWS.rows.Count, tempSourceColIndex).End(xlUp).Row
                If tempLastRow >= SourceDataRow Then
                    MaxRowCount = Application.WorksheetFunction.Max(MaxRowCount, tempLastRow - SourceDataRow + 1)
                End If
            End If
        End If
    Next j
    '*** NEW LOGIC ENDS HERE ***
    
    '--- Loop through mappings
    For i = 6 To lastRow
        Dim TargetSheetName As String
        Dim RowCount As Long
        Dim defaultValue As Variant
        
        TargetSheetName = Trim(wsControl.Cells(i, GetColumn(wsControl, "TargetWS")).Value)
        SourceColName = Trim(wsControl.Cells(i, GetColumn(wsControl, "SourceColName")).Value)
        targetColName = Trim(wsControl.Cells(i, GetColumn(wsControl, "TargetColName")).Value)
        DataType = Trim(wsControl.Cells(i, GetColumn(wsControl, "DataType")).Value)
        FormatStr = Trim(wsControl.Cells(i, GetColumn(wsControl, "Format")).Value)
        defaultValue = wsControl.Cells(i, GetColumn(wsControl, "DEFAULT_VALUE")).Value
        '===========================================================================================
        '--- NEW LOGIC FOR "Not Available" RULES ---
        '===========================================================================================

        If UCase(SourceColName) = "NOT AVAILABLE" And UCase(targetColName) <> "NOT AVAILABLE" And Len(defaultValue) > 0 Then
            On Error Resume Next
            Set targetWS = wbTarget.Sheets(TargetSheetName)
            On Error GoTo 0

            If Not targetWS Is Nothing Then
                TargetHeaders = Application.Transpose(targetWS.Range(targetWS.Cells(TargetHeaderRow, TargetHeaderCol), _
                    targetWS.Cells(TargetHeaderRow, targetWS.Cells(TargetHeaderRow, Columns.Count).End(xlToLeft).Column)))

                targetColIndex = GetColumnIndex(TargetHeaders, targetColName)

                If targetColIndex > 0 Then
                    RowCount = MaxRowCount

                    Dim lastTargetRowNA As Long
                    lastTargetRowNA = targetWS.Cells(targetWS.rows.Count, targetColIndex).End(xlUp).Row
                    If lastTargetRowNA >= TargetDataRow Then TargetDataRow = lastTargetRowNA + 1

                    Call ApplyDefaultValue(targetWS, targetColIndex, TargetDataRow, RowCount, defaultValue, DataType, FormatStr)

                    GoTo SkipSourceCopy
                End If
            End If
        End If

        If UCase(SourceColName) = "NOT AVAILABLE" And UCase(targetColName) = "NOT AVAILABLE" Then
            wsControl.Cells(LogRow, GetColumn(wsControl, "MissingColumn")).Value = _
                "Both Source and Target Not Available for mapping on row " & i
            LogRow = LogRow + 1
            GoTo SkipSourceCopy
        End If

        If UCase(SourceColName) <> "NOT AVAILABLE" And UCase(targetColName) = "NOT AVAILABLE" Then
            wsControl.Cells(LogRow, GetColumn(wsControl, "MissingColumn")).Value = _
                "Missing Target Column (Not Available): " & targetColName & ", row " & i
            LogRow = LogRow + 1
            GoTo SkipSourceCopy
        End If

        If UCase(SourceColName) = "NOT AVAILABLE" And Len(defaultValue) = 0 Then
            wsControl.Cells(LogRow, GetColumn(wsControl, "MissingColumn")).Value = _
                "Source Not Available and no DEFAULT_VALUE for mapping on row " & i
            LogRow = LogRow + 1
            GoTo SkipSourceCopy
        End If

        '===========================================================================================
        '--- END NEW LOGIC INSERT ---
        '===========================================================================================

        
        If TargetSheetName <> "" And SourceColName <> "" And targetColName <> "" Then
            On Error Resume Next
            Set targetWS = wbTarget.Sheets(TargetSheetName)
            On Error GoTo 0
            
            If targetWS Is Nothing Then
                wsControl.Cells(LogRow, GetColumn(wsControl, "MissingColumn")).Value = "Missing Target Sheet: " & TargetSheetName
                LogRow = LogRow + 1
            Else
                TargetHeaders = Application.Transpose(targetWS.Range(targetWS.Cells(TargetHeaderRow, TargetHeaderCol), _
                    targetWS.Cells(TargetHeaderRow, targetWS.Cells(TargetHeaderRow, Columns.Count).End(xlToLeft).Column)))
                
                SourceColIndex = GetColumnIndex(SourceHeaders, SourceColName)
                targetColIndex = GetColumnIndex(TargetHeaders, targetColName)
                
                '--- Handle missing columns
                If SourceColIndex = 0 And targetColIndex = 0 Then
                    wsControl.Cells(LogRow, GetColumn(wsControl, "MissingColumn")).Value = "Missing Source Column in [" & SourceWS.Name & "]: " & SourceColName & _
                        " AND Missing Target Column in [" & targetWS.Name & "]: " & targetColName & _
                        ", for mapping " & SourceColName & " > " & targetColName & "."
                    LogRow = LogRow + 1
                ElseIf SourceColIndex = 0 Then
                    wsControl.Cells(LogRow, GetColumn(wsControl, "MissingColumn")).Value = "Missing Source Column in [" & SourceWS.Name & "]: " & SourceColName & _
                        ", for mapping " & SourceColName & " > " & targetColName & "."
                    LogRow = LogRow + 1
                ElseIf targetColIndex = 0 Then
                    wsControl.Cells(LogRow, GetColumn(wsControl, "MissingColumn")).Value = "Missing Target Column in [" & targetWS.Name & "]: " & targetColName & _
                        ", for mapping " & SourceColName & " > " & targetColName & "."
                    LogRow = LogRow + 1
                Else
                    '--- Determine RowCount
                    If Len(defaultValue) > 0 Then
                        RowCount = MaxRowCount
                    Else
                        Dim lastDataRow As Long
                        lastDataRow = SourceWS.Cells(SourceWS.rows.Count, SourceColIndex).End(xlUp).Row
                        If lastDataRow < SourceDataRow Then
                            RowCount = 1
                        Else
                            RowCount = lastDataRow - SourceDataRow + 1
                        End If
                    End If
                    
                    Dim SrcRange As Range, TgtRange As Range
                    Set SrcRange = SourceWS.Range(SourceWS.Cells(SourceDataRow, SourceColIndex), _
                        SourceWS.Cells(SourceDataRow + RowCount - 1, SourceColIndex))
                    SourceData = SrcRange.Value
                    
                    '--- Append logic
                    Dim lastTargetRow As Long
                    lastTargetRow = targetWS.Cells(targetWS.rows.Count, targetColIndex).End(xlUp).Row
                    If lastTargetRow >= TargetDataRow Then
                        TargetDataRow = lastTargetRow + 1
                    End If
                    
                    Set TgtRange = targetWS.Range(targetWS.Cells(TargetDataRow, targetColIndex), _
                        targetWS.Cells(TargetDataRow + RowCount - 1, targetColIndex))
                    
                    '--- Apply default value if present
                    If Len(defaultValue) > 0 Then
                        ApplyDefaultValue targetWS, targetColIndex, TargetDataRow, RowCount, defaultValue, DataType, FormatStr
                        GoTo SkipSourceCopy
                    End If
                    
                    '--- Apply NumberFormat dynamically
                    If Len(DataType) > 0 Then
                        Select Case UCase(DataType)
                            Case "TEXT"
                                TgtRange.NumberFormat = "@"
                            Case "NUMERIC"
                                If FormatStr <> "" Then
                                    TgtRange.NumberFormat = "0." & String(Val(Replace(FormatStr, " decimal", "")), "0")
                                Else
                                    TgtRange.NumberFormat = "General"
                                End If
                            Case "DATE"
                                If FormatStr <> "" Then
                                    TgtRange.NumberFormat = FormatStr
                                Else
                                    TgtRange.NumberFormat = SourceWS.Cells(SourceDataRow, SourceColIndex).NumberFormat
                                End If
                        End Select
                    End If
                    
                    '--- Copy source data to target
                    If RowCount = 1 Then
                        If IsEmpty(SourceData) Then
                            targetWS.Cells(TargetDataRow, targetColIndex).Value = Empty
                        Else
                            targetWS.Cells(TargetDataRow, targetColIndex).Value = SourceData
                        End If
                    Else
                        ReDim TargetData(1 To RowCount, 1 To 1)
                        For r = 1 To RowCount
                            If IsEmpty(SourceData(r, 1)) Then
                                TargetData(r, 1) = Empty
                            Else
                                TargetData(r, 1) = SourceData(r, 1)
                            End If
                        Next r
                        TgtRange.Value = TargetData
                    End If
                    
                    '--- Validation logic (fixed If/Else)
                    Dim SourceCount As Long, TargetCount As Long
                    Dim SourceSum As Variant, TargetSum As Variant
                    Dim validationStatus As String, sumValidationStatus As String
                    
                    SourceCount = Application.WorksheetFunction.CountA(SrcRange)
                    If UCase(DataType) = "NUMERIC" Then
                        SourceSum = Application.WorksheetFunction.Sum(SrcRange)
                    Else
                        SourceSum = ""
                    End If
                    
                    TargetCount = Application.WorksheetFunction.CountA(TgtRange)
                    If UCase(DataType) = "NUMERIC" Then
                        TargetSum = Application.WorksheetFunction.Sum(TgtRange)
                    Else
                        TargetSum = ""
                    End If
                    
                    '--- Count validation
                    If SourceCount = 0 Then
                        validationStatus = "No Data"
                    ElseIf SourceCount = TargetCount Then
                        validationStatus = "Match"
                    Else
                        validationStatus = "No Match"
                    End If
                    
                    '--- Sum validation
                    If UCase(DataType) = "NUMERIC" Then
                        If SourceCount = 0 Then
                            sumValidationStatus = "No Data"
                        ElseIf SourceSum = TargetSum Then
                            sumValidationStatus = "Match"
                        Else
                            sumValidationStatus = "No Match"
                        End If
                    Else
                        sumValidationStatus = ""
                    End If
                    
                    Call AppendValidationWithSum(SourceWS.Name, targetWS.Name, SourceColName, SourceCount, SourceSum, targetColName, TargetCount, TargetSum, validationStatus, sumValidationStatus)
SkipSourceCopy:
                End If
            End If
        End If
    Next i
    
    '--- Clean whitespace and report
    Call CleanTargetAndReportWhitespace(wbTarget, TargetHeaderRow, TargetDataRow)
    
    SavePath = wbTarget.Path
    NewFileName = SavePath & "\Migrated_Data_" & Format(Now, "YYYYMMDD_HHMMSS") & ".xlsx"
    wbTarget.SaveAs NewFileName, FileFormat:=51
    
    wbSource.Close False
    wbTarget.Close True
    
    '--- Add hyperlink to FrontPage
    With wbControl.Sheets("FrontPage")
        .Hyperlinks.Add Anchor:=.Range("L5"), _
            Address:=NewFileName, _
            TextToDisplay:="Open Latest File"
    End With
    
    MsgBox "Migration completed successfully!" & vbCrLf & "Saved as: " & NewFileName, vbInformation
    Exit Sub

ErrHandler:
    MsgBox "Error: " & Err.Description, vbCritical
End Sub


