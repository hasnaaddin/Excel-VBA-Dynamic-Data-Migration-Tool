Attribute VB_Name = "HELPER"
Function GetColumnDataByTitle(wsControl As Worksheet, title As String) As Variant
    Dim headerRow As Long
    Dim lastCol As Long
    Dim colIndex As Long
    Dim lastRow As Long
    
    headerRow = 5 ' Titles are in row 5
    
    ' Find last column in header row
    lastCol = wsControl.Cells(headerRow, wsControl.Columns.Count).End(xlToLeft).Column
    
    ' Find column index for the given title
    colIndex = 0
    Dim c As Long
    For c = 1 To lastCol
        If Trim(CStr(wsControl.Cells(headerRow, c).Value)) = title Then
            colIndex = c
            Exit For
        End If
    Next c
    
    If colIndex = 0 Then
        GetColumnDataByTitle = CVErr(xlErrNA) ' Title not found
        Exit Function
    End If
    
    ' Find last row of data in that column
    lastRow = wsControl.Cells(wsControl.rows.Count, colIndex).End(xlUp).Row
    
    If lastRow <= headerRow Then
        GetColumnDataByTitle = CVErr(xlErrNA) ' No data below title
        Exit Function
    End If
    
    ' Return the range of data below the title
    GetColumnDataByTitle = wsControl.Range(wsControl.Cells(headerRow + 1, colIndex), wsControl.Cells(lastRow, colIndex)).Value
End Function

Function GetFilePath(Prompt As String) As String
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogOpen)
    fd.title = Prompt
    fd.AllowMultiSelect = False
    If fd.Show = -1 Then
        GetFilePath = fd.SelectedItems(1)
    Else
        GetFilePath = ""
    End If
End Function

Function GetColumnIndex(HeaderArray As Variant, HeaderName As String) As Long
    Dim i As Long
    For i = LBound(HeaderArray) To UBound(HeaderArray)
        If Trim(HeaderArray(i, 1)) = HeaderName Then
            GetColumnIndex = i
            Exit Function
        End If
    Next i
    GetColumnIndex = 0
End Function

Function GetNextEmptyCellAddress(wsControl As Worksheet, title As String) As String
    Dim headerRow As Long
    Dim lastCol As Long
    Dim colIndex As Long
    Dim nextRow As Long
    
    headerRow = 5 ' Titles are in row 5
    
    ' Find last column in header row
    lastCol = wsControl.Cells(headerRow, wsControl.Columns.Count).End(xlToLeft).Column
    
    ' Find column index for the given title
    colIndex = 0
    Dim c As Long
    For c = 1 To lastCol
        If Trim(CStr(wsControl.Cells(headerRow, c).Value)) = title Then
            colIndex = c
            Exit For
        End If
    Next c
    
    If colIndex = 0 Then
        GetNextEmptyCellAddress = "" ' Title not found
        Exit Function
    End If
    
    ' Find next empty row in that column
    nextRow = wsControl.Cells(wsControl.rows.Count, colIndex).End(xlUp).Row + 1
    If nextRow <= headerRow Then nextRow = headerRow + 1
    
    ' Return the cell address
    GetNextEmptyCellAddress = wsControl.Cells(nextRow, colIndex).Address
End Function

Sub ClearColumnDataByTitle(wsControl As Worksheet, title As String)
    Dim headerRow As Long
    Dim lastCol As Long
    Dim colIndex As Long
    Dim lastRow As Long
    
    headerRow = 5 ' Titles are in row 5
    
    ' Find last column in header row
    ' Find column index for the given title
    lastCol = wsControl.Cells(headerRow, wsControl.Columns.Count).End(xlToLeft).Column
    colIndex = 0
    Dim c As Long
    For c = 1 To lastCol
        If Trim(CStr(wsControl.Cells(headerRow, c).Value)) = title Then
            colIndex = c
            Exit For
        End If
    Next c
    
    If colIndex = 0 Then
        MsgBox "Title '" & title & "' not found in row " & headerRow, vbExclamation
        Exit Sub
    End If
    
    ' Find last row of data in that column
    lastRow = wsControl.Cells(wsControl.rows.Count, colIndex).End(xlUp).Row
    
    ' If there is data below the header, clear it
    If lastRow > headerRow Then
        wsControl.Range(wsControl.Cells(headerRow + 1, colIndex), wsControl.Cells(lastRow, colIndex)).ClearContents
    End If
End Sub

Function GetLastRowByTitle(wsControl As Worksheet, title As String) As Long
    Dim headerRow As Long
    Dim lastCol As Long
    Dim colIndex As Long
    
    headerRow = 5 ' Titles are in row 5
    
    ' Find last column in header row
    lastCol = wsControl.Cells(headerRow, wsControl.Columns.Count).End(xlToLeft).Column
    
    ' Find column index for the given title
    colIndex = 0
    Dim c As Long
    For c = 1 To lastCol
        If Trim(CStr(wsControl.Cells(headerRow, c).Value)) = title Then
            colIndex = c
            Exit For
        End If
    Next c
    
    If colIndex = 0 Then
        GetLastRowByTitle = 0 ' Title not found
        Exit Function
    End If
    
    ' Get last non-empty row in that column
    GetLastRowByTitle = wsControl.Cells(wsControl.rows.Count, colIndex).End(xlUp).Row
End Function

Function GetColumn(wsControl As Worksheet, title As String) As Long
    Dim headerRow As Long
    Dim lastCol As Long
    Dim colIndex As Long
    
    headerRow = 5 ' Titles are in row 5
    
    ' Find last column in header row
    lastCol = wsControl.Cells(headerRow, wsControl.Columns.Count).End(xlToLeft).Column
    
    ' Initialize
    colIndex = 0
    
    ' Loop through columns to find the title
    Dim c As Long
    For c = 1 To lastCol
        If StrComp(Trim(CStr(wsControl.Cells(headerRow, c).Value)), title, vbTextCompare) = 0 Then
            colIndex = c
            Exit For
        End If
    Next c
    
    ' Return column index or 0 if not found
    GetColumn = colIndex
End Function
