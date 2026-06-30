Attribute VB_Name = "ExcelReader"
Option Explicit

Private Const XL_UP As Long = -4162
Private Const XL_TO_LEFT As Long = -4159
Private Const HEADER_ROW_NUMBER As Long = 1
Private Const FIRST_DATA_ROW_NUMBER As Long = 2
Private Const DEFAULT_COLUMN_PREFIX As String = "Column"

Private mExcelApplication As Object
Private mWorkbook As Object
Private mWorksheet As Object
Private mWorkbookPath As String

' Creates the Excel application instance using late binding.
Public Function InitializeExcel() As Boolean
    On Error GoTo ErrorHandler

    If mExcelApplication Is Nothing Then
        Set mExcelApplication = CreateObject("Excel.Application")
        mExcelApplication.Visible = False
        mExcelApplication.DisplayAlerts = False
        mExcelApplication.EnableEvents = False
        WriteLog "Excel application initialized."
    Else
        WriteLog "Excel application already initialized."
    End If

    InitializeExcel = True
    Exit Function

ErrorHandler:
    WriteErrorLog "ExcelReader.InitializeExcel", Err.Number, Err.Description
    CloseExcel
    InitializeExcel = False
End Function

' Closes the workbook, quits Excel, and releases all Excel COM objects owned by this module.
Public Sub CloseExcel()
    On Error Resume Next

    CloseWorkbook

    If Not mExcelApplication Is Nothing Then
        mExcelApplication.DisplayAlerts = False
        mExcelApplication.EnableEvents = False
        mExcelApplication.Quit
        WriteLog "Excel application closed."
    End If

    Set mExcelApplication = Nothing
End Sub

' Opens an Excel workbook using late binding. A blank path uses the configured default workbook.
Public Function OpenWorkbook(ByVal WorkbookPath As String) As Boolean
    On Error GoTo ErrorHandler

    WorkbookPath = ResolveWorkbookPath(WorkbookPath)

    If Len(WorkbookPath) = 0 Then
        WriteLog "Workbook path is empty."
        OpenWorkbook = False
        Exit Function
    End If

    If Not FileExists(WorkbookPath) Then
        WriteLog "Workbook not found: " & WorkbookPath
        CloseExcel
        OpenWorkbook = False
        Exit Function
    End If

    CloseWorkbook

    If Not InitializeExcel() Then
        OpenWorkbook = False
        Exit Function
    End If

    Set mWorkbook = mExcelApplication.Workbooks.Open(WorkbookPath, False, True)
    Set mWorksheet = Nothing
    mWorkbookPath = WorkbookPath

    OpenWorkbook = Not mWorkbook Is Nothing
    WriteLog "Workbook opened: " & WorkbookPath
    Exit Function

ErrorHandler:
    WriteErrorLog "ExcelReader.OpenWorkbook", Err.Number, Err.Description
    CloseExcel
    OpenWorkbook = False
End Function

' Closes the active workbook and releases workbook/worksheet objects.
Public Sub CloseWorkbook()
    On Error Resume Next

    If Not mWorkbook Is Nothing Then
        mWorkbook.Close False
        WriteLog "Workbook closed: " & mWorkbookPath
    End If

    Set mWorksheet = Nothing
    Set mWorkbook = Nothing
    mWorkbookPath = vbNullString
End Sub

' Sets and returns the active worksheet by name.
Public Function GetWorksheet(ByVal SheetName As String) As Object
    On Error GoTo ErrorHandler

    SheetName = Trim$(SheetName)

    If mWorkbook Is Nothing Then
        WriteLog "Cannot get worksheet because no workbook is open."
        Set GetWorksheet = Nothing
        Exit Function
    End If

    If Len(SheetName) = 0 Then
        WriteLog "Worksheet name is empty."
        Set GetWorksheet = Nothing
        Exit Function
    End If

    Set mWorksheet = mWorkbook.Worksheets(SheetName)
    Set GetWorksheet = mWorksheet
    WriteLog "Worksheet selected: " & SheetName
    Exit Function

ErrorHandler:
    WriteErrorLog "ExcelReader.GetWorksheet", Err.Number, "Worksheet not found: " & SheetName & ". " & Err.Description
    CloseExcel
    Set mWorksheet = Nothing
    Set GetWorksheet = Nothing
End Function

' Returns the last used row in the active worksheet for the supplied column index.
Public Function GetLastRow(ByVal ColumnIndex As Long) As Long
    Dim candidateRow As Long

    On Error GoTo ErrorHandler

    If mWorksheet Is Nothing Or ColumnIndex < 1 Then
        GetLastRow = 0
        Exit Function
    End If

    candidateRow = mWorksheet.Cells(mWorksheet.Rows.Count, ColumnIndex).End(XL_UP).Row

    If candidateRow = 1 And Len(CellValueToText(mWorksheet.Cells(candidateRow, ColumnIndex).Value)) = 0 Then
        GetLastRow = 0
    Else
        GetLastRow = candidateRow
    End If

    Exit Function

ErrorHandler:
    WriteErrorLog "ExcelReader.GetLastRow", Err.Number, Err.Description
    CloseExcel
    GetLastRow = 0
End Function

' Returns the last used column in the active worksheet for the supplied row index.
Public Function GetLastColumn(ByVal RowIndex As Long) As Long
    Dim candidateColumn As Long

    On Error GoTo ErrorHandler

    If mWorksheet Is Nothing Or RowIndex < 1 Then
        GetLastColumn = 0
        Exit Function
    End If

    candidateColumn = mWorksheet.Cells(RowIndex, mWorksheet.Columns.Count).End(XL_TO_LEFT).Column

    If candidateColumn = 1 And Len(CellValueToText(mWorksheet.Cells(RowIndex, candidateColumn).Value)) = 0 Then
        GetLastColumn = 0
    Else
        GetLastColumn = candidateColumn
    End If

    Exit Function

ErrorHandler:
    WriteErrorLog "ExcelReader.GetLastColumn", Err.Number, Err.Description
    CloseExcel
    GetLastColumn = 0
End Function

' Reads a single cell from the active worksheet.
Public Function ReadCell(ByVal Row As Long, ByVal Column As Long) As Variant
    On Error GoTo ErrorHandler

    If mWorksheet Is Nothing Or Row < 1 Or Column < 1 Then
        ReadCell = Empty
        Exit Function
    End If

    ReadCell = mWorksheet.Cells(Row, Column).Value
    Exit Function

ErrorHandler:
    WriteErrorLog "ExcelReader.ReadCell", Err.Number, Err.Description
    CloseExcel
    ReadCell = Empty
End Function

' Reads all populated columns between StartRow and EndRow from the active worksheet.
Public Function ReadRange(ByVal StartRow As Long, ByVal EndRow As Long) As Variant
    Dim lastColumn As Long

    On Error GoTo ErrorHandler

    If mWorksheet Is Nothing Or StartRow < 1 Or EndRow < StartRow Then
        ReadRange = Empty
        Exit Function
    End If

    lastColumn = GetLastColumn(StartRow)

    If lastColumn = 0 Then
        ReadRange = Empty
        Exit Function
    End If

    ReadRange = mWorksheet.Range(mWorksheet.Cells(StartRow, 1), mWorksheet.Cells(EndRow, lastColumn)).Value
    WriteLog "Range read: rows " & CStr(StartRow) & " to " & CStr(EndRow) & ", columns 1 to " & CStr(lastColumn) & "."
    Exit Function

ErrorHandler:
    WriteErrorLog "ExcelReader.ReadRange", Err.Number, Err.Description
    CloseExcel
    ReadRange = Empty
End Function

' Compatibility helper for existing workflow modules.
Public Function GetCellValue(ByVal SheetName As String, ByVal RowNumber As Long, ByVal ColumnNumber As Long) As Variant
    On Error GoTo ErrorHandler

    If GetWorksheet(SheetName) Is Nothing Then
        GetCellValue = Empty
    Else
        GetCellValue = ReadCell(RowNumber, ColumnNumber)
    End If

    Exit Function

ErrorHandler:
    WriteErrorLog "ExcelReader.GetCellValue", Err.Number, Err.Description
    CloseExcel
    GetCellValue = Empty
End Function

' Compatibility helper for existing workflow modules.
Public Function GetHeaders(ByVal SheetName As String) As Collection
    Dim headers As Collection
    Dim lastColumn As Long
    Dim columnNumber As Long

    On Error GoTo ErrorHandler

    Set headers = New Collection

    If GetWorksheet(SheetName) Is Nothing Then
        Set GetHeaders = headers
        Exit Function
    End If

    lastColumn = GetLastColumn(HEADER_ROW_NUMBER)

    For columnNumber = 1 To lastColumn
        headers.Add GetHeaderName(columnNumber)
    Next columnNumber

    Set GetHeaders = headers
    Exit Function

ErrorHandler:
    WriteErrorLog "ExcelReader.GetHeaders", Err.Number, Err.Description
    CloseExcel
    Set GetHeaders = New Collection
End Function

' Compatibility helper for existing workflow modules.
Public Function ReadRow(ByVal SheetName As String, ByVal RowNumber As Long) As Object
    Dim rowData As Object
    Dim headers As Collection
    Dim columnNumber As Long
    Dim keyName As String

    On Error GoTo ErrorHandler

    Set rowData = CreateDictionary()

    If RowNumber < FIRST_DATA_ROW_NUMBER Then
        Set ReadRow = rowData
        Exit Function
    End If

    Set headers = GetHeaders(SheetName)

    For columnNumber = 1 To headers.Count
        keyName = GetUniqueDictionaryKey(rowData, CStr(headers(columnNumber)))
        rowData.Add keyName, ReadCell(RowNumber, columnNumber)
    Next columnNumber

    Set ReadRow = rowData
    Exit Function

ErrorHandler:
    WriteErrorLog "ExcelReader.ReadRow", Err.Number, Err.Description
    CloseExcel
    Set ReadRow = CreateDictionary()
End Function

' Compatibility helper for existing workflow modules.
Public Function ReadAllRows(ByVal SheetName As String) As Collection
    Dim rows As Collection
    Dim headers As Collection
    Dim rowNumber As Long
    Dim lastRow As Long
    Dim rowData As Object

    On Error GoTo ErrorHandler

    Set rows = New Collection
    Set headers = GetHeaders(SheetName)

    If headers.Count = 0 Then
        Set ReadAllRows = rows
        Exit Function
    End If

    lastRow = GetLastDataRow(headers.Count)

    For rowNumber = FIRST_DATA_ROW_NUMBER To lastRow
        If Not IsRowEmpty(rowNumber, headers.Count) Then
            Set rowData = ReadRow(SheetName, rowNumber)
            rows.Add rowData
        End If
    Next rowNumber

    WriteLog "Worksheet rows read: " & CStr(rows.Count)
    Set ReadAllRows = rows
    Exit Function

ErrorHandler:
    WriteErrorLog "ExcelReader.ReadAllRows", Err.Number, Err.Description
    CloseExcel
    Set ReadAllRows = New Collection
End Function

Private Function ResolveWorkbookPath(ByVal WorkbookPath As String) As String
    On Error GoTo ErrorHandler

    WorkbookPath = Trim$(WorkbookPath)

    If Len(WorkbookPath) = 0 Then
        If Len(ExcelFilePath) = 0 Then
            If Not InitializeConfig() Then
                ResolveWorkbookPath = vbNullString
                Exit Function
            End If
        End If

        WorkbookPath = ExcelFilePath
    End If

    ResolveWorkbookPath = WorkbookPath
    Exit Function

ErrorHandler:
    WriteErrorLog "ExcelReader.ResolveWorkbookPath", Err.Number, Err.Description
    ResolveWorkbookPath = vbNullString
End Function

Private Function GetLastDataRow(ByVal ColumnCount As Long) As Long
    Dim columnNumber As Long
    Dim currentLastRow As Long
    Dim maxLastRow As Long

    On Error GoTo ErrorHandler

    For columnNumber = 1 To ColumnCount
        currentLastRow = GetLastRow(columnNumber)

        If currentLastRow > maxLastRow Then
            maxLastRow = currentLastRow
        End If
    Next columnNumber

    GetLastDataRow = maxLastRow
    Exit Function

ErrorHandler:
    WriteErrorLog "ExcelReader.GetLastDataRow", Err.Number, Err.Description
    GetLastDataRow = 0
End Function

Private Function GetHeaderName(ByVal ColumnNumber As Long) As String
    Dim headerName As String

    On Error GoTo ErrorHandler

    headerName = CellValueToText(ReadCell(HEADER_ROW_NUMBER, ColumnNumber))

    If Len(headerName) = 0 Then
        headerName = DEFAULT_COLUMN_PREFIX & CStr(ColumnNumber)
    End If

    GetHeaderName = headerName
    Exit Function

ErrorHandler:
    WriteErrorLog "ExcelReader.GetHeaderName", Err.Number, Err.Description
    GetHeaderName = DEFAULT_COLUMN_PREFIX & CStr(ColumnNumber)
End Function

Private Function IsRowEmpty(ByVal RowNumber As Long, ByVal ColumnCount As Long) As Boolean
    Dim columnNumber As Long

    On Error GoTo ErrorHandler

    For columnNumber = 1 To ColumnCount
        If Len(CellValueToText(ReadCell(RowNumber, columnNumber))) > 0 Then
            IsRowEmpty = False
            Exit Function
        End If
    Next columnNumber

    IsRowEmpty = True
    Exit Function

ErrorHandler:
    WriteErrorLog "ExcelReader.IsRowEmpty", Err.Number, Err.Description
    IsRowEmpty = True
End Function

Private Function CreateDictionary() As Object
    Dim dictionary As Object

    On Error GoTo ErrorHandler

    Set dictionary = CreateObject("Scripting.Dictionary")
    dictionary.CompareMode = vbTextCompare

    Set CreateDictionary = dictionary
    Exit Function

ErrorHandler:
    WriteErrorLog "ExcelReader.CreateDictionary", Err.Number, Err.Description
    Set CreateDictionary = Nothing
End Function

Private Function GetUniqueDictionaryKey(ByVal dictionary As Object, ByVal keyName As String) As String
    Dim baseKeyName As String
    Dim candidateKeyName As String
    Dim duplicateIndex As Long

    On Error GoTo ErrorHandler

    baseKeyName = Trim$(keyName)

    If Len(baseKeyName) = 0 Then
        baseKeyName = DEFAULT_COLUMN_PREFIX
    End If

    candidateKeyName = baseKeyName
    duplicateIndex = 2

    Do While dictionary.Exists(candidateKeyName)
        candidateKeyName = baseKeyName & "_" & CStr(duplicateIndex)
        duplicateIndex = duplicateIndex + 1
    Loop

    GetUniqueDictionaryKey = candidateKeyName
    Exit Function

ErrorHandler:
    WriteErrorLog "ExcelReader.GetUniqueDictionaryKey", Err.Number, Err.Description
    GetUniqueDictionaryKey = DEFAULT_COLUMN_PREFIX
End Function

Private Function CellValueToText(ByVal CellValue As Variant) As String
    On Error GoTo ErrorHandler

    If IsError(CellValue) Or IsNull(CellValue) Or IsEmpty(CellValue) Then
        CellValueToText = vbNullString
    Else
        CellValueToText = Trim$(CStr(CellValue))
    End If

    Exit Function

ErrorHandler:
    WriteErrorLog "ExcelReader.CellValueToText", Err.Number, Err.Description
    CellValueToText = vbNullString
End Function
