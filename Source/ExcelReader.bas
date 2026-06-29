Attribute VB_Name = "ExcelReader"
Option Explicit

Private Const XL_UP As Long = -4162
Private Const XL_TO_LEFT As Long = -4159
Private Const HEADER_ROW_NUMBER As Long = 1
Private Const FIRST_DATA_ROW_NUMBER As Long = 2
Private Const DEFAULT_COLUMN_PREFIX As String = "Column"

Private mExcelApplication As Object
Private mWorkbook As Object

' Opens an Excel workbook using late binding and stores the Excel objects internally.
' Returns True when the workbook is opened successfully.
Public Function OpenWorkbook(ByVal filePath As String) As Boolean
    On Error GoTo ErrorHandler

    filePath = Trim$(filePath)

    If Len(filePath) = 0 Or Not LocalFileExists(filePath) Then
        OpenWorkbook = False
        Exit Function
    End If

    CloseWorkbook

    Set mExcelApplication = CreateObject("Excel.Application")
    mExcelApplication.Visible = False
    mExcelApplication.DisplayAlerts = False
    mExcelApplication.EnableEvents = False

    Set mWorkbook = mExcelApplication.Workbooks.Open(filePath, False, True)
    OpenWorkbook = Not mWorkbook Is Nothing
    Exit Function

ErrorHandler:
    CloseWorkbook
    OpenWorkbook = False
End Function

' Safely closes the active workbook and releases all Excel COM objects owned by this module.
Public Sub CloseWorkbook()
    On Error Resume Next

    If Not mWorkbook Is Nothing Then
        mWorkbook.Close False
    End If

    Set mWorkbook = Nothing

    If Not mExcelApplication Is Nothing Then
        mExcelApplication.DisplayAlerts = True
        mExcelApplication.EnableEvents = True
        mExcelApplication.Quit
    End If

    Set mExcelApplication = Nothing
End Sub

' Returns the last used row number for the specified worksheet and column.
' Returns 0 when the workbook, worksheet, or column is not valid.
Public Function GetLastRow(ByVal sheetName As String, ByVal columnNumber As Long) As Long
    Dim worksheet As Object
    Dim candidateRow As Long

    On Error GoTo ErrorHandler

    If columnNumber < 1 Then
        GetLastRow = 0
        Exit Function
    End If

    Set worksheet = GetWorksheet(sheetName)

    If worksheet Is Nothing Then
        GetLastRow = 0
        Exit Function
    End If

    candidateRow = worksheet.Cells(worksheet.Rows.Count, columnNumber).End(XL_UP).Row

    If candidateRow = HEADER_ROW_NUMBER And Len(CellValueToText(worksheet.Cells(candidateRow, columnNumber).Value)) = 0 Then
        GetLastRow = 0
    Else
        GetLastRow = candidateRow
    End If

    Exit Function

ErrorHandler:
    GetLastRow = 0
End Function

' Returns the value from the requested worksheet cell.
' Returns Empty when the workbook, worksheet, row, or column is not valid.
Public Function GetCellValue(ByVal sheetName As String, ByVal rowNumber As Long, ByVal columnNumber As Long) As Variant
    Dim worksheet As Object

    On Error GoTo ErrorHandler

    If rowNumber < 1 Or columnNumber < 1 Then
        GetCellValue = Empty
        Exit Function
    End If

    Set worksheet = GetWorksheet(sheetName)

    If worksheet Is Nothing Then
        GetCellValue = Empty
        Exit Function
    End If

    GetCellValue = worksheet.Cells(rowNumber, columnNumber).Value
    Exit Function

ErrorHandler:
    GetCellValue = Empty
End Function

' Returns the first-row column names for the requested worksheet.
' Blank header cells are returned as stable fallback names such as Column3.
Public Function GetHeaders(ByVal sheetName As String) As Collection
    Dim headers As Collection
    Dim worksheet As Object
    Dim lastColumn As Long
    Dim columnNumber As Long

    On Error GoTo ErrorHandler

    Set headers = New Collection
    Set worksheet = GetWorksheet(sheetName)

    If worksheet Is Nothing Then
        Set GetHeaders = headers
        Exit Function
    End If

    lastColumn = GetLastHeaderColumn(worksheet)

    For columnNumber = 1 To lastColumn
        headers.Add GetHeaderName(worksheet, columnNumber)
    Next columnNumber

    Set GetHeaders = headers
    Exit Function

ErrorHandler:
    Set GetHeaders = New Collection
End Function

' Reads a worksheet row into a late-bound Scripting.Dictionary keyed by header name.
' Returns an empty Dictionary when the workbook, worksheet, or row is not valid.
Public Function ReadRow(ByVal sheetName As String, ByVal rowNumber As Long) As Object
    Dim rowData As Object
    Dim headers As Collection
    Dim columnNumber As Long
    Dim keyName As String

    On Error GoTo ErrorHandler

    Set rowData = CreateDictionary()

    If rowNumber < FIRST_DATA_ROW_NUMBER Then
        Set ReadRow = rowData
        Exit Function
    End If

    Set headers = GetHeaders(sheetName)

    For columnNumber = 1 To headers.Count
        keyName = GetUniqueDictionaryKey(rowData, CStr(headers(columnNumber)))
        rowData.Add keyName, GetCellValue(sheetName, rowNumber, columnNumber)
    Next columnNumber

    Set ReadRow = rowData
    Exit Function

ErrorHandler:
    Set ReadRow = CreateDictionary()
End Function

' Reads all data rows from the requested worksheet into a Collection of Dictionaries.
' The first row is treated as headers and data reading begins on row 2.
Public Function ReadAllRows(ByVal sheetName As String) As Collection
    Dim rows As Collection
    Dim headers As Collection
    Dim rowNumber As Long
    Dim lastRow As Long
    Dim rowData As Object

    On Error GoTo ErrorHandler

    Set rows = New Collection
    Set headers = GetHeaders(sheetName)

    If headers.Count = 0 Then
        Set ReadAllRows = rows
        Exit Function
    End If

    lastRow = GetLastDataRow(sheetName, headers.Count)

    For rowNumber = FIRST_DATA_ROW_NUMBER To lastRow
        If Not IsRowEmpty(sheetName, rowNumber, headers.Count) Then
            Set rowData = ReadRow(sheetName, rowNumber)
            rows.Add rowData
        End If
    Next rowNumber

    Set ReadAllRows = rows
    Exit Function

ErrorHandler:
    Set ReadAllRows = New Collection
End Function

Private Function GetWorksheet(ByVal sheetName As String) As Object
    On Error GoTo ErrorHandler

    sheetName = Trim$(sheetName)

    If mWorkbook Is Nothing Or Len(sheetName) = 0 Then
        Set GetWorksheet = Nothing
        Exit Function
    End If

    Set GetWorksheet = mWorkbook.Worksheets(sheetName)
    Exit Function

ErrorHandler:
    Set GetWorksheet = Nothing
End Function

Private Function GetLastHeaderColumn(ByVal worksheet As Object) As Long
    Dim candidateColumn As Long

    On Error GoTo ErrorHandler

    candidateColumn = worksheet.Cells(HEADER_ROW_NUMBER, worksheet.Columns.Count).End(XL_TO_LEFT).Column

    If candidateColumn = 1 And Len(CellValueToText(worksheet.Cells(HEADER_ROW_NUMBER, 1).Value)) = 0 Then
        GetLastHeaderColumn = 0
    Else
        GetLastHeaderColumn = candidateColumn
    End If

    Exit Function

ErrorHandler:
    GetLastHeaderColumn = 0
End Function

Private Function GetLastDataRow(ByVal sheetName As String, ByVal headerCount As Long) As Long
    Dim columnNumber As Long
    Dim currentLastRow As Long
    Dim maxLastRow As Long

    On Error GoTo ErrorHandler

    For columnNumber = 1 To headerCount
        currentLastRow = GetLastRow(sheetName, columnNumber)

        If currentLastRow > maxLastRow Then
            maxLastRow = currentLastRow
        End If
    Next columnNumber

    GetLastDataRow = maxLastRow
    Exit Function

ErrorHandler:
    GetLastDataRow = 0
End Function

Private Function GetHeaderName(ByVal worksheet As Object, ByVal columnNumber As Long) As String
    Dim headerName As String

    On Error GoTo ErrorHandler

    headerName = CellValueToText(worksheet.Cells(HEADER_ROW_NUMBER, columnNumber).Value)

    If Len(headerName) = 0 Then
        headerName = DEFAULT_COLUMN_PREFIX & CStr(columnNumber)
    End If

    GetHeaderName = headerName
    Exit Function

ErrorHandler:
    GetHeaderName = DEFAULT_COLUMN_PREFIX & CStr(columnNumber)
End Function

Private Function IsRowEmpty(ByVal sheetName As String, ByVal rowNumber As Long, ByVal columnCount As Long) As Boolean
    Dim columnNumber As Long
    Dim cellValue As Variant

    On Error GoTo ErrorHandler

    For columnNumber = 1 To columnCount
        cellValue = GetCellValue(sheetName, rowNumber, columnNumber)

        If Len(CellValueToText(cellValue)) > 0 Then
            IsRowEmpty = False
            Exit Function
        End If
    Next columnNumber

    IsRowEmpty = True
    Exit Function

ErrorHandler:
    IsRowEmpty = True
End Function

Private Function CreateDictionary() As Object
    Dim dictionary As Object

    Set dictionary = CreateObject("Scripting.Dictionary")
    dictionary.CompareMode = vbTextCompare

    Set CreateDictionary = dictionary
End Function

Private Function GetUniqueDictionaryKey(ByVal dictionary As Object, ByVal keyName As String) As String
    Dim baseKeyName As String
    Dim candidateKeyName As String
    Dim duplicateIndex As Long

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
End Function

Private Function LocalFileExists(ByVal filePath As String) As Boolean
    Dim attributes As Long

    On Error GoTo ErrorHandler

    attributes = GetAttr(filePath)
    LocalFileExists = ((attributes And vbDirectory) = 0)
    Exit Function

ErrorHandler:
    LocalFileExists = False
End Function

Private Function CellValueToText(ByVal cellValue As Variant) As String
    On Error GoTo ErrorHandler

    If IsError(cellValue) Or IsNull(cellValue) Or IsEmpty(cellValue) Then
        CellValueToText = vbNullString
    Else
        CellValueToText = Trim$(CStr(cellValue))
    End If

    Exit Function

ErrorHandler:
    CellValueToText = vbNullString
End Function
