Attribute VB_Name = "Main"
Option Explicit

Private Const OUTPUT_FILE_PREFIX As String = "GeneratedDocument_"
Private Const DOCX_EXTENSION As String = "docx"
Private Const PDF_EXTENSION As String = "pdf"
Private Const HEADER_ROW_NUMBER As Long = 1
Private Const FIRST_DATA_ROW_NUMBER As Long = 2

' Coordinates batch certificate generation from Excel rows to Word and PDF outputs.
Public Function GenerateCertificates() As Boolean
    Dim worksheet As Object
    Dim headers As Collection
    Dim lastColumn As Long
    Dim lastRow As Long
    Dim rowNumber As Long
    Dim totalRows As Long
    Dim successCount As Long
    Dim failureCount As Long
    Dim workbookOpened As Boolean

    On Error GoTo ErrorHandler

    If Not InitializeConfig() Then
        GenerateCertificates = False
        Exit Function
    End If

    If Not InitLog() Then
        GenerateCertificates = False
        Exit Function
    End If

    WriteLog "Batch certificate generation started."

    If Not ValidateConfiguration() Then
        WriteLog "Configuration validation failed: " & LastConfigError
        GenerateCertificates = False
        GoTo Cleanup
    End If

    If Not EnsureOutputFolder() Then
        WriteLog "Output folder could not be created: " & CurrentOutputFolder
        GenerateCertificates = False
        GoTo Cleanup
    End If

    workbookOpened = OpenWorkbook(ExcelFilePath)

    If Not workbookOpened Then
        WriteLog "Excel workbook could not be opened: " & ExcelFilePath
        GenerateCertificates = False
        GoTo Cleanup
    End If

    Set worksheet = GetWorksheet(GetConfiguredWorksheetName())

    If worksheet Is Nothing Then
        WriteLog "Worksheet could not be selected: " & GetConfiguredWorksheetName()
        GenerateCertificates = False
        GoTo Cleanup
    End If

    lastColumn = GetLastColumn(HEADER_ROW_NUMBER)

    If lastColumn = 0 Then
        WriteLog "No headers found in row 1."
        GenerateCertificates = False
        GoTo Cleanup
    End If

    Set headers = GetHeaders(GetConfiguredWorksheetName())
    lastRow = GetLastDataRow(lastColumn)

    If lastRow < FIRST_DATA_ROW_NUMBER Then
        WriteLog "No data rows found for certificate generation."
        GenerateCertificates = False
        GoTo Cleanup
    End If

    For rowNumber = FIRST_DATA_ROW_NUMBER To lastRow
        If IsDataRowEmpty(rowNumber, lastColumn) Then
            WriteLog "Row " & CStr(rowNumber) & " skipped because it is empty."
        Else
            totalRows = totalRows + 1

            If ProcessCertificateRow(rowNumber, headers, lastColumn) Then
                successCount = successCount + 1
                WriteLog "Row " & CStr(rowNumber) & " completed successfully."
            Else
                failureCount = failureCount + 1
                WriteLog "Row " & CStr(rowNumber) & " failed."
            End If
        End If
    Next rowNumber

    WriteBatchSummary totalRows, successCount, failureCount
    GenerateCertificates = (totalRows > 0 And failureCount = 0)

Cleanup:
    CloseGeneratedDocument

    If workbookOpened Then
        CloseExcel
    End If

    Exit Function

ErrorHandler:
    WriteErrorLog "Main.GenerateCertificates", Err.Number, Err.Description
    GenerateCertificates = False
    Resume Cleanup
End Function

Private Function ProcessCertificateRow(ByVal RowNumber As Long, ByVal Headers As Collection, ByVal LastColumn As Long) As Boolean
    Dim rowValues As Variant
    Dim generatedDocument As Object
    Dim outputBaseName As String
    Dim docxPath As String
    Dim pdfPath As String

    On Error GoTo ErrorHandler

    rowValues = BuildRowValues(RowNumber, LastColumn)

    If IsEmpty(rowValues) Then
        WriteLog "Row " & CStr(RowNumber) & " has no readable values."
        ProcessCertificateRow = False
        GoTo Cleanup
    End If

    outputBaseName = BuildOutputBaseName(RowNumber, Headers, rowValues)
    docxPath = BuildOutputFilePath(outputBaseName, DOCX_EXTENSION)
    pdfPath = BuildOutputFilePath(outputBaseName, PDF_EXTENSION)

    If Len(docxPath) = 0 Or Len(pdfPath) = 0 Then
        WriteLog "Row " & CStr(RowNumber) & " output paths could not be created."
        ProcessCertificateRow = False
        GoTo Cleanup
    End If

    If Not OpenTemplate(TemplateFilePath) Then
        WriteLog "Row " & CStr(RowNumber) & " failed to open template."
        ProcessCertificateRow = False
        GoTo Cleanup
    End If

    Set generatedDocument = GetGeneratedDocument()

    If generatedDocument Is Nothing Then
        WriteLog "Row " & CStr(RowNumber) & " has no generated Word document."
        ProcessCertificateRow = False
        GoTo Cleanup
    End If

    If Not ReplaceAllPlaceholders(generatedDocument, Headers, rowValues) Then
        WriteLog "Row " & CStr(RowNumber) & " placeholder replacement failed."
        ProcessCertificateRow = False
        GoTo Cleanup
    End If

    If Not SaveGeneratedDocument(docxPath) Then
        WriteLog "Row " & CStr(RowNumber) & " Word document save failed."
        ProcessCertificateRow = False
        GoTo Cleanup
    End If

    If Not ExportDocumentToPDF(generatedDocument, pdfPath) Then
        WriteLog "Row " & CStr(RowNumber) & " PDF export failed."
        ProcessCertificateRow = False
        GoTo Cleanup
    End If

    ProcessCertificateRow = True

Cleanup:
    CloseGeneratedDocument
    Set generatedDocument = Nothing
    Exit Function

ErrorHandler:
    WriteErrorLog "Main.ProcessCertificateRow", Err.Number, "Row " & CStr(RowNumber) & ": " & Err.Description
    ProcessCertificateRow = False
    Resume Cleanup
End Function

Private Function GetConfiguredWorksheetName() As String
    GetConfiguredWorksheetName = DEFAULT_WORKSHEET_NAME
End Function

Private Function GetLastDataRow(ByVal LastColumn As Long) As Long
    Dim columnNumber As Long
    Dim currentLastRow As Long
    Dim maxLastRow As Long

    On Error GoTo ErrorHandler

    For columnNumber = 1 To LastColumn
        currentLastRow = GetLastRow(columnNumber)

        If currentLastRow > maxLastRow Then
            maxLastRow = currentLastRow
        End If
    Next columnNumber

    GetLastDataRow = maxLastRow
    Exit Function

ErrorHandler:
    WriteErrorLog "Main.GetLastDataRow", Err.Number, Err.Description
    GetLastDataRow = 0
End Function

Private Function BuildRowValues(ByVal RowNumber As Long, ByVal LastColumn As Long) As Variant
    Dim values() As Variant
    Dim columnNumber As Long

    On Error GoTo ErrorHandler

    ReDim values(1 To 1, 1 To LastColumn)

    For columnNumber = 1 To LastColumn
        values(1, columnNumber) = ReadCell(RowNumber, columnNumber)
    Next columnNumber

    BuildRowValues = values
    Exit Function

ErrorHandler:
    WriteErrorLog "Main.BuildRowValues", Err.Number, Err.Description
    BuildRowValues = Empty
End Function

Private Function IsDataRowEmpty(ByVal RowNumber As Long, ByVal LastColumn As Long) As Boolean
    Dim columnNumber As Long

    On Error GoTo ErrorHandler

    For columnNumber = 1 To LastColumn
        If Len(Trim$(ValueToText(ReadCell(RowNumber, columnNumber)))) > 0 Then
            IsDataRowEmpty = False
            Exit Function
        End If
    Next columnNumber

    IsDataRowEmpty = True
    Exit Function

ErrorHandler:
    WriteErrorLog "Main.IsDataRowEmpty", Err.Number, Err.Description
    IsDataRowEmpty = True
End Function

Private Function BuildOutputBaseName(ByVal RowNumber As Long, ByVal Headers As Collection, ByVal RowValues As Variant) As String
    Dim nameValue As String

    On Error GoTo ErrorHandler

    nameValue = GetNameFieldValue(Headers, RowValues)

    If Len(nameValue) = 0 Then
        BuildOutputBaseName = OUTPUT_FILE_PREFIX & "Row_" & Format$(RowNumber, "0000")
    Else
        BuildOutputBaseName = CleanFileName(nameValue & "_" & Format$(RowNumber, "0000"))
    End If

    Exit Function

ErrorHandler:
    WriteErrorLog "Main.BuildOutputBaseName", Err.Number, Err.Description
    BuildOutputBaseName = OUTPUT_FILE_PREFIX & "Row_" & Format$(RowNumber, "0000")
End Function

Private Function GetNameFieldValue(ByVal Headers As Collection, ByVal RowValues As Variant) As String
    Dim columnIndex As Long
    Dim headerText As String

    On Error GoTo ErrorHandler

    For columnIndex = 1 To Headers.Count
        headerText = LCase$(Trim$(CStr(Headers(columnIndex))))

        If IsNameHeader(headerText) Then
            GetNameFieldValue = CleanFileName(ValueToText(GetRowValue(RowValues, columnIndex)))
            Exit Function
        End If
    Next columnIndex

    GetNameFieldValue = vbNullString
    Exit Function

ErrorHandler:
    WriteErrorLog "Main.GetNameFieldValue", Err.Number, Err.Description
    GetNameFieldValue = vbNullString
End Function

Private Function IsNameHeader(ByVal HeaderText As String) As Boolean
    HeaderText = Replace$(HeaderText, "_", " ")
    HeaderText = Replace$(HeaderText, "-", " ")

    IsNameHeader = (HeaderText = "name" _
        Or HeaderText = "student" _
        Or HeaderText = "student name" _
        Or HeaderText = "full name" _
        Or HeaderText = "recipient name" _
        Or HeaderText = "candidate name" _
        Or (InStr(1, HeaderText, "student", vbTextCompare) > 0 _
            And InStr(1, HeaderText, "name", vbTextCompare) > 0))
End Function

Private Function GetRowValue(ByVal RowValues As Variant, ByVal ColumnIndex As Long) As Variant
    On Error GoTo TryOneDimensionalArray

    If IsArray(RowValues) Then
        GetRowValue = RowValues(LBound(RowValues, 1), LBound(RowValues, 2) + ColumnIndex - 1)
        Exit Function
    End If

TryOneDimensionalArray:
    On Error GoTo ErrorHandler

    If IsArray(RowValues) Then
        GetRowValue = RowValues(LBound(RowValues) + ColumnIndex - 1)
        Exit Function
    End If

ErrorHandler:
    GetRowValue = Empty
End Function

Private Function ValueToText(ByVal Value As Variant) As String
    On Error GoTo ErrorHandler

    If IsError(Value) Or IsNull(Value) Or IsEmpty(Value) Then
        ValueToText = vbNullString
    Else
        ValueToText = Trim$(CStr(Value))
    End If

    Exit Function

ErrorHandler:
    ValueToText = vbNullString
End Function

Private Sub WriteBatchSummary(ByVal TotalRows As Long, ByVal SuccessCount As Long, ByVal FailureCount As Long)
    On Error Resume Next

    WriteLog "Batch certificate generation summary:"
    WriteLog "Total rows: " & CStr(TotalRows)
    WriteLog "Success count: " & CStr(SuccessCount)
    WriteLog "Failure count: " & CStr(FailureCount)
End Sub
