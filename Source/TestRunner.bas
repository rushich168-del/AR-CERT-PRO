Attribute VB_Name = "TestRunner"
Option Explicit

Private Const TEST_DOCX_NAME As String = "TestRunner_Output.docx"
Private Const TEST_PDF_NAME As String = "TestRunner_Output.pdf"

Private mTestLogReady As Boolean

Public Sub TestConfigPaths()
    On Error GoTo ErrorHandler

    EnsureTestLog

    If InitializeConfig() _
            And Len(ProjectFolder) > 0 _
            And Len(ExcelFilePath) > 0 _
            And Len(TemplateFilePath) > 0 _
            And Len(CurrentOutputFolder) > 0 _
            And Len(CurrentLogFolder) > 0 Then
        LogTestResult "TestConfigPaths", True, "Configuration paths initialized."
    Else
        LogTestResult "TestConfigPaths", False, LastConfigError
    End If

    Exit Sub

ErrorHandler:
    LogTestResult "TestConfigPaths", False, Err.Description
End Sub

Public Sub TestLogger()
    On Error GoTo ErrorHandler

    EnsureTestLog

    If mTestLogReady Then
        WriteLog "TestLogger wrote a test log entry."
        LogTestResult "TestLogger", True, "Logger initialized and wrote an entry."
    Else
        LogTestResult "TestLogger", False, "Logger is not ready."
    End If

    Exit Sub

ErrorHandler:
    LogTestResult "TestLogger", False, Err.Description
End Sub

Public Sub TestExcelEngine()
    Dim worksheet As Object
    Dim headerCount As Long

    On Error GoTo ErrorHandler

    EnsureTestLog

    If Not InitializeConfig() Then
        LogTestResult "TestExcelEngine", False, LastConfigError
        Exit Sub
    End If

    If Not FileExists(ExcelFilePath) Then
        LogTestResult "TestExcelEngine", False, "Missing Excel workbook: " & ExcelFilePath
        Exit Sub
    End If

    If Not OpenWorkbook(ExcelFilePath) Then
        LogTestResult "TestExcelEngine", False, "OpenWorkbook returned False."
        GoTo Cleanup
    End If

    Set worksheet = GetWorksheet(DEFAULT_WORKSHEET_NAME)

    If worksheet Is Nothing Then
        LogTestResult "TestExcelEngine", False, "Worksheet not found: " & DEFAULT_WORKSHEET_NAME
        GoTo Cleanup
    End If

    headerCount = GetLastColumn(1)

    If headerCount > 0 Then
        LogTestResult "TestExcelEngine", True, "Workbook opened and headers detected: " & CStr(headerCount)
    Else
        LogTestResult "TestExcelEngine", False, "No headers found in row 1."
    End If

Cleanup:
    CloseExcel
    Exit Sub

ErrorHandler:
    LogTestResult "TestExcelEngine", False, Err.Description
    Resume Cleanup
End Sub

Public Sub TestWordEngine()
    Dim generatedDocument As Object
    Dim outputPath As String

    On Error GoTo ErrorHandler

    EnsureTestLog

    If Not InitializeConfig() Then
        LogTestResult "TestWordEngine", False, LastConfigError
        Exit Sub
    End If

    If Not FileExists(TemplateFilePath) Then
        LogTestResult "TestWordEngine", False, "Missing Word template: " & TemplateFilePath
        Exit Sub
    End If

    If Not OpenTemplate(TemplateFilePath) Then
        LogTestResult "TestWordEngine", False, "OpenTemplate returned False."
        GoTo Cleanup
    End If

    Set generatedDocument = GetGeneratedDocument()

    If generatedDocument Is Nothing Then
        LogTestResult "TestWordEngine", False, "GetGeneratedDocument returned Nothing."
        GoTo Cleanup
    End If

    outputPath = CombinePath(CurrentOutputFolder, TEST_DOCX_NAME)

    If SaveGeneratedDocument(outputPath) And FileExists(outputPath) Then
        LogTestResult "TestWordEngine", True, "Generated document saved: " & outputPath
    Else
        LogTestResult "TestWordEngine", False, "Generated document was not saved."
    End If

Cleanup:
    CloseGeneratedDocument
    Exit Sub

ErrorHandler:
    LogTestResult "TestWordEngine", False, Err.Description
    Resume Cleanup
End Sub

Public Sub TestPlaceholderEngine()
    Dim generatedDocument As Object

    On Error GoTo ErrorHandler

    EnsureTestLog

    If Not InitializeConfig() Then
        LogTestResult "TestPlaceholderEngine", False, LastConfigError
        Exit Sub
    End If

    If Not FileExists(TemplateFilePath) Then
        LogTestResult "TestPlaceholderEngine", False, "Missing Word template: " & TemplateFilePath
        Exit Sub
    End If

    If Not OpenTemplate(TemplateFilePath) Then
        LogTestResult "TestPlaceholderEngine", False, "OpenTemplate returned False."
        GoTo Cleanup
    End If

    Set generatedDocument = GetGeneratedDocument()

    If generatedDocument Is Nothing Then
        LogTestResult "TestPlaceholderEngine", False, "GetGeneratedDocument returned Nothing."
        GoTo Cleanup
    End If

    If ReplacePlaceholder(generatedDocument, "Student Name", "Test Student") Then
        LogTestResult "TestPlaceholderEngine", True, "Placeholder replacement call completed."
    Else
        LogTestResult "TestPlaceholderEngine", False, "ReplacePlaceholder returned False."
    End If

Cleanup:
    CloseGeneratedDocument
    Exit Sub

ErrorHandler:
    LogTestResult "TestPlaceholderEngine", False, Err.Description
    Resume Cleanup
End Sub

Public Sub TestPDFExporter()
    Dim generatedDocument As Object
    Dim pdfPath As String

    On Error GoTo ErrorHandler

    EnsureTestLog

    If Not InitializeConfig() Then
        LogTestResult "TestPDFExporter", False, LastConfigError
        Exit Sub
    End If

    If Not FileExists(TemplateFilePath) Then
        LogTestResult "TestPDFExporter", False, "Missing Word template: " & TemplateFilePath
        Exit Sub
    End If

    If Not OpenTemplate(TemplateFilePath) Then
        LogTestResult "TestPDFExporter", False, "OpenTemplate returned False."
        GoTo Cleanup
    End If

    Set generatedDocument = GetGeneratedDocument()

    If generatedDocument Is Nothing Then
        LogTestResult "TestPDFExporter", False, "GetGeneratedDocument returned Nothing."
        GoTo Cleanup
    End If

    pdfPath = CombinePath(CurrentOutputFolder, TEST_PDF_NAME)

    If ExportDocumentToPDF(generatedDocument, pdfPath) And FileExists(pdfPath) Then
        LogTestResult "TestPDFExporter", True, "PDF exported: " & pdfPath
    Else
        LogTestResult "TestPDFExporter", False, "PDF was not exported."
    End If

Cleanup:
    CloseGeneratedDocument
    Exit Sub

ErrorHandler:
    LogTestResult "TestPDFExporter", False, Err.Description
    Resume Cleanup
End Sub

Public Sub TestCompleteWorkflow()
    On Error GoTo ErrorHandler

    EnsureTestLog
    WriteLog "TestCompleteWorkflow started."

    TestConfigPaths
    TestLogger
    TestExcelEngine
    TestWordEngine
    TestPlaceholderEngine
    TestPDFExporter

    If GenerateCertificates() Then
        LogTestResult "TestCompleteWorkflow", True, "GenerateCertificates completed successfully."
    Else
        LogTestResult "TestCompleteWorkflow", False, "GenerateCertificates returned False. Review earlier log entries."
    End If

    WriteLog "TestCompleteWorkflow finished."
    Exit Sub

ErrorHandler:
    LogTestResult "TestCompleteWorkflow", False, Err.Description
End Sub

Private Sub EnsureTestLog()
    On Error Resume Next

    If mTestLogReady Then
        Exit Sub
    End If

    If Len(ProjectFolder) = 0 Or Len(CurrentLogFolder) = 0 Then
        InitializeConfig
    End If

    mTestLogReady = InitLog()
End Sub

Private Sub LogTestResult(ByVal TestName As String, ByVal Passed As Boolean, ByVal Detail As String)
    On Error Resume Next

    If Passed Then
        WriteLog "PASS - " & TestName & " - " & Detail
    Else
        WriteLog "FAIL - " & TestName & " - " & Detail
    End If
End Sub
