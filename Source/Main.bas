Attribute VB_Name = "Main"
Option Explicit

Private Const DEFAULT_WORKSHEET_NAME As String = "Sheet1"
Private Const OUTPUT_FILE_PREFIX As String = "GeneratedDocument_"
Private Const DOCX_EXTENSION As String = "docx"
Private Const PDF_EXTENSION As String = "pdf"

' Coordinates the certificate generation workflow by calling the approved engine modules.
' Returns True when configuration, data loading, document creation, saving, and export complete successfully.
Public Function GenerateCertificates() As Boolean
    Dim records As Collection
    Dim record As Variant
    Dim generatedDocument As Document
    Dim recordIndex As Long
    Dim outputBaseName As String
    Dim docxPath As String
    Dim pdfPath As String
    Dim workbookOpened As Boolean

    On Error GoTo ErrorHandler

    If Not InitializeConfig() Then
        GenerateCertificates = False
        Exit Function
    End If

    If Not ValidateConfiguration() Then
        GenerateCertificates = False
        Exit Function
    End If

    workbookOpened = OpenWorkbook(ExcelFilePath)

    If Not workbookOpened Then
        GenerateCertificates = False
        Exit Function
    End If

    Set records = ReadAllRows(DEFAULT_WORKSHEET_NAME)

    For Each record In records
        recordIndex = recordIndex + 1
        outputBaseName = BuildOutputBaseName(recordIndex)
        docxPath = BuildOutputFilePath(outputBaseName, DOCX_EXTENSION)
        pdfPath = BuildOutputFilePath(outputBaseName, PDF_EXTENSION)

        Set generatedDocument = OpenTemplate(TemplateFilePath)

        If generatedDocument Is Nothing Then
            GenerateCertificates = False
            GoTo Cleanup
        End If

        If Not ReplacePlaceholders(generatedDocument, record) Then
            GenerateCertificates = False
            GoTo Cleanup
        End If

        If Not SaveDocument(generatedDocument, docxPath) Then
            GenerateCertificates = False
            GoTo Cleanup
        End If

        If Not ExportToPDF(generatedDocument, pdfPath) Then
            GenerateCertificates = False
            GoTo Cleanup
        End If

        CloseDocument generatedDocument, False
        Set generatedDocument = Nothing
    Next record

    GenerateCertificates = True

Cleanup:
    If Not generatedDocument Is Nothing Then
        CloseDocument generatedDocument, False
        Set generatedDocument = Nothing
    End If

    If workbookOpened Then
        CloseWorkbook
    End If

    Exit Function

ErrorHandler:
    GenerateCertificates = False
    Resume Cleanup
End Function

Private Function BuildOutputBaseName(ByVal recordIndex As Long) As String
    BuildOutputBaseName = OUTPUT_FILE_PREFIX & Format$(recordIndex, "0000")
End Function
