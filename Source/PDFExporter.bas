Attribute VB_Name = "PDFExporter"
Option Explicit

Private Const WORD_EXPORT_FORMAT_PDF As Long = 17
Private Const WORD_EXPORT_OPTIMIZE_FOR_PRINT As Long = 0
Private Const WORD_EXPORT_ALL_DOCUMENT As Long = 0
Private Const WORD_EXPORT_DOCUMENT_CONTENT As Long = 0
Private Const PDF_EXTENSION As String = ".pdf"

' Exports the supplied Word document to a PDF file without closing the document.
Public Function ExportDocumentToPDF(ByVal Doc As Object, ByVal OutputPDFPath As String) As Boolean
    Dim safeOutputPath As String
    Dim destinationFolder As String

    On Error GoTo ErrorHandler

    OutputPDFPath = Trim$(OutputPDFPath)

    If Doc Is Nothing Then
        WriteLog "PDF export skipped because document is not available."
        ExportDocumentToPDF = False
        Exit Function
    End If

    If Len(OutputPDFPath) = 0 Then
        WriteLog "PDF export skipped because output path is empty."
        ExportDocumentToPDF = False
        Exit Function
    End If

    safeOutputPath = BuildSafePDFPath(OutputPDFPath)
    destinationFolder = GetParentFolder(safeOutputPath)

    If Len(destinationFolder) > 0 Then
        If Not EnsureFolderExists(destinationFolder) Then
            WriteLog "PDF export folder could not be created: " & destinationFolder
            ExportDocumentToPDF = False
            Exit Function
        End If
    End If

    Doc.ExportAsFixedFormat _
        OutputFileName:=safeOutputPath, _
        ExportFormat:=WORD_EXPORT_FORMAT_PDF, _
        OpenAfterExport:=False, _
        OptimizeFor:=WORD_EXPORT_OPTIMIZE_FOR_PRINT, _
        Range:=WORD_EXPORT_ALL_DOCUMENT, _
        From:=1, _
        To:=1, _
        Item:=WORD_EXPORT_DOCUMENT_CONTENT, _
        IncludeDocProps:=True, _
        KeepIRM:=True, _
        CreateBookmarks:=0, _
        DocStructureTags:=True, _
        BitmapMissingFonts:=True, _
        UseISO19005_1:=False

    WriteLog "PDF exported: " & safeOutputPath
    ExportDocumentToPDF = True
    Exit Function

ErrorHandler:
    WriteErrorLog "PDFExporter.ExportDocumentToPDF", Err.Number, Err.Description
    ExportDocumentToPDF = False
End Function

' Compatibility wrapper for existing workflow code.
Public Function ExportToPDF(ByVal Doc As Object, ByVal PDFPath As String) As Boolean
    On Error GoTo ErrorHandler

    ExportToPDF = ExportDocumentToPDF(Doc, PDFPath)
    Exit Function

ErrorHandler:
    WriteErrorLog "PDFExporter.ExportToPDF", Err.Number, Err.Description
    ExportToPDF = False
End Function

Private Function BuildSafePDFPath(ByVal OutputPDFPath As String) As String
    Dim destinationFolder As String
    Dim fileName As String
    Dim baseName As String

    On Error GoTo ErrorHandler

    destinationFolder = GetParentFolder(OutputPDFPath)
    fileName = GetFileName(OutputPDFPath)

    If Len(fileName) = 0 Then
        fileName = "GeneratedDocument.pdf"
    End If

    If LCase$(Right$(fileName, Len(PDF_EXTENSION))) = PDF_EXTENSION Then
        baseName = Left$(fileName, Len(fileName) - Len(PDF_EXTENSION))
    Else
        baseName = fileName
    End If

    fileName = CleanFileName(baseName) & PDF_EXTENSION

    If Len(destinationFolder) = 0 Then
        BuildSafePDFPath = fileName
    Else
        BuildSafePDFPath = CombinePath(destinationFolder, fileName)
    End If

    Exit Function

ErrorHandler:
    WriteErrorLog "PDFExporter.BuildSafePDFPath", Err.Number, Err.Description
    BuildSafePDFPath = OutputPDFPath
End Function

Private Function GetFileName(ByVal FilePath As String) As String
    Dim separatorPosition As Long

    On Error GoTo ErrorHandler

    FilePath = RemoveTrailingPathSeparator(Trim$(FilePath))
    separatorPosition = InStrRev(FilePath, "\")

    If separatorPosition = 0 Then
        GetFileName = FilePath
    Else
        GetFileName = Mid$(FilePath, separatorPosition + 1)
    End If

    Exit Function

ErrorHandler:
    WriteErrorLog "PDFExporter.GetFileName", Err.Number, Err.Description
    GetFileName = vbNullString
End Function
