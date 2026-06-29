Attribute VB_Name = "PDFExporter"
Option Explicit

Private Const WORD_EXPORT_FORMAT_PDF As Long = 17
Private Const WORD_EXPORT_OPTIMIZE_FOR_PRINT As Long = 0
Private Const WORD_EXPORT_ALL_DOCUMENT As Long = 0
Private Const WORD_EXPORT_DOCUMENT_CONTENT As Long = 0
Private Const WINDOWS_PATH_SEPARATOR As String = "\"

' Exports the supplied Word document to a PDF file without closing the document.
' Returns True when the export completes successfully and False when validation or export fails.
Public Function ExportToPDF(ByVal document As Document, ByVal pdfPath As String) As Boolean
    Dim destinationFolder As String

    On Error GoTo ErrorHandler

    pdfPath = Trim$(pdfPath)

    If document Is Nothing Or Len(pdfPath) = 0 Then
        ExportToPDF = False
        Exit Function
    End If

    destinationFolder = GetParentFolder(pdfPath)

    If Len(destinationFolder) > 0 Then
        If Not CreateFolderIfMissing(destinationFolder) Then
            ExportToPDF = False
            Exit Function
        End If
    End If

    document.ExportAsFixedFormat _
        OutputFileName:=pdfPath, _
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

    ExportToPDF = True
    Exit Function

ErrorHandler:
    ExportToPDF = False
End Function

Private Function GetParentFolder(ByVal filePath As String) As String
    Dim separatorPosition As Long

    On Error GoTo ErrorHandler

    filePath = Trim$(filePath)
    filePath = RemoveTrailingPathSeparator(filePath)
    separatorPosition = InStrRev(filePath, WINDOWS_PATH_SEPARATOR)

    If separatorPosition = 0 Then
        GetParentFolder = vbNullString
    ElseIf separatorPosition = 3 And Mid$(filePath, 2, 1) = ":" Then
        GetParentFolder = Left$(filePath, separatorPosition)
    Else
        GetParentFolder = Left$(filePath, separatorPosition - 1)
    End If

    Exit Function

ErrorHandler:
    GetParentFolder = vbNullString
End Function

Private Function RemoveTrailingPathSeparator(ByVal pathValue As String) As String
    On Error GoTo ErrorHandler

    Do While Len(pathValue) > 1 _
            And Right$(pathValue, 1) = WINDOWS_PATH_SEPARATOR _
            And Not IsDriveRoot(pathValue)
        pathValue = Left$(pathValue, Len(pathValue) - 1)
    Loop

    RemoveTrailingPathSeparator = pathValue
    Exit Function

ErrorHandler:
    RemoveTrailingPathSeparator = vbNullString
End Function

Private Function IsDriveRoot(ByVal pathValue As String) As Boolean
    IsDriveRoot = (Len(pathValue) = 3 _
        And Mid$(pathValue, 2, 1) = ":" _
        And Right$(pathValue, 1) = WINDOWS_PATH_SEPARATOR)
End Function
