Attribute VB_Name = "WordEngine"
Option Explicit

Private Const WORD_FORMAT_DOCUMENT_DEFAULT As Long = 16

Private mGeneratedDocument As Document
Private mTemplatePath As String

' Opens a Word template as a new editable document without modifying the original template.
Public Function OpenTemplate(ByVal TemplatePath As String) As Boolean
    On Error GoTo ErrorHandler

    TemplatePath = ResolveTemplatePath(TemplatePath)

    If Len(TemplatePath) = 0 Then
        WriteLog "Template path is empty."
        OpenTemplate = False
        Exit Function
    End If

    If Not FileExists(TemplatePath) Then
        WriteLog "Template not found: " & TemplatePath
        OpenTemplate = False
        Exit Function
    End If

    CloseGeneratedDocument

    Set mGeneratedDocument = Application.Documents.Add(Template:=TemplatePath, NewTemplate:=False, DocumentType:=0)
    mTemplatePath = TemplatePath

    OpenTemplate = DocumentExists(mGeneratedDocument)

    If OpenTemplate Then
        WriteLog "Template opened as generated document copy: " & TemplatePath
    Else
        WriteLog "Template could not be opened: " & TemplatePath
    End If

    Exit Function

ErrorHandler:
    WriteErrorLog "WordEngine.OpenTemplate", Err.Number, Err.Description
    CloseGeneratedDocument
    OpenTemplate = False
End Function

' Saves the generated document as a DOCX file.
Public Function SaveGeneratedDocument(ByVal OutputPath As String) As Boolean
    Dim outputFolder As String

    On Error GoTo ErrorHandler

    OutputPath = Trim$(OutputPath)

    If Not DocumentExists(mGeneratedDocument) Then
        WriteLog "No generated document is open to save."
        SaveGeneratedDocument = False
        Exit Function
    End If

    If Len(OutputPath) = 0 Then
        WriteLog "Generated document output path is empty."
        SaveGeneratedDocument = False
        Exit Function
    End If

    outputFolder = GetParentFolder(OutputPath)

    If Len(outputFolder) > 0 Then
        If Not EnsureFolderExists(outputFolder) Then
            WriteLog "Unable to create generated document output folder: " & outputFolder
            SaveGeneratedDocument = False
            Exit Function
        End If
    End If

    mGeneratedDocument.SaveAs2 FileName:=OutputPath, FileFormat:=WORD_FORMAT_DOCUMENT_DEFAULT
    WriteLog "Generated document saved: " & OutputPath

    SaveGeneratedDocument = True
    Exit Function

ErrorHandler:
    WriteErrorLog "WordEngine.SaveGeneratedDocument", Err.Number, Err.Description
    SaveGeneratedDocument = False
End Function

' Closes the generated document without saving pending changes.
Public Sub CloseGeneratedDocument()
    On Error Resume Next

    If Not mGeneratedDocument Is Nothing Then
        mGeneratedDocument.Close SaveChanges:=wdDoNotSaveChanges
        WriteLog "Generated document closed."
    End If

    Set mGeneratedDocument = Nothing
    mTemplatePath = vbNullString
End Sub

' Returns the active generated document object.
Public Function GetGeneratedDocument() As Object
    On Error GoTo ErrorHandler

    If DocumentExists(mGeneratedDocument) Then
        Set GetGeneratedDocument = mGeneratedDocument
    Else
        Set GetGeneratedDocument = Nothing
    End If

    Exit Function

ErrorHandler:
    WriteErrorLog "WordEngine.GetGeneratedDocument", Err.Number, Err.Description
    Set GetGeneratedDocument = Nothing
End Function

' Saves the supplied document as a DOCX file.
' Returns True when the document is valid and the save operation completes successfully.
Public Function SaveDocument(ByVal document As Document, ByVal outputPath As String) As Boolean
    On Error GoTo ErrorHandler

    If document Is mGeneratedDocument Then
        SaveDocument = SaveGeneratedDocument(outputPath)
        Exit Function
    End If

    SaveDocument = SaveDocumentAs(document, outputPath, WORD_FORMAT_DOCUMENT_DEFAULT)
    Exit Function

ErrorHandler:
    WriteErrorLog "WordEngine.SaveDocument", Err.Number, Err.Description
    SaveDocument = False
End Function

' Saves the supplied document using the requested Word file format.
' This generic routine is intended for reuse by future Office automation modules.
Public Function SaveDocumentAs(ByVal document As Document, ByVal filePath As String, ByVal fileFormat As Long) As Boolean
    On Error GoTo ErrorHandler

    filePath = Trim$(filePath)

    If Not DocumentExists(document) Or Len(filePath) = 0 Then
        WriteLog "Document save skipped because document or path is invalid."
        SaveDocumentAs = False
        Exit Function
    End If

    If Len(GetParentFolder(filePath)) > 0 Then
        If Not EnsureFolderExists(GetParentFolder(filePath)) Then
            WriteLog "Unable to create document output folder: " & GetParentFolder(filePath)
            SaveDocumentAs = False
            Exit Function
        End If
    End If

    document.SaveAs2 FileName:=filePath, FileFormat:=fileFormat
    WriteLog "Document saved: " & filePath
    SaveDocumentAs = True
    Exit Function

ErrorHandler:
    WriteErrorLog "WordEngine.SaveDocumentAs", Err.Number, Err.Description
    SaveDocumentAs = False
End Function

' Safely closes the supplied document.
' The saveChanges parameter controls whether Word saves pending document changes.
Public Sub CloseDocument(ByVal document As Document, ByVal saveChanges As Boolean)
    Dim closeOption As Long

    On Error Resume Next

    If document Is Nothing Then
        Exit Sub
    End If

    If document Is mGeneratedDocument Then
        CloseGeneratedDocument
        Exit Sub
    End If

    If saveChanges Then
        closeOption = wdSaveChanges
    Else
        closeOption = wdDoNotSaveChanges
    End If

    document.Close SaveChanges:=closeOption
    WriteLog "Document closed."
End Sub

' Creates a duplicate editable document from the supplied document.
' Returns Nothing when the source document is not valid or cannot be copied.
Public Function DuplicateDocument(ByVal document As Document) As Document
    Dim duplicateDocument As Document
    Dim sourceRange As Range

    On Error GoTo ErrorHandler

    If Not DocumentExists(document) Then
        Set DuplicateDocument = Nothing
        Exit Function
    End If

    Set sourceRange = document.Content
    sourceRange.Copy

    Set duplicateDocument = Application.Documents.Add
    duplicateDocument.Range.PasteAndFormat wdFormatOriginalFormatting

    Set DuplicateDocument = duplicateDocument
    Exit Function

ErrorHandler:
    If Not duplicateDocument Is Nothing Then
        CloseDocument duplicateDocument, False
    End If

    WriteErrorLog "WordEngine.DuplicateDocument", Err.Number, Err.Description
    Set DuplicateDocument = Nothing
End Function

' Returns True when the supplied Word document object is valid and accessible.
Public Function DocumentExists(ByVal document As Document) As Boolean
    Dim documentName As String

    On Error GoTo ErrorHandler

    If document Is Nothing Then
        DocumentExists = False
        Exit Function
    End If

    documentName = document.Name
    DocumentExists = (Len(documentName) > 0)
    Exit Function

ErrorHandler:
    DocumentExists = False
End Function

Private Function ResolveTemplatePath(ByVal TemplatePath As String) As String
    On Error GoTo ErrorHandler

    TemplatePath = Trim$(TemplatePath)

    If Len(TemplatePath) = 0 Then
        If Len(TemplateFilePath) = 0 Then
            If Not InitializeConfig() Then
                ResolveTemplatePath = vbNullString
                Exit Function
            End If
        End If

        TemplatePath = TemplateFilePath
    End If

    ResolveTemplatePath = TemplatePath
    Exit Function

ErrorHandler:
    WriteErrorLog "WordEngine.ResolveTemplatePath", Err.Number, Err.Description
    ResolveTemplatePath = vbNullString
End Function
