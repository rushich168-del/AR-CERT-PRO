Attribute VB_Name = "WordEngine"
Option Explicit

Private Const WORD_FORMAT_DOCUMENT_DEFAULT As Long = 16

' Opens a Word template as a new editable document.
' Returns Nothing when the template path is invalid or the document cannot be created.
Public Function OpenTemplate(ByVal templatePath As String) As Document
    On Error GoTo ErrorHandler

    templatePath = Trim$(templatePath)

    If Len(templatePath) = 0 Or Not LocalFileExists(templatePath) Then
        Set OpenTemplate = Nothing
        Exit Function
    End If

    Set OpenTemplate = Application.Documents.Add(Template:=templatePath, NewTemplate:=False, DocumentType:=0)
    Exit Function

ErrorHandler:
    Set OpenTemplate = Nothing
End Function

' Saves the supplied document as a DOCX file.
' Returns True when the document is valid and the save operation completes successfully.
Public Function SaveDocument(ByVal document As Document, ByVal outputPath As String) As Boolean
    On Error GoTo ErrorHandler

    SaveDocument = SaveDocumentAs(document, outputPath, WORD_FORMAT_DOCUMENT_DEFAULT)
    Exit Function

ErrorHandler:
    SaveDocument = False
End Function

' Saves the supplied document using the requested Word file format.
' This generic routine is intended for reuse by future Office automation modules.
Public Function SaveDocumentAs(ByVal document As Document, ByVal filePath As String, ByVal fileFormat As Long) As Boolean
    On Error GoTo ErrorHandler

    filePath = Trim$(filePath)

    If Not DocumentExists(document) Or Len(filePath) = 0 Then
        SaveDocumentAs = False
        Exit Function
    End If

    document.SaveAs2 FileName:=filePath, FileFormat:=fileFormat
    SaveDocumentAs = True
    Exit Function

ErrorHandler:
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

    If saveChanges Then
        closeOption = wdSaveChanges
    Else
        closeOption = wdDoNotSaveChanges
    End If

    document.Close SaveChanges:=closeOption
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

Private Function LocalFileExists(ByVal filePath As String) As Boolean
    Dim attributes As Long

    On Error GoTo ErrorHandler

    filePath = Trim$(filePath)

    If Len(filePath) = 0 Then
        LocalFileExists = False
        Exit Function
    End If

    attributes = GetAttr(filePath)
    LocalFileExists = ((attributes And vbDirectory) = 0)
    Exit Function

ErrorHandler:
    LocalFileExists = False
End Function
