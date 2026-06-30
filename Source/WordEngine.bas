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

' Saves the active generated document as a DOCX file.
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

' Closes the active generated document without saving pending changes.
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

' Creates a duplicate editable document from the supplied document.
Public Function DuplicateDocument(ByVal SourceDocument As Document) As Document
    Dim newDocument As Document
    Dim sourceRange As Range

    On Error GoTo ErrorHandler

    If Not DocumentExists(SourceDocument) Then
        Set newDocument = Nothing
        Exit Function
    End If

    Set sourceRange = SourceDocument.Content
    sourceRange.Copy

    Set newDocument = Application.Documents.Add
    newDocument.Range.PasteAndFormat wdFormatOriginalFormatting

    Set DuplicateDocument = newDocument
    Exit Function

ErrorHandler:
    If Not newDocument Is Nothing Then
        CloseDocumentInternal newDocument, False
    End If

    WriteErrorLog "WordEngine.DuplicateDocument", Err.Number, Err.Description
    Set DuplicateDocument = Nothing
End Function

' Returns True when the supplied Word document object is valid and accessible.
Public Function DocumentExists(ByVal TargetDocument As Document) As Boolean
    Dim documentName As String

    On Error GoTo ErrorHandler

    If TargetDocument Is Nothing Then
        DocumentExists = False
        Exit Function
    End If

    documentName = TargetDocument.Name
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

Private Sub CloseDocumentInternal(ByVal TargetDocument As Document, ByVal SaveChanges As Boolean)
    Dim closeOption As Long

    On Error Resume Next

    If TargetDocument Is Nothing Then
        Exit Sub
    End If

    If SaveChanges Then
        closeOption = wdSaveChanges
    Else
        closeOption = wdDoNotSaveChanges
    End If

    TargetDocument.Close SaveChanges:=closeOption
End Sub