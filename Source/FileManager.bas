Attribute VB_Name = "FileManager"
Option Explicit

' Creates the specified folder when it does not already exist.
' Returns True when the folder exists or is created successfully.
Public Function CreateFolderIfMissing(ByVal folderPath As String) As Boolean
    On Error GoTo ErrorHandler

    CreateFolderIfMissing = EnsureFolderExists(folderPath)
    Exit Function

ErrorHandler:
    CreateFolderIfMissing = False
End Function

' Ensures the configured output folder exists.
' Uses CurrentOutputFolder from the Config module.
Public Function EnsureOutputFolder() As Boolean
    On Error GoTo ErrorHandler

    EnsureOutputFolder = CreateFolderIfMissing(CurrentOutputFolder)
    Exit Function

ErrorHandler:
    EnsureOutputFolder = False
End Function

' Ensures the configured log folder exists.
' Uses CurrentLogFolder from the Config module.
Public Function EnsureLogFolder() As Boolean
    On Error GoTo ErrorHandler

    EnsureLogFolder = CreateFolderIfMissing(CurrentLogFolder)
    Exit Function

ErrorHandler:
    EnsureLogFolder = False
End Function

' Builds a full output file path using the configured output folder.
' The file name is sanitized and the extension is normalized before the path is returned.
Public Function BuildOutputFilePath(ByVal fileName As String, ByVal extension As String) As String
    On Error GoTo ErrorHandler

    fileName = SanitizeFileName(fileName)
    extension = NormalizeExtension(extension)

    If Len(CurrentOutputFolder) = 0 Or Len(fileName) = 0 Or Len(extension) = 0 Then
        BuildOutputFilePath = vbNullString
        Exit Function
    End If

    BuildOutputFilePath = CombinePath(CurrentOutputFolder, fileName & "." & extension)
    Exit Function

ErrorHandler:
    BuildOutputFilePath = vbNullString
End Function

' Replaces characters that are invalid in Windows file names with underscores.
' Leading and trailing whitespace is removed from the returned value.
Public Function SanitizeFileName(ByVal fileName As String) As String
    On Error GoTo ErrorHandler

    SanitizeFileName = CleanFileName(fileName)
    Exit Function

ErrorHandler:
    SanitizeFileName = vbNullString
End Function

Private Function NormalizeExtension(ByVal extension As String) As String
    On Error GoTo ErrorHandler

    extension = Trim$(extension)

    Do While Left$(extension, 1) = "."
        extension = Mid$(extension, 2)
    Loop

    NormalizeExtension = extension
    Exit Function

ErrorHandler:
    NormalizeExtension = vbNullString
End Function
