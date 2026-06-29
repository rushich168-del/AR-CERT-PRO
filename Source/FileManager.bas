Attribute VB_Name = "FileManager"
Option Explicit

Private Const WINDOWS_PATH_SEPARATOR As String = "\"
Private Const DEFAULT_INVALID_CHARACTER_REPLACEMENT As String = "_"
Private Const DEFAULT_FILE_NAME As String = "untitled"

' Creates the specified folder when it does not already exist.
' Returns True when the folder exists or is created successfully.
Public Function CreateFolderIfMissing(ByVal folderPath As String) As Boolean
    On Error GoTo ErrorHandler

    folderPath = Trim$(folderPath)

    If Len(folderPath) = 0 Then
        CreateFolderIfMissing = False
        Exit Function
    End If

    If FolderExists(folderPath) Then
        CreateFolderIfMissing = True
        Exit Function
    End If

    CreateFolderHierarchy NormalizeFolderPath(folderPath)
    CreateFolderIfMissing = FolderExists(folderPath)
    Exit Function

ErrorHandler:
    CreateFolderIfMissing = False
End Function

' Returns True when the supplied path exists and points to a file.
Public Function FileExists(ByVal filePath As String) As Boolean
    Dim attributes As Long

    On Error GoTo ErrorHandler

    filePath = Trim$(filePath)

    If Len(filePath) = 0 Then
        FileExists = False
        Exit Function
    End If

    attributes = GetAttr(filePath)
    FileExists = ((attributes And vbDirectory) = 0)
    Exit Function

ErrorHandler:
    FileExists = False
End Function

' Returns True when the supplied path exists and points to a folder.
Public Function FolderExists(ByVal folderPath As String) As Boolean
    Dim attributes As Long

    On Error GoTo ErrorHandler

    folderPath = Trim$(folderPath)

    If Len(folderPath) = 0 Then
        FolderExists = False
        Exit Function
    End If

    attributes = GetAttr(folderPath)
    FolderExists = ((attributes And vbDirectory) = vbDirectory)
    Exit Function

ErrorHandler:
    FolderExists = False
End Function

' Ensures the configured output folder exists.
' Uses OutputFolder from the Config module.
Public Function EnsureOutputFolder() As Boolean
    On Error GoTo ErrorHandler

    EnsureOutputFolder = CreateFolderIfMissing(OutputFolder)
    Exit Function

ErrorHandler:
    EnsureOutputFolder = False
End Function

' Ensures the configured log folder exists.
' Uses LogFolder from the Config module.
Public Function EnsureLogFolder() As Boolean
    On Error GoTo ErrorHandler

    EnsureLogFolder = CreateFolderIfMissing(LogFolder)
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

    If Len(OutputFolder) = 0 Or Len(fileName) = 0 Or Len(extension) = 0 Then
        BuildOutputFilePath = vbNullString
        Exit Function
    End If

    BuildOutputFilePath = CombinePath(OutputFolder, fileName & "." & extension)
    Exit Function

ErrorHandler:
    BuildOutputFilePath = vbNullString
End Function

' Replaces characters that are invalid in Windows file names with underscores.
' Leading and trailing whitespace is removed from the returned value.
Public Function SanitizeFileName(ByVal fileName As String) As String
    Dim invalidCharacters As Variant
    Dim characterIndex As Long
    Dim characterCode As Long

    On Error GoTo ErrorHandler

    invalidCharacters = Array("<", ">", ":", """", "/", "\", "|", "?", "*")
    fileName = Trim$(fileName)

    For characterIndex = LBound(invalidCharacters) To UBound(invalidCharacters)
        fileName = Replace$(fileName, CStr(invalidCharacters(characterIndex)), DEFAULT_INVALID_CHARACTER_REPLACEMENT)
    Next characterIndex

    For characterCode = 0 To 31
        fileName = Replace$(fileName, Chr$(characterCode), DEFAULT_INVALID_CHARACTER_REPLACEMENT)
    Next characterCode

    Do While Len(fileName) > 0 And Right$(fileName, 1) = "."
        fileName = Left$(fileName, Len(fileName) - 1)
    Loop

    fileName = Trim$(fileName)

    If Len(fileName) = 0 Then
        fileName = DEFAULT_FILE_NAME
    End If

    SanitizeFileName = fileName
    Exit Function

ErrorHandler:
    SanitizeFileName = vbNullString
End Function

' Returns a timestamp in YYYYMMDD_HHMMSS format for file names and log entries.
Public Function GetTimeStamp() As String
    On Error GoTo ErrorHandler

    GetTimeStamp = Format$(Now, "yyyymmdd_hhnnss")
    Exit Function

ErrorHandler:
    GetTimeStamp = vbNullString
End Function

Private Sub CreateFolderHierarchy(ByVal folderPath As String)
    Dim parentFolder As String

    If Len(folderPath) = 0 Or FolderExists(folderPath) Then
        Exit Sub
    End If

    parentFolder = GetParentFolder(folderPath)

    If Len(parentFolder) > 0 And Not FolderExists(parentFolder) Then
        CreateFolderHierarchy parentFolder
    End If

    MkDir folderPath
End Sub

Private Function CombinePath(ByVal basePath As String, ByVal childPath As String) As String
    basePath = RemoveTrailingPathSeparator(basePath)
    childPath = RemoveLeadingPathSeparator(childPath)

    If Len(basePath) = 0 Then
        CombinePath = childPath
    ElseIf Len(childPath) = 0 Then
        CombinePath = basePath
    Else
        CombinePath = basePath & WINDOWS_PATH_SEPARATOR & childPath
    End If
End Function

Private Function NormalizeFolderPath(ByVal folderPath As String) As String
    NormalizeFolderPath = RemoveTrailingPathSeparator(folderPath)
End Function

Private Function NormalizeExtension(ByVal extension As String) As String
    extension = Trim$(extension)

    Do While Left$(extension, 1) = "."
        extension = Mid$(extension, 2)
    Loop

    NormalizeExtension = extension
End Function

Private Function RemoveTrailingPathSeparator(ByVal folderPath As String) As String
    Do While Len(folderPath) > 1 _
            And Right$(folderPath, 1) = WINDOWS_PATH_SEPARATOR _
            And Not IsDriveRoot(folderPath)
        folderPath = Left$(folderPath, Len(folderPath) - 1)
    Loop

    RemoveTrailingPathSeparator = folderPath
End Function

Private Function RemoveLeadingPathSeparator(ByVal pathValue As String) As String
    Do While Len(pathValue) > 0 And Left$(pathValue, 1) = WINDOWS_PATH_SEPARATOR
        pathValue = Mid$(pathValue, 2)
    Loop

    RemoveLeadingPathSeparator = pathValue
End Function

Private Function GetParentFolder(ByVal folderPath As String) As String
    Dim separatorPosition As Long

    folderPath = RemoveTrailingPathSeparator(folderPath)
    separatorPosition = InStrRev(folderPath, WINDOWS_PATH_SEPARATOR)

    If separatorPosition = 0 Then
        GetParentFolder = vbNullString
    ElseIf separatorPosition = 3 And Mid$(folderPath, 2, 1) = ":" Then
        GetParentFolder = Left$(folderPath, separatorPosition)
    Else
        GetParentFolder = Left$(folderPath, separatorPosition - 1)
    End If
End Function

Private Function IsDriveRoot(ByVal folderPath As String) As Boolean
    IsDriveRoot = (Len(folderPath) = 3 _
        And Mid$(folderPath, 2, 1) = ":" _
        And Right$(folderPath, 1) = WINDOWS_PATH_SEPARATOR)
End Function
