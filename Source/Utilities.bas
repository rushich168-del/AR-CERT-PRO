Attribute VB_Name = "Utilities"
Option Explicit

Private Const WINDOWS_PATH_SEPARATOR As String = "\"
Private Const DEFAULT_INVALID_CHARACTER_REPLACEMENT As String = "_"
Private Const DEFAULT_FILE_NAME As String = "untitled"

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

' Creates the supplied folder and any missing parent folders.
Public Function EnsureFolderExists(ByVal folderPath As String) As Boolean
    On Error GoTo ErrorHandler

    folderPath = RemoveTrailingPathSeparator(Trim$(folderPath))

    If Len(folderPath) = 0 Then
        EnsureFolderExists = False
        Exit Function
    End If

    If Not FolderExists(folderPath) Then
        CreateFolderHierarchy folderPath
    End If

    EnsureFolderExists = FolderExists(folderPath)
    Exit Function

ErrorHandler:
    EnsureFolderExists = False
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

' Replaces characters that are invalid in Windows file names with underscores.
Public Function CleanFileName(ByVal fileName As String) As String
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

    CleanFileName = fileName
    Exit Function

ErrorHandler:
    CleanFileName = DEFAULT_FILE_NAME
End Function

' Returns a timestamp in YYYYMMDD_HHMMSS format for file names and run IDs.
Public Function GetTimeStamp() As String
    On Error GoTo ErrorHandler

    GetTimeStamp = Format$(Now, "yyyymmdd_hhnnss")
    Exit Function

ErrorHandler:
    GetTimeStamp = vbNullString
End Function

' Combines two path segments with a single Windows path separator.
Public Function CombinePath(ByVal basePath As String, ByVal childPath As String) As String
    On Error GoTo ErrorHandler

    basePath = RemoveTrailingPathSeparator(Trim$(basePath))
    childPath = RemoveLeadingPathSeparator(Trim$(childPath))

    If Len(basePath) = 0 Then
        CombinePath = childPath
    ElseIf Len(childPath) = 0 Then
        CombinePath = basePath
    Else
        CombinePath = basePath & WINDOWS_PATH_SEPARATOR & childPath
    End If

    Exit Function

ErrorHandler:
    CombinePath = vbNullString
End Function

Public Function RemoveTrailingPathSeparator(ByVal pathValue As String) As String
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

Public Function RemoveLeadingPathSeparator(ByVal pathValue As String) As String
    On Error GoTo ErrorHandler

    Do While Len(pathValue) > 0 And Left$(pathValue, 1) = WINDOWS_PATH_SEPARATOR
        pathValue = Mid$(pathValue, 2)
    Loop

    RemoveLeadingPathSeparator = pathValue
    Exit Function

ErrorHandler:
    RemoveLeadingPathSeparator = vbNullString
End Function

Public Function GetFolderName(ByVal folderPath As String) As String
    Dim separatorPosition As Long

    On Error GoTo ErrorHandler

    folderPath = RemoveTrailingPathSeparator(Trim$(folderPath))
    separatorPosition = InStrRev(folderPath, WINDOWS_PATH_SEPARATOR)

    If separatorPosition = 0 Then
        GetFolderName = folderPath
    Else
        GetFolderName = Mid$(folderPath, separatorPosition + 1)
    End If

    Exit Function

ErrorHandler:
    GetFolderName = vbNullString
End Function

Public Function GetParentFolder(ByVal pathValue As String) As String
    Dim separatorPosition As Long

    On Error GoTo ErrorHandler

    pathValue = RemoveTrailingPathSeparator(Trim$(pathValue))
    separatorPosition = InStrRev(pathValue, WINDOWS_PATH_SEPARATOR)

    If separatorPosition = 0 Then
        GetParentFolder = vbNullString
    ElseIf separatorPosition = 3 And Mid$(pathValue, 2, 1) = ":" Then
        GetParentFolder = Left$(pathValue, separatorPosition)
    Else
        GetParentFolder = Left$(pathValue, separatorPosition - 1)
    End If

    Exit Function

ErrorHandler:
    GetParentFolder = vbNullString
End Function

Private Sub CreateFolderHierarchy(ByVal folderPath As String)
    Dim parentFolder As String

    On Error GoTo ErrorHandler

    If Len(folderPath) = 0 Or FolderExists(folderPath) Then
        Exit Sub
    End If

    parentFolder = GetParentFolder(folderPath)

    If Len(parentFolder) > 0 And Not FolderExists(parentFolder) Then
        CreateFolderHierarchy parentFolder
    End If

    MkDir folderPath
    Exit Sub

ErrorHandler:
End Sub

Private Function IsDriveRoot(ByVal pathValue As String) As Boolean
    On Error GoTo ErrorHandler

    IsDriveRoot = (Len(pathValue) = 3 _
        And Mid$(pathValue, 2, 1) = ":" _
        And Right$(pathValue, 1) = WINDOWS_PATH_SEPARATOR)
    Exit Function

ErrorHandler:
    IsDriveRoot = False
End Function
