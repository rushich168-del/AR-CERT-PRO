Attribute VB_Name = "Config"
Option Explicit

' Application identity used throughout AR-CERT PRO.
Public Const APP_NAME As String = "AR-CERT PRO"
Public Const APP_VERSION As String = "v0.1"

' Global configuration values populated during application startup.
Public ProjectFolder As String
Public ExcelFilePath As String
Public TemplateFilePath As String
Public OutputFolder As String
Public LogFolder As String
Public LastConfigError As String

Private Const DEFAULT_OUTPUT_FOLDER_NAME As String = "Output"
Private Const DEFAULT_LOG_FOLDER_NAME As String = "Output\Logs"
Private Const DEFAULT_EXCEL_FILE_PATH As String = "Excel\Students.xlsx"
Private Const DEFAULT_TEMPLATE_FILE_PATH As String = "Templates\Certificate_Template.docx"
Private Const SOURCE_FOLDER_NAME As String = "Source"

' Initializes the application configuration using the current Word project location.
' This procedure detects the project folder, assigns default Output and Log folders,
' and creates the required folders when they do not already exist.
Public Function InitializeConfig() As Boolean
    On Error GoTo ErrorHandler

    LastConfigError = vbNullString
    ProjectFolder = DetectProjectFolder()

    If Len(ProjectFolder) = 0 Then
        LastConfigError = "Unable to detect the project folder."
        InitializeConfig = False
        Exit Function
    End If

    ExcelFilePath = CombinePath(ProjectFolder, DEFAULT_EXCEL_FILE_PATH)
    TemplateFilePath = CombinePath(ProjectFolder, DEFAULT_TEMPLATE_FILE_PATH)
    OutputFolder = CombinePath(ProjectFolder, DEFAULT_OUTPUT_FOLDER_NAME)
    LogFolder = CombinePath(ProjectFolder, DEFAULT_LOG_FOLDER_NAME)

    EnsureFolderExists OutputFolder
    EnsureFolderExists LogFolder

    InitializeConfig = True
    Exit Function

ErrorHandler:
    LastConfigError = "InitializeConfig failed: " & Err.Description
    InitializeConfig = False
End Function

' Clears all configuration values so the application can be initialized again
' without carrying stale paths from a previous run.
Public Sub ResetConfig()
    On Error GoTo ErrorHandler

    ProjectFolder = vbNullString
    ExcelFilePath = vbNullString
    TemplateFilePath = vbNullString
    OutputFolder = vbNullString
    LogFolder = vbNullString
    LastConfigError = vbNullString
    Exit Sub

ErrorHandler:
    LastConfigError = "ResetConfig failed: " & Err.Description
End Sub

' Validates the minimum configuration required before certificate generation can run.
' Returns True only when the Excel file, Word template, and Output folder all exist.
Public Function ValidateConfiguration() As Boolean
    On Error GoTo ErrorHandler

    LastConfigError = vbNullString

    If Not FileExists(ExcelFilePath) Then
        LastConfigError = "Excel file does not exist: " & ExcelFilePath
        ValidateConfiguration = False
        Exit Function
    End If

    If Not FileExists(TemplateFilePath) Then
        LastConfigError = "Template file does not exist: " & TemplateFilePath
        ValidateConfiguration = False
        Exit Function
    End If

    If Not FolderExists(OutputFolder) Then
        LastConfigError = "Output folder does not exist: " & OutputFolder
        ValidateConfiguration = False
        Exit Function
    End If

    ValidateConfiguration = True
    Exit Function

ErrorHandler:
    LastConfigError = "ValidateConfiguration failed: " & Err.Description
    ValidateConfiguration = False
End Function

Private Function DetectProjectFolder() As String
    Dim basePath As String
    Dim parentPath As String

    basePath = ThisDocument.Path

    If Len(basePath) = 0 Then
        basePath = CurDir$
    End If

    basePath = RemoveTrailingPathSeparator(basePath)

    If StrComp(GetFolderName(basePath), SOURCE_FOLDER_NAME, vbTextCompare) = 0 Then
        parentPath = GetParentFolder(basePath)

        If Len(parentPath) > 0 Then
            DetectProjectFolder = parentPath
            Exit Function
        End If
    End If

    DetectProjectFolder = basePath
End Function

Private Sub EnsureFolderExists(ByVal folderPath As String)
    If Len(folderPath) = 0 Then
        Err.Raise vbObjectError + 1000, "Config.EnsureFolderExists", "Folder path is empty."
    End If

    If Not FolderExists(folderPath) Then
        MkDir folderPath
    End If
End Sub

Private Function FileExists(ByVal filePath As String) As Boolean
    Dim attributes As Long

    On Error GoTo ErrorHandler

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

Private Function FolderExists(ByVal folderPath As String) As Boolean
    Dim attributes As Long

    On Error GoTo ErrorHandler

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

Private Function CombinePath(ByVal basePath As String, ByVal childPath As String) As String
    basePath = RemoveTrailingPathSeparator(basePath)

    If Len(basePath) = 0 Then
        CombinePath = childPath
    Else
        CombinePath = basePath & Application.PathSeparator & childPath
    End If
End Function

Private Function RemoveTrailingPathSeparator(ByVal folderPath As String) As String
    Do While Len(folderPath) > 1 And Right$(folderPath, 1) = Application.PathSeparator
        folderPath = Left$(folderPath, Len(folderPath) - 1)
    Loop

    RemoveTrailingPathSeparator = folderPath
End Function

Private Function GetFolderName(ByVal folderPath As String) As String
    Dim separatorPosition As Long

    folderPath = RemoveTrailingPathSeparator(folderPath)
    separatorPosition = InStrRev(folderPath, Application.PathSeparator)

    If separatorPosition = 0 Then
        GetFolderName = folderPath
    Else
        GetFolderName = Mid$(folderPath, separatorPosition + 1)
    End If
End Function

Private Function GetParentFolder(ByVal folderPath As String) As String
    Dim separatorPosition As Long

    folderPath = RemoveTrailingPathSeparator(folderPath)
    separatorPosition = InStrRev(folderPath, Application.PathSeparator)

    If separatorPosition > 0 Then
        GetParentFolder = Left$(folderPath, separatorPosition - 1)
    Else
        GetParentFolder = vbNullString
    End If
End Function
