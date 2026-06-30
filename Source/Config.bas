Attribute VB_Name = "Config"
Option Explicit

' Project identity and default folder/file configuration.
Public Const PROJECT_NAME As String = "AR-CERT PRO"
Public Const PROJECT_VERSION As String = "v0.9"
Public Const TEMPLATE_FOLDER As String = "Templates"
Public Const EXCEL_FOLDER As String = "Excel"
Public Const OUTPUT_FOLDER As String = "Output"
Public Const LOG_FOLDER As String = "Output\Logs"
Public Const DEFAULT_EXCEL_FILE As String = "Students.xlsx"
Public Const DEFAULT_TEMPLATE_FILE As String = "Certificate_Template.docx"
Public Const DEFAULT_WORKSHEET_NAME As String = "Sheet1"

' Backward-compatible configuration values populated during application startup.
Public ProjectFolder As String
Public ExcelFilePath As String
Public TemplateFilePath As String
Public CurrentOutputFolder As String
Public CurrentLogFolder As String
Public LastConfigError As String

Private Const SOURCE_FOLDER_NAME As String = "Source"

' Initializes the application configuration using the current Word project location.
Public Function InitializeConfig() As Boolean
    On Error GoTo ErrorHandler

    LastConfigError = vbNullString
    ProjectFolder = GetProjectPath()
    ExcelFilePath = GetDefaultExcelFilePath()
    TemplateFilePath = GetDefaultTemplateFilePath()
    CurrentOutputFolder = GetOutputFolderPath()
    CurrentLogFolder = GetLogFolderPath()

    If Len(ProjectFolder) = 0 Then
        LastConfigError = "Unable to detect the project folder."
        InitializeConfig = False
        Exit Function
    End If

    If Not EnsureFolderExists(CurrentOutputFolder) Then
        LastConfigError = "Unable to create output folder: " & CurrentOutputFolder
        InitializeConfig = False
        Exit Function
    End If

    If Not EnsureFolderExists(CurrentLogFolder) Then
        LastConfigError = "Unable to create log folder: " & CurrentLogFolder
        InitializeConfig = False
        Exit Function
    End If

    InitializeConfig = True
    Exit Function

ErrorHandler:
    LastConfigError = "InitializeConfig failed: " & Err.Description
    InitializeConfig = False
End Function

' Clears all configuration values so the application can be initialized again.
Public Sub ResetConfig()
    On Error GoTo ErrorHandler

    ProjectFolder = vbNullString
    ExcelFilePath = vbNullString
    TemplateFilePath = vbNullString
    CurrentOutputFolder = vbNullString
    CurrentLogFolder = vbNullString
    LastConfigError = vbNullString
    Exit Sub

ErrorHandler:
    LastConfigError = "ResetConfig failed: " & Err.Description
End Sub

' Validates the minimum configuration required before certificate generation can run.
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

    If Not FolderExists(CurrentOutputFolder) Then
        LastConfigError = "Output folder does not exist: " & CurrentOutputFolder
        ValidateConfiguration = False
        Exit Function
    End If

    ValidateConfiguration = True
    Exit Function

ErrorHandler:
    LastConfigError = "ValidateConfiguration failed: " & Err.Description
    ValidateConfiguration = False
End Function

' Returns the project root derived from ThisDocument.Path.
Public Function GetProjectPath() As String
    Dim basePath As String
    Dim parentPath As String

    On Error GoTo ErrorHandler

    basePath = Trim$(ThisDocument.Path)

    If Len(basePath) = 0 Then
        basePath = CurDir$
    End If

    basePath = RemoveTrailingPathSeparator(basePath)

    If StrComp(GetFolderName(basePath), SOURCE_FOLDER_NAME, vbTextCompare) = 0 Then
        parentPath = GetParentFolder(basePath)

        If Len(parentPath) > 0 Then
            GetProjectPath = parentPath
            Exit Function
        End If
    End If

    GetProjectPath = basePath
    Exit Function

ErrorHandler:
    LastConfigError = "GetProjectPath failed: " & Err.Description
    GetProjectPath = vbNullString
End Function

Public Function GetTemplateFolderPath() As String
    On Error GoTo ErrorHandler

    GetTemplateFolderPath = CombinePath(GetProjectPath(), TEMPLATE_FOLDER)
    Exit Function

ErrorHandler:
    LastConfigError = "GetTemplateFolderPath failed: " & Err.Description
    GetTemplateFolderPath = vbNullString
End Function

Public Function GetExcelFolderPath() As String
    On Error GoTo ErrorHandler

    GetExcelFolderPath = CombinePath(GetProjectPath(), EXCEL_FOLDER)
    Exit Function

ErrorHandler:
    LastConfigError = "GetExcelFolderPath failed: " & Err.Description
    GetExcelFolderPath = vbNullString
End Function

Public Function GetOutputFolderPath() As String
    On Error GoTo ErrorHandler

    GetOutputFolderPath = CombinePath(GetProjectPath(), OUTPUT_FOLDER)
    Exit Function

ErrorHandler:
    LastConfigError = "GetOutputFolderPath failed: " & Err.Description
    GetOutputFolderPath = vbNullString
End Function

Public Function GetLogFolderPath() As String
    On Error GoTo ErrorHandler

    GetLogFolderPath = CombinePath(GetProjectPath(), LOG_FOLDER)
    Exit Function

ErrorHandler:
    LastConfigError = "GetLogFolderPath failed: " & Err.Description
    GetLogFolderPath = vbNullString
End Function

Public Function GetDefaultExcelFilePath() As String
    On Error GoTo ErrorHandler

    GetDefaultExcelFilePath = CombinePath(GetExcelFolderPath(), DEFAULT_EXCEL_FILE)
    Exit Function

ErrorHandler:
    LastConfigError = "GetDefaultExcelFilePath failed: " & Err.Description
    GetDefaultExcelFilePath = vbNullString
End Function

Public Function GetDefaultTemplateFilePath() As String
    On Error GoTo ErrorHandler

    GetDefaultTemplateFilePath = CombinePath(GetTemplateFolderPath(), DEFAULT_TEMPLATE_FILE)
    Exit Function

ErrorHandler:
    LastConfigError = "GetDefaultTemplateFilePath failed: " & Err.Description
    GetDefaultTemplateFilePath = vbNullString
End Function
