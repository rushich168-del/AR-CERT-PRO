Attribute VB_Name = "Logger"
Option Explicit

Private Const LOG_FILE_PREFIX As String = "AR-CERT-PRO"
Private Const LOG_FILE_EXTENSION As String = ".log"
Private Const LOG_ENTRY_DATE_FORMAT As String = "yyyy-mm-dd hh:nn:ss"

Private mLogFilePath As String

' Initializes a timestamped log file in the configured log folder.
Public Function InitLog() As Boolean
    Dim fileNumber As Integer
    Dim logFileName As String

    On Error GoTo ErrorHandler

    If Len(ProjectFolder) = 0 Or Len(CurrentLogFolder) = 0 Then
        If Not InitializeConfig() Then
            InitLog = False
            Exit Function
        End If
    End If

    If Not EnsureFolderExists(CurrentLogFolder) Then
        InitLog = False
        Exit Function
    End If

    logFileName = CleanFileName(LOG_FILE_PREFIX & "_" & PROJECT_VERSION & "_" & GetTimeStamp()) & LOG_FILE_EXTENSION
    mLogFilePath = CombinePath(CurrentLogFolder, logFileName)

    fileNumber = FreeFile
    Open mLogFilePath For Append As #fileNumber
    Print #fileNumber, Format$(Now, LOG_ENTRY_DATE_FORMAT) & " | INFO | Log initialized for " & PROJECT_NAME & " " & PROJECT_VERSION
    Close #fileNumber

    InitLog = True
    Exit Function

ErrorHandler:
    On Error Resume Next
    If fileNumber > 0 Then Close #fileNumber
    InitLog = False
End Function

' Writes an informational message to the current log file.
Public Sub WriteLog(ByVal message As String)
    Dim fileNumber As Integer

    On Error GoTo ErrorHandler

    If Len(mLogFilePath) = 0 Then
        If Not InitLog() Then
            Exit Sub
        End If
    End If

    If Len(Trim$(message)) = 0 Then
        message = "(blank message)"
    End If

    fileNumber = FreeFile
    Open mLogFilePath For Append As #fileNumber
    Print #fileNumber, Format$(Now, LOG_ENTRY_DATE_FORMAT) & " | INFO | " & message
    Close #fileNumber
    Exit Sub

ErrorHandler:
    On Error Resume Next
    If fileNumber > 0 Then Close #fileNumber
End Sub

' Writes a standardized error entry to the current log file.
Public Sub WriteErrorLog(ByVal procedureName As String, ByVal errorNumber As Long, ByVal errorDescription As String)
    Dim fileNumber As Integer
    Dim logMessage As String

    On Error GoTo ErrorHandler

    If Len(mLogFilePath) = 0 Then
        If Not InitLog() Then
            Exit Sub
        End If
    End If

    procedureName = Trim$(procedureName)

    If Len(procedureName) = 0 Then
        procedureName = "UnknownProcedure"
    End If

    logMessage = Format$(Now, LOG_ENTRY_DATE_FORMAT) _
        & " | ERROR | " & procedureName _
        & " | " & CStr(errorNumber) _
        & " | " & errorDescription

    fileNumber = FreeFile
    Open mLogFilePath For Append As #fileNumber
    Print #fileNumber, logMessage
    Close #fileNumber
    Exit Sub

ErrorHandler:
    On Error Resume Next
    If fileNumber > 0 Then Close #fileNumber
End Sub
