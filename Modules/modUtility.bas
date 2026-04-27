Attribute VB_Name = "modUtility"
Option Explicit

Public Function NzString(ByVal Value As Variant, Optional ByVal DefaultValue As String = "") As String
    If IsNull(Value) Or IsEmpty(Value) Then
        NzString = DefaultValue
    Else
        NzString = Trim$(CStr(Value))
    End If
End Function

Public Function IsNullOrEmpty(ByVal Value As String) As Boolean
    IsNullOrEmpty = (Len(Trim$(Value)) = 0)
End Function

Public Function FileExists(ByVal FilePath As String) As Boolean
    On Error Resume Next
    FileExists = (Len(Dir$(FilePath)) > 0)
End Function

Public Function FolderExists(ByVal FolderPath As String) As Boolean
    On Error Resume Next
    FolderExists = (Len(Dir$(FolderPath, vbDirectory)) > 0)
End Function

Public Function EnsureTrailingBackslash(ByVal FolderPath As String) As String
    If Len(FolderPath) = 0 Then
        EnsureTrailingBackslash = ""
    ElseIf Right$(FolderPath, 1) = "\" Then
        EnsureTrailingBackslash = FolderPath
    Else
        EnsureTrailingBackslash = FolderPath & "\"
    End If
End Function

Public Function GetApplicationPath() As String
    GetApplicationPath = EnsureTrailingBackslash(App.Path)
End Function

Public Function GetLogsFolderPath() As String
    GetLogsFolderPath = GetApplicationPath() & APP_LOG_FOLDER & "\"
End Function

Public Sub EnsureFolderExists(ByVal FolderPath As String)
    If Not FolderExists(FolderPath) Then
        MkDir FolderPath
    End If
End Sub

Public Function BuildLogFileName() As String
    BuildLogFileName = "ImportLog_" & Format$(Now, "yyyymmdd_hhnnss") & ".txt"
End Function

Public Function GetFullLogFilePath() As String
    EnsureFolderExists GetLogsFolderPath()
    GetFullLogFilePath = GetLogsFolderPath() & BuildLogFileName()
End Function

Public Function EscapeSqlText(ByVal Value As String) As String
    EscapeSqlText = Replace(Value, "'", "''")
End Function

Public Function BracketName(ByVal ObjectName As String) As String
    BracketName = "[" & Replace(ObjectName, "]", "]]") & "]"
End Function

Public Function BuildFullTableName(ByVal SchemaName As String, ByVal TableName As String) As String
    BuildFullTableName = BracketName(SchemaName) & "." & BracketName(TableName)
End Function

Public Function GenerateStageTableName(ByVal TableName As String) As String
    Randomize
    GenerateStageTableName = DEFAULT_STAGE_PREFIX & _
                             CleanObjectName(TableName) & "_" & _
                             Format$(Now, "yyyymmdd_hhnnss") & "_" & _
                             CStr(Int((9000 * Rnd) + 1000))
End Function

Public Function CleanObjectName(ByVal Value As String) As String
    Dim Result As String
    
    Result = Trim$(Value)
    Result = Replace(Result, " ", "_")
    Result = Replace(Result, "-", "_")
    Result = Replace(Result, ".", "_")
    Result = Replace(Result, "/", "_")
    Result = Replace(Result, "\", "_")
    
    CleanObjectName = Result
End Function

Public Function IsExcelFileExtensionValid(ByVal FilePath As String) As Boolean
    Dim LowerPath As String
    
    LowerPath = LCase$(Trim$(FilePath))
    
    IsExcelFileExtensionValid = _
        (Right$(LowerPath, 4) = ".xls") Or _
        (Right$(LowerPath, 5) = ".xlsx")
End Function

Public Function GetFileExtension(ByVal FilePath As String) As String
    Dim Pos As Long
    
    Pos = InStrRev(FilePath, ".")
    If Pos > 0 Then
        GetFileExtension = LCase$(Mid$(FilePath, Pos))
    Else
        GetFileExtension = ""
    End If
End Function

Public Function SafeCLng(ByVal Value As Variant, Optional ByVal DefaultValue As Long = 0) As Long
    On Error GoTo ErrorHandler
    
    If IsNull(Value) Or IsEmpty(Value) Or Trim$(CStr(Value)) = "" Then
        SafeCLng = DefaultValue
    Else
        SafeCLng = CLng(Value)
    End If
    
    Exit Function

ErrorHandler:
    SafeCLng = DefaultValue
End Function

Public Function SafeCDbl(ByVal Value As Variant, Optional ByVal DefaultValue As Double = 0) As Double
    On Error GoTo ErrorHandler
    
    If IsNull(Value) Or IsEmpty(Value) Or Trim$(CStr(Value)) = "" Then
        SafeCDbl = DefaultValue
    Else
        SafeCDbl = CDbl(Value)
    End If
    
    Exit Function

ErrorHandler:
    SafeCDbl = DefaultValue
End Function

Public Function BoolToText(ByVal Value As Boolean) As String
    If Value Then
        BoolToText = "Yes"
    Else
        BoolToText = "No"
    End If
End Function
