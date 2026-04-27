VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsImportLogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mLogFilePath As String
Private mSuccessCount As Long
Private mSkipCount As Long
Private mErrorCount As Long
Private mStartedAt As Date

Public Property Get LogFilePath() As String
    LogFilePath = mLogFilePath
End Property

Public Property Get SuccessCount() As Long
    SuccessCount = mSuccessCount
End Property

Public Property Get SkipCount() As Long
    SkipCount = mSkipCount
End Property

Public Property Get ErrorCount() As Long
    ErrorCount = mErrorCount
End Property

Public Sub StartLog( _
    ByVal ServerName As String, _
    ByVal DatabaseName As String, _
    ByVal SchemaName As String, _
    ByVal TableName As String, _
    ByVal ExcelFilePath As String, _
    ByVal ReadModeText As String)

    mStartedAt = Now
    mLogFilePath = GetFullLogFilePath()
    mSuccessCount = 0
    mSkipCount = 0
    mErrorCount = 0
    
    WriteInfo "===================================================="
    WriteInfo "Import started at: " & Format$(mStartedAt, "yyyy-mm-dd hh:nn:ss")
    WriteInfo "Server: " & ServerName
    WriteInfo "Database: " & DatabaseName
    WriteInfo "Target: " & SchemaName & "." & TableName
    WriteInfo "Excel File: " & ExcelFilePath
    WriteInfo "Excel Read Mode: " & ReadModeText
    WriteInfo "===================================================="
End Sub

Public Sub WriteInfo(ByVal MessageText As String)
    WriteTextLine mLogFilePath, "[INFO ] " & MessageText
End Sub

Public Sub WriteWarning(ByVal MessageText As String)
    WriteTextLine mLogFilePath, "[WARN ] " & MessageText
End Sub

Public Sub WriteError(ByVal MessageText As String)
    WriteTextLine mLogFilePath, "[ERROR] " & MessageText
End Sub

Public Sub IncrementSuccess()
    mSuccessCount = mSuccessCount + 1
End Sub

Public Sub IncrementSkip()
    mSkipCount = mSkipCount + 1
End Sub

Public Sub IncrementError()
    mErrorCount = mErrorCount + 1
End Sub

Public Sub FinishLog(ByVal FinalStatus As String)
    WriteInfo "----------------------------------------------------"
    WriteInfo "Finished at: " & Format$(Now, "yyyy-mm-dd hh:nn:ss")
    WriteInfo "Final Status: " & FinalStatus
    WriteInfo "Success Count: " & CStr(mSuccessCount)
    WriteInfo "Skip Count: " & CStr(mSkipCount)
    WriteInfo "Error Count: " & CStr(mErrorCount)
    WriteInfo "===================================================="
End Sub

Public Sub ResetCounters()
    mSuccessCount = 0
    mSkipCount = 0
    mErrorCount = 0
End Sub
