VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsImportEngine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mIdentityInsertEnabled As Boolean

Public Sub ExecuteImport( _
    ByVal Conn As ADODB.Connection, _
    ByVal DatabaseName As String, _
    ByVal SchemaName As String, _
    ByVal TargetTableName As String, _
    ByVal StageTableName As String, _
    ByVal Mappings As Collection, _
    ByVal TargetColumns As Collection, _
    ByVal Options As clsImportOptions, _
    ByVal Logger As clsImportLogger, _
    ByVal OwnerForm As Object)
    
    On Error GoTo ErrorHandler
    
    Dim OriginalErrNumber As Long
    Dim OriginalErrDescription As String
    
    Conn.BeginTrans
    mIdentityInsertEnabled = False
    
    If Options.DeleteExistingData Then
        UpdateImportProgress OwnerForm, 0, 0, Logger.SuccessCount, Logger.SkipCount, Logger.ErrorCount, "Deleting existing rows..."
        Logger.WriteInfo "Deleting existing rows from target table..."
        DeleteExistingRows Conn, DatabaseName, SchemaName, TargetTableName
    End If
    
    If Options.UseIdentityInsert Then
        Logger.WriteInfo "Enabling IDENTITY_INSERT..."
        SetIdentityInsert Conn, DatabaseName, SchemaName, TargetTableName, True
        mIdentityInsertEnabled = True
    End If
    
    Logger.WriteInfo "Starting transfer from staging to target..."
    TransferStageToTarget Conn, DatabaseName, SchemaName, TargetTableName, StageTableName, Mappings, TargetColumns, Options, Logger, OwnerForm
    
    If mIdentityInsertEnabled Then
        Logger.WriteInfo "Disabling IDENTITY_INSERT..."
        SetIdentityInsert Conn, DatabaseName, SchemaName, TargetTableName, False
        mIdentityInsertEnabled = False
    End If
    
    Conn.CommitTrans
    Logger.WriteInfo "Transaction committed successfully."
    Exit Sub

ErrorHandler:
    OriginalErrNumber = Err.Number
    OriginalErrDescription = Err.Description
    
    On Error Resume Next
    
    If mIdentityInsertEnabled Then
        SetIdentityInsert Conn, DatabaseName, SchemaName, TargetTableName, False
        mIdentityInsertEnabled = False
    End If
    
    Conn.RollbackTrans
    
    If Not Logger Is Nothing Then
        Logger.WriteError "Transaction rolled back. Original error " & CStr(OriginalErrNumber) & ": " & OriginalErrDescription
    End If
    
    Err.Clear
    Err.Raise vbObjectError + 2000, "clsImportEngine.ExecuteImport", OriginalErrDescription
End Sub

Public Sub DeleteExistingRows( _
    ByVal Conn As ADODB.Connection, _
    ByVal DatabaseName As String, _
    ByVal SchemaName As String, _
    ByVal TargetTableName As String)
    
    Dim SqlText As String
    Dim SafeDatabaseName As String
    
    SafeDatabaseName = BracketName(DatabaseName)
    
    SqlText = "DELETE FROM " & SafeDatabaseName & "." & BracketName(SchemaName) & "." & BracketName(TargetTableName)
    Conn.Execute SqlText
End Sub

Public Sub SetIdentityInsert( _
    ByVal Conn As ADODB.Connection, _
    ByVal DatabaseName As String, _
    ByVal SchemaName As String, _
    ByVal TargetTableName As String, _
    ByVal EnableValue As Boolean)
    
    Dim SqlText As String
    Dim SafeDatabaseName As String
    
    SafeDatabaseName = BracketName(DatabaseName)
    
    SqlText = "SET IDENTITY_INSERT " & SafeDatabaseName & "." & BracketName(SchemaName) & "." & BracketName(TargetTableName) & " "
    
    If EnableValue Then
        SqlText = SqlText & "ON"
    Else
        SqlText = SqlText & "OFF"
    End If
    
    Conn.Execute SqlText
End Sub

Private Sub TransferStageToTarget( _
    ByVal Conn As ADODB.Connection, _
    ByVal DatabaseName As String, _
    ByVal SchemaName As String, _
    ByVal TargetTableName As String, _
    ByVal StageTableName As String, _
    ByVal Mappings As Collection, _
    ByVal TargetColumns As Collection, _
    ByVal Options As clsImportOptions, _
    ByVal Logger As clsImportLogger, _
    ByVal OwnerForm As Object)
    
    On Error GoTo ErrorHandler
    
    Dim SafeDatabaseName As String
    Dim SqlText As String
    Dim Rs As ADODB.Recordset
    Dim TotalRows As Long
    Dim CurrentRow As Long
    Dim ColumnList As String
    Dim ValueList As String
    Dim MapItem As clsMappingItem
    Dim v As Variant
    Dim ColInfo As clsColumnInfo
    Dim StageValue As Variant
    Dim InsertSql As String
    Dim DuplicateErr As Boolean
    Dim SqlErrNumber As Long
    Dim SqlErrDescription As String
    Dim OriginalErrNumber As Long
    Dim OriginalErrDescription As String
    
    SafeDatabaseName = BracketName(DatabaseName)
    
    TotalRows = GetStageRowCount(Conn, DatabaseName, SchemaName, StageTableName)
    
    If TotalRows = 0 Then Exit Sub
    
    SqlText = ""
    SqlText = SqlText & "SELECT * FROM " & SafeDatabaseName & "." & BracketName(SchemaName) & "." & BracketName(StageTableName) & vbCrLf
    SqlText = SqlText & "ORDER BY StageRowId"
    
    Set Rs = New ADODB.Recordset
    Rs.Open SqlText, Conn, adOpenForwardOnly, adLockReadOnly
    
    CurrentRow = 0
    
    Do While Not Rs.EOF
        If Not OwnerForm Is Nothing Then
            If OwnerForm.CancelRequested Then
                Logger.WriteWarning "Import cancelled by user during target transfer."
                Err.Raise vbObjectError + 3003, "clsImportEngine.TransferStageToTarget", "Import was cancelled by user during target transfer."
            End If
        End If
        
        CurrentRow = CurrentRow + 1
        
        ColumnList = ""
        ValueList = ""
        
        For Each v In Mappings
            Set MapItem = v
            
            If MapItem.IsMapped Then
                Set ColInfo = FindColumnInfo(TargetColumns, MapItem.SqlColumnName)
                
                If Not ColInfo Is Nothing Then
                    If Len(ColumnList) > 0 Then
                        ColumnList = ColumnList & ", "
                        ValueList = ValueList & ", "
                    End If
                    
                    ColumnList = ColumnList & BracketName(ColInfo.ColumnName)
                    StageValue = Rs.Fields(ColInfo.ColumnName).Value
                    ValueList = ValueList & ConvertStageValueToSqlLiteral(StageValue, ColInfo)
                End If
            End If
        Next
        
        InsertSql = ""
        InsertSql = InsertSql & "INSERT INTO " & SafeDatabaseName & "." & BracketName(SchemaName) & "." & BracketName(TargetTableName)
        InsertSql = InsertSql & " (" & ColumnList & ") VALUES (" & ValueList & ")"
        
        SqlErrNumber = 0
        SqlErrDescription = ""
        
        On Error Resume Next
        Err.Clear
        Conn.Execute InsertSql
        SqlErrNumber = Err.Number
        SqlErrDescription = Err.Description
        On Error GoTo ErrorHandler
        
        If SqlErrNumber <> 0 Then
            DuplicateErr = IsDuplicateError(SqlErrNumber, SqlErrDescription)
            
            If DuplicateErr Then
                Logger.IncrementError
                
                If Options.ContinueOnDuplicate Then
                    Logger.IncrementSkip
                    Logger.WriteWarning "Duplicate row skipped during target import. Stage row " & CStr(CurrentRow) & ". Error " & CStr(SqlErrNumber) & ": " & SqlErrDescription
                Else
                    Logger.WriteError "Duplicate error during target import. Stage row " & CStr(CurrentRow) & ". Error " & CStr(SqlErrNumber) & ": " & SqlErrDescription
                    Err.Raise vbObjectError + 1402, "clsImportEngine.TransferStageToTarget", _
                              "Duplicate error during target import at stage row " & CStr(CurrentRow) & ": " & SqlErrDescription
                End If
            Else
                Logger.WriteError "Target import error at stage row " & CStr(CurrentRow) & ". Error " & CStr(SqlErrNumber) & ": " & SqlErrDescription
                Err.Raise vbObjectError + 2001, "clsImportEngine.TransferStageToTarget", _
                          "Target import error at stage row " & CStr(CurrentRow) & ": " & SqlErrDescription
            End If
        Else
            Logger.IncrementSuccess
        End If
        
        UpdateImportProgress OwnerForm, TotalRows, CurrentRow, Logger.SuccessCount, Logger.SkipCount, Logger.ErrorCount, "Importing rows to target..."
        Rs.MoveNext
    Loop
    
    Rs.Close
    Set Rs = Nothing
    Exit Sub

ErrorHandler:
    OriginalErrNumber = Err.Number
    OriginalErrDescription = Err.Description
    
    On Error Resume Next
    
    If Not Rs Is Nothing Then
        If Rs.State = adStateOpen Then Rs.Close
    End If
    Set Rs = Nothing
    
    Err.Clear
    Err.Raise vbObjectError + 2002, "clsImportEngine.TransferStageToTarget", OriginalErrDescription
End Sub

Private Function GetStageRowCount( _
    ByVal Conn As ADODB.Connection, _
    ByVal DatabaseName As String, _
    ByVal SchemaName As String, _
    ByVal StageTableName As String) As Long
    
    Dim SqlText As String
    Dim Rs As ADODB.Recordset
    Dim SafeDatabaseName As String
    
    SafeDatabaseName = BracketName(DatabaseName)
    
    SqlText = "SELECT COUNT(*) FROM " & SafeDatabaseName & "." & BracketName(SchemaName) & "." & BracketName(StageTableName)
    
    Set Rs = New ADODB.Recordset
    Rs.Open SqlText, Conn, adOpenForwardOnly, adLockReadOnly
    
    If Not Rs.EOF Then
        GetStageRowCount = SafeCLng(Rs.Fields(0).Value, 0)
    Else
        GetStageRowCount = 0
    End If
    
    Rs.Close
    Set Rs = Nothing
End Function

Private Function FindColumnInfo(ByVal TargetColumns As Collection, ByVal ColumnName As String) As clsColumnInfo
    Dim i As Long
    Dim ColInfo As clsColumnInfo
    
    For i = 1 To TargetColumns.Count
        Set ColInfo = TargetColumns(i)
        
        If StrComp(Trim$(ColInfo.ColumnName), Trim$(ColumnName), vbTextCompare) = 0 Then
            Set FindColumnInfo = ColInfo
            Exit Function
        End If
    Next i
    
    Set FindColumnInfo = Nothing
End Function

Private Function ConvertStageValueToSqlLiteral(ByVal RawValue As Variant, ByVal ColInfo As clsColumnInfo) As String
    Dim ValueText As String
    Dim DataTypeName As String
    
    DataTypeName = LCase$(Trim$(ColInfo.DataType))
    
    If IsNull(RawValue) Or IsEmpty(RawValue) Then
        ConvertStageValueToSqlLiteral = "NULL"
        Exit Function
    End If
    
    ValueText = Trim$(CStr(RawValue))
    
    If Len(ValueText) = 0 Then
        ConvertStageValueToSqlLiteral = "NULL"
        Exit Function
    End If
    
    Select Case UCase$(ValueText)
        Case "NULL", "(NULL)"
            ConvertStageValueToSqlLiteral = "NULL"
            Exit Function
    End Select
    
    Select Case DataTypeName
        Case "int", "bigint", "smallint", "tinyint", "decimal", "numeric", "money", "float", "real"
            ConvertStageValueToSqlLiteral = Replace(ValueText, ",", ".")
            
        Case "bit"
            Select Case LCase$(ValueText)
                Case "1", "true", "yes", "y"
                    ConvertStageValueToSqlLiteral = "1"
                Case "0", "false", "no", "n"
                    ConvertStageValueToSqlLiteral = "0"
                Case Else
                    If IsNumeric(ValueText) Then
                        If CDbl(ValueText) = 0 Then
                            ConvertStageValueToSqlLiteral = "0"
                        Else
                            ConvertStageValueToSqlLiteral = "1"
                        End If
                    Else
                        Err.Raise vbObjectError + 2101, "clsImportEngine.ConvertStageValueToSqlLiteral", _
                                  "Invalid bit value [" & ValueText & "] for column [" & ColInfo.ColumnName & "]"
                    End If
            End Select
            
        Case "datetime", "date", "smalldatetime"
            ConvertStageValueToSqlLiteral = "'" & Format$(CDate(ValueText), "yyyy-mm-dd hh:nn:ss") & "'"
            
        Case "uniqueidentifier"
            ConvertStageValueToSqlLiteral = "'" & EscapeSqlText(ValueText) & "'"
            
        Case Else
            ConvertStageValueToSqlLiteral = "N'" & EscapeSqlText(ValueText) & "'"
    End Select
End Function

Private Function IsDuplicateError(ByVal ErrorNumber As Long, ByVal ErrorDescription As String) As Boolean
    Dim DescText As String
    
    DescText = LCase$(ErrorDescription)
    
    If ErrorNumber = -2147217873 Then
        IsDuplicateError = True
        Exit Function
    End If
    
    If InStr(1, DescText, "duplicate", vbTextCompare) > 0 Then
        IsDuplicateError = True
        Exit Function
    End If
    
    If InStr(1, DescText, "unique key", vbTextCompare) > 0 Then
        IsDuplicateError = True
        Exit Function
    End If
    
    If InStr(1, DescText, "primary key", vbTextCompare) > 0 Then
        IsDuplicateError = True
        Exit Function
    End If
    
    IsDuplicateError = False
End Function

Private Sub UpdateImportProgress( _
    ByVal OwnerForm As Object, _
    ByVal TotalRows As Long, _
    ByVal CurrentRow As Long, _
    ByVal SuccessCount As Long, _
    ByVal SkipCount As Long, _
    ByVal ErrorCount As Long, _
    ByVal StepText As String)
    
    On Error Resume Next
    
    OwnerForm.lblCurrentStep.Caption = StepText
    OwnerForm.lblRowProgress.Caption = CStr(CurrentRow) & " / " & CStr(TotalRows)
    OwnerForm.lblSuccessCount.Caption = CStr(SuccessCount)
    OwnerForm.lblSkipCount.Caption = CStr(SkipCount)
    OwnerForm.lblErrorCount.Caption = CStr(ErrorCount)
    
    If TotalRows > 0 Then
        OwnerForm.prgImport.Value = CLng((CurrentRow / TotalRows) * PROGRESS_MAX)
    Else
        OwnerForm.prgImport.Value = 0
    End If
    
    If (CurrentRow Mod 5) = 0 Or CurrentRow = TotalRows Then
        DoEvents
    End If
End Sub

