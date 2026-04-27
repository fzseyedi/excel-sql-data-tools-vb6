VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsStagingManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mStageTableName As String

Public Property Get StageTableName() As String
    StageTableName = mStageTableName
End Property

Public Property Let StageTableName(ByVal Value As String)
    mStageTableName = Trim$(Value)
End Property

Public Function CreateStageTable( _
    ByVal Conn As ADODB.Connection, _
    ByVal DatabaseName As String, _
    ByVal SchemaName As String, _
    ByVal BaseTableName As String, _
    ByVal Mappings As Collection, _
    ByVal TargetColumns As Collection) As String
    
    On Error GoTo ErrorHandler
    
    Dim SqlText As String
    Dim SafeDatabaseName As String
    Dim StageName As String
    Dim i As Long
    Dim MapItem As clsMappingItem
    Dim ColInfo As clsColumnInfo
    Dim ColumnDefs As String
    
    SafeDatabaseName = BracketName(DatabaseName)
    StageName = GenerateStageTableName(BaseTableName)
    mStageTableName = StageName
    
    ColumnDefs = ""
    
    For i = 1 To Mappings.Count
        Set MapItem = Mappings(i)
        
        If MapItem.IsMapped Then
            Set ColInfo = FindColumnInfo(TargetColumns, MapItem.SqlColumnName)
            
            If Not ColInfo Is Nothing Then
                If Len(ColumnDefs) > 0 Then
                    ColumnDefs = ColumnDefs & "," & vbCrLf
                End If
                
                ColumnDefs = ColumnDefs & "    " & BracketName(ColInfo.ColumnName) & " " & BuildSqlDataTypeDefinition(ColInfo)
            End If
        End If
    Next i
    
    If Len(ColumnDefs) = 0 Then
        Err.Raise vbObjectError + 1300, "clsStagingManager.CreateStageTable", "No mapped columns were found for staging table creation."
    End If
    
    SqlText = ""
    SqlText = SqlText & "IF OBJECT_ID('" & SchemaName & "." & StageName & "', 'U') IS NOT NULL " & vbCrLf
    SqlText = SqlText & "    DROP TABLE " & SafeDatabaseName & "." & BracketName(SchemaName) & "." & BracketName(StageName)

    Conn.Execute SqlText
    
    SqlText = ""
    SqlText = SqlText & "CREATE TABLE " & SafeDatabaseName & "." & BracketName(SchemaName) & "." & BracketName(StageName) & " (" & vbCrLf
    SqlText = SqlText & "    StageRowId INT IDENTITY(1,1) NOT NULL PRIMARY KEY," & vbCrLf
    SqlText = SqlText & ColumnDefs & vbCrLf
    SqlText = SqlText & ")"
    
    Conn.Execute SqlText
    
    CreateStageTable = StageName
    Exit Function

ErrorHandler:
    Err.Raise Err.Number, "clsStagingManager.CreateStageTable", Err.Description
End Function

Public Sub DropStageTable(ByVal Conn As ADODB.Connection, ByVal DatabaseName As String, ByVal SchemaName As String)
    On Error GoTo ErrorHandler
    
    Dim SqlText As String
    Dim SafeDatabaseName As String
    
    If Len(Trim$(mStageTableName)) = 0 Then Exit Sub
    
    SafeDatabaseName = BracketName(DatabaseName)
    
    SqlText = "DROP TABLE " & SafeDatabaseName & "." & BracketName(SchemaName) & "." & BracketName(mStageTableName)
    Conn.Execute SqlText
    
    mStageTableName = ""
    Exit Sub

ErrorHandler:
    Err.Raise vbObjectError + 1310, "clsStagingManager.DropStageTable", _
              "Failed to drop staging table [" & mStageTableName & "]. " & Err.Description
End Sub

Public Function InsertRowsIntoStage( _
    ByVal Conn As ADODB.Connection, _
    ByVal DatabaseName As String, _
    ByVal SchemaName As String, _
    ByVal ExcelRows As Collection, _
    ByVal Mappings As Collection, _
    ByVal TargetColumns As Collection, _
    ByVal Options As clsImportOptions, _
    ByVal Logger As clsImportLogger, _
    ByVal OwnerForm As Object) As Long
    
    On Error GoTo ErrorHandler
    
    Dim SafeDatabaseName As String
    Dim SqlText As String
    Dim i As Long
    Dim InsertedCount As Long
    Dim RowData As Object
    Dim MapItem As clsMappingItem
    Dim ColInfo As clsColumnInfo
    Dim ColumnList As String
    Dim ValueList As String
    Dim RawValue As Variant
    Dim ConvertedValue As String
    Dim RowHasTypeError As Boolean
    Dim RowErrorMessage As String
    
    SafeDatabaseName = BracketName(DatabaseName)
    InsertedCount = 0
    
    For i = 1 To ExcelRows.Count
        
        If Not OwnerForm Is Nothing Then
            If OwnerForm.CancelRequested Then
                Logger.WriteWarning "Import cancelled by user during staging."
                Err.Raise vbObjectError + 3002, "clsStagingManager.InsertRowsIntoStage", "Import was cancelled by user during staging."
            End If
        End If
        
        Set RowData = ExcelRows(i)
        
        ColumnList = ""
        ValueList = ""
        RowHasTypeError = False
        RowErrorMessage = ""
        
        For Each MapItem In Mappings
            If MapItem.IsMapped Then
                Set ColInfo = FindColumnInfo(TargetColumns, MapItem.SqlColumnName)
                
                If Not ColInfo Is Nothing Then
                    RawValue = Null
                    
                    If RowData.Exists(MapItem.ExcelColumnName) Then
                        RawValue = RowData(MapItem.ExcelColumnName)
                    End If
                    
                    ConvertedValue = ConvertValueForSqlLiteral(RawValue, ColInfo, RowHasTypeError, RowErrorMessage)
                    
                    If RowHasTypeError Then Exit For
                    
                    If Len(ColumnList) > 0 Then
                        ColumnList = ColumnList & ", "
                        ValueList = ValueList & ", "
                    End If
                    
                    ColumnList = ColumnList & BracketName(ColInfo.ColumnName)
                    ValueList = ValueList & ConvertedValue
                End If
            End If
        Next
        
        If RowHasTypeError Then
            Logger.IncrementError
            
            If Options.ContinueOnTypeError Then
                Logger.IncrementSkip
                Logger.WriteWarning "Row " & CStr(i + 1) & " skipped due to data validation error: " & RowErrorMessage
            Else
                Logger.WriteError "Data validation error at Excel row " & CStr(i + 1) & ": " & RowErrorMessage
                Err.Raise vbObjectError + 1301, "clsStagingManager.InsertRowsIntoStage", _
                          "Data validation error at Excel row " & CStr(i + 1) & ": " & RowErrorMessage
            End If
        Else
            SqlText = ""
            SqlText = SqlText & "INSERT INTO " & SafeDatabaseName & "." & BracketName(SchemaName) & "." & BracketName(mStageTableName)
            SqlText = SqlText & " (" & ColumnList & ") VALUES (" & ValueList & ")"
            
            Conn.Execute SqlText
            InsertedCount = InsertedCount + 1
            Logger.IncrementSuccess
        End If
        
        UpdateProgress OwnerForm, ExcelRows.Count, i, InsertedCount, Logger.SkipCount, Logger.ErrorCount, "Writing rows to staging..."
    Next i
    
    InsertRowsIntoStage = InsertedCount
    Exit Function

ErrorHandler:
    Dim OriginalErrNumber As Long
    Dim OriginalErrDescription As String
    
    OriginalErrNumber = Err.Number
    OriginalErrDescription = Err.Description
    
    Err.Clear
    Err.Raise vbObjectError + 1302, "clsStagingManager.InsertRowsIntoStage", OriginalErrDescription
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

Private Function BuildSqlDataTypeDefinition(ByVal ColInfo As clsColumnInfo) As String
    Select Case LCase$(ColInfo.DataType)
        Case "int"
            BuildSqlDataTypeDefinition = "INT NULL"
        Case "bigint"
            BuildSqlDataTypeDefinition = "BIGINT NULL"
        Case "smallint"
            BuildSqlDataTypeDefinition = "SMALLINT NULL"
        Case "tinyint"
            BuildSqlDataTypeDefinition = "TINYINT NULL"
        Case "bit"
            BuildSqlDataTypeDefinition = "BIT NULL"
        Case "decimal", "numeric"
            BuildSqlDataTypeDefinition = "DECIMAL(18,4) NULL"
        Case "money"
            BuildSqlDataTypeDefinition = "MONEY NULL"
        Case "float"
            BuildSqlDataTypeDefinition = "FLOAT NULL"
        Case "real"
            BuildSqlDataTypeDefinition = "REAL NULL"
        Case "datetime"
            BuildSqlDataTypeDefinition = "DATETIME NULL"
        Case "date"
            BuildSqlDataTypeDefinition = "DATE NULL"
        Case "smalldatetime"
            BuildSqlDataTypeDefinition = "SMALLDATETIME NULL"
        Case "uniqueidentifier"
            BuildSqlDataTypeDefinition = "UNIQUEIDENTIFIER NULL"
        Case "nvarchar"
            BuildSqlDataTypeDefinition = "NVARCHAR(4000) NULL"
        Case "varchar"
            BuildSqlDataTypeDefinition = "VARCHAR(4000) NULL"
        Case "nchar"
            BuildSqlDataTypeDefinition = "NCHAR(255) NULL"
        Case "char"
            BuildSqlDataTypeDefinition = "CHAR(255) NULL"
        Case "text"
            BuildSqlDataTypeDefinition = "VARCHAR(MAX) NULL"
        Case "ntext"
            BuildSqlDataTypeDefinition = "NVARCHAR(MAX) NULL"
        Case Else
            BuildSqlDataTypeDefinition = "NVARCHAR(4000) NULL"
    End Select
End Function

Private Function ConvertValueForSqlLiteral( _
    ByVal RawValue As Variant, _
    ByVal ColInfo As clsColumnInfo, _
    ByRef HasTypeError As Boolean, _
    ByRef ErrorMessage As String) As String
    
    On Error GoTo ConversionError
    
    Dim ValueText As String
    Dim DataTypeName As String
    Dim IntegerPart As String
    Dim DecimalPart As String
    Dim DotPos As Long
    
    HasTypeError = False
    ErrorMessage = ""
    DataTypeName = LCase$(Trim$(ColInfo.DataType))
    
    If IsNull(RawValue) Or IsEmpty(RawValue) Then
        ConvertValueForSqlLiteral = "NULL"
        Exit Function
    End If
    
    ValueText = Trim$(CStr(RawValue))
    
    If Len(ValueText) = 0 Then
        ConvertValueForSqlLiteral = "NULL"
        Exit Function
    End If
    
    Select Case UCase$(ValueText)
        Case "NULL", "(NULL)"
            ConvertValueForSqlLiteral = "NULL"
            Exit Function
    End Select
    
    Select Case DataTypeName
        Case "int", "bigint", "smallint", "tinyint"
            If Not IsNumeric(ValueText) Then GoTo ConversionError
            ConvertValueForSqlLiteral = CStr(CLng(ValueText))
            
        Case "decimal", "numeric", "money", "float", "real"
            If Not IsNumeric(ValueText) Then GoTo ConversionError
            
            ValueText = Replace(ValueText, ",", ".")
            
            If ColInfo.NumericPrecision > 0 Then
                DotPos = InStr(1, ValueText, ".")
                
                If DotPos > 0 Then
                    IntegerPart = Replace(Left$(ValueText, DotPos - 1), "-", "")
                    DecimalPart = Mid$(ValueText, DotPos + 1)
                Else
                    IntegerPart = Replace(ValueText, "-", "")
                    DecimalPart = ""
                End If
                
                If Len(IntegerPart) + Len(DecimalPart) > ColInfo.NumericPrecision Then GoTo ConversionError
                
                If ColInfo.NumericScale > 0 Then
                    If Len(DecimalPart) > ColInfo.NumericScale Then GoTo ConversionError
                Else
                    If Len(DecimalPart) > 0 Then GoTo ConversionError
                End If
            End If
            
            ConvertValueForSqlLiteral = ValueText
            
        Case "bit"
            Select Case LCase$(ValueText)
                Case "1", "true", "yes", "y"
                    ConvertValueForSqlLiteral = "1"
                Case "0", "false", "no", "n"
                    ConvertValueForSqlLiteral = "0"
                Case Else
                    If IsNumeric(ValueText) Then
                        If CDbl(ValueText) = 0 Then
                            ConvertValueForSqlLiteral = "0"
                        Else
                            ConvertValueForSqlLiteral = "1"
                        End If
                    Else
                        GoTo ConversionError
                    End If
            End Select
            
        Case "datetime", "date", "smalldatetime"
            If Not IsDate(ValueText) Then GoTo ConversionError
            ConvertValueForSqlLiteral = "'" & Format$(CDate(ValueText), "yyyy-mm-dd hh:nn:ss") & "'"
            
        Case "uniqueidentifier"
            If Len(ValueText) > 50 Then GoTo ConversionError
            ConvertValueForSqlLiteral = "'" & EscapeSqlText(ValueText) & "'"
            
        Case "nvarchar", "varchar", "nchar", "char", "text", "ntext"
            If ColInfo.CharacterMaxLength > 0 Then
                If Len(ValueText) > ColInfo.CharacterMaxLength Then
                    HasTypeError = True
                    ErrorMessage = "Text value is too long for SQL column [" & ColInfo.ColumnName & "]. Max length=" & CStr(ColInfo.CharacterMaxLength) & ", actual length=" & CStr(Len(ValueText))
                    ConvertValueForSqlLiteral = "NULL"
                    Exit Function
                End If
            End If
            
            ConvertValueForSqlLiteral = "N'" & EscapeSqlText(ValueText) & "'"
            
        Case Else
            If ColInfo.CharacterMaxLength > 0 Then
                If Len(ValueText) > ColInfo.CharacterMaxLength Then
                    HasTypeError = True
                    ErrorMessage = "Text value is too long for SQL column [" & ColInfo.ColumnName & "]. Max length=" & CStr(ColInfo.CharacterMaxLength) & ", actual length=" & CStr(Len(ValueText))
                    ConvertValueForSqlLiteral = "NULL"
                    Exit Function
                End If
            End If
            
            ConvertValueForSqlLiteral = "N'" & EscapeSqlText(ValueText) & "'"
    End Select
    
    Exit Function

ConversionError:
    HasTypeError = True
    ErrorMessage = "Invalid value [" & ValueText & "] for SQL column [" & ColInfo.ColumnName & "] with type [" & ColInfo.DataType & "]"
    ConvertValueForSqlLiteral = "NULL"
End Function

Private Sub UpdateProgress( _
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

