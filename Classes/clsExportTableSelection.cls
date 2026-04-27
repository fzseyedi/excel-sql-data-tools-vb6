VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsExportTableSelection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mSchemaName As String
Private mTableName As String
Private mIsChecked As Boolean
Private mAllFields As Collection
Private mSelectedFields As Collection
Private mIsFieldsLoaded As Boolean

Private Sub Class_Initialize()
    Set mAllFields = New Collection
    Set mSelectedFields = New Collection
    mIsFieldsLoaded = False
End Sub

Public Property Get SchemaName() As String
    SchemaName = mSchemaName
End Property

Public Property Let SchemaName(ByVal Value As String)
    mSchemaName = Trim$(Value)
End Property

Public Property Get TableName() As String
    TableName = mTableName
End Property

Public Property Let TableName(ByVal Value As String)
    mTableName = Trim$(Value)
End Property

Public Property Get IsChecked() As Boolean
    IsChecked = mIsChecked
End Property

Public Property Let IsChecked(ByVal Value As Boolean)
    mIsChecked = Value
End Property

Public Property Get IsFieldsLoaded() As Boolean
    IsFieldsLoaded = mIsFieldsLoaded
End Property

Public Property Let IsFieldsLoaded(ByVal Value As Boolean)
    mIsFieldsLoaded = Value
End Property

Public Property Get FullTableName() As String
    If Len(mSchemaName) > 0 Then
        FullTableName = mSchemaName & "." & mTableName
    Else
        FullTableName = mTableName
    End If
End Property

Public Property Get AllFields() As Collection
    Set AllFields = mAllFields
End Property

Public Property Get SelectedFields() As Collection
    Set SelectedFields = mSelectedFields
End Property

Public Sub ClearFields()
    Set mAllFields = New Collection
    Set mSelectedFields = New Collection
    mIsFieldsLoaded = False
End Sub

Public Sub AddField(ByVal FieldName As String, Optional ByVal IsSelected As Boolean = True)
    If Len(Trim$(FieldName)) = 0 Then Exit Sub
    
    mAllFields.Add Trim$(FieldName)
    
    If IsSelected Then
        AddSelectedField Trim$(FieldName)
    End If
End Sub

Public Sub AddSelectedField(ByVal FieldName As String)
    If Len(Trim$(FieldName)) = 0 Then Exit Sub
    
    If Not IsFieldSelected(FieldName) Then
        mSelectedFields.Add Trim$(FieldName)
    End If
End Sub

Public Sub RemoveSelectedField(ByVal FieldName As String)
    Dim i As Long
    
    For i = mSelectedFields.Count To 1 Step -1
        If StrComp(CStr(mSelectedFields(i)), Trim$(FieldName), vbTextCompare) = 0 Then
            mSelectedFields.Remove i
            Exit For
        End If
    Next i
End Sub

Public Function IsFieldSelected(ByVal FieldName As String) As Boolean
    Dim i As Long
    
    For i = 1 To mSelectedFields.Count
        If StrComp(CStr(mSelectedFields(i)), Trim$(FieldName), vbTextCompare) = 0 Then
            IsFieldSelected = True
            Exit Function
        End If
    Next i
    
    IsFieldSelected = False
End Function

Public Sub ClearSelectedFields()
    Set mSelectedFields = New Collection
End Sub

Public Sub RebuildSelectedFieldsFromListBox(ByVal SourceList As ListBox)
    Dim i As Long
    
    ClearSelectedFields
    
    For i = 0 To SourceList.ListCount - 1
        If SourceList.Selected(i) Then
            AddSelectedField SourceList.List(i)
        End If
    Next i
End Sub

