VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAppContext"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mServerName As String
Private mAuthenticationType As String
Private mUserName As String
Private mPassword As String
Private mRememberConnectionSettings As Boolean
Private mSelectedDatabase As String
Private mSelectedSchema As String
Private mSelectedTable As String
Private mExcelReadMode As Long
Private mIsConnected As Boolean

Public Property Get ServerName() As String
    ServerName = mServerName
End Property

Public Property Let ServerName(ByVal Value As String)
    mServerName = Trim$(Value)
End Property

Public Property Get AuthenticationType() As String
    AuthenticationType = mAuthenticationType
End Property

Public Property Let AuthenticationType(ByVal Value As String)
    mAuthenticationType = Trim$(Value)
End Property

Public Property Get UserName() As String
    UserName = mUserName
End Property

Public Property Let UserName(ByVal Value As String)
    mUserName = Trim$(Value)
End Property

Public Property Get Password() As String
    Password = mPassword
End Property

Public Property Let Password(ByVal Value As String)
    mPassword = Value
End Property

Public Property Get RememberConnectionSettings() As Boolean
    RememberConnectionSettings = mRememberConnectionSettings
End Property

Public Property Let RememberConnectionSettings(ByVal Value As Boolean)
    mRememberConnectionSettings = Value
End Property

Public Property Get SelectedDatabase() As String
    SelectedDatabase = mSelectedDatabase
End Property

Public Property Let SelectedDatabase(ByVal Value As String)
    mSelectedDatabase = Trim$(Value)
End Property

Public Property Get SelectedSchema() As String
    SelectedSchema = mSelectedSchema
End Property

Public Property Let SelectedSchema(ByVal Value As String)
    mSelectedSchema = Trim$(Value)
End Property

Public Property Get SelectedTable() As String
    SelectedTable = mSelectedTable
End Property

Public Property Let SelectedTable(ByVal Value As String)
    mSelectedTable = Trim$(Value)
End Property

Public Property Get ExcelReadMode() As Long
    ExcelReadMode = mExcelReadMode
End Property

Public Property Let ExcelReadMode(ByVal Value As Long)
    mExcelReadMode = Value
End Property

Public Property Get IsConnected() As Boolean
    IsConnected = mIsConnected
End Property

Public Property Let IsConnected(ByVal Value As Boolean)
    mIsConnected = Value
End Property

Public Sub ResetConnectionState()
    mIsConnected = False
    mSelectedDatabase = ""
    mSelectedSchema = ""
    mSelectedTable = ""
End Sub

Public Sub ResetAll()
    mServerName = ""
    mAuthenticationType = ""
    mUserName = ""
    mPassword = ""
    mRememberConnectionSettings = False
    mSelectedDatabase = ""
    mSelectedSchema = ""
    mSelectedTable = ""
    mExcelReadMode = 0
    mIsConnected = False
End Sub

Public Function HasSelectedTarget() As Boolean
    HasSelectedTarget = (Len(mSelectedDatabase) > 0 And Len(mSelectedTable) > 0)
End Function

Public Function GetFullTargetName() As String
    If Len(mSelectedSchema) > 0 Then
        GetFullTargetName = mSelectedSchema & "." & mSelectedTable
    Else
        GetFullTargetName = mSelectedTable
    End If
End Function
