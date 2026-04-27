VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Excel / SQL Data Tools"
   ClientHeight    =   9240
   ClientLeft      =   14040
   ClientTop       =   5640
   ClientWidth     =   5430
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9240
   ScaleWidth      =   5430
   Begin VB.Frame fraStatus 
      Caption         =   " Status "
      Height          =   1935
      Left            =   120
      TabIndex        =   24
      Top             =   7200
      Width           =   5175
      Begin VB.Label lblAppVersion 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1.1.0"
         Height          =   240
         Left            =   1320
         TabIndex        =   32
         Top             =   1440
         Width           =   435
      End
      Begin VB.Label lblAppVersionCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Version    :"
         Height          =   240
         Left            =   240
         TabIndex        =   31
         Top             =   1440
         Width           =   960
      End
      Begin VB.Label lblSelectedTable 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         Height          =   240
         Left            =   1320
         TabIndex        =   30
         Top             =   1080
         Width           =   2385
      End
      Begin VB.Label lblSelectedTableCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Table       :"
         Height          =   240
         Left            =   240
         TabIndex        =   29
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label lblSelectedDatabase 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         Height          =   240
         Left            =   1320
         TabIndex        =   28
         Top             =   720
         Width           =   2385
      End
      Begin VB.Label lblSelectedDatabseCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Database  :"
         Height          =   240
         Left            =   240
         TabIndex        =   27
         Top             =   720
         Width           =   990
      End
      Begin VB.Label lblConnectionStatus 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Not Connected"
         Height          =   240
         Left            =   1320
         TabIndex        =   26
         Top             =   360
         Width           =   2385
      End
      Begin VB.Label lblConnectionStatusCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Connection: "
         Height          =   240
         Left            =   240
         TabIndex        =   25
         Top             =   360
         Width           =   1080
      End
   End
   Begin VB.Frame fraTools 
      Caption         =   " Tools "
      Height          =   1575
      Left            =   120
      TabIndex        =   23
      Top             =   5640
      Width           =   5175
      Begin VB.CommandButton cmdOpenExport 
         Caption         =   "Export SQL to Excel"
         Height          =   960
         Left            =   2700
         TabIndex        =   11
         Top             =   360
         Width           =   2070
      End
      Begin VB.CommandButton cmdOpenImport 
         Caption         =   "Import Excel to SQL"
         Height          =   960
         Left            =   360
         TabIndex        =   10
         Top             =   360
         Width           =   2070
      End
   End
   Begin VB.Frame fraDatabaseTable 
      Caption         =   " Database and Table Selection "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2520
      Left            =   120
      TabIndex        =   18
      Top             =   3120
      Width           =   5175
      Begin VB.ComboBox cmbDatabase 
         Height          =   360
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   780
         Width           =   3375
      End
      Begin VB.ComboBox cmbTable 
         Height          =   360
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   2040
         Width           =   3375
      End
      Begin VB.TextBox txtSearchDatabase 
         Height          =   375
         Left            =   1680
         TabIndex        =   6
         Top             =   360
         Width           =   3375
      End
      Begin VB.TextBox txtSearchTable 
         Height          =   375
         Left            =   1680
         TabIndex        =   8
         Top             =   1560
         Width           =   3375
      End
      Begin VB.Label lblDatabase 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Database"
         Height          =   240
         Left            =   120
         TabIndex        =   22
         Top             =   780
         Width           =   795
      End
      Begin VB.Label lblTable 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Table"
         Height          =   240
         Left            =   120
         TabIndex        =   21
         Top             =   2040
         Width           =   480
      End
      Begin VB.Label lblSearchDatabase 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Search Database"
         Height          =   240
         Left            =   120
         TabIndex        =   20
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label lblSearchTable 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Search Table"
         Height          =   240
         Left            =   120
         TabIndex        =   19
         Top             =   1560
         Width           =   1140
      End
   End
   Begin VB.Frame FraSQLServer 
      Caption         =   " SQL Server Connection "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2520
      Left            =   120
      TabIndex        =   13
      Top             =   600
      Width           =   5175
      Begin VB.TextBox txtServerName 
         Height          =   375
         Left            =   1680
         TabIndex        =   0
         Top             =   360
         Width           =   3375
      End
      Begin VB.TextBox txtUserName 
         Height          =   375
         Left            =   1680
         TabIndex        =   2
         Top             =   1140
         Width           =   3375
      End
      Begin VB.TextBox txtPassword 
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   1680
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   1560
         Width           =   3375
      End
      Begin VB.ComboBox cmbAuthentication 
         Height          =   360
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   780
         Width           =   3375
      End
      Begin VB.CommandButton cmdConnect 
         Caption         =   "Connect"
         Height          =   360
         Left            =   3465
         TabIndex        =   5
         Top             =   2040
         Width           =   1590
      End
      Begin VB.CheckBox chkRememberConnectionSettings 
         Caption         =   "Remember connection settings"
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   2040
         Width           =   3015
      End
      Begin VB.Label lblServerName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Server Name"
         Height          =   240
         Left            =   240
         TabIndex        =   17
         Top             =   360
         Width           =   1125
      End
      Begin VB.Label lblUserName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User Name"
         Height          =   240
         Left            =   240
         TabIndex        =   16
         Top             =   1140
         Width           =   945
      End
      Begin VB.Label lblPassword 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
         Height          =   240
         Left            =   240
         TabIndex        =   15
         Top             =   1560
         Width           =   825
      End
      Begin VB.Label lblAuthentication 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Authentication"
         Height          =   240
         Left            =   240
         TabIndex        =   14
         Top             =   780
         Width           =   1215
      End
   End
   Begin VB.Label lblMainTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Excel / SQL Data Tools"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1260
      TabIndex        =   12
      Top             =   120
      Width           =   2910
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private mDatabaseBrowser As clsDatabaseBrowser
Private mSqlConnection As clsSqlServerConnection
Private mSuppressUiEvents As Boolean
Private mAllDatabases As Collection
Private mAllTables As Collection

Private Const SETTINGS_FILE_NAME As String = "AppSettings.ini"
Private Const SETTINGS_SECTION_CONNECTION As String = "Connection"

Private Sub cmbDatabase_Click()
    On Error GoTo ErrorHandler
    
    If mSuppressUiEvents Then Exit Sub
    
    ResetAfterTableSelectionChange
    
    EnsureAppContext
    gAppContext.SelectedDatabase = Trim$(cmbDatabase.Text)
    
    LoadTables
    RefreshMainUiState
    
    Exit Sub

ErrorHandler:
    MsgBox "Error loading tables." & vbCrLf & _
           Err.Number & " - " & Err.Description, vbExclamation, APP_NAME
End Sub

Private Sub cmbTable_Click()
    On Error GoTo ErrorHandler
    
    If cmbTable.ListIndex >= 0 Then
        UpdateContextFromTargetControls
        RefreshMainUiState
    End If
    
    Exit Sub

ErrorHandler:
    MsgBox "Error selecting table." & vbCrLf & _
           Err.Number & " - " & Err.Description, vbExclamation, APP_NAME
End Sub

Private Sub cmdConnect_Click()
    On Error GoTo ErrorHandler
    
    Screen.MousePointer = vbHourglass
    lblConnectionStatus.Caption = "Connecting..."
    
    If Not ValidateConnectionInputs() Then GoTo SafeExit
    
    If mSqlConnection.OpenConnection( _
        Trim$(txtServerName.Text), _
        cmbAuthentication.Text, _
        Trim$(txtUserName.Text), _
        txtPassword.Text) Then
        
        SetConnectedState
        LoadDatabases
        
        UpdateContextFromConnectionControls
        gAppContext.IsConnected = True
        ResetTargetInContext
        
        fraDatabaseTable.Enabled = True
        RefreshMainUiState
        
        If chkRememberConnectionSettings.Value = vbChecked Then
            SaveConnectionSettings
        Else
            ClearSavedConnectionSettings
        End If
        
        MsgBox "Connection established successfully.", vbInformation, APP_NAME
    End If
    
SafeExit:
    Screen.MousePointer = vbDefault
    Exit Sub

ErrorHandler:
    Screen.MousePointer = vbDefault
    
    ResetAfterConnectionSettingsChange
    
    EnsureAppContext
    gAppContext.ResetConnectionState
    RefreshMainUiState
    
    MsgBox "Connection failed." & vbCrLf & _
           Err.Number & " - " & Err.Description, vbCritical, APP_NAME
End Sub

Private Sub cmdOpenExport_Click()
    EnsureAppContext
    
    If Not gAppContext.IsConnected Then
        MsgBox "Please connect to SQL Server first.", vbExclamation, APP_NAME
        Exit Sub
    End If
    
    If Not gAppContext.HasSelectedTarget Then
        MsgBox "Please select a target database and table first.", vbExclamation, APP_NAME
        Exit Sub
    End If
    
    frmExportSqlToExcel.Show
End Sub

Private Sub cmdOpenImport_Click()
    EnsureAppContext
    
    If Not gAppContext.IsConnected Then
        MsgBox "Please connect to SQL Server first.", vbExclamation, APP_NAME
        Exit Sub
    End If
    
    If Not gAppContext.HasSelectedTarget Then
        MsgBox "Please select a target database and table first.", vbExclamation, APP_NAME
        Exit Sub
    End If
    
    frmImportExcelToSql.Show
End Sub

Private Sub Form_Load()
    InitializeForm
End Sub

Private Sub InitializeForm()
    On Error GoTo ErrorHandler
    
    mSuppressUiEvents = True
    
    EnsureAppContext
    EnsureMainObjects
    
    Me.Caption = "Excel / SQL Data Tools"
    
    InitializeAuthenticationCombo
    
    lblAppVersion.Caption = "1.1.0"
    
    fraDatabaseTable.Enabled = False
    fraTools.Enabled = False
    cmdOpenImport.Enabled = False
    cmdOpenExport.Enabled = False
    
    RefreshMainUiState
    LoadConnectionSettings
    
    mSuppressUiEvents = False
    Exit Sub

ErrorHandler:
    mSuppressUiEvents = False
    MsgBox "Error initializing main form." & vbCrLf & _
           Err.Number & " - " & Err.Description, vbCritical, APP_NAME
End Sub

Private Sub EnsureAppContext()
    If gAppContext Is Nothing Then
        Set gAppContext = New clsAppContext
    End If
End Sub

Private Sub UpdateContextFromConnectionControls()
    EnsureAppContext
    
    gAppContext.ServerName = Trim$(txtServerName.Text)
    gAppContext.AuthenticationType = Trim$(cmbAuthentication.Text)
    gAppContext.UserName = Trim$(txtUserName.Text)
    gAppContext.Password = txtPassword.Text
    gAppContext.RememberConnectionSettings = (chkRememberConnectionSettings.Value = vbChecked)
End Sub

Private Sub UpdateContextFromTargetControls()
    EnsureAppContext
    
    gAppContext.SelectedDatabase = Trim$(cmbDatabase.Text)
    gAppContext.SelectedSchema = GetSelectedSchemaName()
    gAppContext.SelectedTable = GetSelectedTableName()
End Sub

Private Sub ResetTargetInContext()
    EnsureAppContext
    
    gAppContext.SelectedDatabase = ""
    gAppContext.SelectedSchema = ""
    gAppContext.SelectedTable = ""
End Sub

Private Sub RefreshMainStatus()
    EnsureAppContext
    
    If gAppContext.IsConnected Then
        lblConnectionStatus.Caption = "Connected"
    Else
        lblConnectionStatus.Caption = "Not connected"
    End If
    
    If Len(gAppContext.SelectedDatabase) > 0 Then
        lblSelectedDatabase.Caption = gAppContext.SelectedDatabase
    Else
        lblSelectedDatabase.Caption = "-"
    End If
    
    If Len(gAppContext.GetFullTargetName()) > 0 Then
        lblSelectedTable.Caption = gAppContext.GetFullTargetName()
    Else
        lblSelectedTable.Caption = "-"
    End If
End Sub

Private Sub RefreshToolButtonsState()
    EnsureAppContext
    
    If gAppContext.IsConnected And gAppContext.HasSelectedTarget Then
        fraTools.Enabled = True
        cmdOpenImport.Enabled = True
        cmdOpenExport.Enabled = True
    Else
        fraTools.Enabled = False
        cmdOpenImport.Enabled = False
        cmdOpenExport.Enabled = False
    End If
End Sub

Private Sub RefreshMainUiState()
    RefreshMainStatus
    RefreshToolButtonsState
End Sub

Private Sub cmbAuthentication_Click()
    On Error GoTo ErrorHandler
    
    If mSuppressUiEvents Then Exit Sub
    
    UpdateAuthenticationControls
    ResetAfterConnectionSettingsChange
    UpdateContextFromConnectionControls
    ResetTargetInContext
    RefreshMainUiState
    
    Exit Sub

ErrorHandler:
    MsgBox "Error changing authentication mode." & vbCrLf & _
           Err.Number & " - " & Err.Description, vbExclamation, APP_NAME
End Sub

Private Sub ResetAfterConnectionSettingsChange()
    On Error Resume Next
    
    mSuppressUiEvents = True
    
    If Not mSqlConnection Is Nothing Then
        mSqlConnection.CloseConnection
    End If
    
    cmbDatabase.Clear
    cmbTable.Clear
    
    txtSearchDatabase.Text = ""
    txtSearchTable.Text = ""
    
    fraDatabaseTable.Enabled = False
    fraTools.Enabled = False
    
    cmdOpenImport.Enabled = False
    cmdOpenExport.Enabled = False
    
    EnsureAppContext
    gAppContext.ResetConnectionState
    
    RefreshMainUiState
    
    mSuppressUiEvents = False
End Sub

Private Sub ResetAfterTableSelectionChange()
    On Error Resume Next
    
    mSuppressUiEvents = True
    
    cmbTable.Clear
    txtSearchTable.Text = ""
    
    fraTools.Enabled = False
    cmdOpenImport.Enabled = False
    cmdOpenExport.Enabled = False
    
    EnsureAppContext
    gAppContext.SelectedSchema = ""
    gAppContext.SelectedTable = ""
    
    RefreshMainUiState
    
    mSuppressUiEvents = False
End Sub

Private Function ValidateConnectionInputs() As Boolean
    ValidateConnectionInputs = False
    
    If Len(Trim$(txtServerName.Text)) = 0 Then
        MsgBox "Please enter the SQL Server name.", vbExclamation, APP_NAME
        txtServerName.SetFocus
        Exit Function
    End If
    
    If cmbAuthentication.ListIndex < 0 Or Len(Trim$(cmbAuthentication.Text)) = 0 Then
        MsgBox "Please select an authentication type.", vbExclamation, APP_NAME
        cmbAuthentication.SetFocus
        Exit Function
    End If
    
    If StrComp(Trim$(cmbAuthentication.Text), AUTH_SQL_SERVER, vbTextCompare) = 0 Then
        If Len(Trim$(txtUserName.Text)) = 0 Then
            MsgBox "Please enter the SQL Server user name.", vbExclamation, APP_NAME
            txtUserName.SetFocus
            Exit Function
        End If
        
        If Len(txtPassword.Text) = 0 Then
            MsgBox "Please enter the SQL Server password.", vbExclamation, APP_NAME
            txtPassword.SetFocus
            Exit Function
        End If
    End If
    
    ValidateConnectionInputs = True
End Function

Private Sub SetConnectedState()
    fraDatabaseTable.Enabled = True
    
    lblConnectionStatus.Caption = "Connected"
    
    cmdOpenImport.Enabled = False
    cmdOpenExport.Enabled = False
    fraTools.Enabled = False
    
    lblSelectedDatabase.Caption = "-"
    lblSelectedTable.Caption = "-"
End Sub

Private Sub LoadDatabases()
    On Error GoTo ErrorHandler
    
    Dim Rs As ADODB.Recordset
    Dim DbName As String
    
    If mSqlConnection Is Nothing Then Exit Sub
    If mSqlConnection.Connection Is Nothing Then Exit Sub
    If mSqlConnection.Connection.State <> adStateOpen Then Exit Sub
    
    cmbDatabase.Clear
    cmbTable.Clear
    txtSearchDatabase.Text = ""
    txtSearchTable.Text = ""
    
    Set mAllDatabases = New Collection
    Set mAllTables = New Collection
    
    Set Rs = mDatabaseBrowser.GetUserDatabases(mSqlConnection.Connection)
    
    Do While Not Rs.EOF
        DbName = NzString(Rs.Fields("name").Value)
        
        If Len(DbName) > 0 Then
            mAllDatabases.Add DbName
            cmbDatabase.AddItem DbName
        End If
        
        Rs.MoveNext
    Loop
    
    Rs.Close
    Set Rs = Nothing
    
    If cmbDatabase.ListCount > 0 Then
        cmbDatabase.ListIndex = 0
    End If
    
    Exit Sub

ErrorHandler:
    On Error Resume Next
    
    If Not Rs Is Nothing Then
        If Rs.State = adStateOpen Then Rs.Close
    End If
    Set Rs = Nothing
    
    MsgBox "Failed to load databases." & vbCrLf & _
           Err.Number & " - " & Err.Description, vbCritical, APP_NAME
End Sub

Private Sub SaveConnectionSettings()
    Dim SettingsFilePath As String
    
    SettingsFilePath = GetSettingsFilePath()
    
    WriteIniValue SETTINGS_SECTION_CONNECTION, "RememberSettings", _
                  IIf(chkRememberConnectionSettings.Value = vbChecked, "1", "0"), _
                  SettingsFilePath
    
    WriteIniValue SETTINGS_SECTION_CONNECTION, "ServerName", _
                  Trim$(txtServerName.Text), _
                  SettingsFilePath
    
    WriteIniValue SETTINGS_SECTION_CONNECTION, "AuthenticationType", _
                  Trim$(cmbAuthentication.Text), _
                  SettingsFilePath
    
    WriteIniValue SETTINGS_SECTION_CONNECTION, "UserName", _
                  Trim$(txtUserName.Text), _
                  SettingsFilePath
End Sub

Private Sub ClearSavedConnectionSettings()
    Dim SettingsFilePath As String
    
    SettingsFilePath = GetSettingsFilePath()
    
    DeleteIniKey SETTINGS_SECTION_CONNECTION, "RememberSettings", SettingsFilePath
    DeleteIniKey SETTINGS_SECTION_CONNECTION, "ServerName", SettingsFilePath
    DeleteIniKey SETTINGS_SECTION_CONNECTION, "AuthenticationType", SettingsFilePath
    DeleteIniKey SETTINGS_SECTION_CONNECTION, "UserName", SettingsFilePath
End Sub

Private Function GetSettingsFilePath() As String
    GetSettingsFilePath = App.Path
    
    If Right$(GetSettingsFilePath, 1) <> "\" Then
        GetSettingsFilePath = GetSettingsFilePath & "\"
    End If
    
    GetSettingsFilePath = GetSettingsFilePath & SETTINGS_FILE_NAME
End Function

Private Sub InitializeAuthenticationCombo()
    cmbAuthentication.Clear
    
    cmbAuthentication.AddItem AUTH_WINDOWS
    cmbAuthentication.AddItem AUTH_SQL_SERVER
    
    cmbAuthentication.ListIndex = 0
End Sub

Private Sub LoadConnectionSettings()
    On Error GoTo ErrorHandler
    
    Dim SettingsFilePath As String
    Dim RememberValue As String
    Dim ServerName As String
    Dim AuthenticationType As String
    Dim UserName As String
    
    mSuppressUiEvents = True
    
    SettingsFilePath = GetSettingsFilePath()
    
    RememberValue = ReadIniValue(SETTINGS_SECTION_CONNECTION, "RememberSettings", "0", SettingsFilePath)
    
    If RememberValue <> "1" Then
        chkRememberConnectionSettings.Value = vbUnchecked
        mSuppressUiEvents = False
        Exit Sub
    End If
    
    chkRememberConnectionSettings.Value = vbChecked
    
    ServerName = ReadIniValue(SETTINGS_SECTION_CONNECTION, "ServerName", "", SettingsFilePath)
    AuthenticationType = ReadIniValue(SETTINGS_SECTION_CONNECTION, "AuthenticationType", AUTH_WINDOWS, SettingsFilePath)
    UserName = ReadIniValue(SETTINGS_SECTION_CONNECTION, "UserName", "", SettingsFilePath)
    
    txtServerName.Text = ServerName
    txtUserName.Text = UserName
    txtPassword.Text = ""
    
    If StrComp(AuthenticationType, AUTH_SQL_SERVER, vbTextCompare) = 0 Then
        cmbAuthentication.ListIndex = 1
    Else
        cmbAuthentication.ListIndex = 0
    End If
    
    UpdateAuthenticationControls
    UpdateContextFromConnectionControls
    
    mSuppressUiEvents = False
    Exit Sub

ErrorHandler:
    mSuppressUiEvents = False
    chkRememberConnectionSettings.Value = vbUnchecked
End Sub

Private Function GetSelectedSchemaName() As String
    Dim FullName As String
    Dim DotPos As Long
    
    FullName = Trim$(cmbTable.Text)
    DotPos = InStr(1, FullName, ".")
    
    If DotPos > 0 Then
        GetSelectedSchemaName = Left$(FullName, DotPos - 1)
    Else
        GetSelectedSchemaName = "dbo"
    End If
End Function

Private Function GetSelectedTableName() As String
    Dim FullName As String
    Dim DotPos As Long
    
    FullName = Trim$(cmbTable.Text)
    
    If Len(FullName) = 0 Then
        GetSelectedTableName = ""
        Exit Function
    End If
    
    DotPos = InStr(1, FullName, ".")
    
    If DotPos > 0 Then
        GetSelectedTableName = Mid$(FullName, DotPos + 1)
    Else
        GetSelectedTableName = FullName
    End If
End Function

Private Sub UpdateAuthenticationControls()
    Dim IsSqlAuthentication As Boolean
    
    IsSqlAuthentication = (StrComp(Trim$(cmbAuthentication.Text), AUTH_SQL_SERVER, vbTextCompare) = 0)
    
    txtUserName.Enabled = IsSqlAuthentication
    txtPassword.Enabled = IsSqlAuthentication
    
    If IsSqlAuthentication Then
        txtUserName.BackColor = vbWindowBackground
        txtPassword.BackColor = vbWindowBackground
    Else
        txtUserName.Text = ""
        txtPassword.Text = ""
        
        txtUserName.BackColor = &H8000000F
        txtPassword.BackColor = &H8000000F
    End If
End Sub

Private Sub EnsureMainObjects()
    If mSqlConnection Is Nothing Then
        Set mSqlConnection = New clsSqlServerConnection
    End If
    
    If mDatabaseBrowser Is Nothing Then
        Set mDatabaseBrowser = New clsDatabaseBrowser
    End If
    
    If mAllDatabases Is Nothing Then
        Set mAllDatabases = New Collection
    End If
    
    If mAllTables Is Nothing Then
        Set mAllTables = New Collection
    End If
End Sub

Private Sub LoadTables()
    On Error GoTo ErrorHandler
    
    Dim Rs As ADODB.Recordset
    Dim DisplayText As String
    
    cmbTable.Clear
    txtSearchTable.Text = ""
    Set mAllTables = New Collection
    
    If cmbDatabase.ListIndex < 0 Or Len(Trim$(cmbDatabase.Text)) = 0 Then
        MsgBox "Please select a database first.", vbExclamation, APP_NAME
        cmbDatabase.SetFocus
        Exit Sub
    End If
    
    Set Rs = mDatabaseBrowser.GetBaseTables(mSqlConnection.Connection, cmbDatabase.Text)
    
    Do While Not Rs.EOF
        DisplayText = NzString(Rs.Fields("TABLE_SCHEMA").Value) & "." & NzString(Rs.Fields("TABLE_NAME").Value)
        
        If Len(Trim$(DisplayText)) > 0 Then
            mAllTables.Add DisplayText
            cmbTable.AddItem DisplayText
        End If
        
        Rs.MoveNext
    Loop
    
    Rs.Close
    Set Rs = Nothing
    
    If cmbTable.ListCount > 0 Then
        cmbTable.ListIndex = 0
    End If
    
    Exit Sub

ErrorHandler:
    On Error Resume Next
    
    If Not Rs Is Nothing Then
        If Rs.State = adStateOpen Then Rs.Close
    End If
    Set Rs = Nothing
    
    MsgBox "Failed to load tables." & vbCrLf & _
           Err.Number & " - " & Err.Description, vbCritical, APP_NAME
End Sub

Private Sub txtSearchDatabase_Change()
    If mSuppressUiEvents Then Exit Sub
    
    If fraDatabaseTable.Enabled Then
        FilterDatabases
    End If
End Sub

Private Sub txtSearchTable_Change()
    If mSuppressUiEvents Then Exit Sub
    
    If fraDatabaseTable.Enabled Then
        FilterTables
    End If
End Sub

Private Sub FilterDatabases()
    Dim i As Long
    Dim DbName As String
    Dim SearchText As String
    
    SearchText = Trim$(txtSearchDatabase.Text)
    
    cmbDatabase.Clear
    
    If mAllDatabases Is Nothing Then Exit Sub
    
    For i = 1 To mAllDatabases.Count
        DbName = CStr(mAllDatabases(i))
        
        If ContainsText(DbName, SearchText) Then
            cmbDatabase.AddItem DbName
        End If
    Next i
    
    If cmbDatabase.ListCount > 0 Then
        cmbDatabase.ListIndex = 0
    End If
End Sub

Private Sub FilterTables()
    Dim i As Long
    Dim TableName As String
    Dim SearchText As String
    
    SearchText = Trim$(txtSearchTable.Text)
    
    cmbTable.Clear
    
    If mAllTables Is Nothing Then Exit Sub
    
    For i = 1 To mAllTables.Count
        TableName = CStr(mAllTables(i))
        
        If ContainsText(TableName, SearchText) Then
            cmbTable.AddItem TableName
        End If
    Next i
    
    If cmbTable.ListCount > 0 Then
        cmbTable.ListIndex = 0
    End If
End Sub

Private Function ContainsText(ByVal SourceText As String, ByVal SearchText As String) As Boolean
    If Len(Trim$(SearchText)) = 0 Then
        ContainsText = True
    Else
        ContainsText = (InStr(1, SourceText, SearchText, vbTextCompare) > 0)
    End If
End Function

Private Function NzString(ByVal Value As Variant) As String
    If IsNull(Value) Or IsEmpty(Value) Then
        NzString = ""
    Else
        NzString = CStr(Value)
    End If
End Function

Private Sub txtServerName_LostFocus()
    If mSuppressUiEvents Then Exit Sub
    If gAppContext Is Nothing Then Exit Sub
    
    If gAppContext.IsConnected Then
        If StrComp(Trim$(txtServerName.Text), gAppContext.ServerName, vbTextCompare) <> 0 Then
            ResetAfterConnectionSettingsChange
            RefreshMainUiState
        End If
    End If
End Sub

Public Property Get SqlConnectionManager() As clsSqlServerConnection
    Set SqlConnectionManager = mSqlConnection
End Property
