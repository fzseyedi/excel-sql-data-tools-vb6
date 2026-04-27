VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmImportExcelToSql 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Import Excel Data To Sql"
   ClientHeight    =   9015
   ClientLeft      =   16665
   ClientTop       =   8175
   ClientWidth     =   14880
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
   ScaleHeight     =   9015
   ScaleWidth      =   14880
   Begin MSComDlg.CommonDialog dlgExcelFile 
      Left            =   7200
      Top             =   4080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin TabDlg.SSTab sstMain 
      Height          =   6615
      Left            =   120
      TabIndex        =   18
      Top             =   1560
      Width           =   14655
      _ExtentX        =   25850
      _ExtentY        =   11668
      _Version        =   393216
      Tab             =   2
      TabHeight       =   520
      TabCaption(0)   =   "Mapping"
      TabPicture(0)   =   "frmImportExcelToSql.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "lblExcelPreview"
      Tab(0).Control(1)=   "grdMapping"
      Tab(0).Control(2)=   "cmdAutoMatch"
      Tab(0).Control(3)=   "cmdLoadSavedMapping"
      Tab(0).Control(4)=   "cmdSaveMapping"
      Tab(0).Control(5)=   "cmdValidateMapping"
      Tab(0).Control(6)=   "cmbGridSqlColumns"
      Tab(0).Control(7)=   "grdExcelPreview"
      Tab(0).ControlCount=   8
      TabCaption(1)   =   "Import Options"
      TabPicture(1)   =   "frmImportExcelToSql.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraImportOption"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Execution"
      TabPicture(2)   =   "frmImportExcelToSql.frx":0038
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "fraExceution"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      Begin MSFlexGridLib.MSFlexGrid grdExcelPreview 
         Height          =   1935
         Left            =   -74880
         TabIndex        =   20
         Top             =   4200
         Width           =   14415
         _ExtentX        =   25426
         _ExtentY        =   3413
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.ComboBox cmbGridSqlColumns 
         Height          =   360
         Left            =   -66840
         Style           =   2  'Dropdown List
         TabIndex        =   57
         Top             =   6150
         Visible         =   0   'False
         Width           =   3855
      End
      Begin VB.Frame fraExceution 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6135
         Left            =   120
         TabIndex        =   30
         Top             =   360
         Width           =   14415
         Begin VB.CommandButton cmdCancelImport 
            Caption         =   "Cancel Import"
            Enabled         =   0   'False
            Height          =   360
            Left            =   1920
            TabIndex        =   32
            Top             =   360
            Width           =   1590
         End
         Begin VB.TextBox txtStatus 
            Height          =   2535
            Left            =   360
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   35
            Top             =   3480
            Width           =   13935
         End
         Begin MSComctlLib.ProgressBar prgImport 
            Height          =   375
            Left            =   240
            TabIndex        =   34
            Top             =   840
            Width           =   14055
            _ExtentX        =   24791
            _ExtentY        =   661
            _Version        =   393216
            Appearance      =   1
            Scrolling       =   1
         End
         Begin VB.CommandButton cmdExit 
            Caption         =   "Exit"
            Height          =   360
            Left            =   12720
            TabIndex        =   33
            Top             =   360
            Width           =   1590
         End
         Begin VB.CommandButton cmdStartImport 
            Caption         =   "Start Import"
            Height          =   360
            Left            =   240
            TabIndex        =   31
            Top             =   360
            Width           =   1590
         End
         Begin VB.Label lblErrorCount 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            ForeColor       =   &H000000FF&
            Height          =   315
            Left            =   1680
            TabIndex        =   54
            Top             =   3120
            Width           =   2190
         End
         Begin VB.Label lblErrorCaption 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Errors"
            ForeColor       =   &H000000FF&
            Height          =   240
            Left            =   360
            TabIndex        =   53
            Top             =   3150
            Width           =   525
         End
         Begin VB.Label lblSkipCount 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            Height          =   315
            Left            =   1680
            TabIndex        =   52
            Top             =   2685
            Width           =   2190
         End
         Begin VB.Label lblSkippedCaption 
            BackStyle       =   0  'Transparent
            Caption         =   "Skipped"
            Height          =   315
            Left            =   360
            TabIndex        =   51
            Top             =   2685
            Width           =   990
         End
         Begin VB.Label lblSuccessCount 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            Height          =   315
            Left            =   1680
            TabIndex        =   50
            Top             =   2205
            Width           =   2190
         End
         Begin VB.Label lblSuccesCaption 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Succes"
            Height          =   240
            Left            =   360
            TabIndex        =   49
            Top             =   2235
            Width           =   600
         End
         Begin VB.Label lblRowProgress 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0 / 0"
            Height          =   315
            Left            =   1680
            TabIndex        =   48
            Top             =   1770
            Width           =   2175
         End
         Begin VB.Label lblRowProgressCaption 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Row Progress"
            Height          =   240
            Left            =   360
            TabIndex        =   47
            Top             =   1770
            Width           =   1185
         End
         Begin VB.Label lblCurrentStep 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Ready"
            Height          =   315
            Left            =   1680
            TabIndex        =   46
            Top             =   1320
            Width           =   12615
         End
         Begin VB.Label lblCurrentStepCaption 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Current Step"
            Height          =   240
            Left            =   360
            TabIndex        =   45
            Top             =   1320
            Width           =   1095
         End
      End
      Begin VB.Frame fraImportOption 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6135
         Left            =   -74880
         TabIndex        =   25
         Top             =   360
         Width           =   14415
         Begin VB.CheckBox chkAutoLoadSavedMapping 
            Caption         =   "Auto load saved mapping if available"
            Height          =   495
            Left            =   240
            TabIndex        =   29
            Top             =   1800
            Width           =   3615
         End
         Begin VB.CheckBox chkContinueOnDuplicate 
            Caption         =   "Continue on duplicate rows"
            Height          =   495
            Left            =   240
            TabIndex        =   28
            Top             =   1320
            Width           =   3615
         End
         Begin VB.CheckBox chkContinueOnTypeError 
            Caption         =   "Continue on data type errors"
            Height          =   495
            Left            =   240
            TabIndex        =   27
            Top             =   840
            Width           =   3615
         End
         Begin VB.CheckBox chkDeleteExisting 
            Caption         =   "Delete exisiting rows before import"
            Height          =   495
            Left            =   240
            TabIndex        =   26
            Top             =   360
            Width           =   3615
         End
         Begin VB.Label lblOptionsInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "These options effect how import errors and saved mappings are handled."
            Height          =   240
            Left            =   240
            TabIndex        =   44
            Top             =   2520
            Width           =   6315
         End
      End
      Begin VB.CommandButton cmdValidateMapping 
         Caption         =   "Validate Mapping"
         Height          =   360
         Left            =   -70800
         TabIndex        =   23
         Top             =   6150
         Width           =   1830
      End
      Begin VB.CommandButton cmdSaveMapping 
         Caption         =   "Save Mapping"
         Height          =   360
         Left            =   -68880
         TabIndex        =   24
         Top             =   6150
         Width           =   1830
      End
      Begin VB.CommandButton cmdLoadSavedMapping 
         Caption         =   "Load Saved Mapping"
         Height          =   360
         Left            =   -72960
         TabIndex        =   22
         Top             =   6150
         Width           =   2070
      End
      Begin VB.CommandButton cmdAutoMatch 
         Caption         =   "Auto Match"
         Height          =   360
         Left            =   -74880
         TabIndex        =   21
         Top             =   6150
         Width           =   1830
      End
      Begin MSFlexGridLib.MSFlexGrid grdMapping 
         Height          =   3495
         Left            =   -74880
         TabIndex        =   19
         Top             =   480
         Width           =   14415
         _ExtentX        =   25426
         _ExtentY        =   6165
         _Version        =   393216
         ScrollTrack     =   -1  'True
         SelectionMode   =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblExcelPreview 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Excel preview(First Rows)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -74880
         TabIndex        =   60
         Top             =   3960
         Width           =   1860
      End
   End
   Begin VB.Frame fraExcel 
      Caption         =   "Excel File"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   14655
      Begin VB.OptionButton optReadByExcel 
         Caption         =   "Read via Excel Automation"
         Height          =   375
         Left            =   9720
         TabIndex        =   16
         Top             =   600
         Value           =   -1  'True
         Width           =   2655
      End
      Begin VB.OptionButton optReadByAdo 
         Caption         =   "Read via ADO/OLEDB"
         Height          =   375
         Left            =   7440
         TabIndex        =   15
         Top             =   600
         Width           =   2175
      End
      Begin VB.CommandButton cmdLoadExcelColumns 
         Caption         =   "Load Excel Columns"
         Height          =   360
         Left            =   12720
         TabIndex        =   17
         Top             =   600
         Width           =   1830
      End
      Begin VB.CommandButton cmdBrowseExcel 
         Caption         =   "Browse"
         Height          =   360
         Left            =   4440
         TabIndex        =   14
         Top             =   700
         Width           =   990
      End
      Begin VB.TextBox txtExcelFile 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   13
         Top             =   700
         Width           =   4215
      End
      Begin VB.Label lblExcelFile 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Excel File"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   43
         Top             =   360
         Width           =   690
      End
   End
   Begin VB.Frame fraDatabaseTable 
      Caption         =   "Database and Table Selection"
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
      Left            =   5400
      TabIndex        =   6
      Top             =   -120
      Visible         =   0   'False
      Width           =   4575
      Begin VB.TextBox txtSearchTable 
         Height          =   375
         Left            =   1680
         TabIndex        =   10
         Top             =   1560
         Width           =   2775
      End
      Begin VB.TextBox txtSearchDatabase 
         Height          =   375
         Left            =   1680
         TabIndex        =   7
         Top             =   360
         Width           =   2775
      End
      Begin VB.CommandButton cmdLoadTables 
         Caption         =   "Load Tables"
         Height          =   360
         Left            =   2850
         TabIndex        =   9
         Top             =   1140
         Width           =   1590
      End
      Begin VB.ComboBox cmbTable 
         Height          =   360
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   2040
         Width           =   2775
      End
      Begin VB.ComboBox cmbDatabase 
         Height          =   360
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   780
         Width           =   2775
      End
      Begin VB.Label lblSearchTable 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Search Table"
         Height          =   240
         Left            =   120
         TabIndex        =   59
         Top             =   1560
         Width           =   1140
      End
      Begin VB.Label lblSearchDatabase 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Search Database"
         Height          =   240
         Left            =   120
         TabIndex        =   58
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label lblTable 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Table"
         Height          =   240
         Left            =   120
         TabIndex        =   42
         Top             =   2040
         Width           =   480
      End
      Begin VB.Label lblDatabase 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Database"
         Height          =   240
         Left            =   120
         TabIndex        =   41
         Top             =   780
         Width           =   795
      End
   End
   Begin VB.Frame FraConnection 
      Caption         =   "SQL Server Connection"
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
      TabIndex        =   36
      Top             =   -120
      Visible         =   0   'False
      Width           =   5175
      Begin VB.CheckBox chkRememberConnectionSettings 
         Caption         =   "Remember connection settings"
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   2040
         Width           =   3015
      End
      Begin VB.CommandButton cmdConnect 
         Caption         =   "Connect"
         Height          =   360
         Left            =   3465
         TabIndex        =   5
         Top             =   2040
         Width           =   1590
      End
      Begin VB.ComboBox cmbAuthentication 
         Height          =   360
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   780
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
      Begin VB.TextBox txtUserName 
         Height          =   375
         Left            =   1680
         TabIndex        =   2
         Top             =   1140
         Width           =   3375
      End
      Begin VB.TextBox txtServerName 
         Height          =   375
         Left            =   1680
         TabIndex        =   0
         Top             =   360
         Width           =   3375
      End
      Begin VB.Label lblAuthentication 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Authentication"
         Height          =   240
         Left            =   240
         TabIndex        =   40
         Top             =   780
         Width           =   1215
      End
      Begin VB.Label lblPassword 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
         Height          =   240
         Left            =   240
         TabIndex        =   39
         Top             =   1560
         Width           =   825
      End
      Begin VB.Label lblUserName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User Name"
         Height          =   240
         Left            =   240
         TabIndex        =   38
         Top             =   1140
         Width           =   945
      End
      Begin VB.Label lblLblServerName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Server Name"
         Height          =   240
         Left            =   240
         TabIndex        =   37
         Top             =   360
         Width           =   1125
      End
   End
   Begin VB.Label lblTableInfo 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "-"
      ForeColor       =   &H00FF0000&
      Height          =   345
      Left            =   10920
      TabIndex        =   66
      Top             =   8220
      Width           =   3855
   End
   Begin VB.Label lblTableInfoCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Table :"
      Height          =   240
      Left            =   10200
      TabIndex        =   65
      Top             =   8220
      Width           =   615
   End
   Begin VB.Label lblDatabaseInfo 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "-"
      ForeColor       =   &H00FF0000&
      Height          =   345
      Left            =   6120
      TabIndex        =   64
      Top             =   8220
      Width           =   3855
   End
   Begin VB.Label lblDatabaseInfoCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Database :"
      Height          =   240
      Left            =   5040
      TabIndex        =   63
      Top             =   8220
      Width           =   930
   End
   Begin VB.Label lblConnectionInfo 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "-"
      ForeColor       =   &H00FF0000&
      Height          =   345
      Left            =   1080
      TabIndex        =   62
      Top             =   8220
      Width           =   3855
   End
   Begin VB.Label lblConnectionInfoCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Server :"
      Height          =   240
      Left            =   120
      TabIndex        =   61
      Top             =   8220
      Width           =   705
   End
   Begin VB.Label lblGlobalStatus 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ready"
      ForeColor       =   &H00FF0000&
      Height          =   345
      Left            =   1080
      TabIndex        =   56
      Top             =   8595
      Width           =   13695
   End
   Begin VB.Label lblGlobalStatusCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Status  :"
      Height          =   240
      Left            =   120
      TabIndex        =   55
      Top             =   8595
      Width           =   735
   End
End
Attribute VB_Name = "frmImportExcelToSql"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const SETTINGS_FILE_NAME As String = "AppSettings.ini"
Private Const SETTINGS_SECTION_CONNECTION As String = "Connection"

'Private mSqlConnection As clsSqlServerConnection
Private mDatabaseBrowser As clsDatabaseBrowser
Private mMappingManager As clsMappingManager
Private mTargetColumns As Collection
Private mExcelReader As clsExcelReader
Private mExcelHeaders As Collection
Private mCurrentGridRow As Long
Private mCurrentGridCol As Long
Private mImportLogger As clsImportLogger
Private mStagingManager As clsStagingManager
Private mImportEngine As clsImportEngine
'Private mAllDatabases As Collection
'Private mAllTables As Collection
Private mCancelRequested As Boolean

Private Sub cmbGridSqlColumns_Click()
    Dim SelectedSqlColumn As String
    
    If mCurrentGridRow <= 0 Then Exit Sub
    
    SelectedSqlColumn = NzString(cmbGridSqlColumns.Text)
    
    If Not CanAssignSqlColumnToRow(mCurrentGridRow, SelectedSqlColumn) Then
        cmbGridSqlColumns.Visible = False
        grdMapping.Row = mCurrentGridRow
        grdMapping.Col = 1
        grdMapping.RowSel = mCurrentGridRow
        grdMapping.ColSel = 1
        grdMapping.SetFocus
        Exit Sub
    End If
    
    UpdateGridRowBySelectedSqlColumn mCurrentGridRow, SelectedSqlColumn
    
    grdMapping.Row = mCurrentGridRow
    grdMapping.Col = 1
    grdMapping.RowSel = mCurrentGridRow
    grdMapping.ColSel = 1
    
    cmbGridSqlColumns.Visible = False
    grdMapping.SetFocus
End Sub

Private Sub cmbGridSqlColumns_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim SelectedSqlColumn As String
    
    Select Case KeyCode
        Case vbKeyReturn
            If mCurrentGridRow > 0 Then
                SelectedSqlColumn = NzString(cmbGridSqlColumns.Text)
                
                If CanAssignSqlColumnToRow(mCurrentGridRow, SelectedSqlColumn) Then
                    UpdateGridRowBySelectedSqlColumn mCurrentGridRow, SelectedSqlColumn
                End If
                
                grdMapping.Row = mCurrentGridRow
                grdMapping.Col = 1
                grdMapping.RowSel = mCurrentGridRow
                grdMapping.ColSel = 1
                
                cmbGridSqlColumns.Visible = False
                grdMapping.SetFocus
            End If
            
            KeyCode = 0
            
        Case vbKeyEscape
            cmbGridSqlColumns.Visible = False
            grdMapping.SetFocus
            KeyCode = 0
    End Select
End Sub

Private Sub cmbGridSqlColumns_LostFocus()
    If cmbGridSqlColumns.Visible Then
        cmbGridSqlColumns.Visible = False
    End If
End Sub

Private Sub cmdAutoMatch_Click()
    On Error GoTo ErrorHandler
    
    Dim Mappings As Collection
    
    If mExcelHeaders Is Nothing Then Exit Sub
    If mTargetColumns Is Nothing Then Exit Sub
    
    SetGlobalStatus "Running auto match..."
    AppendStatusMessage "Running auto match..."
    DoEvents
    
    Set Mappings = mMappingManager.AutoMatch(mExcelHeaders, mTargetColumns)
    ApplyMappingsToGrid Mappings
    Call ValidateNoDuplicateSqlMappings
    
    AppendStatusMessage "Auto match completed."
    SetGlobalStatus "Auto match completed"
    cmdValidateMapping_Click
    
    '''AppendStatusMessage "Please click Validate Mapping button to enable Start Import button in Execution tab."
    '''MsgBox "Please click Validate Mapping button to enable Start Import button in Execution tab."
    
    Exit Sub

ErrorHandler:
    AppendStatusMessage "Auto match failed: " & Err.Description
    MsgBox "Auto match failed." & vbCrLf & Err.Description, vbExclamation, APP_NAME
End Sub

Private Sub cmdBrowseExcel_Click()
    On Error GoTo ErrorHandler
    
    Set mExcelHeaders = New Collection
    ResetMappingGridRows
    
    grdMapping.Enabled = False
    cmdAutoMatch.Enabled = False
    cmdLoadSavedMapping.Enabled = False
    cmdSaveMapping.Enabled = False
    cmdValidateMapping.Enabled = False
    cmdStartImport.Enabled = False
    
    chkDeleteExisting.Enabled = False
    chkContinueOnTypeError.Enabled = False
    chkContinueOnDuplicate.Enabled = False
    chkAutoLoadSavedMapping.Enabled = False
    
    ResetExecutionProgress
    
    dlgExcelFile.CancelError = True
    dlgExcelFile.DialogTitle = "Select Excel File"
    dlgExcelFile.Filter = "Excel Files (*.xls;*.xlsx)|*.xls;*.xlsx|Excel 97-2003 Workbook (*.xls)|*.xls|Excel Workbook (*.xlsx)|*.xlsx|All Files (*.*)|*.*"
    dlgExcelFile.FileName = ""
    dlgExcelFile.ShowOpen
    
    If Len(Trim$(dlgExcelFile.FileName)) > 0 Then
        txtExcelFile.Text = dlgExcelFile.FileName
        
        Set mExcelHeaders = New Collection
        ResetMappingGridRows
        
        grdMapping.Enabled = False
        cmdAutoMatch.Enabled = False
        cmdLoadSavedMapping.Enabled = False
        cmdSaveMapping.Enabled = False
        cmdValidateMapping.Enabled = False
        cmdStartImport.Enabled = False
        
        chkDeleteExisting.Enabled = False
        chkContinueOnTypeError.Enabled = False
        chkContinueOnDuplicate.Enabled = False
        chkAutoLoadSavedMapping.Enabled = False
        
        ResetExecutionProgress
        ResetPreviewGrid
        
        SetGlobalStatus "Excel file selected"
        AppendStatusMessage "Excel file selected: " & dlgExcelFile.FileName
    End If
    
    Exit Sub

ErrorHandler:
    If Err.Number <> 32755 Then
        AppendStatusMessage "Error selecting Excel file: " & Err.Description
        MsgBox "Error selecting Excel file." & vbCrLf & Err.Number & " - " & Err.Description, vbCritical, APP_NAME
    End If
End Sub

Private Sub cmdCancelImport_Click()
    If MsgBox("Do you want to cancel the current import operation?", vbQuestion + vbYesNo, APP_NAME) = vbYes Then
        mCancelRequested = True
        cmdCancelImport.Enabled = False
        SetGlobalStatus "Cancel requested..."
        AppendStatusMessage "Cancel requested by user."
    End If
End Sub

Private Sub cmdLoadExcelColumns_Click()
    On Error GoTo ErrorHandler
    
    Screen.MousePointer = vbHourglass
    SetGlobalStatus "Loading Excel headers..."
    AppendStatusMessage "Loading Excel headers..."
    DoEvents
    
    If Not ValidateExcelSelectionInputs() Then
        GoTo SafeExit
    End If
    
    mExcelReader.ReadMode = GetSelectedExcelReadMode()
    
    Set mExcelHeaders = mExcelReader.LoadHeaders(Trim$(txtExcelFile.Text))
    
    PopulateMappingGridFromExcelHeaders
    LoadExcelPreview
    SetExcelHeadersLoadedState
    
    If chkAutoLoadSavedMapping.Value = vbChecked Then
        On Error Resume Next
        cmdLoadSavedMapping_Click
        On Error GoTo ErrorHandler
    End If
    
    AppendStatusMessage CStr(mExcelHeaders.Count) & " Excel columns loaded successfully."
    SetGlobalStatus "Excel headers loaded"
    
SafeExit:
    Screen.MousePointer = vbDefault
    Exit Sub

ErrorHandler:
    Screen.MousePointer = vbDefault
    AppendStatusMessage "Failed to load Excel headers: " & Err.Description
    MsgBox "Failed to load Excel headers." & vbCrLf & Err.Number & " - " & Err.Description, vbCritical, APP_NAME
End Sub

Private Sub cmdLoadSavedMapping_Click()
    On Error GoTo ErrorHandler
    
    Dim Conn As ADODB.Connection
    Dim Mappings As Collection
    
    If gAppContext Is Nothing Then
        MsgBox "Application context is not initialized.", vbExclamation, APP_NAME
        Exit Sub
    End If
    
    If Not gAppContext.IsConnected Then
        MsgBox "No active SQL connection is available.", vbExclamation, APP_NAME
        Exit Sub
    End If
    
    If Not gAppContext.HasSelectedTarget Then
        MsgBox "No target table is selected.", vbExclamation, APP_NAME
        Exit Sub
    End If
    
    Set Conn = GetActiveConnection()
    If Conn Is Nothing Then
        MsgBox "No active database connection is available.", vbExclamation, APP_NAME
        Exit Sub
    End If
    
    SetGlobalStatus "Loading saved mapping..."
    AppendStatusMessage "Loading saved mapping..."
    DoEvents
    
    Set Mappings = mMappingManager.LoadMapping( _
        Conn, _
        gAppContext.SelectedDatabase, _
        gAppContext.SelectedSchema, _
        gAppContext.SelectedTable)
    
    If Mappings.Count = 0 Then
        AppendStatusMessage "No saved mapping found."
        MsgBox "No saved mapping found for the selected table.", vbInformation, APP_NAME
        Exit Sub
    End If
    
    ApplyMappingsToGrid Mappings
    
    AppendStatusMessage "Saved mapping loaded successfully."
    SetGlobalStatus "Saved mapping loaded"
    
    Exit Sub

ErrorHandler:
    AppendStatusMessage "Failed to load saved mapping: " & Err.Description
    MsgBox "Failed to load saved mapping." & vbCrLf & Err.Description, vbExclamation, APP_NAME
End Sub

Private Sub cmdSaveMapping_Click()
    On Error GoTo ErrorHandler
    
    Dim Conn As ADODB.Connection
    Dim Mappings As Collection
    
    If gAppContext Is Nothing Then
        MsgBox "Application context is not initialized.", vbExclamation, APP_NAME
        Exit Sub
    End If
    
    If Not gAppContext.IsConnected Then
        MsgBox "No active SQL connection is available.", vbExclamation, APP_NAME
        Exit Sub
    End If
    
    If Not gAppContext.HasSelectedTarget Then
        MsgBox "No target table is selected.", vbExclamation, APP_NAME
        Exit Sub
    End If
    
    Set Conn = GetActiveConnection()
    If Conn Is Nothing Then
        MsgBox "No active database connection is available.", vbExclamation, APP_NAME
        Exit Sub
    End If
    
    Set Mappings = CollectMappingsFromGrid()
    
    SetGlobalStatus "Saving mapping..."
    AppendStatusMessage "Saving mapping..."
    DoEvents
    
    mMappingManager.EnsureMappingTableExists Conn, gAppContext.SelectedDatabase
    
    mMappingManager.SaveMapping _
        Conn, _
        gAppContext.SelectedDatabase, _
        gAppContext.SelectedSchema, _
        gAppContext.SelectedTable, _
        Mappings
    
    AppendStatusMessage "Mapping saved successfully."
    SetGlobalStatus "Mapping saved"
    
    MsgBox "Mapping saved successfully.", vbInformation, APP_NAME
    Exit Sub

ErrorHandler:
    AppendStatusMessage "Failed to save mapping: " & Err.Description
    MsgBox "Failed to save mapping." & vbCrLf & Err.Description, vbExclamation, APP_NAME
End Sub

Private Sub cmdStartImport_Click()
    On Error GoTo ErrorHandler
    
    Dim Conn As ADODB.Connection
    Dim Options As clsImportOptions
    Dim Mappings As Collection
    Dim MissingRequired As String
    Dim ExcelRows As Collection
    Dim InsertedToStageCount As Long
    Dim StageTableName As String
    Dim FinalTotalRows As Long
    Dim FriendlyMessage As String
    Dim OriginalErrNumber As Long
    Dim OriginalErrDescription As String
    
    If gAppContext Is Nothing Then
        MsgBox "Application context is not initialized.", vbExclamation, APP_NAME
        Exit Sub
    End If
    
    If Not gAppContext.IsConnected Then
        MsgBox "No active SQL connection is available.", vbExclamation, APP_NAME
        Exit Sub
    End If
    
    If Not gAppContext.HasSelectedTarget Then
        MsgBox "No target table is selected.", vbExclamation, APP_NAME
        Exit Sub
    End If
    
    Set Conn = GetActiveConnection()
    If Conn Is Nothing Then
        MsgBox "No active database connection is available.", vbExclamation, APP_NAME
        Exit Sub
    End If
    
    If Not ValidateExcelSelectionInputs() Then Exit Sub
    
    Set Mappings = CollectMappingsFromGrid()
    MissingRequired = mMappingManager.ValidateRequiredMappings(mTargetColumns, Mappings)
    
    If Len(MissingRequired) > 0 Then
        MsgBox "Required target columns are not mapped:" & vbCrLf & vbCrLf & MissingRequired, vbExclamation, APP_NAME
        Exit Sub
    End If
    
    If chkDeleteExisting.Value = vbChecked Then
        If MsgBox(MSG_CONFIRM_DELETE_EXISTING, vbQuestion + vbYesNo, APP_NAME) <> vbYes Then
            AppendStatusMessage MSG_OPERATION_CANCELLED
            Exit Sub
        End If
    End If
    
    Set Options = BuildImportOptionsFromForm()
    
    mCancelRequested = False
    SetImportUiBusyState True
    ResetExecutionProgress
    
    ' --------------------------------------------------
    ' Phase 1: Read Excel + Create Stage + Insert to Stage
    ' --------------------------------------------------
    Set mImportLogger = New clsImportLogger
    mImportLogger.StartLog _
        gAppContext.ServerName, _
        gAppContext.SelectedDatabase, _
        gAppContext.SelectedSchema, _
        gAppContext.SelectedTable, _
        Trim$(txtExcelFile.Text), _
        GetExcelReadModeText(Options.ExcelReadMode)
    
    Screen.MousePointer = vbHourglass
    
    lblCurrentStep.Caption = "Reading Excel rows..."
    SetGlobalStatus "Reading Excel rows..."
    AppendStatusMessage "Reading Excel rows from file..."
    DoEvents
    
    CheckForCancelRequest
    
    mExcelReader.ReadMode = Options.ExcelReadMode
    Set ExcelRows = mExcelReader.LoadRows(Trim$(txtExcelFile.Text), Me)
    
    AppendStatusMessage CStr(ExcelRows.Count) & " data rows loaded from Excel."
    mImportLogger.WriteInfo CStr(ExcelRows.Count) & " data rows loaded from Excel."
    
    lblCurrentStep.Caption = "Creating staging table..."
    SetGlobalStatus "Creating staging table..."
    DoEvents
    
    CheckForCancelRequest
    
    StageTableName = mStagingManager.CreateStageTable( _
        Conn, _
        gAppContext.SelectedDatabase, _
        gAppContext.SelectedSchema, _
        gAppContext.SelectedTable, _
        Mappings, _
        mTargetColumns)
    
    AppendStatusMessage "Staging table created: " & StageTableName
    mImportLogger.WriteInfo "Staging table created: " & StageTableName
    
    lblCurrentStep.Caption = "Writing rows to staging..."
    SetGlobalStatus "Writing rows to staging..."
    ResetExecutionCountersOnly
    DoEvents
    
    CheckForCancelRequest
    
    InsertedToStageCount = mStagingManager.InsertRowsIntoStage( _
        Conn, _
        gAppContext.SelectedDatabase, _
        gAppContext.SelectedSchema, _
        ExcelRows, _
        Mappings, _
        mTargetColumns, _
        Options, _
        mImportLogger, _
        Me)
    
    AppendStatusMessage CStr(InsertedToStageCount) & " rows inserted into staging."
    mImportLogger.WriteInfo CStr(InsertedToStageCount) & " rows inserted into staging."
    
    ' --------------------------------------------------
    ' Phase 2: Import from Stage to Target
    ' IMPORTANT: use a NEW logger so counts belong only to target phase
    ' --------------------------------------------------
    lblCurrentStep.Caption = "Importing rows to target..."
    SetGlobalStatus "Importing rows to target..."
    ResetExecutionCountersOnly
    
    Set mImportLogger = New clsImportLogger
    mImportLogger.StartLog _
        gAppContext.ServerName, _
        gAppContext.SelectedDatabase, _
        gAppContext.SelectedSchema, _
        gAppContext.SelectedTable, _
        Trim$(txtExcelFile.Text), _
        GetExcelReadModeText(Options.ExcelReadMode)
    
    mImportLogger.WriteInfo "Target import phase started."
    mImportLogger.WriteInfo "Stage table: " & StageTableName
    
    DoEvents
    
    CheckForCancelRequest
    
    mImportEngine.ExecuteImport _
        Conn, _
        gAppContext.SelectedDatabase, _
        gAppContext.SelectedSchema, _
        gAppContext.SelectedTable, _
        StageTableName, _
        Mappings, _
        mTargetColumns, _
        Options, _
        mImportLogger, _
        Me
    
    lblCurrentStep.Caption = "Cleaning up..."
    SetGlobalStatus "Cleaning up..."
    DoEvents
    
    On Error Resume Next
    mStagingManager.DropStageTable Conn, gAppContext.SelectedDatabase, gAppContext.SelectedSchema
    If Err.Number = 0 Then
        AppendStatusMessage "Staging table dropped."
    Else
        AppendStatusMessage "Warning: staging table could not be dropped. Error " & CStr(Err.Number) & " - " & Err.Description
        If Not mImportLogger Is Nothing Then
            mImportLogger.WriteWarning "Staging table could not be dropped. Error " & CStr(Err.Number) & " - " & Err.Description
        End If
        Err.Clear
    End If
    On Error GoTo ErrorHandler
    
    AppendStatusMessage "Import completed successfully."
    AppendStatusMessage "Log file: " & mImportLogger.LogFilePath
    
    mImportLogger.FinishLog "SUCCESS"
    
    lblCurrentStep.Caption = "Import completed"
    SetGlobalStatus "Import completed"
    
    lblSuccessCount.Caption = CStr(mImportLogger.SuccessCount)
    lblSkipCount.Caption = CStr(mImportLogger.SkipCount)
    lblErrorCount.Caption = CStr(mImportLogger.ErrorCount)
    
    FinalTotalRows = mImportLogger.SuccessCount + _
                     mImportLogger.SkipCount + _
                     mImportLogger.ErrorCount
    
    SetExecutionSummary _
        mImportLogger.SuccessCount, _
        mImportLogger.SkipCount, _
        mImportLogger.ErrorCount, _
        FinalTotalRows
    
    MsgBox "Import completed successfully." & vbCrLf & vbCrLf & _
           "Target table: " & gAppContext.GetFullTargetName() & vbCrLf & _
           "Imported rows: " & CStr(mImportLogger.SuccessCount) & vbCrLf & _
           "Skipped rows: " & CStr(mImportLogger.SkipCount) & vbCrLf & _
           "Error rows: " & CStr(mImportLogger.ErrorCount) & vbCrLf & vbCrLf & _
           "Log file:" & vbCrLf & mImportLogger.LogFilePath, vbInformation, APP_NAME
    
SafeExit:
    Screen.MousePointer = vbDefault
    mCancelRequested = False
    SetImportUiBusyState False
    cmdStartImport.Enabled = True
    Exit Sub

ErrorHandler:
    OriginalErrNumber = Err.Number
    OriginalErrDescription = Err.Description
    
    Screen.MousePointer = vbDefault
    
    If OriginalErrNumber = vbObjectError + 3001 _
       Or OriginalErrNumber = vbObjectError + 3002 _
       Or OriginalErrNumber = vbObjectError + 3003 _
       Or OriginalErrNumber = vbObjectError + 3004 Then
        
        FriendlyMessage = "Import was cancelled by user."
        SetCancelledExecutionState
    Else
        FriendlyMessage = GetFriendlyImportErrorMessage(OriginalErrNumber, OriginalErrDescription)
    End If
    
    On Error Resume Next
    
    If Not mStagingManager Is Nothing Then
        mStagingManager.DropStageTable Conn, gAppContext.SelectedDatabase, gAppContext.SelectedSchema
    End If
    
    If Not mImportLogger Is Nothing Then
        mImportLogger.WriteError "Error " & CStr(OriginalErrNumber) & ": " & OriginalErrDescription
        
        If OriginalErrNumber = vbObjectError + 3001 _
           Or OriginalErrNumber = vbObjectError + 3002 _
           Or OriginalErrNumber = vbObjectError + 3003 _
           Or OriginalErrNumber = vbObjectError + 3004 Then
            mImportLogger.FinishLog "CANCELLED"
        Else
            mImportLogger.FinishLog "FAILED"
        End If
    End If
    
    AppendStatusMessage "Import failed: " & FriendlyMessage
    SetGlobalStatus "Import failed"
    
    SetImportUiBusyState False
    cmdStartImport.Enabled = True
    
    MsgBox FriendlyMessage & vbCrLf & vbCrLf & _
           "More details have been written to the log file.", _
           vbExclamation, APP_NAME
End Sub

Private Sub cmdValidateMapping_Click()
    On Error GoTo ErrorHandler
    
    Dim Mappings As Collection
    Dim MissingRequired As String
    
    If gAppContext Is Nothing Then
        MsgBox "Application context is not initialized.", vbExclamation, APP_NAME
        Exit Sub
    End If
    
    If Not gAppContext.IsConnected Then
        MsgBox "No active SQL connection is available.", vbExclamation, APP_NAME
        Exit Sub
    End If
    
    If Not gAppContext.HasSelectedTarget Then
        MsgBox "No target table is selected.", vbExclamation, APP_NAME
        Exit Sub
    End If
    
    Set Mappings = CollectMappingsFromGrid()
    
    If Not ValidateNoDuplicateSqlMappings() Then
        cmdStartImport.Enabled = False
        Exit Sub
    End If
    
    MissingRequired = mMappingManager.ValidateRequiredMappings(mTargetColumns, Mappings)
    
    If Len(MissingRequired) > 0 Then
        AppendStatusMessage "Mapping validation failed. Required fields are missing."
        SetGlobalStatus "Mapping validation failed"
        cmdStartImport.Enabled = False
        
        MsgBox "Required target columns are not mapped:" & vbCrLf & vbCrLf & _
               MissingRequired, vbExclamation, APP_NAME
        Exit Sub
    End If
    
    cmdStartImport.Enabled = True
    AppendStatusMessage "Mapping validation successful."
    SetGlobalStatus "Mapping validation successful"
    
    MsgBox "Mapping is valid. You can now start the import.", vbInformation, APP_NAME
    Exit Sub

ErrorHandler:
    AppendStatusMessage "Failed to validate mapping: " & Err.Description
    MsgBox "Failed to validate mapping." & vbCrLf & Err.Description, vbExclamation, APP_NAME
End Sub

Private Sub Form_Load()
    InitializeForm
End Sub

Private Sub InitializeForm()
    On Error GoTo ErrorHandler
    
    EnsureImportObjects
    
    If gAppContext Is Nothing Then
        MsgBox "Application context is not initialized.", vbExclamation, APP_NAME
        Unload Me
        Exit Sub
    End If
    
    If Not gAppContext.IsConnected Then
        MsgBox "No active SQL connection is available.", vbExclamation, APP_NAME
        Unload Me
        Exit Sub
    End If
    
    If Not gAppContext.HasSelectedTarget Then
        MsgBox "No target table is selected.", vbExclamation, APP_NAME
        Unload Me
        Exit Sub
    End If
    
    If GetActiveConnection() Is Nothing Then
        MsgBox "No active database connection is available.", vbExclamation, APP_NAME
        Unload Me
        Exit Sub
    End If
    
    ApplyAppContextToImportForm
    LoadTargetTableStructureFromContext
    
    Exit Sub

ErrorHandler:
    MsgBox "Error initializing import form." & vbCrLf & _
           Err.Number & " - " & Err.Description, vbCritical, APP_NAME
    Unload Me
End Sub

'Private Sub InitializeAuthenticationCombo()
'    cmbAuthentication.Clear
'    cmbAuthentication.AddItem AUTH_WINDOWS
'    cmbAuthentication.AddItem AUTH_SQL_SERVER
'    cmbAuthentication.ListIndex = 0
'
'    UpdateAuthenticationControls
'End Sub

Private Sub InitializeExcelReadOptions()
    optReadByAdo.Value = False
    optReadByExcel.Value = True
End Sub

Private Sub InitializeMappingGrid()
    With grdMapping
        .Rows = 2
        .Cols = 5
        .FixedRows = 1
        .FixedCols = 0
        .SelectionMode = 0
        
        .TextMatrix(0, 0) = "Excel Column"
        .TextMatrix(0, 1) = "SQL Column"
        .TextMatrix(0, 2) = "SQL Data Type"
        .TextMatrix(0, 3) = "Required"
        .TextMatrix(0, 4) = "Status"
        
        .TextMatrix(1, 0) = ""
        .TextMatrix(1, 1) = ""
        .TextMatrix(1, 2) = ""
        .TextMatrix(1, 3) = ""
        .TextMatrix(1, 4) = ""
    
        .ColWidth(0) = 2500
        .ColWidth(1) = 2500
        .ColWidth(2) = 1800
        .ColWidth(3) = 1200
        .ColWidth(4) = 1800
    End With
End Sub

Private Sub InitializeCounters()
    prgImport.Min = 0
    prgImport.Max = PROGRESS_MAX
    prgImport.Value = 0
    
    lblCurrentStep.Caption = "Ready"
    lblGlobalStatus.Caption = "Ready"
    lblRowProgress.Caption = "0 / 0"
    lblSuccessCount.Caption = "0"
    lblSkipCount.Caption = "0"
    lblErrorCount.Caption = "0"
    
    txtStatus.Text = ""
    
    txtStatus.Text = ""
End Sub

Private Sub SetInitialControlState()
    'FraConnection.Enabled = True
    'fraDatabaseTable.Enabled = False
    fraExcel.Enabled = False
    
    grdMapping.Enabled = False
    cmdAutoMatch.Enabled = False
    cmdLoadSavedMapping.Enabled = False
    cmdSaveMapping.Enabled = False
    cmdValidateMapping.Enabled = False
    
    chkDeleteExisting.Enabled = False
    chkContinueOnTypeError.Enabled = False
    chkContinueOnDuplicate.Enabled = False
    chkAutoLoadSavedMapping.Enabled = False
    chkRememberConnectionSettings.Enabled = True
    
    cmdStartImport.Enabled = False
    cmdExit.Enabled = True
    cmdCancelImport.Enabled = False

    txtUserName.Enabled = False
    txtPassword.Enabled = False
End Sub

'Private Sub cmbAuthentication_Click()
'    On Error GoTo ErrorHandler
'
'    UpdateAuthenticationControls
'    ResetAfterConnectionSettingsChange
'
'    AppendStatusMessage "Authentication mode changed. Database and table selections were cleared."
'    Exit Sub
'
'ErrorHandler:
'    MsgBox "Error changing authentication mode." & vbCrLf & _
'           Err.Number & " - " & Err.Description, vbExclamation, APP_NAME
'End Sub
'
'Private Sub UpdateAuthenticationControls()
'    Dim IsSqlAuth As Boolean
'
'    IsSqlAuth = (cmbAuthentication.Text = AUTH_SQL_SERVER)
'
'    txtUserName.Enabled = IsSqlAuth
'    txtPassword.Enabled = IsSqlAuth
'
'    If Not IsSqlAuth Then
'        txtUserName.Text = ""
'        txtPassword.Text = ""
'    End If
'End Sub
'
Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub SetConnectedState()
    'fraDatabaseTable.Enabled = True
    fraExcel.Enabled = False
    
    grdMapping.Enabled = False
    cmdAutoMatch.Enabled = False
    cmdLoadSavedMapping.Enabled = False
    cmdSaveMapping.Enabled = False
    cmdValidateMapping.Enabled = False
    
    chkDeleteExisting.Enabled = False
    chkContinueOnTypeError.Enabled = False
    chkContinueOnDuplicate.Enabled = False
    chkAutoLoadSavedMapping.Enabled = False
    
    cmdStartImport.Enabled = False
    
    lblCurrentStep.Caption = "Connected to SQL Server"
    SetGlobalStatus "Connected to SQL Server"
    AppendStatusMessage "Connected successfully."
End Sub

Private Sub SetDisconnectedState()
    'fraDatabaseTable.Enabled = False
    fraExcel.Enabled = False
    
    'cmbDatabase.Clear
    'cmbTable.Clear
    
    grdMapping.Enabled = False
    cmdAutoMatch.Enabled = False
    cmdLoadSavedMapping.Enabled = False
    cmdSaveMapping.Enabled = False
    cmdValidateMapping.Enabled = False
    
    chkDeleteExisting.Enabled = False
    chkContinueOnTypeError.Enabled = False
    chkContinueOnDuplicate.Enabled = False
    chkAutoLoadSavedMapping.Enabled = False
    
    cmdStartImport.Enabled = False
    
    lblCurrentStep.Caption = "Disconnected"
    SetGlobalStatus "Disconnected"
End Sub

Private Sub AppendStatusMessage(ByVal MessageText As String)
    Dim FinalText As String
    
    FinalText = Format$(Now, "hh:nn:ss") & " - " & MessageText
    
    If Len(txtStatus.Text) = 0 Then
        txtStatus.Text = FinalText
    Else
        txtStatus.Text = txtStatus.Text & vbCrLf & FinalText
    End If
    
    txtStatus.SelStart = Len(txtStatus.Text)
    SetGlobalStatus MessageText
End Sub

'Private Function ValidateConnectionInputs() As Boolean
'    ValidateConnectionInputs = False
'
'    If IsNullOrEmpty(txtServerName.Text) Then
'        MsgBox "Please enter server name.", vbExclamation, APP_NAME
'        txtServerName.SetFocus
'        Exit Function
'    End If
'
'    If cmbAuthentication.ListIndex < 0 Then
'        MsgBox "Please select authentication type.", vbExclamation, APP_NAME
'        cmbAuthentication.SetFocus
'        Exit Function
'    End If
'
'    If cmbAuthentication.Text = AUTH_SQL_SERVER Then
'        If IsNullOrEmpty(txtUserName.Text) Then
'            MsgBox "Please enter user name.", vbExclamation, APP_NAME
'            txtUserName.SetFocus
'            Exit Function
'        End If
'
'        If IsNullOrEmpty(txtPassword.Text) Then
'            MsgBox "Please enter password.", vbExclamation, APP_NAME
'            txtPassword.SetFocus
'            Exit Function
'        End If
'    End If
'
'    ValidateConnectionInputs = True
'End Function

'Private Sub cmdConnect_Click()
'    On Error GoTo ErrorHandler
'
'    Screen.MousePointer = vbHourglass
'    SetGlobalStatus "Connecting to SQL Server..."
'    AppendStatusMessage "Connecting to SQL Server..."
'    DoEvents
'
'    If Not ValidateConnectionInputs() Then GoTo SafeExit
'
'    If mSqlConnection.OpenConnection( _
'        Trim$(txtServerName.Text), _
'        cmbAuthentication.Text, _
'        Trim$(txtUserName.Text), _
'        txtPassword.Text) Then
'
'        SetConnectedState
'        LoadDatabases
'
'        If chkRememberConnectionSettings.Value = vbChecked Then
'            SaveConnectionSettings
'        Else
'            ClearSavedConnectionSettings
'        End If
'
'        MsgBox "Connection established successfully.", vbInformation, APP_NAME
'    End If
'
'SafeExit:
'    Screen.MousePointer = vbDefault
'    Exit Sub
'
'ErrorHandler:
'    Screen.MousePointer = vbDefault
'
'    ResetAfterConnectionSettingsChange
'
'    AppendStatusMessage "Connection failed: " & Err.Description
'    MsgBox "Connection failed." & vbCrLf & _
'           Err.Number & " - " & Err.Description, vbCritical, APP_NAME
'End Sub

'Private Sub Form_Unload(Cancel As Integer)
'    On Error Resume Next
'
'    If Not mSqlConnection Is Nothing Then
'        mSqlConnection.CloseConnection
'    End If
'
'    Set mSqlConnection = Nothing
'End Sub

Private Sub SetGlobalStatus(ByVal StatusText As String)
    lblGlobalStatus.Caption = StatusText
    DoEvents
End Sub

'Private Sub LoadDatabases()
'    On Error GoTo ErrorHandler
'
'    Dim Rs As ADODB.Recordset
'    Dim DbName As String
'
'    cmbDatabase.Clear
'    cmbTable.Clear
'    txtSearchDatabase.Text = ""
'    txtSearchTable.Text = ""
'
'    Set mAllDatabases = New Collection
'    Set mAllTables = New Collection
'
'    SetGlobalStatus "Loading databases..."
'    AppendStatusMessage "Loading user databases..."
'    DoEvents
'
'    Set Rs = mDatabaseBrowser.GetUserDatabases(mSqlConnection.Connection)
'
'    Do While Not Rs.EOF
'        DbName = NzString(Rs.Fields("name").Value)
'        mAllDatabases.Add DbName
'        cmbDatabase.AddItem DbName
'        Rs.MoveNext
'    Loop
'
'    Rs.Close
'    Set Rs = Nothing
'
'    If cmbDatabase.ListCount > 0 Then
'        cmbDatabase.ListIndex = 0
'    End If
'
'    SetGlobalStatus "Databases loaded"
'    AppendStatusMessage "Database list loaded successfully."
'
'    Exit Sub
'
'ErrorHandler:
'    AppendStatusMessage "Failed to load databases: " & Err.Description
'    MsgBox "Failed to load databases." & vbCrLf & Err.Number & " - " & Err.Description, vbCritical, APP_NAME
'End Sub

'Private Sub LoadTables()
'    On Error GoTo ErrorHandler
'
'    Dim Rs As ADODB.Recordset
'    Dim DisplayText As String
'
'    cmbTable.Clear
'    txtSearchTable.Text = ""
'    Set mAllTables = New Collection
'
'    If cmbDatabase.ListIndex < 0 Then
'        MsgBox MSG_NO_DATABASE_SELECTED, vbExclamation, APP_NAME
'        cmbDatabase.SetFocus
'        Exit Sub
'    End If
'
'    SetGlobalStatus "Loading tables..."
'    AppendStatusMessage "Loading base tables for database: " & cmbDatabase.Text
'    DoEvents
'
'    Set Rs = mDatabaseBrowser.GetBaseTables(mSqlConnection.Connection, cmbDatabase.Text)
'
'    Do While Not Rs.EOF
'        DisplayText = NzString(Rs.Fields("TABLE_SCHEMA").Value) & "." & NzString(Rs.Fields("TABLE_NAME").Value)
'        mAllTables.Add DisplayText
'        cmbTable.AddItem DisplayText
'        Rs.MoveNext
'    Loop
'
'    Rs.Close
'    Set Rs = Nothing
'
'    If cmbTable.ListCount > 0 Then
'        cmbTable.ListIndex = 0
'    End If
'
'    SetGlobalStatus "Tables loaded"
'    AppendStatusMessage "Table list loaded successfully."
'
'    Exit Sub
'
'ErrorHandler:
'    AppendStatusMessage "Failed to load tables: " & Err.Description
'    MsgBox "Failed to load tables." & vbCrLf & Err.Number & " - " & Err.Description, vbCritical, APP_NAME
'End Sub

'Private Sub cmdLoadTables_Click()
'    LoadTables
'End Sub

'Private Sub cmbDatabase_Click()
'    cmbTable.Clear
'    fraExcel.Enabled = False
'
'    ResetAfterTableSelectionChange
'
'    SetGlobalStatus "Database selected: " & cmbDatabase.Text
'End Sub

'Private Sub cmbTable_Click()
'    On Error GoTo ErrorHandler
'
'    If cmbTable.ListIndex >= 0 Then
'        ResetAfterTableSelectionChange
'        fraExcel.Enabled = True
'
'        LoadTargetTableStructure
'        mMappingManager.EnsureMappingTableExists mSqlConnection.Connection, cmbDatabase.Text
'
'        SetGlobalStatus "Table selected: " & cmbTable.Text
'        AppendStatusMessage "Target table selected: " & cmbTable.Text
'        AppendStatusMessage "Mapping table is ready."
'    End If
'
'    Exit Sub
'
'ErrorHandler:
'    AppendStatusMessage "Error preparing selected table: " & Err.Description
'    MsgBox "Error preparing selected table." & vbCrLf & Err.Number & " - " & Err.Description, vbCritical, APP_NAME
'End Sub

'Private Function GetSelectedSchemaName() As String
'    Dim FullName As String
'    Dim DotPos As Long
'
'    FullName = Trim$(cmbTable.Text)
'    DotPos = InStr(1, FullName, ".")
'
'    If DotPos > 0 Then
'        GetSelectedSchemaName = Left$(FullName, DotPos - 1)
'    Else
'        GetSelectedSchemaName = "dbo"
'    End If
'End Function

'Private Function GetSelectedTableName() As String
'    Dim FullName As String
'    Dim DotPos As Long
'
'    FullName = Trim$(cmbTable.Text)
'    DotPos = InStr(1, FullName, ".")
'
'    If DotPos > 0 Then
'        GetSelectedTableName = Mid$(FullName, DotPos + 1)
'    Else
'        GetSelectedTableName = FullName
'    End If
'End Function

'Private Sub LoadTargetTableStructure()
'    On Error GoTo ErrorHandler
'
'    SetGlobalStatus "Loading target table structure..."
'    AppendStatusMessage "Loading table structure for: " & cmbTable.Text
'    DoEvents
'
'    Set mTargetColumns = mDatabaseBrowser.GetTableColumns( _
'        mSqlConnection.Connection, _
'        cmbDatabase.Text, _
'        GetSelectedSchemaName(), _
'        GetSelectedTableName())
'
'    SetGlobalStatus "Target table structure loaded"
'    AppendStatusMessage "Loaded " & CStr(mTargetColumns.Count) & " target columns."
'
'    Exit Sub
'
'ErrorHandler:
'    AppendStatusMessage "Failed to load target table structure: " & Err.Description
'    MsgBox "Failed to load target table structure." & vbCrLf & Err.Number & " - " & Err.Description, vbCritical, APP_NAME
'End Sub

Private Function GetSelectedExcelReadMode() As Integer
    If optReadByExcel.Value = True Then
        GetSelectedExcelReadMode = EXCEL_READ_MODE_AUTOMATION
    Else
        GetSelectedExcelReadMode = EXCEL_READ_MODE_ADO
    End If
End Function

Private Function ValidateExcelSelectionInputs() As Boolean
    ValidateExcelSelectionInputs = False
    
    If gAppContext Is Nothing Then
        MsgBox "Application context is not initialized.", vbExclamation, APP_NAME
        Exit Function
    End If
    
    If Not gAppContext.IsConnected Then
        MsgBox "No active SQL connection is available.", vbExclamation, APP_NAME
        Exit Function
    End If
    
    If Len(Trim$(gAppContext.SelectedDatabase)) = 0 Then
        MsgBox MSG_NO_DATABASE_SELECTED, vbExclamation, APP_NAME
        Exit Function
    End If
    
    If Not gAppContext.HasSelectedTarget Then
        MsgBox MSG_NO_TABLE_SELECTED, vbExclamation, APP_NAME
        Exit Function
    End If
    
    If IsNullOrEmpty(txtExcelFile.Text) Then
        MsgBox MSG_NO_EXCEL_FILE_SELECTED, vbExclamation, APP_NAME
        txtExcelFile.SetFocus
        Exit Function
    End If
    
    If Not FileExists(txtExcelFile.Text) Then
        MsgBox "Selected Excel file does not exist.", vbExclamation, APP_NAME
        txtExcelFile.SetFocus
        Exit Function
    End If
    
    If Not IsExcelFileExtensionValid(txtExcelFile.Text) Then
        MsgBox "Please select a valid Excel file (.xls or .xlsx).", vbExclamation, APP_NAME
        txtExcelFile.SetFocus
        Exit Function
    End If
    
    ValidateExcelSelectionInputs = True
End Function

Private Sub ResetMappingGridRows()
    With grdMapping
        .Rows = 2
        .Cols = 5
        .FixedRows = 1
        .FixedCols = 0
        
        .TextMatrix(0, 0) = "Excel Column"
        .TextMatrix(0, 1) = "SQL Column"
        .TextMatrix(0, 2) = "SQL Data Type"
        .TextMatrix(0, 3) = "Required"
        .TextMatrix(0, 4) = "Status"
        
        .TextMatrix(1, 0) = ""
        .TextMatrix(1, 1) = ""
        .TextMatrix(1, 2) = ""
        .TextMatrix(1, 3) = ""
        .TextMatrix(1, 4) = ""
    End With
End Sub

Private Sub PopulateMappingGridFromExcelHeaders()
    Dim i As Long
    Dim HeaderName As String
    
    ResetMappingGridRows
    
    If mExcelHeaders Is Nothing Then Exit Sub
    If mExcelHeaders.Count = 0 Then Exit Sub
    
    With grdMapping
        .Rows = mExcelHeaders.Count + 1
        
        For i = 1 To mExcelHeaders.Count
            HeaderName = NzString(mExcelHeaders(i))
            
            .TextMatrix(i, 0) = HeaderName
            .TextMatrix(i, 1) = ""
            .TextMatrix(i, 2) = ""
            .TextMatrix(i, 3) = ""
            .TextMatrix(i, 4) = "Unmapped"
        Next i
    End With
End Sub

Private Sub SetExcelHeadersLoadedState()
    grdMapping.Enabled = True
    
    cmdAutoMatch.Enabled = True
    cmdLoadSavedMapping.Enabled = True
    cmdSaveMapping.Enabled = True
    cmdValidateMapping.Enabled = True
    
    chkDeleteExisting.Enabled = True
    chkContinueOnTypeError.Enabled = True
    chkContinueOnDuplicate.Enabled = True
    chkAutoLoadSavedMapping.Enabled = True
    
    cmdStartImport.Enabled = False
    
    FillGridSqlColumnsCombo
    cmbGridSqlColumns.Visible = False
    
    sstMain.Tab = 0
End Sub

Private Sub FillGridSqlColumnsCombo()
    Dim i As Long
    Dim ColInfo As clsColumnInfo
    
    cmbGridSqlColumns.Clear
    cmbGridSqlColumns.AddItem ""
    
    For i = 1 To mTargetColumns.Count
        Set ColInfo = mTargetColumns(i)
        cmbGridSqlColumns.AddItem ColInfo.ColumnName
    Next i
End Sub

Private Function FindTargetColumn(ByVal ColumnName As String) As clsColumnInfo
    Dim i As Long
    Dim ColInfo As clsColumnInfo
    
    For i = 1 To mTargetColumns.Count
        Set ColInfo = mTargetColumns(i)
        
        If StrComp(Trim$(ColInfo.ColumnName), Trim$(ColumnName), vbTextCompare) = 0 Then
            Set FindTargetColumn = ColInfo
            Exit Function
        End If
    Next i
    
    Set FindTargetColumn = Nothing
End Function

Private Sub UpdateGridRowBySelectedSqlColumn(ByVal RowIndex As Long, ByVal SqlColumnName As String)
    Dim ColInfo As clsColumnInfo
    
    If RowIndex < 1 Then Exit Sub
    
    If Len(Trim$(SqlColumnName)) = 0 Then
        grdMapping.TextMatrix(RowIndex, 1) = ""
        grdMapping.TextMatrix(RowIndex, 2) = ""
        grdMapping.TextMatrix(RowIndex, 3) = ""
        grdMapping.TextMatrix(RowIndex, 4) = "Unmapped"
        Exit Sub
    End If
    
    Set ColInfo = FindTargetColumn(SqlColumnName)
    
    grdMapping.TextMatrix(RowIndex, 1) = SqlColumnName
    
    If ColInfo Is Nothing Then
        grdMapping.TextMatrix(RowIndex, 2) = ""
        grdMapping.TextMatrix(RowIndex, 3) = ""
        grdMapping.TextMatrix(RowIndex, 4) = "Mapped"
    Else
        grdMapping.TextMatrix(RowIndex, 2) = ColInfo.DataType
        grdMapping.TextMatrix(RowIndex, 3) = BoolToText(ColInfo.IsRequired)
        grdMapping.TextMatrix(RowIndex, 4) = "Mapped"
    End If
End Sub

Private Function CollectMappingsFromGrid() As Collection
    Dim Result As Collection
    Dim Item As clsMappingItem
    Dim i As Long
    Dim SqlColName As String
    Dim ColInfo As clsColumnInfo
    
    Set Result = New Collection
    
    For i = 1 To grdMapping.Rows - 1
        Set Item = New clsMappingItem
        
        Item.ExcelColumnName = NzString(grdMapping.TextMatrix(i, 0))
        Item.SqlColumnName = NzString(grdMapping.TextMatrix(i, 1))
        Item.SqlDataType = NzString(grdMapping.TextMatrix(i, 2))
        Item.IsMapped = (Len(Item.SqlColumnName) > 0)
        
        If Item.IsMapped Then
            Set ColInfo = FindTargetColumn(Item.SqlColumnName)
            If Not ColInfo Is Nothing Then
                Item.IsRequired = ColInfo.IsRequired
                If Len(Item.SqlDataType) = 0 Then
                    Item.SqlDataType = ColInfo.DataType
                End If
            End If
        Else
            Item.IsRequired = False
        End If
        
        Result.Add Item
    Next i
    
    Set CollectMappingsFromGrid = Result
End Function

Private Sub ApplyMappingsToGrid(ByVal Mappings As Collection)
    Dim i As Long
    Dim j As Long
    Dim MapItem As clsMappingItem
    Dim ExcelName As String
    
    For i = 1 To grdMapping.Rows - 1
        ExcelName = NzString(grdMapping.TextMatrix(i, 0))
        
        grdMapping.TextMatrix(i, 1) = ""
        grdMapping.TextMatrix(i, 2) = ""
        grdMapping.TextMatrix(i, 3) = ""
        grdMapping.TextMatrix(i, 4) = "Unmapped"
        
        For j = 1 To Mappings.Count
            Set MapItem = Mappings(j)
            
            If StrComp(Trim$(MapItem.ExcelColumnName), Trim$(ExcelName), vbTextCompare) = 0 Then
                If MapItem.IsMapped Then
                    UpdateGridRowBySelectedSqlColumn i, MapItem.SqlColumnName
                End If
                Exit For
            End If
        Next j
    Next i
End Sub

'Private Sub grdMapping_Click()
'    ShowGridSqlColumnsComboIfNeeded
'End Sub

'Private Sub grdMapping_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    ShowGridSqlColumnsComboIfNeeded
'End Sub

Private Sub grdMapping_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ShowGridSqlColumnsComboIfNeeded
End Sub

Private Sub ShowGridSqlColumnsComboIfNeeded()
    On Error GoTo ErrorHandler
    
    Dim TargetRow As Long
    Dim TargetCol As Long
    Dim CurrentSqlColumn As String
    Dim SuggestedSqlColumn As String
    Dim ComboIndex As Integer
    
    cmbGridSqlColumns.Visible = False
    
    TargetRow = grdMapping.MouseRow
    TargetCol = grdMapping.MouseCol
    
    If TargetRow < 1 Then Exit Sub
    If TargetCol <> 1 Then Exit Sub
    
    grdMapping.Row = TargetRow
    grdMapping.Col = TargetCol
    grdMapping.RowSel = TargetRow
    grdMapping.ColSel = TargetCol
    
    mCurrentGridRow = TargetRow
    mCurrentGridCol = TargetCol
    
    CurrentSqlColumn = NzString(grdMapping.TextMatrix(TargetRow, 1))
    
    With cmbGridSqlColumns
        .Left = grdMapping.Left + grdMapping.CellLeft
        .Top = grdMapping.Top + grdMapping.CellTop
        .Width = grdMapping.CellWidth
        .Visible = True
        .ZOrder 0
    End With
    
    If Len(CurrentSqlColumn) > 0 Then
        ComboIndex = FindComboIndexByText(cmbGridSqlColumns, CurrentSqlColumn)
        If ComboIndex >= 0 Then
            cmbGridSqlColumns.ListIndex = ComboIndex
        Else
            cmbGridSqlColumns.ListIndex = 0
        End If
    Else
        SuggestedSqlColumn = FindBestSqlColumnMatchForExcelHeader(NzString(grdMapping.TextMatrix(TargetRow, 0)))
        If Len(SuggestedSqlColumn) > 0 Then
            ComboIndex = FindComboIndexByText(cmbGridSqlColumns, SuggestedSqlColumn)
            If ComboIndex >= 0 Then
                cmbGridSqlColumns.ListIndex = ComboIndex
            Else
                cmbGridSqlColumns.ListIndex = 0
            End If
        Else
            cmbGridSqlColumns.ListIndex = 0
        End If
    End If
    
    cmbGridSqlColumns.SetFocus
    Exit Sub

ErrorHandler:
    cmbGridSqlColumns.Visible = False
End Sub

Private Function FindComboIndexByText(ByVal Combo As ComboBox, ByVal SearchText As String) As Integer
    Dim i As Integer
    
    FindComboIndexByText = -1
    
    For i = 0 To Combo.ListCount - 1
        If StrComp(Trim$(Combo.List(i)), Trim$(SearchText), vbTextCompare) = 0 Then
            FindComboIndexByText = i
            Exit Function
        End If
    Next i
End Function

Private Function NormalizeNameForUiMatch(ByVal Value As String) As String
    Dim Result As String
    
    Result = LCase$(Trim$(Value))
    Result = Replace(Result, " ", "")
    Result = Replace(Result, "_", "")
    Result = Replace(Result, "-", "")
    Result = Replace(Result, ".", "")
    
    NormalizeNameForUiMatch = Result
End Function

Private Function FindBestSqlColumnMatchForExcelHeader(ByVal ExcelHeader As String) As String
    Dim i As Long
    Dim ColInfo As clsColumnInfo
    Dim NormalizedExcel As String
    Dim NormalizedSql As String
    
    FindBestSqlColumnMatchForExcelHeader = ""
    
    NormalizedExcel = NormalizeNameForUiMatch(ExcelHeader)
    
    If Len(NormalizedExcel) = 0 Then Exit Function
    
    For i = 1 To mTargetColumns.Count
        Set ColInfo = mTargetColumns(i)
        NormalizedSql = NormalizeNameForUiMatch(ColInfo.ColumnName)
        
        If StrComp(NormalizedExcel, NormalizedSql, vbTextCompare) = 0 Then
            FindBestSqlColumnMatchForExcelHeader = ColInfo.ColumnName
            Exit Function
        End If
    Next i
End Function

Private Function BuildImportOptionsFromForm() As clsImportOptions
    Dim Options As clsImportOptions
    Dim i As Long
    Dim ColInfo As clsColumnInfo
    Dim HasIdentityMapping As Boolean
    Dim Mappings As Collection
    Dim MapItem As clsMappingItem
    
    Set Options = New clsImportOptions
    
    Options.DeleteExistingData = (chkDeleteExisting.Value = vbChecked)
    Options.ContinueOnTypeError = (chkContinueOnTypeError.Value = vbChecked)
    Options.ContinueOnDuplicate = (chkContinueOnDuplicate.Value = vbChecked)
    Options.AutoLoadSavedMapping = (chkAutoLoadSavedMapping.Value = vbChecked)
    Options.ExcelReadMode = GetSelectedExcelReadMode()
    
    HasIdentityMapping = False
    Set Mappings = CollectMappingsFromGrid()
    
    For i = 1 To Mappings.Count
        Set MapItem = Mappings(i)
        
        If MapItem.IsMapped Then
            Set ColInfo = FindTargetColumn(MapItem.SqlColumnName)
            If Not ColInfo Is Nothing Then
                If ColInfo.IsIdentity Then
                    HasIdentityMapping = True
                    Exit For
                End If
            End If
        End If
    Next i
    
    Options.UseIdentityInsert = HasIdentityMapping
    
    Set BuildImportOptionsFromForm = Options
End Function

Private Function GetExcelReadModeText(ByVal ReadMode As Integer) As String
    Select Case ReadMode
        Case EXCEL_READ_MODE_AUTOMATION
            GetExcelReadModeText = "Excel Automation"
        Case Else
            GetExcelReadModeText = "ADO/OLEDB"
    End Select
End Function

Private Sub ResetExecutionProgress()
    prgImport.Min = 0
    prgImport.Max = PROGRESS_MAX
    prgImport.Value = 0
    
    lblCurrentStep.Caption = "Ready"
    lblRowProgress.Caption = "0 / 0"
    lblSuccessCount.Caption = "0"
    lblSkipCount.Caption = "0"
    lblErrorCount.Caption = "0"
End Sub

Private Sub SetImportUiBusyState(ByVal IsBusy As Boolean)
    'FraConnection.Enabled = Not IsBusy
    'fraDatabaseTable.Enabled = Not IsBusy
    fraExcel.Enabled = Not IsBusy
    
    grdMapping.Enabled = Not IsBusy
    cmbGridSqlColumns.Visible = False
    
    cmdAutoMatch.Enabled = Not IsBusy
    cmdLoadSavedMapping.Enabled = Not IsBusy
    cmdSaveMapping.Enabled = Not IsBusy
    cmdValidateMapping.Enabled = Not IsBusy
    
    chkDeleteExisting.Enabled = Not IsBusy
    chkContinueOnTypeError.Enabled = Not IsBusy
    chkContinueOnDuplicate.Enabled = Not IsBusy
    chkAutoLoadSavedMapping.Enabled = Not IsBusy
    
    cmdStartImport.Enabled = Not IsBusy
    cmdExit.Enabled = Not IsBusy
    cmdCancelImport.Enabled = IsBusy
End Sub

Private Sub ResetExecutionCountersOnly()
    lblRowProgress.Caption = "0 / 0"
    lblSuccessCount.Caption = "0"
    lblSkipCount.Caption = "0"
    lblErrorCount.Caption = "0"
    prgImport.Value = 0
End Sub

Private Sub SetExecutionSummary( _
    ByVal SuccessCount As Long, _
    ByVal SkipCount As Long, _
    ByVal ErrorCount As Long, _
    ByVal TotalRows As Long)
    
    lblSuccessCount.Caption = CStr(SuccessCount)
    lblSkipCount.Caption = CStr(SkipCount)
    lblErrorCount.Caption = CStr(ErrorCount)
    lblRowProgress.Caption = CStr(TotalRows) & " / " & CStr(TotalRows)
    prgImport.Value = PROGRESS_MAX
End Sub

'Private Sub ResetAfterTableSelectionChange()
'    txtExcelFile.Text = ""
'
'    Set mExcelHeaders = New Collection
'    ResetMappingGridRows
'
'    grdMapping.Enabled = False
'    cmdAutoMatch.Enabled = False
'    cmdLoadSavedMapping.Enabled = False
'    cmdSaveMapping.Enabled = False
'    cmdValidateMapping.Enabled = False
'
'    chkDeleteExisting.Value = vbUnchecked
'    chkContinueOnTypeError.Value = vbUnchecked
'    chkContinueOnDuplicate.Value = vbUnchecked
'
'    chkDeleteExisting.Enabled = False
'    chkContinueOnTypeError.Enabled = False
'    chkContinueOnDuplicate.Enabled = False
'    chkAutoLoadSavedMapping.Enabled = False
'
'    cmdStartImport.Enabled = False
'    cmbGridSqlColumns.Visible = False
'
'    ResetExecutionProgress
'    ResetPreviewGrid
'End Sub

Private Function ContainsText(ByVal SourceText As String, ByVal SearchText As String) As Boolean
    If Len(Trim$(SearchText)) = 0 Then
        ContainsText = True
    Else
        ContainsText = (InStr(1, SourceText, SearchText, vbTextCompare) > 0)
    End If
End Function

'Private Sub FilterDatabases()
'    Dim i As Long
'    Dim DbName As String
'    Dim SearchText As String
'
'    SearchText = Trim$(txtSearchDatabase.Text)
'
'    cmbDatabase.Clear
'
'    For i = 1 To mAllDatabases.Count
'        DbName = CStr(mAllDatabases(i))
'
'        If ContainsText(DbName, SearchText) Then
'            cmbDatabase.AddItem DbName
'        End If
'    Next i
'
'    If cmbDatabase.ListCount > 0 Then
'        cmbDatabase.ListIndex = 0
'    End If
'End Sub

'Private Sub txtSearchDatabase_Change()
'    If fraDatabaseTable.Enabled Then
'        FilterDatabases
'    End If
'End Sub

'Private Sub FilterTables()
'    Dim i As Long
'    Dim TableName As String
'    Dim SearchText As String
'
'    SearchText = Trim$(txtSearchTable.Text)
'
'    cmbTable.Clear
'
'    For i = 1 To mAllTables.Count
'        TableName = CStr(mAllTables(i))
'
'        If ContainsText(TableName, SearchText) Then
'            cmbTable.AddItem TableName
'        End If
'    Next i
'
'    If cmbTable.ListCount > 0 Then
'        cmbTable.ListIndex = 0
'    End If
'End Sub

'Private Sub txtSearchTable_Change()
'    If fraDatabaseTable.Enabled Then
'        FilterTables
'    End If
'End Sub

'Private Sub ResetAfterConnectionSettingsChange()
'    On Error Resume Next
'
'    If Not mSqlConnection Is Nothing Then
'        mSqlConnection.CloseConnection
'    End If
'
'    cmbDatabase.Clear
'    cmbTable.Clear
'
'    txtSearchDatabase.Text = ""
'    txtSearchTable.Text = ""
'
'    Set mAllDatabases = New Collection
'    Set mAllTables = New Collection
'    Set mTargetColumns = New Collection
'    Set mExcelHeaders = New Collection
'
'    txtExcelFile.Text = ""
'
'    ResetMappingGridRows
'    ResetExecutionProgress
'
'    fraDatabaseTable.Enabled = False
'    fraExcel.Enabled = False
'
'    grdMapping.Enabled = False
'    cmbGridSqlColumns.Visible = False
'
'    cmdAutoMatch.Enabled = False
'    cmdLoadSavedMapping.Enabled = False
'    cmdSaveMapping.Enabled = False
'    cmdValidateMapping.Enabled = False
'    cmdStartImport.Enabled = False
'
'    chkDeleteExisting.Value = vbUnchecked
'    chkContinueOnTypeError.Value = vbUnchecked
'    chkContinueOnDuplicate.Value = vbUnchecked
'
'    chkDeleteExisting.Enabled = False
'    chkContinueOnTypeError.Enabled = False
'    chkContinueOnDuplicate.Enabled = False
'    chkAutoLoadSavedMapping.Enabled = False
'
'    lblCurrentStep.Caption = "Ready"
'    SetGlobalStatus "Not connected"
'End Sub

'Private Sub txtServerName_Change()
'    If fraDatabaseTable.Enabled Or fraExcel.Enabled Then
'        ResetAfterConnectionSettingsChange
'        AppendStatusMessage "Server name changed. Database and table selections were cleared."
'    End If
'End Sub

'Private Sub txtServerName_LostFocus()
'    If fraDatabaseTable.Enabled Then
'        ResetAfterConnectionSettingsChange
'        AppendStatusMessage "Server name changed. Database and table selections were cleared."
'    End If
'End Sub

Private Sub InitializePreviewGrid()
    With grdExcelPreview
        .Rows = 2
        .Cols = 1
        .FixedRows = 1
        .FixedCols = 0
        .SelectionMode = 0
        
        .TextMatrix(0, 0) = "Preview"
        .TextMatrix(1, 0) = ""
    End With
End Sub

Private Sub ResetPreviewGrid()
    With grdExcelPreview
        .Rows = 2
        .Cols = 1
        .FixedRows = 1
        .FixedCols = 0
        
        .TextMatrix(0, 0) = "Preview"
        .TextMatrix(1, 0) = ""
    End With
End Sub

Private Sub LoadExcelPreview()
    On Error GoTo ErrorHandler
    
    Dim PreviewRows As Collection
    Dim RowData As Object
    Dim PreviewRowCount As Long
    Dim PreviewColCount As Long
    Dim r As Long
    Dim c As Long
    Dim HeaderName As String
    Dim CellValue As String
    
    SetGlobalStatus "Loading Excel preview..."
    AppendStatusMessage "Loading Excel preview..."
    DoEvents
    
    mExcelReader.ReadMode = GetSelectedExcelReadMode()
    Set PreviewRows = mExcelReader.LoadRows(Trim$(txtExcelFile.Text), Nothing, 5)
    
    If mExcelHeaders Is Nothing Then Exit Sub
    If mExcelHeaders.Count = 0 Then Exit Sub
    
    PreviewColCount = mExcelHeaders.Count
    
    If PreviewRows.Count >= 5 Then
        PreviewRowCount = 5
    Else
        PreviewRowCount = PreviewRows.Count
    End If
    
    With grdExcelPreview
        .Rows = IIf(PreviewRowCount = 0, 2, PreviewRowCount + 1)
        .Cols = IIf(PreviewColCount = 0, 1, PreviewColCount)
        .FixedRows = 1
        .FixedCols = 0
        
        For c = 1 To mExcelHeaders.Count
            .TextMatrix(0, c - 1) = CStr(mExcelHeaders(c))
        Next c
        
        If PreviewRowCount = 0 Then
            If .Cols > 0 Then
                .TextMatrix(1, 0) = ""
            End If
        Else
            For r = 1 To PreviewRowCount
                Set RowData = PreviewRows(r)
                
                For c = 1 To mExcelHeaders.Count
                    HeaderName = CStr(mExcelHeaders(c))
                    
                    If RowData.Exists(HeaderName) Then
                        If IsNull(RowData(HeaderName)) Or IsEmpty(RowData(HeaderName)) Then
                            CellValue = ""
                        Else
                            CellValue = CStr(RowData(HeaderName))
                        End If
                    Else
                        CellValue = ""
                    End If
                    
                    .TextMatrix(r, c - 1) = CellValue
                Next c
            Next r
        End If
    End With
    
    AdjustPreviewGridColumnWidths
    
    AppendStatusMessage CStr(PreviewRowCount) & " preview rows loaded."
    SetGlobalStatus "Excel preview loaded"
    Exit Sub

ErrorHandler:
    AppendStatusMessage "Failed to load Excel preview: " & Err.Description
    MsgBox "Failed to load Excel preview." & vbCrLf & Err.Description, vbExclamation, APP_NAME
End Sub

Private Sub AdjustPreviewGridColumnWidths()
    Dim c As Long
    
    On Error Resume Next
    
    For c = 0 To grdExcelPreview.Cols - 1
        grdExcelPreview.ColWidth(c) = 1800
    Next c
End Sub

Private Function FindMappedRowBySqlColumn(ByVal SqlColumnName As String, ByVal IgnoreRowIndex As Long) As Long
    Dim i As Long
    Dim CurrentSqlColumn As String
    
    FindMappedRowBySqlColumn = 0
    
    If Len(Trim$(SqlColumnName)) = 0 Then Exit Function
    
    For i = 1 To grdMapping.Rows - 1
        If i <> IgnoreRowIndex Then
            CurrentSqlColumn = NzString(grdMapping.TextMatrix(i, 1))
            
            If StrComp(Trim$(CurrentSqlColumn), Trim$(SqlColumnName), vbTextCompare) = 0 Then
                FindMappedRowBySqlColumn = i
                Exit Function
            End If
        End If
    Next i
End Function

Private Function CanAssignSqlColumnToRow(ByVal RowIndex As Long, ByVal SqlColumnName As String) As Boolean
    Dim DuplicateRow As Long
    
    CanAssignSqlColumnToRow = False
    
    If Len(Trim$(SqlColumnName)) = 0 Then
        CanAssignSqlColumnToRow = True
        Exit Function
    End If
    
    DuplicateRow = FindMappedRowBySqlColumn(SqlColumnName, RowIndex)
    
    If DuplicateRow > 0 Then
        MsgBox "This SQL column is already mapped in another row." & vbCrLf & vbCrLf & _
               "SQL Column: " & SqlColumnName & vbCrLf & _
               "Already used in Excel Column: " & grdMapping.TextMatrix(DuplicateRow, 0), _
               vbExclamation, APP_NAME
        Exit Function
    End If
    
    CanAssignSqlColumnToRow = True
End Function

Private Function ValidateNoDuplicateSqlMappings() As Boolean
    Dim i As Long
    Dim j As Long
    Dim SqlCol1 As String
    Dim SqlCol2 As String
    
    ValidateNoDuplicateSqlMappings = True
    
    For i = 1 To grdMapping.Rows - 1
        SqlCol1 = NzString(grdMapping.TextMatrix(i, 1))
        
        If Len(SqlCol1) > 0 Then
            For j = i + 1 To grdMapping.Rows - 1
                SqlCol2 = NzString(grdMapping.TextMatrix(j, 1))
                
                If Len(SqlCol2) > 0 Then
                    If StrComp(SqlCol1, SqlCol2, vbTextCompare) = 0 Then
                        MsgBox "Duplicate SQL mapping found." & vbCrLf & vbCrLf & _
                               "SQL Column: " & SqlCol1 & vbCrLf & _
                               "Excel Columns: " & grdMapping.TextMatrix(i, 0) & " and " & grdMapping.TextMatrix(j, 0), _
                               vbExclamation, APP_NAME
                        ValidateNoDuplicateSqlMappings = False
                        Exit Function
                    End If
                End If
            Next j
        End If
    Next i
End Function

Public Property Get CancelRequested() As Boolean
    CancelRequested = mCancelRequested
End Property

Private Sub CheckForCancelRequest()
    If mCancelRequested Then
        Err.Raise vbObjectError + 3001, "frmImportExcelToSql.CheckForCancelRequest", "Import was cancelled by user."
    End If
End Sub

Private Sub SetCancelledExecutionState()
    lblCurrentStep.Caption = "Import cancelled"
    SetGlobalStatus "Import cancelled by user"
    prgImport.Value = 0
End Sub

Private Function GetSettingsFilePath() As String
    GetSettingsFilePath = App.Path
    
    If Right$(GetSettingsFilePath, 1) <> "\" Then
        GetSettingsFilePath = GetSettingsFilePath & "\"
    End If
    
    GetSettingsFilePath = GetSettingsFilePath & SETTINGS_FILE_NAME
End Function

'Private Sub SaveConnectionSettings()
'    Dim SettingsFilePath As String
'
'    SettingsFilePath = GetSettingsFilePath()
'
'    WriteIniValue SETTINGS_SECTION_CONNECTION, "RememberSettings", IIf(chkRememberConnectionSettings.Value = vbChecked, "1", "0"), SettingsFilePath
'    WriteIniValue SETTINGS_SECTION_CONNECTION, "ServerName", Trim$(txtServerName.Text), SettingsFilePath
'    WriteIniValue SETTINGS_SECTION_CONNECTION, "AuthenticationType", Trim$(cmbAuthentication.Text), SettingsFilePath
'    WriteIniValue SETTINGS_SECTION_CONNECTION, "UserName", Trim$(txtUserName.Text), SettingsFilePath
'    WriteIniValue SETTINGS_SECTION_CONNECTION, "ExcelReadMode", CStr(GetSelectedExcelReadMode()), SettingsFilePath
'End Sub
'
'Private Sub ClearSavedConnectionSettings()
'    Dim SettingsFilePath As String
'
'    SettingsFilePath = GetSettingsFilePath()
'
'    DeleteIniKey SETTINGS_SECTION_CONNECTION, "RememberSettings", SettingsFilePath
'    DeleteIniKey SETTINGS_SECTION_CONNECTION, "ServerName", SettingsFilePath
'    DeleteIniKey SETTINGS_SECTION_CONNECTION, "AuthenticationType", SettingsFilePath
'    DeleteIniKey SETTINGS_SECTION_CONNECTION, "UserName", SettingsFilePath
'    DeleteIniKey SETTINGS_SECTION_CONNECTION, "ExcelReadMode", SettingsFilePath
'End Sub

'Private Sub LoadConnectionSettings()
'    On Error GoTo ErrorHandler
'
'    Dim SettingsFilePath As String
'    Dim RememberValue As String
'    Dim ServerName As String
'    Dim AuthenticationType As String
'    Dim UserName As String
'    Dim ExcelReadModeText As String
'
'    SettingsFilePath = GetSettingsFilePath()
'
'    RememberValue = ReadIniValue(SETTINGS_SECTION_CONNECTION, "RememberSettings", "0", SettingsFilePath)
'
'    If RememberValue <> "1" Then Exit Sub
'
'    chkRememberConnectionSettings.Value = vbChecked
'
'    ServerName = ReadIniValue(SETTINGS_SECTION_CONNECTION, "ServerName", "", SettingsFilePath)
'    AuthenticationType = ReadIniValue(SETTINGS_SECTION_CONNECTION, "AuthenticationType", AUTH_WINDOWS, SettingsFilePath)
'    UserName = ReadIniValue(SETTINGS_SECTION_CONNECTION, "UserName", "", SettingsFilePath)
'    ExcelReadModeText = ReadIniValue(SETTINGS_SECTION_CONNECTION, "ExcelReadMode", CStr(EXCEL_READ_MODE_ADO), SettingsFilePath)
'
'    txtServerName.Text = ServerName
'
'    If StrComp(AuthenticationType, AUTH_SQL_SERVER, vbTextCompare) = 0 Then
'        cmbAuthentication.ListIndex = 1
'    Else
'        cmbAuthentication.ListIndex = 0
'    End If
'
'    UpdateAuthenticationControls
'
'    txtUserName.Text = UserName
'    txtPassword.Text = ""
'
'    If Val(ExcelReadModeText) = EXCEL_READ_MODE_AUTOMATION Then
'        optReadByExcel.Value = True
'        optReadByAdo.Value = False
'    Else
'        optReadByAdo.Value = True
'        optReadByExcel.Value = False
'    End If
'
'    Exit Sub
'
'ErrorHandler:
'    chkRememberConnectionSettings.Value = vbUnchecked
'End Sub

'Private Sub chkRememberConnectionSettings_Click()
'    If chkRememberConnectionSettings.Value = vbUnchecked Then
'        ClearSavedConnectionSettings
'        AppendStatusMessage "Saved connection settings were cleared."
'    End If
'End Sub

'Private Sub ApplyAppContextToImportForm()
'    EnsureImportObjects
'
'    lblConnectionInfo.Caption = gAppContext.ServerName
'    lblDatabaseInfo.Caption = gAppContext.SelectedDatabase
'    lblTableInfo.Caption = gAppContext.GetFullTargetName()
'End Sub

Private Function GetActiveConnection() As ADODB.Connection
    On Error GoTo ErrorHandler
    
    If frmMain Is Nothing Then
        Set GetActiveConnection = Nothing
        Exit Function
    End If
    
    If frmMain.SqlConnectionManager Is Nothing Then
        Set GetActiveConnection = Nothing
        Exit Function
    End If
    
    If frmMain.SqlConnectionManager.Connection Is Nothing Then
        Set GetActiveConnection = Nothing
        Exit Function
    End If
    
    Set GetActiveConnection = frmMain.SqlConnectionManager.Connection
    Exit Function

ErrorHandler:
    Set GetActiveConnection = Nothing
End Function

Private Sub ApplyAppContextToImportForm()
    lblConnectionInfo.Caption = gAppContext.ServerName
    lblDatabaseInfo.Caption = gAppContext.SelectedDatabase
    lblTableInfo.Caption = gAppContext.GetFullTargetName()
End Sub

Private Sub LoadTargetTableStructureFromContext()
    On Error GoTo ErrorHandler
    
    Dim Conn As ADODB.Connection
    
    Set Conn = GetActiveConnection()
    
    If Conn Is Nothing Then
        MsgBox "No active database connection is available.", vbExclamation, APP_NAME
        Exit Sub
    End If
    
    Set mTargetColumns = mDatabaseBrowser.GetTableColumns( _
        Conn, _
        gAppContext.SelectedDatabase, _
        gAppContext.SelectedSchema, _
        gAppContext.SelectedTable)
    
    If Not mMappingManager Is Nothing Then
        mMappingManager.EnsureMappingTableExists Conn, gAppContext.SelectedDatabase
    End If
    
    Exit Sub

ErrorHandler:
    MsgBox "Failed to load target table structure." & vbCrLf & _
           Err.Number & " - " & Err.Description, vbCritical, APP_NAME
End Sub

Private Sub EnsureImportObjects()
    If mExcelReader Is Nothing Then Set mExcelReader = New clsExcelReader
    If mMappingManager Is Nothing Then Set mMappingManager = New clsMappingManager
    If mStagingManager Is Nothing Then Set mStagingManager = New clsStagingManager
    If mImportEngine Is Nothing Then Set mImportEngine = New clsImportEngine
    If mDatabaseBrowser Is Nothing Then Set mDatabaseBrowser = New clsDatabaseBrowser
    If mTargetColumns Is Nothing Then Set mTargetColumns = New Collection
    If mExcelHeaders Is Nothing Then Set mExcelHeaders = New Collection
End Sub
