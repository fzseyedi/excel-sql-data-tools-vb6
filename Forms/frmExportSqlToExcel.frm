VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmExportSqlToExcel 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Export SQL Table(s) to Excel / CSV"
   ClientHeight    =   9015
   ClientLeft      =   12585
   ClientTop       =   6750
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
   Begin VB.Frame fraExecution 
      Caption         =   " Execution "
      Height          =   1215
      Left            =   8760
      TabIndex        =   10
      Top             =   5760
      Width           =   6015
      Begin VB.CommandButton cmdStartExport 
         Caption         =   "Start Export"
         Height          =   360
         Left            =   720
         TabIndex        =   11
         Top             =   480
         Width           =   2190
      End
      Begin VB.CommandButton cmdCancelExport 
         Caption         =   "Cancel Export"
         Height          =   360
         Left            =   3000
         TabIndex        =   12
         Top             =   480
         Width           =   2190
      End
   End
   Begin VB.TextBox txtStatus 
      Height          =   1095
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   29
      Top             =   7800
      Width           =   14655
   End
   Begin VB.Frame fraExportOptions 
      Caption         =   " Export Options "
      Height          =   2535
      Left            =   8760
      TabIndex        =   3
      Top             =   840
      Width           =   6015
      Begin VB.Frame FraOutputFolder 
         Caption         =   " Output Folder "
         Height          =   975
         Left            =   120
         TabIndex        =   7
         Top             =   1440
         Width           =   5775
         Begin VB.CommandButton cmdBrowseOutputFolder 
            Caption         =   "Browse"
            Height          =   360
            Left            =   4800
            TabIndex        =   9
            Top             =   360
            Width           =   870
         End
         Begin VB.TextBox txtOutputFolder 
            Height          =   375
            Left            =   240
            TabIndex        =   8
            Top             =   360
            Width           =   4575
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   " Output Type "
         Height          =   975
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   5775
         Begin VB.OptionButton optExportCsv 
            Caption         =   "CSV"
            Height          =   375
            Left            =   3360
            TabIndex        =   6
            Top             =   360
            Width           =   1215
         End
         Begin VB.OptionButton optExportExcel 
            Caption         =   "Excel"
            Height          =   375
            Left            =   1920
            TabIndex        =   5
            Top             =   360
            Width           =   1215
         End
      End
   End
   Begin VB.Frame fraFields 
      Caption         =   " Fields "
      Height          =   6165
      Left            =   4440
      TabIndex        =   20
      Top             =   840
      Width           =   4215
      Begin VB.ListBox lstFields 
         Height          =   5190
         Left            =   120
         Style           =   1  'Checkbox
         TabIndex        =   2
         Top             =   840
         Width           =   3975
      End
      Begin VB.Label lblSelectedTable 
         BackStyle       =   0  'Transparent
         Height          =   240
         Left            =   2280
         TabIndex        =   31
         Top             =   360
         Width           =   1740
      End
      Begin VB.Label lblSelectedTableCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fields for selected table"
         Height          =   240
         Left            =   120
         TabIndex        =   30
         Top             =   360
         Width           =   2040
      End
   End
   Begin VB.Frame fraContextInfo 
      Caption         =   "Current Context"
      Height          =   855
      Left            =   113
      TabIndex        =   13
      Top             =   0
      Width           =   14655
      Begin VB.Label lblDatabaseInfo 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         Height          =   360
         Left            =   9480
         TabIndex        =   17
         Top             =   360
         Width           =   4995
      End
      Begin VB.Label lblDatabaseInfoCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Database :"
         Height          =   240
         Left            =   8400
         TabIndex        =   16
         Top             =   360
         Width           =   930
      End
      Begin VB.Label lblServerInfo 
         BackColor       =   &H00C0FFFF&
         Height          =   360
         Left            =   960
         TabIndex        =   15
         Top             =   360
         Width           =   4995
      End
      Begin VB.Label lblServerInfoCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Server :"
         Height          =   240
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Width           =   705
      End
   End
   Begin VB.Frame fraTables 
      Caption         =   " Tables "
      Height          =   6165
      Left            =   120
      TabIndex        =   18
      Top             =   840
      Width           =   4215
      Begin VB.ListBox lstTables 
         Height          =   5190
         Left            =   120
         Style           =   1  'Checkbox
         TabIndex        =   1
         Top             =   840
         Width           =   3975
      End
      Begin VB.TextBox txtSearchTables 
         Height          =   375
         Left            =   1680
         TabIndex        =   0
         Top             =   300
         Width           =   2415
      End
      Begin VB.Label lblSearchTables 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Search tables :"
         Height          =   240
         Left            =   240
         TabIndex        =   19
         Top             =   300
         Width           =   1305
      End
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   120
      TabIndex        =   21
      Top             =   6960
      Width           =   14655
      Begin MSComctlLib.ProgressBar prgExport 
         Height          =   375
         Left            =   7920
         TabIndex        =   28
         Top             =   240
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
      Begin VB.Label lblSuccessCount 
         BackColor       =   &H00C0FFFF&
         Caption         =   "-"
         Height          =   360
         Left            =   3120
         TabIndex        =   33
         Top             =   240
         Width           =   1005
      End
      Begin VB.Label lblSuccessCountCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Success :"
         Height          =   240
         Left            =   2280
         TabIndex        =   32
         Top             =   240
         Width           =   825
      End
      Begin VB.Label lblErrorCount 
         BackColor       =   &H00C0FFFF&
         Caption         =   "-"
         Height          =   360
         Left            =   6840
         TabIndex        =   27
         Top             =   240
         Width           =   1005
      End
      Begin VB.Label lblErrorCountCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Errors :"
         Height          =   240
         Left            =   6120
         TabIndex        =   26
         Top             =   240
         Width           =   660
      End
      Begin VB.Label lblSkipCount 
         BackColor       =   &H00C0FFFF&
         Caption         =   "-"
         Height          =   360
         Left            =   5040
         TabIndex        =   25
         Top             =   240
         Width           =   1005
      End
      Begin VB.Label lblSkipCountCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Skipped :"
         Height          =   240
         Left            =   4200
         TabIndex        =   24
         Top             =   240
         Width           =   810
      End
      Begin VB.Label lblCurrentStep 
         BackColor       =   &H00C0FFFF&
         Caption         =   "-"
         Height          =   360
         Left            =   720
         TabIndex        =   23
         Top             =   240
         Width           =   1485
      End
      Begin VB.Label lblCurrentStepCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Step :"
         Height          =   240
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   525
      End
   End
   Begin VB.Menu mnuListPopup 
      Caption         =   "List Popup"
      Visible         =   0   'False
      Begin VB.Menu mnuListSelectAll 
         Caption         =   "Select All"
      End
      Begin VB.Menu mnuCaption 
         Caption         =   "-"
      End
      Begin VB.Menu mnuListSelectNone 
         Caption         =   "Select None"
      End
   End
End
Attribute VB_Name = "frmExportSqlToExcel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mDatabaseBrowser As clsDatabaseBrowser
Private mExportSelections As Collection
Private mSuppressUiEvents As Boolean
Private mCancelRequested As Boolean
Private mPopupListSource As String
Private mIsExporting As Boolean

Private Sub cmdBrowseOutputFolder_Click()
    On Error GoTo ErrorHandler
    
    Dim ShellApp As Object
    Dim FolderObj As Object
    Dim SelectedPath As String
    
    Set ShellApp = CreateObject("Shell.Application")
    Set FolderObj = ShellApp.BrowseForFolder(0, "Select output folder", 0)
    
    If FolderObj Is Nothing Then Exit Sub
    
    SelectedPath = FolderObj.Items.Item.Path
    
    If Len(Trim$(SelectedPath)) > 0 Then
        txtOutputFolder.Text = SelectedPath
        RefreshExportUiState
    End If
    
    Exit Sub

ErrorHandler:
    MsgBox "Failed to select output folder." & vbCrLf & _
           Err.Number & " - " & Err.Description, vbExclamation, APP_NAME
End Sub

Private Sub cmdCancelExport_Click()
    If MsgBox("Do you want to cancel the current export operation?", _
              vbQuestion + vbYesNo, APP_NAME) = vbYes Then
        
        mCancelRequested = True
        cmdCancelExport.Enabled = False
        lblCurrentStep.Caption = "Cancel requested..."
        AppendExportStatus "Cancel requested by user."
        DoEvents
    End If
End Sub

Private Sub cmdStartExport_Click()
    On Error GoTo ErrorHandler
    
    Dim i As Long
    Dim Selection As clsExportTableSelection
    Dim TotalTables As Long
    Dim CurrentTable As Long
    Dim ExportedRows As Long
    Dim SuccessCount As Long
    Dim SkipCount As Long
    Dim ErrorCount As Long
    Dim OutputFolder As String
    Dim ExcelExportMode As String
    
    If Not ValidateExportInputs() Then Exit Sub
    If Not ValidateExcelExportAvailability(ExcelExportMode) Then Exit Sub
    
    OutputFolder = Trim$(txtOutputFolder.Text)
    
    mCancelRequested = False
    
    ResetExportProgress
    SetExportUiBusyState True
    
    DoEvents

    TotalTables = GetCheckedTableCount()
    CurrentTable = 0
    SuccessCount = 0
    SkipCount = 0
    ErrorCount = 0
    
    AppendExportStatus "Export started."
        
    For i = 1 To mExportSelections.Count
        Set Selection = mExportSelections(i)
        
        If Selection.IsChecked Then
            CheckForExportCancel
            
            CurrentTable = CurrentTable + 1
            
            lblCurrentStep.Caption = "Exporting " & Selection.FullTableName
            '''lblProgress.Caption = CStr(CurrentTable) & " / " & CStr(TotalTables)
            
            If TotalTables > 0 Then
                prgExport.Value = CLng((CurrentTable / TotalTables) * PROGRESS_MAX)
            End If
            
            DoEvents
            
            On Error GoTo TableError
            
            If optExportCsv.Value Then
                ExportedRows = ExportTableToCsv(Selection, OutputFolder)
            
            ElseIf optExportExcel.Value Then
                If ExcelExportMode = "XLSX_AUTOMATION" Then
                    ExportedRows = ExportTableToXlsxWithAutomation(Selection, OutputFolder)
                    
                ElseIf ExcelExportMode = "XLS_ADO" Then
                    ExportedRows = ExportTableToXlsWithAdo(Selection, OutputFolder)
                    
                Else
                    MsgBox "Excel export is not available. Please select CSV export instead.", vbExclamation, APP_NAME
                    GoTo SafeExit
                End If
            End If
            
            If ExportedRows = -1 Then
                SkipCount = SkipCount + 1
                AppendExportStatus "Skipped: " & Selection.FullTableName
            Else
                SuccessCount = SuccessCount + 1
                AppendExportStatus "Exported: " & Selection.FullTableName & " (" & CStr(ExportedRows) & " rows)"
            End If
            
            lblSuccessCount.Caption = CStr(SuccessCount)
            lblSkipCount.Caption = CStr(SkipCount)
            lblErrorCount.Caption = CStr(ErrorCount)
            
            On Error GoTo ErrorHandler
        End If
    Next i
    
    lblCurrentStep.Caption = "Export completed"
    '''lblProgress.Caption = CStr(TotalTables) & " / " & CStr(TotalTables)
    prgExport.Value = PROGRESS_MAX
    
    AppendExportStatus "Export completed."
    
    MsgBox "Export completed." & vbCrLf & vbCrLf & _
           "Tables exported: " & CStr(SuccessCount) & vbCrLf & _
           "Tables skipped: " & CStr(SkipCount) & vbCrLf & _
           "Tables failed: " & CStr(ErrorCount), _
           vbInformation, APP_NAME
    
SafeExit:
    mCancelRequested = False
    SetExportUiBusyState False
    RefreshExportUiState
    Exit Sub

TableError:
    ErrorCount = ErrorCount + 1
    AppendExportStatus "Failed: " & Selection.FullTableName & " - " & Err.Description
    
    lblErrorCount.Caption = CStr(ErrorCount)
    
    Err.Clear
    Resume Next

ErrorHandler:
    If Err.Number = vbObjectError + 5001 Then
        lblCurrentStep.Caption = "Export cancelled"
        AppendExportStatus "Export cancelled by user."
        MsgBox "Export was cancelled by user.", vbInformation, APP_NAME
    Else
        AppendExportStatus "Export failed: " & Err.Description
        MsgBox "Export failed." & vbCrLf & Err.Description, vbExclamation, APP_NAME
    End If
    
    Resume SafeExit
End Sub

Private Sub Form_Load()
    InitializeForm
End Sub

Private Sub InitializeForm()
    On Error GoTo ErrorHandler
    
    mSuppressUiEvents = True
    
    EnsureExportObjects
    
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
    
    If Len(Trim$(gAppContext.SelectedDatabase)) = 0 Then
        MsgBox "No database is selected.", vbExclamation, APP_NAME
        Unload Me
        Exit Sub
    End If
    
    If GetActiveConnection() Is Nothing Then
        MsgBox "No active database connection is available.", vbExclamation, APP_NAME
        Unload Me
        Exit Sub
    End If
    
    lblServerInfo.Caption = gAppContext.ServerName
    lblDatabaseInfo.Caption = gAppContext.SelectedDatabase
    
    optExportCsv.Value = True
    optExportExcel.Value = False
    
    cmdCancelExport.Enabled = False
    cmdStartExport.Enabled = False
    
    lblCurrentStep.Caption = "Ready"
    '''lblProgress.Caption = "0 / 0"
    lblCurrentStep.Caption = "0"
    lblSkipCount.Caption = "0"
    lblErrorCount.Caption = "0"
    prgExport.Value = 0
    txtStatus.Text = ""
    
    LoadTablesForExport
    ApplyDefaultSelectedTableFromContext
    RefreshExportUiState
    
    mSuppressUiEvents = False
    Exit Sub

ErrorHandler:
    mSuppressUiEvents = False
    MsgBox "Error initializing export form." & vbCrLf & _
           Err.Number & " - " & Err.Description, vbCritical, APP_NAME
    Unload Me
End Sub

Private Sub EnsureExportObjects()
    If mDatabaseBrowser Is Nothing Then
        Set mDatabaseBrowser = New clsDatabaseBrowser
    End If
    
    If mExportSelections Is Nothing Then
        Set mExportSelections = New Collection
    End If
End Sub

Private Function FindTableSelection(ByVal FullTableName As String) As clsExportTableSelection
    Dim i As Long
    Dim Item As clsExportTableSelection
    
    For i = 1 To mExportSelections.Count
        Set Item = mExportSelections(i)
        If StrComp(Item.FullTableName, Trim$(FullTableName), vbTextCompare) = 0 Then
            Set FindTableSelection = Item
            Exit Function
        End If
    Next i
    
    Set FindTableSelection = Nothing
End Function

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

Private Sub LoadTablesForExport()
    On Error GoTo ErrorHandler
    
    Dim Rs As ADODB.Recordset
    Dim Item As clsExportTableSelection
    
    lstTables.Clear
    Set mExportSelections = New Collection
    
    Set Rs = mDatabaseBrowser.GetBaseTables(GetActiveConnection(), gAppContext.SelectedDatabase)
    
    Do While Not Rs.EOF
        Set Item = New clsExportTableSelection
        Item.SchemaName = NzString(Rs.Fields("TABLE_SCHEMA").Value)
        Item.TableName = NzString(Rs.Fields("TABLE_NAME").Value)
        Item.IsChecked = False
        
        mExportSelections.Add Item
        lstTables.AddItem Item.FullTableName
        
        Rs.MoveNext
    Loop
    
    Rs.Close
    Set Rs = Nothing
    Exit Sub

ErrorHandler:
    On Error Resume Next
    If Not Rs Is Nothing Then
        If Rs.State = adStateOpen Then Rs.Close
    End If
    Set Rs = Nothing
    
    MsgBox "Failed to load tables for export." & vbCrLf & _
           Err.Number & " - " & Err.Description, vbExclamation, APP_NAME
End Sub

Private Sub ApplyDefaultSelectedTableFromContext()
    On Error GoTo ErrorHandler
    
    Dim DefaultFullName As String
    Dim i As Long
    Dim Selection As clsExportTableSelection
    
    If gAppContext Is Nothing Then Exit Sub
    If Not gAppContext.HasSelectedTarget Then Exit Sub
    If lstTables.ListCount = 0 Then Exit Sub
    
    DefaultFullName = gAppContext.GetFullTargetName()
    
    For i = 0 To lstTables.ListCount - 1
        If StrComp(lstTables.List(i), DefaultFullName, vbTextCompare) = 0 Then
            
            ' «‰ Œ«» —œÌ› Ã«—Ì œ— ·Ì”  ÃœÊ·ùÂ«
            lstTables.ListIndex = i
            
            '  Ìò “œ‰ Â„«‰ ÃœÊ·
            lstTables.Selected(i) = True
            
            ' –ŒÌ—Â Ê÷⁄Ì  checked »Êœ‰ œ— ”«Œ «— Õ«›ŸÂù«Ì
            Set Selection = FindTableSelection(DefaultFullName)
            If Not Selection Is Nothing Then
                Selection.IsChecked = True
            End If
            
            ' ·Êœ ò—œ‰ ›Ì·œÂ«Ì Â„Ì‰ ÃœÊ· œ— ·Ì”  ›Ì·œÂ«
            LoadFieldsForSelectedTable
            
            ' »Âù—Ê“—”«‰Ì Ê÷⁄Ì  œò„ÂùÂ« Ê ›—„
            RefreshExportUiState
            
            Exit For
        End If
    Next i
    
    Exit Sub

ErrorHandler:
    MsgBox "Failed to apply default selected table." & vbCrLf & _
           Err.Number & " - " & Err.Description, vbExclamation, APP_NAME
End Sub

Private Sub LoadFieldsForSelectedTable()
    On Error GoTo ErrorHandler
    
    Dim Selection As clsExportTableSelection
    Dim i As Long
    Dim OldSuppress As Boolean
    
    If lstTables.ListIndex < 0 Then Exit Sub
    
    Set Selection = FindTableSelection(lstTables.List(lstTables.ListIndex))
    If Selection Is Nothing Then Exit Sub
    
    OldSuppress = mSuppressUiEvents
    mSuppressUiEvents = True
    
    lstFields.Clear
    lblSelectedTable.Caption = "-"
    
    If Not Selection.IsChecked Then
        lblSelectedTable.Caption = "-"
        mSuppressUiEvents = OldSuppress
        Exit Sub
    End If
    
    lblSelectedTable.Caption = Selection.FullTableName
    
    If Not Selection.IsFieldsLoaded Then
        LoadFieldsForTableSelection Selection
    End If
    
    For i = 1 To Selection.AllFields.Count
        lstFields.AddItem CStr(Selection.AllFields(i))
    Next i
    
    RestoreFieldChecksForCurrentTable
    
    mSuppressUiEvents = OldSuppress
    Exit Sub

ErrorHandler:
    mSuppressUiEvents = OldSuppress
    MsgBox "Failed to load fields for selected table." & vbCrLf & _
           Err.Number & " - " & Err.Description, vbExclamation, APP_NAME
End Sub

Private Sub RestoreFieldChecksForCurrentTable()
    Dim Selection As clsExportTableSelection
    Dim i As Long
    Dim FieldName As String
    
    If lstTables.ListIndex < 0 Then Exit Sub
    If lstFields.ListCount = 0 Then Exit Sub
    
    Set Selection = FindTableSelection(lstTables.List(lstTables.ListIndex))
    If Selection Is Nothing Then Exit Sub
    
    For i = 0 To lstFields.ListCount - 1
        FieldName = lstFields.List(i)
        lstFields.Selected(i) = Selection.IsFieldSelected(FieldName)
    Next i
End Sub

Private Sub SaveFieldChecksForCurrentTable()
    Dim Selection As clsExportTableSelection
    
    If mSuppressUiEvents Then Exit Sub
    If lstTables.ListIndex < 0 Then Exit Sub
    If lstFields.ListCount = 0 Then Exit Sub
    
    Set Selection = FindTableSelection(lstTables.List(lstTables.ListIndex))
    If Selection Is Nothing Then Exit Sub
    
    Selection.RebuildSelectedFieldsFromListBox lstFields
End Sub

Private Sub lstFields_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        mPopupListSource = "FIELDS"
        PopupMenu mnuListPopup
    End If
End Sub

Private Sub lstTables_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        mPopupListSource = "TABLES"
        PopupMenu mnuListPopup
    End If
End Sub

Private Sub lstTables_Click()
    Dim Selection As clsExportTableSelection
    
    If mSuppressUiEvents Then Exit Sub
    If lstTables.ListIndex < 0 Then Exit Sub
    
    Set Selection = FindTableSelection(lstTables.List(lstTables.ListIndex))
    If Selection Is Nothing Then Exit Sub
    
    If Selection.IsChecked Then
        LoadFieldsForSelectedTable
    Else
        lstFields.Clear
        lblSelectedTable.Caption = "-"
    End If
    
    RefreshExportUiState
End Sub

Private Sub lstTables_ItemCheck(Item As Integer)
    If mSuppressUiEvents Then Exit Sub
    
    SaveTableCheckState Item
    RefreshExportUiState
End Sub

Private Sub lstFields_ItemCheck(Item As Integer)
    If mSuppressUiEvents Then Exit Sub
    
    SaveFieldChecksForCurrentTable
    RefreshExportUiState
    
    MoveFieldCursorToNextItem Item
End Sub

Private Sub SaveTableCheckState(ByVal ItemIndex As Integer)
    Dim Selection As clsExportTableSelection
    Dim FullTableName As String
    
    If mSuppressUiEvents Then Exit Sub
    If ItemIndex < 0 Then Exit Sub
    If ItemIndex > lstTables.ListCount - 1 Then Exit Sub
    
    FullTableName = lstTables.List(ItemIndex)
    
    Set Selection = FindTableSelection(FullTableName)
    If Selection Is Nothing Then Exit Sub
    
    Selection.IsChecked = lstTables.Selected(ItemIndex)
    
    If Selection.IsChecked Then
        lstTables.ListIndex = ItemIndex
        LoadFieldsForSelectedTable
    Else
        If lstTables.ListIndex = ItemIndex Then
            lstFields.Clear
            lblSelectedTable.Caption = "-"
        End If
    End If
End Sub

Private Function HasCheckedTables() As Boolean
    Dim i As Long
    Dim Selection As clsExportTableSelection
    
    If mExportSelections Is Nothing Then Exit Function
    
    For i = 1 To mExportSelections.Count
        Set Selection = mExportSelections(i)
        
        If Selection.IsChecked Then
            HasCheckedTables = True
            Exit Function
        End If
    Next i
    
    HasCheckedTables = False
End Function

Private Function IsOutputFolderSelected() As Boolean
    IsOutputFolderSelected = (Len(Trim$(txtOutputFolder.Text)) > 0)
End Function

Private Sub RefreshExportUiState()
    Dim CanExport As Boolean
    
    If mIsExporting Then
        cmdStartExport.Enabled = False
        cmdCancelExport.Enabled = True
        Exit Sub
    End If
    
    CanExport = False
    
    If HasCheckedTables() Then
        If IsOutputFolderSelected() Then
            CanExport = True
        End If
    End If
    
    cmdStartExport.Enabled = CanExport
    cmdCancelExport.Enabled = False
End Sub

Private Function ValidateExportInputs() As Boolean
    On Error GoTo ErrorHandler
    
    Dim i As Long
    Dim Selection As clsExportTableSelection
    Dim HasAnyTable As Boolean
    
    ValidateExportInputs = False
    
    If gAppContext Is Nothing Then
        MsgBox "Application context is not initialized.", vbExclamation, APP_NAME
        Exit Function
    End If
    
    If Not gAppContext.IsConnected Then
        MsgBox "No active SQL connection is available.", vbExclamation, APP_NAME
        Exit Function
    End If
    
    If Len(Trim$(gAppContext.SelectedDatabase)) = 0 Then
        MsgBox "No database is selected.", vbExclamation, APP_NAME
        Exit Function
    End If
    
    If GetActiveConnection() Is Nothing Then
        MsgBox "No active database connection is available.", vbExclamation, APP_NAME
        Exit Function
    End If
    
    If Len(Trim$(txtOutputFolder.Text)) = 0 Then
        MsgBox "Please select an output folder.", vbExclamation, APP_NAME
        txtOutputFolder.SetFocus
        Exit Function
    End If
    
    If Not FolderExists(Trim$(txtOutputFolder.Text)) Then
        MsgBox "Selected output folder does not exist.", vbExclamation, APP_NAME
        txtOutputFolder.SetFocus
        Exit Function
    End If
    
    If mExportSelections Is Nothing Then
        MsgBox "No tables are loaded.", vbExclamation, APP_NAME
        Exit Function
    End If
    
    HasAnyTable = False
    
    For i = 1 To mExportSelections.Count
        Set Selection = mExportSelections(i)
        
        If Selection.IsChecked Then
            HasAnyTable = True
            
            If Not Selection.IsFieldsLoaded Then
                LoadFieldsForTableSelection Selection
            End If
            
            If Selection.SelectedFields.Count = 0 Then
                MsgBox "Please select at least one field for table:" & vbCrLf & _
                       Selection.FullTableName, vbExclamation, APP_NAME
                Exit Function
            End If
        End If
    Next i
    
    If Not HasAnyTable Then
        MsgBox "Please select at least one table to export.", vbExclamation, APP_NAME
        Exit Function
    End If
    
    ValidateExportInputs = True
    Exit Function

ErrorHandler:
    MsgBox "Failed to validate export inputs." & vbCrLf & _
           Err.Number & " - " & Err.Description, vbExclamation, APP_NAME
End Function

Private Sub LoadFieldsForTableSelection(ByVal Selection As clsExportTableSelection)
    On Error GoTo ErrorHandler
    
    Dim Columns As Collection
    Dim ColInfo As clsColumnInfo
    Dim i As Long
    
    If Selection Is Nothing Then Exit Sub
    If Selection.IsFieldsLoaded Then Exit Sub
    
    Selection.ClearFields
    
    Set Columns = mDatabaseBrowser.GetTableColumns( _
        GetActiveConnection(), _
        gAppContext.SelectedDatabase, _
        Selection.SchemaName, _
        Selection.TableName)
    
    For i = 1 To Columns.Count
        Set ColInfo = Columns(i)
        Selection.AddField ColInfo.ColumnName, True
    Next i
    
    Selection.IsFieldsLoaded = True
    Exit Sub

ErrorHandler:
    Err.Raise vbObjectError + 4101, "frmExportSqlToExcel.LoadFieldsForTableSelection", _
              "Failed to load fields for table [" & Selection.FullTableName & "]. " & Err.Description
End Sub

Private Sub txtOutputFolder_Change()
    If mSuppressUiEvents Then Exit Sub
    
    RefreshExportUiState
End Sub

Private Function BuildCsvOutputFilePath(ByVal OutputFolder As String, ByVal Selection As clsExportTableSelection) As String
    Dim FolderPath As String
    Dim FileName As String
    
    FolderPath = NormalizeFolderPath(OutputFolder)
    FileName = CleanFileName(Selection.TableName) & ".csv"
    
    BuildCsvOutputFilePath = FolderPath & FileName
End Function

Private Function NormalizeFolderPath(ByVal FolderPath As String) As String
    NormalizeFolderPath = Trim$(FolderPath)
    
    If Len(NormalizeFolderPath) = 0 Then Exit Function
    
    If Right$(NormalizeFolderPath, 1) <> "\" Then
        NormalizeFolderPath = NormalizeFolderPath & "\"
    End If
End Function

Private Function CleanFileName(ByVal FileName As String) As String
    Dim Result As String
    
    Result = Trim$(FileName)
    
    Result = Replace(Result, "\", "_")
    Result = Replace(Result, "/", "_")
    Result = Replace(Result, ":", "_")
    Result = Replace(Result, "*", "_")
    Result = Replace(Result, "?", "_")
    Result = Replace(Result, """", "_")
    Result = Replace(Result, "<", "_")
    Result = Replace(Result, ">", "_")
    Result = Replace(Result, "|", "_")
    
    CleanFileName = Result
End Function

Private Sub mnuListSelectAll_Click()
    SelectAllItemsInPopupSource
End Sub

Private Sub mnuListSelectNone_Click()
    SelectNoItemsInPopupSource
End Sub

Private Sub SelectAllItemsInPopupSource()
    On Error GoTo ErrorHandler
    
    mSuppressUiEvents = True
    
    If mPopupListSource = "TABLES" Then
        SelectAllTables
    ElseIf mPopupListSource = "FIELDS" Then
        SelectAllFields
    End If
    
    mSuppressUiEvents = False
    
    RefreshExportUiState
    Exit Sub

ErrorHandler:
    mSuppressUiEvents = False
    MsgBox "Failed to select all items." & vbCrLf & _
           Err.Number & " - " & Err.Description, vbExclamation, APP_NAME
End Sub

Private Sub SelectNoItemsInPopupSource()
    On Error GoTo ErrorHandler
    
    mSuppressUiEvents = True
    
    If mPopupListSource = "TABLES" Then
        SelectNoTables
    ElseIf mPopupListSource = "FIELDS" Then
        SelectNoFields
    End If
    
    mSuppressUiEvents = False
    
    RefreshExportUiState
    Exit Sub

ErrorHandler:
    mSuppressUiEvents = False
    MsgBox "Failed to clear selected items." & vbCrLf & _
           Err.Number & " - " & Err.Description, vbExclamation, APP_NAME
End Sub

Private Sub SelectAllTables()
    Dim i As Long
    Dim Selection As clsExportTableSelection
    
    For i = 0 To lstTables.ListCount - 1
        lstTables.Selected(i) = True
        
        Set Selection = FindTableSelection(lstTables.List(i))
        If Not Selection Is Nothing Then
            Selection.IsChecked = True
            
            If Not Selection.IsFieldsLoaded Then
                LoadFieldsForTableSelection Selection
            End If
        End If
    Next i
    
    If lstTables.ListCount > 0 Then
        If lstTables.ListIndex < 0 Then
            lstTables.ListIndex = 0
        End If
        
        LoadFieldsForSelectedTable
    End If
End Sub

Private Sub SelectNoTables()
    Dim i As Long
    Dim Selection As clsExportTableSelection
    
    For i = 0 To lstTables.ListCount - 1
        lstTables.Selected(i) = False
        
        Set Selection = FindTableSelection(lstTables.List(i))
        If Not Selection Is Nothing Then
            Selection.IsChecked = False
            Selection.ClearSelectedFields
        End If
    Next i
    
    lstFields.Clear
    lblSelectedTable.Caption = "-"
End Sub

Private Sub SelectAllFields()
    Dim i As Long
    Dim Selection As clsExportTableSelection
    
    If lstFields.ListCount = 0 Then Exit Sub
    
    Set Selection = GetCurrentTableSelection()
    If Selection Is Nothing Then Exit Sub
    
    Selection.ClearSelectedFields
    
    For i = 0 To lstFields.ListCount - 1
        lstFields.Selected(i) = True
        Selection.AddSelectedField lstFields.List(i)
    Next i
End Sub

Private Sub SelectNoFields()
    Dim i As Long
    Dim Selection As clsExportTableSelection
    
    If lstFields.ListCount = 0 Then Exit Sub
    
    Set Selection = GetCurrentTableSelection()
    If Selection Is Nothing Then Exit Sub
    
    Selection.ClearSelectedFields
    
    For i = 0 To lstFields.ListCount - 1
        lstFields.Selected(i) = False
    Next i
End Sub

Private Sub txtSearchTables_Change()
    If mSuppressUiEvents Then Exit Sub
    
    FilterTablesList
End Sub

Private Sub FilterTablesList()
    On Error GoTo ErrorHandler
    
    Dim i As Long
    Dim Selection As clsExportTableSelection
    Dim SearchText As String
    Dim CurrentFullName As String
    Dim NewIndex As Long
    
    SearchText = Trim$(txtSearchTables.Text)
    
    CurrentFullName = ""
    If lstTables.ListIndex >= 0 Then
        CurrentFullName = lstTables.List(lstTables.ListIndex)
    End If
    
    mSuppressUiEvents = True
    
    lstTables.Clear
    
    If Not mExportSelections Is Nothing Then
        For i = 1 To mExportSelections.Count
            Set Selection = mExportSelections(i)
            
            If ContainsText(Selection.FullTableName, SearchText) Then
                lstTables.AddItem Selection.FullTableName
                
                NewIndex = lstTables.ListCount - 1
                lstTables.Selected(NewIndex) = Selection.IsChecked
            End If
        Next i
    End If
    
    If lstTables.ListCount > 0 Then
        If Len(CurrentFullName) > 0 Then
            SelectTableInListIfExists CurrentFullName
        Else
            lstTables.ListIndex = 0
        End If
    Else
        lstFields.Clear
        lblSelectedTable.Caption = "-"
    End If
    
    mSuppressUiEvents = False
    
    RefreshExportUiState
    Exit Sub

ErrorHandler:
    mSuppressUiEvents = False
    MsgBox "Failed to filter table list." & vbCrLf & _
           Err.Number & " - " & Err.Description, vbExclamation, APP_NAME
End Sub

Private Function ContainsText(ByVal SourceText As String, ByVal SearchText As String) As Boolean
    If Len(Trim$(SearchText)) = 0 Then
        ContainsText = True
    Else
        ContainsText = (InStr(1, SourceText, SearchText, vbTextCompare) > 0)
    End If
End Function

Private Sub SelectTableInListIfExists(ByVal FullTableName As String)
    Dim i As Long
    
    For i = 0 To lstTables.ListCount - 1
        If StrComp(lstTables.List(i), FullTableName, vbTextCompare) = 0 Then
            lstTables.ListIndex = i
            Exit Sub
        End If
    Next i
    
    If lstTables.ListCount > 0 Then
        lstTables.ListIndex = 0
    End If
End Sub

Private Function GetCurrentTableSelection() As clsExportTableSelection
    If lstTables.ListIndex < 0 Then
        Set GetCurrentTableSelection = Nothing
        Exit Function
    End If
    
    Set GetCurrentTableSelection = FindTableSelection(lstTables.List(lstTables.ListIndex))
End Function

Private Sub AppendExportStatus(ByVal MessageText As String)
    txtStatus.Text = txtStatus.Text & _
                     Format$(Now, "hh:nn:ss") & " - " & MessageText & vbCrLf
    
    txtStatus.SelStart = Len(txtStatus.Text)
End Sub

Private Sub SetExportUiBusyState(ByVal IsBusy As Boolean)
    mIsExporting = IsBusy
    
    fraTables.Enabled = Not IsBusy
    fraFields.Enabled = Not IsBusy
    fraExportOptions.Enabled = Not IsBusy
    
    fraExecution.Enabled = True
    
    cmdStartExport.Enabled = Not IsBusy
    cmdCancelExport.Enabled = IsBusy
End Sub

Private Sub ResetExportProgress()
    lblCurrentStep.Caption = "Ready"
    '''lblProgress.Caption = "0 / 0"
    lblSuccessCount.Caption = "0"
    lblSkipCount.Caption = "0"
    lblErrorCount.Caption = "0"
    prgExport.Value = 0
    txtStatus.Text = ""
End Sub

Private Sub CheckForExportCancel()
    If mCancelRequested Then
        Err.Raise vbObjectError + 5001, "frmExportSqlToExcel.CheckForExportCancel", "Export was cancelled by user."
    End If
End Sub

Private Function BuildSelectColumnList(ByVal Selection As clsExportTableSelection) As String
    Dim i As Long
    Dim Result As String
    
    Result = ""
    
    For i = 1 To Selection.SelectedFields.Count
        If Len(Result) > 0 Then
            Result = Result & ", "
        End If
        
        Result = Result & BracketName(CStr(Selection.SelectedFields(i)))
    Next i
    
    BuildSelectColumnList = Result
End Function

Private Function BuildExportSql(ByVal Selection As clsExportTableSelection) As String
    Dim SafeDatabaseName As String
    Dim ColumnList As String
    
    SafeDatabaseName = BracketName(gAppContext.SelectedDatabase)
    ColumnList = BuildSelectColumnList(Selection)
    
    BuildExportSql = "SELECT " & ColumnList & _
                     " FROM " & SafeDatabaseName & "." & _
                     BracketName(Selection.SchemaName) & "." & _
                     BracketName(Selection.TableName)
End Function

Private Function CsvEscape(ByVal Value As Variant) As String
    Dim TextValue As String
    
    If IsNull(Value) Then
        CsvEscape = ""
        Exit Function
    End If
    
    TextValue = CStr(Value)
    
    TextValue = Replace(TextValue, """", """""")
    
    If InStr(1, TextValue, ",") > 0 _
       Or InStr(1, TextValue, vbCr) > 0 _
       Or InStr(1, TextValue, vbLf) > 0 _
       Or InStr(1, TextValue, """") > 0 Then
        
        CsvEscape = """" & TextValue & """"
    Else
        CsvEscape = TextValue
    End If
End Function

Private Function WriteRecordsetToCsv( _
    ByVal Rs As ADODB.Recordset, _
    ByVal FilePath As String, _
    ByVal Selection As clsExportTableSelection) As Long
    
    On Error GoTo ErrorHandler
    
    Dim FileNo As Integer
    Dim i As Long
    Dim LineText As String
    Dim RowCount As Long
    
    FileNo = FreeFile
    
    Open FilePath For Output As #FileNo
    
    ' Header row
    LineText = ""
    For i = 1 To Selection.SelectedFields.Count
        If Len(LineText) > 0 Then LineText = LineText & ","
        LineText = LineText & CsvEscape(CStr(Selection.SelectedFields(i)))
    Next i
    
    Print #FileNo, LineText
    
    RowCount = 0
    
    Do While Not Rs.EOF
        CheckForExportCancel
        
        LineText = ""
        
        For i = 1 To Selection.SelectedFields.Count
            If Len(LineText) > 0 Then LineText = LineText & ","
            LineText = LineText & CsvEscape(Rs.Fields(CStr(Selection.SelectedFields(i))).Value)
        Next i
        
        Print #FileNo, LineText
        
        RowCount = RowCount + 1
        
        If (RowCount Mod 5) = 0 Then
            DoEvents
        End If
        
        Rs.MoveNext
    Loop
    
    Close #FileNo
    
    WriteRecordsetToCsv = RowCount
    Exit Function

ErrorHandler:
    On Error Resume Next
    If FileNo > 0 Then Close #FileNo
    
    Err.Raise vbObjectError + 5101, "frmExportSqlToExcel.WriteRecordsetToCsv", _
              "Failed to write CSV file [" & FilePath & "]. " & Err.Description
End Function

Private Function ExportTableToCsv(ByVal Selection As clsExportTableSelection, ByVal OutputFolder As String) As Long
    On Error GoTo ErrorHandler
    
    Dim Conn As ADODB.Connection
    Dim Rs As ADODB.Recordset
    Dim SqlText As String
    Dim FilePath As String
    
    Set Conn = GetActiveConnection()
    
    If Conn Is Nothing Then
        Err.Raise vbObjectError + 5102, "frmExportSqlToExcel.ExportTableToCsv", "No active database connection is available."
    End If
    
    FilePath = BuildCsvOutputFilePath(OutputFolder, Selection)
    
    If FileExists(FilePath) Then
        Select Case MsgBox("File already exists:" & vbCrLf & FilePath & vbCrLf & vbCrLf & _
                           "Do you want to overwrite it?", _
                           vbQuestion + vbYesNoCancel, APP_NAME)
            Case vbYes
                Kill FilePath
            
            Case vbNo
                ExportTableToCsv = -1
                Exit Function
            
            Case vbCancel
                Err.Raise vbObjectError + 5001, "frmExportSqlToExcel.ExportTableToCsv", "Export was cancelled by user."
        End Select
    End If
    
    SqlText = BuildExportSql(Selection)
    
    Set Rs = New ADODB.Recordset
    Rs.Open SqlText, Conn, adOpenForwardOnly, adLockReadOnly
    
    ExportTableToCsv = WriteRecordsetToCsv(Rs, FilePath, Selection)
    
    Rs.Close
    Set Rs = Nothing
    Exit Function

ErrorHandler:
    On Error Resume Next
    
    If Not Rs Is Nothing Then
        If Rs.State = adStateOpen Then Rs.Close
    End If
    Set Rs = Nothing
    
    Err.Raise Err.Number, "frmExportSqlToExcel.ExportTableToCsv", Err.Description
End Function

Private Function GetCheckedTableCount() As Long
    Dim i As Long
    Dim Selection As clsExportTableSelection
    
    If mExportSelections Is Nothing Then Exit Function
    
    For i = 1 To mExportSelections.Count
        Set Selection = mExportSelections(i)
        If Selection.IsChecked Then
            GetCheckedTableCount = GetCheckedTableCount + 1
        End If
    Next i
End Function

Private Sub MoveFieldCursorToNextItem(ByVal CurrentIndex As Integer)
    On Error Resume Next
    
    If lstFields.ListCount = 0 Then Exit Sub
    If CurrentIndex < 0 Then Exit Sub
    
    If CurrentIndex < lstFields.ListCount - 1 Then
        lstFields.ListIndex = CurrentIndex + 1
    Else
        lstFields.ListIndex = CurrentIndex
    End If
End Sub

Private Function IsExcelInstalled() As Boolean
    On Error GoTo ErrorHandler
    
    Dim ExcelApp As Object
    
    Set ExcelApp = CreateObject("Excel.Application")
    
    IsExcelInstalled = True
    
    ExcelApp.Quit
    Set ExcelApp = Nothing
    Exit Function

ErrorHandler:
    IsExcelInstalled = False
    
    On Error Resume Next
    If Not ExcelApp Is Nothing Then ExcelApp.Quit
    Set ExcelApp = Nothing
End Function

Private Function BuildExcelOutputFilePath(ByVal OutputFolder As String, ByVal Selection As clsExportTableSelection) As String
    Dim FolderPath As String
    Dim FileName As String
    
    FolderPath = NormalizeFolderPath(OutputFolder)
    FileName = CleanFileName(Selection.SchemaName & "." & Selection.TableName) & ".xlsx"
    
    BuildExcelOutputFilePath = FolderPath & FileName
End Function

Private Function CanCreateXlsWithAdo() As Boolean
    On Error GoTo ErrorHandler
    
    Dim TestFolder As String
    Dim TestFilePath As String
    Dim Conn As ADODB.Connection
    Dim ConnStr As String
    
    TestFolder = NormalizeFolderPath(App.Path)
    TestFilePath = TestFolder & "__xls_export_test_" & Format$(Now, "yyyymmdd_hhnnss") & ".xls"
    
    ConnStr = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
              "Data Source=" & TestFilePath & ";" & _
              "Extended Properties=""Excel 8.0;HDR=YES"";"
    
    Set Conn = New ADODB.Connection
    Conn.Open ConnStr
    
    Conn.Execute "CREATE TABLE [TestSheet] ([TestValue] TEXT)"
    Conn.Execute "INSERT INTO [TestSheet] ([TestValue]) VALUES ('OK')"
    
    Conn.Close
    Set Conn = Nothing
    
    On Error Resume Next
    Kill TestFilePath
    
    CanCreateXlsWithAdo = True
    Exit Function

ErrorHandler:
    CanCreateXlsWithAdo = False
    
    On Error Resume Next
    If Not Conn Is Nothing Then
        If Conn.State = adStateOpen Then Conn.Close
    End If
    Set Conn = Nothing
    
    Kill TestFilePath
End Function

Private Function BuildXlsOutputFilePath(ByVal OutputFolder As String, ByVal Selection As clsExportTableSelection) As String
    Dim FolderPath As String
    Dim FileName As String
    
    FolderPath = NormalizeFolderPath(OutputFolder)
    FileName = CleanFileName(Selection.TableName) & ".xls"
    
    BuildXlsOutputFilePath = FolderPath & FileName
End Function

Private Function BuildXlsxOutputFilePath(ByVal OutputFolder As String, ByVal Selection As clsExportTableSelection) As String
    Dim FolderPath As String
    Dim FileName As String
    
    FolderPath = NormalizeFolderPath(OutputFolder)
    FileName = CleanFileName(Selection.TableName) & ".xlsx"
    
    BuildXlsxOutputFilePath = FolderPath & FileName
End Function

Private Function GetExcelExportMode() As String
    If IsExcelInstalled() Then
        GetExcelExportMode = "XLSX_AUTOMATION"
    ElseIf CanCreateXlsWithAdo() Then
        GetExcelExportMode = "XLS_ADO"
    Else
        GetExcelExportMode = ""
    End If
End Function

Private Function ValidateExcelExportAvailability(ByRef ExcelExportMode As String) As Boolean
    ValidateExcelExportAvailability = False
    ExcelExportMode = ""
    
    If Not optExportExcel.Value Then
        ValidateExcelExportAvailability = True
        Exit Function
    End If
    
    ExcelExportMode = GetExcelExportMode()
    
    If Len(ExcelExportMode) = 0 Then
        MsgBox "Excel export is not available on this computer." & vbCrLf & vbCrLf & _
               "Microsoft Excel is not installed and the required OLEDB provider for XLS export is not available." & vbCrLf & vbCrLf & _
               "Please select CSV export instead.", _
               vbExclamation, APP_NAME
        Exit Function
    End If
    
    ValidateExcelExportAvailability = True
End Function

Private Function CleanSheetName(ByVal SheetName As String) As String
    Dim Result As String
    
    Result = Trim$(SheetName)
    
    Result = Replace(Result, "\", "_")
    Result = Replace(Result, "/", "_")
    Result = Replace(Result, ":", "_")
    Result = Replace(Result, "*", "_")
    Result = Replace(Result, "?", "_")
    Result = Replace(Result, "[", "_")
    Result = Replace(Result, "]", "_")
    
    If Len(Result) = 0 Then Result = "Sheet1"
    
    If Len(Result) > 31 Then
        Result = Left$(Result, 31)
    End If
    
    CleanSheetName = Result
End Function

Private Function CleanExcelColumnName(ByVal ColumnName As String) As String
    Dim Result As String
    
    Result = Trim$(ColumnName)
    
    Result = Replace(Result, "[", "_")
    Result = Replace(Result, "]", "_")
    Result = Replace(Result, ".", "_")
    Result = Replace(Result, "'", "_")
    Result = Replace(Result, """", "_")
    
    If Len(Result) = 0 Then Result = "Column1"
    
    CleanExcelColumnName = Result
End Function

Private Function AdoXlsValueLiteral(ByVal Value As Variant) As String
    Dim TextValue As String
    
    If IsNull(Value) Then
        AdoXlsValueLiteral = "NULL"
        Exit Function
    End If
    
    TextValue = CStr(Value)
    TextValue = Replace(TextValue, "'", "''")
    
    AdoXlsValueLiteral = "'" & TextValue & "'"
End Function

Private Function BuildXlsCreateTableSql(ByVal SheetName As String, ByVal Selection As clsExportTableSelection) As String
    Dim i As Long
    Dim SqlText As String
    Dim ColumnDefs As String
    Dim FieldName As String
    
    ColumnDefs = ""
    
    For i = 1 To Selection.SelectedFields.Count
        FieldName = CleanExcelColumnName(CStr(Selection.SelectedFields(i)))
        
        If Len(ColumnDefs) > 0 Then
            ColumnDefs = ColumnDefs & ", "
        End If
        
        ColumnDefs = ColumnDefs & "[" & FieldName & "] TEXT"
    Next i
    
    SqlText = "CREATE TABLE [" & SheetName & "] (" & ColumnDefs & ")"
    
    BuildXlsCreateTableSql = SqlText
End Function

Private Function BuildXlsInsertSql( _
    ByVal SheetName As String, _
    ByVal Selection As clsExportTableSelection, _
    ByVal Rs As ADODB.Recordset) As String
    
    Dim i As Long
    Dim ColumnList As String
    Dim ValueList As String
    Dim FieldName As String
    Dim CleanName As String
    
    ColumnList = ""
    ValueList = ""
    
    For i = 1 To Selection.SelectedFields.Count
        FieldName = CStr(Selection.SelectedFields(i))
        CleanName = CleanExcelColumnName(FieldName)
        
        If Len(ColumnList) > 0 Then
            ColumnList = ColumnList & ", "
            ValueList = ValueList & ", "
        End If
        
        ColumnList = ColumnList & "[" & CleanName & "]"
        ValueList = ValueList & AdoXlsValueLiteral(Rs.Fields(FieldName).Value)
    Next i
    
    BuildXlsInsertSql = "INSERT INTO [" & SheetName & "] (" & ColumnList & ") VALUES (" & ValueList & ")"
End Function

Private Function WriteRecordsetToXlsWithAdo( _
    ByVal Rs As ADODB.Recordset, _
    ByVal FilePath As String, _
    ByVal Selection As clsExportTableSelection) As Long
    
    On Error GoTo ErrorHandler
    
    Dim XlsConn As ADODB.Connection
    Dim ConnStr As String
    Dim SheetName As String
    Dim RowCount As Long
    
    SheetName = CleanSheetName(Selection.TableName)
    
    ConnStr = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
              "Data Source=" & FilePath & ";" & _
              "Extended Properties=""Excel 8.0;HDR=YES"";"
    
    Set XlsConn = New ADODB.Connection
    XlsConn.Open ConnStr
    
    XlsConn.Execute BuildXlsCreateTableSql(SheetName, Selection)
    
    RowCount = 0
    
    Do While Not Rs.EOF
        CheckForExportCancel
        
        XlsConn.Execute BuildXlsInsertSql(SheetName, Selection, Rs)
        
        RowCount = RowCount + 1
        
        If (RowCount Mod 50) = 0 Then
            DoEvents
        End If
        
        Rs.MoveNext
    Loop
    
    XlsConn.Close
    Set XlsConn = Nothing
    
    WriteRecordsetToXlsWithAdo = RowCount
    Exit Function

ErrorHandler:
    On Error Resume Next
    
    If Not XlsConn Is Nothing Then
        If XlsConn.State = adStateOpen Then XlsConn.Close
    End If
    Set XlsConn = Nothing
    
    Err.Raise vbObjectError + 5201, "frmExportSqlToExcel.WriteRecordsetToXlsWithAdo", _
              "Failed to write XLS file [" & FilePath & "]. " & Err.Description
End Function

Private Function ExportTableToXlsWithAdo(ByVal Selection As clsExportTableSelection, ByVal OutputFolder As String) As Long
    On Error GoTo ErrorHandler
    
    Dim Conn As ADODB.Connection
    Dim Rs As ADODB.Recordset
    Dim SqlText As String
    Dim FilePath As String
    
    Set Conn = GetActiveConnection()
    
    If Conn Is Nothing Then
        Err.Raise vbObjectError + 5202, "frmExportSqlToExcel.ExportTableToXlsWithAdo", "No active database connection is available."
    End If
    
    FilePath = BuildXlsOutputFilePath(OutputFolder, Selection)
    
    If FileExists(FilePath) Then
        Select Case MsgBox("File already exists:" & vbCrLf & FilePath & vbCrLf & vbCrLf & _
                           "Do you want to overwrite it?", _
                           vbQuestion + vbYesNoCancel, APP_NAME)
            Case vbYes
                Kill FilePath
            
            Case vbNo
                ExportTableToXlsWithAdo = -1
                Exit Function
            
            Case vbCancel
                Err.Raise vbObjectError + 5001, "frmExportSqlToExcel.ExportTableToXlsWithAdo", "Export was cancelled by user."
        End Select
    End If
    
    SqlText = BuildExportSql(Selection)
    
    Set Rs = New ADODB.Recordset
    Rs.Open SqlText, Conn, adOpenForwardOnly, adLockReadOnly
    
    ExportTableToXlsWithAdo = WriteRecordsetToXlsWithAdo(Rs, FilePath, Selection)
    
    Rs.Close
    Set Rs = Nothing
    Exit Function

ErrorHandler:
    On Error Resume Next
    
    If Not Rs Is Nothing Then
        If Rs.State = adStateOpen Then Rs.Close
    End If
    Set Rs = Nothing
    
    Err.Raise Err.Number, "frmExportSqlToExcel.ExportTableToXlsWithAdo", Err.Description
End Function

Private Function WriteRecordsetToXlsxWithAutomation( _
    ByVal Rs As ADODB.Recordset, _
    ByVal FilePath As String, _
    ByVal Selection As clsExportTableSelection, _
    ByVal TotalRows As Long) As Long
    
    On Error GoTo ErrorHandler
    
    Dim ExcelApp As Object
    Dim Workbook As Object
    Dim Worksheet As Object
    Dim RowIndex As Long
    Dim ColIndex As Long
    Dim FieldName As String
    Dim RowCount As Long
    Dim SheetName As String
    
    SheetName = CleanSheetName(Selection.TableName)
    
    Set ExcelApp = CreateObject("Excel.Application")
'    ExcelApp.Visible = False
'    ExcelApp.DisplayAlerts = False
    
    Set Workbook = ExcelApp.Workbooks.Add
    Set Worksheet = Workbook.Worksheets(1)
    
    Worksheet.Name = SheetName
    
    ' Header row
    For ColIndex = 1 To Selection.SelectedFields.Count
        FieldName = CStr(Selection.SelectedFields(ColIndex))
        Worksheet.Cells(1, ColIndex).Value = FieldName
        Worksheet.Cells(1, ColIndex).Font.Bold = True
    Next ColIndex
    
    RowIndex = 2
    RowCount = 0
    
    UpdateExportRowProgress 0, TotalRows, "Exporting rows to Excel..."
        
    Do While Not Rs.EOF
        DoEvents
        CheckForExportCancel
        
        For ColIndex = 1 To Selection.SelectedFields.Count
            FieldName = CStr(Selection.SelectedFields(ColIndex))
            
            If IsNull(Rs.Fields(FieldName).Value) Then
                Worksheet.Cells(RowIndex, ColIndex).Value = ""
            Else
                Worksheet.Cells(RowIndex, ColIndex).Value = Rs.Fields(FieldName).Value
            End If
        Next ColIndex
        
        RowIndex = RowIndex + 1
        RowCount = RowCount + 1
        
        If (RowCount Mod 5) = 0 Or RowCount = TotalRows Then
            UpdateExportRowProgress RowCount, TotalRows, _
                "Exporting " & Selection.FullTableName & " to Excel..."
        End If
        
        Rs.MoveNext
    Loop
    
    Worksheet.Columns.AutoFit
    
    Workbook.SaveAs FilePath, 51   ' 51 = xlOpenXMLWorkbook (.xlsx)
    Workbook.Close False
    
    ExcelApp.Quit
    
    Set Worksheet = Nothing
    Set Workbook = Nothing
    Set ExcelApp = Nothing
    
    WriteRecordsetToXlsxWithAutomation = RowCount
    Exit Function

ErrorHandler:
    On Error Resume Next
    
    If Not Workbook Is Nothing Then Workbook.Close False
    If Not ExcelApp Is Nothing Then ExcelApp.Quit
    
    Set Worksheet = Nothing
    Set Workbook = Nothing
    Set ExcelApp = Nothing
    
    Err.Raise vbObjectError + 5301, "frmExportSqlToExcel.WriteRecordsetToXlsxWithAutomation", _
              "Failed to write XLSX file [" & FilePath & "]. " & Err.Description
End Function

Private Function ExportTableToXlsxWithAutomation(ByVal Selection As clsExportTableSelection, ByVal OutputFolder As String) As Long
    On Error GoTo ErrorHandler
    
    Dim Conn As ADODB.Connection
    Dim Rs As ADODB.Recordset
    Dim SqlText As String
    Dim FilePath As String
    Dim TotalRows As Long
    
    Set Conn = GetActiveConnection()
    
    If Conn Is Nothing Then
        Err.Raise vbObjectError + 5302, "frmExportSqlToExcel.ExportTableToXlsxWithAutomation", _
                  "No active database connection is available."
    End If
    
    FilePath = BuildXlsxOutputFilePath(OutputFolder, Selection)
    
    If FileExists(FilePath) Then
        Select Case MsgBox("File already exists:" & vbCrLf & FilePath & vbCrLf & vbCrLf & _
                           "Do you want to overwrite it?", _
                           vbQuestion + vbYesNoCancel, APP_NAME)
            Case vbYes
                Kill FilePath
            
            Case vbNo
                ExportTableToXlsxWithAutomation = -1
                Exit Function
            
            Case vbCancel
                Err.Raise vbObjectError + 5001, _
                          "frmExportSqlToExcel.ExportTableToXlsxWithAutomation", _
                          "Export was cancelled by user."
        End Select
    End If
    
    TotalRows = GetExportTableRowCount(Selection)
    
    SqlText = BuildExportSql(Selection)
    
    Set Rs = New ADODB.Recordset
    Rs.Open SqlText, Conn, adOpenForwardOnly, adLockReadOnly
    
    ExportTableToXlsxWithAutomation = WriteRecordsetToXlsxWithAutomation( _
        Rs, _
        FilePath, _
        Selection, _
        TotalRows)
    
    Rs.Close
    Set Rs = Nothing
    Exit Function

ErrorHandler:
    Dim OriginalErrNumber As Long
    Dim OriginalErrDescription As String
    
    OriginalErrNumber = Err.Number
    OriginalErrDescription = Err.Description
    
    On Error Resume Next
    
    If Not Rs Is Nothing Then
        If Rs.State = adStateOpen Then Rs.Close
    End If
    Set Rs = Nothing
    
    If OriginalErrNumber = vbObjectError + 5001 Then
        If Len(FilePath) > 0 Then
            If FileExists(FilePath) Then Kill FilePath
        End If
    End If
    
    Err.Clear
    Err.Raise OriginalErrNumber, "frmExportSqlToExcel.ExportTableToXlsxWithAutomation", OriginalErrDescription
End Function

Private Sub UpdateExportRowProgress( _
    ByVal CurrentRow As Long, _
    ByVal TotalRows As Long, _
    ByVal StepText As String)
    
    On Error Resume Next
    
    lblCurrentStep.Caption = StepText
    'lblProgress.Caption = CStr(CurrentRow) & " / " & CStr(TotalRows)
    
    If TotalRows > 0 Then
        prgExport.Value = CLng((CurrentRow / TotalRows) * PROGRESS_MAX)
    Else
        prgExport.Value = 0
    End If
    
    If (CurrentRow Mod 25) = 0 Or CurrentRow = TotalRows Then
        DoEvents
    End If
End Sub

Private Function GetExportTableRowCount(ByVal Selection As clsExportTableSelection) As Long
    On Error GoTo ErrorHandler
    
    Dim Conn As ADODB.Connection
    Dim Rs As ADODB.Recordset
    Dim SqlText As String
    Dim SafeDatabaseName As String
    
    Set Conn = GetActiveConnection()
    
    If Conn Is Nothing Then
        GetExportTableRowCount = 0
        Exit Function
    End If
    
    SafeDatabaseName = BracketName(gAppContext.SelectedDatabase)
    
    SqlText = "SELECT COUNT(*) FROM " & SafeDatabaseName & "." & _
              BracketName(Selection.SchemaName) & "." & _
              BracketName(Selection.TableName)
    
    Set Rs = New ADODB.Recordset
    Rs.Open SqlText, Conn, adOpenForwardOnly, adLockReadOnly
    
    If Not Rs.EOF Then
        GetExportTableRowCount = SafeCLng(Rs.Fields(0).Value, 0)
    Else
        GetExportTableRowCount = 0
    End If
    
    Rs.Close
    Set Rs = Nothing
    Exit Function

ErrorHandler:
    On Error Resume Next
    If Not Rs Is Nothing Then
        If Rs.State = adStateOpen Then Rs.Close
    End If
    Set Rs = Nothing
    
    GetExportTableRowCount = 0
End Function

