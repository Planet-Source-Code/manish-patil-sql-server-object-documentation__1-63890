VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDocumentation 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Documentation 1.0"
   ClientHeight    =   8715
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11100
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8715
   ScaleWidth      =   11100
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraFiles 
      Height          =   6255
      Left            =   150
      TabIndex        =   10
      Top             =   2280
      Width           =   10785
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   285
         Left            =   1230
         TabIndex        =   13
         Top             =   5820
         Width           =   7695
         _ExtentX        =   13573
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.CommandButton cmdDocument 
         Caption         =   "Create &Document"
         Height          =   375
         Left            =   9060
         TabIndex        =   14
         Top             =   5730
         Width           =   1605
      End
      Begin VB.CheckBox chkSelect 
         Caption         =   "&Select All"
         Height          =   285
         Left            =   180
         TabIndex        =   12
         Top             =   5820
         Width           =   1035
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   5415
         Left            =   180
         TabIndex        =   11
         Top             =   240
         Width           =   10485
         _ExtentX        =   18494
         _ExtentY        =   9551
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
   End
   Begin VB.Frame fraOptions 
      Height          =   2025
      Left            =   150
      TabIndex        =   9
      Top             =   120
      Width           =   10785
      Begin VB.CommandButton cmdAbout 
         Caption         =   "&About"
         Height          =   345
         Left            =   5430
         TabIndex        =   3
         Top             =   270
         Width           =   1065
      End
      Begin VB.OptionButton optObjectOptions 
         Caption         =   "&2. Stored Procedure"
         Height          =   285
         Index           =   1
         Left            =   150
         TabIndex        =   8
         Top             =   1590
         Width           =   1845
      End
      Begin VB.OptionButton optObjectOptions 
         Caption         =   "&1. Tables"
         Height          =   285
         Index           =   0
         Left            =   150
         TabIndex        =   7
         Top             =   1170
         Width           =   1065
      End
      Begin VB.CommandButton cmdPath 
         Caption         =   "..."
         Height          =   315
         Left            =   10290
         TabIndex        =   6
         Top             =   750
         Width           =   345
      End
      Begin VB.TextBox txtPath 
         Height          =   315
         Left            =   990
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   750
         Width           =   9195
      End
      Begin VB.ComboBox cboDatabases 
         Height          =   315
         Left            =   990
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   270
         Width           =   3075
      End
      Begin VB.CommandButton cmdConnect 
         Caption         =   "&Connect"
         Height          =   345
         Left            =   4230
         TabIndex        =   2
         Top             =   270
         Width           =   1065
      End
      Begin VB.Label lblOptions 
         AutoSize        =   -1  'True
         Caption         =   "&Path"
         Height          =   195
         Index           =   1
         Left            =   150
         TabIndex        =   4
         Top             =   810
         Width           =   330
      End
      Begin VB.Label lblOptions 
         AutoSize        =   -1  'True
         Caption         =   "&Database"
         Height          =   195
         Index           =   0
         Left            =   150
         TabIndex        =   0
         Top             =   330
         Width           =   690
      End
   End
End
Attribute VB_Name = "frmDocumentation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private objFSO As FileSystemObject

Private Sub Form_Load()
    Dim oDB As SQLDMO.Database2
    Dim strLastOpenedDB As String
    Dim strSelectedPath As String
    
    On Error GoTo errHand
    
    Me.Show
    Me.Caption = APPLICATION_NAME
    
    Set objFSO = New FileSystemObject
    For Each oDB In objServer.Databases
        If oDB.SystemObject = False Then
            cboDatabases.AddItem oDB.Name
        End If
    Next
    
    With ListView1
        .Checkboxes = True
        .ColumnHeaders.Add 1, , "Table Name": .ColumnHeaders.Item(1).Width = 2500
        .ColumnHeaders.Add 2, , "File Name": .ColumnHeaders.Item(2).Width = .Width - 2600
        .View = lvwReport
        .FullRowSelect = True
        .LabelEdit = lvwManual
    End With
    
    'START - Loading Data From The INI File.
    If ReadFromInIFile("AppData", "LastOpenedDB", strLastOpenedDB, App.Path & "\Config.ini") = True Then
        If Len(Trim(strLastOpenedDB)) > 0 Then
            cboDatabases.Text = strLastOpenedDB
            If ReadFromInIFile(strLastOpenedDB, "SelectedPath", strSelectedPath, App.Path & "\Config.ini") = True Then
                If Len(Trim(strSelectedPath)) > 0 Then
                    txtPath.Text = strSelectedPath
                Else
                    txtPath.Text = ""
                End If
            Else
                txtPath.Text = ""
            End If
        Else
            cboDatabases.ListIndex = -1
        End If
    Else
        cboDatabases.ListIndex = -1
    End If
    'END - Loading Data From The INI File.
        
    ProgressBar1.Visible = False
    
    Exit Sub
errHand:
    If Err.Number = 383 Then Resume Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If cboDatabases.ListIndex >= 0 Then
        WriteToInIFile "AppData", "LastOpenedDB", cboDatabases.Text, App.Path & "\Config.ini"
        If Len(Trim(txtPath)) > 0 Then
            WriteToInIFile cboDatabases.Text, "SelectedPath", txtPath.Text, App.Path & "\Config.ini"
        End If
    End If
    
    Set objFSO = Nothing
End Sub

Private Sub cboDatabases_Click()
    Dim strSelectedPath As String
    
    If cboDatabases.ListIndex >= 0 Then
        If ReadFromInIFile(cboDatabases.Text, "SelectedPath", strSelectedPath, App.Path & "\Config.ini") = True Then
            If Len(Trim(strSelectedPath)) > 0 Then
                txtPath.Text = strSelectedPath
            Else
                txtPath.Text = ""
            End If
        Else
            txtPath.Text = ""
        End If
        cmdConnect_Click
    End If
End Sub

Private Sub cmdConnect_Click()
    If cboDatabases.ListIndex >= 0 Then
        Set objDatabase = objServer.Databases(cboDatabases.Text)
        If optObjectOptions(1).Value = True Then
            optObjectOptions(1).Value = True
            optObjectOptions_Click 1
        Else
            optObjectOptions(0).Value = True
            optObjectOptions_Click 0
        End If
    Else
        MsgBox "Please Select the Database to Conenct.", vbCritical, APPLICATION_NAME
        cboDatabases.SetFocus
    End If
End Sub

Private Sub cmdAbout_Click()
    frmAbout.Show vbModal
End Sub

Private Sub cmdPath_Click()
    Dim strPath As String
    
    On Error GoTo errHand
    
    If cboDatabases.ListIndex < 0 Then
        MsgBox "Please Select the Database to Conenct.", vbCritical, APPLICATION_NAME
        Exit Sub
    End If
    
    strPath = SelectFolder(Me.hWnd)
    txtPath.Text = strPath
    
    If objFSO.FolderExists(strPath & "\" & cboDatabases.Text) = False Then
        objFSO.CreateFolder strPath & "\" & cboDatabases.Text
    End If
    If objFSO.FolderExists(strPath & "\" & cboDatabases.Text & "\Tables") = False Then
        objFSO.CreateFolder strPath & "\" & cboDatabases.Text & "\Tables"
    End If
    If objFSO.FolderExists(strPath & "\" & cboDatabases.Text & "\Stored Procedures") = False Then
        objFSO.CreateFolder strPath & "\" & cboDatabases.Text & "\Stored Procedures"
    End If
    
    If cboDatabases.ListIndex >= 0 Then
        WriteToInIFile "AppData", "LastOpenedDB", cboDatabases.Text, App.Path & "\Config.ini"
        If Len(Trim(txtPath)) > 0 Then
            WriteToInIFile cboDatabases.Text, "SelectedPath", txtPath.Text, App.Path & "\Config.ini"
        End If
    End If
    
    Exit Sub
errHand:
    MsgBox Err.Description, vbCritical, APPLICATION_NAME
End Sub

Private Sub optObjectOptions_Click(Index As Integer)
    Dim itmX As ListItem

    Select Case Index
        Case 0 'Tables
            Dim objTables As New SQLDMO.Table
            
            ListView1.ListItems.Clear
            For Each objTables In objDatabase.Tables
                If objTables.SystemObject = False Then
                    Set itmX = ListView1.ListItems.Add(, , objTables.Name)
                    If objFSO.FolderExists(txtPath.Text) = True Then
                        If objFSO.FileExists(txtPath.Text & "\" & cboDatabases.Text & "\Tables\" & objTables.Name & ".doc") = True Then
                            itmX.SubItems(1) = txtPath.Text & "\" & cboDatabases.Text & "\Tables\" & objTables.Name & ".doc"
                        Else
                            itmX.SubItems(1) = "<File Does Not Exist>"
                        End If
                    Else
                        itmX.SubItems(1) = "<File Does Not Exist>"
                    End If
                End If
            Next
            
            Set objTables = Nothing
        Case 1 'Stored Procedure
            Dim objStoredProcedure As New SQLDMO.StoredProcedure2
            
            ListView1.ListItems.Clear
            For Each objStoredProcedure In objDatabase.StoredProcedures
                If objStoredProcedure.SystemObject = False Then
                    Set itmX = ListView1.ListItems.Add(, , objStoredProcedure.Name)
                    If objFSO.FolderExists(txtPath.Text) = True Then
                        If objFSO.FileExists(txtPath.Text & "\" & cboDatabases.Text & "\Stored Procedures\" & objStoredProcedure.Name & ".doc") = True Then
                            itmX.SubItems(1) = txtPath.Text & "\" & cboDatabases.Text & "\Stored Procedures\" & objStoredProcedure.Name & ".doc"
                        Else
                            itmX.SubItems(1) = "<File Does Not Exist>"
                        End If
                    Else
                        itmX.SubItems(1) = "<File Does Not Exist>"
                    End If
                End If
            Next
            
            Set objStoredProcedure = Nothing
    End Select
    If chkSelect.Value = 1 Then chkSelect.Value = 0
End Sub

Private Sub ListView1_DblClick()
    With ListView1
        If Len(Trim(.ListItems.Item(.SelectedItem.Index).SubItems(1))) > 0 Then
            If Trim(.ListItems.Item(.SelectedItem.Index).SubItems(1)) <> "<File Does Not Exist>" Then
                If objFSO.FileExists(Trim(.ListItems.Item(.SelectedItem.Index).SubItems(1))) = True Then
                    ShowFile Me.hWnd, Trim(.ListItems.Item(.SelectedItem.Index).SubItems(1))
                Else
                    MsgBox "File Deleted Or Moved.", vbCritical, APPLICATION_NAME
                End If
            End If
        End If
    End With
End Sub

Private Sub chkSelect_Click()
    Dim intCounter As Integer
        
    For intCounter = 1 To ListView1.ListItems.Count
        If chkSelect.Value = 1 Then
            ListView1.ListItems.Item(intCounter).Checked = True
        Else
            ListView1.ListItems.Item(intCounter).Checked = False
        End If
    Next
End Sub

Private Sub cmdDocument_Click()
    Dim intCounter As Integer
    Dim intSelected As Integer
    
    On Error GoTo errHand
    
    If objFSO.FolderExists(Trim(txtPath.Text)) = True Then
        If objFSO.FolderExists(Trim(txtPath.Text) & "\" & cboDatabases.Text) = False Then
            objFSO.CreateFolder Trim(txtPath.Text) & "\" & cboDatabases.Text
        End If
        If objFSO.FolderExists(Trim(txtPath.Text) & "\" & cboDatabases.Text & "\Tables") = False Then
            objFSO.CreateFolder Trim(txtPath.Text) & "\" & cboDatabases.Text & "\Tables"
        End If
        If objFSO.FolderExists(Trim(txtPath.Text) & "\" & cboDatabases.Text & "\Stored Procedures") = False Then
            objFSO.CreateFolder Trim(txtPath.Text) & "\" & cboDatabases.Text & "\Stored Procedures"
        End If
    Else
        MsgBox "Please Specify The Correct Path", vbCritical, APPLICATION_NAME
        txtPath.SetFocus
        Exit Sub
    End If
    
    For intCounter = 1 To ListView1.ListItems.Count
        If ListView1.ListItems.Item(intCounter).Checked = True Then
            intSelected = intSelected + 1
        End If
    Next
    
    If intSelected <= 0 Then
        MsgBox "Please Select Objects To Document.", vbCritical, APPLICATION_NAME
        ListView1.SetFocus
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    DisableControls
    
    ProgressBar1.Min = 0: ProgressBar1.Max = (intSelected * 10) + 1: ProgressBar1.Value = 0
    If ProgressBar1.Max > 1 Then ProgressBar1.Visible = True: Progress ProgressBar1, 1 Else ProgressBar1.Visible = False
    
    For intCounter = 1 To ListView1.ListItems.Count
        If ListView1.ListItems.Item(intCounter).Checked = True Then
            If optObjectOptions(0).Value = True Then 'Tables
                Dim objTable As SQLDMO.Table2
                
                Set objTable = objDatabase.Tables(ListView1.ListItems.Item(intCounter).Text)
                If UCase(TypeName(objTable)) <> "NOTHING" Then
                    CreateTableDocument objTable, Trim(txtPath.Text)
                    If objFSO.FileExists(txtPath.Text & "\" & cboDatabases.Text & "\Tables\" & ListView1.ListItems.Item(intCounter).Text & ".doc") = True Then
                        ListView1.ListItems.Item(intCounter).SubItems(1) = txtPath.Text & "\" & cboDatabases.Text & "\Tables\" & ListView1.ListItems.Item(intCounter).Text & ".doc"
                    End If
                    ListView1.ListItems.Item(intCounter).Checked = False
                    Set objTable = Nothing
                End If
                Progress ProgressBar1, ProgressBar1.Value + 10
            ElseIf optObjectOptions(1).Value = True Then 'Stored Procedure
                Dim objStoredProcedure As SQLDMO.StoredProcedure2
                
                Set objStoredProcedure = objDatabase.StoredProcedures(ListView1.ListItems.Item(intCounter).Text)
                If UCase(TypeName(objStoredProcedure)) <> "NOTHING" Then
                    CreateSPDocument objStoredProcedure, Trim(txtPath.Text)
                    If objFSO.FileExists(txtPath.Text & "\" & cboDatabases.Text & "\Stored Procedures\" & ListView1.ListItems.Item(intCounter).Text & ".doc") = True Then
                        ListView1.ListItems.Item(intCounter).SubItems(1) = txtPath.Text & "\" & cboDatabases.Text & "\Stored Procedures\" & ListView1.ListItems.Item(intCounter).Text & ".doc"
                    End If
                    ListView1.ListItems.Item(intCounter).Checked = False
                    Set objStoredProcedure = Nothing
                End If
                Progress ProgressBar1, ProgressBar1.Value + 10
            End If
        End If
    Next
    
    If chkSelect.Value = 1 Then chkSelect.Value = 0
    EnableControls
    Screen.MousePointer = vbNormal
    MsgBox "Documentation Compleated.", vbInformation, APPLICATION_NAME
    
    Exit Sub
errHand:
    If Err.Number = -2147199728 Then Resume Next
    If chkSelect.Value = 1 Then chkSelect.Value = 0
    EnableControls
    Screen.MousePointer = vbNormal
End Sub

Private Function CreateTableDocument(objDMO_Table As SQLDMO.Table2, ByVal strPath As String)
    Dim objWord As Word.Application
    Dim objDocument As Word.Document
    Dim objTable As Word.Table
    Dim objColumns As Word.Columns
    Dim objCells As Word.Cell
    
    Dim objColList As SQLDMO.SQLObjectList
    Dim objDMOColumn As SQLDMO.Column2
    
    Dim intColumnCounter As Integer
    Dim intIndexCounter As Integer
    Dim intTriggerCounter As Integer
    Dim intKeyCounter As Integer
    Dim intKeyColumnsCounter As Integer
    Dim intViewCounter As Integer
    Dim intStored_ProcedureCounter As Integer
    Dim blnNextColumn As Boolean
    
    Set objWord = CreateObject("Word.Application")
    Set objDocument = objWord.Documents.Add

    objWord.Selection.PageSetup.TopMargin = 35
    objWord.Selection.PageSetup.LeftMargin = 25
    objWord.Selection.PageSetup.BottomMargin = 35
    objWord.Selection.PageSetup.RightMargin = 25
    'objWord.Selection.PageSetup.Orientation = wdOrientLandscape
    
    'objWord.Selection.Font.Name = "Courier New"
    objWord.Selection.Font.Name = "Verdana"
    objWord.Selection.Font.Size = 10
    objWord.Selection.Font.Bold = True
    objWord.Selection.TypeText "Database Name: " & objDatabase.Name & vbCrLf
    objWord.Selection.TypeText "Table Name: " & objDMO_Table.Owner & "." & objDMO_Table.Name & vbCrLf
    objWord.Selection.TypeText "Description: " & vbCrLf & vbCrLf
    
    'START - Columns
    objWord.Selection.TypeText "Columns " & vbCrLf
    objWord.Selection.Font.Size = 8
    objWord.Selection.Font.Bold = False
    
    Set objTable = objWord.Selection.Tables.Add(objWord.Selection.Range, objDMO_Table.Columns.Count + 1, 4)
    Set objColumns = objTable.Columns
    objColumns(1).Width = 35 'Type
    objColumns(2).Width = 150 'Field Name
    objColumns(3).Width = 80 'Data Type
    objColumns(4).Width = 230 'Description
    Set objCells = objTable.Cell(1, 1): objCells.Select: objWord.Selection.TypeText "Type"
    Set objCells = objTable.Cell(1, 2): objCells.Select: objWord.Selection.TypeText "Field Name"
    Set objCells = objTable.Cell(1, 3): objCells.Select: objWord.Selection.TypeText "Data Type"
    Set objCells = objTable.Cell(1, 4): objCells.Select: objWord.Selection.TypeText "Description"
        
    For intColumnCounter = 1 To objDMO_Table.Columns.Count
        With objDMO_Table.Columns(intColumnCounter)
            If .InPrimaryKey = True Then
                Set objCells = objTable.Cell(intColumnCounter + 1, 1): objCells.Select: objWord.Selection.TypeText "Pk"
            End If
            Set objCells = objTable.Cell(intColumnCounter + 1, 2): objCells.Select: objWord.Selection.TypeText .Name
            Set objCells = objTable.Cell(intColumnCounter + 1, 3): objCells.Select: objWord.Selection.TypeText .Datatype
            If UCase(.Datatype) = ("VARCHAR") Or UCase(.Datatype) = ("CHAR") Then
                objWord.Selection.TypeText "(" & .Length & ")"
            End If
        End With
    Next
    SetTableSetting objTable
    objWord.Selection.GoToNext wdGoToLine
    objWord.Selection.TypeText vbCrLf & vbCrLf
    
    Set objTable = Nothing
    Set objColumns = Nothing
    Set objCells = Nothing
    'END - Columns
    
    'START - Keys
    'objWord.Selection.Font.Name = "Courier New"
    objWord.Selection.Font.Name = "Verdana"
    objWord.Selection.Font.Size = 10
    objWord.Selection.Font.Bold = True
    
    objWord.Selection.TypeText "Constraints " & vbCrLf
    objWord.Selection.Font.Size = 8
    objWord.Selection.Font.Bold = False
    
    Set objTable = objWord.Selection.Tables.Add(objWord.Selection.Range, objDMO_Table.Keys.Count + 1, 4)
    Set objColumns = objTable.Columns
    objColumns(1).Width = 150 'Key Name
    objColumns(2).Width = 50 'Key Type
    objColumns(3).Width = 120 'Key Columns
    objColumns(4).Width = 230 'Description
    
    Set objCells = objTable.Cell(1, 1): objCells.Select: objWord.Selection.TypeText "Key Name"
    Set objCells = objTable.Cell(1, 2): objCells.Select: objWord.Selection.TypeText "Key Type"
    Set objCells = objTable.Cell(1, 3): objCells.Select: objWord.Selection.TypeText "Key Columns"
    Set objCells = objTable.Cell(1, 4): objCells.Select: objWord.Selection.TypeText "Description"
        
    Dim objKey As SQLDMO.Key '*****
    
    For Each objKey In objDMO_Table.Keys
        intKeyCounter = intKeyCounter + 1
        With objKey
            Set objCells = objTable.Cell(intKeyCounter + 1, 1): objCells.Select: objWord.Selection.TypeText .Name
            Set objCells = objTable.Cell(intKeyCounter + 1, 2): objCells.Select
            If .Type = SQLDMOKey_Primary Then
                objWord.Selection.TypeText "Primery"
            ElseIf .Type = SQLDMOKey_Unique Then
                objWord.Selection.TypeText "Unique"
            ElseIf .Type = SQLDMOKey_Foreign Then
                objWord.Selection.TypeText "Foregin"
            End If
            Set objCells = objTable.Cell(intKeyCounter + 1, 3): objCells.Select
            For intKeyColumnsCounter = 1 To .KeyColumns.Count
                If blnNextColumn = True Then objWord.Selection.TypeText ", "
                objWord.Selection.TypeText .KeyColumns.Item(intKeyColumnsCounter): blnNextColumn = True
            Next
            blnNextColumn = False
            If .Type = SQLDMOKey_Foreign Then
                objWord.Selection.Font.Color = wdColorSeaGreen
                objWord.Selection.TypeText " REFERENCES " & .ReferencedColumns(1) & "(" & .ReferencedTable & ")"
                objWord.Selection.Font.Color = wdColorBlack
            End If
        End With
    Next
    SetTableSetting objTable
    objWord.Selection.GoToNext wdGoToLine
    objWord.Selection.TypeText vbCrLf & vbCrLf
    
    Set objTable = Nothing
    Set objColumns = Nothing
    Set objCells = Nothing
    'END - Keys
    
    'START - Indexes
    'objWord.Selection.Font.Name = "Courier New"
    objWord.Selection.Font.Name = "Verdana"
    objWord.Selection.Font.Size = 10
    objWord.Selection.Font.Bold = True
        
    objWord.Selection.TypeText "Indexes " & vbCrLf
    objWord.Selection.Font.Size = 8
    objWord.Selection.Font.Bold = False
    
    Set objTable = objWord.Selection.Tables.Add(objWord.Selection.Range, objDMO_Table.Indexes.Count + 1, 4)
    Set objColumns = objTable.Columns
    objColumns(1).Width = 150 'Index Name
    objColumns(2).Width = 35 'Index Columns
    objColumns(3).Width = 80 'Index Type
    objColumns(4).Width = 230 'Description
    Set objCells = objTable.Cell(1, 1): objCells.Select: objWord.Selection.TypeText "Index Name"
    Set objCells = objTable.Cell(1, 2): objCells.Select: objWord.Selection.TypeText "Index Columns"
    Set objCells = objTable.Cell(1, 3): objCells.Select: objWord.Selection.TypeText "Index Type"
    Set objCells = objTable.Cell(1, 4): objCells.Select: objWord.Selection.TypeText "Description"
     
    For intIndexCounter = 1 To objDMO_Table.Indexes.Count
        With objDMO_Table.Indexes(intIndexCounter)
            Set objCells = objTable.Cell(intIndexCounter + 1, 1): objCells.Select: objWord.Selection.TypeText .Name
            
            Set objColList = .ListIndexedColumns
            Set objCells = objTable.Cell(intIndexCounter + 1, 2): objCells.Select:
            For Each objDMOColumn In objColList
                If blnNextColumn = True Then objWord.Selection.TypeText ", "
                objWord.Selection.TypeText objDMOColumn.Name: blnNextColumn = True
            Next
            blnNextColumn = False
            
            Set objCells = objTable.Cell(intIndexCounter + 1, 3): objCells.Select
            If .Type = (SQLDMOIndex_Default) Then
                objWord.Selection.TypeText "Non Clustered"
            ElseIf .Type = SQLDMOIndex_Clustered Then
                objWord.Selection.TypeText "Clustered"
            ElseIf .Type = SQLDMOIndex_DRIPrimaryKey Then
                objWord.Selection.TypeText "Primery"
            ElseIf .Type = SQLDMOIndex_Unique Then
                objWord.Selection.TypeText "Unique"
            ElseIf .Type = (SQLDMOIndex_Clustered Or SQLDMOIndex_Unique Or SQLDMOIndex_DropExist Or SQLDMOIndex_Valid) Then
                objWord.Selection.TypeText "Unique/Clustered"
            End If
        End With
    Next
    SetTableSetting objTable
    objWord.Selection.GoToNext wdGoToLine
    objWord.Selection.TypeText vbCrLf & vbCrLf
    
    Set objTable = Nothing
    Set objColumns = Nothing
    Set objCells = Nothing
    'END - Indexes
    
    'START - Triggers
    'objWord.Selection.Font.Name = "Courier New"
    objWord.Selection.Font.Name = "Verdana"
    objWord.Selection.Font.Size = 10
    objWord.Selection.Font.Bold = True
    
    objWord.Selection.TypeText "Triggers " & vbCrLf
    objWord.Selection.Font.Size = 8
    objWord.Selection.Font.Bold = False
    
    Set objTable = objWord.Selection.Tables.Add(objWord.Selection.Range, objDMO_Table.Triggers.Count + 1, 3)
    Set objColumns = objTable.Columns
    objColumns(1).Width = 150 'Trigger Name
    objColumns(2).Width = 120 'Trigger Type
    objColumns(3).Width = 230 'Description
    
    Set objCells = objTable.Cell(1, 1): objCells.Select: objWord.Selection.TypeText "Trigger Name"
    Set objCells = objTable.Cell(1, 2): objCells.Select: objWord.Selection.TypeText "Trigger Type"
    Set objCells = objTable.Cell(1, 3): objCells.Select: objWord.Selection.TypeText "Description"
        
    Dim objTrigger As SQLDMO.Trigger2
    
    For Each objTrigger In objDMO_Table.Triggers
        intTriggerCounter = intTriggerCounter + 1
        With objTrigger
            Set objCells = objTable.Cell(intTriggerCounter + 1, 1): objCells.Select: objWord.Selection.TypeText .Name
            
            Set objCells = objTable.Cell(intTriggerCounter + 1, 2): objCells.Select
            If .InsteadOfTrigger = True Then
                If .Type = SQLDMOTrig_Delete Then objWord.Selection.TypeText "Instead Of Delete"
                If .Type = SQLDMOTrig_Update Then objWord.Selection.TypeText "Instead Of Update"
                If .Type = SQLDMOTrig_Insert Then objWord.Selection.TypeText "Instead Of Insert"
                If .Type = SQLDMOTrig_All Then objWord.Selection.TypeText "Instead Of Insert, Update, Delete"
                If .Type = SQLDMOTrig_Insert Or SQLDMOTrig_Update Then objWord.Selection.TypeText "Instead Of Insert, Update"
                If .Type = SQLDMOTrig_Insert Or SQLDMOTrig_Delete Then objWord.Selection.TypeText "Instead Of Insert, Delete"
                If .Type = SQLDMOTrig_Update Or SQLDMOTrig_Delete Then objWord.Selection.TypeText "Instead Of Update, Delete"
            ElseIf .AfterTrigger = True Then
                If .Type = SQLDMOTrig_Delete Then objWord.Selection.TypeText "After Delete"
                If .Type = SQLDMOTrig_Update Then objWord.Selection.TypeText "After Update"
                If .Type = SQLDMOTrig_Insert Then objWord.Selection.TypeText "After Insert"
                If .Type = SQLDMOTrig_All Then objWord.Selection.TypeText "After Insert, Update, Delete"
                If .Type = (SQLDMOTrig_Insert Or SQLDMOTrig_Update) Then objWord.Selection.TypeText "After Insert, Update"
                If .Type = (SQLDMOTrig_Insert Or SQLDMOTrig_Delete) Then objWord.Selection.TypeText "After Insert, Delete"
                If .Type = (SQLDMOTrig_Update Or SQLDMOTrig_Delete) Then objWord.Selection.TypeText "After Update, Delete"
            End If
        End With
    Next
    SetTableSetting objTable
    objWord.Selection.GoToNext wdGoToLine
    objWord.Selection.TypeText vbCrLf & vbCrLf
    
    Set objTable = Nothing
    Set objColumns = Nothing
    Set objCells = Nothing
    'END - Triggers
    
    'START - Dependencies
    Dim objResult As SQLDMO.QueryResults2
    Dim i, j As Integer '*****
    Dim strViews() As String
    Dim strStored_Procedures() As String
    
    Set objResult = objDMO_Table.EnumDependencies(SQLDMODep_Children)
    
    ReDim strViews(0) As String
    ReDim strStored_Procedures(0) As String
    
    For i = 1 To objResult.ResultSets
        objResult.CurrentResultSet = i
        For j = 1 To objResult.Rows
            If Val(objResult.GetColumnString(j, 1)) = 4 Then
                ReDim Preserve strViews(UBound(strViews) + 1) As String
                strViews(UBound(strViews)) = objResult.GetColumnString(j, 2)
            End If
            If Val(objResult.GetColumnString(j, 1)) = 16 Then
                ReDim Preserve strStored_Procedures(UBound(strStored_Procedures) + 1) As String
                strStored_Procedures(UBound(strStored_Procedures)) = objResult.GetColumnString(j, 2)
            End If
        Next
    Next
      
    'START - View
    'objWord.Selection.Font.Name = "Courier New"
    objWord.Selection.Font.Name = "Verdana"
    objWord.Selection.Font.Size = 10
    objWord.Selection.Font.Bold = True

    objWord.Selection.TypeText "Views " & vbCrLf
    objWord.Selection.Font.Size = 8
    objWord.Selection.Font.Bold = False

    Set objTable = objWord.Selection.Tables.Add(objWord.Selection.Range, UBound(strViews) + 1, 2)
    Set objColumns = objTable.Columns
    objColumns(1).Width = 150 'View Name
    objColumns(2).Width = 230 'Description

    Set objCells = objTable.Cell(1, 1): objCells.Select: objWord.Selection.TypeText "View Name"
    Set objCells = objTable.Cell(1, 2): objCells.Select: objWord.Selection.TypeText "Description"

    For intViewCounter = 1 To UBound(strViews)
        Set objCells = objTable.Cell(intViewCounter + 1, 1): objCells.Select: objWord.Selection.TypeText strViews(intViewCounter)
    Next
    SetTableSetting objTable
    objWord.Selection.GoToNext wdGoToLine
    objWord.Selection.TypeText vbCrLf & vbCrLf

    Set objTable = Nothing
    Set objColumns = Nothing
    Set objCells = Nothing
    'END - View
    
    'START - Stored Procedure
    'objWord.Selection.Font.Name = "Courier New"
    objWord.Selection.Font.Name = "Verdana"
    objWord.Selection.Font.Size = 10
    objWord.Selection.Font.Bold = True

    objWord.Selection.TypeText "Stored Procedures " & vbCrLf
    objWord.Selection.Font.Size = 8
    objWord.Selection.Font.Bold = False

    Set objTable = objWord.Selection.Tables.Add(objWord.Selection.Range, UBound(strStored_Procedures) + 1, 2)
    Set objColumns = objTable.Columns
    objColumns(1).Width = 150 'Stored Procedure Name
    objColumns(2).Width = 230 'Description

    Set objCells = objTable.Cell(1, 1): objCells.Select: objWord.Selection.TypeText "Stored Procedure Name"
    Set objCells = objTable.Cell(1, 2): objCells.Select: objWord.Selection.TypeText "Description"

    For intStored_ProcedureCounter = 1 To UBound(strStored_Procedures)
        Set objCells = objTable.Cell(intStored_ProcedureCounter + 1, 1): objCells.Select: objWord.Selection.TypeText strStored_Procedures(intStored_ProcedureCounter)
    Next
    SetTableSetting objTable
    objWord.Selection.GoToNext wdGoToLine
    objWord.Selection.TypeText vbCrLf & vbCrLf

    Set objTable = Nothing
    Set objColumns = Nothing
    Set objCells = Nothing
    'END - Stored Procedure
    'END - Dependencies
    
    objDocument.SaveAs strPath & "\" & cboDatabases.Text & "\Tables\" & objDMO_Table.Name & ".doc"
    objDocument.Close
  
    objWord.Quit
    Set objWord = Nothing
    Set objDocument = Nothing
    Set objTable = Nothing
    Set objColumns = Nothing
    Set objCells = Nothing
End Function

Private Function CreateSPDocument(objDMO_SP As SQLDMO.StoredProcedure2, ByVal strPath As String)
    Dim objWord As Word.Application
    Dim objDocument As Word.Document
    Dim objTable As Word.Table
    Dim objColumns As Word.Columns
    Dim objCells As Word.Cell
    Dim objResult As SQLDMO.QueryResults2
    
    Dim intResultCounter As Integer
    
    Set objWord = CreateObject("Word.Application")
    Set objDocument = objWord.Documents.Add

    Set objResult = objDMO_SP.EnumParameters

    objWord.Selection.PageSetup.TopMargin = 35
    objWord.Selection.PageSetup.LeftMargin = 25
    objWord.Selection.PageSetup.BottomMargin = 35
    objWord.Selection.PageSetup.RightMargin = 25
    'objWord.Selection.PageSetup.Orientation = wdOrientLandscape

    objWord.Selection.Font.Name = "Verdana"
    objWord.Selection.Font.Size = 10
    objWord.Selection.Font.Bold = True
    objWord.Selection.TypeText "Database Name: " & objDatabase.Name & vbCrLf
    objWord.Selection.TypeText "Stored Procedure Name: " & objDMO_SP.Owner & "." & objDMO_SP.Name & vbCrLf
    objWord.Selection.TypeText "Description: " & vbCrLf & vbCrLf

    objWord.Selection.Font.Size = 8
    objWord.Selection.Font.Bold = False
            
    Set objTable = objWord.Selection.Tables.Add(objWord.Selection.Range, objResult.Rows + 1, 4)
    Set objColumns = objTable.Columns
    objColumns(1).Width = 150 'Parameter Name
    objColumns(2).Width = 80 'Data Type
    objColumns(3).Width = 35 'In/Out
    objColumns(4).Width = 230 'Description
    Set objCells = objTable.Cell(1, 1): objCells.Select: objWord.Selection.TypeText "Parameter Name"
    Set objCells = objTable.Cell(1, 2): objCells.Select: objWord.Selection.TypeText "Data Type"
    Set objCells = objTable.Cell(1, 3): objCells.Select: objWord.Selection.TypeText "In/Out"
    Set objCells = objTable.Cell(1, 4): objCells.Select: objWord.Selection.TypeText "Description"
    
    For intResultCounter = 1 To objResult.Rows
        With objResult
            Set objCells = objTable.Cell(intResultCounter + 1, 1): objCells.Select: objWord.Selection.TypeText .GetColumnString(intResultCounter, 1)
            Set objCells = objTable.Cell(intResultCounter + 1, 2): objCells.Select: objWord.Selection.TypeText .GetColumnString(intResultCounter, 2)
            If UCase(.GetColumnString(intResultCounter, 2)) = ("VARCHAR") Or UCase(.GetColumnString(intResultCounter, 3)) = ("CHAR") Then
                objWord.Selection.TypeText "(" & .GetColumnString(intResultCounter, 3) & ")"
            End If
            If .GetColumnString(intResultCounter, 5) = 1 Then
                Set objCells = objTable.Cell(intResultCounter + 1, 3): objCells.Select: objWord.Selection.TypeText "Output"
            Else
                Set objCells = objTable.Cell(intResultCounter + 1, 3): objCells.Select: objWord.Selection.TypeText "Input"
            End If
        End With
    Next
    SetTableSetting objTable
    objWord.Selection.GoToNext wdGoToLine
    objWord.Selection.TypeText vbCrLf & vbCrLf
    
    objDocument.SaveAs strPath & "\" & cboDatabases.Text & "\Stored Procedures\" & objDMO_SP.Name & ".doc"
    objDocument.Close
  
    objWord.Quit
    Set objWord = Nothing
    Set objDocument = Nothing
    Set objTable = Nothing
    Set objColumns = Nothing
    Set objCells = Nothing
End Function

Private Function SetTableSetting(objTable As Word.Table)
    objTable.AllowAutoFit = True
    objTable.Borders.InsideLineStyle = wdLineStyleSingle
    objTable.Borders.OutsideLineStyle = wdLineStyleSingle
    objTable.Borders.InsideColor = wdColorGray30
    objTable.Borders.OutsideColor = wdColorGray30
    objTable.AutoFitBehavior wdAutoFitWindow
    objTable.Rows(1).Shading.Texture = wdTexture10Percent
End Function

Private Function EnableControls() As Boolean
    lblOptions(0).Enabled = True: cboDatabases.Enabled = True: cmdConnect.Enabled = True
    lblOptions(1).Enabled = True: txtPath.Enabled = True: cmdPath.Enabled = True
    optObjectOptions(0).Enabled = True
    optObjectOptions(1).Enabled = True
    chkSelect.Enabled = True
    cmdDocument.Enabled = True
End Function

Private Function DisableControls() As Boolean
    lblOptions(0).Enabled = False: cboDatabases.Enabled = False: cmdConnect.Enabled = False
    lblOptions(1).Enabled = False: txtPath.Enabled = False: cmdPath.Enabled = False
    optObjectOptions(0).Enabled = False
    optObjectOptions(1).Enabled = False
    chkSelect.Enabled = False
    cmdDocument.Enabled = False
End Function
