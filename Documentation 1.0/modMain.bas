Attribute VB_Name = "modMain"
Option Explicit
'START - API Declarations
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Type BrowseInfo
    hWndOwner As Long
    pIDLRoot As String
    pszDisplayName As String
    lpszTitle As Long
    ulFlags As Long
    lpfnCallback As Long
    lParam As Long
    iImage As Long
End Type
Private Const BIF_RETURNONLYFSDIRS = 1
Private Const MAX_PATH = 260
Private Const SW_SHOWNORMAL = 1
'END - API Declarations

Public APPLICATION_NAME As String
Public objServer As New SQLDMO.SQLServer
Public objDatabase As New SQLDMO.Database2

Public Sub Main()
    APPLICATION_NAME = "Documentation 1.0"
    
    Load frmLogin
    frmLogin.Show
End Sub

Public Function AvailableSQLServers(objCombo As ComboBox) As Boolean
    Dim objServers As New SQLDMO.Application
    Dim oNameList As SQLDMO.NameList
    Dim iElement As Integer
    Dim lCtr As Long, lCount As Long

    On Error GoTo ErrorHandler

    Set oNameList = objServers.ListAvailableSQLServers

    With oNameList
        lCount = .Count
        If lCount > 0 Then
            For lCtr = 1 To .Count
                objCombo.AddItem oNameList.Item(lCtr)
            Next
        End If
    End With

    AvailableSQLServers = True
    Set objServers = Nothing
    Set oNameList = Nothing
    Exit Function
ErrorHandler:
    MsgBox Err.Description
    AvailableSQLServers = False
End Function

Public Function SelectFolder(ByVal hWnd As Long) As String
    Dim iNull As Integer, lpIDList As Long, lResult As Long
    Dim sPath As String, udtBI As BrowseInfo

    With udtBI
        .hWndOwner = hWnd 'Set the owner window
        .lpszTitle = lstrcat("Select Folder", "") 'lstrcat appends the two strings and returns the memory address
        .ulFlags = BIF_RETURNONLYFSDIRS 'Return only if the user selected a directory
    End With

    'Show the 'Browse for folder' dialog
    lpIDList = SHBrowseForFolder(udtBI)
    If lpIDList Then
        sPath = String$(MAX_PATH, 0)
        'Get the path from the IDList
        SHGetPathFromIDList lpIDList, sPath
        'free the block of memory
        CoTaskMemFree lpIDList
        iNull = InStr(sPath, vbNullChar)
        If iNull Then
            sPath = Left$(sPath, iNull - 1)
        End If
    End If
    SelectFolder = sPath
End Function

Public Function WriteToInIFile(ByVal strSection As String, ByVal strKey As String, ByVal strKeyValue As String, ByVal strFileName As String) As Boolean
    Dim dl As Long
    
    dl = WritePrivateProfileString(strSection, strKey, strKeyValue, strFileName)
    If dl <> 0 Then
        WriteToInIFile = True
    Else
        WriteToInIFile = False
    End If
End Function

Public Function ReadFromInIFile(ByVal strSection As String, ByVal strKey As String, ByRef strKeyValue As String, ByVal strFileName As String) As Boolean
    Dim dl As Long
    Dim strReturnValue As String
    
    strReturnValue = Space(255)
    dl = GetPrivateProfileString(strSection, strKey, "", strReturnValue, 255, strFileName)
    If dl <> 0 Then
        strKeyValue = Left(strReturnValue, dl)
        ReadFromInIFile = True
    Else
        strKeyValue = ""
        ReadFromInIFile = False
    End If
End Function

Public Function Progress(objProgress As ProgressBar, ByVal intProgressValue As Integer) As Boolean
    Dim intCounter As Integer
    
    If intProgressValue < objProgress.Min Then
        Progress = False
        Exit Function
    End If
    
    If objProgress.Value < intProgressValue Then
        For intCounter = 1 To (intProgressValue - objProgress.Value)
            objProgress.Value = objProgress.Value + 1
            Sleep 50
        Next
    ElseIf objProgress.Value > intProgressValue Then
        For intCounter = (objProgress.Value - intProgressValue) To 1 Step -1
            objProgress.Value = objProgress.Value - 1
            Sleep 50
        Next
    End If
    
    DoEvents
    Progress = True
End Function

Public Function ShowFile(ByVal hWnd As Long, ByVal strFilePath As String)
    ShellExecute hWnd, vbNullString, strFilePath, vbNullString, "C:\", SW_SHOWNORMAL
End Function
