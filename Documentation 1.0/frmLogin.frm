VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Connect"
   ClientHeight    =   2685
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4425
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2685
   ScaleWidth      =   4425
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkLoginSecure 
      Caption         =   "Login Secure"
      Height          =   315
      Left            =   1170
      TabIndex        =   6
      Top             =   1650
      Width           =   1275
   End
   Begin VB.TextBox txtPassword 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1170
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   1170
      Width           =   3075
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancle"
      Height          =   435
      Left            =   2355
      TabIndex        =   8
      Top             =   2070
      Width           =   1215
   End
   Begin VB.CommandButton cmdLog_In 
      Caption         =   "&Log In"
      Default         =   -1  'True
      Height          =   435
      Left            =   855
      TabIndex        =   7
      Top             =   2070
      Width           =   1215
   End
   Begin VB.TextBox txtUser_Id 
      Height          =   315
      Left            =   1170
      TabIndex        =   3
      Top             =   690
      Width           =   3075
   End
   Begin VB.ComboBox cboServers 
      Height          =   315
      Left            =   1170
      TabIndex        =   1
      Top             =   210
      Width           =   3075
   End
   Begin VB.Label lblLogin 
      AutoSize        =   -1  'True
      Caption         =   "&Password"
      Height          =   195
      Index           =   2
      Left            =   330
      TabIndex        =   4
      Top             =   1230
      Width           =   690
   End
   Begin VB.Label lblLogin 
      AutoSize        =   -1  'True
      Caption         =   "&User Id"
      Height          =   195
      Index           =   1
      Left            =   300
      TabIndex        =   2
      Top             =   750
      Width           =   510
   End
   Begin VB.Label lblLogin 
      AutoSize        =   -1  'True
      Caption         =   "Server"
      Height          =   195
      Index           =   0
      Left            =   300
      TabIndex        =   0
      Top             =   270
      Width           =   465
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    AvailableSQLServers cboServers
End Sub

Private Sub chkLoginSecure_Click()
    If chkLoginSecure.Value = 1 Then
        lblLogin(1).Enabled = False
        txtUser_Id.Text = ""
        txtUser_Id.Enabled = False
        lblLogin(2).Enabled = False
        txtPassword.Text = ""
        txtPassword.Enabled = False
    Else
        lblLogin(1).Enabled = True
        txtUser_Id.Text = ""
        txtUser_Id.Enabled = True
        lblLogin(2).Enabled = True
        txtPassword.Text = ""
        txtPassword.Enabled = True
    End If
End Sub

Private Sub cmdLog_In_Click()
    If Login(True) = True Then
        Unload Me
        Load frmDocumentation
        frmDocumentation.Show
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
    End
End Sub

Private Function Login(ByVal blnPrompt As Boolean) As Boolean
    Dim oDB As New SQLDMO.Database
    
    On Error GoTo errHand
        
    Screen.MousePointer = vbHourglass
    If cboServers.ListIndex >= 0 Or Len(Trim(cboServers.Text)) > 0 Then
        If chkLoginSecure.Value <> 1 Then
            If Len(Trim(txtUser_Id.Text)) > 0 Then
                objServer.LoginSecure = False
            Else
                If blnPrompt = True Then
                    MsgBox "Please Specify the Login Id."
                    txtUser_Id.SetFocus
                End If
            End If
        Else
            objServer.LoginSecure = True
        End If
    Else
        If blnPrompt = True Then
            MsgBox "Please Select the Server to Login in."
            cboServers.SetFocus
        End If
    End If

    objServer.Connect cboServers.Text, Trim(txtUser_Id.Text), Trim(txtPassword.Text)
    Screen.MousePointer = vbNormal

    Set oDB = Nothing
    Login = True
    Exit Function
errHand:
    Screen.MousePointer = vbNormal
    If blnPrompt = True Then MsgBox Err.Description
    Login = False
End Function
