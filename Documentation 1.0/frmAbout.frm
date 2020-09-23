VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About"
   ClientHeight    =   1425
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6570
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1425
   ScaleWidth      =   6570
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   4500
      Picture         =   "frmAbout.frx":0000
      ScaleHeight     =   855
      ScaleWidth      =   1965
      TabIndex        =   0
      Top             =   450
      Width           =   1965
   End
   Begin VB.Label lblApplication 
      BackColor       =   &H00FFFFFF&
      Caption         =   "lblApplication"
      Height          =   285
      Left            =   90
      TabIndex        =   2
      Top             =   120
      Width           =   5295
   End
   Begin VB.Label lblAbout 
      BackColor       =   &H00FFFFFF&
      Caption         =   "lblAbout"
      Height          =   855
      Left            =   90
      TabIndex        =   1
      Top             =   450
      Width           =   4335
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Me.Caption = "About " & APPLICATION_NAME
    lblApplication.Caption = APPLICATION_NAME
    lblAbout.Caption = "Author : Manish Patil (Nagpur+Pune - India)" & vbCrLf
    lblAbout.Caption = lblAbout.Caption & "Application : " & APPLICATION_NAME & vbCrLf
    lblAbout.Caption = lblAbout.Caption & Space(21) & "Used to Document a SQL Server Database." & vbCrLf
    lblAbout.Caption = lblAbout.Caption & "e-mail : manish.patil@hotmail.com"
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then Unload Me
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then Unload Me
End Sub

Private Sub lblAbout_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then Unload Me
End Sub
