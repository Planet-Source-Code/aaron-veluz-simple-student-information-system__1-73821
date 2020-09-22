VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "User Login"
   ClientHeight    =   2295
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4215
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2295
   ScaleWidth      =   4215
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraLogin 
      Height          =   2175
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   3975
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   2880
         Picture         =   "frmLogin.frx":1082
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1200
         Width           =   975
      End
      Begin VB.CommandButton cmdLogin 
         Caption         =   "&Login"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   1920
         Picture         =   "frmLogin.frx":2104
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1200
         Width           =   975
      End
      Begin VB.TextBox txtPassword 
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   1320
         MaxLength       =   255
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   720
         Width           =   2535
      End
      Begin VB.TextBox txtUsername 
         Height          =   375
         Left            =   1320
         MaxLength       =   255
         TabIndex        =   1
         Top             =   240
         Width           =   2535
      End
      Begin VB.Label lblPassword 
         Caption         =   "Password:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label lblUsername 
         Caption         =   "Username:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdCancel_Click()
If MsgBox("Are you sure you want to exit the application?", vbYesNo + vbExclamation, "Confirm Exit Application") = vbNo Then Exit Sub
End
End Sub

Private Sub cmdLogin_Click()
If WarnBlank(txtUsername, "Username") = True Then Exit Sub
If WarnBlank(txtPassword, "Password") = True Then Exit Sub
If recfound("UserName", txtUsername.Text) = False Then
    ErrCounter = ErrCounter + 1
    Call MistakeCounter("Username")
    Call SelText(txtUsername)
Else
    If Password <> txtPassword.Text Then
        ErrCounter = ErrCounter + 1
        Call MistakeCounter("Password")
        Call SelText(txtPassword)
    Else
        Call Login
        Unload Me
        
        
        With frmMain
            .sBar.Panels(2).Text = Format(Now, "DDDD, mmmm dd, yyyy hh:mm:ss am/pm")
            .sBar.Panels(1).Text = "Current User: " & Username & " (" & Privilege & ")"
            .Show
        End With
        Call EnableControls
        MsgBox "Welcome to the system, " & Username & ".", vbOKOnly + vbInformation, "Login"
        ErrCounter = 0
    End If
End If
End Sub

Private Sub Form_Activate()
If con.State = 0 Then Call konek
End Sub

Private Sub Form_Load()
ErrCounter = 0
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)

End Sub

Private Sub txtUsername_KeyPress(KeyAscii As Integer)

End Sub
