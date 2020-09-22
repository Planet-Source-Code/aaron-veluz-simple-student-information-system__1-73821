VERSION 5.00
Begin VB.Form frmSetAdmin 
   Caption         =   "Set Default User"
   ClientHeight    =   3000
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4920
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
   Icon            =   "frmSetAdmin.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3000
   ScaleWidth      =   4920
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraDefault 
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4695
      Begin VB.CommandButton cmdSet 
         Caption         =   "Set"
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
         Height          =   855
         Left            =   2640
         Picture         =   "frmSetAdmin.frx":1082
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1800
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   3600
         Picture         =   "frmSetAdmin.frx":2104
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   1800
         Width           =   975
      End
      Begin VB.TextBox txtUsername 
         Height          =   375
         Left            =   2040
         MaxLength       =   255
         TabIndex        =   1
         Top             =   240
         Width           =   2535
      End
      Begin VB.TextBox txtPassword 
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   2040
         MaxLength       =   255
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   720
         Width           =   2535
      End
      Begin VB.TextBox txtRetype 
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   2040
         MaxLength       =   255
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   1200
         Width           =   2535
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
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   1095
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
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label lblRetype 
         Caption         =   "Retype Password:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1200
         Width           =   1935
      End
   End
End
Attribute VB_Name = "frmSetAdmin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdCancel_Click()
If MsgBox("Are you sure you want to exit the application?", vbYesNo + vbExclamation, "Confirm Exit Application") = vbNo Then Exit Sub
End
End Sub

Private Sub cmdSet_Click()
If WarnBlank(txtUsername, "Username") = True Then Exit Sub
If WarnBlank(txtPassword, "Password") = True Then Exit Sub

If Len(txtUsername.Text) < 6 Then
    MsgBox "Username must consist of 6 or more characters.", vbOKOnly + vbExclamation, "Invalid Length"
    Call SelText(txtUsername)
ElseIf Len(txtPassword.Text) < 6 Then
    MsgBox "Password must consist of 6 or more characters.", vbOKOnly + vbExclamation, "Invalid Length"
    Call SelText(txtPassword)
ElseIf txtRetype.Text <> txtPassword.Text Then
    MsgBox "Password entries did not match.", vbOKOnly + vbExclamation, "Invalid Password"
    txtPassword.Text = ""
    txtRetype.Text = ""
    txtPassword.SetFocus
Else
    con.Execute ("INSERT INTO Users VALUES(1, '" & txtUsername.Text & "', '" & txtPassword.Text & "', 'SuperAdministrator')")
    MsgBox "New SuperAdministrator account successfully added.", vbOKOnly + vbInformation, "Success"
    Username = txtUsername.Text
    Password = txtPassword.Text
    Privilege = "SuperAdministrator"
    Call Login
    Unload Me
    frmMain.Show
End If
End Sub

Private Sub Form_Activate()
If con.State = 0 Then Call konek
End Sub

