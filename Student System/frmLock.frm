VERSION 5.00
Begin VB.Form frmLock 
   Caption         =   "Enter Password to Unlock"
   ClientHeight    =   1455
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4440
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
   LinkTopic       =   "Form1"
   ScaleHeight     =   1455
   ScaleWidth      =   4440
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraLock 
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4215
      Begin VB.CommandButton cmdUnlock 
         Caption         =   "Unlock"
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
         Left            =   3120
         Picture         =   "frmLock.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtPassword 
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   240
         MaxLength       =   255
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   360
         Width           =   2655
      End
   End
End
Attribute VB_Name = "frmLock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdUnlock_Click()
If txtPassword.Text <> Password Then
    ErrCounter = ErrCounter + 1
    Call MistakeCounter("Password")
    Call SelText(txtPassword)
Else
    Unload Me
End If
End Sub

Private Sub Form_Activate()
txtPassword.SetFocus
End Sub

Private Sub Form_Load()
ErrCounter = 0
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)

End Sub
