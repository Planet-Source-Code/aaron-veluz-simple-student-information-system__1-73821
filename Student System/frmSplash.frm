VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSplash 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   ClientHeight    =   3825
   ClientLeft      =   210
   ClientTop       =   1365
   ClientWidth     =   7380
   ClipControls    =   0   'False
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
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3825
   ScaleWidth      =   7380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   75
      Left            =   480
      TabIndex        =   1
      Top             =   3120
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   132
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   6120
      Top             =   1800
   End
   Begin VB.Label lblProgress 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "---"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   480
      TabIndex        =   2
      Top             =   3360
      Width           =   6495
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "---"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1575
      Left            =   0
      TabIndex        =   0
      Top             =   1080
      Width           =   7335
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
lblTitle.Caption = App.FileDescription
End Sub

Private Sub Timer1_Timer()
pBar.Value = pBar.Value + 1
lblProgress.Caption = pBar.Value & "%"
If pBar.Value = 100 Then
    Unload Me
    Set rs = con.Execute("SELECT * FROM Users")
    If rs.RecordCount = 0 Then
        frmSetAdmin.Show
    Else
        frmLogin.Show
    End If
End If
End Sub

