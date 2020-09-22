VERSION 5.00
Begin VB.Form frmBackupRestore 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Backup/Restore Database"
   ClientHeight    =   3615
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5550
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmBackupRestore.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   5550
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Height          =   735
      Left            =   4320
      Picture         =   "frmBackupRestore.frx":1082
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2760
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
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
      Left            =   3240
      Picture         =   "frmBackupRestore.frx":2104
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Frame fraAction 
      Caption         =   "Select Action"
      Height          =   855
      Left            =   120
      TabIndex        =   7
      Top             =   2640
      Width           =   2655
      Begin VB.OptionButton optRestore 
         Caption         =   "Restore"
         Height          =   255
         Left            =   1320
         TabIndex        =   9
         Top             =   360
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.OptionButton optBackup 
         Caption         =   "Backup"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Value           =   -1  'True
         Width           =   1095
      End
   End
   Begin VB.Frame frabackupRestore 
      Caption         =   "Select Database Path"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5295
      Begin VB.FileListBox fil 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   1200
         Pattern         =   "*.mdb"
         TabIndex        =   3
         Top             =   1800
         Width           =   3975
      End
      Begin VB.DirListBox dir 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1050
         Left            =   1200
         TabIndex        =   2
         Top             =   720
         Width           =   3975
      End
      Begin VB.DriveListBox drv 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1200
         TabIndex        =   1
         Top             =   360
         Width           =   3975
      End
      Begin VB.Label Label1 
         Caption         =   "Files"
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
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label lblDirectory 
         Caption         =   "Directory"
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
         TabIndex        =   5
         Top             =   720
         Width           =   975
      End
      Begin VB.Label lblDrive 
         Caption         =   "Drive"
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
         TabIndex        =   4
         Top             =   360
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmBackupRestore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()
On Error GoTo Hell

con.Close
If optBackup.Value = True Then
    strTest = dir.Path
    If Right(Directory, 1) = Chr(92) Then strTest = Left(strTest, (Len(strTest) - 1))
    FileCopy App.Path & "\MainDB.mdb", strTest & "\Backup_" & Format(Date, "mm-dd-yyyy") & ".mdb"
    MsgBox "Backup procedure was successful.", vbOKOnly + vbInformation, "Success"
    Unload Me
Else
    If fil.ListIndex = -1 Then MsgBox "Please select a database to be used for restoration.", vbOKOnly + vbExclamation, "Select Database": Exit Sub
    If MsgBox("Restoring a database will cause the system to terminate. Proceed?", vbQuestion + vbYesNo + vbDefaultButton2, "Confirm Restore") = vbYes Then
        FileCopy fil.Path & "\" & fil.FileName, App.Path & "\MainDB.mdb"
        MsgBox "Database was restored. The application will now be closed.", vbOKOnly + vbInformation, "Success"
        End
    End If
End If

Exit Sub
Hell:
    MsgBox Err.Description, vbCritical, "Error"
    Exit Sub
End Sub

Private Sub dir_Change()
    fil.Path = dir.Path
End Sub

Private Sub drv_Change()
    On Error Resume Next
    dir.Path = drv.Drive
End Sub

Private Sub optBackup_Click()
fil.enabled = False
End Sub

Private Sub optRestore_Click()
fil.enabled = True
End Sub
