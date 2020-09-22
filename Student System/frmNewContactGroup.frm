VERSION 5.00
Begin VB.Form frmNewContactGroup 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "---"
   ClientHeight    =   3630
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5415
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmNewContactGroup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3630
   ScaleWidth      =   5415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraUser 
      Height          =   3495
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   5175
      Begin VB.TextBox txtGroupName 
         Height          =   375
         Left            =   1560
         MaxLength       =   255
         TabIndex        =   4
         Top             =   600
         Width           =   3495
      End
      Begin VB.TextBox txtNotes 
         Height          =   1215
         IMEMode         =   3  'DISABLE
         Left            =   1560
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   1080
         Width           =   3495
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add Group"
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
         Left            =   2880
         Picture         =   "frmNewContactGroup.frx":1082
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   2520
         Width           =   1215
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
         Height          =   735
         Left            =   4080
         Picture         =   "frmNewContactGroup.frx":2104
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   2520
         Width           =   975
      End
      Begin VB.Label lblUsername 
         Alignment       =   1  'Right Justify
         Caption         =   "Group Name"
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
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label lblPassword 
         Alignment       =   1  'Right Justify
         Caption         =   "Notes"
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
         Left            =   240
         TabIndex        =   7
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label lblRecord 
         Caption         =   "Record #"
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
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lblRecordID 
         Caption         =   "---"
         Height          =   255
         Left            =   1560
         TabIndex        =   5
         Top             =   240
         Width           =   3495
      End
   End
End
Attribute VB_Name = "frmNewContactGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdAdd_Click()
If WarnBlank(txtGroupName, "GroupName") = True Then Exit Sub

If cmdAdd.Caption = "Add Group" Then
    con.Execute ("INSERT INTO ContactGroups VALUES(" & lblRecordID.Caption & ", '" & Replace(txtGroupName.Text, "'", "''") & "', '" & Replace(txtNotes.Text, "'", "''") & "'," & Replace(UserID, "'", "''") & ")")
    MsgBox "New Contact Group successfully added.", vbOKOnly + vbInformation, "Success"
ElseIf cmdAdd.Caption = "Update" Then
    con.Execute ("UPDATE ContactGroups SET GroupName='" & Replace(txtGroupName.Text, "'", "''") & "',Notes='" & Replace(txtNotes.Text, "'", "''") & "' WHERE GroupID=" & Val(lblRecordID.Caption) & "")
    MsgBox "Contact Group information successfully updated.", vbOKOnly + vbInformation, "Success"
End If
Unload Me
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub Form_Activate()
If con.State = 0 Then Call konek
End Sub

