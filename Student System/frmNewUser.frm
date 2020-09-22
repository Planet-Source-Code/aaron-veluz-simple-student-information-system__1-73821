VERSION 5.00
Begin VB.Form frmNewUser 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "---"
   ClientHeight    =   3615
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5430
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmNewUser.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   5430
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraUser 
      Height          =   3495
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   5175
      Begin VB.CheckBox chkMask 
         Caption         =   "Show/Hide Password"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   3000
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.TextBox txtRetype 
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   2520
         MaxLength       =   255
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   1560
         Width           =   2535
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
         Picture         =   "frmNewUser.frx":1082
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   2520
         Width           =   975
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add User"
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
         Picture         =   "frmNewUser.frx":2104
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   2520
         Width           =   1215
      End
      Begin VB.ComboBox cboPrivilege 
         Height          =   315
         ItemData        =   "frmNewUser.frx":3186
         Left            =   2520
         List            =   "frmNewUser.frx":3190
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   2040
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.TextBox txtPassword 
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   2520
         MaxLength       =   255
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   1080
         Width           =   2535
      End
      Begin VB.TextBox txtUsername 
         Height          =   375
         Left            =   2520
         MaxLength       =   255
         TabIndex        =   1
         Top             =   600
         Width           =   2535
      End
      Begin VB.Label lblUserPrivilege 
         Caption         =   "---"
         Height          =   255
         Left            =   2520
         TabIndex        =   14
         Top             =   2040
         Visible         =   0   'False
         Width           =   2535
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
         TabIndex        =   12
         Top             =   1560
         Width           =   1935
      End
      Begin VB.Label lblRecordID 
         Caption         =   "---"
         Height          =   255
         Left            =   2520
         TabIndex        =   11
         Top             =   240
         Width           =   2535
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
         TabIndex        =   10
         Top             =   240
         Width           =   975
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
         TabIndex        =   9
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label lblPrivilege 
         Caption         =   "Privilege:"
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
         Top             =   2040
         Visible         =   0   'False
         Width           =   1095
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
         TabIndex        =   7
         Top             =   600
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmNewUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkMask_Click()
If chkMask.Value = 0 Then
    txtPassword.PasswordChar = ""
    txtRetype.PasswordChar = ""
Else
    txtPassword.PasswordChar = "*"
    txtRetype.PasswordChar = "*"
End If
End Sub

Private Sub cmdAdd_Click()
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
    If cmdAdd.Caption = "Add User" Then
        Set rs = con.Execute("SELECT * FROM Users WHERE Username='" & txtUsername.Text & "'")
        If rs.RecordCount >= 1 Then
            MsgBox "Username already exists. Please specify another.", vbOKOnly + vbExclamation, "Invalid Username"
            Call SelText(txtUsername)
            Exit Sub
        Else
            con.Execute ("INSERT INTO Users VALUES(" & lblRecordID.Caption & ", '" & txtUsername.Text & "', '" & txtPassword.Text & "', '" & cboPrivilege.Text & "')")
            MsgBox "New user successfully added.", vbOKOnly + vbInformation, "Success"
            Unload Me
        End If
    ElseIf cmdAdd.Caption = "Update" Then
        Set rs = con.Execute("SELECT * FROM Users WHERE Username='" & txtUsername.Text & "' AND UserID<>" & lblRecordID.Caption & "")
        If rs.RecordCount = 0 Then
            If lblUserPrivilege.Caption = "SuperAdministrator" Then
                con.Execute ("UPDATE Users SET UserName='" & txtUsername.Text & "',UPassword='" & txtPassword.Text & "' WHERE UserID=1")
                con.Execute ("UPDATE LogTrail SET Username='" & txtUsername.Text & "' WHERE LUserID=1")
            Else
                con.Execute ("UPDATE Users SET Username='" & txtUsername.Text & "',UPassword='" & txtPassword.Text & "',Privilege='" & cboPrivilege.Text & "' WHERE UserID=" & Val(lblRecordID.Caption) & "")
                con.Execute ("UPDATE LogTrail SET Username='" & txtUsername.Text & "',Privilege='" & cboPrivilege.Text & "' WHERE LUserID=" & Val(lblRecordID.Caption) & "")
            End If
            If UserID = Val(lblRecordID.Caption) Then
                frmMain.sBar.Panels(1).Text = "Current User: " & txtUsername.Text & " (" & IIf((Privilege = "SuperAdministrator"), "SuperAdministrator", cboPrivilege.Text) & ")"
                Privilege = IIf((Privilege = "SuperAdministrator"), "SuperAdministrator", cboPrivilege.Text)
                Call EnableControls
            End If
            MsgBox "User information successfully updated.", vbOKOnly + vbInformation, "Success"
            Unload Me
        Else
            MsgBox "Username already exists. Please specify another.", vbOKOnly + vbExclamation, "Invalid Username"
            Call SelText(txtUsername)
            Exit Sub
        End If
    End If
End If
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub Form_Activate()
If con.State = 0 Then Call konek
End Sub

Private Sub Form_Load()
Call chkMask_Click
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)

End Sub

Private Sub txtRetype_KeyPress(KeyAscii As Integer)

End Sub

Private Sub txtUsername_KeyPress(KeyAscii As Integer)
End Sub
