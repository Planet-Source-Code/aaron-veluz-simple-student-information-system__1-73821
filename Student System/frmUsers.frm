VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmUsers 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "User List"
   ClientHeight    =   6510
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6255
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmUsers.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6510
   ScaleWidth      =   6255
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraUsers 
      Height          =   5775
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   6015
      Begin VB.Frame fraSearch 
         Caption         =   "Search User"
         Height          =   975
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   5775
         Begin VB.TextBox txtSearch 
            Height          =   375
            Left            =   2880
            MaxLength       =   255
            TabIndex        =   7
            Top             =   480
            Visible         =   0   'False
            Width           =   2655
         End
         Begin VB.ComboBox cboCategory 
            Height          =   315
            ItemData        =   "frmUsers.frx":1082
            Left            =   1200
            List            =   "frmUsers.frx":1092
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   480
            Width           =   1575
         End
         Begin VB.Label lblCategory 
            Caption         =   "Category:"
            Height          =   255
            Left            =   120
            TabIndex        =   8
            Top             =   480
            Width           =   975
         End
      End
      Begin VB.CheckBox chkSelect 
         Caption         =   "Select All"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   5400
         Width           =   1095
      End
      Begin MSComctlLib.ListView lvUsers 
         Height          =   3975
         Left            =   120
         TabIndex        =   2
         ToolTipText     =   "Double-click an item to view/edit details"
         Top             =   1320
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   7011
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FlatScrollBar   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "User ID"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Username"
            Object.Width           =   4498
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Privilege"
            Object.Width           =   3528
         EndProperty
      End
      Begin VB.Label lblRecord 
         Alignment       =   1  'Right Justify
         Caption         =   "---"
         Height          =   255
         Left            =   1320
         TabIndex        =   4
         Top             =   5400
         Width           =   4575
      End
   End
   Begin MSComctlLib.Toolbar tbrUsers 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   1005
      ButtonWidth     =   2275
      ButtonHeight    =   1005
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   8
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Add New"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Delete"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Print"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Show All"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
      EndProperty
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   5640
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   4
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmUsers.frx":10C1
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmUsers.frx":2153
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmUsers.frx":31E5
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmUsers.frx":4277
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "frmUsers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cboCategory_Click()
If cboCategory.Text = "ALL RECORDS" Then
    Call LoadUsers(lvUsers)
    Call CheckCount(lvUsers, lblRecord)
    txtSearch.Visible = False
Else
    txtSearch.Visible = True
End If
End Sub

Private Sub chkSelect_Click()
Call SelectAll(lvUsers, chkSelect)
End Sub



Private Sub cmdShowAll_Click()

End Sub

Private Sub Form_Activate()
If con.State = 0 Then Call konek
End Sub

Private Sub Form_Load()
Call EnableControls
cboCategory.ListIndex = 3
Call cboCategory_Click
Call LoadUsers(lvUsers)
Call CheckCount(lvUsers, lblRecord)
End Sub

Private Sub lvUsers_Click()
Call CheckCount(lvUsers, lblRecord)
End Sub

Private Sub lvUsers_DblClick()
If lvUsers.ListItems.Count = 0 Then Exit Sub

If Privilege <> "SuperAdministrator" And lvUsers.SelectedItem.SubItems(2) = "SuperAdministrator" Then
    MsgBox "You do not have sufficient rights to edit the default account.", vbOKOnly + vbCritical, "Access Denied"
    Exit Sub
ElseIf Privilege = "Staff" And (UserID <> Val(lvUsers.SelectedItem)) Then
    MsgBox "You do not have sufficient rights to edit an account other than yours.", vbOKOnly + vbCritical, "Access Denied"
    Exit Sub
Else
    Set rs = con.Execute("SELECT * FROM Users WHERE UserID=" & IIf((lvUsers.SelectedItem = 0), 1, lvUsers.SelectedItem) & "")
    With frmNewUser
        .lblRecordID.Caption = rs!UserID
        .txtUsername.Text = rs!Username
        If Privilege = "SuperAdministrator" Then
            .txtPassword.Text = rs!UPassword
            .txtPassword.PasswordChar = ""
            .txtRetype.Text = rs!UPassword
            .txtRetype.PasswordChar = ""
            .lblPrivilege.Visible = True
            .cboPrivilege.Visible = False
            .chkMask.Visible = True
            .chkMask.Value = 1
            .lblUserPrivilege.Visible = True
            .lblUserPrivilege.Caption = rs!Privilege
        Else
            .txtPassword.PasswordChar = "*"
            .txtPassword.Text = ""
            .txtRetype.PasswordChar = "*"
            .txtRetype.Text = ""
            .lblUserPrivilege.Visible = False
            .chkMask.Visible = False
        End If
        If rs!Privilege <> "SuperAdministrator" Then
            .cboPrivilege.Visible = True
            .lblPrivilege.Visible = True
            .cboPrivilege.Text = rs!Privilege
            .lblUserPrivilege.Visible = False
        End If
        .Caption = "Edit User - " & lvUsers.SelectedItem.SubItems(1)
        .cmdAdd.Caption = "Update"
        .Show vbModal, Me
    End With
    Call LoadUsers(lvUsers)
    Call CheckCount(lvUsers, lblRecord)
End If
End Sub

Private Sub lvUsers_KeyDown(KeyCode As Integer, Shift As Integer)
Call CheckCount(lvUsers, lblRecord)
End Sub

Private Sub lvUsers_KeyUp(KeyCode As Integer, Shift As Integer)
Call CheckCount(lvUsers, lblRecord)
End Sub

Private Sub tbrUsers_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
    Case 1 'Add
        Set rs = con.Execute("SELECT * FROM Users ORDER BY UserID")
        If rs.RecordCount = 0 Then
            xCount = 1
        Else
            rs.MoveLast
            xCount = Val(rs!UserID) + 1
        End If
        With frmNewUser
            .Caption = "New User"
            .lblRecordID.Caption = xCount
            .cboPrivilege.ListIndex = 0
            .cboPrivilege.Visible = True
            .lblPrivilege.Visible = True
            .cmdAdd.Caption = "Add User"
            .Show vbModal, Me
        End With
        Call LoadUsers(lvUsers)
        Call CheckCount(lvUsers, lblRecord)
    Case 3 'Delete
        If lvUsers.ListItems.Count = 0 Then Exit Sub
        Call CountSelected(lvUsers)
        If yCount <> 0 Then
            If MsgBox("Are you sure you want to delete the selected item(s)?", vbYesNo + vbExclamation, "Confirm Delete") = vbNo Then Exit Sub
            For xCount = 1 To lvUsers.ListItems.Count
                If lvUsers.ListItems(xCount).Checked = True Then
                    If lvUsers.ListItems(xCount).SubItems(2) = "SuperAdministrator" Then
                        If Privilege = "SuperAdministrator" Then
                            MsgBox "You cannot delete your own account.", vbOKOnly + vbCritical, "Invalid Operation"
                            Exit Sub
                        Else
                            MsgBox "You cannot delete a default user account.", vbOKOnly + vbCritical, "Invalid Operation"
                            Exit Sub
                        End If
                    ElseIf lvUsers.ListItems(xCount).SubItems(1) = Username Then
                        MsgBox "You cannot delete your own account.", vbOKOnly + vbCritical, "Invalid Operation"
                        Exit Sub
                    Else
                        CurrRec = lvUsers.ListItems(xCount)
                        con.Execute ("DELETE FROM Users WHERE UserID = " & CurrRec & "")
                    End If
                End If
            Next xCount
            Call LoadUsers(lvUsers)
            Call CheckCount(lvUsers, lblRecord)
        End If
    Case 5 'Print''
        If lvUsers.ListItems.Count = 0 Then Exit Sub
        Call CountSelected(lvUsers)
        If yCount = 1 Then
            For xCount = 1 To lvUsers.ListItems.Count
                If lvUsers.ListItems(xCount).Checked = True Then
                    CurrRec = lvUsers.ListItems(xCount)
                    Set rs = con.Execute("SELECT * FROM Users WHERE UserID = " & CurrRec & "")
                End If
            Next xCount
            With rptUsers
                Set .DataSource = rs
                .Show vbModal, Me
            End With
        ElseIf yCount > 1 And yCount <> lvUsers.ListItems.Count Then
            inPart = ""
            inPart2 = ""
            inWhole = ""
            For xCount = 1 To lvUsers.ListItems.Count
                If lvUsers.ListItems(xCount).Checked = True Then
                    inPart = lvUsers.ListItems(xCount) & ", "
                    inPart2 = inPart2 & inPart
                End If
            inWhole = "IN ( " & inPart2 & ")"
            Next xCount
            Set rs = con.Execute("SELECT * FROM Users WHERE UserID " & inWhole)
            With rptUsers
                Set .DataSource = rs
                .Show vbModal, Me
            End With
        Else
            Set rs = con.Execute("SELECT * FROM Users ORDER BY UserID")
            With rptUsers
                Set .DataSource = rs
                .Show vbModal, Me
            End With
        End If
    Case 7 'Show All
        cboCategory.ListIndex = 3
        Call cboCategory_Click
End Select
End Sub

Private Sub txtSearch_Change()

Select Case cboCategory.Text
    Case "User ID"
        strTest = "UserID"
    Case "Username"
        strTest = "Username"
    Case "Privilege"
        strTest = "Privilege"
End Select
If txtSearch.Text <> "" Then
    If cboCategory.Text <> "User ID" Then
        Set rs = con.Execute("SELECT UserID,Username, Privilege FROM Users WHERE " & strTest & " LIKE '" & txtSearch.Text & "%' ORDER BY UserID")
    Else
        Set rs = con.Execute("SELECT UserID,Username, Privilege FROM Users WHERE UserID=" & Val(txtSearch.Text) & " ORDER BY UserID")
    End If
    lvUsers.ListItems.Clear
    
    For xCount = 1 To rs.RecordCount
        With ls
            Set ls = lvUsers.ListItems.Add(, , rs!UserID)
            ls.SubItems(1) = rs!Username
            ls.SubItems(2) = rs!Privilege
            rs.MoveNext
        End With
    Next xCount
    Call CheckCount(lvUsers, lblRecord)
Else
    Call LoadUsers(lvUsers)
    Call CheckCount(lvUsers, lblRecord)
End If
End Sub


