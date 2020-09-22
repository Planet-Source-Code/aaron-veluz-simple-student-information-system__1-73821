VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmContactGroups 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Contact Groups"
   ClientHeight    =   5550
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8895
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmContactGroups.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5550
   ScaleWidth      =   8895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraStaff 
      Height          =   4815
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   8655
      Begin VB.Frame fraSearch 
         Caption         =   "Search Contact Groups"
         Height          =   975
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   8295
         Begin VB.TextBox txtSearch 
            Height          =   375
            Left            =   3360
            MaxLength       =   255
            TabIndex        =   5
            Top             =   480
            Visible         =   0   'False
            Width           =   2775
         End
         Begin VB.ComboBox cboCategory 
            Height          =   315
            ItemData        =   "frmContactGroups.frx":1082
            Left            =   1080
            List            =   "frmContactGroups.frx":1092
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   480
            Width           =   2055
         End
         Begin VB.Label lblCategory 
            Caption         =   "Category:"
            Height          =   255
            Left            =   120
            TabIndex        =   6
            Top             =   480
            Width           =   1215
         End
      End
      Begin VB.CheckBox chkSelect 
         Caption         =   "Select All"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   4440
         Width           =   1095
      End
      Begin MSComctlLib.ListView lvContactGroups 
         Height          =   2895
         Left            =   120
         TabIndex        =   7
         ToolTipText     =   "Double-click an item to view/edit details"
         Top             =   1320
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   5106
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
            Text            =   "Contact ID"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Name"
            Object.Width           =   7056
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Contact Group"
            Object.Width           =   5768
         EndProperty
      End
      Begin VB.Label lblRecord 
         Alignment       =   1  'Right Justify
         Caption         =   "---"
         Height          =   255
         Left            =   1320
         TabIndex        =   8
         Top             =   4440
         Width           =   7095
      End
   End
   Begin MSComctlLib.Toolbar tbrContactGroups 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8895
      _ExtentX        =   15690
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
         Left            =   5760
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
               Picture         =   "frmContactGroups.frx":10C0
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmContactGroups.frx":2152
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmContactGroups.frx":31E4
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmContactGroups.frx":4276
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "frmContactGroups"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cboCategory_Click()
If cboCategory.Text = "ALL RECORDS" Then
    Call LoadContactGroups(lvContactGroups)
    Call CheckCount(lvContactGroups, lblRecord)
    txtSearch.Visible = False
Else
    txtSearch.Visible = True
End If
End Sub

Private Sub chkSelect_Click()
Call SelectAll(lvContactGroups, chkSelect)
End Sub

Private Sub Form_Activate()
If con.State = 0 Then Call konek
End Sub

Private Sub Form_Load()
Call LoadContactGroups(lvContactGroups)
Call CheckCount(lvContactGroups, lblRecord)
cboCategory.ListIndex = 3
Call cboCategory_Click
End Sub

Private Sub fraStaff_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub lvContactGroups_Click()
Call CheckCount(lvContactGroups, lblRecord)
End Sub

Private Sub lvContactGroups_DblClick()
If lvContactGroups.ListItems.Count = 0 Then Exit Sub
Set rs = con.Execute("SELECT * FROM ContactGroups WHERE LUserID=" & UserID & " AND GroupID=" & IIf((lvContactGroups.SelectedItem = 0), 1, lvContactGroups.SelectedItem) & "")
With frmNewContactGroup
    .lblRecordID.Caption = rs!GroupID
    .txtGroupName.Text = rs!GroupName
    .txtNotes.Text = rs!Notes
    .Caption = "Edit Groups - " & lvContactGroups.SelectedItem.SubItems(1)
    .cmdAdd.Caption = "Update"
    .Show vbModal, Me
End With
Call LoadContactGroups(lvContactGroups)
Call CheckCount(lvContactGroups, lblRecord)
End Sub

Private Sub lvContactGroups_KeyDown(KeyCode As Integer, Shift As Integer)
Call CheckCount(lvContactGroups, lblRecord)
End Sub

Private Sub lvContactGroups_KeyUp(KeyCode As Integer, Shift As Integer)
Call CheckCount(lvContactGroups, lblRecord)
End Sub

Private Sub tbrStaff_ButtonClick(ByVal Button As MSComctlLib.Button)
End Sub

Private Sub tbrContactGroups_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
    Case 1 'Add
        Set rs = con.Execute("SELECT * FROM ContactGroups WHERE LUserID=" & UserID & " ORDER BY GroupID")
        If rs.RecordCount = 0 Then
            xCount = 1
        Else
            rs.MoveLast
            xCount = Val(rs!GroupID) + 1
        End If
        With frmNewContactGroup
            .Caption = "Add Contact Group"
            .lblRecordID.Caption = xCount
            .cmdAdd.Caption = "Add Group"
            .Show vbModal, Me
        End With
        Call LoadContactGroups(lvContactGroups)
        Call CheckCount(lvContactGroups, lblRecord)
    Case 3 'Delete
        If lvContactGroups.ListItems.Count = 0 Then Exit Sub
        Call CountSelected(lvContactGroups)
        If yCount <> 0 Then
            If MsgBox("Are you sure you want to delete the selected item(s)?", vbYesNo + vbExclamation, "Confirm Delete") = vbNo Then Exit Sub
            For xCount = 1 To lvContactGroups.ListItems.Count
                If lvContactGroups.ListItems(xCount).Checked = True Then
                        CurrRec = lvContactGroups.ListItems(xCount)
                        con.Execute ("DELETE FROM ContactGroups WHERE LUserID=" & UserID & " AND GroupID = " & CurrRec & "")
                End If
            Next xCount
            Call LoadContactGroups(lvContactGroups)
            Call CheckCount(lvContactGroups, lblRecord)
        End If
    Case 5 'Print''
        If lvContactGroups.ListItems.Count = 0 Then Exit Sub
        Call CountSelected(lvContactGroups)
        If yCount >= 1 And yCount <> lvContactGroups.ListItems.Count Then
            inPart = ""
            inPart2 = ""
            inWhole = ""
            For xCount = 1 To lvContactGroups.ListItems.Count
                If lvContactGroups.ListItems(xCount).Checked = True Then
                    inPart = lvContactGroups.ListItems(xCount) & ", "
                    inPart2 = inPart2 & inPart
                End If
            inWhole = "IN ( " & inPart2 & ")"
            Next xCount
            Set rs = con.Execute("SELECT * FROM ContactGroups WHERE LUserID=" & UserID & " AND GroupID " & inWhole)
            With rptContactGroups
                Set .DataSource = rs
                .Show vbModal, Me
            End With
        Else
            Set rs = con.Execute("SELECT * FROM ContactGroups WHERE LUserID=" & UserID & " ORDER BY GroupID")
            With rptContactGroups
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
    Case "Group ID"
        strTest = "GroupID"
    Case "Group Name"
        strTest = "GroupName"
    Case "Notes"
        strTest = "Notes"
End Select
If txtSearch.Text <> "" Then
    If cboCategory.Text <> "Group ID" Then
        Set rs = con.Execute("SELECT * FROM ContactGroups WHERE LUserID=" & UserID & " AND " & strTest & " LIKE '" & txtSearch.Text & "%' ORDER BY GroupID")
    Else
        Set rs = con.Execute("SELECT * FROM ContactGroups WHERE LUserID=" & UserID & " AND GroupID=" & Val(txtSearch.Text) & " ORDER BY GroupID")
    End If
    lvContactGroups.ListItems.Clear
    
    For xCount = 1 To rs.RecordCount
        With ls
            Set ls = lvContactGroups.ListItems.Add(, , rs!GroupID)
            ls.SubItems(1) = rs!GroupName
            ls.SubItems(2) = rs!Notes
            rs.MoveNext
        End With
    Next xCount
    Call CheckCount(lvContactGroups, lblRecord)
Else
    Call LoadContactGroups(lvContactGroups)
    Call CheckCount(lvContactGroups, lblRecord)
End If
End Sub
