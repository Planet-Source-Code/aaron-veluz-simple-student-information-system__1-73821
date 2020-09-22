VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmStaff 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Staff List"
   ClientHeight    =   5550
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8790
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmStaff.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5550
   ScaleWidth      =   8790
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraStaff 
      Height          =   4815
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   8535
      Begin VB.Frame fraSearch 
         Caption         =   "Search Staff"
         Height          =   975
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   8295
         Begin VB.ComboBox cboCivilStatus 
            Height          =   315
            ItemData        =   "frmStaff.frx":1082
            Left            =   3360
            List            =   "frmStaff.frx":1098
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   480
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.ComboBox cboGender 
            Height          =   315
            ItemData        =   "frmStaff.frx":10DA
            Left            =   3360
            List            =   "frmStaff.frx":10E4
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   480
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.TextBox txtSearch 
            Height          =   375
            Left            =   3360
            MaxLength       =   255
            TabIndex        =   8
            Top             =   480
            Visible         =   0   'False
            Width           =   2775
         End
         Begin VB.ComboBox cboCategory 
            Height          =   315
            ItemData        =   "frmStaff.frx":10F6
            Left            =   1080
            List            =   "frmStaff.frx":1115
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   480
            Width           =   2055
         End
         Begin VB.Label lblCategory 
            Caption         =   "Category:"
            Height          =   255
            Left            =   120
            TabIndex        =   7
            Top             =   480
            Width           =   1215
         End
      End
      Begin VB.CheckBox chkSelect 
         Caption         =   "Select All"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   4440
         Width           =   1095
      End
      Begin MSComctlLib.ListView lvStaff 
         Height          =   2895
         Left            =   120
         TabIndex        =   2
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
            Text            =   "Staff ID"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Name"
            Object.Width           =   7056
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Designation"
            Object.Width           =   5768
         EndProperty
      End
      Begin VB.Label lblRecord 
         Alignment       =   1  'Right Justify
         Caption         =   "---"
         Height          =   255
         Left            =   1320
         TabIndex        =   4
         Top             =   4440
         Width           =   7095
      End
   End
   Begin MSComctlLib.Toolbar tbrStaff 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   8790
      _ExtentX        =   15505
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
               Picture         =   "frmStaff.frx":1181
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmStaff.frx":2213
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmStaff.frx":32A5
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmStaff.frx":4337
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "frmStaff"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cboCategory_Click()
If cboCategory.Text = "ALL RECORDS" Then
    Call LoadStaff(lvStaff)
    Call CheckCount(lvStaff, lblRecord)
    cboGender.Visible = False
    txtSearch.Visible = False
    cboCivilStatus.Visible = False
ElseIf cboCategory.Text = "Gender" Then
    cboGender.Visible = True
    txtSearch.Visible = False
    cboCivilStatus.Visible = False
    cboGender.ListIndex = 0
ElseIf cboCategory.Text = "Civil Status" Then
    cboGender.Visible = True
    txtSearch.Visible = False
    cboCivilStatus.Visible = True
    cboCivilStatus.ListIndex = 0
Else
    cboGender.Visible = False
    txtSearch.Visible = True
    cboCivilStatus.Visible = False
End If
End Sub

Private Sub cboCivilStatus_Click()
If lvStaff.ListItems.Count = 0 Then Exit Sub
Set rs = con.Execute("SELECT StaffID, FirstName+' '+MiddleName+' '+LastName AS FullName, Designation, Gender FROM Staff WHERE CivilStatus='" & cboCivilStatus.Text & "' ORDER BY StaffID")
lvStaff.ListItems.Clear

For xCount = 1 To rs.RecordCount
    With ls
        Set ls = lvStaff.ListItems.Add(, , rs!StaffID)
        ls.SubItems(1) = rs!FullName
        ls.SubItems(2) = rs!Designation
        rs.MoveNext
    End With
Next xCount
Call CheckCount(lvStaff, lblRecord)
End Sub

Private Sub cboGender_Click()
If lvStaff.ListItems.Count = 0 Then Exit Sub
Set rs = con.Execute("SELECT StaffID, FirstName+' '+MiddleName+' '+LastName AS FullName, Designation, Gender FROM Staff WHERE Gender='" & cboGender.Text & "' ORDER BY StaffID")
lvStaff.ListItems.Clear

For xCount = 1 To rs.RecordCount
    With ls
        Set ls = lvStaff.ListItems.Add(, , rs!StaffID)
        ls.SubItems(1) = rs!FullName
        ls.SubItems(2) = rs!Designation
        rs.MoveNext
    End With
Next xCount
Call CheckCount(lvStaff, lblRecord)
End Sub

Private Sub chkSelect_Click()
Call SelectAll(lvStaff, chkSelect)
End Sub

Private Sub cmdShoAll_Click()

End Sub

Private Sub Form_Activate()
If con.State = 0 Then Call konek
End Sub

Private Sub Form_Load()
Call EnableControls
Call LoadStaff(lvStaff)
Call CheckCount(lvStaff, lblRecord)
cboCategory.ListIndex = 8
Call cboCategory_Click
End Sub

Private Sub lvStaff_Click()
Call CheckCount(lvStaff, lblRecord)
End Sub

Private Sub lvStaff_DblClick()
If lvStaff.ListItems.Count = 0 Then Exit Sub
Set rs = con.Execute("SELECT * FROM Staff WHERE StaffID=" & Val(IIf((lvStaff.SelectedItem = 0), 1, lvStaff.SelectedItem)) & "")
With frmNewStaff
    .lblRecordID.Caption = rs!StaffID
    .txtLastName.Text = rs!LastName
    .txtMiddleName.Text = rs!MiddleName
    .txtFirstName.Text = rs!FirstName
    If rs!Gender = "Male" Then
        .optMale.Value = True
    Else
        .optFemale.Value = True
    End If
    .txtDesignation.Text = rs!Designation
    .txtContactNum.Text = rs!ContactNumber
    .cboCivilStatus.Text = rs!CivilStatus
    .txtReligion.Text = rs!Religion
    .dtpBirthDate.Value = rs!BirthDate
    Age = rs!Age
    .txtEmailAddress.Text = IIf(IsNull(rs!EmailAddress), "", rs!EmailAddress)
    .txtAddress.Text = rs!Address
    If rs!ImagePath <> Null Or rs!ImagePath <> "" Then
        strTest = App.Path & "\Images\" & rs!ImagePath
        If fs.FileExists(strTest) = True Then
            Call loadLogo(strTest, .img, .pic)
            FullPath = rs!ImagePath
        Else
            MsgBox "Image file not found.", vbOKOnly + vbExclamation, "Not Found"
        End If
    End If
    .Caption = "Edit Staff"
    .cmdAdd.Caption = "Update"
    .Show vbModal, Me
End With
Call LoadStaff(lvStaff)
Call CheckCount(lvStaff, lblRecord)
End Sub

Private Sub lvStaff_KeyDown(KeyCode As Integer, Shift As Integer)
Call CheckCount(lvStaff, lblRecord)
End Sub

Private Sub lvStaff_KeyUp(KeyCode As Integer, Shift As Integer)
Call CheckCount(lvStaff, lblRecord)
End Sub

Private Sub tbrStaff_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
    Case 1 'Add
        Set rs = con.Execute("SELECT * FROM Staff ORDER BY StaffID")
        If rs.RecordCount = 0 Then
            xCount = 1
        Else
            rs.MoveLast
            xCount = Val(rs!StaffID) + 1
        End If
        With frmNewStaff
            .Caption = "New Staff"
            .cmdAdd.Caption = "Add Staff"
            .cboCivilStatus.ListIndex = 0
            .lblRecordID.Caption = xCount
            .Show vbModal, Me
        End With
        Call LoadStaff(lvStaff)
        Call CheckCount(lvStaff, lblRecord)
    Case 3 'Delete
        If lvStaff.ListItems.Count = 0 Then Exit Sub
        Call CountSelected(lvStaff)
        If yCount <> 0 Then
            If MsgBox("Are you sure you want to delete the selected item(s)?", vbYesNo + vbExclamation, "Confirm Delete") = vbNo Then Exit Sub
            For xCount = 1 To lvStaff.ListItems.Count
                If lvStaff.ListItems(xCount).Checked = True Then
                    CurrRec = lvStaff.ListItems(xCount)
                    con.Execute ("DELETE FROM Staff WHERE StaffID = " & CurrRec & "")
                End If
            Next xCount
            Call LoadStaff(lvStaff)
            Call CheckCount(lvStaff, lblRecord)
        End If
    Case 5 'Print
        If lvStaff.ListItems.Count = 0 Then Exit Sub
        Call CountSelected(lvStaff)
        If yCount = 1 Then
            For xCount = 1 To lvStaff.ListItems.Count
                If lvStaff.ListItems(xCount).Checked = True Then
                    CurrRec = lvStaff.ListItems(xCount)
                    Set rs = con.Execute("SELECT * FROM Staff WHERE StaffID=" & CurrRec & "")
                End If
            Next xCount
            With rptSpecificStaff
                If fs.FileExists(App.Path & "\Images\" & rs!ImagePath) = True Then Set .Sections(3).Controls("img").Picture = LoadPicture(App.Path & "\Images\" & rs!ImagePath)
                Set .DataSource = rs
                .Caption = "Staff Info: " & rs!FirstName & " " & rs!MiddleName & " " & rs!LastName
                .Show vbModal, Me
            End With
        ElseIf yCount > 1 And yCount <> lvStaff.ListItems.Count Then
            inPart = ""
            inPart2 = ""
            inWhole = ""
            For xCount = 1 To lvStaff.ListItems.Count
                If lvStaff.ListItems(xCount).Checked = True Then
                    inPart = lvStaff.ListItems(xCount) & ", "
                    inPart2 = inPart2 & inPart
                End If
            inWhole = "IN ( " & inPart2 & ")"
            Next xCount
            Set rs = con.Execute("SELECT StaffID, FirstName+' '+MiddleName+' '+LastName AS FullName, Designation FROM Staff WHERE StaffID " & inWhole)
            With rptStaff
                Set .DataSource = rs
                .Show vbModal, Me
            End With
        Else
            Set rs = con.Execute("SELECT StaffID, FirstName+' '+MiddleName+' '+LastName AS FullName, Designation FROM Staff ORDER BY StaffID")
            With rptStaff
                Set .DataSource = rs
                .Show vbModal, Me
            End With
        End If
    Case 7 'Show All
        cboCategory.ListIndex = 8
        Call cboCategory_Click
End Select

End Sub

Private Sub txtSearch_Change()

Select Case cboCategory.Text
    Case "Staff ID"
        strTest = "StaffID"
    Case "Last Name"
        strTest = "LastName"
    Case "First Name"
        strTest = "FirstName"
    Case "Middle Name"
        strTest = "MiddleName"
    Case "Designation"
        strTest = "Designation"
    Case "Religion"
        strTest = "Religion"
End Select
If txtSearch.Text <> "" Then
    If cboCategory.Text <> "Staff ID" Then
        Set rs = con.Execute("SELECT StaffID, FirstName+' '+MiddleName+' '+LastName AS FullName, Designation FROM Staff WHERE " & strTest & " LIKE '" & txtSearch.Text & "%' ORDER BY StaffID")
    Else
        Set rs = con.Execute("SELECT StaffID, FirstName+' '+MiddleName+' '+LastName AS FullName, Designation FROM Staff WHERE StaffID=" & Val(txtSearch.Text) & " ORDER BY StaffID")
    End If
    lvStaff.ListItems.Clear
    
    For xCount = 1 To rs.RecordCount
        With ls
            Set ls = lvStaff.ListItems.Add(, , rs!StaffID)
            ls.SubItems(1) = rs!FullName
            ls.SubItems(2) = rs!Designation
            rs.MoveNext
        End With
    Next xCount
    Call CheckCount(lvStaff, lblRecord)
Else
    Call LoadStaff(lvStaff)
    Call CheckCount(lvStaff, lblRecord)
End If
End Sub

