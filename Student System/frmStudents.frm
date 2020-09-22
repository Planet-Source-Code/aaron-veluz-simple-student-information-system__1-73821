VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmStudents 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Student List"
   ClientHeight    =   6030
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8415
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmStudents.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6030
   ScaleWidth      =   8415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraStudents 
      Height          =   5295
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   8175
      Begin VB.Frame fraSearch 
         Caption         =   "Search Student"
         Height          =   975
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   7935
         Begin VB.ComboBox cboCivilStatus 
            Height          =   315
            ItemData        =   "frmStudents.frx":1082
            Left            =   3360
            List            =   "frmStudents.frx":1098
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   480
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.ComboBox cboCategory 
            Height          =   315
            ItemData        =   "frmStudents.frx":10DA
            Left            =   1080
            List            =   "frmStudents.frx":1102
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   480
            Width           =   2055
         End
         Begin VB.TextBox txtSearch 
            Height          =   375
            Left            =   3360
            MaxLength       =   255
            TabIndex        =   7
            Top             =   480
            Visible         =   0   'False
            Width           =   2535
         End
         Begin VB.ComboBox cboGender 
            Height          =   315
            ItemData        =   "frmStudents.frx":118C
            Left            =   3360
            List            =   "frmStudents.frx":1196
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   480
            Visible         =   0   'False
            Width           =   1695
         End
         Begin MSComCtl2.DTPicker BirthDate 
            Height          =   375
            Left            =   3360
            TabIndex        =   11
            Top             =   480
            Visible         =   0   'False
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   661
            _Version        =   393216
            Format          =   16908289
            CurrentDate     =   40575
         End
         Begin VB.Label lblCategory 
            Caption         =   "Category:"
            Height          =   255
            Left            =   120
            TabIndex        =   9
            Top             =   480
            Width           =   1215
         End
      End
      Begin VB.CheckBox chkSelect 
         Caption         =   "Select All"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   4920
         Width           =   1095
      End
      Begin MSComctlLib.ListView lvStudents 
         Height          =   3495
         Left            =   120
         TabIndex        =   2
         ToolTipText     =   "Double-click an item to view/edit details"
         Top             =   1320
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   6165
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
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Student ID"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Name"
            Object.Width           =   5821
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Section"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Student Number"
            Object.Width           =   2999
         EndProperty
      End
      Begin VB.Label lblRecord 
         Alignment       =   1  'Right Justify
         Caption         =   "---"
         Height          =   255
         Left            =   1320
         TabIndex        =   4
         Top             =   4920
         Width           =   6735
      End
   End
   Begin MSComctlLib.Toolbar tbrStudents 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   8415
      _ExtentX        =   14843
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
            ImageIndex      =   2
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
         Left            =   6600
         Top             =   120
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
               Picture         =   "frmStudents.frx":11A8
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmStudents.frx":223A
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmStudents.frx":32CC
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmStudents.frx":435E
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "frmStudents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub BirthDate_Change()
Call CheckBirthdate
End Sub

Private Sub BirthDate_Click()
Call CheckBirthdate
End Sub

Private Sub BirthDate_DblClick()
Call CheckBirthdate
End Sub

Private Sub BirthDate_DropDown()
Call CheckBirthdate
End Sub

Private Sub cboCategory_Click()
If cboCategory.Text = "ALL RECORDS" Then
    Call LoadStudents(lvStudents)
    Call CheckCount(lvStudents, lblRecord)
    cboGender.Visible = False
    txtSearch.Visible = False
    cboCivilStatus.Visible = False
    BirthDate.Visible = False
ElseIf cboCategory.Text = "Birthdate" Then
    BirthDate.Visible = True
    cboGender.Visible = True
    txtSearch.Visible = False
    cboCivilStatus.Visible = False
    cboGender.ListIndex = 0
ElseIf cboCategory.Text = "Gender" Then
    BirthDate.Visible = False
    cboGender.Visible = True
    txtSearch.Visible = False
    cboCivilStatus.Visible = False
    cboGender.ListIndex = 0
ElseIf cboCategory.Text = "Civil Status" Then
    BirthDate.Visible = False
    cboCivilStatus.Visible = True
    cboGender.Visible = True
    txtSearch.Visible = False
    cboCivilStatus.ListIndex = 0
Else
    cboGender.Visible = False
    BirthDate.Visible = False
    txtSearch.Visible = True
    cboCivilStatus.Visible = False
End If
End Sub

Private Sub cboCivilStatus_Click()
'If lvStudents.ListItems.Count = 0 Then Exit Sub
Set rs = con.Execute("SELECT * FROM Students WHERE CivilStatus='" & cboCivilStatus.Text & "'ORDER BY StudentID")
lvStudents.ListItems.Clear

For xCount = 1 To rs.RecordCount
    With ls
        Set ls = lvStudents.ListItems.Add(, , rs!StudentID)
        ls.SubItems(1) = rs!FirstName & " " & rs!MiddleName & " " & rs!LastName
        ls.SubItems(2) = rs!SectionName
        ls.SubItems(3) = rs!StudentNumber
        rs.MoveNext
    End With
Next xCount
Call CheckCount(lvStudents, lblRecord)
End Sub

Private Sub cboGender_Click()
'If lvStudents.ListItems.Count = 0 Then Exit Sub
Set rs = con.Execute("SELECT * FROM Students WHERE Gender='" & cboGender.Text & "'ORDER BY StudentID")
lvStudents.ListItems.Clear

For xCount = 1 To rs.RecordCount
    With ls
        Set ls = lvStudents.ListItems.Add(, , rs!StudentID)
        ls.SubItems(1) = rs!FirstName & " " & rs!MiddleName & " " & rs!LastName
        ls.SubItems(2) = rs!SectionName
        ls.SubItems(3) = rs!StudentNumber
        rs.MoveNext
    End With
Next xCount
Call CheckCount(lvStudents, lblRecord)
End Sub

Private Sub chkSelect_Click()
Call SelectAll(lvStudents, chkSelect)
End Sub



Private Sub cmdShowAll_Click()

End Sub

Private Sub Form_Activate()
If con.State = 0 Then Call konek
End Sub

Private Sub Form_Load()
Call EnableControls
cboCategory.ListIndex = 11
Call cboCategory_Click
Call LoadStudents(lvStudents)
Call CheckCount(lvStudents, lblRecord)
End Sub



Private Sub lvStudents_Click()
Call CheckCount(lvStudents, lblRecord)
End Sub

Private Sub lvStudents_DblClick()
If lvStudents.ListItems.Count = 0 Then Exit Sub

Set rs = con.Execute("SELECT * FROM Students WHERE StudentID=" & IIf((lvStudents.SelectedItem = 0), 1, lvStudents.SelectedItem) & "")
With frmNewStudent
    .lblRecordID.Caption = rs!StudentID
    .txtLastName.Text = rs!LastName
    .txtMiddleName.Text = rs!MiddleName
    .txtFirstName.Text = rs!FirstName
    .txtStudentNumber.Text = rs!StudentNumber
    If rs!Gender = "Male" Then
        .optMale.Value = True
    Else
        .optFemale.Value = True
    End If
    .txtSection.Text = rs!SectionName
    .txtContactNum.Text = rs!ContactNumber
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
    .txtEmailAddress.Text = IIf(IsNull(rs!EmailAddress), "", rs!EmailAddress)
    .txtContactNum2.Text = rs!ContactPersonNumber
    .txtContactPerson.Text = rs!ContactPerson
    .dtpBirthDate.Value = rs!BirthDate
    .cboCivilStatus.Text = rs!CivilStatus
    .txtReligion.Text = rs!Religion
    Age = rs!Age
    .Caption = "Edit Student - " & lvStudents.SelectedItem.SubItems(1)
    .cmdAdd.Caption = "Update"
    .Show vbModal, Me
End With
Call LoadStudents(lvStudents)
Call CheckCount(lvStudents, lblRecord)
End Sub

Private Sub lvStudents_KeyDown(KeyCode As Integer, Shift As Integer)
Call CheckCount(lvStudents, lblRecord)
End Sub

Private Sub lvStudents_KeyUp(KeyCode As Integer, Shift As Integer)
Call CheckCount(lvStudents, lblRecord)
End Sub

Private Sub optLastName_Click()

End Sub

Private Sub optStudentNumber_Click()

End Sub

Private Sub tbrStudents_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
    Case 1 'Add
        Set rs = con.Execute("SELECT * FROM Students ORDER BY StudentID")
        If rs.RecordCount = 0 Then
            xCount = 1
        Else
            rs.MoveLast
            xCount = Val(rs!StudentID) + 1
        End If
        With frmNewStudent
            .lblRecordID.Caption = xCount
            .cboCivilStatus.ListIndex = 0
            .Show vbModal, Me
            .cmdAdd.Caption = "New Student"
        End With
        Call LoadStudents(lvStudents)
        Call CheckCount(lvStudents, lblRecord)
    Case 3 'Delete
        If lvStudents.ListItems.Count = 0 Then Exit Sub
        Call CountSelected(lvStudents)
        If yCount <> 0 Then
            If MsgBox("Are you sure you want to delete the selected item(s)?", vbYesNo + vbExclamation, "Confirm Delete") = vbNo Then Exit Sub
            For xCount = 1 To lvStudents.ListItems.Count
                If lvStudents.ListItems(xCount).Checked = True Then
                    CurrRec = lvStudents.ListItems(xCount)
                    con.Execute ("DELETE FROM Students WHERE StudentID = " & CurrRec & "")
                End If
            Next xCount
            Call LoadStudents(lvStudents)
            Call CheckCount(lvStudents, lblRecord)
        End If
    Case 5 'Print
        If lvStudents.ListItems.Count = 0 Then Exit Sub
        Call CountSelected(lvStudents)
        If yCount = 1 Then
            For xCount = 1 To lvStudents.ListItems.Count
                If lvStudents.ListItems(xCount).Checked = True Then
                    CurrRec = lvStudents.ListItems(xCount)
                    Set rs = con.Execute("SELECT * FROM Students WHERE StudentID=" & CurrRec & "")
                End If
            Next xCount
            With rptStudent
                If fs.FileExists(App.Path & "\Images\" & rs!ImagePath) = True Then Set .Sections(3).Controls("img").Picture = LoadPicture(App.Path & "\Images\" & rs!ImagePath)
                Set .DataSource = rs
                .Caption = "Student Info: " & rs!FirstName & " " & rs!MiddleName & " " & rs!LastName
                .Show vbModal, Me
            End With
        ElseIf yCount > 1 And yCount <> lvStudents.ListItems.Count Then
            inPart = ""
            inPart2 = ""
            inWhole = ""
            For xCount = 1 To lvStudents.ListItems.Count
                If lvStudents.ListItems(xCount).Checked = True Then
                    inPart = lvStudents.ListItems(xCount) & ", "
                    inPart2 = inPart2 & inPart
                End If
            inWhole = "IN ( " & inPart2 & ")"
            Next xCount
            Set rs = con.Execute("SELECT StudentID,SectionName,StudentNumber,LastName+', '+FirstName+' '+MiddleName AS FullName FROM Students WHERE StudentID " & inWhole)
            With rptStudents
                Set .DataSource = rs
                .Orientation = rptOrientLandscape
                .Show vbModal, Me
            End With
        Else
            Set rs = con.Execute("SELECT StudentID,SectionName,StudentNumber,LastName+', '+FirstName+' '+MiddleName AS FullName FROM Students ORDER BY StudentID")
            With rptStudents
                Set .DataSource = rs
                .Orientation = rptOrientLandscape
                .Show vbModal, Me
            End With
        End If
    Case 7 'Show All
        cboCategory.ListIndex = 11
        Call cboCategory_Click
End Select
End Sub

Private Sub txtSearch_Change()

Select Case cboCategory.Text
    Case "Student ID"
        strTest = "StudentID"
    Case "Last Name"
        strTest = "LastName"
    Case "First Name"
        strTest = "FirstName"
    Case "Middle Name"
        strTest = "MiddleName"
    Case "Student Number"
    Call CheckCount(lvStudents, lblRecord)
        strTest = "StudentNumber"
    Case "Section"
        strTest = "Section"
    Case "Religion"
        strTest = "Religion"
    Case "Age"
        strTest = "Age"
End Select
If txtSearch.Text <> "" Then
    If cboCategory.Text = "Student ID" Then
        Set rs = con.Execute("SELECT * FROM Students WHERE StudentID=" & Val(txtSearch.Text) & "  ORDER BY StudentID")
    ElseIf cboCategory.Text = "Age" Then
        Set rs = con.Execute("SELECT * FROM Students WHERE Age=" & Val(txtSearch.Text) & "  ORDER BY StudentID")
    Else
        Set rs = con.Execute("SELECT * FROM Students WHERE " & strTest & " LIKE '" & txtSearch.Text & "%'  ORDER BY StudentID")
    End If
    lvStudents.ListItems.Clear
    
    For xCount = 1 To rs.RecordCount
        With ls
            Set ls = lvStudents.ListItems.Add(, , rs!StudentID)
            ls.SubItems(1) = rs!FirstName & " " & rs!MiddleName & " " & rs!LastName
            ls.SubItems(2) = rs!SectionName
            ls.SubItems(3) = rs!StudentNumber
            rs.MoveNext
        End With
    Next xCount
    Call CheckCount(lvStudents, lblRecord)
Else
    Call LoadStudents(lvStudents)
End If
End Sub
Public Sub CheckBirthdate()
Set rs = con.Execute("SELECT * FROM Students WHERE Birthdate=#" & BirthDate.Value & "# ORDER BY StudentID")
lvStudents.ListItems.Clear

For xCount = 1 To rs.RecordCount
    With ls
        Set ls = lvStudents.ListItems.Add(, , rs!StudentID)
        ls.SubItems(1) = rs!FirstName & " " & rs!MiddleName & " " & rs!LastName
        ls.SubItems(2) = rs!SectionName
        ls.SubItems(3) = rs!StudentNumber
        rs.MoveNext
    End With
Next xCount
Call CheckCount(lvStudents, lblRecord)
End Sub
