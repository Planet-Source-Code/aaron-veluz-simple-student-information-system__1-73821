VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmContacts 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Contacts List"
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
   Icon            =   "frmContacts.frx":0000
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
      Begin VB.CheckBox chkSelect 
         Caption         =   "Select All"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   4440
         Width           =   1095
      End
      Begin VB.Frame fraSearch 
         Caption         =   "Search Contacts"
         Height          =   975
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   8295
         Begin VB.ComboBox cboContactGroup 
            Height          =   315
            ItemData        =   "frmContacts.frx":1082
            Left            =   3360
            List            =   "frmContacts.frx":1084
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   480
            Visible         =   0   'False
            Width           =   1935
         End
         Begin VB.ComboBox cboCategory 
            Height          =   315
            ItemData        =   "frmContacts.frx":1086
            Left            =   1080
            List            =   "frmContacts.frx":10C0
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   480
            Width           =   2055
         End
         Begin VB.TextBox txtSearch 
            Height          =   375
            Left            =   3360
            MaxLength       =   255
            TabIndex        =   5
            Top             =   480
            Visible         =   0   'False
            Width           =   2775
         End
         Begin VB.ComboBox cboGender 
            Height          =   315
            ItemData        =   "frmContacts.frx":11A4
            Left            =   3360
            List            =   "frmContacts.frx":11AE
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   480
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.ComboBox cboCivilStatus 
            Height          =   315
            ItemData        =   "frmContacts.frx":11C0
            Left            =   3360
            List            =   "frmContacts.frx":11D6
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   480
            Visible         =   0   'False
            Width           =   1695
         End
         Begin MSComCtl2.DTPicker BirthDate 
            Height          =   375
            Left            =   3360
            TabIndex        =   12
            Top             =   480
            Visible         =   0   'False
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   661
            _Version        =   393216
            Format          =   53805057
            CurrentDate     =   40575
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
      Begin MSComctlLib.ListView lvContacts 
         Height          =   2895
         Left            =   120
         TabIndex        =   9
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
            Text            =   "ID"
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
         TabIndex        =   10
         Top             =   4440
         Width           =   7095
      End
   End
   Begin MSComctlLib.Toolbar tbrContacts 
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
               Picture         =   "frmContacts.frx":1218
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmContacts.frx":22AA
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmContacts.frx":333C
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmContacts.frx":43CE
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "frmContacts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BirthDate_Change()
Call CheckDates
End Sub

Private Sub BirthDate_Click()
Call CheckDates
End Sub

Private Sub BirthDate_DblClick()
Call CheckDates
End Sub

Private Sub BirthDate_DropDown()
Call CheckDates
End Sub

Private Sub cboCategory_Click()
If cboCategory.Text = "ALL RECORDS" Then
    Call LoadContacts(lvContacts)
    Call CheckCount(lvContacts, lblRecord)
    cboGender.Visible = False
    txtSearch.Visible = False
    cboCivilStatus.Visible = False
    BirthDate.Visible = False
    cboContactGroup.Visible = False
ElseIf cboCategory.Text = "Birthdate" Then
    BirthDate.Visible = True
    cboGender.Visible = True
    txtSearch.Visible = False
    cboCivilStatus.Visible = False
    cboGender.ListIndex = 0
    cboContactGroup.Visible = False
ElseIf cboCategory.Text = "Gender" Then
    BirthDate.Visible = False
    cboGender.Visible = True
    txtSearch.Visible = False
    cboCivilStatus.Visible = False
    cboGender.ListIndex = 0
    cboContactGroup.Visible = False
ElseIf cboCategory.Text = "Civil Status" Then
    BirthDate.Visible = False
    cboCivilStatus.Visible = True
    cboGender.Visible = True
    txtSearch.Visible = False
    cboCivilStatus.ListIndex = 0
    cboContactGroup.Visible = False
ElseIf cboCategory.Text = "Group Name" Then
    BirthDate.Visible = False
    cboCivilStatus.Visible = False
    cboGender.Visible = True
    txtSearch.Visible = False
    Set rs = con.Execute("SELECT * FROM ContactGroups WHERE LUserID=" & UserID & " ORDER BY GroupID")
    For xCount = 1 To rs.RecordCount
        With rs
            cboContactGroup.AddItem !GroupName
            rs.MoveNext
        End With
    Next xCount
    cboContactGroup.Visible = True
    cboContactGroup.ListIndex = 0
Else
    cboGender.Visible = False
    BirthDate.Visible = False
    txtSearch.Visible = True
    cboCivilStatus.Visible = False
    cboContactGroup.Visible = False
End If
End Sub

Private Sub cboCivilStatus_Click()
Set rs = con.Execute("SELECT * FROM Contacts WHERE LUserID=" & UserID & " AND CivilStatus='" & cboCivilStatus.Text & "'ORDER BY ContactID")
lvContacts.ListItems.Clear

For xCount = 1 To rs.RecordCount
    With ls
        Set ls = lvContacts.ListItems.Add(, , rs!ContactID)
        ls.SubItems(1) = rs!FirstName & " " & rs!MiddleName & " " & rs!LastName
        ls.SubItems(2) = rs!GroupName
        rs.MoveNext
    End With
Next xCount
Call CheckCount(lvContacts, lblRecord)
End Sub

Private Sub cboContactGroup_Click()
Set rs = con.Execute("SELECT * FROM Contacts WHERE LUserID=" & UserID & " AND GroupName='" & cboContactGroup.Text & "'ORDER BY ContactID")
lvContacts.ListItems.Clear

For xCount = 1 To rs.RecordCount
    With ls
        Set ls = lvContacts.ListItems.Add(, , rs!ContactID)
        ls.SubItems(1) = rs!FirstName & " " & rs!MiddleName & " " & rs!LastName
        ls.SubItems(2) = rs!GroupName
        rs.MoveNext
    End With
Next xCount
Call CheckCount(lvContacts, lblRecord)
End Sub

Private Sub cboGender_Click()
Set rs = con.Execute("SELECT * FROM Contacts WHERE LUserID=" & UserID & "  AND Gender='" & cboGender.Text & "'ORDER BY ContactID")
lvContacts.ListItems.Clear

For xCount = 1 To rs.RecordCount
    With ls
        Set ls = lvContacts.ListItems.Add(, , rs!ContactID)
        ls.SubItems(1) = rs!FirstName & " " & rs!MiddleName & " " & rs!LastName
        ls.SubItems(2) = rs!GroupName
        rs.MoveNext
    End With
Next xCount
Call CheckCount(lvContacts, lblRecord)
End Sub

Private Sub chkSelect_Click()
Call SelectAll(lvContacts, chkSelect)
End Sub

Private Sub Form_Activate()
If con.State = 0 Then Call konek
End Sub

Private Sub Form_Load()
Call LoadContacts(lvContacts)
Call CheckCount(lvContacts, lblRecord)
cboCategory.ListIndex = 17
Call cboCategory_Click
End Sub

Private Sub lvContacts_Click()
Call CheckCount(lvContacts, lblRecord)
End Sub

Private Sub lvContacts_DblClick()
If lvContacts.ListItems.Count = 0 Then Exit Sub
Set rs = con.Execute("SELECT * FROM ContactGroups WHERE LUserID=" & UserID & " ORDER BY GroupID")
If rs.RecordCount <> 0 Then
    frmNewContact.cboContactGroup.Clear
    For xCount = 1 To rs.RecordCount
        With rs
            frmNewContact.cboContactGroup.AddItem !GroupName
            rs.MoveNext
        End With
    Next xCount
End If
Set rs = con.Execute("SELECT * FROM Contacts WHERE LUserID=" & UserID & " AND ContactID=" & IIf((lvContacts.SelectedItem = 0), 1, lvContacts.SelectedItem) & "")
With frmNewContact
    .lblRecordID.Caption = rs!ContactID
    .lblUserID.Caption = rs!LUserID
    .txtLastName.Text = rs!LastName
    .txtMiddleName.Text = rs!MiddleName
    .txtFirstName.Text = rs!FirstName
    .txtPositionDesignation = rs!PositionDesignation
    .txtCompanySchool.Text = rs!CompanySchool
    .txtPhoneNumber.Text = rs!PhoneNumber
    .txtMobileNumber.Text = rs!MobileNumber
    .txtFaxNumber.Text = rs!FaxNumber
    If rs!Gender = "Male" Then
        .optMale.Value = True
    Else
        .optFemale.Value = True
    End If
    .txtAddress.Text = IIf(IsNull(rs!Address), "", rs!Address)
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
    .dtpBirthDate.Value = rs!BirthDate
    .cboCivilStatus.Text = rs!CivilStatus
    .txtReligion.Text = rs!Religion
    Age = rs!Age
    .txtNotes.Text = rs!Notes
    .Caption = "Edit Contact - " & lvContacts.SelectedItem.SubItems(1)
    .cmdAdd.Caption = "Update"
    .cboContactGroup.Text = rs!GroupName
    .Show vbModal, Me
End With
Call LoadContacts(lvContacts)
Call CheckCount(lvContacts, lblRecord)
End Sub

Private Sub lvContacts_KeyDown(KeyCode As Integer, Shift As Integer)
Call CheckCount(lvContacts, lblRecord)
End Sub

Private Sub lvContacts_KeyUp(KeyCode As Integer, Shift As Integer)
Call CheckCount(lvContacts, lblRecord)
End Sub

Private Sub tbrStaff_ButtonClick(ByVal Button As MSComctlLib.Button)

End Sub
Public Sub CheckDates()
If lvContacts.ListItems.Count = 0 Then Exit Sub
Set rs = con.Execute("SELECT * FROM Contacts WHERE LUserID=" & UserID & "  AND BirthDate=#" & BirthDate.Value & "# ORDER BY ContactID")
lv.ListItems.Clear

For xCount = 1 To rs.RecordCount
    With ls
        Set ls = lv.ListItems.Add(, , rs!ContactID)
        ls.SubItems(1) = rs!rs!FirstName & " " & rs!MiddleName & " " & rs!LastName
        ls.SubItems(2) = rs!GroupName
        rs.MoveNext
    End With
Next xCount
End Sub

Private Sub tbrContacts_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
    Case 1 'Add
        Set rs = con.Execute("SELECT * FROM ContactGroups WHERE LUserID=" & UserID & " ORDER BY GroupID")
        If rs.RecordCount = 0 Then
            If MsgBox("No Contact Groups found. Would you like to add a new Contact Group?", vbYesNo + vbQuestion, "No Records") = vbNo Then
                Exit Sub
            Else
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
            End If
        Else
            frmNewContact.cboContactGroup.Clear
            For xCount = 1 To rs.RecordCount
                With rs
                    frmNewContact.cboContactGroup.AddItem !GroupName
                    rs.MoveNext
                End With
            Next xCount
            With frmNewContact
                Set rs = con.Execute("SELECT * FROM Contacts ORDER BY ContactID")
                If rs.RecordCount = 0 Then
                    .lblRecordID.Caption = 1
                    .lblUserID.Caption = 1
                Else
                    rs.MoveLast
                    .lblRecordID.Caption = rs!ContactID + 1
                    Set rs = con.Execute("SELECT * FROM Contacts WHERE LUserID=" & UserID & " ORDER BY ContactID")
                    If rs.RecordCount = 0 Then
                        .lblUserID.Caption = 1
                    Else
                        rs.MoveLast
                        .lblUserID.Caption = rs!LUserID + 1
                    End If
                End If
                .cboContactGroup.ListIndex = 0
                .Caption = "Add Contact"
                .Show vbModal, Me
            End With
        End If
        Call LoadContacts(lvContacts)
        Call CheckCount(lvContacts, lblRecord)
    Case 3 'Delete
        If lvContacts.ListItems.Count = 0 Then Exit Sub
        Call CountSelected(lvContacts)
        If yCount <> 0 Then
            If MsgBox("Are you sure you want to delete the selected item(s)?", vbYesNo + vbExclamation, "Confirm Delete") = vbNo Then Exit Sub
            For xCount = 1 To lvContacts.ListItems.Count
                If lvContacts.ListItems(xCount).Checked = True Then
                        CurrRec = lvContacts.ListItems(xCount)
                        con.Execute ("DELETE FROM Contacts WHERE LUserID=" & UserID & " AND ContactID = " & CurrRec & "")
                End If
            Next xCount
            Call LoadContacts(lvContacts)
            Call CheckCount(lvContacts, lblRecord)
        End If
    Case 5 'Print''
        If lvContacts.ListItems.Count = 0 Then Exit Sub
        Call CountSelected(lvContacts)
        If yCount = 1 Then
            For xCount = 1 To lvContacts.ListItems.Count
                If lvContacts.ListItems(xCount).Checked = True Then
                    CurrRec = lvContacts.ListItems(xCount)
                    Set rs = con.Execute("SELECT * FROM Contacts WHERE LUserID=" & UserID & " AND ContactID=" & CurrRec & "")
                End If
            Next xCount
            With rptSpecificContact
                If fs.FileExists(App.Path & "\Images\" & rs!ImagePath) = True Then Set .Sections(3).Controls("img").Picture = LoadPicture(App.Path & "\Images\" & rs!ImagePath)
                Set .DataSource = rs
                .Caption = "Contact Info: " & rs!FirstName & " " & rs!MiddleName & " " & rs!LastName
                .Show vbModal, Me
            End With
        ElseIf yCount > 1 And yCount <> lvContacts.ListItems.Count Then
            inPart = ""
            inPart2 = ""
            inWhole = ""
            For xCount = 1 To lvContacts.ListItems.Count
                If lvContacts.ListItems(xCount).Checked = True Then
                    inPart = lvContacts.ListItems(xCount) & ", "
                    inPart2 = inPart2 & inPart
                End If
            inWhole = "IN ( " & inPart2 & ")"
            Next xCount
            Set rs = con.Execute("SELECT ContactID,FirstName+' '+MiddleName+' '+LastName AS FullName,GroupName,PhoneNumber,MobileNumber,FaxNumber,EmailAddress FROM Contacts WHERE LUserID=" & UserID & " AND ContactID " & inWhole)
            With rptContacts
                Set .DataSource = rs
                .Orientation = rptOrientLandscape
                .Show vbModal, Me
            End With
        Else
            Set rs = con.Execute("SELECT ContactID,FirstName+' '+MiddleName+' '+LastName AS FullName,GroupName,PhoneNumber,MobileNumber,FaxNumber,EmailAddress FROM Contacts WHERE LUserID=" & UserID & " ORDER BY ContactID")
            With rptContacts
                Set .DataSource = rs
                .Orientation = rptOrientLandscape
                .Show vbModal, Me
            End With
        End If
    Case 7 'Show All
        cboCategory.ListIndex = 17
        Call cboCategory_Click
End Select
End Sub

Private Sub txtSearch_Change()

Select Case cboCategory.Text
    Case "Contact ID"
        strTest = "ContactID"
    Case "Last Name"
        strTest = "LastName"
    Case "First Name"
        strTest = "FirstName"
    Case "Middle Name"
        strTest = "MiddleName"
    Case "Religion"
        strTest = "Religion"
    Case "Age"
        strTest = "Age"
    Case "Position/Designation"
        strTest = "PositionDesignation"
    Case "Company/School"
        strTest = "CompanySchool"
    Case "Phone Number"
        strTest = "PhoneNumber"
    Case "Mobile Number"
        strTest = "MobileNumber"
    Case "Fax Number"
        strTest = "FaxNumber"
    Case "Email Address"
        strTest = "EmailAddress"
    Case "Address"
        strTest = "Address"
End Select
If txtSearch.Text <> "" Then
    If cboCategory.Text = "Contact ID" Then
        Set rs = con.Execute("SELECT * FROM Contacts WHERE LUserID=" & UserID & " AND ContactID=" & Val(txtSearch.Text) & "  ORDER BY ContactID")
    ElseIf cboCategory.Text = "Age" Then
        Set rs = con.Execute("SELECT * FROM Contacts WHERE  LUserID=" & UserID & " AND Age=" & Val(txtSearch.Text) & "  ORDER BY ContactID")
    Else
        Set rs = con.Execute("SELECT * FROM Contacts WHERE  LUserID=" & UserID & "  AND " & strTest & " LIKE '" & txtSearch.Text & "%'  ORDER BY ContactID")
    End If
    lvContacts.ListItems.Clear
    
    For xCount = 1 To rs.RecordCount
        With ls
            Set ls = lvContacts.ListItems.Add(, , rs!ContactID)
            ls.SubItems(1) = rs!FirstName & " " & rs!MiddleName & " " & rs!LastName
            ls.SubItems(2) = rs!GroupName
            rs.MoveNext
        End With
    Next xCount
    Call CheckCount(lvContacts, lblRecord)
Else
    Call LoadContacts(lvContacts)
End If
End Sub
