VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmViolations 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Violations List"
   ClientHeight    =   5775
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10575
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmViolations.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5775
   ScaleWidth      =   10575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraViolations 
      Height          =   5055
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   10335
      Begin VB.Frame fraSearch 
         Caption         =   "Search Student"
         Height          =   975
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   10095
         Begin MSComCtl2.DTPicker dtpFrom 
            Height          =   375
            Left            =   4080
            TabIndex        =   10
            Top             =   480
            Visible         =   0   'False
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   661
            _Version        =   393216
            Format          =   17039361
            CurrentDate     =   40567
         End
         Begin VB.ComboBox cboCategory 
            Height          =   315
            ItemData        =   "frmViolations.frx":1082
            Left            =   1080
            List            =   "frmViolations.frx":10A1
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   480
            Width           =   2055
         End
         Begin MSComCtl2.DTPicker dtpTo 
            Height          =   375
            Left            =   6480
            TabIndex        =   11
            Top             =   480
            Visible         =   0   'False
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   661
            _Version        =   393216
            Format          =   17039361
            CurrentDate     =   40567
         End
         Begin VB.ComboBox cboGender 
            Height          =   315
            ItemData        =   "frmViolations.frx":1118
            Left            =   3480
            List            =   "frmViolations.frx":1122
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   480
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.TextBox txtSearch 
            Height          =   375
            Left            =   3480
            MaxLength       =   255
            TabIndex        =   7
            Top             =   480
            Visible         =   0   'False
            Width           =   2535
         End
         Begin VB.Label lblFrom 
            Caption         =   "From"
            Height          =   255
            Left            =   3480
            TabIndex        =   12
            Top             =   480
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.Label lblTo 
            Caption         =   "To"
            Height          =   255
            Left            =   6000
            TabIndex        =   13
            Top             =   480
            Visible         =   0   'False
            Width           =   375
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
         TabIndex        =   2
         Top             =   4680
         Width           =   1095
      End
      Begin MSComctlLib.ListView lvViolations 
         Height          =   3255
         Left            =   120
         TabIndex        =   1
         ToolTipText     =   "Double-click an item to view/edit details"
         Top             =   1320
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   5741
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
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Violation ID"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Name"
            Object.Width           =   5468
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Student #"
            Object.Width           =   2999
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Section"
            Object.Width           =   2999
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Violation"
            Object.Width           =   3704
         EndProperty
      End
      Begin VB.Label lblRecord 
         Alignment       =   1  'Right Justify
         Caption         =   "---"
         Height          =   255
         Left            =   1320
         TabIndex        =   3
         Top             =   4680
         Width           =   8895
      End
   End
   Begin MSComctlLib.Toolbar tbrViolations 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   10575
      _ExtentX        =   18653
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
         Left            =   7440
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
               Picture         =   "frmViolations.frx":1134
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmViolations.frx":21C6
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmViolations.frx":3258
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmViolations.frx":42EA
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "frmViolations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cboCategory_Click()
If cboCategory.Text = "Gender" Then
    cboGender.Visible = True
    txtSearch.Visible = False
    dtpFrom.Visible = False
    dtpTo.Visible = False
    lblFrom.Visible = False
    lblTo.Visible = False
    cboGender.ListIndex = 0
ElseIf cboCategory.Text = "Violation Date" Then
    cboGender.Visible = False
    txtSearch.Visible = False
    dtpFrom.Visible = True
    dtpTo.Visible = True
    lblFrom.Visible = True
    lblTo.Visible = True
ElseIf cboCategory.Text = "ALL RECORDS" Then
    cboGender.Visible = False
    txtSearch.Visible = False
    dtpFrom.Visible = False
    dtpTo.Visible = False
    lblFrom.Visible = False
    lblTo.Visible = False
    Call LoadViolations(lvViolations)
    Call CheckCount(lvViolations, lblRecord)
Else
    cboGender.Visible = False
    txtSearch.Visible = True
    dtpFrom.Visible = False
    dtpTo.Visible = False
    lblFrom.Visible = False
    lblTo.Visible = False
End If

End Sub

Private Sub cboGender_Click()
If lvViolations.ListItems.Count = 0 Then Exit Sub
Set rs = con.Execute("SELECT ViolationID,ViolationDate,Violation,Sanction,Students.StudentNumber,Students.SectionName,Students.FirstName+' '+Students.MiddleName+' '+Students.LastName AS FullName FROM Violations INNER JOIN Students ON Violations.VStudentID=Students.StudentID ORDER BY ViolationID")
lvViolations.ListItems.Clear

For xCount = 1 To rs.RecordCount
    With ls
        Set ls = lvViolations.ListItems.Add(, , rs!ViolationID)
        ls.SubItems(1) = rs!FullName
        ls.SubItems(2) = rs!StudentNumber
        ls.SubItems(3) = rs!SectionName
        ls.SubItems(4) = rs!Violation
        rs.MoveNext
    End With
Next xCount
Call CheckCount(lvViolations, lblRecord)
End Sub

Private Sub chkSelect_Click()
Call SelectAll(lvViolations, chkSelect)
End Sub





Private Sub cmdShowAll_Click()

End Sub

Private Sub dtpFrom_Change()
Call CheckDates
End Sub

Private Sub dtpFrom_Click()
Call CheckDates
End Sub

Private Sub dtpFrom_DblClick()
Call CheckDates
End Sub

Private Sub dtpFrom_DropDown()
Call CheckDates
End Sub

Private Sub dtpTo_Change()
Call CheckDates
End Sub

Private Sub dtpTo_Click()
Call CheckDates
End Sub

Private Sub dtpTo_DblClick()
Call CheckDates
End Sub

Private Sub dtpTo_DropDown()
Call CheckDates
End Sub

Private Sub Form_Activate()
If con.State = 0 Then Call konek
End Sub

Private Sub Form_Load()
Call EnableControls
cboCategory.ListIndex = 8
Call cboCategory_Click
Call LoadViolations(lvViolations)
Call CheckCount(lvViolations, lblRecord)
End Sub

Private Sub lvViolations_Click()
Call CheckCount(lvViolations, lblRecord)
End Sub

Private Sub lvViolations_DblClick()
If lvViolations.ListItems.Count = 0 Then Exit Sub
Set rs = con.Execute("SELECT * FROM Violations INNER JOIN Students ON Violations.VStudentID=Students.StudentID WHERE ViolationID=" & IIf((lvViolations.SelectedItem = 0), 1, lvViolations.SelectedItem) & "")
With frmNewViolation
    .lvStudents.ListItems.Clear
    Set ls = .lvStudents.ListItems.Add(, , rs!StudentID)
    ls.SubItems(1) = rs!FirstName & " " & rs!MiddleName & " " & rs!LastName
    .lblRecordID.Caption = rs!ViolationID
    .lblStudentID.Caption = rs!StudentID
    .lblSection.Caption = rs!SectionName
    .lblStudentNumber.Caption = rs!StudentNumber
    .dtDate.Value = rs!ViolationDate
    .dtTime.Value = rs!ViolationTime
    .txtViolation.Text = rs!Violation
    .txtSanction.Text = rs!Sanction
    .lblGender.Caption = rs!Gender
    .Caption = "Edit Violation - " & .lvStudents.SelectedItem.SubItems(1)
    .cmdAdd.Caption = "Update"
    .Show vbModal, Me
End With
Call LoadViolations(lvViolations)
Call CheckCount(lvViolations, lblRecord)
End Sub

Private Sub lvViolations_KeyDown(KeyCode As Integer, Shift As Integer)
Call CheckCount(lvViolations, lblRecord)
End Sub

Private Sub lvViolations_KeyUp(KeyCode As Integer, Shift As Integer)
Call CheckCount(lvViolations, lblRecord)
End Sub

Private Sub tbrViolations_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
    Case 1 'Add
        Set rs = con.Execute("SELECT * FROM Violations ORDER BY ViolationID")
        If rs.RecordCount = 0 Then
            xCount = 1
        Else
            rs.MoveLast
            xCount = Val(rs!ViolationID) + 1
        End If
        With frmNewViolation
            Call LoadStudents2(.lvStudents)
            .cmdAdd.Caption = "New Violation"
            .lblRecordID.Caption = xCount
            .dtDate.Value = Date
            .dtTime.Value = Time
            .Show vbModal, Me
        End With
        Call LoadViolations(lvViolations)
        Call CheckCount(lvViolations, lblRecord)
    Case 3 'Delete
        If lvViolations.ListItems.Count = 0 Then Exit Sub
        Call CountSelected(lvViolations)
        If yCount <> 0 Then
            If MsgBox("Are you sure you want to delete the selected item(s)?", vbYesNo + vbExclamation, "Confirm Delete") = vbNo Then Exit Sub
            For xCount = 1 To lvViolations.ListItems.Count
                If lvViolations.ListItems(xCount).Checked = True Then
                    CurrRec = lvViolations.ListItems(xCount)
                    con.Execute ("DELETE FROM Violations INNER JOIN Students ON Violations.StudentID=Students.StudentID WHERE ViolationID = " & CurrRec & "")
                End If
            Next xCount
            Call LoadViolations(lvViolations)
            Call CheckCount(lvViolations, lblRecord)
        End If
    Case 5 'Print
        If lvViolations.ListItems.Count = 0 Then Exit Sub
        Call CountSelected(lvViolations)
        If yCount = 1 Then
            For xCount = 1 To lvViolations.ListItems.Count
                If lvViolations.ListItems(xCount).Checked = True Then
                    CurrRec = lvViolations.ListItems(xCount)
                    Set rs = con.Execute("SELECT * FROM Violations INNER JOIN Students ON Violations.VStudentID=Students.StudentID WHERE ViolationID=" & CurrRec & "")
                End If
            Next xCount
            With rptViolation
                If fs.FileExists(App.Path & "\Images\" & rs!ImagePath) Then Set .Sections(3).Controls("img").Picture = LoadPicture(App.Path & "\Images\" & rs!ImagePath)
                Set .DataSource = rs
                .Caption = "Student Violation Info: " & rs!FirstName & " " & rs!MiddleName & " " & rs!LastName
                .Show vbModal, Me
            End With
        ElseIf yCount > 1 And yCount <> lvViolations.ListItems.Count Then
            inPart = ""
            inPart2 = ""
            inWhole = ""
            For xCount = 1 To lvViolations.ListItems.Count
                If lvViolations.ListItems(xCount).Checked = True Then
                    inPart = lvViolations.ListItems(xCount) & ", "
                    inPart2 = inPart2 & inPart
                End If
                inWhole = "IN ( " & inPart2 & ")"
            Next xCount
            Set rs = con.Execute("SELECT ViolationID,ViolationDate,Violation,Sanction,Students.StudentNumber,Students.SectionName,Students.FirstName+' '+Students.MiddleName+' '+Students.LastName AS FullName FROM Violations INNER JOIN Students ON Violations.VStudentID=Students.StudentID WEHRE ViolationID " & inWhole & " ORDER BY ViolationID")
            With rptViolations
                Set .DataSource = rs
                .Orientation = rptOrientLandscape
                .Show vbModal, Me
            End With
        Else
            Set rs = con.Execute("SELECT ViolationID,ViolationDate,Violation,Sanction,Students.StudentNumber,Students.SectionName,Students.FirstName+' '+Students.MiddleName+' '+Students.LastName AS FullName FROM Violations INNER JOIN Students ON Violations.VStudentID=Students.StudentID ORDER BY ViolationID")

            With rptViolations
                Set .DataSource = rs
                .Orientation = rptOrientLandscape
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
    Case "Violation ID"
        strTest = "ViolationID"
    Case "Last Name"
        strTest = "LastName"
    Case "First Name"
        strTest = "FirstName"
    Case "Middle Name"
        strTest = "MiddleName"
    Case "Section"
        strTest = "Section"
    Case "Violation"
        strTest = "Violation"
    Case "Student Number"
        strTest = "StudentNumber"
End Select
If txtSearch.Text <> "" Then
    If cboCategory.Text = "Violation ID" Then
        Set rs = con.Execute("SELECT ViolationID,ViolationDate,Violation,Sanction,Students.StudentNumber,Students.SectionName,Students.FirstName+' '+Students.MiddleName+' '+Students.LastName AS FullName FROM Violations INNER JOIN Students ON Violations.VStudentID=Students.StudentID WHERE ViolationID=" & Val(txtSearch.Text) & " ORDER BY ViolationID")
    Else
        Set rs = con.Execute("SELECT ViolationID,ViolationDate,Violation,Sanction,Students.StudentNumber,Students.SectionName,Students.FirstName+' '+Students.MiddleName+' '+Students.LastName AS FullName FROM Violations INNER JOIN Students ON Violations.VStudentID=Students.StudentID WHERE " & strTest & " LIKE '" & txtSearch.Text & "%' ORDER BY ViolationID")
    End If
    lvViolations.ListItems.Clear
    For xCount = 1 To rs.RecordCount
        With ls
            Set ls = lvViolations.ListItems.Add(, , rs!ViolationID)
            ls.SubItems(1) = rs!FullName
            ls.SubItems(2) = rs!StudentNumber
            ls.SubItems(3) = rs!SectionName
            ls.SubItems(4) = rs!Violation
            rs.MoveNext
        End With
    Next xCount
    Call CheckCount(lvViolations, lblRecord)
Else
    Call LoadViolations(lvViolations)
    Call CheckCount(lvViolations, lblRecord)
End If
End Sub
Public Sub CheckDates()
If lvViolations.ListItems.Count = 0 Then Exit Sub
Set rs = con.Execute("SELECT ViolationID,ViolationDate,Violation,Sanction,Students.StudentNumber,Students.SectionName,Students.FirstName+' '+Students.MiddleName+' '+Students.LastName AS FullName FROM Violations INNER JOIN Students ON Violations.VStudentID=Students.StudentID WHERE ViolationDate BETWEEN #" & dtpFrom.Value & "# AND #" & dtpTo.Value & "# ORDER BY ViolationID")
lvViolations.ListItems.Clear

For xCount = 1 To rs.RecordCount
    With ls
        Set ls = lvViolations.ListItems.Add(, , rs!ViolationID)
        ls.SubItems(1) = rs!FullName
        ls.SubItems(2) = rs!StudentNumber
        ls.SubItems(3) = rs!SectionName
        ls.SubItems(4) = rs!Violation
        rs.MoveNext
    End With
Next xCount
Call CheckCount(lvViolations, lblRecord)
End Sub
