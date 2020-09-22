VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSelect 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "---"
   ClientHeight    =   4200
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8550
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSelect.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4200
   ScaleWidth      =   8550
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView lvContacts 
      Height          =   3495
      Left            =   120
      TabIndex        =   5
      ToolTipText     =   "Double-click an item to view/edit details"
      Top             =   120
      Visible         =   0   'False
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   6165
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
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
   Begin MSComctlLib.ListView lvStaff 
      Height          =   3495
      Left            =   120
      TabIndex        =   3
      ToolTipText     =   "Double-click an item to view/edit details"
      Top             =   120
      Visible         =   0   'False
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   6165
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
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
   Begin MSComctlLib.ListView lvStudents 
      Height          =   3495
      Left            =   120
      TabIndex        =   1
      ToolTipText     =   "Double-click an item to view/edit details"
      Top             =   120
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   6165
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
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
   Begin MSComctlLib.ListView lvUsers 
      Height          =   3495
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "Double-click an item to view/edit details"
      Top             =   120
      Visible         =   0   'False
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   6165
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
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
   Begin MSComctlLib.ListView lvEvents 
      Height          =   3495
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   6165
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FlatScrollBar   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Event ID"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Date"
         Object.Width           =   2522
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Time From"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Time To"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Topic"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Venue"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Details"
         Object.Width           =   3528
      EndProperty
   End
   Begin VB.Label lblRecord 
      Alignment       =   1  'Right Justify
      Caption         =   "---"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   3720
      Width           =   5775
   End
End
Attribute VB_Name = "frmSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
If CallForm = "Specific User" Then
    Call LoadUsers(lvUsers)
    lvUsers.Visible = True
    lvStudents.Visible = False
    lvStaff.Visible = False
    lvEvents.Visible = False
    lvContacts.Visible = False
ElseIf CallForm = "Specific Student" Or CallForm = "Specific Violation" Then
    Call LoadStudents(lvStudents)
    Call CheckCount(lvStudents, lblRecord)
    lvUsers.Visible = False
    lvStudents.Visible = True
    lvStaff.Visible = False
    lvEvents.Visible = False
ElseIf CallForm = "Specific Staff" Then
    Call LoadStaff(lvStaff)
    Call CheckCount(lvStaff, lblRecord)
    lvUsers.Visible = False
    lvStudents.Visible = False
    lvStaff.Visible = True
    lvEvents.Visible = False
ElseIf CallForm = "Specific Event" Then
    Call LoadEvents(lvEvents)
    Call CheckCount(lvEvents, lblRecord)
    lvUsers.Visible = False
    lvStudents.Visible = False
    lvStaff.Visible = False
    lvEvents.Visible = True
ElseIf CallForm = "Specific Contact" Then
    Call LoadContacts(lvContacts)
    lvUsers.Visible = False
    lvStudents.Visible = False
    lvStaff.Visible = False
    lvEvents.Visible = False
    lvContacts.Visible = True
End If
End Sub




Private Sub lvContacts_Click()
Call CheckCount(lvContacts, lblRecord)
End Sub

Private Sub lvContacts_DblClick()
Set rs = con.Execute("SELECT * FROM Contacts WHERE LUserID=" & UserID & " AND ContactID=" & Val(lvContacts.SelectedItem) & "")
If rs.RecordCount <> 0 Then
    With rptSpecificContact
        If fs.FileExists(App.Path & "\Images\" & rs!ImagePath) = True Then Set .Sections(3).Controls("img").Picture = LoadPicture(App.Path & "\Images\" & rs!ImagePath)
        Set .DataSource = rs
        .Caption = "Contact Info: " & rs!FirstName & " " & rs!MiddleName & " " & rs!LastName
        .Show vbModal, Me
    End With
Else
    MsgBox "No Records found for the selected contact.", vbOKOnly + vbExclamation, "No Records"
    Exit Sub
End If
End Sub

Private Sub lvContacts_KeyDown(KeyCode As Integer, Shift As Integer)
Call CheckCount(lvContacts, lblRecord)
End Sub

Private Sub lvContacts_KeyUp(KeyCode As Integer, Shift As Integer)
Call CheckCount(lvContacts, lblRecord)
End Sub

Private Sub lvEvents_Click()
Call CheckCount(lvEvents, lblRecord)
End Sub

Private Sub lvEvents_DblClick()
Set rs = con.Execute("SELECT * FROM EventsCalendar WHERE EventID=" & Val(lvEvents.SelectedItem) & "")
If rs.RecordCount <> 0 Then
    With rptSpecificEvent
        Set .DataSource = rs
        .Caption = "Event Info"
        .Show vbModal, Me
    End With
Else
    MsgBox "No Records found for the selected event.", vbOKOnly + vbExclamation, "No Records"
    Exit Sub
End If
End Sub

Private Sub lvEvents_KeyDown(KeyCode As Integer, Shift As Integer)
Call CheckCount(lvEvents, lblRecord)
End Sub

Private Sub lvEvents_KeyUp(KeyCode As Integer, Shift As Integer)
Call CheckCount(lvEvents, lblRecord)
End Sub

Private Sub lvStaff_Click()
Call CheckCount(lvStaff, lblRecord)
End Sub

Private Sub lvStaff_DblClick()
Set rs = con.Execute("SELECT * FROM Staff WHERE StaffID=" & Val(lvStaff.SelectedItem) & "")
If rs.RecordCount <> 0 Then
    With rptSpecificStaff
        If fs.FileExists(App.Path & "\Images\" & rs!ImagePath) = True Then Set .Sections(3).Controls("img").Picture = LoadPicture(App.Path & "\Images\" & rs!ImagePath)
        Set .DataSource = rs
        .Caption = "Staff Info: " & rs!FirstName & " " & rs!MiddleName & " " & rs!LastName
        .Show vbModal, Me
    End With
Else
    MsgBox "No Records found for the selected staff.", vbOKOnly + vbExclamation, "No Records"
    Exit Sub
End If
End Sub

Private Sub lvStaff_KeyDown(KeyCode As Integer, Shift As Integer)
Call CheckCount(lvStaff, lblRecord)
End Sub

Private Sub lvStaff_KeyUp(KeyCode As Integer, Shift As Integer)
Call CheckCount(lvStaff, lblRecord)
End Sub

Private Sub lvStudents_DblClick()
If CallForm = "Specific Violation" Then
    Set rs = con.Execute("SELECT * FROM Violations INNER JOIN Students ON Violations.VStudentID=Students.StudentID WHERE VStudentID=" & Val(lvStudents.SelectedItem) & " ORDER BY ViolationID")
    If rs.RecordCount <> 0 Then
        With rptViolationHistory
            .Caption = "Violation History - " & rs!FirstName & " " & rs!MiddleName & " " & rs!LastName
            Set .DataSource = rs
            .Show vbModal, Me
        End With
    Else
        MsgBox "No Records found for the selected student.", vbOKOnly + vbExclamation, "No Records"
        Exit Sub
    End If
Else
    Set rs = con.Execute("SELECT * FROM Students WHERE StudentID=" & Val(lvStudents.SelectedItem) & " ORDER BY StudentID")
    If rs.RecordCount <> 0 Then
        With rptStudent
            If fs.FileExists(App.Path & "\Images\" & rs!ImagePath) = True Then Set .Sections(3).Controls("img").Picture = LoadPicture(App.Path & "\Images\" & rs!ImagePath)
            Set .DataSource = rs
            .Caption = "Student Info: " & rs!FirstName & " " & rs!MiddleName & " " & rs!LastName
            .Show vbModal, Me
        End With
    Else
        MsgBox "No Records found for the selected user.", vbOKOnly + vbExclamation, "No Records"
        Exit Sub
    End If
End If
End Sub

Private Sub lvUsers_Click()
Call CheckCount(lvUsers, lblRecord)
End Sub

Private Sub lvUsers_DblClick()
Set rs = con.Execute("SELECT * FROM LogTrail WHERE Username='" & lvUsers.SelectedItem.SubItems(1) & "' ORDER BY LogID")
If rs.RecordCount <> 0 Then
    With rptUserTrail
        .Caption = "Log Details - " & strTest & " (" & rs!Privilege & ")"
        Set .DataSource = rs
        .Show vbModal, Me
    End With
Else
    MsgBox "No Records found for the selected user.", vbOKOnly + vbExclamation, "No Records"
    Exit Sub
End If
End Sub

Private Sub lvUsers_KeyDown(KeyCode As Integer, Shift As Integer)
Call CheckCount(lvUsers, lblRecord)
End Sub

Private Sub lvUsers_KeyUp(KeyCode As Integer, Shift As Integer)
Call CheckCount(lvUsers, lblRecord)
End Sub



