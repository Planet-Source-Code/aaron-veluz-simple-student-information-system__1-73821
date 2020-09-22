VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H8000000C&
   Caption         =   "Student System"
   ClientHeight    =   7395
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   10305
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   7395
   ScaleWidth      =   10305
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   7200
      Top             =   5520
   End
   Begin MSComctlLib.StatusBar sBar 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   7020
      Width           =   10305
      _ExtentX        =   18177
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   4948
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   4948
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            AutoSize        =   2
            Enabled         =   0   'False
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            AutoSize        =   2
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   3
            AutoSize        =   2
            TextSave        =   "INS"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mFile 
      Caption         =   "File"
      Begin VB.Menu mNew 
         Caption         =   "New"
         Begin VB.Menu mNewContactGroup 
            Caption         =   "Contact Group"
         End
         Begin VB.Menu mNewContact 
            Caption         =   "Contact"
         End
         Begin VB.Menu mNewEvent 
            Caption         =   "Event"
         End
         Begin VB.Menu mNewStaff 
            Caption         =   "Staff"
         End
         Begin VB.Menu mNewStudent 
            Caption         =   "Student"
         End
         Begin VB.Menu mNewViolation 
            Caption         =   "Student Violation"
         End
         Begin VB.Menu mNewUser 
            Caption         =   "User"
         End
      End
      Begin VB.Menu mEditSchoolInfo 
         Caption         =   "Edit School Info"
      End
      Begin VB.Menu mLock 
         Caption         =   "Lock System"
      End
      Begin VB.Menu mLogout 
         Caption         =   "Log Out"
      End
      Begin VB.Menu mExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mView 
      Caption         =   "View"
      Begin VB.Menu mViewCalendarofEvents 
         Caption         =   "Calendar of Events"
         Begin VB.Menu mViewCalendarSelect 
            Caption         =   "Select Date"
         End
         Begin VB.Menu mViewCalendarAll 
            Caption         =   "All Events"
         End
      End
      Begin VB.Menu mViewContactGroups 
         Caption         =   "Contact Groups"
      End
      Begin VB.Menu mViewContactsList 
         Caption         =   "Contacts List"
      End
      Begin VB.Menu mLogTrail 
         Caption         =   "Log Trail"
      End
      Begin VB.Menu mViewStaff 
         Caption         =   "Staff List"
      End
      Begin VB.Menu mViewStudents 
         Caption         =   "Student List"
      End
      Begin VB.Menu mViewUsers 
         Caption         =   "User List"
      End
      Begin VB.Menu mViewViolations 
         Caption         =   "Violations List"
      End
   End
   Begin VB.Menu mReports 
      Caption         =   "Reports"
      Begin VB.Menu mRptContactGroups 
         Caption         =   "Contact Groups"
      End
      Begin VB.Menu mRptContacts 
         Caption         =   "Contacts"
         Begin VB.Menu mRptSpecificContact 
            Caption         =   "Specific"
         End
         Begin VB.Menu mRptAllContacts 
            Caption         =   "All"
         End
      End
      Begin VB.Menu mRptEvents 
         Caption         =   "Events"
         Begin VB.Menu mRptSpecificEvent 
            Caption         =   "Specific Event"
         End
         Begin VB.Menu mRptAllEvents 
            Caption         =   "All Events"
         End
      End
      Begin VB.Menu mRptLogTrail 
         Caption         =   "Log Trail"
         Begin VB.Menu mRptSpecificUser 
            Caption         =   "History (Specific User)"
         End
         Begin VB.Menu mRptAllUsers 
            Caption         =   "All Users"
         End
      End
      Begin VB.Menu mRptStaffList 
         Caption         =   "Staff List"
         Begin VB.Menu mRptSpecificStaff 
            Caption         =   "Specific Staff"
         End
         Begin VB.Menu mRptAllStaff 
            Caption         =   "All Staff"
         End
      End
      Begin VB.Menu mRptStudents 
         Caption         =   "Students"
         Begin VB.Menu mRptSpecificStudent 
            Caption         =   "Specific Student"
         End
         Begin VB.Menu mRptAllStudents 
            Caption         =   "All Students"
         End
      End
      Begin VB.Menu mRptUserList 
         Caption         =   "User List"
      End
      Begin VB.Menu mRptViolations 
         Caption         =   "Violations"
         Begin VB.Menu mRptViolationSpecific 
            Caption         =   "History (Specific Student)"
         End
         Begin VB.Menu mRptViolationAll 
            Caption         =   "All Violations"
         End
      End
   End
   Begin VB.Menu mUtilities 
      Caption         =   "Utilities"
      Begin VB.Menu mBackupRestore 
         Caption         =   "Backup/Restore Database"
      End
      Begin VB.Menu mReset 
         Caption         =   "Reset Database"
      End
   End
   Begin VB.Menu mAbout 
      Caption         =   "About"
      Begin VB.Menu mAboutDeveloper 
         Caption         =   "Developer"
      End
      Begin VB.Menu mAboutSchoolInfo 
         Caption         =   "School Info"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Activate()
If con.State = 0 Then Call konek
End Sub

Private Sub Form_Unload(Cancel As Integer)
If MsgBox("Are you sure you want to exit the application?", vbYesNo + vbExclamation, "Confirm Exit Application") = vbNo Then
    Cancel = 1
Else
    Call Logout
    End
End If
End Sub

Private Sub mAboutDeveloper_Click()
With frmAbout
    .fraDeveloperInfo.Visible = True
    .fraSchoolInfo.Visible = False
    .Width = 4680
    .Height = 3270
    .Caption = "Developer Information"
    .Show vbModal, Me
End With
End Sub

Private Sub mAboutSchoolInfo_Click()
Set rs = con.Execute("SELECT * FROM SchoolInfo")
If rs.RecordCount = 0 Then
    If MsgBox("School information has not been set yet. Would you like to set it now?", vbYesNo + vbQuestion, "Set School Info") = vbNo Then Exit Sub
    frmSchoolInfo.Show vbModal, Me
Else
    With frmAbout
        .fraSchoolInfo.Visible = True
        .fraDeveloperInfo.Visible = False
        .Width = 8475
        .Height = 4950
        .lblAddress.Caption = rs!SchoolAddress
        .lblContactNumber.Caption = rs!ContactNumber
        .lblEmailAddress.Caption = rs!EmailAddress
        .lblName.Caption = rs!SchoolName
        .lblOwnerHead.Caption = rs!OwnerHead
        Call loadLogo(App.Path & "\Images\" & rs!ImagePath, .img, .pic)
        .Caption = "About " & rs!SchoolName
        .Show vbModal, Me
    End With
End If

End Sub

Private Sub mBackupRestore_Click()
frmBackupRestore.Show vbModal, Me
End Sub

Private Sub mEditSchoolInfo_Click()
Set rs = con.Execute("SELECT * FROM SchoolInfo")
With frmSchoolInfo
    If rs.RecordCount = 0 Then
        .cmdSave.Caption = "Save"
    Else
        .Caption = "School Info - " & rs!SchoolName
        .cmdSave.Caption = "Update"
    End If
    .Show vbModal, Me
End With
End Sub

Private Sub mExit_Click()
Unload Me
End Sub

Private Sub mLock_Click()
frmLock.Show vbModal, Me
End Sub

Private Sub mLogout_Click()
If MsgBox("Are you sure you want to Log Off?", vbYesNo + vbExclamation, "Confirm Log Off") = vbNo Then Exit Sub
Call Logout
MsgBox "Goodbye, " & Username & ". Thank you for using this application.", vbOKOnly + vbInformation, "Closing"
Username = ""
Password = ""
Privilege = ""
sBar.Panels(1).Text = ""
sBar.Panels(2).Text = ""
Me.Hide
frmLogin.Show vbModal, Me
End Sub

Private Sub mLogTrail_Click()
frmLogTrail.Show vbModal, Me
End Sub

Private Sub mNewContact_Click()
Set rs = con.Execute("SELECT * FROM ContactGroups WHERE LUserID=" & UserID & " ORDER BY GroupID")
If rs.RecordCount = 0 Then
    If MsgBox("No Contact Groups found. Would you like to add a new Contact Group?", vbYesNo + vbQuestion, "No Records") = vbNo Then
        Exit Sub
    Else
        Call mNewContactGroup_Click
    End If
Else
    frmNewContact.cboContactGroup.Clear
    With rs
        frmNewContact.cboContactGroup.AddItem !GroupName
        rs.MoveNext
    End With
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
        .Caption = "New Contact"
        .cboContactGroup.ListIndex = 0
        .Show vbModal, Me
    End With
End If
End Sub

Private Sub mNewContactGroup_Click()
Set rs = con.Execute("SELECT * FROM ContactGroups WHERE LUserID=" & UserID & " ORDER BY GroupID")
If rs.RecordCount = 0 Then
    xCount = 1
Else
    rs.MoveLast
    xCount = Val(rs!GroupID) + 1
End If
With frmNewContactGroup
    .Caption = "New Contact Group"
    .lblRecordID.Caption = xCount
    .cmdAdd.Caption = "Add Group"
    .Show vbModal, Me
End With
End Sub

Private Sub mNewEvent_Click()
With frmNewEvent
    Set rs = con.Execute("SELECT * FROM EventsCalendar ORDER BY EventID")
    If rs.RecordCount = 0 Then
        .lblRecordID.Caption = 1
        .lblSequence.Caption = 1
    Else
        rs.MoveLast
        .lblRecordID.Caption = rs!EventID + 1
        Set rs = con.Execute("SELECT * FROM EventsCalendar WHERE EventDate=#" & .EventDate.Value & "# ORDER BY Sequence")
        If rs.RecordCount = 0 Then
            .lblSequence.Caption = 1
        Else
            rs.MoveLast
            .lblSequence.Caption = rs!Sequence + 1
        End If
    End If
    .EventDate.Value = Format(Now, "mm/dd/yyyy")
    .Show vbModal, Me
End With
End Sub

Private Sub mNewStaff_Click()
Set rs = con.Execute("SELECT * FROM Staff ORDER BY StaffID")
If rs.RecordCount = 0 Then
    xCount = 1
Else
    rs.MoveLast
    xCount = Val(rs!StaffID) + 1
End If
With frmNewStaff
    .Caption = "Add Staff"
    .cmdAdd.Caption = "Add Staff"
    .cboCivilStatus.ListIndex = 0
    .lblRecordID.Caption = xCount
    .Show vbModal, Me
End With
End Sub

Private Sub mNewStudent_Click()
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
    .cmdAdd.Caption = "Add Student"
End With
End Sub

Private Sub mNewUser_Click()
Set rs = con.Execute("SELECT * FROM Users ORDER BY UserID")
If rs.RecordCount = 0 Then
    xCount = 1
Else
    rs.MoveLast
    xCount = Val(rs!UserID) + 1
End If
With frmNewUser
    .Caption = "Add User"
    .lblRecordID.Caption = xCount
    .cboPrivilege.ListIndex = 0
    .cboPrivilege.Visible = True
    .lblPrivilege.Visible = True
    .Show vbModal, Me
    .cmdAdd.Caption = "Add User"
End With
End Sub

Private Sub mNewViolation_Click()
Set rs = con.Execute("SELECT * FROM Students ORDER BY StudentID")
If rs.RecordCount = 0 Then
    If MsgBox("No records found. Would you like to add a new Student?", vbYesNo + vbQuestion, "No Records") = vbNo Then
        Exit Sub
    Else
        Call mNewStudent_Click
    End If
Else
    Set rs = con.Execute("SELECT * FROM Violations ORDER BY ViolationID")
    If rs.RecordCount = 0 Then
        xCount = 1
    Else
        rs.MoveLast
        xCount = Val(rs!ViolationID) + 1
    End If
    With frmNewViolation
        Call LoadStudents2(.lvStudents)
        .cmdAdd.Caption = "Add Violation"
        .lblRecordID.Caption = xCount
        .dtDate.Value = Date
        .dtTime.Value = Time
        .Show vbModal, Me
    End With
End If
End Sub

Private Sub mReset_Click()
If MsgBox("Are you sure you want to delete all system records including:" & vbCrLf & vbCrLf & "* Log Trail" & vbCrLf & "* Staff List" & vbCrLf & "* Student List" & vbCrLf & "* User List" & vbCrLf & "* Violations List" & vbCrLf & "* School Info" & vbCrLf & "* Contacts List" & vbCrLf & "* Contact Groups" & vbCrLf & "* and all image files?" & vbCrLf, vbYesNo + vbExclamation, "Confirm System Reset") = vbNo Then Exit Sub
If InputBox("Enter password to reset", "RESET") <> Password Then MsgBox "Invalid password.", vbOKOnly + vbExclamation, "Error": Exit Sub
With con
    .Execute "DELETE * FROM LogTrail"
    .Execute "DELETE * FROM Staff"
    .Execute "DELETE * FROM Students"
    .Execute "DELETE * FROM Users"
    .Execute "DELETE * FROM Violations"
    .Execute "DELETE * FROM SchoolInfo"
    .Execute "DELETE * FROM Contacts"
    .Execute "DELETE * FROM ContactGroups"
End With
fs.DeleteFolder (App.Path & "\Images")
MsgBox "All records have been deleted. The application will now be closed.", vbOKOnly + vbInformation, "Reset Complete"
End
End Sub

Private Sub mRptAllContacts_Click()
Set rs = con.Execute("SELECT ContactID,FirstName+' '+MiddleName+' '+LastName AS FullName,GroupName,PhoneNumber,MobileNumber,FaxNumber,EmailAddress FROM Contacts WHERE LUserID=" & UserID & " ORDER BY ContactID")
If rs.RecordCount = 0 Then
    If MsgBox("No records found. Would you like to add a new Contact?", vbYesNo + vbQuestion, "No Records") = vbNo Then
        Exit Sub
    Else
        Call mNewContact_Click
    End If
Else
    Set rptContacts.DataSource = rs
    rptContacts.Show vbModal, Me
End If
End Sub

Private Sub mRptAllEvents_Click()
Set rs = con.Execute("SELECT * FROM EventsCalendar ORDER BY EventID")
If rs.RecordCount = 0 Then
    If MsgBox("No records found. Would you like to add a new Event?", vbYesNo + vbQuestion, "No Records") = vbNo Then
        Exit Sub
    Else
        Call mNewEvent_Click
    End If
Else
    Set rptEvents.DataSource = rs
    rptEvents.Show vbModal, Me
End If
End Sub

Private Sub mRptAllStaff_Click()
Set rs = con.Execute("SELECT StaffID, FirstName+' '+MiddleName+' '+LastName AS FullName, Designation FROM Staff ORDER BY StaffID")
If rs.RecordCount = 0 Then
    If MsgBox("No records found. Would you like to add a new Staff?", vbYesNo + vbQuestion, "No Records") = vbNo Then
        Exit Sub
    Else
        Call mNewStaff_Click
    End If
Else
    Set rptStaff.DataSource = rs
    rptStaff.Show vbModal, Me
End If
End Sub

Private Sub mRptAllStudents_Click()
Set rs = con.Execute("SELECT StudentID,SectionName,StudentNumber,LastName+', '+FirstName+' '+MiddleName AS FullName FROM Students ORDER BY StudentID")
If rs.RecordCount = 0 Then
    If MsgBox("No records found. Would you like to add a new Student?", vbYesNo + vbQuestion, "No Records") = vbNo Then
        Exit Sub
    Else
        Call mNewStudent_Click
    End If
Else
    With rptStudents
        Set .DataSource = rs
        .Show vbModal, Me
    End With
End If
End Sub

Private Sub mRptAllUsers_Click()
Set rs = con.Execute("SELECT * FROM LogTrail ORDER BY LogID")
With rptLogTrail
    Set .DataSource = rs
    .Show vbModal, Me
End With
End Sub

Private Sub mRptContactGroups_Click()
Set rs = con.Execute("SELECT * FROM ContactGroups WHERE LUserID=" & UserID & " ORDER BY GroupID")
If rs.RecordCount = 0 Then
    If MsgBox("No records found. Would you like to add a new Contact?", vbYesNo + vbQuestion, "No Records") = vbNo Then
        Exit Sub
    Else
        Call mNewContactGroup_Click
    End If
Else
    Set rptContactGroups.DataSource = rs
    rptContactGroups.Show vbModal, Me
End If
End Sub

Private Sub mRptSpecificContact_Click()
Set rs = con.Execute("SELECT * FROM Contacts ORDER BY ContactID")
If rs.RecordCount <> 0 Then
    CallForm = "Specific Contact"
    With frmSelect
        .Caption = "Select Contact from list"
        .Width = 8640
        .lblRecord.Width = .lvContacts.Width
        .Show vbModal, Me
    End With
    
Else
    Call mViewContactsList_Click
End If
End Sub

Private Sub mRptSpecificEvent_Click()
Set rs = con.Execute("SELECT * FROM EventsCalendar ORDER BY EventID")
If rs.RecordCount <> 0 Then
    CallForm = "Specific Event"
    With frmSelect
        .Caption = "Select Event from list"
        .Width = 11625
        .lblRecord.Width = .lvEvents.Width
        .Show vbModal, Me
    End With
    
Else
    Call mViewCalendarAll_Click
End If
End Sub

Private Sub mRptSpecificStaff_Click()
Set rs = con.Execute("SELECT * FROM Staff ORDER BY StaffID")
If rs.RecordCount <> 0 Then
    CallForm = "Specific Staff"
    With frmSelect
        .Caption = "Select Staff from list"
        .Width = 8640
        .lblRecord.Width = .lvStaff.Width
        .Show vbModal, Me
    End With
    
Else
    Call mViewStaff_Click
End If
End Sub

Private Sub mRptSpecificStudent_Click()
Set rs = con.Execute("SELECT * FROM Students ORDER BY StudentID")
If rs.RecordCount <> 0 Then
    CallForm = "Specific Student"
    With frmSelect
        .Caption = "Select Student from list"
        .Width = 8280
        .lblRecord.Width = .lvStudents.Width
        .Show vbModal, Me
    End With
    
Else
    Call mViewStudents_Click
End If
End Sub

Private Sub mRptSpecificUser_Click()
Set rs = con.Execute("SELECT * FROM LogTrail ORDER BY LogID")
If rs.RecordCount <> 0 Then
    CallForm = "Specific User"
    With frmSelect
        .Caption = "Select User from list"
        .Width = 6135
        .lblRecord.Width = .lvUsers.Width
        .Show vbModal, Me
    End With
    
Else
    Call mViewUsers_Click
End If
End Sub

Private Sub mRptStudentList_Click()


End Sub

Private Sub mRptUserList_Click()
Set rs = con.Execute("SELECT * FROM Users ORDER BY UserID")
Set rptUsers.DataSource = rs
rptUsers.Show vbModal, Me
End Sub

Private Sub mRptViolationAll_Click()
Set rs = con.Execute("SELECT ViolationDate,Violation,Sanction,Students.StudentNumber,Students.SectionName,Students.FirstName+' '+Students.MiddleName+' '+Students.LastName AS FullName FROM Violations INNER JOIN Students ON Violations.VStudentID=Students.StudentID ORDER BY ViolationID")
If rs.RecordCount = 0 Then
    If MsgBox("No records found. Would you like to add a new Student Violation?", vbYesNo + vbQuestion, "No Records") = vbNo Then
        Exit Sub
    Else
        Call mNewViolation_Click
    End If
Else
    With rptViolations
        Set .DataSource = rs
        .Orientation = rptOrientLandscape
        .Show vbModal, Me
    End With
End If
End Sub

Private Sub mRptViolationSpecific_Click()
Set rs = con.Execute("SELECT * FROM Violations ORDER BY ViolationID")
If rs.RecordCount <> 0 Then
    CallForm = "Specific Violation"
    With frmSelect
        .Caption = "Select Student from list"
        .Width = 8280
        .lblRecord.Width = .lvStudents.Width
        .Show vbModal, Me
    End With

Else
    Call mViewViolations_Click
End If
End Sub

Private Sub mViewCalendarAll_Click()
frmEvents.Show vbModal, Me
End Sub

Private Sub mViewCalendarSelect_Click()
frmEventCalendar.Show vbModal, Me
End Sub

Private Sub mViewContactGroups_Click()
frmContactGroups.Show vbModal, Me
End Sub

Private Sub mViewContactsList_Click()
frmContacts.Show vbModal, Me
End Sub

Private Sub mViewStaff_Click()
Set rs = con.Execute("SELECT * FROM Staff ORDER BY StaffID")
If rs.RecordCount = 0 Then
    If MsgBox("No records found. Would you like to add a new Staff?", vbYesNo + vbQuestion, "No Records") = vbNo Then
        Exit Sub
    Else
        Call mNewStaff_Click
    End If
Else
    frmStaff.Show vbModal, Me
End If
End Sub

Private Sub mViewStudents_Click()
Set rs = con.Execute("SELECT * FROM Students ORDER BY StudentID")
If rs.RecordCount = 0 Then
    If MsgBox("No records found. Would you like to add a new Student?", vbYesNo + vbQuestion, "No Records") = vbNo Then
        Exit Sub
    Else
        Call mNewStudent_Click
    End If
Else
    frmStudents.Show vbModal, Me
End If

End Sub

Private Sub mViewUsers_Click()
Set rs = con.Execute("SELECT * FROM Users ORDER BY UserID")
If rs.RecordCount = 0 Then
    If MsgBox("No records found. Would you like to add a new User?", vbYesNo + vbQuestion, "No Records") = vbNo Then
        Exit Sub
    Else
        Call mNewUser_Click
    End If
Else
    frmUsers.Show vbModal, Me
End If

End Sub

Private Sub mViewViolations_Click()
Set rs = con.Execute("SELECT * FROM Violations ORDER BY ViolationID")
If rs.RecordCount = 0 Then
    If MsgBox("No records found. Would you like to add a new Student Violation?", vbYesNo + vbQuestion, "No Records") = vbNo Then
        Exit Sub
    Else
        Call mNewViolation_Click
    End If
Else
    frmViolations.Show vbModal, Me
End If

End Sub

Private Sub Timer1_Timer()
sBar.Panels(2).Text = Format(Now, "DDDD, mmmm dd, yyyy hh:mm:ss am/pm")
End Sub
