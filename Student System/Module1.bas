Attribute VB_Name = "Module1"
Option Explicit
Public rs As New ADODB.Recordset
Public rsUsers As New ADODB.Recordset
Public con As New ADODB.Connection
Public Username, Password, Privilege, CallForm, FullName, FullPath, OldPath, strTest As String
Public UserID, xCount, yCount, zCount, CurrRec, ErrCounter, Age As Long
Public ls As ListItem
Public fs As New FileSystemObject
Public inPart, inPart2, inWhole As String
Public MM, DD, YYYY, NowM, NowD, NowY As Integer
Public Sub konek()
On Error GoTo Hell
con.CursorLocation = adUseClient
If con.State <> 1 Then
    con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\MainDB.mdb;Jet OLEDB:Database Password=123!@#"
End If
Exit Sub
Hell:
    MsgBox Err.Description, vbCritical, "Error"
    End
End Sub
Public Function recfound(ByVal sField As String, ByVal sfindtext As String) As Boolean

Set rs = con.Execute("SELECT * FROM Users where StrComp(UserName, '" & sfindtext & "', 0) = 0")

If rs.EOF Then
    recfound = False
Else
    With rs
        recfound = True
        UserID = !UserID
        Username = !Username
        Password = !UPassword
        Privilege = !Privilege
    End With
End If
End Function
Sub Main()
On Error GoTo Hell
If App.PrevInstance = True Then
    MsgBox "Application is already running.", vbOKOnly + vbExclamation, "Run"
    End
End If
If fs.FolderExists(App.Path & "\Images") = False Then fs.CreateFolder App.Path & "\Images"
Call konek
frmSplash.Show
Exit Sub
Hell:
    MsgBox Err.Description, vbCritical, "Error"
    End
End Sub
Public Sub SelText(txt As TextBox)
txt.SelStart = 0
txt.SelLength = Len(txt.Text)
txt.SetFocus
End Sub
Public Function WarnBlank(txt As TextBox, txtField As String) As Boolean
If txt.Text = "" Then
    MsgBox txtField & " is required. Please enter a valid value.", vbOKOnly + vbExclamation, txtField & " Required"
    txt.SetFocus
    WarnBlank = True
Else
    WarnBlank = False
End If
End Function
Public Function CheckCount(lv As ListView, lbl As Label)
If lv.ListItems.Count = 0 Then
    lbl.Caption = "---"
Else
    lbl.Caption = lv.SelectedItem.Index & " of " & lv.ListItems.Count
End If
End Function
Public Sub SelectAll(lv As ListView, chk As CheckBox)
If lv.ListItems.Count = 0 Then Exit Sub
For xCount = 1 To lv.ListItems.Count
    If chk.Value = 1 Then
    lv.ListItems(xCount).Checked = True
    Else
    lv.ListItems(xCount).Checked = False
    End If
Next xCount
End Sub
Public Sub Login()
If rs.State = 1 Then rs.Close
rs.Open "SELECT * FROM LogTrail ORDER BY LogID", con, adOpenKeyset, adLockOptimistic
If rs.RecordCount = 0 Then
    xCount = 1
ElseIf rs.RecordCount <> 0 Then
    rs.MoveLast
    xCount = rs!LogID + 1
End If
With rs
    .AddNew
    !LogID = xCount
    !LUserID = UserID
    !Username = Username
    !Privilege = Privilege
    !LogDate = Format(Now, "mm/dd/yyyy")
    !TimeIn = Format(Now, "hh:mm:ss am/pm")
    .UpdateBatch adAffectCurrent
    .Close
End With
End Sub
Public Sub Logout()
If rs.State = 1 Then rs.Close
rs.Open "SELECT * FROM LogTrail ORDER BY LogID", con, adOpenKeyset, adLockOptimistic
If rs.RecordCount = 0 Then Exit Sub
With rs
    .MoveLast
    !TimeOut = Format(Time, "hh:mm:ss am/pm")
    .UpdateBatch adAffectCurrent
    .Close
End With

End Sub
Function MistakeCounter(ErrStr As String)
If ErrCounter >= 1 And ErrCounter < 4 Then
    MsgBox "Please enter a valid " & ErrStr & ". You have " & 5 - ErrCounter & " attempts remaining.", vbOKOnly + vbExclamation, "Invalid " & ErrStr
ElseIf ErrCounter = 4 Then
    MsgBox "Please enter a valid " & ErrStr & ". This is your last attempt.", vbOKOnly + vbExclamation, "Invalid " & ErrStr
ElseIf ErrCounter = 5 Then
    MsgBox "Maximum number of mistaken attempts has been reached. The application will now be closed.", vbOKOnly + vbCritical, "Exit Application"
    End
End If
End Function
Public Sub ClearPic(txt As TextBox, dia As CommonDialog, ima As Image, pict As PictureBox)
txt.Text = ""
dia.FileName = ""
ima.Picture = Nothing
pict.Cls
FullPath = ""
End Sub
Public Sub loadLogo(loc As String, ima As Image, pict As PictureBox)
ima.Stretch = False
ima.Picture = LoadPicture(loc)
ima.Stretch = True
xCount = IIf(ima.Width > ima.Height, ima.Width / pict.Width, ima.Height / pict.Height)
ima.Width = ima.Width / xCount
ima.Height = ima.Height / xCount

ima.Top = (pict.Height - ima.Height) / 2
ima.Left = (pict.Width - ima.Width) / 2
End Sub
Public Sub Upload(txt As TextBox)
If txt.Text <> "" Then
    xCount = 1
    yCount = 999999999
    Randomize
    zCount = Int(Rnd * (yCount + 1 - xCount) + xCount)
    strTest = zCount & "." & fs.GetExtensionName(txt.Text)
    FileCopy txt.Text, App.Path & "\" & strTest
    
    Do Until fs.FileExists(App.Path & "\Images\" & strTest) = False
        Randomize
        zCount = Int(Rnd * (yCount + 1 - xCount) + xCount)
        Exit Do
    Loop
    FileCopy App.Path & "\" & strTest, App.Path & "\Images\" & strTest
    FullPath = zCount & "." & fs.GetExtensionName(txt.Text)
    Kill (App.Path & "\" & zCount & "." & fs.GetExtensionName(txt.Text))
End If
End Sub
Public Sub CountSelected(lvw As ListView)
yCount = 0
For xCount = 1 To lvw.ListItems.Count
    If lvw.ListItems(xCount).Checked = True Then
        yCount = yCount + 1
    End If
Next xCount
End Sub
Public Sub ValidateAge(dtp As DTPicker)


NowM = Val(Format(Now, "MM"))
NowD = Val(Format(Now, "dd"))
NowY = Val(Format(Now, "yyyy"))
MM = Val(Format(dtp.Value, "MM"))
DD = Val(Format(dtp.Value, "dd"))
YYYY = Val(Format(dtp.Value, "yyyy"))
Age = NowY - YYYY
If MM > NowM Or (MM = NowM And DD > NowD) Then
    Age = Age - 1
End If
If Age < 0 Then
    MsgBox "Birthdate is before current date.", vbOKOnly + vbExclamation, "Invalid BirthDate"
    dtp.SetFocus
End If
End Sub
Public Sub EnableControls()
If Privilege = "SuperAdministrator" Then
    With frmMain
        .mNewStaff.Visible = True
        .mNewUser.Visible = True
        .mBackupRestore.Visible = True
        .mEditSchoolInfo.Visible = True
        .mLogTrail.Visible = True
        .mReset.Visible = True
        .mReports.Visible = True
        .mUtilities.Visible = True
        .mViewUsers.Visible = True
        .mBackupRestore.Caption = "Backup/Restore Database"
    End With
    Call EnableDisable(frmStaff, frmStaff.tbrStaff, True)
    Call EnableDisable(frmStudents, frmStudents.tbrStudents, True)
    Call EnableDisable(frmUsers, frmUsers.tbrUsers, True)
    Call EnableDisable(frmViolations, frmViolations.tbrViolations, True)
    Call EnableDisable(frmContacts, frmContacts.tbrContacts, True)
    Call EnableDisable(frmContactGroups, frmContactGroups.tbrContactGroups, True)
    With frmBackupRestore
        .Caption = "Backup/Restore Database"
        .optRestore.Visible = True
    End With
ElseIf Privilege = "Administrator" Then
    With frmMain
        .mNewStaff.Visible = True
        .mNewUser.Visible = True
        .mBackupRestore.Visible = True
        .mEditSchoolInfo.Visible = True
        .mLogTrail.Visible = True
        .mReset.Visible = False
        .mReports.Visible = True
        .mUtilities.Visible = True
        .mViewUsers.Visible = True
        .mBackupRestore.Caption = "Backup Database"
    End With
    Call EnableDisable(frmStaff, frmStaff.tbrStaff, True)
    Call EnableDisable(frmStudents, frmStudents.tbrStudents, True)
    Call EnableDisable(frmUsers, frmUsers.tbrUsers, True)
    Call EnableDisable(frmViolations, frmViolations.tbrViolations, True)
    Call EnableDisable(frmContacts, frmContacts.tbrContacts, True)
    Call EnableDisable(frmContactGroups, frmContactGroups.tbrContactGroups, True)
    With frmBackupRestore
        .Caption = "Backup Database"
        .optRestore.Visible = False
    End With

ElseIf Privilege = "Staff" Then
    With frmMain
        .mNewStaff.Visible = False
        .mNewUser.Visible = False
        .mEditSchoolInfo.Visible = False
        .mReports.Visible = False
        .mUtilities.Visible = False
        .mViewUsers.Visible = False
    End With
    Call EnableDisable(frmStaff, frmStaff.tbrStaff, False)
    Call EnableDisable(frmStudents, frmStudents.tbrStudents, False)
    Call EnableDisable(frmUsers, frmUsers.tbrUsers, False)
    Call EnableDisable(frmViolations, frmViolations.tbrViolations, False)
    Call EnableDisable(frmContacts, frmContacts.tbrContacts, False)
    Call EnableDisable(frmContactGroups, frmContactGroups.tbrContactGroups, False)
End If
End Sub
Public Sub LoadStaff(lv As ListView)
Set rs = con.Execute("SELECT StaffID, FirstName+' '+MiddleName+' '+LastName AS FullName, Designation FROM Staff ORDER BY StaffID")
lv.ListItems.Clear

For xCount = 1 To rs.RecordCount
    With ls
        Set ls = lv.ListItems.Add(, , rs!StaffID)
        ls.SubItems(1) = rs!FullName
        ls.SubItems(2) = rs!Designation
        rs.MoveNext
    End With
Next xCount
End Sub
Public Sub LoadStudents(lv As ListView)
Set rs = con.Execute("SELECT * FROM Students ORDER BY StudentID")
lv.ListItems.Clear

For xCount = 1 To rs.RecordCount
    With ls
        Set ls = lv.ListItems.Add(, , rs!StudentID)
        ls.SubItems(1) = rs!FirstName & " " & rs!MiddleName & " " & rs!LastName
        ls.SubItems(2) = rs!SectionName
        ls.SubItems(3) = rs!StudentNumber
        rs.MoveNext
    End With
Next xCount
End Sub
Public Sub LoadLogTrail(lv As ListView)
Set rs = con.Execute("SELECT * FROM LogTrail ORDER BY LogID")
lv.ListItems.Clear

For xCount = 1 To rs.RecordCount
    With ls
        Set ls = lv.ListItems.Add(, , rs!LogID)
        ls.SubItems(1) = rs!LogDate
        ls.SubItems(2) = rs!TimeIn
        ls.SubItems(3) = IIf(IsNull(rs!TimeOut), "", rs!TimeOut)
        ls.SubItems(4) = rs!Username
        ls.SubItems(5) = rs!Privilege
        rs.MoveNext
    End With
Next xCount
End Sub
Public Sub LoadUsers(lv As ListView)
Set rs = con.Execute("SELECT UserID,Username, Privilege FROM Users ORDER BY UserID")
lv.ListItems.Clear

For xCount = 1 To rs.RecordCount
    With ls
        Set ls = lv.ListItems.Add(, , rs!UserID)
        ls.SubItems(1) = rs!Username
        ls.SubItems(2) = rs!Privilege
        rs.MoveNext
    End With
Next xCount
End Sub
Public Sub LoadViolations(lv As ListView)
Set rs = con.Execute("SELECT ViolationID,ViolationDate,Violation,Sanction,Students.StudentNumber,Students.SectionName,Students.FirstName+' '+Students.MiddleName+' '+Students.LastName AS FullName FROM Violations INNER JOIN Students ON Violations.VStudentID=Students.StudentID ORDER BY ViolationID")
lv.ListItems.Clear

For xCount = 1 To rs.RecordCount
    With ls
        Set ls = lv.ListItems.Add(, , rs!ViolationID)
        ls.SubItems(1) = rs!FullName
        ls.SubItems(2) = rs!StudentNumber
        ls.SubItems(3) = rs!SectionName
        ls.SubItems(4) = rs!Violation
        rs.MoveNext
    End With
Next xCount
End Sub
Public Sub LoadStudents2(lv As ListView)
Set rs = con.Execute("SELECT * FROM Students ORDER BY StudentID")
lv.ListItems.Clear

For xCount = 1 To rs.RecordCount
    With ls
        Set ls = lv.ListItems.Add(, , rs!StudentID)
        ls.SubItems(1) = rs!FirstName & " " & rs!MiddleName & " " & rs!LastName
        rs.MoveNext
    End With
Next xCount
End Sub
Public Sub LoadEvents(lv As ListView)
Set rs = con.Execute("SELECT * FROM EventsCalendar ORDER BY EventID,Sequence")
lv.ListItems.Clear

For xCount = 1 To rs.RecordCount
    With ls
        Set ls = lv.ListItems.Add(, , rs!EventID)
        ls.SubItems(1) = rs!EventDate
        ls.SubItems(2) = rs!TimeFrom
        ls.SubItems(3) = rs!TimeTo
        ls.SubItems(4) = rs!Topic
        ls.SubItems(5) = rs!Venue
        ls.SubItems(6) = rs!Details
        rs.MoveNext
    End With
Next xCount
End Sub
Public Sub LoadContactGroups(lv As ListView)
Set rs = con.Execute("SELECT * FROM ContactGroups WHERE LUserID=" & UserID & " ORDER BY GroupID")
lv.ListItems.Clear

For xCount = 1 To rs.RecordCount
    With ls
        Set ls = lv.ListItems.Add(, , rs!GroupID)
        ls.SubItems(1) = rs!GroupName
        ls.SubItems(2) = rs!Notes
        rs.MoveNext
    End With
Next xCount
End Sub
Public Sub LoadContacts(lv As ListView)
Set rs = con.Execute("SELECT ContactID,FirstName+' '+MiddleName+' '+LastName AS FullName,GroupName,PhoneNumber,MobileNumber,FaxNumber,EmailAddress FROM Contacts WHERE LUserID=" & UserID & " ORDER BY ContactID")
lv.ListItems.Clear

For xCount = 1 To rs.RecordCount
    With ls
        Set ls = lv.ListItems.Add(, , rs!ContactID)
        ls.SubItems(1) = rs!FullName
        ls.SubItems(2) = rs!GroupName
        rs.MoveNext
    End With
Next xCount
End Sub
Public Sub EnableDisable(frm As Form, tbr As Toolbar, isEnabled As Boolean)
With frm
    For xCount = 1 To 6
        tbr.Buttons(xCount).Visible = isEnabled
    Next xCount
End With
End Sub
