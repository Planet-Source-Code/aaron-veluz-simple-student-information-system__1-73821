VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmEvents 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Events List"
   ClientHeight    =   5670
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11895
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmEvents.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5670
   ScaleWidth      =   11895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraStaff 
      Height          =   4815
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   11655
      Begin VB.CheckBox chkSelect 
         Caption         =   "Select All"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   4440
         Width           =   1095
      End
      Begin VB.Frame fraSearch 
         Caption         =   "Search Events"
         Height          =   975
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   11295
         Begin VB.ComboBox cboCategory 
            Height          =   315
            ItemData        =   "frmEvents.frx":1082
            Left            =   1080
            List            =   "frmEvents.frx":1092
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
            ItemData        =   "frmEvents.frx":10B7
            Left            =   3360
            List            =   "frmEvents.frx":10C1
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   480
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.ComboBox cboCivilStatus 
            Height          =   315
            ItemData        =   "frmEvents.frx":10D3
            Left            =   3360
            List            =   "frmEvents.frx":10E9
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   480
            Visible         =   0   'False
            Width           =   1695
         End
         Begin MSComCtl2.DTPicker dtpFrom 
            Height          =   375
            Left            =   3960
            TabIndex        =   11
            Top             =   480
            Visible         =   0   'False
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   661
            _Version        =   393216
            Format          =   53608449
            CurrentDate     =   40567
         End
         Begin MSComCtl2.DTPicker dtpTo 
            Height          =   375
            Left            =   6360
            TabIndex        =   12
            Top             =   480
            Visible         =   0   'False
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   661
            _Version        =   393216
            Format          =   53608449
            CurrentDate     =   40567
         End
         Begin VB.Label lblFrom 
            Caption         =   "From"
            Height          =   255
            Left            =   3360
            TabIndex        =   14
            Top             =   480
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.Label lblTo 
            Caption         =   "To"
            Height          =   255
            Left            =   5880
            TabIndex        =   13
            Top             =   480
            Visible         =   0   'False
            Width           =   375
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
      Begin MSComctlLib.ListView lvEvents 
         Height          =   2895
         Left            =   120
         TabIndex        =   10
         Top             =   1320
         Width           =   11295
         _ExtentX        =   19923
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
         Left            =   1320
         TabIndex        =   9
         Top             =   4440
         Width           =   10095
      End
   End
   Begin MSComctlLib.Toolbar tbrStaff 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11895
      _ExtentX        =   20981
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
         Left            =   6840
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
               Picture         =   "frmEvents.frx":112B
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEvents.frx":21BD
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEvents.frx":324F
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEvents.frx":42E1
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "frmEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cboCategory_Click()
If cboCategory.Text = "Date" Then
    txtSearch.Visible = False
    dtpFrom.Visible = True
    dtpTo.Visible = True
    lblFrom.Visible = True
    lblTo.Visible = True
ElseIf cboCategory.Text = "ALL RECORDS" Then
    txtSearch.Visible = False
    dtpFrom.Visible = False
    dtpTo.Visible = False
    lblFrom.Visible = False
    lblTo.Visible = False
    Call LoadEvents(lvEvents)
    Call CheckCount(lvEvents, lblRecord)
Else
    txtSearch.Visible = True
    dtpFrom.Visible = False
    dtpTo.Visible = False
    lblFrom.Visible = False
    lblTo.Visible = False
End If
End Sub

Private Sub chkSelect_Click()
Call SelectAll(lvEvents, chkSelect)
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
Call LoadEvents(lvEvents)
Call CheckCount(lvEvents, lblRecord)
cboCategory.ListIndex = 3
Call cboCategory_Click
End Sub

Private Sub lvEvents_Click()
Call CheckCount(lvEvents, lblRecord)
End Sub

Private Sub lvEvents_DblClick()
If lvEvents.ListItems.Count <> 0 Then
    Set rs = con.Execute("SELECT * FROM EventsCalendar WHERE EventID=" & Val(lvEvents.SelectedItem) & "")
    With frmNewEvent
        .lblRecordID.Caption = lvEvents.SelectedItem
        .lblSequence.Caption = rs!Sequence
        .TimeFrom.Value = rs!TimeFrom
        .TimeTo.Value = rs!TimeTo
        .txtTopic.Text = rs!Topic
        .txtDetails.Text = rs!Details
        .txtVenue.Text = rs!Venue
        .cmdAdd.Caption = "Update"
        .Show vbModal, Me
    End With
End If
Call LoadEvents(lvEvents)
Call CheckCount(lvEvents, lblRecord)
End Sub

Private Sub lvEvents_KeyDown(KeyCode As Integer, Shift As Integer)
Call CheckCount(lvEvents, lblRecord)
End Sub

Private Sub lvEvents_KeyUp(KeyCode As Integer, Shift As Integer)
Call CheckCount(lvEvents, lblRecord)
End Sub

Private Sub tbrStaff_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
    Case 1 'Add New
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
        Call LoadEvents(lvEvents)
        Call CheckCount(lvEvents, lblRecord)
    Case 3 'Delete
        If lvEvents.ListItems.Count = 0 Then Exit Sub
        Call CountSelected(lvEvents)
        If yCount <> 0 Then
            If MsgBox("Are you sure you want to delete the selected item(s)?", vbYesNo + vbExclamation, "Confirm Delete") = vbNo Then Exit Sub
            For xCount = 1 To lvEvents.ListItems.Count
                If lvEvents.ListItems(xCount).Checked = True Then
                    CurrRec = lvEvents.ListItems(xCount)
                    con.Execute ("DELETE FROM EventsCalendar WHERE EventID = " & CurrRec & "")
                End If
            Next xCount
            Call LoadEvents(lvEvents)
            Call CheckCount(lvEvents, lblRecord)
        End If
    Case 5 'Print
        If lvEvents.ListItems.Count = 0 Then Exit Sub
        Call CountSelected(lvEvents)
        If yCount = 1 Then
            For xCount = 1 To lvEvents.ListItems.Count
                If lvEvents.ListItems(xCount).Checked = True Then
                    CurrRec = lvEvents.ListItems(xCount)
                    Set rs = con.Execute("SELECT * FROM EventsCalendar WHERE EventID=" & CurrRec & "")
                End If
            Next xCount
            With rptSpecificEvent
                Set .DataSource = rs
                .Caption = "Event Info "
                .Show vbModal, Me
            End With
        ElseIf yCount > 1 And yCount <> lvEvents.ListItems.Count Then
            inPart = ""
            inPart2 = ""
            inWhole = ""
            For xCount = 1 To lvEvents.ListItems.Count
                If lvEvents.ListItems(xCount).Checked = True Then
                    inPart = lvEvents.ListItems(xCount) & ", "
                    inPart2 = inPart2 & inPart
                End If
                inWhole = "IN ( " & inPart2 & ")"
            Next xCount
            Set rs = con.Execute("SELECT * FROM EventsCalendar WHERE EventID " & inWhole & " ORDER BY EventID")
            With rptEvents
                Set .DataSource = rs
                .Orientation = rptOrientLandscape
                .Show vbModal, Me
            End With
        Else
            Set rs = con.Execute("SELECT * FROM EventsCalendar ORDER BY EventID, Sequence")

            With rptEvents
                Set .DataSource = rs
                .Orientation = rptOrientLandscape
                .Show vbModal, Me
            End With
        End If
    Case 7 'Show All
        cboCategory.ListIndex = 3
        Call cboCategory_Click
End Select
End Sub

Private Sub txtSearch_Change()
strTest = cboCategory.Text
If txtSearch.Text <> "" Then
    Set rs = con.Execute("SELECT * FROM EventsCalendar WHERE " & strTest & " LIKE '" & txtSearch.Text & "%' ORDER BY EventID,Sequence")
    lvEvents.ListItems.Clear
    For xCount = 1 To rs.RecordCount
        With ls
            Set ls = lvEvents.ListItems.Add(, , rs!EventID)
            ls.SubItems(1) = rs!EventDate
            ls.SubItems(2) = rs!TimeFrom
            ls.SubItems(3) = rs!TimeTo
            ls.SubItems(4) = rs!Topic
            ls.SubItems(5) = rs!Venue
            ls.SubItems(6) = rs!Details
            rs.MoveNext
        End With
    Next xCount
    Call CheckCount(lvEvents, lblRecord)
Else
    Call LoadEvents(lvEvents)
    Call CheckCount(lvEvents, lblRecord)
End If
End Sub
Public Sub CheckDates()
If lvEvents.ListItems.Count = 0 Then Exit Sub
Set rs = con.Execute("SELECT * FROM EventsCalendar WHERE EventDate BETWEEN #" & dtpFrom.Value & "# AND #" & dtpTo.Value & "# ORDER BY EventID")
lvEvents.ListItems.Clear

For xCount = 1 To rs.RecordCount
    With ls
        Set ls = lvEvents.ListItems.Add(, , rs!EventID)
        ls.SubItems(1) = rs!EventDate
        ls.SubItems(2) = rs!TimeFrom
        ls.SubItems(3) = rs!TimeTo
        ls.SubItems(4) = rs!Topic
        ls.SubItems(5) = rs!Venue
        ls.SubItems(6) = rs!Details
        rs.MoveNext
    End With
Next xCount
Call CheckCount(lvEvents, lblRecord)
End Sub
