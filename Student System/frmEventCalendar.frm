VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Begin VB.Form frmEventCalendar 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Calendar of Activities per Day"
   ClientHeight    =   4215
   ClientLeft      =   1230
   ClientTop       =   1605
   ClientWidth     =   10560
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmEventCalendar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   10560
   ShowInTaskbar   =   0   'False
   Begin MSACAL.Calendar Calendar1 
      Height          =   3975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4575
      _Version        =   524288
      _ExtentX        =   8070
      _ExtentY        =   7011
      _StockProps     =   1
      BackColor       =   -2147483633
      Year            =   2011
      Month           =   1
      Day             =   31
      DayLength       =   1
      MonthLength     =   2
      DayFontColor    =   0
      FirstDay        =   7
      GridCellEffect  =   1
      GridFontColor   =   10485760
      GridLinesColor  =   -2147483632
      ShowDateSelectors=   -1  'True
      ShowDays        =   -1  'True
      ShowHorizontalGrid=   -1  'True
      ShowTitle       =   -1  'True
      ShowVerticalGrid=   -1  'True
      TitleFontColor  =   10485760
      ValueIsNull     =   0   'False
      BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame fraRecords 
      Caption         =   "---"
      Height          =   4095
      Left            =   4800
      TabIndex        =   15
      Top             =   0
      Visible         =   0   'False
      Width           =   5655
      Begin VB.CommandButton cmdBack 
         Caption         =   "<<<"
         Height          =   375
         Left            =   240
         TabIndex        =   21
         Top             =   360
         Width           =   615
      End
      Begin VB.CheckBox chkSelect 
         Caption         =   "Select All"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   3600
         Width           =   1095
      End
      Begin VB.CommandButton cmdAddNew 
         Caption         =   "Add New"
         Height          =   375
         Left            =   2760
         TabIndex        =   17
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Delete"
         Height          =   375
         Left            =   4080
         TabIndex        =   16
         Top             =   360
         Width           =   1335
      End
      Begin MSComctlLib.ListView lvEvents 
         Height          =   2655
         Left            =   240
         TabIndex        =   19
         Top             =   840
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   4683
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
            Text            =   "Event ID"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Topic"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Details"
            Object.Width           =   3528
         EndProperty
      End
      Begin VB.Label lblRecord 
         Alignment       =   1  'Right Justify
         Caption         =   "---"
         Height          =   255
         Left            =   1440
         TabIndex        =   20
         Top             =   3600
         Width           =   3975
      End
   End
   Begin VB.Frame fraEvent 
      Caption         =   "---"
      Height          =   4095
      Left            =   4800
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   5655
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add Event"
         Default         =   -1  'True
         Height          =   495
         Left            =   3000
         TabIndex        =   13
         Top             =   3360
         Width           =   1335
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   495
         Left            =   4320
         TabIndex        =   14
         Top             =   3360
         Width           =   1095
      End
      Begin VB.TextBox txtDetails 
         Height          =   855
         Left            =   960
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   2280
         Width           =   4455
      End
      Begin VB.TextBox txtTopic 
         Height          =   375
         Left            =   960
         MaxLength       =   255
         TabIndex        =   4
         Top             =   1800
         Width           =   4455
      End
      Begin VB.TextBox txtVenue 
         Height          =   375
         Left            =   960
         MaxLength       =   255
         TabIndex        =   3
         Top             =   1320
         Width           =   4455
      End
      Begin MSComCtl2.DTPicker TimeFrom 
         Height          =   375
         Left            =   1560
         TabIndex        =   9
         Top             =   360
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         _Version        =   393216
         Format          =   16973826
         CurrentDate     =   40574
      End
      Begin MSComCtl2.DTPicker TimeTo 
         Height          =   375
         Left            =   1560
         TabIndex        =   10
         Top             =   840
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         _Version        =   393216
         Format          =   16973826
         CurrentDate     =   40574
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "From"
         Height          =   375
         Left            =   960
         TabIndex        =   12
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "To"
         Height          =   375
         Left            =   840
         TabIndex        =   11
         Top             =   840
         Width           =   615
      End
      Begin VB.Label lblDetails 
         Alignment       =   1  'Right Justify
         Caption         =   "Details"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   2280
         Width           =   735
      End
      Begin VB.Label lblVenue 
         Alignment       =   1  'Right Justify
         Caption         =   "Venue"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label lblTopic 
         Alignment       =   1  'Right Justify
         Caption         =   "Topic"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   1800
         Width           =   615
      End
      Begin VB.Label lblTime 
         Caption         =   "Time"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   2
         Top             =   360
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmEventCalendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Calendar1_Click()
Me.Width = 4860
Set rs = con.Execute("SELECT * FROM EventsCalendar WHERE EventDate=#" & Calendar1.Value & "# ORDER BY EventID")

If rs.RecordCount = 0 Then
    If MsgBox("No records found for the selected date. Would you like to add one now?", vbYesNo + vbInformation, "No Records") = vbNo Then Exit Sub
    With frmNewEvent
        Set rs = con.Execute("SELECT * FROM EventsCalendar ORDER BY EventID")
        If rs.RecordCount = 0 Then
            .lblRecordID.Caption = 1
            .lblSequence.Caption = 1
        Else
            rs.MoveLast
            .lblRecordID.Caption = rs!EventID + 1
            Set rs = con.Execute("SELECT * FROM EventsCalendar WHERE EventDate=#" & Calendar1.Value & "# ORDER BY Sequence")
            If rs.RecordCount = 0 Then
                .lblSequence.Caption = 1
            Else
                rs.MoveLast
                .lblSequence.Caption = rs!Sequence + 1
            End If
        End If
        .EventDate.Value = Calendar1.Value
        .Show vbModal, Me
    End With
    Call LoadEventsinDate
    Call CheckCount(lvEvents, lblRecord)
    Call Calendar1_Click
Else
    Me.Width = 10680
    fraRecords.Caption = "Event(s) for " & Calendar1.Value
    fraRecords.Visible = True
    fraEvent.Visible = False
    Call LoadEventsinDate
    Call CheckCount(lvEvents, lblRecord)
End If
End Sub

Private Sub chkSelect_Click()
Call SelectAll(lvEvents, chkSelect)
End Sub

Private Sub cmdAdd_Click()
If WarnBlank(txtTopic, "Topic") = True Then Exit Sub
If WarnBlank(txtDetails, "Details") = True Then Exit Sub
If cmdAdd.Caption = "Add Event" Then
    Set rs = con.Execute("SELECT * FROM EventsCalendar ORDER BY EventID")
    If rs.RecordCount = 0 Then
        xCount = 1
        yCount = 1
    Else
        rs.MoveLast
        xCount = rs!EventID + 1
        Set rs = con.Execute("SELECT * FROM EventsCalendar WHERE EventDate=#" & Calendar1.Value & "# ORDER BY Sequence")
        If rs.RecordCount = 0 Then
            yCount = 1
        Else
            rs.MoveLast
            yCount = rs!Sequence + 1
        End If
    End If
    con.Execute ("INSERT INTO EventsCalendar VALUES(" & xCount & ", " & yCount & ", #" & Calendar1.Value & "#, #" & Format(TimeFrom.Value, "hh:mm:ss am/pm") & "#,#" & Format(TimeTo.Value, "hh:mm:ss am/pm") & "#, '" & txtVenue.Text & "', '" & txtTopic.Text & "', '" & txtDetails.Text & "')")
    MsgBox "New event has been added successfully.", vbOKOnly + vbInformation, "Success"
    Call ClearObjects
    Call LoadEventsinDate
    Call CheckCount(lvEvents, lblRecord)
    Call cmdCancel_Click
ElseIf cmdAdd.Caption = "Update" Then
    con.Execute ("UPDATE EventsCalendar SET TimeFrom=#" & TimeFrom.Value & "#, TimeTo=#" & TimeTo.Value & "#, Topic='" & txtTopic.Text & "',Venue='" & txtVenue.Text & "', Details='" & txtDetails.Text & "' WHERE EventDate=#" & Calendar1.Value & "# AND Sequence=" & lvEvents.SelectedItem & "")
    MsgBox "Event has been updated successfully.", vbOKOnly + vbInformation, "Success"
    Call ClearObjects
    Call LoadEventsinDate
    Call CheckCount(lvEvents, lblRecord)
    Call cmdCancel_Click
End If
End Sub

Private Sub cmdAddNew_Click()
With frmNewEvent
    Set rs = con.Execute("SELECT * FROM EventsCalendar ORDER BY EventID")
    If rs.RecordCount = 0 Then
        .lblRecordID.Caption = 1
        .lblSequence.Caption = 1
    Else
        rs.MoveLast
        .lblRecordID.Caption = rs!EventID + 1
        Set rs = con.Execute("SELECT * FROM EventsCalendar WHERE EventDate=#" & Calendar1.Value & "# ORDER BY Sequence")
        If rs.RecordCount = 0 Then
            .lblSequence.Caption = 1
        Else
            rs.MoveLast
            .lblSequence.Caption = rs!Sequence + 1
        End If
    End If
    .cmdAdd.Caption = "Add Event"
    .EventDate.Value = Format(Calendar1.Value, "mm/dd/yyyy")
    .Show vbModal, Me
End With
Call LoadEventsinDate
Call CheckCount(lvEvents, lblRecord)
End Sub

Private Sub cmdBack_Click()
Me.Width = 4860
End Sub

Private Sub cmdCancel_Click()
fraEvent.Visible = False
fraRecords.Visible = True
End Sub

Private Sub cmdDelete_Click()
If lvEvents.ListItems.Count = 0 Then Exit Sub
Call CountSelected(lvEvents)
If yCount <> 0 Then
    If MsgBox("Are you sure you want to delete the selected item(s)?", vbYesNo + vbExclamation, "Confirm Delete") = vbNo Then Exit Sub
    For xCount = 1 To lvEvents.ListItems.Count
        If lvEvents.ListItems(xCount).Checked = True Then
            con.Execute ("DELETE FROM EventsCalendar WHERE EventDate=#" & Calendar1.Value & "# AND Sequence=" & Val(lvEvents.ListItems(xCount)) & "")
        End If
    Next xCount
    Call LoadEventsinDate
    Call CheckCount(lvEvents, lblRecord)
End If
End Sub

Private Sub Form_Activate()
If con.State = 0 Then Call konek
End Sub

Public Sub LoadEventsinDate()
Set rs = con.Execute("SELECT * FROM EventsCalendar WHERE EventDate=#" & Calendar1.Value & "# ORDER BY EventID,Sequence")
lvEvents.ListItems.Clear

For xCount = 1 To rs.RecordCount
    With ls
        Set ls = lvEvents.ListItems.Add(, , rs!Sequence)
        ls.SubItems(1) = rs!Topic
        ls.SubItems(2) = rs!Details
        rs.MoveNext
    End With
Next xCount
End Sub

Private Sub Form_Load()
Me.Width = 4860
Calendar1.Value = Date
End Sub

Private Sub lvEvents_Click()
Call CheckCount(lvEvents, lblRecord)
End Sub

Private Sub lvEvents_DblClick()
If lvEvents.ListItems.Count <> 0 Then
    Set rs = con.Execute("SELECT * FROM EventsCalendar WHERE EventDate=#" & Calendar1.Value & "# AND Sequence=" & Val(lvEvents.SelectedItem) & "")
    With rs
        TimeFrom.Value = !TimeFrom
        TimeTo.Value = !TimeTo
        txtTopic.Text = !Topic
        txtDetails.Text = !Details
        txtVenue.Text = !Venue
    End With
    cmdAdd.Caption = "Update"
End If
fraRecords.Visible = False
fraEvent.Visible = True
fraEvent.Caption = "Edit Event for " & Calendar1.Value
End Sub

Private Sub lvEvents_KeyDown(KeyCode As Integer, Shift As Integer)
Call CheckCount(lvEvents, lblRecord)
End Sub

Private Sub lvEvents_KeyUp(KeyCode As Integer, Shift As Integer)
Call CheckCount(lvEvents, lblRecord)
End Sub
Public Sub ClearObjects()
txtTopic.Text = ""
txtVenue.Text = ""
txtDetails.Text = ""
End Sub
