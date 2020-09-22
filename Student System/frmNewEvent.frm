VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmNewEvent 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "---"
   ClientHeight    =   5070
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6750
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmNewEvent.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5070
   ScaleWidth      =   6750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraEvent 
      Height          =   4935
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   6495
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   5160
         Picture         =   "frmNewEvent.frx":1082
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   3960
         Width           =   1095
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add Event"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   3840
         Picture         =   "frmNewEvent.frx":2104
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   3960
         Width           =   1335
      End
      Begin VB.TextBox txtVenue 
         Height          =   375
         Left            =   1800
         MaxLength       =   255
         TabIndex        =   11
         Top             =   2040
         Width           =   4455
      End
      Begin VB.TextBox txtTopic 
         Height          =   375
         Left            =   1800
         MaxLength       =   255
         TabIndex        =   10
         Top             =   2520
         Width           =   4455
      End
      Begin VB.TextBox txtDetails 
         Height          =   855
         Left            =   1800
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         Top             =   3000
         Width           =   4455
      End
      Begin MSComCtl2.DTPicker EventDate 
         Height          =   375
         Left            =   1800
         TabIndex        =   6
         Top             =   600
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         _Version        =   393216
         Format          =   53673985
         CurrentDate     =   40574
      End
      Begin MSComCtl2.DTPicker TimeFrom 
         Height          =   375
         Left            =   1800
         TabIndex        =   7
         Top             =   1080
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         _Version        =   393216
         Format          =   53673986
         CurrentDate     =   40574
      End
      Begin MSComCtl2.DTPicker TimeTo 
         Height          =   375
         Left            =   1800
         TabIndex        =   8
         Top             =   1560
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         _Version        =   393216
         Format          =   53673986
         CurrentDate     =   40574
      End
      Begin VB.Label lblSequence 
         Caption         =   "---"
         Height          =   375
         Left            =   3960
         TabIndex        =   18
         Top             =   240
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Label lblRecord 
         Caption         =   "Record #"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lblRecordID 
         Caption         =   "---"
         Height          =   255
         Left            =   1800
         TabIndex        =   16
         Top             =   240
         Width           =   4455
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "To"
         Height          =   375
         Left            =   1080
         TabIndex        =   13
         Top             =   1560
         Width           =   615
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "From"
         Height          =   375
         Left            =   1200
         TabIndex        =   12
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Topic"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   5
         Top             =   2520
         Width           =   615
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Details"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   3000
         Width           =   735
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Time"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   3
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Venue"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   2
         Top             =   2040
         Width           =   615
      End
      Begin VB.Label lblDate 
         Alignment       =   1  'Right Justify
         Caption         =   "Date"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   1
         Top             =   600
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmNewEvent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdAdd_Click()
If WarnBlank(txtTopic, "Topic") = True Then Exit Sub
If WarnBlank(txtDetails, "Details") = True Then Exit Sub
If cmdAdd.Caption = "Add Event" Then
    con.Execute ("INSERT INTO EventsCalendar VALUES(" & lblRecordID.Caption & ", " & lblSequence.Caption & ", #" & EventDate.Value & "#, #" & Format(TimeFrom.Value, "hh:mm:ss am/pm") & "#,#" & Format(TimeTo.Value, "hh:mm:ss am/pm") & "#, '" & Replace(txtVenue.Text, "'", "''") & "', '" & Replace(txtTopic.Text, "'", "''") & "', '" & Replace(txtDetails.Text, "'", "''") & "')")
    MsgBox "New event has been added successfully.", vbOKOnly + vbInformation, "Success"
ElseIf cmdAdd.Caption = "Update" Then
    con.Execute ("UPDATE EventsCalendar SET EventDate=#" & Format(EventDate.Value, "mm/dd/yyyy") & "#,TimeFrom=#" & Format(TimeFrom.Value, "hh:mm:ss am/pm") & "#,TimeTo=#" & Format(TimeTo.Value, "hh:mm:ss am/pm") & "#,Topic='" & Replace(txtTopic.Text, "'", "''") & "',Venue='" & Replace(txtVenue.Text, "'", "''") & "',Details='" & Replace(txtDetails.Text, "'", "''") & "' WHERE EventID=" & Val(lblRecordID.Caption) & " AND Sequence=" & Val(lblSequence.Caption) & "")
    MsgBox "Selected event has been updated successfully.", vbOKOnly + vbInformation, "Success"
End If
Unload Me
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub EventDate_Change()
Call CheckDates
End Sub

Private Sub EventDate_Click()
Call CheckDates
End Sub

Private Sub EventDate_DblClick()
Call CheckDates
End Sub

Private Sub EventDate_DropDown()
Call CheckDates
End Sub

Private Sub EventDate_KeyDown(KeyCode As Integer, Shift As Integer)
Call CheckDates
End Sub

Private Sub EventDate_KeyUp(KeyCode As Integer, Shift As Integer)
Call CheckDates
End Sub

Public Sub CheckDates()
Set rs = con.Execute("SELECT * FROM EventsCalendar ORDER BY EventID")
If rs.RecordCount = 0 Then
    lblRecordID.Caption = 1
    lblSequence.Caption = 1
Else
    rs.MoveLast
    lblRecordID.Caption = rs!EventID + 1
    Set rs = con.Execute("SELECT * FROM EventsCalendar WHERE EventDate=#" & EventDate.Value & "# ORDER BY Sequence")
    If rs.RecordCount = 0 Then
        lblSequence.Caption = 1
    Else
        rs.MoveLast
        lblSequence.Caption = rs!Sequence + 1
    End If
End If
End Sub

Private Sub Form_Activate()
If con.State = 0 Then Call konek
End Sub

