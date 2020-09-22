VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmNewViolation 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "---"
   ClientHeight    =   7695
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
   Icon            =   "frmNewViolation.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7695
   ScaleWidth      =   8415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraStudent 
      Height          =   7575
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   8175
      Begin VB.TextBox txtSanction 
         Height          =   735
         Left            =   2640
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   5760
         Width           =   5415
      End
      Begin MSComCtl2.DTPicker dtDate 
         Height          =   375
         Left            =   2640
         TabIndex        =   1
         Top             =   600
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   661
         _Version        =   393216
         Format          =   17039361
         CurrentDate     =   40564
      End
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
         Left            =   6960
         Picture         =   "frmNewViolation.frx":1082
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   6600
         Width           =   1095
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add Violation"
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
         Left            =   5280
         Picture         =   "frmNewViolation.frx":2104
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   6600
         Width           =   1695
      End
      Begin VB.TextBox txtViolation 
         Height          =   735
         Left            =   2640
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   4920
         Width           =   5415
      End
      Begin MSComctlLib.ListView lvStudents 
         Height          =   2055
         Left            =   2640
         TabIndex        =   3
         Top             =   1560
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   3625
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
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Student ID"
            Object.Width           =   2999
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Name"
            Object.Width           =   6526
         EndProperty
      End
      Begin MSComCtl2.DTPicker dtTime 
         Height          =   375
         Left            =   2640
         TabIndex        =   2
         Top             =   1080
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "hh:mm am/pm"
         Format          =   17039362
         CurrentDate     =   40564
      End
      Begin VB.Label lblGender 
         Caption         =   "---"
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
         Left            =   2640
         TabIndex        =   21
         Top             =   4560
         Width           =   5415
      End
      Begin VB.Label Label8 
         Caption         =   "Gender:"
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
         TabIndex        =   20
         Top             =   4560
         Width           =   975
      End
      Begin VB.Label lblStudentID 
         Caption         =   "---"
         Height          =   255
         Left            =   2640
         TabIndex        =   19
         Top             =   6720
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label6 
         Caption         =   "Sanction"
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
         TabIndex        =   18
         Top             =   5760
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "Date:"
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
         Top             =   720
         Width           =   2415
      End
      Begin VB.Label Label2 
         Caption         =   "Time:"
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
         TabIndex        =   16
         Top             =   1080
         Width           =   2415
      End
      Begin VB.Label Label5 
         Caption         =   "Student #:"
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
         TabIndex        =   15
         Top             =   3840
         Width           =   1935
      End
      Begin VB.Label lblStudentNumber 
         Caption         =   "---"
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
         Left            =   2640
         TabIndex        =   14
         Top             =   3840
         Width           =   5415
      End
      Begin VB.Label Label3 
         Caption         =   "Section:"
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
         TabIndex        =   13
         Top             =   4200
         Width           =   975
      End
      Begin VB.Label lblSection 
         Caption         =   "---"
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
         Left            =   2640
         TabIndex        =   12
         Top             =   4200
         Width           =   5415
      End
      Begin VB.Label Label1 
         Caption         =   "Select Student from List"
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
         TabIndex        =   11
         Top             =   1560
         Width           =   2415
      End
      Begin VB.Label Label7 
         Caption         =   "Violation"
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
         TabIndex        =   10
         Top             =   4920
         Width           =   1575
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
         TabIndex        =   9
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lblRecordID 
         Caption         =   "---"
         Height          =   255
         Left            =   2640
         TabIndex        =   8
         Top             =   240
         Width           =   5415
      End
   End
End
Attribute VB_Name = "frmNewViolation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdAdd_Click()
If WarnBlank(txtViolation, "Violation Name") = True Then Exit Sub
If WarnBlank(txtSanction, "Sanction") = True Then Exit Sub
If lblSection.Caption = "---" Then
    MsgBox "Select a student from the list.", vbOKOnly + vbExclamation, "No Student Selected"
    Exit Sub
Else
    If cmdAdd.Caption = "Add Violation" Then
            con.Execute ("INSERT INTO Violations VALUES('" & lblRecordID.Caption & "', #" & Format(dtDate.Value, "mm/dd/yyyy") & "#, #" & Format(dtTime.Value, "hh:mm:ss AM/PM") & "#,'" & Replace(txtViolation.Text, "'", "''") & "'," & lvStudents.SelectedItem.Text & ", '" & Replace(txtSanction.Text, "'", "''") & "')")
            MsgBox "New student violation successfully added.", vbOKOnly + vbInformation, "Success"
            Unload Me
    ElseIf cmdAdd.Caption = "Update" Then
        con.Execute ("UPDATE Violations SET ViolationDate=#" & dtDate.Value & "#,ViolationTime=#" & dtTime.Value & "#,Violation='" & Replace(txtViolation.Text, "'", "''") & "',Sanction='" & Replace(txtSanction.Text, "'", "''") & "' WHERE ViolationID=" & Val(lblRecordID.Caption) & "")
        MsgBox "Student violation successfully updated.", vbOKOnly + vbInformation, "Success"
        Unload Me
    End If
End If
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub Form_Activate()
If con.State = 0 Then Call konek
End Sub

Private Sub lvStudents_Click()
If lvStudents.ListItems.Count >= 1 Then
    Set rs = con.Execute("SELECT * FROM Students WHERE StudentID=" & lvStudents.SelectedItem.Text & "")
    With rs
        lblSection.Caption = !SectionName
        lblStudentNumber.Caption = !StudentNumber
        lblStudentID.Caption = !StudentID
        lblGender.Caption = !Gender
    End With
End If
End Sub
