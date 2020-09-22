VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmNewStaff 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "---"
   ClientHeight    =   6975
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
   Icon            =   "frmNewStaff.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6975
   ScaleWidth      =   8895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraStudent 
      Height          =   6855
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   8655
      Begin VB.TextBox txtEmailAddress 
         Height          =   375
         Left            =   1800
         MaxLength       =   255
         TabIndex        =   32
         Top             =   6240
         Width           =   3735
      End
      Begin VB.TextBox txtAddress 
         Height          =   1095
         Left            =   1800
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   27
         Top             =   4200
         Width           =   3735
      End
      Begin VB.ComboBox cboCivilStatus 
         Height          =   315
         ItemData        =   "frmNewStaff.frx":1082
         Left            =   1800
         List            =   "frmNewStaff.frx":1098
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   3360
         Width           =   1935
      End
      Begin VB.TextBox txtReligion 
         Height          =   375
         Left            =   1800
         MaxLength       =   255
         TabIndex        =   24
         Top             =   3750
         Width           =   3735
      End
      Begin VB.TextBox txtLocation 
         Height          =   375
         Left            =   5880
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   3480
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "Browse"
         Height          =   375
         Left            =   6720
         TabIndex        =   21
         Top             =   3000
         Width           =   855
      End
      Begin VB.PictureBox pic 
         BackColor       =   &H00FFFFFF&
         Height          =   2295
         Left            =   5880
         ScaleHeight     =   2235
         ScaleWidth      =   2475
         TabIndex        =   20
         Top             =   600
         Width           =   2535
         Begin VB.Image img 
            Height          =   2295
            Left            =   0
            Stretch         =   -1  'True
            Top             =   0
            Width           =   2535
         End
      End
      Begin VB.CommandButton cmdRemove 
         Caption         =   "Remove"
         Height          =   375
         Left            =   7560
         TabIndex        =   19
         Top             =   3000
         Width           =   855
      End
      Begin VB.TextBox txtDesignation 
         Height          =   375
         Left            =   1800
         MaxLength       =   255
         TabIndex        =   6
         Top             =   2400
         Width           =   3735
      End
      Begin VB.TextBox txtLastName 
         Height          =   375
         Left            =   1800
         MaxLength       =   255
         TabIndex        =   3
         Top             =   1560
         Width           =   3735
      End
      Begin VB.TextBox txtMiddleName 
         Height          =   375
         Left            =   1800
         MaxLength       =   255
         TabIndex        =   2
         Top             =   1080
         Width           =   3735
      End
      Begin VB.TextBox txtFirstName 
         Height          =   375
         Left            =   1800
         MaxLength       =   255
         TabIndex        =   1
         Top             =   600
         Width           =   3735
      End
      Begin VB.TextBox txtContactNum 
         Height          =   735
         Left            =   1800
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   7
         Top             =   5400
         Width           =   3735
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add Staff"
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
         Left            =   5880
         Picture         =   "frmNewStaff.frx":10DA
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   5880
         Width           =   1335
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
         Left            =   7200
         Picture         =   "frmNewStaff.frx":215C
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   5880
         Width           =   1095
      End
      Begin VB.OptionButton optMale 
         Caption         =   "Male"
         Height          =   255
         Left            =   1800
         TabIndex        =   4
         Top             =   2040
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.OptionButton optFemale 
         Caption         =   "Female"
         Height          =   255
         Left            =   2640
         TabIndex        =   5
         Top             =   2040
         Width           =   1215
      End
      Begin MSComDlg.CommonDialog dlg 
         Left            =   6000
         Top             =   2880
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSComCtl2.DTPicker dtpBirthDate 
         Height          =   375
         Left            =   1800
         TabIndex        =   26
         Top             =   2880
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         Format          =   17039361
         CurrentDate     =   40568
      End
      Begin VB.Label Label12 
         Caption         =   "Email Address"
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
         TabIndex        =   33
         Top             =   6240
         Width           =   1575
      End
      Begin VB.Label Label8 
         Caption         =   "Address"
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
         TabIndex        =   31
         Top             =   4200
         Width           =   1575
      End
      Begin VB.Label Label13 
         Caption         =   "Date of Birth"
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
         TabIndex        =   30
         Top             =   2880
         Width           =   1335
      End
      Begin VB.Label Label14 
         Caption         =   "Civil Status"
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
         TabIndex        =   29
         Top             =   3360
         Width           =   1335
      End
      Begin VB.Label Label15 
         Caption         =   "Religion"
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
         TabIndex        =   28
         Top             =   3750
         Width           =   1575
      End
      Begin VB.Label Label9 
         Caption         =   "Image"
         Height          =   255
         Left            =   5880
         TabIndex        =   23
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label6 
         Caption         =   "Designation"
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
         Top             =   2400
         Width           =   1575
      End
      Begin VB.Label lblRecordID 
         Caption         =   "---"
         Height          =   255
         Left            =   1800
         TabIndex        =   17
         Top             =   240
         Width           =   3735
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
         TabIndex        =   16
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Name"
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
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Last"
         Height          =   255
         Left            =   1080
         TabIndex        =   14
         Top             =   1560
         Width           =   615
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Middle"
         Height          =   255
         Left            =   1080
         TabIndex        =   13
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "First"
         Height          =   255
         Left            =   1080
         TabIndex        =   12
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label5 
         Caption         =   "Gender"
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
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label Label7 
         Caption         =   "Contact Number"
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
         Top             =   5400
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frmNewStaff"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdAdd_Click()
Call ValidateAge(dtpBirthDate)
If WarnBlank(txtLastName, "Last Name") = True Then Exit Sub
If WarnBlank(txtFirstName, "First Name") Then Exit Sub
If WarnBlank(txtMiddleName, "Middle Name") = True Then Exit Sub
If WarnBlank(txtContactNum, "Contact Number") Then Exit Sub
If WarnBlank(txtAddress, "Address") = True Then Exit Sub
If WarnBlank(txtDesignation, "Designation") Then Exit Sub
If WarnBlank(txtReligion, "Religion") = True Then Exit Sub
Call Upload(txtLocation)
If Age < 18 Then
    MsgBox "Please select a valid age.", vbOKOnly + vbExclamation, "Invalid Age"
    Exit Sub
Else
    If cmdAdd.Caption = "Add Staff" Then
        con.Execute ("INSERT INTO Staff VALUES(" & lblRecordID.Caption & ", '" & Replace(txtLastName.Text, "'", "''") & "', '" & Replace(txtFirstName.Text, "'", "''") & "', '" & Replace(txtMiddleName.Text, "'", "''") & "', '" & IIf((optMale.Value = True), optMale.Caption, optFemale.Caption) & "', #" & dtpBirthDate.Value & "#," & Age & ",'" & Replace(txtDesignation.Text, "'", "''") & "', '" & cboCivilStatus.Text & "','" & Replace(txtReligion.Text, "'", "''") & "','" & Replace(txtContactNum.Text, "'", "''") & "','" & Replace(txtEmailAddress.Text, "'", "''") & "','" & Replace(txtAddress.Text, "'", "''") & "','" & IIf((FullPath = ""), "", FullPath) & "')")
        MsgBox "New staff successfully added.", vbOKOnly + vbInformation, "Success"
        Unload Me
    ElseIf cmdAdd.Caption = "Update" Then
        con.Execute ("UPDATE Staff SET LastName='" & Replace(txtLastName.Text, "'", "''") & "',FirstName='" & Replace(txtFirstName.Text, "'", "''") & "',MiddleName='" & Replace(txtMiddleName.Text, "'", "''") & "',Designation='" & Replace(txtDesignation.Text, "'", "''") & "',ContactNumber='" & Replace(txtContactNum.Text, "'", "''") & "',BirthDate=#" & dtpBirthDate.Value & "#,Age=" & Age & ",CivilStatus='" & cboCivilStatus.Text & "',Religion='" & Replace(txtReligion.Text, "'", "''") & "',EmailAddress='" & Replace(txtEmailAddress.Text, "'", "''") & "',Address='" & Replace(txtAddress.Text, "'", "''") & "',ImagePath='" & IIf((FullPath = ""), "", FullPath) & "' WHERE StaffID=" & Val(lblRecordID.Caption) & "")
        MsgBox "Staff information successfully updated.", vbOKOnly + vbInformation, "Success"
        Unload Me
    End If
End If
End Sub

Private Sub cmdBrowse_Click()
With dlg
    .Filter = "Pictures (*.bmp;*.jpg;*.gif)|*.bmp;*.jpg;*.gif"
    .ShowOpen
    txtLocation.Text = .FileName
    If txtLocation.Text <> "" Then Call loadLogo(txtLocation.Text, img, pic)
End With
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdRemove_Click()
Call ClearPic(txtLocation, dlg, img, pic)
End Sub

Private Sub Form_Activate()
If con.State = 0 Then Call konek
End Sub



