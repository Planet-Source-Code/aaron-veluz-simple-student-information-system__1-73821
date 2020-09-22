VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmNewContact 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "---"
   ClientHeight    =   5430
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11175
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmNewContact.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5430
   ScaleWidth      =   11175
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraStudent 
      Height          =   5295
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   10935
      Begin TabDlg.SSTab SSTab1 
         Height          =   4455
         Left            =   240
         TabIndex        =   11
         Top             =   600
         Width           =   7575
         _ExtentX        =   13361
         _ExtentY        =   7858
         _Version        =   393216
         TabHeight       =   520
         TabCaption(0)   =   "Basic Information"
         TabPicture(0)   =   "frmNewContact.frx":1082
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label4"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Label3"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Label2"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "Label1"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "Label13"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "Label5"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "Label16"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "dtpBirthDate"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "txtFirstName"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).Control(9)=   "txtMiddleName"
         Tab(0).Control(9).Enabled=   0   'False
         Tab(0).Control(10)=   "txtLastName"
         Tab(0).Control(10).Enabled=   0   'False
         Tab(0).Control(11)=   "optFemale"
         Tab(0).Control(11).Enabled=   0   'False
         Tab(0).Control(12)=   "optMale"
         Tab(0).Control(12).Enabled=   0   'False
         Tab(0).Control(13)=   "cboContactGroup"
         Tab(0).Control(13).Enabled=   0   'False
         Tab(0).ControlCount=   14
         TabCaption(1)   =   "Contact Information"
         TabPicture(1)   =   "frmNewContact.frx":109E
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Label10"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "Label17"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).Control(2)=   "Label18"
         Tab(1).Control(2).Enabled=   0   'False
         Tab(1).Control(3)=   "Label20"
         Tab(1).Control(3).Enabled=   0   'False
         Tab(1).Control(4)=   "Label8"
         Tab(1).Control(4).Enabled=   0   'False
         Tab(1).Control(5)=   "txtPhoneNumber"
         Tab(1).Control(5).Enabled=   0   'False
         Tab(1).Control(6)=   "txtMobileNumber"
         Tab(1).Control(6).Enabled=   0   'False
         Tab(1).Control(7)=   "txtFaxNumber"
         Tab(1).Control(7).Enabled=   0   'False
         Tab(1).Control(8)=   "txtEmailAddress"
         Tab(1).Control(8).Enabled=   0   'False
         Tab(1).Control(9)=   "txtAddress"
         Tab(1).Control(9).Enabled=   0   'False
         Tab(1).ControlCount=   10
         TabCaption(2)   =   "Others"
         TabPicture(2)   =   "frmNewContact.frx":10BA
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Label15"
         Tab(2).Control(0).Enabled=   0   'False
         Tab(2).Control(1)=   "Label14"
         Tab(2).Control(1).Enabled=   0   'False
         Tab(2).Control(2)=   "lblPassword"
         Tab(2).Control(2).Enabled=   0   'False
         Tab(2).Control(3)=   "Label6"
         Tab(2).Control(3).Enabled=   0   'False
         Tab(2).Control(4)=   "lblStudentNumber"
         Tab(2).Control(4).Enabled=   0   'False
         Tab(2).Control(5)=   "txtReligion"
         Tab(2).Control(5).Enabled=   0   'False
         Tab(2).Control(6)=   "cboCivilStatus"
         Tab(2).Control(6).Enabled=   0   'False
         Tab(2).Control(7)=   "txtNotes"
         Tab(2).Control(7).Enabled=   0   'False
         Tab(2).Control(8)=   "txtCompanySchool"
         Tab(2).Control(8).Enabled=   0   'False
         Tab(2).Control(9)=   "txtPositionDesignation"
         Tab(2).Control(9).Enabled=   0   'False
         Tab(2).ControlCount=   10
         Begin VB.TextBox txtPositionDesignation 
            Height          =   375
            Left            =   -72240
            MaxLength       =   255
            TabIndex        =   44
            Top             =   1680
            Width           =   4335
         End
         Begin VB.TextBox txtCompanySchool 
            Height          =   855
            Left            =   -72240
            MaxLength       =   255
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   43
            Top             =   2160
            Width           =   4335
         End
         Begin VB.TextBox txtNotes 
            Height          =   855
            IMEMode         =   3  'DISABLE
            Left            =   -72240
            MaxLength       =   255
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   41
            Top             =   3120
            Width           =   4335
         End
         Begin VB.ComboBox cboCivilStatus 
            Height          =   315
            ItemData        =   "frmNewContact.frx":10D6
            Left            =   -72240
            List            =   "frmNewContact.frx":10EC
            Style           =   2  'Dropdown List
            TabIndex        =   38
            Top             =   840
            Width           =   1935
         End
         Begin VB.TextBox txtReligion 
            Height          =   375
            Left            =   -72240
            MaxLength       =   255
            TabIndex        =   37
            Top             =   1200
            Width           =   4335
         End
         Begin VB.TextBox txtAddress 
            Height          =   1095
            Left            =   -72360
            MaxLength       =   255
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   35
            Top             =   2760
            Width           =   4455
         End
         Begin VB.TextBox txtEmailAddress 
            Height          =   375
            Left            =   -72360
            MaxLength       =   255
            TabIndex        =   33
            Top             =   2280
            Width           =   4455
         End
         Begin VB.TextBox txtFaxNumber 
            Height          =   375
            Left            =   -72360
            MaxLength       =   255
            TabIndex        =   30
            Top             =   1800
            Width           =   4455
         End
         Begin VB.TextBox txtMobileNumber 
            Height          =   375
            Left            =   -72360
            MaxLength       =   255
            TabIndex        =   28
            Top             =   1320
            Width           =   4455
         End
         Begin VB.TextBox txtPhoneNumber 
            Height          =   375
            Left            =   -72360
            MaxLength       =   255
            TabIndex        =   26
            Top             =   840
            Width           =   4455
         End
         Begin VB.ComboBox cboContactGroup 
            Height          =   315
            ItemData        =   "frmNewContact.frx":112E
            Left            =   2640
            List            =   "frmNewContact.frx":1130
            Style           =   2  'Dropdown List
            TabIndex        =   24
            Top             =   3120
            Width           =   1935
         End
         Begin VB.OptionButton optMale 
            Caption         =   "Male"
            Height          =   255
            Left            =   2640
            TabIndex        =   21
            Top             =   2280
            Value           =   -1  'True
            Width           =   735
         End
         Begin VB.OptionButton optFemale 
            Caption         =   "Female"
            Height          =   255
            Left            =   3480
            TabIndex        =   20
            Top             =   2280
            Width           =   1215
         End
         Begin VB.TextBox txtLastName 
            Height          =   375
            Left            =   2640
            MaxLength       =   255
            TabIndex        =   14
            Top             =   1800
            Width           =   4455
         End
         Begin VB.TextBox txtMiddleName 
            Height          =   375
            Left            =   2640
            MaxLength       =   255
            TabIndex        =   13
            Top             =   1320
            Width           =   4455
         End
         Begin VB.TextBox txtFirstName 
            Height          =   375
            Left            =   2640
            MaxLength       =   255
            TabIndex        =   12
            Top             =   840
            Width           =   4455
         End
         Begin MSComCtl2.DTPicker dtpBirthDate 
            Height          =   375
            Left            =   2640
            TabIndex        =   19
            Top             =   2640
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   661
            _Version        =   393216
            Format          =   53739521
            CurrentDate     =   40568
         End
         Begin VB.Label lblStudentNumber 
            Caption         =   "Position/Designation"
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
            Left            =   -74520
            TabIndex        =   46
            Top             =   1680
            Width           =   2055
         End
         Begin VB.Label Label6 
            Caption         =   "Company/School"
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
            Left            =   -74520
            TabIndex        =   45
            Top             =   2160
            Width           =   1815
         End
         Begin VB.Label lblPassword 
            Caption         =   "Notes"
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
            Left            =   -74520
            TabIndex        =   42
            Top             =   3120
            Width           =   1095
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
            Left            =   -74520
            TabIndex        =   40
            Top             =   840
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
            Left            =   -74520
            TabIndex        =   39
            Top             =   1230
            Width           =   1575
         End
         Begin VB.Label Label8 
            Caption         =   "Home Address"
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
            Left            =   -74520
            TabIndex        =   36
            Top             =   2760
            Width           =   1575
         End
         Begin VB.Label Label20 
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
            Left            =   -74520
            TabIndex        =   34
            Top             =   2280
            Width           =   1575
         End
         Begin VB.Label Label18 
            Caption         =   "Fax Number"
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
            Left            =   -74520
            TabIndex        =   31
            Top             =   1800
            Width           =   1575
         End
         Begin VB.Label Label17 
            Caption         =   "Mobile Number"
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
            Left            =   -74520
            TabIndex        =   29
            Top             =   1320
            Width           =   1575
         End
         Begin VB.Label Label10 
            Caption         =   "Phone Number"
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
            Left            =   -74520
            TabIndex        =   27
            Top             =   840
            Width           =   1575
         End
         Begin VB.Label Label16 
            Caption         =   "Group Name"
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
            Left            =   480
            TabIndex        =   25
            Top             =   3120
            Width           =   1335
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
            Left            =   480
            TabIndex        =   23
            Top             =   2280
            Width           =   1095
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
            Left            =   480
            TabIndex        =   22
            Top             =   2640
            Width           =   1335
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
            Left            =   480
            TabIndex        =   18
            Top             =   840
            Width           =   1095
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "Last"
            Height          =   255
            Left            =   1920
            TabIndex        =   17
            Top             =   1800
            Width           =   615
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Caption         =   "Middle"
            Height          =   255
            Left            =   1920
            TabIndex        =   16
            Top             =   1320
            Width           =   615
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "First"
            Height          =   255
            Left            =   1920
            TabIndex        =   15
            Top             =   840
            Width           =   615
         End
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add Contact"
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
         Left            =   8160
         Picture         =   "frmNewContact.frx":1132
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   4320
         Width           =   1455
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
         Left            =   9600
         Picture         =   "frmNewContact.frx":21B4
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   4320
         Width           =   1095
      End
      Begin VB.TextBox txtLocation 
         Height          =   375
         Left            =   8160
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   3600
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "Browse"
         Height          =   375
         Left            =   9000
         TabIndex        =   3
         Top             =   3120
         Width           =   855
      End
      Begin VB.PictureBox pic 
         BackColor       =   &H00FFFFFF&
         Height          =   2295
         Left            =   8160
         ScaleHeight     =   2235
         ScaleWidth      =   2475
         TabIndex        =   2
         Top             =   720
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
         Left            =   9840
         TabIndex        =   1
         Top             =   3120
         Width           =   855
      End
      Begin MSComDlg.CommonDialog dlg 
         Left            =   8280
         Top             =   3000
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label lblUserID 
         Caption         =   "---"
         Height          =   255
         Left            =   3240
         TabIndex        =   47
         Top             =   240
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.Label lblRecordID 
         Caption         =   "---"
         Height          =   255
         Left            =   1320
         TabIndex        =   10
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
         Left            =   240
         TabIndex        =   9
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label9 
         Caption         =   "Image"
         Height          =   255
         Left            =   8160
         TabIndex        =   8
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label11 
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
         Left            =   360
         TabIndex        =   7
         Top             =   8400
         Width           =   1575
      End
   End
   Begin VB.Label Label19 
      Alignment       =   1  'Right Justify
      Caption         =   "Mobile Number"
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
      Left            =   0
      TabIndex        =   32
      Top             =   0
      Width           =   1575
   End
End
Attribute VB_Name = "frmNewContact"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAdd_Click()
Call ValidateAge(dtpBirthDate)
If WarnBlank(txtFirstName, "First Name") Then SSTab1.Tab = 0: Exit Sub
If WarnBlank(txtMiddleName, "Middle Name") = True Then SSTab1.Tab = 0: Exit Sub
If WarnBlank(txtLastName, "Last Name") = True Then SSTab1.Tab = 0: Exit Sub
If WarnBlank(txtAddress, "Address") = True Then SSTab1.Tab = 1: Exit Sub
If ErrCounter = 0 Then
    MsgBox "Please specify at least one contact number.", vbOKOnly + vbExclamation, "No Contact Number Specified"
    SSTab1.Tab = 1
    Exit Sub
Else
    Call Upload(txtLocation)
    If Age < 6 Then
        MsgBox "Please select a valid age.", vbOKOnly + vbExclamation, "Invalid Age"
        SSTab1.Tab = 0
        Exit Sub
    Else
        If cmdAdd.Caption = "Add Contact" Then
            con.Execute ("INSERT INTO Contacts VALUES('" & lblRecordID.Caption & "', " & Replace(UserID, "'", "''") & ",'" & Replace(cboContactGroup.Text, "'", "''") & "','" & Replace(txtLastName.Text, "'", "''") & "', '" & Replace(txtFirstName.Text, "'", "''") & "', '" & Replace(txtMiddleName.Text, "'", "''") & "', '" & Replace(txtPositionDesignation.Text, "'", "''") & "','" & Replace(txtCompanySchool.Text, "'", "''") & "', '" & Replace(txtPhoneNumber.Text, "'", "''") & "','" & Replace(txtMobileNumber.Text, "'", "''") & "','" & Replace(txtFaxNumber.Text, "'", "''") & "','" & Replace(txtEmailAddress.Text, "'", "''") & "','" & Replace(txtAddress.Text, "'", "''") & "',#" & dtpBirthDate.Value & "#," & Age & ",'" & IIf((optMale.Value = True), optMale.Caption, optFemale.Caption) & "', '" & cboCivilStatus.Text & "','" & Replace(txtReligion.Text, "'", "''") & "','" & Replace(txtNotes.Text, "'", "''") & "','" & Replace(FullPath, "'", "''") & "')")
            MsgBox "New contact successfully added.", vbOKOnly + vbInformation, "Success"
        ElseIf cmdAdd.Caption = "Update" Then
            con.Execute ("UPDATE Contacts SET GroupName='" & Replace(cboContactGroup.Text, "'", "''") & "',LastName='" & Replace(txtLastName.Text, "'", "''") & "',FirstName='" & Replace(txtFirstName.Text, "'", "''") & "',MiddleName='" & Replace(txtMiddleName.Text, "'", "''") & "',PositionDesignation='" & Replace(txtPositionDesignation.Text, "'", "''") & "',CompanySchool='" & Replace(txtCompanySchool.Text, "'", "''") & "',PhoneNumber='" & Replace(txtPhoneNumber.Text, "'", "''") & "',MobileNumber='" & Replace(txtMobileNumber.Text, "'", "''") & "',FaxNumber='" & Replace(txtFaxNumber.Text, "'", "''") & "',EmailAddress='" & Replace(txtEmailAddress.Text, "'", "''") & "',Address='" & Replace(txtAddress.Text, "'", "''") & "',BirthDate=#" & dtpBirthDate.Value & "#, " & _
                "Age=" & Age & ",Gender='" & IIf((optMale.Value = True), optMale.Caption, optFemale.Caption) & "',CivilStatus='" & cboCivilStatus.Text & "',Religion='" & Replace(txtReligion.Text, "'", "''") & "',Notes='" & Replace(txtNotes.Text, "'", "''") & "',ImagePath='" & Replace(FullPath, "'", "''") & "' WHERE ContactID=" & lblRecordID.Caption & " AND LUserID=" & lblUserID.Caption & "")
            MsgBox "Contact information successfully updated.", vbOKOnly + vbInformation, "Success"
        End If
    End If
End If
Unload Me
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

Private Sub Form_Load()
ErrCounter = 0
SSTab1.Tab = 0
cboCivilStatus.ListIndex = 0
End Sub

Private Sub txtAddress_Change()
Call ErrChecker(txtAddress)
End Sub

Private Sub txtEmailAddress_Change()
Call ErrChecker(txtEmailAddress)
End Sub

Private Sub txtFaxNumber_Change()
Call ErrChecker(txtFaxNumber)
End Sub

Private Sub txtMobileNumber_Change()
Call ErrChecker(txtMobileNumber)
End Sub

Private Sub txtPhoneNumber_Change()
Call ErrChecker(txtPhoneNumber)
End Sub
Public Sub ErrChecker(txt As TextBox)
If Len(txt.Text) <> 0 Then ErrCounter = ErrCounter + 1
End Sub
