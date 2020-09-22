VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmSchoolInfo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "School Info"
   ClientHeight    =   4710
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8535
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSchoolInfo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4710
   ScaleWidth      =   8535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraSchoolInfo 
      Height          =   4575
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   8295
      Begin VB.PictureBox pic 
         BackColor       =   &H00FFFFFF&
         Height          =   2000
         Left            =   6120
         ScaleHeight     =   1935
         ScaleWidth      =   1935
         TabIndex        =   17
         Top             =   600
         Width           =   2000
         Begin VB.Image img 
            Height          =   1935
            Left            =   0
            Stretch         =   -1  'True
            Top             =   0
            Width           =   1935
         End
      End
      Begin VB.TextBox txtLocation 
         Height          =   375
         Left            =   6000
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   3000
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "Browse"
         Height          =   375
         Left            =   6360
         TabIndex        =   14
         Top             =   2640
         Width           =   855
      End
      Begin VB.CommandButton cmdRemove 
         Caption         =   "Remove"
         Height          =   375
         Left            =   7200
         TabIndex        =   13
         Top             =   2640
         Width           =   855
      End
      Begin VB.TextBox txtOwnerHead 
         Height          =   375
         Left            =   2280
         MaxLength       =   255
         TabIndex        =   12
         Top             =   3000
         Width           =   3495
      End
      Begin VB.TextBox txtEmailAddress 
         Height          =   375
         Left            =   2280
         MaxLength       =   255
         TabIndex        =   10
         Top             =   2520
         Width           =   3495
      End
      Begin VB.TextBox txtContactNumber 
         Height          =   495
         Left            =   2280
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   8
         Top             =   1920
         Width           =   3495
      End
      Begin VB.TextBox txtAddress 
         Height          =   855
         Left            =   2280
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Top             =   960
         Width           =   3495
      End
      Begin VB.TextBox txtSchoolName 
         Height          =   495
         Left            =   2280
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   360
         Width           =   3495
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
         Height          =   855
         Left            =   6960
         Picture         =   "frmSchoolInfo.frx":1082
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   3480
         Width           =   1095
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save"
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
         Height          =   855
         Left            =   5760
         Picture         =   "frmSchoolInfo.frx":2104
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   3480
         Width           =   1215
      End
      Begin MSComDlg.CommonDialog dlg 
         Left            =   7200
         Top             =   120
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label lblLogo 
         Caption         =   "Logo"
         Height          =   255
         Left            =   6120
         TabIndex        =   16
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label lblOwnerHead 
         Caption         =   "Owner/Head"
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
         Top             =   3000
         Width           =   1935
      End
      Begin VB.Label lblEmailAddress 
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
         TabIndex        =   9
         Top             =   2520
         Width           =   1935
      End
      Begin VB.Label lblContactNumber 
         Caption         =   "Contact Number(s)"
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
         TabIndex        =   7
         Top             =   1920
         Width           =   1935
      End
      Begin VB.Label lblAddress 
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
         TabIndex        =   5
         Top             =   960
         Width           =   855
      End
      Begin VB.Label lblName 
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
         TabIndex        =   3
         Top             =   360
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmSchoolInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
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

Private Sub cmdSave_Click()
If WarnBlank(txtSchoolName, "School Name") = True Then Exit Sub
If WarnBlank(txtAddress, "School Address") Then Exit Sub
If WarnBlank(txtContactNumber, "Contact Number") = True Then Exit Sub
Call Upload(txtLocation)
If cmdSave.Caption = "Save" Then
    con.Execute ("INSERT INTO SchoolInfo VALUES(#" & Format(Now, "mm/dd/yyyy hh:mm:ss am/pm") & "#, '" & Replace(txtSchoolName.Text, "'", "''") & "', '" & Replace(txtAddress.Text, "'", "''") & "', '" & Replace(txtContactNumber.Text, "'", "''") & "', '" & Replace(txtEmailAddress.Text, "'", "''") & "', '" & Replace(txtOwnerHead.Text, "'", "''") & "','" & FullPath & "')")
    MsgBox "School Information successfully added.", vbOKOnly + vbInformation, "Success"
    Unload Me
ElseIf cmdSave.Caption = "Update" Then
    con.Execute ("UPDATE SchoolInfo SET DateModified=#" & Format(Now, "mm/dd/yyyy hh:mm:ss am/pm") & "#,SchoolName='" & Replace(txtSchoolName.Text, "'", "''") & "',SchoolAddress='" & Replace(txtAddress.Text, "'", "''") & "',ContactNumber='" & Replace(txtContactNumber.Text, "'", "''") & "',EmailAddress='" & Replace(txtEmailAddress.Text, "'", "''") & "',OwnerHead='" & Replace(txtOwnerHead.Text, "'", "''") & "',ImagePath='" & FullPath & "'")
    MsgBox "School Information successfully updated.", vbOKOnly + vbInformation, "Success"
    Unload Me
End If

End Sub

Private Sub Form_Activate()
If con.State = 0 Then Call konek
End Sub

Private Sub Text1_Change()

End Sub

Private Sub Form_Load()
Set rs = con.Execute("SELECT * FROM SchoolInfo")
If rs.RecordCount = 0 Then
    cmdSave.Caption = "Save"
Else
    cmdSave.Caption = "Update"
    With rs
        txtSchoolName.Text = !SchoolName
        txtAddress.Text = !SchoolAddress
        txtContactNumber.Text = !ContactNumber
        txtEmailAddress.Text = !EmailAddress
        txtOwnerHead.Text = !OwnerHead
        If !ImagePath <> Null Or !ImagePath <> "" Then
            strTest = App.Path & "\Images\" & !ImagePath
            If fs.FileExists(strTest) = True Then
                Call loadLogo(strTest, img, pic)
                FullPath = !ImagePath
            Else
                MsgBox "Image file not found.", vbOKOnly + vbExclamation, "Not Found"
            End If
        End If
    End With
End If
End Sub

