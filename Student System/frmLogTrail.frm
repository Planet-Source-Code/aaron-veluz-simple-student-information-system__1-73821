VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmLogTrail 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Log Trail"
   ClientHeight    =   6270
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9270
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLogTrail.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6270
   ScaleWidth      =   9270
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmLogTrail 
      Height          =   5535
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   9015
      Begin VB.Frame fraSearch 
         Caption         =   "Search Log Trail"
         Height          =   975
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   8775
         Begin VB.ComboBox cboCategory 
            Height          =   315
            ItemData        =   "frmLogTrail.frx":1082
            Left            =   1080
            List            =   "frmLogTrail.frx":1095
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   480
            Width           =   2055
         End
         Begin VB.TextBox txtSearch 
            Height          =   375
            Left            =   3360
            MaxLength       =   255
            TabIndex        =   7
            Top             =   480
            Visible         =   0   'False
            Width           =   2535
         End
         Begin VB.ComboBox cboPrivilege 
            Height          =   315
            ItemData        =   "frmLogTrail.frx":10C9
            Left            =   3360
            List            =   "frmLogTrail.frx":10D6
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   480
            Visible         =   0   'False
            Width           =   1695
         End
         Begin MSComCtl2.DTPicker dtpDateFrom 
            Height          =   375
            Left            =   4680
            TabIndex        =   5
            Top             =   480
            Visible         =   0   'False
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   661
            _Version        =   393216
            Format          =   53542913
            CurrentDate     =   40567
         End
         Begin MSComCtl2.DTPicker dtpDateTo 
            Height          =   375
            Left            =   6960
            TabIndex        =   9
            Top             =   480
            Visible         =   0   'False
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   661
            _Version        =   393216
            Format          =   53542913
            CurrentDate     =   40567
         End
         Begin VB.Label lblDate 
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
            Left            =   3360
            TabIndex        =   13
            Top             =   480
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.Label lblDateFrom 
            Caption         =   "From"
            Height          =   255
            Left            =   4080
            TabIndex        =   12
            Top             =   480
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.Label lblDateTo 
            Caption         =   "To"
            Height          =   255
            Left            =   6600
            TabIndex        =   11
            Top             =   480
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.Label lblCategory 
            Caption         =   "Category:"
            Height          =   255
            Left            =   120
            TabIndex        =   10
            Top             =   480
            Width           =   1215
         End
      End
      Begin MSComctlLib.ListView lvLogTrail 
         Height          =   3735
         Left            =   120
         TabIndex        =   1
         ToolTipText     =   "Double-click an item to view/edit details"
         Top             =   1320
         Width           =   8775
         _ExtentX        =   15478
         _ExtentY        =   6588
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
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Log ID"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Date"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Time In"
            Object.Width           =   2364
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Time Out"
            Object.Width           =   2364
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Username"
            Object.Width           =   3069
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Privilege"
            Object.Width           =   3298
         EndProperty
      End
      Begin VB.Label lblRecord 
         Alignment       =   1  'Right Justify
         Caption         =   "---"
         Height          =   255
         Left            =   1320
         TabIndex        =   2
         Top             =   5160
         Width           =   7455
      End
   End
   Begin MSComctlLib.Toolbar tbrLogTrail 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   9270
      _ExtentX        =   16351
      _ExtentY        =   1005
      ButtonWidth     =   2672
      ButtonHeight    =   1005
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Print Report"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Show All"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
      EndProperty
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   4200
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   2
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLogTrail.frx":1104
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLogTrail.frx":2196
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "frmLogTrail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboCategory_Click()
If cboCategory.Text = "Date" Then
    Call ShowDateTime(True)
    txtSearch.Visible = False
    cboPrivilege.Visible = False
ElseIf cboCategory.Text = "Privilege" Then
    Call ShowDateTime(False)
    txtSearch.Visible = False
    cboPrivilege.Visible = True
ElseIf cboCategory.Text = "Username" Or cboCategory.Text = "Log ID" Then
    Call ShowDateTime(False)
    txtSearch.Visible = True
    cboPrivilege.Visible = False
Else
    Call ShowDateTime(False)
    txtSearch.Visible = False
    cboPrivilege.Visible = False
    Call LoadLogTrail(lvLogTrail)
End If
End Sub

Private Sub cboPrivilege_Click()
Set rs = con.Execute("SELECT * FROM LogTrail WHERE Privilege='" & cboPrivilege.Text & "' ORDER BY LogID")
lvLogTrail.ListItems.Clear

For xCount = 1 To rs.RecordCount
    With ls
        Set ls = lvLogTrail.ListItems.Add(, , rs!LogID)
        ls.SubItems(1) = rs!LogDate
        ls.SubItems(2) = rs!TimeIn
        ls.SubItems(3) = IIf(IsNull(rs!TimeOut), "", rs!TimeOut)
        ls.SubItems(4) = rs!Username
        ls.SubItems(5) = rs!Privilege
        rs.MoveNext
    End With
Next xCount
Call CheckCount(lvLogTrail, lblRecord)
End Sub



Private Sub dtpDateFrom_Change()
Call CheckDates
End Sub

Private Sub dtpDateFrom_Click()
Call CheckDates
End Sub

Private Sub dtpDateFrom_DblClick()
Call CheckDates
End Sub

Private Sub dtpDateFrom_DropDown()
Call CheckDates
End Sub

Private Sub dtpDateTo_Change()
Call CheckDates
End Sub

Private Sub dtpDateTo_Click()
Call CheckDates
End Sub

Private Sub dtpDateTo_DblClick()
Call CheckDates
End Sub

Private Sub dtpDateTo_DropDown()
Call CheckDates
End Sub

Private Sub Form_Activate()
If con.State = 0 Then Call konek
End Sub

Private Sub Form_Load()
cboCategory.ListIndex = 4
Call cboCategory_Click
Call LoadLogTrail(lvLogTrail)
Call CheckCount(lvLogTrail, lblRecord)
End Sub


Private Sub Label2_Click()

End Sub

Private Sub lblTimeTo_Click()

End Sub

Private Sub lvLogTrail_Click()
Call CheckCount(lvLogTrail, lblRecord)
End Sub

Private Sub lvLogTrail_KeyDown(KeyCode As Integer, Shift As Integer)
Call CheckCount(lvLogTrail, lblRecord)
End Sub

Private Sub lvLogTrail_KeyUp(KeyCode As Integer, Shift As Integer)
Call CheckCount(lvLogTrail, lblRecord)
End Sub

Private Sub tbrLogTrail_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
    Case 1 'Print
        Set rptLogTrail.DataSource = rs
        rptLogTrail.Show vbModal, Me
    Case 3 'Show All
        cboCategory.ListIndex = 4
        Call cboCategory_Click
End Select
End Sub
Public Sub ShowDateTime(enabled As Boolean)
lblDate.Visible = enabled
lblDateTo.Visible = enabled
lblDateFrom.Visible = enabled
dtpDateFrom.Visible = enabled
dtpDateTo.Visible = enabled

End Sub
Public Sub CheckDates()
Set rs = con.Execute("SELECT * FROM LogTrail WHERE LogDate BETWEEN #" & dtpDateFrom.Value & "# AND #" & dtpDateTo.Value & "# ORDER BY LogID")
lvLogTrail.ListItems.Clear

For xCount = 1 To rs.RecordCount
    With ls
        Set ls = lvLogTrail.ListItems.Add(, , rs!LogID)
        ls.SubItems(1) = rs!LogDate
        ls.SubItems(2) = rs!TimeIn
        ls.SubItems(3) = IIf(IsNull(rs!TimeOut), "", rs!TimeOut)
        ls.SubItems(4) = rs!Username
        ls.SubItems(5) = rs!Privilege
        rs.MoveNext
    End With
Next xCount
Call CheckCount(lvLogTrail, lblRecord)
End Sub

Private Sub txtSearch_Change()

Select Case cboCategory.Text
    Case "Log ID"
        strTest = "LogID"
    Case "Username"
        strTest = "Username"
End Select
If txtSearch.Text <> "" Then
    If cboCategory.Text <> "Log ID" Then
        Set rs = con.Execute("SELECT * FROM LogTrail WHERE Username LIKE '" & txtSearch.Text & "%' ORDER BY LogID")
    Else
        Set rs = con.Execute("SELECT * FROM LogTrail WHERE LogID=" & Val(txtSearch.Text) & " ORDER BY LogID")
    End If
    lvLogTrail.ListItems.Clear
    
    For xCount = 1 To rs.RecordCount
        With ls
            Set ls = lvLogTrail.ListItems.Add(, , rs!LogID)
            ls.SubItems(1) = rs!LogDate
            ls.SubItems(2) = rs!TimeIn
            ls.SubItems(3) = IIf(IsNull(rs!TimeOut), "", rs!TimeOut)
            ls.SubItems(4) = rs!Username
            ls.SubItems(5) = rs!Privilege
            rs.MoveNext
        End With
    Next xCount
    Call CheckCount(lvLogTrail, lblRecord)
Else
    Call LoadLogTrail(lvLogTrail)
End If
End Sub
