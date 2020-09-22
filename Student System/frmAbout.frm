VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About"
   ClientHeight    =   2790
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4590
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2790
   ScaleWidth      =   4590
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraDeveloperInfo 
      Height          =   2655
      Left            =   120
      TabIndex        =   13
      Top             =   0
      Width           =   4335
      Begin VB.Image Image1 
         Height          =   2055
         Left            =   120
         Picture         =   "frmAbout.frx":1082
         Stretch         =   -1  'True
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lblApp 
         Alignment       =   2  'Center
         Caption         =   "Student Information and Violation Record Keeping System by Aaron Villamejor Veluz"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   1200
         TabIndex        =   14
         Top             =   600
         Width           =   3015
      End
   End
   Begin VB.Frame fraSchoolInfo 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4335
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   8415
      Begin VB.PictureBox pic 
         BackColor       =   &H00FFFFFF&
         Height          =   2295
         Left            =   5640
         ScaleHeight     =   2235
         ScaleWidth      =   2475
         TabIndex        =   6
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
      Begin VB.Label lblOwnerHead 
         AutoSize        =   -1  'True
         Caption         =   "---"
         Height          =   195
         Left            =   360
         TabIndex        =   12
         Top             =   3720
         Width           =   225
      End
      Begin VB.Label lblEmailAddress 
         AutoSize        =   -1  'True
         Caption         =   "---"
         Height          =   195
         Left            =   360
         TabIndex        =   11
         Top             =   3000
         Width           =   225
      End
      Begin VB.Label lblContactNumber 
         AutoSize        =   -1  'True
         Caption         =   "---"
         Height          =   195
         Left            =   360
         TabIndex        =   10
         Top             =   2280
         Width           =   225
      End
      Begin VB.Label lblAddress 
         AutoSize        =   -1  'True
         Caption         =   "---"
         Height          =   195
         Left            =   360
         TabIndex        =   9
         Top             =   1440
         Width           =   225
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         Caption         =   "---"
         Height          =   195
         Left            =   360
         TabIndex        =   8
         Top             =   600
         Width           =   225
      End
      Begin VB.Label Label6 
         Caption         =   "Logo"
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
         Left            =   5640
         TabIndex        =   7
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label5 
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
         TabIndex        =   5
         Top             =   3480
         Width           =   1815
      End
      Begin VB.Label Label4 
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
         TabIndex        =   4
         Top             =   2760
         Width           =   1815
      End
      Begin VB.Label Label3 
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
         TabIndex        =   3
         Top             =   2040
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "Address/Location"
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
         TabIndex        =   2
         Top             =   1200
         Width           =   1815
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
         TabIndex        =   1
         Top             =   360
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

