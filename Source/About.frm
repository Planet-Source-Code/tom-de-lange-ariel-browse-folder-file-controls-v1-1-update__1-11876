VERSION 5.00
Begin VB.Form dlgAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About Ariel Folder Control"
   ClientHeight    =   2850
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   6090
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "About.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1967.12
   ScaleMode       =   0  'User
   ScaleWidth      =   5718.824
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fr 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   60
      TabIndex        =   1
      Top             =   0
      Width           =   5955
      Begin VB.CommandButton cmdOK 
         Cancel          =   -1  'True
         Caption         =   "OK"
         Default         =   -1  'True
         Height          =   375
         Left            =   4500
         TabIndex        =   0
         ToolTipText     =   "Close window"
         Top             =   2220
         Width           =   1275
      End
      Begin VB.PictureBox picIcon 
         AutoSize        =   -1  'True
         ClipControls    =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1050
         Left            =   180
         Picture         =   "About.frx":030A
         ScaleHeight     =   695.31
         ScaleMode       =   0  'User
         ScaleWidth      =   705.845
         TabIndex        =   2
         Top             =   240
         Width           =   1065
      End
      Begin VB.Label lblMail 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "E-mail: tomdl@attglobal.net"
         Height          =   195
         Left            =   3780
         TabIndex        =   8
         Tag             =   "Company"
         Top             =   1920
         Width           =   1995
      End
      Begin VB.Label lblCopyright 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Copyright"
         Height          =   195
         Left            =   5070
         TabIndex        =   7
         Tag             =   "Copyright"
         Top             =   1440
         Width           =   705
      End
      Begin VB.Label lblCompany 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Company"
         Height          =   195
         Left            =   5100
         TabIndex        =   6
         Tag             =   "Company"
         Top             =   1665
         Width           =   675
      End
      Begin VB.Label lblProductName 
         AutoSize        =   -1  'True
         Caption         =   "Product Name"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1320
         TabIndex        =   5
         Tag             =   "Product"
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label lblDescription 
         AutoSize        =   -1  'True
         Caption         =   "Description (App.comments)"
         Height          =   195
         Left            =   1320
         TabIndex        =   4
         Top             =   540
         Width           =   4500
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblVersion 
         AutoSize        =   -1  'True
         Caption         =   "Version"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1320
         TabIndex        =   3
         Tag             =   "Version"
         Top             =   1020
         Width           =   735
      End
   End
End
Attribute VB_Name = "dlgAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--------------------------------------------------------------------
'Module     : ArAbout
'Description: About Ariel
'Release    : 2000 VB6
'Copyright  : © T De Lange
'--------------------------------------------------------------------
Option Explicit
Option Base 0
DefLng H-N
DefBool O

Private Sub cmdOK_Click()
'---------------------------------
'Unload form
'---------------------------------
Unload Me

End Sub

Private Sub Form_Load()
'-----------------------------------------
'Load defaults
'-----------------------------------------
Me.Caption = "About " & App.Title
lblVersion.Caption = "Version " & App.Major & "." & Format(App.Minor, "00") & "." & Format(App.Revision, "0000")
lblProductName.Caption = App.Title
lblDescription.Caption = App.Comments
lblCopyright.Caption = "© " & App.LegalCopyright
lblCompany.Caption = App.CompanyName
  
End Sub

