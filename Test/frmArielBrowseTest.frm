VERSION 5.00
Object = "{C1C2430B-978A-11D4-9744-004F490561B3}#7.0#0"; "Ariel Browse Ctrl.ocx"
Begin VB.Form frmArielBrowseTest 
   Caption         =   "Ariel Browse Demo"
   ClientHeight    =   3315
   ClientLeft      =   2190
   ClientTop       =   2475
   ClientWidth     =   7245
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmArielBrowseTest.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3315
   ScaleWidth      =   7245
   Begin VB.Frame frFile 
      Caption         =   "File"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   2880
      TabIndex        =   23
      Top             =   2100
      Width           =   1935
      Begin VB.CheckBox chkInclPath 
         Caption         =   "Include Path"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   7
         ToolTipText     =   "Converts DOS uppercase files to 'proper' format"
         Top             =   300
         Width           =   1575
      End
   End
   Begin VB.ComboBox cmbDialog 
      Height          =   315
      ItemData        =   "frmArielBrowseTest.frx":0E42
      Left            =   1440
      List            =   "frmArielBrowseTest.frx":0E4C
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   2580
      Width           =   1275
   End
   Begin VB.TextBox txtFilterIndex 
      Height          =   315
      Left            =   1440
      TabIndex        =   5
      Text            =   "1"
      Top             =   2160
      Width           =   1275
   End
   Begin VB.TextBox txtFilter 
      Height          =   315
      Left            =   1440
      TabIndex        =   4
      Text            =   "Text files (*.txt)|*.txt|All Files (*.*)|*.*|"
      Top             =   1740
      Width           =   3375
   End
   Begin VB.Frame frGnl 
      Caption         =   "General"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   5100
      TabIndex        =   19
      Top             =   1560
      Width           =   1935
      Begin VB.CheckBox chkItalic 
         Caption         =   "Italic"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   180
         TabIndex        =   12
         Top             =   600
         Width           =   915
      End
      Begin VB.CheckBox chkProper 
         Caption         =   "Proper"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   180
         TabIndex        =   11
         ToolTipText     =   "Converts DOS uppercase files to 'proper' format"
         Top             =   300
         Width           =   915
      End
      Begin VB.CheckBox chkBold 
         Caption         =   "Bold"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   180
         TabIndex        =   13
         Top             =   900
         Width           =   915
      End
   End
   Begin VB.Frame frFolder 
      Caption         =   "Folder"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1395
      Left            =   5100
      TabIndex        =   18
      Top             =   0
      Width           =   1935
      Begin VB.CheckBox chkDomain 
         Caption         =   "Domain"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   180
         TabIndex        =   8
         Top             =   360
         Width           =   1575
      End
      Begin VB.CheckBox chkAncestors 
         Caption         =   "Return Ancestors"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   180
         TabIndex        =   9
         Top             =   660
         Width           =   1575
      End
      Begin VB.CheckBox chkRetFSDirs 
         Caption         =   "Return FS Dirs"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   180
         TabIndex        =   10
         Top             =   960
         Width           =   1575
      End
   End
   Begin VB.TextBox txtRoot 
      Height          =   315
      Left            =   1440
      TabIndex        =   1
      Top             =   480
      Width           =   3375
   End
   Begin VB.ComboBox cmbRoot 
      Height          =   315
      ItemData        =   "frmArielBrowseTest.frx":0E66
      Left            =   1440
      List            =   "frmArielBrowseTest.frx":0E82
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   60
      Width           =   3375
   End
   Begin ArielBrowseCtrl.ArielBrowseFolder ArielBrowseFolder 
      Height          =   315
      Left            =   1440
      TabIndex        =   2
      ToolTipText     =   "Click to select a folder"
      Top             =   900
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   556
      ForeColor       =   -2147483630
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Object.ToolTipText     =   "Click to select a folder"
      DropdownPicture =   "frmArielBrowseTest.frx":0EDB
   End
   Begin ArielBrowseCtrl.ArielBrowseFile ArielBrowseFile 
      Height          =   315
      Left            =   1440
      TabIndex        =   3
      ToolTipText     =   "Click to select a file to open/save"
      Top             =   1320
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   556
      ForeColor       =   -2147483630
      Object.ToolTipText     =   "Click to select a file to open/save"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropdownPicture =   "frmArielBrowseTest.frx":0FD5
      Proper          =   -1  'True
      Filter          =   "All Files (*.*)|*.*|Access Files (*.mdb)|*.mdb"
      FilterIndex     =   2
      Path            =   "D:\Ariel"
      FileName        =   "Ariel 2000.mdb"
      File            =   "D:\Ariel\Ariel 2000.mdb"
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "File Dialog Mode"
      Height          =   195
      Index           =   9
      Left            =   180
      TabIndex        =   22
      Top             =   2640
      Width           =   1155
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "Filter Index"
      Height          =   195
      Index           =   8
      Left            =   180
      TabIndex        =   21
      Top             =   2220
      Width           =   825
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "File Filter"
      Height          =   195
      Index           =   7
      Left            =   180
      TabIndex        =   20
      Top             =   1800
      Width           =   645
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "Custom Root"
      Height          =   195
      Index           =   4
      Left            =   180
      TabIndex        =   17
      Top             =   540
      Width           =   930
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "Root Folder"
      Height          =   195
      Index           =   3
      Left            =   180
      TabIndex        =   16
      Top             =   120
      Width           =   840
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "File"
      Height          =   195
      Index           =   2
      Left            =   180
      TabIndex        =   15
      Top             =   1380
      Width           =   240
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "Folder"
      Height          =   195
      Index           =   1
      Left            =   180
      TabIndex        =   14
      Top             =   960
      Width           =   450
   End
End
Attribute VB_Name = "frmArielBrowseTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBrowse_Click()
'cdlOFNReadOnly &H1
'Causes the Read Only check box to be initially checked when the
'dialog box is created. This flag also indicates the state of the
'Read Only check box when the dialog box is closed.

'cdlOFNOverwritePrompt &H2
'Causes the Save As dialog box to generate a message box if
'the selected file already exists. The user must confirm whether
'to overwrite the file.

'cdlOFNHideReadOnly &H4
'Hides the Read Onlycheck box.

'cdlOFNNoChangeDir &H8
'Forces the dialog box to set the current directory to
'what it was when the dialog box was opened.

'cdlOFNHelpButton &H10
'Causes the dialog box to display the Help button.

'cdlOFNNoValidate &H100
'Specifies that the common dialog box allows invalid characters
'in the returned filename.

'cdlOFNAllowMultiselect &H200
'Specifies that the File Namelist box
'allows multiple selections. The user can select more than one
'file at run time by pressing the SHIFT key and using the UP ARROW
'and DOWN ARROW keys to select the desired files.
'When this is done, the FileName property returns a string
'containing the names of all selected files.
'The names in the string are delimited by spaces.
 
'CdlOFNExtensionDifferent &H400
'Indicates that the extension of the returned filename is
'different from the extension specified by the DefaultExt property.
'This flag isn't set if the DefaultExt property is Null,
'if the extensions match, or if the file has no extension.
'This flag value can be checked upon closing the dialog box.

'cdlOFNPathMustExist &H800
'Specifies that the user can enter only valid paths.
'If this flag is set and the user enters an invalid path,
'a warning message is displayed.

'cdlOFNFileMustExist &H1000
'Specifies that the user can enter only names of existing files
'in the File Name text box. If this flag is set and the user
'enters an invalid filename, a warning is displayed.
'This flag automatically sets the cdlOFNPathMustExist flag.

'cdlOFNCreatePrompt &H2000
'Specifies that the dialog box prompts the user to create
'a file that doesn't currently exist. This flag automatically
'sets the cdlOFNPathMustExist and cdlOFNFileMustExist flags.

'cdlOFNShareAware &H4000
'Specifies that sharing violation errors will be ignored.


'CdlOFNNoReadOnlyReturn &H8000
'Specifies that the returned file won't have the Read Only
'attribute set and won't be in a write-protected directory.

'cdlOFNNoLongNames &H40000
'No long file names.

'cdlOFNExplorer &H80000
'Use the Explorer-like Open A File dialog box template.
'Works with Windows 95 and Windows NT 4.0.

'CdlOFNNoDereferenceLinks &H100000
'Do not dereference shell links (also known as shortcuts).
'By default, choosing a shell link causes it to be dereferenced
'by the shell.

'cdlOFNLongNames &H200000
'Use long filenames.


dlg.ShowOpen

End Sub


Private Sub ArielBrowseFile_Click(File As String, Path As String, FileName As String)
'--------------------------------------------------------
'Update references
'--------------------------------------------------------
ArielBrowseFolder.Text = Path
txtFilterIndex = ArielBrowseFile.FilterIndex

End Sub

Private Sub ArielBrowseFile_Dropdown()
'------------------------------------------------
'Use this event to set properties prior to the
'dialog box opening
'------------------------------------------------
ArielBrowseFile.FileDialogType = cmbDialog.ListIndex
If ArielBrowseFile.FileDialogType = OpenFile Then
  ArielBrowseFile.Caption = "Select file to open"
Else
  ArielBrowseFile.Caption = "Select file to save"
End If
ArielBrowseFile.Filter = txtFilter
ArielBrowseFile.FilterIndex = txtFilterIndex
ArielBrowseFile.Path = ArielBrowseFolder.Text
ArielBrowseFile.IncludePath = chkInclPath And 1


End Sub

Private Sub chkInclPath_Click()

ArielBrowseFile.IncludePath = chkInclPath And 1

End Sub

Private Sub Form_Load()
cmbRoot.ListIndex = 1
chkItalic = 0
chkBold = 0
chkProper = 1
chkRetFSDirs = 1
cmbDialog.ListIndex = 0

End Sub

Private Sub ArielBrowseFolder_Dropdown()
'------------------------------------------
'Use this event to set properties
'prior to opening of the folder dialog box
'------------------------------------------
With ArielBrowseFolder
  If .RootFolder = Custom Then
    .CustomRootFolder = txtRoot
  End If
End With

End Sub

Private Sub chkAncestors_Click()
ArielBrowseFolder.ReturnAncestors = chkAncestors And 1

End Sub

Private Sub chkBold_Click()
ArielBrowseFolder.Font.Bold = chkBold And 1
ArielBrowseFile.Font.Bold = chkBold And 1

End Sub

Private Sub chkDomain_Click()
ArielBrowseFolder.Domain = chkDomain And 1

End Sub

Private Sub chkItalic_Click()
ArielBrowseFolder.Font.Italic = chkItalic And 1
ArielBrowseFile.Font.Italic = chkItalic And 1

End Sub


Private Sub chkProper_Click()
ArielBrowseFolder.Proper = chkProper And 1
ArielBrowseFile.Proper = chkProper And 1

End Sub

Private Sub chkRetFSDirs_Click()
ArielBrowseFolder.ReturnFSDirs = chkRetFSDirs And 1

End Sub

Private Sub cmbRoot_Click()
ArielBrowseFolder.RootFolder = cmbRoot.ListIndex
txtRoot.Enabled = (ArielBrowseFolder.RootFolder = asfCustom)

End Sub


