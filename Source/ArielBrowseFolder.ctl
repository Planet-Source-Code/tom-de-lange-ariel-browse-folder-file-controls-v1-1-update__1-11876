VERSION 5.00
Begin VB.UserControl ArielBrowseFolder 
   ClientHeight    =   450
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2115
   KeyPreview      =   -1  'True
   PropertyPages   =   "ArielBrowseFolder.ctx":0000
   ScaleHeight     =   30
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   141
   ToolboxBitmap   =   "ArielBrowseFolder.ctx":0035
   Begin VB.TextBox txtFolder 
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   1695
   End
End
Attribute VB_Name = "ArielBrowseFolder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'--------------------------------------------------------------------
'Module     : ArielBrowseFolder
'Description: Ariel Broswe Folder ActiveX Control
'Version    : V1.00 Sep 2000
'Release    : VB6
'Copyright  : © T De Lange, 2000
'--------------------------------------------------------------------
'V1.00    Sep 00 Original version, based on ColorCombo
'--------------------------------------------------------------------
'Credits:
'All code obtained from www.planet-source-code.com
'Per Andersson, FireStorm@GoToMy.com, www.FireStormEntertainment.cjb.net
'Roman Blachman, romaz@inter.net.il
'Stephen Fonnesbeck, steev@xmission.com, http://www.xmission.com/~steev
'Max Raskin, www.planet-source-code.com
'--------------------------------------------------------------------
Option Base 0
Option Explicit
DefLng A-N, P-Z
DefBool O

'---------------------------------------------------
'Public Enumerations
'---------------------------------------------------
'Limited set of CSIDL-Folders
Public Enum SpecialFolders
  asfCustom = 0
  asfDesktop
  asfMyComputer
  asfMyDocuments
  asfNetwork
  asfProgramGroups
  asfFavorites
  asfRecent
End Enum

'---------------------------------------------------
'Internal constants
'---------------------------------------------------
Const mMinWidth = 35    '21 pixels for edit and 13 for dropdown
Const mDropWidth = 13   'Width of dropdown button

'---------------------------------------------------
'Internal Control Variables
'---------------------------------------------------
Dim rDrp As Rect                'DropDown rectangle (incl border)

'---------------------------------------------------
'Control Properties & Events (using Wizard)
'---------------------------------------------------
'Default Property Values:
Const m_def_MaskColor = &HC0C0C0
Const m_def_Proper = False
Const m_def_Domain = False
Const m_def_ReturnAncestors = False
Const m_def_ReturnFSDirs = True
Const m_def_Caption = "Select a folder"
Const m_def_RootFolder = 1  'Desktop
Const m_def_CustomRootFolder = ""
Const m_def_Text = ""

'Property Variables:
Dim m_MaskColor As OLE_COLOR
Dim m_Picture As StdPicture
Dim m_Proper As Boolean
Dim m_Domain As Boolean
Dim m_ReturnAncestors As Boolean
Dim m_ReturnFSDirs As Boolean
Dim m_Caption As String
Dim m_Text As String
Dim m_RootFolder As SpecialFolders
Dim m_CustomRootFolder As String
Private WithEvents mFont As StdFont
Attribute mFont.VB_VarHelpID = -1

'Event Declarations:
Event Dropdown()
Attribute Dropdown.VB_Description = "Occurs when the user clicks the dropdown button"
Event Click(SelectedPath As String)
Event Change(SelectedPath As String)
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=txtFolder,txtFolder,-1,KeyDown
Attribute KeyDown.VB_Description = "Occurs when the user presses a key while an object has the focus."
Event KeyPress(KeyAscii As Integer) 'MappingInfo=txtFolder,txtFolder,-1,KeyPress
Attribute KeyPress.VB_Description = "Occurs when the user presses and releases an ANSI key."
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=txtFolder,txtFolder,-1,KeyUp
Attribute KeyUp.VB_Description = "Occurs when the user releases a key while an object has the focus."
Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseDown
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseUp
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."

'---------------------------------------------------
'Api Type Declarations
'---------------------------------------------------
Private Type PointApi
  x As Long
  y As Long
End Type

Private Type Rect
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

'==============================================================
'Api Function Declarations
'==============================================================
'For general drawing
Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, lpPoint As PointApi) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As Rect, ByVal edge As Long, ByVal grfFlags As Long) As Long
'Note: the following declaration in the API viewer is incorrect!
'Private Declare Function PtInRect Lib "user32" (lpRect As Rect, pt As PointApi) As Long
Private Declare Function PtInRect Lib "user32" (lpRect As Rect, ByVal x As Long, ByVal y As Long) As Long

'--------------------------------------------------------------------
'General Windows User Interface
'--------------------------------------------------------------------
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As Rect) As Long

'---------------------------------------------------
'Api Constants
'---------------------------------------------------
'For determining which part of the border to draw
Private Const BF_LEFT = &H1
Private Const BF_TOP = &H2
Private Const BF_RIGHT = &H4
Private Const BF_BOTTOM = &H8
Private Const BF_SOFT = &H1000
Private Const BF_FLAT = &H4000
Private Const BF_MONO = &H8000
Private Const BF_BOTTOMLEFT = (BF_BOTTOM Or BF_LEFT)
Private Const BF_BOTTOMRIGHT = (BF_BOTTOM Or BF_RIGHT)
Private Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
Private Const BF_TOPLEFT = (BF_TOP Or BF_LEFT)
Private Const BF_TOPRIGHT = (BF_TOP Or BF_RIGHT)

'For drawing borders with the DrawEdge() function
Private Const BDR_INNER = &HC
Private Const BDR_OUTER = &H3
Private Const BDR_RAISED = &H5
Private Const BDR_RAISEDINNER = &H4
Private Const BDR_RAISEDOUTER = &H1
Private Const BDR_SUNKEN = &HA
Private Const BDR_SUNKENINNER = &H8
Private Const BDR_SUNKENOUTER = &H2
Private Const EDGE_BUMP = BDR_RAISEDOUTER Or BDR_SUNKENINNER
Private Const EDGE_ETCHED = BDR_SUNKENOUTER Or BDR_RAISEDINNER
Private Const EDGE_RAISED = BDR_RAISEDOUTER Or BDR_RAISEDINNER
Private Const EDGE_SUNKEN = BDR_SUNKENOUTER Or BDR_SUNKENINNER

'Flag Constants of the BrowseForFolder API function
Private Const BIF_RETURNONLYFSDIRS = &H1       'Only return file system directories. If the user selects folders that are not part of the file system, the OK button is grayed.
Private Const BIF_DONTGOBELOWDOMAIN = &H2      'Do not include network folders below the domain level in the tree view control.
Private Const BIF_STATUSTEXT = &H4             'Include a status area in the dialog box. The callback function can set the status text by sending messages to the dialog box.
Private Const BIF_RETURNFSANCESTORS As Long = &H8     'Only return file system ancestors. If the user selects anything other than a file system ancestor, the OK button is grayed.
Private Const BIF_EDITBOX As Long = &H10              'Version 4.71. The browse dialog includes an edit control in which the user can type the name of an item.
Private Const BIF_BROWSEFORCOMPUTER As Long = &H1000  'Only return computers. If the user selects anything other than a computer, the OK button is grayed.
Private Const BIF_BROWSEFORPRINTER As Long = &H2000   'Only return printers. If the user selects anything other than a printer, the OK button is grayed.
Private Const BIF_BROWSEINCLUDEFILES As Long = &H4000 'The browse dialog will display files as well as folders.
'private const BIF_VALIDATE             Version 4.71. If the user types an invalid name into the edit box, the browse dialog will call the application's BrowseCallbackProc with the BFFM_VALIDATEFAILED message. This flag is ignored if BIF_EDITBOX is not specified.

'CSIDL Constants for Special Folders
Private Const CSIDL_DESKTOP = &H0   'Desktop Folder
Private Const CSIDL_INTERNET = &H1  'Internet
Private Const CSIDL_PROGRAMS = &H2  'Program Groups
Private Const CSIDL_CONTROLS = &H3  'Control Panel
Private Const CSIDL_PRINTERS = &H4  'Printers
Private Const CSIDL_PERSONAL = &H5  'My Documents
Private Const CSIDL_FAVORITES = &H6 'Favorites
Private Const CSIDL_STARTUP = &H7   'Startup
Private Const CSIDL_RECENT = &H8    'Recent
Private Const CSIDL_SENDTO = &H9    'SendTo
Private Const CSIDL_BUTBUCKET = &HA 'RecylceBin
Private Const CSIDL_STARTMENU = &HB 'StartMenu
Private Const CSIDL_DESKTOPDIRECTORY = &H10   'Windows\Desktop Folder
Private Const CSIDL_DRIVES = &H11   'Devices Virtual Folder (My Computer)
Private Const CSIDL_NETWORK = &H12  'Network Neigborhood Virtual Folder
Private Const CSIDL_NETHOOD = &H13  'Network Neighborhood Folder
Private Const CSIDL_FONTS = &H14    'Fonts Folder
Private Const CSIDL_TEMPLATES = &H15 'ShellNew Folder
Private Const CSIDL_COMMONSTARTMENU = &H16
Private Const CSIDL_COMMONPROGRAMS = &H17
Private Const CSIDL_COMMONSTARTUP = &H18
Private Const CSIDL_COMMONDESKTOPDIR = &H19
Private Const CSIDL_APPLICATIONDATA = &H1A
Private Const CSIDL_PRINTHOOD = &H1B
Private Const CSIDL_ALTSTARTUP = &H1D
Private Const CSIDL_COMMONALTSTARTUP = &H1E
Private Const CSIDL_COMMONFAVORITES = &H1F
Private Const CSIDL_INTERNETCACHE = &H20
Private Const CSIDL_COOKIES = &H21
Private Const CSIDL_HISTORY = &H22


'MappingInfo=txtFolder,txtFolder,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
'-------------------------------------------
'Get Property
'-------------------------------------------
BackColor = txtFolder.BackColor

End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
'-------------------------------------------
'Set Property
'-------------------------------------------
txtFolder.BackColor() = New_BackColor
UserControl.BackColor = New_BackColor
PropertyChanged "BackColor"

End Property

Private Sub ChangePath(sPath As String, Optional ByTyping As Boolean = False)
'-----------------------------------------
'Internal routine to maintain integrity
'of changes to the selected folder
'If ByTyping is true, the text was entered
'by typing in the txtFolder edit box
'-----------------------------------------
If m_Proper Then
  m_Text = StringToProper(sPath)
Else
  m_Text = sPath
End If
If Not ByTyping Then
  txtFolder = m_Text
End If
PropertyChanged "Text"
RaiseEvent Change(m_Text)
'UserControl.Refresh

End Sub

'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
'-------------------------------------------
'Get Property
'-------------------------------------------
Enabled = UserControl.Enabled

End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
'-------------------------------------------
'Set Property
'-------------------------------------------
UserControl.Enabled() = New_Enabled
PropertyChanged "Enabled"

End Property

'MemberInfo=5
Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
'-------------------------------------------
'Refresh the ctrl
'-------------------------------------------
UserControl.Refresh
   
End Sub

Public Function Show() As Boolean
'-------------------------------------------------------
'Shows the Browse Folder dialog box
'Returns: True if as folder was selected
'         False if user clicked cancel
'         or an invalid folder was selected
'-------------------------------------------------------
Dim vStartFolder As Variant   'Startup folder, custom (string) or system (enum)
Dim vRootFolder As Variant    'Topmost folder, custom (string) or system (enum)
Dim sPath As String           'Selected folder
Dim lFlags

'In this implimentation, it makes sense not to use
'a system folder, but rather the current path
vStartFolder = m_Text

'Select the rootfolder (topmost)
Select Case m_RootFolder
Case asfCustom
  If m_CustomRootFolder <> "" Then
    vRootFolder = m_CustomRootFolder
  Else
    vRootFolder = CSIDL_DESKTOP
  End If
Case asfDesktop
  vRootFolder = CSIDL_DESKTOP
Case asfMyComputer
  vRootFolder = CSIDL_DRIVES
Case asfMyDocuments
  vRootFolder = CSIDL_PERSONAL
Case asfNetwork
  vRootFolder = CSIDL_NETWORK
Case asfProgramGroups
  vRootFolder = CSIDL_PROGRAMS
Case asfFavorites
  vRootFolder = CSIDL_FAVORITES
Case asfRecent
  vRootFolder = CSIDL_RECENT
End Select

'Set flags
lFlags = BIF_STATUSTEXT
'Notes on flags
'BIF_BROWSEFORCOMPUTER allows ONLY ComputerNames to be selected, not applicable
'BIF_BROWSEFORPRINTERS allows ONLY Printers to be selected, not applicable
'BIF_DONTGOBELOWDOMAIN limits the search to the same computer
'Normally True to allow network folder selection
'To exclude networks, set to False
'BIF_RETURNONLYFSDIRS limits the search to true folders, excludes Internet, Recycle bin etc
'and is normally TRUE
'BIF_RETURNFSANCESTORS limits the search to "My Documents" etc
'and is normally FALSE

If m_Domain Then
  lFlags = lFlags Or BIF_DONTGOBELOWDOMAIN
End If
If m_ReturnAncestors Then
  lFlags = lFlags Or BIF_RETURNFSANCESTORS
End If
If m_ReturnFSDirs Then
  lFlags = lFlags Or BIF_RETURNONLYFSDIRS
End If

sPath = BrowseFolder(UserControl.hWnd, lFlags, m_Caption, vStartFolder, vRootFolder, m_Proper)
If sPath <> "" Then
  ChangePath sPath
  RaiseEvent Click(m_Text)
  Show = True
Else
  Show = False
End If

End Function
Private Sub mFont_FontChanged(ByVal PropertyName As String)
'----------------------------------------------------------
'Change the fonts
'----------------------------------------------------------
'Set UserControl.Font = New_Font
'Set txtFolder.Font = New_Font
Set UserControl.Font = mFont
Set txtFolder.Font = mFont
ResizeCtrl
Refresh

End Sub

Private Sub txtFolder_Change()
'-----------------------------------------------
'Notify UserControl of change
'Use static variable to prevent recursive calls
'-----------------------------------------------
Static OldText As String

With txtFolder
  If OldText <> .Text Then
    ChangePath .Text
    OldText = .Text
  End If
End With

End Sub

Private Sub txtFolder_GotFocus()
'---------------------------------
'Select all of the text
'---------------------------------
With txtFolder
  .SelStart = 0
  .SelLength = Len(.Text)
End With

End Sub

Private Sub txtFolder_KeyDown(KeyCode As Integer, Shift As Integer)
'-------------------------------------------
'Handle KeyDown events
'-------------------------------------------
RaiseEvent KeyDown(KeyCode, Shift)

End Sub

Private Sub txtFolder_KeyPress(KeyAscii As Integer)
'-------------------------------------------
'Handle Keypress events
'-------------------------------------------
RaiseEvent KeyPress(KeyAscii)

End Sub

Private Sub txtFolder_KeyUp(KeyCode As Integer, Shift As Integer)
'-------------------------------------------
'Handle KeyUp events
'-------------------------------------------
RaiseEvent KeyUp(KeyCode, Shift)

End Sub

Private Sub txtFolder_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'----------------------------------------------------------------------
'Convert x,y positions to user control and
'then to Container scale
'----------------------------------------------------------------------
RaiseEvent MouseDown(Button, Shift, _
      ScaleX(x + txtFolder.Left, vbPixels, vbContainerPosition), _
      ScaleY(y + txtFolder.Top, vbPixels, vbContainerPosition))

End Sub

Private Sub txtFolder_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'----------------------------------------------------------------------
'Convert x,y positions to user control and
'then to Container scale
'----------------------------------------------------------------------
RaiseEvent MouseMove(Button, Shift, _
      ScaleX(x + txtFolder.Left, vbPixels, vbContainerPosition), _
      ScaleY(y + txtFolder.Top, vbPixels, vbContainerPosition))

End Sub

Private Sub txtFolder_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
'----------------------------------------------------------------------
'Convert x,y positions to user control and
'then to Container scale
'----------------------------------------------------------------------
RaiseEvent MouseUp(Button, Shift, _
      ScaleX(x + txtFolder.Left, vbPixels, vbContainerPosition), _
      ScaleY(y + txtFolder.Top, vbPixels, vbContainerPosition))

End Sub

Private Sub UserControl_AmbientChanged(PropertyName As String)
'-------------------------------------------------------------
'Update ambient properties
'-------------------------------------------------------------
Select Case PropertyName
Case "Font"
  Set Font = Ambient.Font
  mFont_FontChanged "Font"
Case "ForeColor"
  txtFolder.ForeColor = Ambient.ForeColor
End Select
End Sub

Public Sub ShowAboutBox()
Attribute ShowAboutBox.VB_UserMemId = -552
'----------------------------------
'Show the about box
'----------------------------------
dlgAbout.Show vbModal
Unload dlgAbout
Set dlgAbout = Nothing
End Sub

Private Sub UserControl_Click()
'-----------------------------------------------------
'Show browse dialog
'-----------------------------------------------------

'Raise the dropdown event
RaiseEvent Dropdown

'Show the browse for folder dialog
Show

End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
'----------------------------------------------------------------------
'Keypreview is set, so we get all of the keypresses here first.
'Check for keypresses which should cause the Browse dialog to show
'Alt and down arrow.
If KeyCode = vbKeyDown And Shift = 4 Then
  Show
End If

End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'--------------------------------------------------------------------------
'No matter where the UserCtrl is clicked, emulate the Dropdown button
'being clicked
'Also, pass on the event to the user
'--------------------------------------------------------------------------
PaintDropDown True
RaiseEvent MouseDown(Button, Shift, _
      ScaleX(x, vbPixels, vbContainerPosition), _
      ScaleY(y, vbPixels, vbContainerPosition))


End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'----------------------------------------------------------------------
'Control the state of DropDown button. To show the dropdown button in
'the 'down' state, two conditions must be fulfilled:-
'a) The left mouse button must be pressed (down)
'b) The mouse must be over the button
'In all other cases, the dropdown button is set to the 'Up' state
'Don't paint if state has not changed
'Also, pass on the event to the user
'----------------------------------------------------------------------
Dim cDown As Boolean        'New state

If Button = vbLeftButton Then
  cDown = PtInRect(rDrp, x, y) <> 0   'Returns 1 (inside) or 0 (outside)
Else
  cDown = False
End If
PaintDropDown cDown
RaiseEvent MouseDown(Button, Shift, _
      ScaleX(x, vbPixels, vbContainerPosition), _
      ScaleY(y, vbPixels, vbContainerPosition))

End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
'--------------------------------------------------------------------------
'Set the dropdown button state to 'Up'
'Pass on the event to the user
'--------------------------------------------------------------------------
PaintDropDown False
RaiseEvent MouseUp(Button, Shift, _
      ScaleX(x, vbPixels, vbContainerPosition), _
      ScaleY(y, vbPixels, vbContainerPosition))

End Sub

Private Sub UserControl_InitProperties()
'-------------------------------------------
'Initialize Properties for User Control
'-------------------------------------------
UserControl.BackColor = txtFolder.BackColor
txtFolder.ForeColor = Ambient.ForeColor
Set Font = Ambient.Font
mFont_FontChanged "Font"

m_RootFolder = m_def_RootFolder
m_CustomRootFolder = m_def_CustomRootFolder
m_Text = m_def_Text
ChangePath m_Text

m_Caption = m_def_Caption
m_Domain = m_def_Domain
m_ReturnAncestors = m_def_ReturnAncestors
m_ReturnFSDirs = m_def_ReturnFSDirs
m_Proper = m_def_Proper
Set m_Picture = Nothing
Set UserControl.Picture = LoadPicture("")
m_MaskColor = m_def_MaskColor

End Sub

Private Sub UserControl_Paint()
'-------------------------------------------------
'Repaint the control
'-------------------------------------------------
PaintMain

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
'------------------------------------------------------------
'Load property values from storage
'------------------------------------------------------------
txtFolder.BackColor = PropBag.ReadProperty("BackColor", &H80000005)
UserControl.BackColor = txtFolder.BackColor
UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
txtFolder.ForeColor = PropBag.ReadProperty("ForeColor", Ambient.ForeColor)
Set Font = PropBag.ReadProperty("Font", Ambient.Font)
mFont_FontChanged "Font"

m_RootFolder = PropBag.ReadProperty("RootFolder", m_def_RootFolder)
m_CustomRootFolder = PropBag.ReadProperty("CustomRootFolder", m_def_CustomRootFolder)
m_Text = PropBag.ReadProperty("Text", m_def_Text)
ChangePath m_Text

m_Caption = PropBag.ReadProperty("Caption", m_def_Caption)
m_Domain = PropBag.ReadProperty("Domain", m_def_Domain)
m_ReturnAncestors = PropBag.ReadProperty("ReturnAncestors", m_def_ReturnAncestors)
m_ReturnFSDirs = PropBag.ReadProperty("ReturnFSDirs", m_def_ReturnFSDirs)
m_Proper = PropBag.ReadProperty("Proper", m_def_Proper)

txtFolder.Locked = PropBag.ReadProperty("Locked", False)
txtFolder.ToolTipText = PropBag.ReadProperty("ToolTipText", "")
txtFolder.SelLength = PropBag.ReadProperty("SelLength", 0)
txtFolder.SelStart = PropBag.ReadProperty("SelStart", 0)
txtFolder.SelText = PropBag.ReadProperty("SelText", "")
Set m_Picture = PropBag.ReadProperty("DropdownPicture", Nothing)
Set UserControl.Picture = LoadPicture("")
m_MaskColor = PropBag.ReadProperty("MaskColor", m_def_MaskColor)

End Sub

Private Sub UserControl_Resize()
'---------------------------------------------------
'Resize the control and its constituents
'---------------------------------------------------
On Error Resume Next
ResizeCtrl

End Sub

Private Sub UserControl_Show()
'-------------------------------------------------
'Get the tooltip
'-------------------------------------------------
txtFolder.ToolTipText = UserControl.Extender.ToolTipText

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
'------------------------------------------------------------
'Write property values to storage
'------------------------------------------------------------

Call PropBag.WriteProperty("BackColor", txtFolder.BackColor, &H80000005)
Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
Call PropBag.WriteProperty("ForeColor", txtFolder.ForeColor, &H80000008)
Call PropBag.WriteProperty("Font", mFont, Ambient.Font)

Call PropBag.WriteProperty("RootFolder", m_RootFolder, m_def_RootFolder)
Call PropBag.WriteProperty("CustomRootFolder", m_CustomRootFolder, m_def_CustomRootFolder)
Call PropBag.WriteProperty("Text", m_Text, m_def_Text)

Call PropBag.WriteProperty("Caption", m_Caption, m_def_Caption)
Call PropBag.WriteProperty("Domain", m_Domain, m_def_Domain)
Call PropBag.WriteProperty("ReturnAncestors", m_ReturnAncestors, m_def_ReturnAncestors)
Call PropBag.WriteProperty("ReturnFSDirs", m_ReturnFSDirs, m_def_ReturnFSDirs)

Call PropBag.WriteProperty("Proper", m_Proper, m_def_Proper)
Call PropBag.WriteProperty("Locked", txtFolder.Locked, False)
Call PropBag.WriteProperty("ToolTipText", txtFolder.ToolTipText, "")
Call PropBag.WriteProperty("SelLength", txtFolder.SelLength, 0)
Call PropBag.WriteProperty("SelStart", txtFolder.SelStart, 0)
Call PropBag.WriteProperty("SelText", txtFolder.SelText, "")
Call PropBag.WriteProperty("DropdownPicture", m_Picture, Nothing)
Call PropBag.WriteProperty("MaskColor", m_MaskColor, m_def_MaskColor)

End Sub

'MappingInfo=txtFolder,txtFolder,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
'------------------------------------------------------------
'Get Property
'------------------------------------------------------------
ForeColor = txtFolder.ForeColor

End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
'------------------------------------------------------------
'Set Property
'------------------------------------------------------------
txtFolder.ForeColor() = New_ForeColor
PropertyChanged "ForeColor"

End Property

Private Sub PaintDropDown(Optional vDown As Variant)
'-----------------------------------------------------------------
'Paint the dropdown button
'State  : Normal (cDown=false), Down (cDown=true)
'         If cDown is omitted, use the previous state
'Width  : 13 pixels, including 2 pixel border
'Height : Normally 17 pixels (inside of edit box)
'         but changes with scaleheight
'-----------------------------------------------------------------
Static pDown As Boolean       'Previous state
Dim cDown As Boolean          'Current state
Dim pt As PointApi            'Not used, but required by API call
Dim c, Ok
Dim wd, hg                    'Height & Width of bitmap
Dim rWidth, rHeight           'Height & Width of rectangle
Dim x, y                      'Offsets for picture

If IsMissing(vDown) Then
  cDown = pDown
  Ok = True                   'Force a repaint
Else
  cDown = vDown
  Ok = cDown <> pDown         'Repaint only if state has changed
End If

If Ok Then
  'Get the rectangle area of the UsrCtrl and adjust to size
  'to get the rectangle of the dropdown button
  Ok = GetClientRect(UserControl.hWnd, rDrp)
  With rDrp
    .Top = .Top + 2
    .Right = .Right - 2
    .Left = .Right - mDropWidth
    .Bottom = .Bottom - 2
    rHeight = .Bottom - .Top + 1
    rWidth = .Right - .Left + 1
    c = .Bottom \ 2        'Center height
    If cDown Then
      '-----------------------------------
      'Button is in down (pressed) state
      '-----------------------------------
      'Draw the border
      Ok = DrawEdge(UserControl.hdc, rDrp, EDGE_SUNKEN, BF_RECT Or BF_FLAT)
      'Draw the face
      Line (.Left + 2, .Top + 2)-(.Right - 3, .Bottom - 3), vbButtonFace, BF
      If m_Picture Is Nothing Then
        'Draw triangle
        'Triangle is 3 lines high, 5 pixels first line, then 3, then 1
        'Remember that LineTo command does not draw last point
        UserControl.ForeColor = vbButtonText                  'Normally black
        'Ok = MoveToEx(UserControl.hDc, x, y, pt)             'Sample
        Ok = MoveToEx(UserControl.hdc, .Left + 5, c + 1, pt)  'Top line, left
        Ok = LineTo(UserControl.hdc, .Left + 10, c + 1)       'Top right
        Ok = MoveToEx(UserControl.hdc, .Left + 6, c + 2, pt)  'Mdl left
        Ok = LineTo(UserControl.hdc, .Left + 9, c + 2)        'Mdl right
        Ok = MoveToEx(UserControl.hdc, .Left + 7, c + 3, pt)  'Bot left
        Ok = LineTo(UserControl.hdc, .Left + 8, c + 3)        'Bot right
      Else
        'Get the size of the picture to draw
        GetPicSize m_Picture, hg, wd
        x = (rWidth - wd) \ 2
        If hg Mod 2 = 0 Then
          y = (rHeight - hg) \ 2
        Else
          y = (rHeight - hg) \ 2 + 1
        End If
        'TransparentBlt DestDC As Long, ByVal SrcBmp As Long, ByVal DestX As Long, ByVal DestY As Long, ByVal TransColor As Long
        TransparentBlt UserControl.hdc, m_Picture, .Left + x, .Top + y, m_MaskColor
      End If
    Else
      '-----------------------------------
      'Button is in up (normal) state
      '-----------------------------------
      'Draw the border
      Ok = DrawEdge(UserControl.hdc, rDrp, EDGE_RAISED, BF_RECT)
      'Draw the face
      Line (.Left + 2, .Top + 2)-(.Right - 3, .Bottom - 3), vbButtonFace, BF
      If m_Picture Is Nothing Then
        'Draw triangle
        'Triangle is 3 lines high, 5 pixels first line, then 3, then 1
        'Remember that LineTo command does not draw last point
        'Triangle moves one pixel down & one to the right
        UserControl.ForeColor = vbButtonText                 'Normally black
        'Ok = MoveToEx(UserControl.hdc, x, y, pt)            'Sample
        Ok = MoveToEx(UserControl.hdc, .Left + 4, c, pt)     'Top line, left
        Ok = LineTo(UserControl.hdc, .Left + 9, c)           'Top right
        Ok = MoveToEx(UserControl.hdc, .Left + 5, c + 1, pt)    'Mdl left
        Ok = LineTo(UserControl.hdc, .Left + 8, c + 1)         'Mdl right
        Ok = MoveToEx(UserControl.hdc, .Left + 6, c + 2, pt)    'Bot left
        Ok = LineTo(UserControl.hdc, .Left + 7, c + 2)          'Bot right
      Else
        'Get the size of the picture to draw
        GetPicSize m_Picture, hg, wd
        x = (rWidth - wd) \ 2 - 1
        If hg Mod 2 = 0 Then
          y = (rHeight - hg) \ 2 - 1
        Else
          y = (rHeight - hg) \ 2
        End If
        TransparentBlt UserControl.hdc, m_Picture, .Left + x, .Top + y, m_MaskColor
      End If
    End If
  End With
End If
pDown = cDown

End Sub

Private Sub ResizeCtrl()
'-------------------------------------------------------
'Resize the user control
'Use the Busy flag to prevent recursive calls
'-------------------------------------------------------
Static Busy As Boolean
Dim h, w

If Not (Busy) Then
  Busy = True
  '--------------------------------------
  'Validate height
  '--------------------------------------
  'Restrict to height of text + 2 pixels white space + 2 pxls border
  h = TextHeight("Dummy") + 4 * 2
  'Remember, UsrCtrl height may be in different scale modes,
  'depending on the container setting.
  'Therefore, scale from pixels to the appropriate size
  Height = ScaleY(h, vbPixels, vbContainerSize)
  '--------------------------------------
  'Validate width of main box
  '--------------------------------------
  'Same scaling applies to width
  w = ScaleX(Width, vbContainerSize, vbPixels)
  'Restrict minimum width
  If w < mMinWidth Then
    Width = ScaleX(mMinWidth, vbPixels, vbContainerSize)
  End If
  'Change the txtFolder size
  'UserControl.BackColor = RGB(255, 0, 0) 'for debug purposes
  'The txtfolder requires a 3 pixel border, 2 for the ctrl edge and
  '1 for the ctrl background. This is in sync with the std combobox
  'behaviour. Right edge requires only a 2 pixel border
  txtFolder.Move 3, 3, w - mDropWidth - 5, h - 6
  Busy = False
End If
 
End Sub

Private Sub PaintMain()
'---------------------------------------------------------------------------
'Paint the main edit box
'a) Paint 3D border
'---------------------------------------------------------------------------
'Notes:
'a) The dropdown button is not painted here - see PaintDropDown()
'b) The singlborder could not be used, as it adds a border to the
'   outside of the client area.
'---------------------------------------------------------------------------
Dim cFocus As Boolean         'Current focus
Dim rct As Rect               'UsrCtrl rectangle
Dim h, w                      'UsrCtrl height & width
Dim pt As PointApi            'Not used, but required by API call
Dim TextColor As Long         'Saved to draw text
Dim Ok

On Error GoTo PaintMainErr
'Get environment info
h = ScaleHeight
w = ScaleWidth
'Cls                           'Clear Usercontrol

'----------------------------------------------
'Draw the control border
'Reduces client area with 2 pixels on all sides
'The API DrawEdge() function is used with
'EDGE_SUNKEN : the type of border (raised/sunken)
'BF_RECT     : the sides to draw  (all 4 sides)
'----------------------------------------------
'Get the rectangle area of the UsrCtrl
Ok = GetClientRect(UserControl.hWnd, rct)
'Draw the border
Ok = DrawEdge(UserControl.hdc, rct, EDGE_SUNKEN, BF_RECT)

'Paint the dropdown button too
'Do not adjust the button state
PaintDropDown

Exit Sub

PaintMainErr:
Debug.Print "PaintMainErr: ", Err, Error
Resume Next
End Sub

Public Property Get Font() As StdFont
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
'------------------------------------------------------------
'Get Property
'------------------------------------------------------------
'Set Font = UserControl.Font
Set Font = mFont

End Property

Public Property Set Font(ByVal New_Font As StdFont)
'------------------------------------------------------------
'Set Property
'------------------------------------------------------------
If Not (New_Font Is Nothing) Then
  Set mFont = New_Font
  If Not (Ambient.UserMode) Then
    'In design mode the font changed event is not triggered
    'so manually trigger it to display the new font
    mFont_FontChanged "Font"
  End If
  PropertyChanged "Font"
End If

End Property

'MemberInfo=8,0,0,0
Public Property Get RootFolder() As SpecialFolders
Attribute RootFolder.VB_Description = "Sets/Returns the system folder to be used as the topmost folder in the dialog"
'-------------------------------------------------
'Read Rootfolder property
'Contains the system folder id's or -1 for Custom
'-------------------------------------------------
RootFolder = m_RootFolder

End Property

Public Property Let RootFolder(ByVal New_RootFolder As SpecialFolders)
'-------------------------------------------------
'Set Rootfolder property
'Contains the system folder id's or -1 for Custom
'-------------------------------------------------
m_RootFolder = New_RootFolder
PropertyChanged "RootFolder"

End Property

'MemberInfo=13,0,0,
Public Property Get CustomRootFolder() As String
Attribute CustomRootFolder.VB_Description = "Sets/Returns the custom root (topmost) folder. RootFolder property must be set to Custom."
'------------------------------------------------------------
'Returns the custom root folder. Only applicable when
'the RootFolder property is set to Custom
'------------------------------------------------------------
CustomRootFolder = m_CustomRootFolder

End Property

Public Property Let CustomRootFolder(ByVal New_CustomRootFolder As String)
'----------------------------------------------------------------------
'If the RootFolder is set to Custom, change the CustomRootFolder here
'Otherwise it has no effect
'----------------------------------------------------------------------
m_CustomRootFolder = New_CustomRootFolder
PropertyChanged "CustomRootFolder"

End Property

'MemberInfo=13,0,0,0
Public Property Get Text() As String
Attribute Text.VB_Description = "Returns the folder selected in the Folder Browse dialog."
Attribute Text.VB_ProcData.VB_Invoke_Property = ";Text"
Attribute Text.VB_UserMemId = 0
'------------------------------------------------
'Returns the current selected path, as displayed
'in the textbox
'------------------------------------------------
Text = m_Text

End Property

Public Property Let Text(ByVal New_Text As String)
'------------------------------------------------
'Sets the current path, which will be highlighted
'when the folderdialog opens
'------------------------------------------------
ChangePath New_Text

End Property

'MemberInfo=13,0,0,Select a folder
Public Property Get Caption() As String
Attribute Caption.VB_Description = "Sets/Returns the captio of the folder browse dialog."
Attribute Caption.VB_ProcData.VB_Invoke_Property = ";Data"
Attribute Caption.VB_UserMemId = -518
'-------------------------------------------------------
'Gets the caption of the folder dialog
'i.e. 'Please select a folder'
'-------------------------------------------------------
Caption = m_Caption

End Property

Public Property Let Caption(ByVal New_Caption As String)
'-------------------------------------------------------
'Sets the caption of the folder dialog
'i.e. 'Please select a folder'
'-------------------------------------------------------
m_Caption = New_Caption
PropertyChanged "Caption"

End Property

'MemberInfo=0,0,0,False
Public Property Get Domain() As Boolean
Attribute Domain.VB_Description = "Sets/Return the domain flag. If True, limits the folderselection to within the same domain."
'------------------------------------------------------
'Returns the BIF_DONTGOBELOWDOMAIN Flag
'------------------------------------------------------
Domain = m_Domain

End Property

Public Property Let Domain(ByVal New_Domain As Boolean)
'------------------------------------------------------
'Sets the BIF_DONTGOBELOWDOMAIN Flag
'------------------------------------------------------
m_Domain = New_Domain
PropertyChanged "Domain"

End Property

'MemberInfo=0,0,0,False
Public Property Get ReturnAncestors() As Boolean
Attribute ReturnAncestors.VB_Description = "Sets/Returns the ReturnFSAncestors flag. If true, only allows FSAncestors to be selected."
'-------------------------------------------------------
'Returns the BIF_ReturnFSAncestors Flag
'-------------------------------------------------------
ReturnAncestors = m_ReturnAncestors

End Property

Public Property Let ReturnAncestors(ByVal New_ReturnAncestors As Boolean)
'-------------------------------------------------------
'Sets the BIF_ReturnFSAncestors Flag
'-------------------------------------------------------
m_ReturnAncestors = New_ReturnAncestors
PropertyChanged "ReturnAncestors"

End Property

'MemberInfo=0,0,0,True
Public Property Get ReturnFSDirs() As Boolean
Attribute ReturnFSDirs.VB_Description = "Sets/Returns the ReturnOnlyFSDirs flag. If true, limits the folder selection to File System directories."
'-----------------------------------------------------
'Returns the BIF_RETURNONLYFSDIRS Flag
'-----------------------------------------------------
ReturnFSDirs = m_ReturnFSDirs

End Property

Public Property Let ReturnFSDirs(ByVal New_ReturnFSDirs As Boolean)
'-----------------------------------------------------
'Sets the BIF_RETURNONLYFSDIRS Flag
'-----------------------------------------------------
m_ReturnFSDirs = New_ReturnFSDirs
PropertyChanged "ReturnFSDirs"

End Property

'MemberInfo=0,0,0,False
Public Property Get Proper() As Boolean
Attribute Proper.VB_Description = "Sets/Returns a flag that converts the selected path to a proper string (first letter in each segment capitalised)"
'---------------------------------------------------
'Allows conversion of Path to proper string
'---------------------------------------------------
Proper = m_Proper

End Property

Public Property Let Proper(ByVal New_Proper As Boolean)
'---------------------------------------------------
'Allows conversion of Path to proper string
'---------------------------------------------------
m_Proper = New_Proper
PropertyChanged "Proper"
ChangePath m_Text

End Property

'MappingInfo=txtFolder,txtFolder,-1,Locked
Public Property Get Locked() As Boolean
Attribute Locked.VB_Description = "Determines whether a control can be edited."
'----------------------------------------
'Read locked property
'----------------------------------------
Locked = txtFolder.Locked

End Property

Public Property Let Locked(ByVal New_Locked As Boolean)
'----------------------------------------
'Set locked property
'----------------------------------------
txtFolder.Locked() = New_Locked
PropertyChanged "Locked"

End Property

'MappingInfo=txtFolder,txtFolder,-1,ToolTipText
Public Property Get ToolTipText() As String
Attribute ToolTipText.VB_Description = "Returns/sets the text displayed when the mouse is paused over the control."
'---------------------------------------------------
'Get tooltip text property
'---------------------------------------------------
ToolTipText = txtFolder.ToolTipText

End Property

Public Property Let ToolTipText(ByVal New_ToolTipText As String)
'---------------------------------------------------
'Set tooltip text property
'---------------------------------------------------
txtFolder.ToolTipText() = New_ToolTipText
PropertyChanged "ToolTipText"

End Property

'MappingInfo=txtFolder,txtFolder,-1,SelLength
Public Property Get SelLength() As Long
Attribute SelLength.VB_Description = "Returns/sets the number of characters selected."
SelLength = txtFolder.SelLength
End Property

Public Property Let SelLength(ByVal New_SelLength As Long)
txtFolder.SelLength() = New_SelLength
PropertyChanged "SelLength"
End Property

'MappingInfo=txtFolder,txtFolder,-1,SelStart
Public Property Get SelStart() As Long
Attribute SelStart.VB_Description = "Returns/sets the starting point of text selected."
SelStart = txtFolder.SelStart
End Property

Public Property Let SelStart(ByVal New_SelStart As Long)
txtFolder.SelStart() = New_SelStart
PropertyChanged "SelStart"
End Property

'MappingInfo=txtFolder,txtFolder,-1,SelText
Public Property Get SelText() As String
Attribute SelText.VB_Description = "Returns/sets the string containing the currently selected text."
SelText = txtFolder.SelText
End Property

Public Property Let SelText(ByVal New_SelText As String)
txtFolder.SelText() = New_SelText
PropertyChanged "SelText"
End Property

'MemberInfo=11,0,0,0
Public Property Get DropdownPicture() As Picture
Attribute DropdownPicture.VB_Description = "Returns/sets a 8x7 (wxh) pixel bitmap graphic to be displayed on the dropdown button of the control. If nothing, a std triangle is shown."
'--------------------------------------------------------------
'Dropdown picture is copied transparently to the dropdown box
'If set to nothing, an inverse triangle is drawn
'--------------------------------------------------------------
Set DropdownPicture = m_Picture

End Property

Public Property Set DropdownPicture(ByVal New_Picture As StdPicture)
'--------------------------------------------------------------
'Dropdown picture is copied transparently to the dropdown box
'If set to nothing, an inverse triangle is drawn
'--------------------------------------------------------------
Set m_Picture = New_Picture
PropertyChanged "DropdownPicture"
PaintDropDown

End Property

'MemberInfo=10,0,0,0
Public Property Get MaskColor() As OLE_COLOR
Attribute MaskColor.VB_Description = "Sets/Returns the transparent color for the picture"
'--------------------------------------------------------------
'Maskcolor is used to make the Dropdown picture transparent
'--------------------------------------------------------------
MaskColor = m_MaskColor

End Property

Public Property Let MaskColor(ByVal New_MaskColor As OLE_COLOR)
'--------------------------------------------------------------
'Maskcolor is used to make the Dropdown picture transparent
'--------------------------------------------------------------
m_MaskColor = New_MaskColor
PropertyChanged "MaskColor"

End Property

