VERSION 5.00
Begin VB.UserControl ArielBrowseFile 
   ClientHeight    =   570
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2520
   ScaleHeight     =   38
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   168
   ToolboxBitmap   =   "ArielBrowseFile.ctx":0000
   Begin VB.TextBox txtFile 
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "ArielBrowseFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'--------------------------------------------------------------------
'Module     : ArielBrowseFolder
'Description: Ariel Broswe Folder ActiveX Control
'Version    : V1.1.8 Sep 2000
'Release    : VB6
'Copyright  : © T De Lange, 2000
'--------------------------------------------------------------------
'V1.00    Sep 00 Original version, based on ColorCombo
'V1.1.8   Sep 00 Minor bugs removed
'--------------------------------------------------------------------
'Credits
'Brian Gillham, SafeCtx controls
'http: www.FailSafe.co.za
'MailTo: Brian@ FailSafe.co.za
'--------------------------------------------------------------------
Option Base 0
Option Explicit
DefLng A-N, P-Z
DefBool O

'---------------------------------------------------
'Public Enumerations
'---------------------------------------------------
Public Enum DialogType
  OpenFile = 0
  SaveFile = 1
End Enum

'---------------------------------------------------
'Internal constants
'---------------------------------------------------
Const mDropWidth = 13   'Width of dropdown button
Const mMinWidth = 100   'Min 13 for dropdown

'---------------------------------------------------
'Internal Control Variables
'---------------------------------------------------
Dim rDrp As Rect        'DropDown rectangle (incl border)

'---------------------------------------------------
'Control Properties & Events (using Wizard)
'---------------------------------------------------
'Default Property Values:
Const m_def_Caption = "Select a file"
Const m_def_File = ""
Const m_def_FileName = ""
Const m_def_FileDialogType = 0
Const m_def_Filter = "All Files (*.*) | *.*"
Const m_def_FilterIndex = 1
Const m_def_IncludePath = True
Const m_def_MaskColor = &HC0C0C0
Const m_def_Path = ""
Const m_def_Proper = False

'Property Variables:
Dim m_Caption As String             'Dialog Title
Dim m_File As String                'FileName incl path
Dim m_FileName As String            'Filename excl path, e.g. 'readme.txt'
Dim m_FileDialogType As DialogType  'Open or Save
Dim m_Filter As String              'Filter string e.g. 'All Files (*.*) | *.*'
Dim m_FilterIndex As Long           'Default filter element to show
Dim m_IncludePath As Boolean
Dim m_MaskColor As OLE_COLOR        'Transparent color for DropdownPicture
Dim m_Path As String                'Initial/Selected folder, e.g. 'C:\Windows'
Dim m_Picture As StdPicture         'Dropdown picture 8x10 pxl wxh. Height may vary
Dim m_Proper As Boolean             'Converts UCASE filename to Proper format
Dim m_Text As String                'Selected File name incl path, e.g. 'C:\Windows\readme.txt'
Private WithEvents mFont As StdFont
Attribute mFont.VB_VarHelpID = -1

'Event Declarations:
Event Dropdown()
Attribute Dropdown.VB_Description = "Fired when the dialog box is about to open. Useful to set properties on entry"
Event Click(File As String, Path As String, FileName As String)
Attribute Click.VB_Description = "Fired when a new file is selected"
Event Change(Text As String)
Attribute Change.VB_Description = "Fired when the text property changes"
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=txtFile,txtFile,-1,KeyDown
Attribute KeyDown.VB_Description = "Fired when a key is pressed"
Event KeyPress(KeyAscii As Integer) 'MappingInfo=txtFile,txtFile,-1,KeyPress
Attribute KeyPress.VB_Description = "Fired when a key is pressed"
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=txtFile,txtFile,-1,KeyUp
Attribute KeyUp.VB_Description = "Fired when a key has been pressed"
Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseDown
Attribute MouseDown.VB_Description = "Fired when the mouse button is pressed"
Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Attribute MouseMove.VB_Description = "Fired when the mouse is moved over the control"
Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseUp
Attribute MouseUp.VB_Description = "Fired when the mousebutton is released"

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

Private Type OPENFILENAME
  lStructSize As Long           'Length of structure, in bytes
  hWndOwner As Long             'Window that owns the dialog, or NULL
  hInstance As Long             'Handle of mem object containing template (not used)
  lpstrFilter As String         'File types/descriptions, delimited with chr(0), ends with 2xchr(0)
  lpstrCustomFilter As String   'Filters typed in by user
  nMaxCustFilter As Long        'Length of CustomFilter, min 40x chars
  nFilterIndex As Long          'Filter Index to use (1,2,etc) or 0 for custom
  lpstrFile As String           'Initial file/returned file(s), delimited with chr(0) for multi files
  nMaxFile As Long              'Size of Initial File string, min 256
  lpstrFileTitle As String      'File.ext excluding path
  nMaxFileTitle As Long         'Length of FileTitle
  lpstrInitialDir As String     'Initial file dir, null for current dir
  lpstrTitle As String          'Title bar of dialog
  flags As Long                 'See OFN_Flags
  nFileOffset As Integer        'Offset to file name in full path, 0-based
  nFileExtension As Integer     'Offset to file ext in full path, 0-based (excl '.')
  lpstrDefExt As String         'Default ext appended, excl '.', max 3 chars
  lCustData As Long             'Appl defined data for lpfnHook
  lpfnHook As Long              'Pointer to hook procedure
  lpTemplateName As String      'Template Name (not used)
End Type

'---------------------------------------------------
'Api Function Declarations
'---------------------------------------------------
'Common dialog
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long

'For general drawing
Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, lpPoint As PointApi) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As Rect, ByVal edge As Long, ByVal grfFlags As Long) As Long
'Note: the following declaration in the API viewer is incorrect!
'Private Declare Function PtInRect Lib "user32" (lpRect As Rect, pt As PointApi) As Long
Private Declare Function PtInRect Lib "user32" (lpRect As Rect, ByVal x As Long, ByVal y As Long) As Long

'General Windows User Interface
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As Rect) As Long

'---------------------------------------------------
'Api Constant Declarations
'---------------------------------------------------
Private Const OFN_READONLY = &H1
Private Const OFN_OVERWRITEPROMPT = &H2
Private Const OFN_HIDEREADONLY = &H4
Private Const OFN_NOCHANGEDIR = &H8
Private Const OFN_SHOWHELP = &H10
Private Const OFN_ENABLEHOOK = &H20
Private Const OFN_ENABLETEMPLATE = &H40
Private Const OFN_ENABLETEMPLATEHANDLE = &H80
Private Const OFN_NOVALIDATE = &H100
Private Const OFN_ALLOWMULTISELECT = &H200
Private Const OFN_EXTENSIONDIFFERENT = &H400
Private Const OFN_PATHMUSTEXIST = &H800
Private Const OFN_FILEMUSTEXIST = &H1000
Private Const OFN_CREATEPROMPT = &H2000
Private Const OFN_SHAREAWARE = &H4000
Private Const OFN_NOREADONLYRETURN = &H8000
Private Const OFN_NOTESTFILECREATE = &H10000
Private Const OFN_NONETWORKBUTTON = &H20000
Private Const OFN_NOLONGNAMES = &H40000
Private Const OFN_EXPLORER = &H80000
Private Const OFN_NODEREFERENCELINKS = &H100000
Private Const OFN_LONGNAMES = &H200000

Private Const OFN_SHAREFALLTHROUGH = 2
Private Const OFN_SHARENOWARN = 1
Private Const OFN_SHAREWARN = 0

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

Private Function ShowSaveDialog(sFilter As String, ByRef nFilterIndex As Long, sTitle As String, ByRef sPath As String, ByRef sFileName As String) As String
'--------------------------------------------------------------------------------
'Show the Save File dialog
'Filter       : Filter List i.e. "All Files (*.*)|*.*"
'FilterIndex  : Element no in Filter list, starting from 1
'               Updated on exit
'Title        : Dialog Title
'InitialPath  : Default Folder
'FileName     : On entry contains the initial file name excl path
'               On exit contains file name excl path, i.e. "command.com"
'--------------------------------------------------------------------------------
Dim ofn As OPENFILENAME
Dim i

ofn.lStructSize = Len(ofn)
ofn.hWndOwner = UserControl.hWnd
ofn.hInstance = App.hInstance
If Right(sFilter, 1) <> "|" Then sFilter = sFilter + "|"
For i = 1 To Len(sFilter)
  If Mid(sFilter, i, 1) = "|" Then
    Mid(sFilter, i, 1) = Chr$(0)
  End If
Next
ofn.lpstrFilter = sFilter
ofn.nFilterIndex = nFilterIndex
ofn.lpstrFile = Left(sFileName & Space(254), 254)
ofn.nMaxFile = 255
ofn.lpstrFileTitle = Space(254) 'on exit contains the filename
ofn.nMaxFileTitle = 255
ofn.lpstrInitialDir = sPath
ofn.lpstrTitle = sTitle
ofn.flags = OFN_HIDEREADONLY Or OFN_OVERWRITEPROMPT Or OFN_CREATEPROMPT
i = GetSaveFileName(ofn)

If (i) Then
  ShowSaveDialog = Left(ofn.lpstrFile, InStr(ofn.lpstrFile, vbNullChar) - 1)
  sPath = ValidateFolder(Left(ofn.lpstrFile, ofn.nFileOffset))
  sFileName = Left(ofn.lpstrFileTitle, InStr(ofn.lpstrFileTitle, vbNullChar) - 1)
  nFilterIndex = ofn.nFilterIndex
Else
  ShowSaveDialog = ""
End If

End Function

'MappingInfo=txtFile,txtFile,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Sets/returns the backcolor of the control (textbox portion)"
'-------------------------------------------
'Get Property
'-------------------------------------------
BackColor = txtFile.BackColor

End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
'-------------------------------------------
'Set Property
'-------------------------------------------
txtFile.BackColor() = New_BackColor
UserControl.BackColor = New_BackColor
PropertyChanged "BackColor"

End Property

'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Sets/return a boolean value that enables/disables all mouse action/editing."
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
Attribute Refresh.VB_Description = "Refreshes the control"
'-------------------------------------------
'Refresh the ctrl
'-------------------------------------------
UserControl.Refresh
   
End Sub
Private Sub mFont_FontChanged(ByVal PropertyName As String)
'----------------------------------------------------------
'Change the fonts
'----------------------------------------------------------
'Set UserControl.Font = New_Font
'Set txtFile.Font = New_Font
Set UserControl.Font = mFont
Set txtFile.Font = mFont
ResizeCtrl
Refresh

End Sub

Private Sub txtFile_Change()
'-----------------------------------------------
'Notify UserControl of change
'-----------------------------------------------
With txtFile
  If .Text <> m_Text Then
    ChangeText .Text, True
  End If
End With

End Sub

Private Sub txtFile_GotFocus()
'---------------------------------
'Select all of the text
'---------------------------------
With txtFile
  .SelStart = 0
  .SelLength = Len(.Text)
End With

End Sub
Private Sub txtFile_KeyDown(KeyCode As Integer, Shift As Integer)
'-------------------------------------------
'Handle KeyDown events
'-------------------------------------------
RaiseEvent KeyDown(KeyCode, Shift)

End Sub

Private Sub txtFile_KeyPress(KeyAscii As Integer)
'-------------------------------------------
'Handle Keypress events
'-------------------------------------------
RaiseEvent KeyPress(KeyAscii)

End Sub

Private Sub txtFile_KeyUp(KeyCode As Integer, Shift As Integer)
'-------------------------------------------
'Handle KeyUp events
'-------------------------------------------
RaiseEvent KeyUp(KeyCode, Shift)

End Sub
Private Sub txtFile_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'----------------------------------------------------------------------
'Convert x,y positions to user control and
'then to Container scale
'----------------------------------------------------------------------
RaiseEvent MouseDown(Button, Shift, _
      ScaleX(x + txtFile.Left, vbPixels, vbContainerPosition), _
      ScaleY(y + txtFile.Top, vbPixels, vbContainerPosition))

End Sub

Private Sub txtFile_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'----------------------------------------------------------------------
'Convert x,y positions to user control and
'then to Container scale
'----------------------------------------------------------------------
RaiseEvent MouseMove(Button, Shift, _
      ScaleX(x + txtFile.Left, vbPixels, vbContainerPosition), _
      ScaleY(y + txtFile.Top, vbPixels, vbContainerPosition))

End Sub
Private Sub txtFile_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
'----------------------------------------------------------------------
'Convert x,y positions to user control and
'then to Container scale
'----------------------------------------------------------------------
RaiseEvent MouseUp(Button, Shift, _
      ScaleX(x + txtFile.Left, vbPixels, vbContainerPosition), _
      ScaleY(y + txtFile.Top, vbPixels, vbContainerPosition))

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
  txtFile.ForeColor = Ambient.ForeColor
End Select

End Sub
Public Sub ShowAboutBox()
Attribute ShowAboutBox.VB_Description = "Displays the about box"
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
Dim sFile As String, sFileName As String, sPath As String
Dim nFilterIndex As Long
Dim Changed As Boolean
    
'Raise the dropdown event (prior to opening the dlg)
RaiseEvent Dropdown

'Show the browse for folder dialog
sPath = m_Path
sFileName = m_FileName
nFilterIndex = m_FilterIndex
If m_FileDialogType = OpenFile Then
  sFile = ShowOpenDialog(m_Filter, nFilterIndex, m_Caption, sPath, sFileName)
Else
  sFile = ShowSaveDialog(m_Filter, nFilterIndex, m_Caption, sPath, sFileName)
End If
If sFile <> "" Then
  'Change properties and raise click event
  'Also updates text and raise change event
  ChangeFile sFile, sPath, sFileName
End If
If nFilterIndex <> m_FilterIndex Then
  FilterIndex = nFilterIndex
End If

End Sub
Private Sub ChangeFile(sFile As String, sPath As String, sFileName As String, Optional ByTyping As Boolean = False)
'-----------------------------------------------------------------------
'Internal routine to maintain integrity of changes to the selected file
'sFile    : New Full path & file name
'sFileName: New Filename only
'(Note: InclPath property is used to determine what to show)
'If ByTyping is true, the text was entered by typing in the txtFile edit box
'-----------------------------------------------------------------------
Dim sText As String

If m_Proper Then
  m_File = StringToProper(sFile)
  m_Path = StringToProper(sPath)
  m_FileName = StringToProper(sFileName)
Else
  m_File = sFile
  m_Path = sPath
  m_FileName = sFileName
End If
If m_IncludePath Then
  sText = m_File
Else
  sText = m_FileName
End If

PropertyChanged "File"
PropertyChanged "Path"
PropertyChanged "FileName"
ChangeText sText, False
RaiseEvent Click(m_File, m_Path, m_FileName)

End Sub

Private Sub ChangeText(sText As String, Optional ByTyping As Boolean = False)
'-----------------------------------------------------------------------
'Internal routine to maintain integrity of changes to the text property
'If ByTyping is true, the text was entered by typing in the txtFile edit box
'-----------------------------------------------------------------------
If sText <> m_Text Then
  If m_Proper Then
    m_Text = StringToProper(sText)
  Else
    m_Text = sText
  End If
  If Not ByTyping Then
    txtFile = m_Text
  End If
  PropertyChanged "Text"
  RaiseEvent Change(m_Text)
End If

End Sub



Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
'----------------------------------------------------------------------
'Keypreview is set, so we get all of the keypresses here first.
'Check for keypresses which should cause the Browse dialog to show
'Alt and down arrow.
If KeyCode = vbKeyDown And Shift = 4 Then
  'Show
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
'Colors & Fonts
UserControl.BackColor = txtFile.BackColor
txtFile.ForeColor = Ambient.ForeColor
Set Font = Ambient.Font
mFont_FontChanged "Font"

'Pictures
m_MaskColor = m_def_MaskColor
Set m_Picture = Nothing
Set UserControl.Picture = LoadPicture("")

'Non display Properties
m_Caption = m_def_Caption
m_FileDialogType = m_def_FileDialogType
m_Filter = m_def_Filter
m_FilterIndex = m_def_FilterIndex
m_IncludePath = m_def_IncludePath
m_Proper = m_def_Proper

'Display
m_Path = m_def_Path
m_FileName = m_def_FileName
m_File = m_def_File
ChangeFile m_File, m_Path, m_FileName

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
'Constituent Properties
txtFile.BackColor = PropBag.ReadProperty("BackColor", &H80000005)
txtFile.ForeColor = PropBag.ReadProperty("ForeColor", Ambient.ForeColor)
txtFile.Locked = PropBag.ReadProperty("Locked", False)
txtFile.ToolTipText = PropBag.ReadProperty("ToolTipText", "")
'txtFile.SelLength = PropBag.ReadProperty("SelLength", 0)
'txtFile.SelStart = PropBag.ReadProperty("SelStart", 0)
'txtFile.SelText = PropBag.ReadProperty("SelText", "")

'UserControl Properties
UserControl.BackColor = txtFile.BackColor
UserControl.Enabled = PropBag.ReadProperty("Enabled", True)

'Font & Pictures
Set Font = PropBag.ReadProperty("Font", Ambient.Font)
mFont_FontChanged "Font"
m_MaskColor = PropBag.ReadProperty("MaskColor", m_def_MaskColor)
Set m_Picture = PropBag.ReadProperty("DropdownPicture", Nothing)
Set UserControl.Picture = LoadPicture("")

'Non-display Properties
m_Caption = PropBag.ReadProperty("Caption", m_def_Caption)
m_Proper = PropBag.ReadProperty("Proper", m_def_Proper)
m_FileDialogType = PropBag.ReadProperty("FileDialogType", m_def_FileDialogType)
m_Filter = PropBag.ReadProperty("Filter", m_def_Filter)
m_FilterIndex = PropBag.ReadProperty("FilterIndex", m_def_FilterIndex)
m_IncludePath = PropBag.ReadProperty("IncludePath", m_def_IncludePath)

'Display Properties
m_Path = PropBag.ReadProperty("Path", m_def_Path)
m_FileName = PropBag.ReadProperty("FileName", m_def_FileName)
m_File = PropBag.ReadProperty("File", m_def_File)
ChangeFile m_File, m_Path, m_FileName

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
txtFile.ToolTipText = UserControl.Extender.ToolTipText

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
'------------------------------------------------------------
'Write property values to storage
'------------------------------------------------------------
'Constituent Properties
Call PropBag.WriteProperty("BackColor", txtFile.BackColor, &H80000005)
Call PropBag.WriteProperty("ForeColor", txtFile.ForeColor, &H80000008)
Call PropBag.WriteProperty("Locked", txtFile.Locked, False)
Call PropBag.WriteProperty("ToolTipText", txtFile.ToolTipText, "")
'Call PropBag.WriteProperty("SelLength", txtFile.SelLength, 0)
'Call PropBag.WriteProperty("SelStart", txtFile.SelStart, 0)
'Call PropBag.WriteProperty("SelText", txtFile.SelText, "")

'UserControl
Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)

'Font & Pictures
Call PropBag.WriteProperty("Font", mFont, Ambient.Font)
Call PropBag.WriteProperty("MaskColor", m_MaskColor, m_def_MaskColor)
Call PropBag.WriteProperty("DropdownPicture", m_Picture, Nothing)

'Non-display Properties
Call PropBag.WriteProperty("Caption", m_Caption, m_def_Caption)
Call PropBag.WriteProperty("Proper", m_Proper, m_def_Proper)
Call PropBag.WriteProperty("FileDialogType", m_FileDialogType, m_def_FileDialogType)
Call PropBag.WriteProperty("Filter", m_Filter, m_def_Filter)
Call PropBag.WriteProperty("FilterIndex", m_FilterIndex, m_def_FilterIndex)
Call PropBag.WriteProperty("IncludePath", m_IncludePath, m_def_IncludePath)

'Display Properties
Call PropBag.WriteProperty("Path", m_Path, m_def_Path)
Call PropBag.WriteProperty("FileName", m_FileName, m_def_FileName)
Call PropBag.WriteProperty("File", m_File, m_def_File)

End Sub

'MappingInfo=txtFile,txtFile,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Sets/returns the forecolor in which the text property is displayed"
'------------------------------------------------------------
'Get Property
'------------------------------------------------------------
ForeColor = txtFile.ForeColor

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
    c = .Bottom \ 2        'Center height
    rHeight = .Bottom - .Top + 1
    rWidth = .Right - .Left + 1
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
  'Change the txtFile size
  'UserControl.BackColor = RGB(255, 0, 0) 'for debug purposes
  'The txtFile requires a 3 pixel border, 2 for the ctrl edge and
  '1 for the ctrl background. This is in sync with the std combobox
  'behaviour. Right edge requires only a 2 pixel border
  txtFile.Move 3, 3, w - mDropWidth - 5, h - 6
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
Attribute Font.VB_Description = "Sets/returns the font in which the text property is displayed"
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

'MemberInfo=13,0,0,0
Public Property Get Text() As String
Attribute Text.VB_Description = "Sets/returns the text to be displayed in the text box. If IncludePath is true, the full path/file name is displayed. If false, only the filename is displayed."
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
'ChangePath New_Text

End Property
'MemberInfo=13,0,0,Select a folder
Public Property Get Caption() As String
Attribute Caption.VB_Description = "Sets/returns the dialog title"
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
Public Property Get Proper() As Boolean
Attribute Proper.VB_Description = "Sets/returns a boolean value that allows all upper case path/file names (DOS) to be converted to proper format."
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
ChangeFile m_File, m_Path, m_FileName

End Property
'MappingInfo=txtFile,txtFile,-1,Locked
Public Property Get Locked() As Boolean
Attribute Locked.VB_Description = "Sets/returns the locked property of the text box portion"
'----------------------------------------
'Read locked property
'----------------------------------------
Locked = txtFile.Locked

End Property

Private Function ShowOpenDialog(ByVal sFilter As String, ByRef nFilterIndex As Long, ByVal sTitle As String, ByRef sPath As String, ByRef sFileName As String) As String
'------------------------------------------------------------------------------------
'Displays the Open File Dialog
'Filter       : Filter List i.e. "All Files (*.*)|*.*"
'FilterIndex  : Element no in Filter list, starting from 1
'               Updated on exit
'Title        : Dialog Title
'Path         : Default Folder on entry, selected folder on exit
'FileName     : On entry contains the initial file name excl path
'               On exit contains file name excl path, i.e. "command.com"
'------------------------------------------------------------------------------------
Dim ofn As OPENFILENAME
Dim i

ofn.lStructSize = Len(ofn)
ofn.hWndOwner = UserControl.hWnd
ofn.hInstance = App.hInstance

'Ensure | character added to end
If Right(sFilter, 1) <> "|" Then
  sFilter = sFilter + "|"
End If
'Replace the | character with chr(0)
For i = 1 To Len(sFilter)
  If Mid(sFilter, i, 1) = "|" Then
    Mid(sFilter, i, 1) = Chr(0)
  End If
Next
ofn.lpstrFilter = sFilter
ofn.nFilterIndex = nFilterIndex
'ofn.lpstrCustomFilter      'Not Used
'ofn.nMaxCustomFilter       'Not Used
ofn.lpstrFile = Left(sFileName & Space(254), 254)
ofn.nMaxFile = 255
ofn.lpstrFileTitle = Space(254) 'on exit contains the filename
ofn.nMaxFileTitle = 255
ofn.lpstrInitialDir = sPath
ofn.lpstrTitle = sTitle
ofn.flags = OFN_HIDEREADONLY Or OFN_FILEMUSTEXIST 'Or OFN_EXPLORER 'Or OFN_LONGNAMES
i = GetOpenFileName(ofn)

If (i) Then
  ShowOpenDialog = Left(ofn.lpstrFile, InStr(ofn.lpstrFile, vbNullChar) - 1)
  sPath = ValidateFolder(Left(ofn.lpstrFile, ofn.nFileOffset))
  sFileName = Left(ofn.lpstrFileTitle, InStr(ofn.lpstrFileTitle, vbNullChar) - 1)
  nFilterIndex = ofn.nFilterIndex
Else
  ShowOpenDialog = ""
End If

End Function

Public Property Let Locked(ByVal New_Locked As Boolean)
'----------------------------------------
'Set locked property
'----------------------------------------
txtFile.Locked() = New_Locked
PropertyChanged "Locked"

End Property

'MappingInfo=txtFile,txtFile,-1,ToolTipText
Public Property Get ToolTipText() As String
Attribute ToolTipText.VB_Description = "Standard tooltips"
'---------------------------------------------------
'Get tooltip text property
'---------------------------------------------------
ToolTipText = txtFile.ToolTipText

End Property

Public Property Let ToolTipText(ByVal New_ToolTipText As String)
'---------------------------------------------------
'Set tooltip text property
'---------------------------------------------------
txtFile.ToolTipText() = New_ToolTipText
PropertyChanged "ToolTipText"

End Property

'MappingInfo=txtFile,txtFile,-1,SelLength
Public Property Get SelLength() As Long
Attribute SelLength.VB_Description = "Standard text box"
'-------------------------------------------------------
'Expose textbox property
'-------------------------------------------------------
SelLength = txtFile.SelLength

End Property

Public Property Let SelLength(ByVal New_SelLength As Long)
'-------------------------------------------------------
'Expose textbox property
'-------------------------------------------------------
txtFile.SelLength() = New_SelLength
PropertyChanged "SelLength"

End Property

'MappingInfo=txtFile,txtFile,-1,SelStart
Public Property Get SelStart() As Long
Attribute SelStart.VB_Description = "Standard text box"
'-------------------------------------------------------
'Expose textbox property
'-------------------------------------------------------
SelStart = txtFile.SelStart

End Property

Public Property Let SelStart(ByVal New_SelStart As Long)
'-------------------------------------------------------
'Expose textbox property
'-------------------------------------------------------
txtFile.SelStart() = New_SelStart
PropertyChanged "SelStart"

End Property

'MappingInfo=txtFile,txtFile,-1,SelText
Public Property Get SelText() As String
Attribute SelText.VB_Description = "Standard text box"
'-------------------------------------------------------
'Expose textbox property
'-------------------------------------------------------
SelText = txtFile.SelText

End Property

Public Property Let SelText(ByVal New_SelText As String)
'-------------------------------------------------------
'Expose textbox property
'-------------------------------------------------------
txtFile.SelText() = New_SelText
PropertyChanged "SelText"
End Property

'MemberInfo=11,0,0,0
Public Property Get DropdownPicture() As Picture
Attribute DropdownPicture.VB_Description = "Sets/returns a picture to be drawn on the dropdown button. 8 pixels wide x 8 to 12 pixels height. If none, a triangle is drawn."
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
Attribute MaskColor.VB_Description = "Sets/returns the transparent color of the dropdown picture."
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





Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
'------------------------------------------------------------
'Set Property
'------------------------------------------------------------
txtFile.ForeColor() = New_ForeColor
PropertyChanged "ForeColor"

End Property



'MemberInfo=14,0,0,0
Public Property Get FileDialogType() As DialogType
Attribute FileDialogType.VB_Description = "Sets/Returns the behaviour of the file dialog to Open File or Save File"
'----------------------------------------------------------
'Gets the behaviour of the dialog (Open/Save mode)
'----------------------------------------------------------
FileDialogType = m_FileDialogType

End Property

Public Property Let FileDialogType(ByVal New_FileDialogType As DialogType)
'----------------------------------------------------------
'Sets the behaviour of the dialog (Open/Save mode)
'----------------------------------------------------------
m_FileDialogType = New_FileDialogType
PropertyChanged "FileDialogType"

End Property

'MemberInfo=13,0,0,All Files (*.*) | *.*
Public Property Get Filter() As String
Attribute Filter.VB_Description = "Sets/returns the file extension filters "
'----------------------------------------------
'Gets the file filter list
'----------------------------------------------
Filter = m_Filter

End Property

Public Property Let Filter(ByVal New_Filter As String)
'----------------------------------------------
'Sets the file filter list
'----------------------------------------------
Dim n, s() As String

If Right(New_Filter, 1) <> "|" Then
  New_Filter = New_Filter + "|"
End If
m_Filter = New_Filter
PropertyChanged "Filter"

'Adjust filter index if out of bounds

s = Split(m_Filter, "|")
n = (UBound(s) + 1) \ 2
If m_FilterIndex > n Then
  FilterIndex = n
End If

End Property

'MemberInfo=13,0,0,
Public Property Get Path() As String
Attribute Path.VB_Description = "Sets/returns the initial path (on entry) and the selected path (on exit)"
'----------------------------------------------------------
'On entry: Initial path
'On exit : Selected path
'----------------------------------------------------------
Path = m_Path

End Property

Public Property Let Path(ByVal New_Path As String)
'----------------------------------------------------------
'On entry: Initial path
'On exit : Selected path
'----------------------------------------------------------
If m_Proper Then
  m_Path = StringToProper(New_Path)
Else
  m_Path = New_Path
End If
PropertyChanged "Path"

End Property

'MemberInfo=13,0,0,
Public Property Get FileName() As String
Attribute FileName.VB_Description = "Sets/returns the initial filename on entry and the selected filename on exit, excluding the path."
'----------------------------------------------------------------
'On entry: Default file name (excl path)
'On exit : Selected file name (excl path)
'----------------------------------------------------------------
FileName = m_FileName

End Property

Public Property Let FileName(ByVal New_FileName As String)
'----------------------------------------------------------------
'On entry: Default file name (excl path)
'On exit : Selected file name (excl path)
'----------------------------------------------------------------
If m_Proper Then
  m_FileName = StringToProper(New_FileName)
Else
  m_FileName = New_FileName
End If
PropertyChanged "FileName"

End Property

'MemberInfo=8,0,0,1
Public Property Get FilterIndex() As Long
Attribute FilterIndex.VB_Description = "Sets/Returns the filter index, starting from 1, corresponding to the Filter list"
'-------------------------------------------------------------
'Get the default filter element no
'-------------------------------------------------------------
FilterIndex = m_FilterIndex

End Property

Public Property Let FilterIndex(ByVal New_FilterIndex As Long)
'-------------------------------------------------------------
'Sets the default filter element no
'-------------------------------------------------------------
Dim n, s() As String

'Adjust filter index if out of bounds
s = Split(m_Filter, "|")
n = (UBound(s) + 1) \ 2
If New_FilterIndex > n Then
  New_FilterIndex = n
End If
m_FilterIndex = New_FilterIndex
PropertyChanged "FilterIndex"

End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,True
Public Property Get IncludePath() As Boolean
Attribute IncludePath.VB_Description = "Sets/Returns a boolean value indicating if the path should be included in the edit box, along with the filename."
  IncludePath = m_IncludePath
End Property

Public Property Let IncludePath(ByVal New_IncludePath As Boolean)
'----------------------------------------------------------------
'Determines if the full path/file name must be shown in the
'text box (true) or only the filename (false)
'----------------------------------------------------------------
m_IncludePath = New_IncludePath
ChangeFile m_File, m_Path, m_FileName
PropertyChanged "IncludePath"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,
Public Property Get File() As String
Attribute File.VB_Description = "Sets/returns the full file name, incl path"
  File = m_File
End Property

Public Property Let File(ByVal New_File As String)
  m_File = New_File
  PropertyChanged "File"
End Property

