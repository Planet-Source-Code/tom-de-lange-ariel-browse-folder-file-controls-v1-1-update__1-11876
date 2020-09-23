http://www.planet-source-code.com
=============================================================
Ariel Browse Folder/File Controls
by Tom de Lange
e-mail: tomdl@attglobal.net
=============================================================
Two active-x controls are included in this OCX library, a folder and file selector. The controls resemble a combobox, and when the dropdown button is clicked, a dialog is shown in which a folder or file may be selected. The BrowseFolder control uses the SHBrowseForFolder API function, and the BrowseFile the GetOpenFileName and GetSaveFileName functions. A call back procedure is implemented showing the currently selected folder in the browse for folder dialog. The user has the choice to select any system folder (Desktop, My Computer etc) as the root, or a custom folder. The BrowseFile control returns the filename and path separately from the full path. The standard inverse triangle on the dropdown boxes may be substituted for any 8 pixel wide bitmap (with varying heights), with a choice of mask colors. Full set of events are included, click() upon dialog close, Change() and DropDown() prior to opening of the dialogs. Source also implements BitBlt() API function to copy transparent images. This is an update to the ArielBrowseFolder Control

Method
The BrowseFolder control uses the SHBrowseForFolder function, and the
BrowseFile the GetOpenFileName and GetSaveFileName functions of the
comdlg32.dll, amongst other APIs. A call back procedure is implemented 
showing the currently selected folder in the browse for folder dialog. 
User has the choice to select any system folder (Desktop, My Computer etc) 
as the root, or a custom folder. The BrowseFile control returns the filename
and path separately from the full path.
The standard inverse triangle on the dropdown boxes may be substituted for any 
8 pixel wide bitmap (with varying heights), with a choice of mask colors.
Full set of events are included, click() upon dialog close, Change() and 
DropDown() prior to opening of the dialogs. Source also implements BitBlt() 
API function to copy transparent images.

APIs used
FOLDER BROWSE
SHBrowseForFolder Lib "shell32" 
SHGetPathFromIDList Lib "shell32" 
SHGetFolderLocation Lib "shell32" 
SHGetSpecialFolderLocation Lib "shell32" 
SHSimpleIDListFromPath Lib "shell32" 

FILE BROWSE
GetOpenFileName Lib "comdlg32.dll"
GetSaveFileName Lib "comdlg32.dll"

GRAPHICS
MoveToEx Lib "gdi32" 
LineTo Lib "gdi32" 
DrawEdge Lib "user32" 
PtInRect Lib "user32"
GetClientRect Lib "user32"

BITMAPS
BitBlt Lib "gdi32" 
SetBkColor Lib "gdi32" 
CreateCompatibleDC Lib "gdi32" 
DeleteDC Lib "gdi32" 
CreateBitmap Lib "gdi32" 
CreateCompatibleBitmap Lib "gdi32" 
SelectObject Lib "gdi32" 
DeleteObject Lib "gdi32" 
GetObject Lib "gdi32" 


Distribution
Zip file contains full control source, images, screenshots and a testprogram
Source was developed in VB6 Professional, SP4 with Windows 98SE
Compatible with VB5 but not VB4-32bit or VB4-16bit.

Disclaimer
This example program is provided "as is" with no warranty of any kind. It is
intended for demonstration purposes only. You can use the example in any form, 
but please mention the author

Credits
Per Andersson, FireStorm@GoToMy.com, www.FireStormEntertainment.cjb.net
Roman Blachman, romaz@inter.net.il
Stephen Fonnesbeck, steev@xmission.com, http://www.xmission.com/~steev
Max Raskin, www.planet-source-code.com
Brian Gillham, SafeCtx controls, http: www.FailSafe.co.za, Brian@ FailSafe.co.za
