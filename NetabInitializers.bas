Attribute VB_Name = "NetabInitializers"
Option Explicit
'Constants
Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

Private Const TRANSPARENT = 1
Private Const DT_CALCRECT = &H400
Private Const DT_CENTER = &H1
Private Const DT_VCENTER = &H4
Private Const DT_SINGLELINE = &H20
Private Const DT_RIGHT = &H2
Private Const DT_BOTTOM = &H8
Public Const MAX_PATH = 260

'Variables
Dim lL As Long, lT As Long, sB As Long
Dim i As Byte
Dim z As Byte
Dim Registry As New RegEdit
Global TypedUrls As Variant
Public Function GetSystemPath() As String
On Error Resume Next
Dim strFolder As String
Dim lngResult As Long
strFolder = String(MAX_PATH, 0)
lngResult = GetSystemDirectory(strFolder, MAX_PATH)
If lngResult <> 0 Then
    GetSystemPath = Left(strFolder, InStr(strFolder, Chr(0)) - 1)
Else
    GetSystemPath = ""
End If
End Function
Public Function GetTypedUrl()
On Error Resume Next
Registry.OpenRegistry HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\TypedURLs"
TypedUrls = Registry.GetAllValues
For z = LBound(TypedUrls) To UBound(TypedUrls)
    frmMain.cboAddress.AddItem Registry.GetValue(TypedUrls(z))
Next z
End Function

Public Function SaveTypedUrl()

End Function
Public Function InterfaceSetup()
On Error Resume Next
With frmMain

'Initializes the Status Bar

.StatusBar.AddPanel estbrStandard, , , , 550, , , , "1"
.StatusBar.AddPanel estbrStandard, , , , , , , , "2"
.StatusBar.AddPanel estbrStandard, "Internet"
'-----------------------
.Tabmain.AddTab "about:blank", -1, -1, "Tab"
'------------------------

'-----TEMP
.IE(1).Navigate "about:blank"
.cboAddress.Text = "about:blank"
.IE(1).Visible = True
'---------


   With frmMain.tbrTools
      .ImageSource = CTBExternalImageList
      .SetImageList frmMain.m_cILNormal, CTBImageListNormal
      .SetImageList frmMain.m_cILHot, CTBImageListHot
      
      .CreateToolbar 22, , False, False, 16
      .AddButton "New Tab", 11, , , "", CTBNormal, "NEW"
      .AddButton "Delete Tab", 9, , , "", CTBNormal, "DELETETAB"
      .AddButton , , , , , CTBSeparator
      .AddButton "Back", 0, , , "", CTBNormal, "BACK"
      .AddButton "Forward", 1, , , "", CTBNormal, "FORWARD"
      .AddButton "Stop", 12, , , "", , "STOP"
      .AddButton "Refresh", 10, , , "", CTBNormal, "REFRESH"
      .AddButton "Home", 17, , , "", , "HOME"
      .AddButton , , , , , CTBSeparator
      .AddButton "Search", 19, , , "", , "SEARCH"
      .AddButton "Favorites", 15, , , "", , "FAVORITES"
      .AddButton "History", 22, , , "", , "HISTORY"
      .AddButton , , , , , CTBSeparator
      .AddButton "Copy", 3, , , "", , "COPY"
      .AddButton "Paste", 4, , , "", , "PASTE"
      .AddButton , , , , , CTBSeparator
      
      .AddButton "IE Options", 8, , , "", , "IEOPTIONS"
      .AddButton "Settings", 30, , , "", CTBNormal, "SETTINGS"
      .ButtonChecked(15) = frmMain.TreeView1.Visible
   End With
         
  
   With frmMain.rbrMain
      .Position = erbPositionTop
      .CreateRebar frmMain.hwnd
      .AddBandByHwnd frmMain.tbrmenu.hwnd, , , , "MENU"
      
      .AddBandByHwnd frmMain.tbrTools.hwnd, , , , "TOOLBAR"
      .AddBandByHwnd frmMain.picAnim.hwnd, , False, True, "ANIM"
      .AddBandByHwnd frmMain.cboAddress.hwnd, "Address", , , "ADDRESS"
      For i = 0 To .BandCount - 1
         If i <> 1 Then
            .BandChildMinWidth(i) = 16
         End If
      Next i
   End With
   
   ' These borders are only visible so you can see at Design
   .picTitle.BorderStyle = 0
   .picFolders.BorderStyle = 0
   
   With frmMain.tbrClose
      .ImageSource = CTBExternalImageList
      .SetImageList frmMain.imgClose, CTBImageListNormal
      .CreateToolbar 13
      .AddButton "Close", 0
      
   End With
   
   With frmMain.tbrHost
      .Capture frmMain.picTitle
      .Capture frmMain.tbrClose
   End With
   frmMain.tbrHost.BorderStyle = etbhBorderStyleNone
 
 'Setup the status bar
 With frmMain.rbrMain
.BandGripper(0) = False
.BandGripper(2) = False
.BandGripper(3) = False
.BandChildEdge(0) = False
End With

 frmMain.Visible = True
End With
End Function

Public Function PrepareUnload()
On Error Resume Next
With frmMain
   .rbrMain.RemoveAllRebarBands
   .rbrMain.DestroyRebar
   .rbrSide.RemoveAllRebarBands
   .rbrSide.DestroyRebar
End With
Unload frmMain
End Function
Public Function ResizingMain()
On Error Resume Next

frmMain.rbrMain.RebarSize
frmMain.rbrSide.RebarSize
lL = (frmMain.rbrSide.RebarWidth + 4) * Screen.TwipsPerPixelX
lT = frmMain.rbrMain.RebarHeight * Screen.TwipsPerPixelY
sB = frmMain.StatusBar.Height + 4 * Screen.TwipsPerPixelY
frmMain.Tabmain.Move lL, lT, frmMain.ScaleWidth - lL, frmMain.ScaleHeight - sB - lT
frmMain.TreeView1.Move frmMain.TreeView1.Left, 240, frmMain.ScaleWidth - frmMain.Tabmain.Width, frmMain.Tabmain.Height - 300
frmMain.Search.Move frmMain.Search.Left, 240, frmMain.ScaleWidth - frmMain.Tabmain.Width, frmMain.Tabmain.Height - 300

For i = 1 To frmMain.IE.UBound
If frmMain.Tabmain.TabAlign = etaTop Then frmMain.IE(i).Move frmMain.IE(i).Left, 360, frmMain.Tabmain.Width - 240, frmMain.Tabmain.Height - 480
If frmMain.Tabmain.TabAlign = etaBottom Then frmMain.IE(i).Move frmMain.IE(i).Left, 120, frmMain.Tabmain.Width - 240, frmMain.Tabmain.Height - 480
Next
frmMain.Refresh
End Function

