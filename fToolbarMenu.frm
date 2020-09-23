VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{54F463F3-0135-11D2-8D52-00C04FA4EE99}#7.2#0"; "vbalTbar.ocx"
Object = "{F1909D6D-FB9D-11D3-B06C-00500427A693}#1.0#0"; "xuiTreeView6.ocx"
Object = "{1FE9A10D-50A4-431B-89AE-610EC623D3F1}#1.0#0"; "vbalIml.ocx"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   Caption         =   "Netab Preview Release 1"
   ClientHeight    =   7368
   ClientLeft      =   3768
   ClientTop       =   2220
   ClientWidth     =   9348
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.4
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "fToolbarMenu.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7368
   ScaleWidth      =   9348
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin vbalIml6.vbalImageList imgClose 
      Left            =   2760
      Top             =   960
      _ExtentX        =   762
      _ExtentY        =   762
      IconSizeX       =   14
      IconSizeY       =   14
      ColourDepth     =   8
      Size            =   10444
      Images          =   "fToolbarMenu.frx":02EA
      KeyCount        =   14
      Keys            =   "ÿÿÿÿÿÿÿÿÿÿÿÿÿ"
   End
   Begin VB.Timer tmranim 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   4800
      Top             =   960
   End
   Begin VB.PictureBox Picanimsrc 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   5748
      Left            =   8880
      Picture         =   "fToolbarMenu.frx":2BD6
      ScaleHeight     =   5700
      ScaleWidth      =   300
      TabIndex        =   10
      Top             =   1200
      Visible         =   0   'False
      Width           =   348
   End
   Begin Netab.vbalStatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   252
      Left            =   0
      TabIndex        =   9
      Top             =   7116
      Width           =   9348
      _ExtentX        =   16489
      _ExtentY        =   445
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   -2147483633
   End
   Begin vbalTBar.cToolbar tbrmenu 
      Left            =   4560
      Top             =   360
      _ExtentX        =   5313
      _ExtentY        =   445
   End
   Begin vbalIml6.vbalImageList m_cILMenu 
      Left            =   4200
      Top             =   960
      _ExtentX        =   762
      _ExtentY        =   762
      ColourDepth     =   16
      Size            =   31960
      Images          =   "fToolbarMenu.frx":B91E
      KeyCount        =   34
      Keys            =   "ÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿ"
   End
   Begin vbalIml6.vbalImageList m_cILNormal 
      Left            =   3720
      Top             =   960
      _ExtentX        =   762
      _ExtentY        =   762
      IconSizeX       =   24
      IconSizeY       =   24
      ColourDepth     =   16
      Size            =   62592
      Images          =   "fToolbarMenu.frx":13616
      KeyCount        =   32
      Keys            =   "ÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿ"
   End
   Begin vbalIml6.vbalImageList m_cILHot 
      Left            =   3240
      Top             =   960
      _ExtentX        =   762
      _ExtentY        =   762
      IconSizeX       =   24
      IconSizeY       =   24
      ColourDepth     =   16
      Size            =   62592
      Images          =   "fToolbarMenu.frx":22AB6
      KeyCount        =   32
      Keys            =   "ÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿ"
   End
   Begin VB.PictureBox picFolders 
      Height          =   5592
      Left            =   120
      ScaleHeight     =   5544
      ScaleWidth      =   2244
      TabIndex        =   6
      Top             =   1320
      Visible         =   0   'False
      Width           =   2292
      Begin xuiTreeView6.TreeView TreeView1 
         Height          =   5532
         Left            =   0
         TabIndex        =   8
         Top             =   240
         Width           =   2292
         _ExtentX        =   4043
         _ExtentY        =   9758
         Lines           =   0   'False
         LabelEditing    =   0   'False
         PlusMinus       =   0   'False
         RootLines       =   0   'False
         ToolTips        =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxScrollTime   =   0
         InternalBorderX =   2
         InternalBorderY =   2
      End
      Begin vbalTBar.cToolbarHost tbrHost 
         Height          =   192
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Visible         =   0   'False
         Width           =   2172
         _ExtentX        =   3831
         _ExtentY        =   339
      End
   End
   Begin Netab.TabControl Tabmain 
      Height          =   5772
      Left            =   2520
      TabIndex        =   4
      Top             =   1200
      Width           =   6732
      _ExtentX        =   11875
      _ExtentY        =   10181
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin SHDocVwCtl.WebBrowser IE 
         Height          =   5292
         Index           =   1
         Left            =   120
         TabIndex        =   5
         Top             =   480
         Width           =   6972
         ExtentX         =   12298
         ExtentY         =   9334
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   "res://C:\WINNT\system32\shdoclc.dll/dnserror.htm#http:///"
      End
   End
   Begin VB.PictureBox picTitle 
      Height          =   252
      Left            =   120
      ScaleHeight     =   204
      ScaleWidth      =   1140
      TabIndex        =   2
      Top             =   840
      Width           =   1188
      Begin VB.Label lblCaption 
         Caption         =   "Folders"
         Height          =   252
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   732
      End
   End
   Begin vbalTBar.cToolbar tbrClose 
      Left            =   2160
      Top             =   840
      _ExtentX        =   550
      _ExtentY        =   445
   End
   Begin vbalTBar.cReBar rbrSide 
      Left            =   60
      Top             =   840
      _ExtentX        =   4149
      _ExtentY        =   10922
   End
   Begin VB.PictureBox picAnim 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   348
      Left            =   8880
      ScaleHeight     =   348
      ScaleWidth      =   348
      TabIndex        =   1
      Top             =   120
      Width           =   348
   End
   Begin vbalTBar.cToolbar tbrTools 
      Left            =   720
      Top             =   120
      _ExtentX        =   5525
      _ExtentY        =   656
   End
   Begin VB.ComboBox cboAddress 
      Height          =   300
      Left            =   5400
      TabIndex        =   0
      Top             =   840
      Width           =   3852
   End
   Begin vbalTBar.cReBar rbrMain 
      Left            =   120
      Top             =   120
      _ExtentX        =   16108
      _ExtentY        =   1185
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public WithEvents m_cMenu As cPopupMenu
Attribute m_cMenu.VB_VarHelpID = -1
Public Function SidePanel(Status As String, Optional Section As String)
Status = LCase(Status)
If Len(Section) Then Section = LCase(Section)
LockWindowUpdate Me.hwnd
With rbrSide
If Status = "show" Then
picFolders.Visible = True
tbrHost.Visible = True
.Position = erbPositionLeft
.CreateRebar Me.hwnd
.AddBandByHwnd picFolders.hwnd, , False, True, "FOLDERBAND"
.BandGripper(0) = False
.BandChildEdge(0) = False
.BandChildEdge(1) = False
ResizingMain
Else
.RemoveAllRebarBands
.DestroyRebar
ResizingMain
End If
End With
Select Case Section
Case "favorites"
If Status = "show" Then
lblCaption = "Favorites"
TreeView1.Visible = True
Else
lblCaption = ""
TreeView1.Visible = False
End If
Case Else
ResizingMain
End Select
LockWindowUpdate 0
End Function
Private Sub RemoveTab()
Unload IE(Tabmain.SelectedTab)
Tabmain.RemoveTab (Tabmain.SelectedTab)
End Sub
Private Sub NewTab()
Dim IndexAdd As Integer
IndexAdd = IE.UBound + 1
Load IE(IndexAdd)
Tabmain.AddTab "about:blank", , IndexAdd, IndexAdd
Tabmain.SelectTab IndexAdd
IE(IndexAdd).Navigate "about:blank"
IE(IndexAdd).Visible = True
IE(IndexAdd).ZOrder 0
End Sub
Private Sub pShowBitmap(ByVal bState As Boolean)
      ' To change the background bitmap, we remove all bands
      ' and add them in again.
      ' In order to prevent flickering whilst the rebar builds,
      ' use LockWindowUpdate.  See Tips on vbAccelerator for
      ' more info.
'   LockWindowUpdate Me.hwnd
'   With rbrMain
'      .ImageSource = CRBLoadFromFile
'      If (bState) Then
'         .DestroyRebar
'         .ImageFile = App.Path & "\iebar.bmp"
'         .CreateRebar Me.hwnd
'      Else
'         .DestroyRebar
'         .ImageFile = ""
'         .CreateRebar Me.hwnd
'      End If
'      .AddBandByHwnd tbrmenu.hwnd, , , , "MENU"
'      .AddBandByHwnd picAnim.hwnd, , False, True, "ANIM"
'      If m_cMenu.Checked(m_cMenu.IndexForKey("mnuToolbar(0)")) Then
'         .AddBandByHwnd tbrTools.hwnd, , , , "TOOLBAR"
'      End If
'      If m_cMenu.Checked(m_cMenu.IndexForKey("mnuToolbar(1)")) Then
'         .AddBandByHwnd cboAddress.hwnd, "Address", , , "ADDRESS"
'      End If
'   End With
'   LockWindowUpdate 0

End Sub



Private Sub cboAddress_Click()
IE(Tabmain.SelectedTab).Navigate cboAddress.List(cboAddress.ListIndex)
End Sub
Private Sub Form_Load()
'Builds the menu windows
pBuildMenu
tbrmenu.CreateFromMenu m_cMenu
'-----------------------

InterfaceSetup
GetTypedUrl
ResizingMain
ResizeControls Me, 0
End Sub
Private Sub Form_Paint()
Me.Refresh
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
PrepareUnload
End Sub

Private Sub IE_ProgressChange(Index As Integer, ByVal Progress As Long, ByVal ProgressMax As Long)
If Progress > 0 Then 'Show percentage loaded in statusbar
StatusBar.PanelText(2) = Int((Progress / ProgressMax) * 100) & "%"
Else
Progress = 100
End If

If Progress = ProgressMax Then 'turn animation on and off for browser
tmranim.Enabled = False
BitBlt picAnim.hdc, 3, 4, picAnim.Width, picAnim.Height, Picanimsrc.hdc, 0, 3, vbSrcCopy
Else
tmranim.Enabled = True
End If
End Sub


Private Sub IE_StatusTextChange(Index As Integer, ByVal Text As String)
'shows the current status in the statusbar
StatusBar.PanelText(1) = Text
End Sub

Private Sub IE_TitleChange(Index As Integer, ByVal Text As String)
'On Error Resume Next
'Changes the tabs caption, and the main caption!
Dim TitleCaption As String
TitleCaption = Text
If Len(TitleCaption) >= 20 Then
TitleCaption = Left(Text, 15) & "..."
End If
LockWindowUpdate Me.hwnd
Tabmain.RemoveTab Index
Tabmain.AddTab TitleCaption, , Index, Index
IE(Index).Visible = True
IE(Index).ZOrder 0
Tabmain.SelectTab Index
IE(Index).SetFocus
LockWindowUpdate 0
Me.Caption = Text & App.ProductName & " " & App.Major & "." & App.Minor & cboAddress.Text = IE(Index).LocationURL
End Sub

Private Sub m_cMenu_Click(ItemNumber As Long)
Dim bS As Boolean

   Select Case m_cMenu.ItemKey(ItemNumber)
   Case "NewTab"
            NewTab
   Case "Open"
           OpenURL
   Case "SaveAs"
             IE(Tabmain.SelectedTab).ExecWB OLECMDID_SAVEAS, OLECMDEXECOPT_PROMPTUSER
   Case "PageSetup"
            IE(Tabmain.SelectedTab).ExecWB OLECMDID_PAGESETUP, OLECMDEXECOPT_DODEFAULT
   Case "Print"
            IE(Tabmain.SelectedTab).ExecWB OLECMDID_PRINT, OLECMDEXECOPT_DODEFAULT
   Case "CreateShortcut"
   CreateInternetShortcut IE(Tabmain.SelectedTab).LocationName, IE(Tabmain.SelectedTab).LocationURL
   Case "Properties"
            IE(Tabmain.SelectedTab).ExecWB OLECMDID_PROPERTIES, OLECMDEXECOPT_DODEFAULT
      

   Case "Close"
        Unload Me
        End
   Case "Copy"
            IE(Tabmain.SelectedTab).ExecWB OLECMDID_COPY, OLECMDEXECOPT_DODEFAULT
   Case "Paste"
            IE(Tabmain.SelectedTab).ExecWB OLECMDID_PASTE, OLECMDEXECOPT_DODEFAULT
   Case "SelectAll"
            IE(Tabmain.SelectedTab).ExecWB OLECMDID_SELECTALL, OLECMDEXECOPT_DODEFAULT
   Case "Find"
            IE(Tabmain.SelectedTab).ExecWB OLECMDID_FIND, OLECMDEXECOPT_DODEFAULT
   Case "VAddressBar"
            
   Case "VStatusBar"
   
   Case "Stop"
            IE(Tabmain.SelectedTab).Stop
   Case "Refresh"
            IE(Tabmain.SelectedTab).Refresh
   Case "Source"
        
   Case "FullScreen"
   
   Case "AddFavorites"
            
   Case "OrganizeFavorites"
   
   Case "ShowFavorites"
   lblCaption = "Favorites"
   TreeView1.Visible = True
   SidePanel ("Show")

   Case "HideFavorites"
   lblCaption = ""
   TreeView1.Visible = False
    SidePanel ("Hide")

   Case "WindowsUpdate"
   
   Case "NetabUpdate"
   
   Case "InternetOptions"
   
   Case "NetabSettings"
   
   Case "TipOfDay"
   
   Case "Feedback"
   
   Case "Elucid"
   
   Case "About"
    frmAbout.Show , Me
   Case Else
      Dim FavStr$, a%
   FavStr$ = m_cMenu.ItemKey(ItemNumber)
   If InStr(FavStr$, "Folder") = 0 Then
   a% = InStr(FavStr$, "URL")
    If a% > 0 Then
       FavStr$ = Right$(FavStr$, Len(FavStr$) - 3)
       FavStr$ = Trim(FavStr$)
       IE(Tabmain.SelectedTab).Navigate FavStr$
       cboAddress.Text = FavStr$
       End If
     Else
     MsgBox "This folder is empty!  You may add an item by going to add to favorites and clicking on this folder.", vbExclamation + vbOKOnly, "Empty Favorite Folder"

    End If
   End Select
   
End Sub

Private Sub newtab_filemnu_Click()
NewTab
End Sub

Private Sub picFolders_Resize()
Dim lW As Long
Dim llW As Long
Dim lH As Long
   
   lW = picFolders.ScaleWidth
   llW = lW - (tbrClose.ToolbarWidth + 4) * Screen.TwipsPerPixelX
   If llW > 0 Then
      picTitle.Width = llW
   Else
      picTitle.Width = 0
   End If
   lH = (tbrClose.ToolbarHeight + 2) * Screen.TwipsPerPixelY
   lblCaption.Move 2 * Screen.TwipsPerPixelX, (lH - lblCaption.Height) \ 2
   tbrHost.Move 0, 0, lW, lH
   tbrHost.Refresh
   
End Sub

Private Sub rbrMain_BandChildResize(ByVal wID As Long, ByVal lBandLeft As Long, ByVal lBandTop As Long, ByVal lBandRight As Long, ByVal lBandBottom As Long, lChildLeft As Long, lChildTop As Long, lChildRight As Long, lChildBottom As Long)
'   If rbrMain.BandData(rbrMain.BandIndexForId(wID)) = "ANIM" Then
'      picAnim.Width = (lChildRight - lChildLeft + 30) * Screen.TwipsPerPixelX
'      picAnim.Height = (lChildBottom - lChildTop - 10) * Screen.TwipsPerPixelX
'   End If
End Sub

Private Sub rbrMain_HeightChanged(lNewHeight As Long)
ResizingMain
End Sub

Private Sub refresh_viewmnu_Click()
IE(Tabmain.SelectedTab).Refresh
End Sub

Private Sub Tabmain_TabClick(ByVal lTab As Long)
IE(Tabmain.SelectedTab).Visible = True
IE(Tabmain.SelectedTab).ZOrder 0
End Sub

Private Sub tbrClose_ButtonClick(ByVal lButton As Long)
SidePanel ("Hide")
ResizingMain
End Sub
'========================================FAVORITES============

Private Sub tbrTools_ButtonClick(ByVal lButton As Long)
On Error Resume Next
Select Case tbrTools.ButtonKey(lButton)
Case "NEW"
    NewTab
    
Case "DELETETAB"
    RemoveTab
Case "BACK"
IE(Tabmain.SelectedTab).GoBack
Case "FORWARD"

IE(Tabmain.SelectedTab).GoForward
Case "HOME"
IE(Tabmain.SelectedTab).GoHome
Case "STOP"
IE(Tabmain.SelectedTab).Stop
Case "REFRESH"
IE(Tabmain.SelectedTab).Refresh
Case "CUT"
IE(Tabmain.SelectedTab).ExecWB OLECMDID_CUT, OLECMDEXECOPT_DODEFAULT
Case "COPY"
IE(Tabmain.SelectedTab).ExecWB OLECMDID_COPY, OLECMDEXECOPT_DODEFAULT
Case "PASTE"
IE(Tabmain.SelectedTab).ExecWB OLECMDID_PASTE, OLECMDEXECOPT_DODEFAULT
Case "FAVORITES"
If TreeView1.Visible = False Then
   SidePanel "Show", "Favorites"
Else
   SidePanel "Hide", "Favorites"
End If

End Select
End Sub
Private Sub TreeView1_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyF5 Then
'         TreeView1.Nodes.Clear
         TreeView1.Clear
         TreeView1.Refresh
         
         'retrieve the special folder path
         'to the internet favorites
         favpath = GetFolderPath(CSIDL_FAVORITES)
         
         'Initializes the Root Item in the TreeView
         Call LoadTreeView("Internet Favorites", True, True)
        
         If Len(favpath) > 0 Then
        
          'set up the search UDT
           With FP
              .sFileRoot = favpath
              .sFileNameExt = "*.url"
              .bRecurse = True
           End With
           
          'get the files
           Call SearchForFilesArray(FP)
'           TreeView1.Nodes("R").Expanded = True
           TreeView1.ItemExpanded("") = True
         Else
         
            MsgBox " Could not locate favorites folder! " & _
                "This program requires Microsoft's Internet " & _
                "Explorer to be installed. Program will shutdown now!", _
                vbCritical + vbOKOnly, "FavMenu Error"
            End
         End If
    End If
End Sub
Private Sub tmranim_Timer()
Static y As Integer
Static up As Boolean
Static down As Boolean
'MsgBox y
If y = 0 Then
down = True
up = False
y = 3
Else
    If y >= 378 Then
    up = True
    down = False
    ElseIf y = 3 Then
    up = False
    down = True
    End If
    
If down = True Then
y = y + 25
ElseIf up = True Then
y = y - 25
End If

End If
BitBlt picAnim.hdc, 3, 4, picAnim.Width, picAnim.Height, Picanimsrc.hdc, 0, y, vbSrcCopy
End Sub

Private Sub TreeView1_ItemClick(hItem As Long, RightButton As Boolean)
Dim UrlFav$
If Len(TreeView1.ItemKey(hItem)) > 0 Then
If InStr(TreeView1.ItemKey(hItem), "Folder") = 0 Then
UrlFav$ = TreeView1.ItemKey(hItem)
UrlFav$ = Left$(UrlFav$, InStr(UrlFav$, " ") - 1)
IE(Tabmain.SelectedTab).Navigate UrlFav$
cboAddress.Text = TreeView1.ItemKey(hItem)
End If
End If
End Sub

Private Sub TreeView1_ItemDblClick(hItem As Long)
'MsgBox TreeView1.ItemKey(hItem)
End Sub
Public Sub pBuildMenu()
Set m_cMenu = New cPopupMenu

'The menu settings
With frmMain.m_cMenu
.ImageList = m_cILMenu.hIml
.hWndOwner = Me.hwnd
.GradientHighlight = False
'----------------------------

' File menu:
iP(0) = .AddItem("&File", , , , , , , "mnuFileTOP")
iP(1) = .AddItem("&New Tab", , , iP(0), 1, , , "NewTab")
iP(1) = .AddItem("&Delete Tab", , , iP(0), 2, , , "DeleteTab")
iP(1) = .AddItem("-", , , iP(0), , , , "line2")
iP(1) = .AddItem("&Open" & vbTab & "Ctrl+O", , , iP(0), 19, , , "Open")
iP(1) = .AddItem("&Save As..", , , iP(0), , , , "SaveAs")
iP(1) = .AddItem("-", , , iP(0), , , , "line1")
iP(1) = .AddItem("Create &Shortcut", , , iP(0), , , , "CreateShortcut")
iP(1) = .AddItem("P&roperties", , , iP(0), 3, , , "Properties")
iP(1) = .AddItem("-", , , iP(0), , , , "line2")
iP(1) = .AddItem("&History", , , iP(0), , , , "History")
iP(2) = .AddItem("<empty>", , , iP(1), , , , "mnuHist(0)")
iP(1) = .AddItem("-", , , iP(0), , , , "line3")
'      iP(1) = .AddItem("&Work Offline", , , iP(0), , , , "WorkOffline")
iP(1) = .AddItem("&Close", , , iP(0), , , , "Close")

' Edit menu
iP(0) = .AddItem("&Edit", , , , , , , "mnuEditTOP")
iP(1) = .AddItem("-", , , iP(0), , , , "line1")
iP(1) = .AddItem("&Copy" & vbTab & "Ctrl+C", , , iP(0), 6, , , "Copy")
iP(1) = .AddItem("&Paste" & vbTab & "Ctrl+V", , , iP(0), 7, , , "Paste")
iP(1) = .AddItem("-", , , iP(0), , , , "line2")
iP(1) = .AddItem("Select &All" & vbTab & "Ctrl+A", , , iP(0), , , , "SelectAll")

' View menu
iP(0) = .AddItem("&View", , , , , , , "mnuViewTOP")
iP(1) = .AddItem("&Toolbars", , , iP(0), , , , "mnuView(0)")
iP(2) = .AddItem("&Standard Buttons", , , iP(1), , True, , "mnuToolbar(0)")
iP(2) = .AddItem("&Address", , , iP(1), , True, , "mnuToolbar(1)")
iP(2) = .AddItem("&Links", , , iP(1), 8, True, , "mnuToolbar(2)")
iP(2) = .AddItem("-", , , iP(1))
iP(2) = .AddItem("&Background Bitmap", , , iP(1), , , , "mnuToolbar(5)")
iP(1) = .AddItem("&Status Bar", , , iP(0), , , , "VStatusBar")
iP(2) = .AddItem("&Search", , , iP(1), 9, True, , "VSearch")
iP(2) = .AddItem("&Favorites", , , iP(1), 10, , , "VFavorites")
iP(2) = .AddItem("&History", , , iP(1), 11, , , "VHistory")
iP(2) = .AddItem("F&olders", , , iP(1), 12, , , "VFolders")
iP(2) = .AddItem("-", , , iP(1), , , , "line1")
iP(2) = .AddItem("&Tip of the Day", , , iP(1), 13, , , "VTipOfDay")
iP(1) = .AddItem("-", , , iP(0), , , , "line1")
iP(1) = .AddItem("&Stop" & vbTab & "Esc", , , iP(0), , , , "Stop")
iP(1) = .AddItem("&Refresh" & vbTab & "F5", , , iP(0), 18, , , "Refresh")

'Favorites Menu
iP(0) = .AddItem("&Favorites", , , , , , , "Favorites")
iP(1) = .AddItem("&Add To Favorites", , , iP(0), , , , "AddFavorites")
iP(1) = .AddItem("&Organize Favorites", , , iP(0), , , , "OrganizeFavorites")
iP(1) = .AddItem("-", , , iP(0), , , , "line1")

'Setup the favorites data
favpath = GetFolderPath(CSIDL_FAVORITES)
frmMain.TreeView1.Clear
frmMain.TreeView1.Refresh
Call LoadTreeView("Internet Favorites", True, True)
If Len(favpath) > 0 Then
With FP
.sFileRoot = favpath
.sFileNameExt = "*.url"
.bRecurse = True
End With
frmMain.TreeView1.ItemExpanded(a) = True
End If
Call SearchForFilesArray(FP)
'-------------------------------------

' Help menu.
iP(0) = .AddItem("&Help", , , , , , , "mnuHelpTOP")
iP(1) = .AddItem("&Contents", , , iP(0), 29, , , "Contents")
iP(1) = .AddItem("&Tip of the Day", , , iP(0), , , , "TipOfDay")
iP(1) = .AddItem("Elucid Software on the &Web", , , iP(0), 33, , , "ElucidSoftware")
iP(1) = .AddItem("&Send Feedback", , , iP(0), 23, , , "Feedback")
iP(1) = .AddItem("-", , , iP(0), , , , "line2")
iP(1) = .AddItem("&About...", , , iP(0), , , , "About")
End With
End Sub
