VERSION 5.00
Begin VB.UserControl vbalStatusBar 
   Alignable       =   -1  'True
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4788
   ControlContainer=   -1  'True
   ScaleHeight     =   3600
   ScaleWidth      =   4788
End
Attribute VB_Name = "vbalStatusBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' =========================================================================
' vbAccelerator Statusbar control
' Copyright © 1998 Steve McMahon (steve@dogma.demon.co.uk)
'
' This is a status bar control implemented in VB using COMCTL32.DLL
' Features include
'  * Status bar icons
'  * Panels resize right up to end of the bar!
'  * Owner draw status bar panel style allows you to draw your own
'    panel styles.
'  * Support for standard VB status panel types
'
' Visit vbAccelerator at http://vbaccelerator.com
' =========================================================================

' ==============================================================================
' Declares, constants and types required for status bar:
' ==============================================================================

' Win API declares:
Private Type POINTAPI
   x As Long
   y As Long
End Type
Private Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type
Private Type DRAWITEMSTRUCT
   CtlType As Long
   CtlID As Long
   itemID As Long
   itemAction As Long
   itemState As Long
   hwndItem As Long
   hdc As Long
   rcItem As RECT
   itemData As Long
End Type
Private Const WM_USER = &H400
Private Const WM_DRAWITEM = &H2B
Private Const WM_NOTIFY = &H4E
Private Const WM_WININICHANGE = &H1A
Private Const WM_SIZE = &H5
Private Const WM_SETFONT = &H30
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Const GWL_STYLE = (-16)
Private Const WS_CHILD = &H40000000
Private Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Const SW_HIDE = 0
Private Const SW_SHOW = 5
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SendMessageString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, lColorRef As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSrc As Any, ByVal ByteLen As Long)
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function EnableWindow Lib "user32" (ByVal hwnd As Long, ByVal fEnable As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Const DT_CALCRECT = &H400
Private Const DT_CENTER = &H1
Private Const DT_VCENTER = &H4
Private Const DT_SINGLELINE = &H20
Private Const DT_RIGHT = &H2
Private Const DT_BOTTOM = &H8
Private Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long
Private Declare Function GetKeyboardState Lib "user32" (pbKeyState As Byte) As Long
Private Declare Function DrawStateString Lib "user32" Alias "DrawStateA" (ByVal hdc As Long, ByVal hBrush As Long, ByVal lpDrawStateProc As Long, ByVal lpString As String, ByVal cbStringLen As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal fuFlags As Long) As Long
'/* Image type */
Private Const DST_COMPLEX = &H0
Private Const DST_TEXT = &H1
Private Const DST_PREFIXTEXT = &H2
Private Const DST_ICON = &H3
Private Const DST_BITMAP = &H4
' /* State type */
Private Const DSS_NORMAL = &H0
Private Const DSS_UNION = &H10 ' Dither
Private Const DSS_DISABLED = &H20
Private Const DSS_MONO = &H80 ' Draw in colour of brush specified in hBrush
Private Const DSS_RIGHT = &H8000
Private Const TRANSPARENT = 1
Private Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Private Declare Function InvalidateRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT, ByVal bErase As Long) As Long
Private Declare Function UpdateWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function PtInRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long

' Font:
Private Const LF_FACESIZE = 32
Private Type LOGFONT
    lfHeight As Long
    lfWidth As Long
    lfEscapement As Long
    lfOrientation As Long
    lfWeight As Long
    lfItalic As Byte
    lfUnderline As Byte
    lfStrikeOut As Byte
    lfCharSet As Byte
    lfOutPrecision As Byte
    lfClipPrecision As Byte
    lfQuality As Byte
    lfPitchAndFamily As Byte
    lfFaceName(LF_FACESIZE) As Byte
End Type
Private Const FW_NORMAL = 400
Private Const FW_BOLD = 700
Private Const FF_DONTCARE = 0
Private Const DEFAULT_QUALITY = 0
Private Const DEFAULT_PITCH = 0
Private Const DEFAULT_CHARSET = 1
Private Declare Function CreateFontIndirect& Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT)
Private Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Const LOGPIXELSY = 90        '  Logical pixels/inch in Y

' Common Controls declares:
Private Declare Sub InitCommonControls Lib "Comctl32.dll" ()
Private Const CCM_FIRST = &H2000                   '// Common control shared messages
Private Const CCM_SETBKCOLOR = (CCM_FIRST + 1)         '// lParam is bkColor
Private Const H_MAX As Long = &HFFFF + 1
Private Const SBN_FIRST = -880&                  '// status bar
Private Type NMHDR
   hwndFrom As Long
   idfrom As Long
   code As Long
End Type
Private Const NM_FIRST = H_MAX
Private Const NM_CLICK = (NM_FIRST - 2)                '// uses NMCLICK struct
Private Const NM_DBLCLK = (NM_FIRST - 3)
Private Const NM_RCLICK = (NM_FIRST - 5)               '// uses NMCLICK struct
Private Const NM_RDBLCLK = (NM_FIRST - 6)
Private Declare Function ImageList_GetIcon Lib "COMCTL32" (ByVal hImageList As Long, ByVal ImgIndex As Long, ByVal fuFlags As Long) As Long
Private Declare Function ImageList_GetIconSize Lib "COMCTL32" (ByVal hImageList As Long, cx As Long, cy As Long) As Long

'//====== STATUS BAR CONTROL ===================================================
Private Const SBARS_SIZEGRIP = &H100

Private Declare Function DrawStatusText Lib "COMCTL32" Alias "DrawStatusTextA" (ByVal hdc As Long, lprc As RECT, ByVal pszText As String, ByVal uFlags As Long) As Long
Private Declare Function CreateStatusWindow Lib "COMCTL32" Alias "CreateStatusWindowA" (ByVal style As Long, ByVal lpszText As String, ByVal hWndParent As Long, ByVal wID As Long) As Long

Private Const STATUSCLASSNAMEA = "msctls_statusbar32"
Private Const STATUSCLASSNAME = STATUSCLASSNAMEA

Private Const SB_SETTEXTA = (WM_USER + 1)
'Private Const SB_SETTEXTW = (WM_USER + 11)
Private Const SB_GETTEXTA = (WM_USER + 2)
'Private Const SB_GETTEXTW = (WM_USER + 13)
Private Const SB_GETTEXTLENGTHA = (WM_USER + 3)
'Private Const SB_GETTEXTLENGTHW = (WM_USER + 12)
Private Const SB_SETTIPTEXTA = (WM_USER + 16)
'Private Const SB_SETTIPTEXTW = (WM_USER + 17)
Private Const SB_GETTIPTEXTA = (WM_USER + 18)
'Private Const SB_GETTIPTEXTW = (WM_USER + 19)
Private Const SB_GETTEXT = SB_GETTEXTA
Private Const SB_SETTEXT = SB_SETTEXTA
Private Const SB_GETTEXTLENGTH = SB_GETTEXTLENGTHA
Private Const SB_SETTIPTEXT = SB_SETTIPTEXTA
Private Const SB_GETTIPTEXT = SB_GETTIPTEXTA

Private Const SB_SETPARTS = (WM_USER + 4)
Private Const SB_GETPARTS = (WM_USER + 6)
Private Const SB_GETBORDERS = (WM_USER + 7)
Private Const SB_SETMINHEIGHT = (WM_USER + 8)
Private Const SB_SIMPLE = (WM_USER + 9)
Private Const SB_GETRECT = (WM_USER + 10)
Private Const SB_ISSIMPLE = (WM_USER + 14)
Private Const SB_SETICON = (WM_USER + 15)
Private Const SB_GETICON = (WM_USER + 20)
'private const SB_SETUNICODEFORMAT     =CCM_SETUNICODEFORMAT
'private const SB_GETUNICODEFORMAT     =CCM_GETUNICODEFORMAT

Private Const SBT_OWNERDRAW = &H1000
Private Const SBT_NOBORDERS = &H100
Private Const SBT_POPOUT = &H200
Private Const SBT_RTLREADING = &H400
Private Const SBT_TOOLTIPS = &H800

Private Const SB_SETBKCOLOR = CCM_SETBKCOLOR               '// lParam = bkColor

'/// status bar notifications
Private Const SBN_SIMPLEMODECHANGE = (SBN_FIRST - 0)

'//====== STATUS BAR CONTROL ===================================================

Implements ISubclass

Public Enum ESTBRSimplePanelStyle
   estbrsStandard = &H0&
   estbrsNoBorders = SBT_NOBORDERS
   estbrsRaisedBorder = SBT_POPOUT
   estbrsTooltips = SBT_TOOLTIPS
End Enum

Public Enum ESTBRPanelStyle
   estbrStandard = &H0&
   estbrNoBorders = SBT_NOBORDERS
   estbrRaisedBorder = SBT_POPOUT
   estbrTooltips = SBT_TOOLTIPS
   estbrOwnerDraw = SBT_OWNERDRAW
   estbrCaps = SBT_OWNERDRAW + 1
   estbrNum = SBT_OWNERDRAW + 2
   estbrIns = SBT_OWNERDRAW + 3
   estbrScrl = SBT_OWNERDRAW + 4
   estbrTime = SBT_OWNERDRAW + 5
   estbrDate = SBT_OWNERDRAW + 6
   estbrDateTime = SBT_OWNERDRAW + 7
End Enum

Private Type tStatusPanel
   lID As Long
   sKey As String
   lItemData As Long
   iImgIndex As Long
   hIcon As Long ' 4.71+ only
   sText As String
   sToolTipText As String ' 4.71+ only
   lMinWidth As Long
   lIdealWidth As Long
   lSetWidth As Long
   bSpring As Boolean
   bFit As Boolean
   eStyle As ESTBRPanelStyle
   bState As Boolean
End Type

Private m_tPanels() As tStatusPanel
Private m_iPanelCount As Long
Private m_hWnd As Long
Private m_bSizeGrip As Boolean
Private m_bSimpleMode As Boolean
Private m_sSimpleText As String
Private m_eSimpleStyle As ESTBRPanelStyle
Private m_bSubClassing As Boolean
Private m_hIml As Long
Private m_hUFnt As Long
Private m_lIconSize As Long
Private WithEvents m_tmr As CTimer
Attribute m_tmr.VB_VarHelpID = -1

Public Event Click(ByVal iPanel As Long, ByVal x As Single, ByVal y As Single, ByVal eButton As MouseButtonConstants)
Attribute Click.VB_Description = "Raised when a panel in the status bar is clicked."
Public Event DblClick(ByVal iPanel As Long, ByVal x As Single, ByVal y As Single, ByVal eButton As MouseButtonConstants)
Attribute DblClick.VB_Description = "Raised when a panel in the status bar is double clicked."
Public Event DrawItem(ByVal lHDC As Long, ByVal iPanel As Long, ByVal lLeftPixels As Long, ByVal lTopPixels As Long, ByVal lRightPixels As Long, ByVal lBottomPixels As Long)
Attribute DrawItem.VB_Description = "Raised when a panel with the Owner-Draw style set needs to be redrawn."
Public Event Timer()
Attribute Timer.VB_Description = "Raised when an the internal status bar timer fires.  The internal status bar timer only operates to check key states and/or draw dates or times, so if your control does not use these custom styles then this event will not be raised."
Public Event Resize()
Attribute Resize.VB_Description = "Raised when the control is resized or one of the panels is resized."

Public Property Get SimpleMode() As Boolean
Attribute SimpleMode.VB_Description = "Gets/sets whether the status bar works in simple mode (a single panel only) or normal mode,"
   SimpleMode = m_bSimpleMode
End Property
Public Property Let SimpleMode(ByVal bState As Boolean)
Dim tR As RECT
   m_bSimpleMode = bState
   If (m_hWnd <> 0) Then
      SendMessageLong m_hWnd, SB_SIMPLE, Abs(bState), 0
   End If
   PropertyChanged "SimpleMode"
End Property

Public Property Get PanelKey(ByVal lIndex As Long) As Variant
Attribute PanelKey.VB_Description = "Gets/sets the key used to identify a panel."
Dim iPanel As Long
   If (lIndex > 0) And (lIndex <= m_iPanelCount) Then
      PanelKey = m_tPanels(lIndex).sKey
   Else
      Err.Raise vbObjectError + 1050, App.EXEName & ".vbalStatusBar", "Invalid Panel Index: " & lIndex
   End If
   
End Property
Public Property Let PanelKey(ByVal lIndex As Long, ByVal vKey As Variant)
   If (lIndex > 0) And (lIndex <= m_iPanelCount) Then
      m_tPanels(lIndex).sKey = vKey
   Else
      Err.Raise vbObjectError + 1050, App.EXEName & ".vbalStatusBar", "Invalid Panel Index: " & lIndex
   End If
   
End Property
Public Property Get PanelIndex(ByVal vKey As Variant) As Long
Attribute PanelIndex.VB_Description = "Returns the index of a panel given the panel's key."
Dim i As Long
Dim iFound As Long

   If (IsNumeric(vKey)) Then
      If (vKey > 0) And (vKey <= m_iPanelCount) Then
         PanelIndex = vKey
      Else
         Err.Raise vbObjectError + 1050, App.EXEName & ".vbalStatusBar", "Invalid Panel Index: " & vKey
      End If
   Else
      For i = 1 To m_iPanelCount
         If m_tPanels(i).sKey = vKey Then
            iFound = i
            Exit For
         End If
      Next i
      If (iFound > 0) Then
         PanelIndex = iFound
      Else
         Err.Raise vbObjectError + 1050, App.EXEName & ".vbalStatusBar", "Invalid Panel Index: " & vKey
      End If
   End If
   
End Property
Public Property Let PanelText(ByVal vKey As Variant, ByVal sText As String)
Attribute PanelText.VB_Description = "Gets/sets the text to show in a panel."
Dim iPanel As Long
Dim iPartuType As Long
Dim lR As Long
   iPanel = PanelIndex(vKey)
   If (iPanel > 0) Then
      m_tPanels(iPanel).sText = sText
      iPartuType = ((iPanel - 1) And &HFF&) Or (m_tPanels(iPanel).eStyle And &HFF00)
      If (m_tPanels(iPanel).eStyle And estbrOwnerDraw) <> estbrOwnerDraw Then
         If (Len(sText) > 0) Then
            lR = SendMessageString(m_hWnd, SB_SETTEXT, iPartuType, sText & Chr$(0))
         Else
            lR = SendMessageLong(m_hWnd, SB_SETTEXT, iPartuType, 0&)
         End If
         Debug.Assert (lR <> 0)
      Else
         SendMessageLong m_hWnd, SB_SETTEXT, iPartuType, m_tPanels(iPanel).lItemData
      End If
   End If
End Property
Public Property Get PanelText(ByVal vKey As Variant) As String
Dim iPanel As Long
   iPanel = PanelIndex(vKey)
   If (iPanel > 0) Then
      PanelText = m_tPanels(iPanel).sText
   End If
End Property
Public Property Let SimpleText(ByVal sText As String)
Attribute SimpleText.VB_Description = "Gets/sets the text displayed in the status bar when in simple mode.  Note this text is independent of any panels added to the control."
Dim iPartuType As Long
Dim lR As Long
   m_sSimpleText = sText
   If (m_hWnd <> 0) Then
      iPartuType = &HFF Or m_eSimpleStyle
      lR = SendMessageString(m_hWnd, SB_SETTEXT, iPartuType, m_sSimpleText & Chr$(0))
   End If
   PropertyChanged "SimpleText"
End Property
Public Property Get SimpleText() As String
   SimpleText = m_sSimpleText
End Property
Public Property Get SimpleStyle() As ESTBRSimplePanelStyle
Attribute SimpleStyle.VB_Description = "Gets/sets the style used to draw the status bar when it is in Simple Mode."
   SimpleStyle = m_eSimpleStyle
End Property
Public Property Let SimpleStyle(ByVal eStyle As ESTBRSimplePanelStyle)
   m_eSimpleStyle = eStyle
   PropertyChanged "SimpleStyle"
End Property
Public Property Let PanelToolTipText(ByVal vKey As Variant, ByVal sText As String)
Attribute PanelToolTipText.VB_Description = "Gets/sets the tool tip text to show in a panel.  Note that tool tips are only displayed for panels that have an icon and either have no text or not all text is visible in the panel."
Dim iPanel As Long
Dim lR As Long
   iPanel = PanelIndex(vKey)
   If (iPanel > 0) Then
      m_tPanels(iPanel).sToolTipText = sText
      lR = SendMessageString(m_hWnd, SB_SETTIPTEXT, iPanel - 1, sText & Chr$(0))
   End If
End Property
Public Property Get PanelToolTipText(ByVal vKey As Variant) As String
Dim iPanel As Long
Dim sTest As String
   iPanel = PanelIndex(vKey)
   If (iPanel > 0) Then
      PanelToolTipText = m_tPanels(iPanel).sToolTipText
   End If
End Property
Public Property Let PanelSpring(ByVal vKey As Variant, ByVal bState As Boolean)
Attribute PanelSpring.VB_Description = "Gets/sets whether a panel springs to fit the available space.  Only one panel at a time can have the spring property set to true."
Dim iPanel As Long
Dim i As Long
   iPanel = PanelIndex(vKey)
   If (iPanel > 0) Then
      If (m_tPanels(iPanel).bSpring <> bState) Then
         For i = 1 To m_iPanelCount
            If i = iPanel Then
               m_tPanels(iPanel).bSpring = bState
            Else
               m_tPanels(iPanel).bSpring = False
            End If
         Next i
         pEvaluateIdealSize iPanel
         pResizeStatus
      End If
   End If
End Property
Public Property Get PanelSpring(ByVal vKey As Variant) As Boolean
Dim iPanel As Long
   iPanel = PanelIndex(vKey)
   If (iPanel > 0) Then
      PanelSpring = m_tPanels(iPanel).bSpring
   End If
End Property
Public Property Let PanelFitToContents(ByVal vKey As Variant, ByVal bState As Boolean)
Attribute PanelFitToContents.VB_Description = "Gets/sets whether a panel should automatically size to its contents."
Dim iPanel As Long
   iPanel = PanelIndex(vKey)
   If (iPanel > 0) Then
      If (m_tPanels(iPanel).bFit <> bState) Then
         m_tPanels(iPanel).bFit = bState
         pEvaluateIdealSize iPanel
         pResizeStatus
      End If
   End If
End Property
Public Property Get PanelFitToContents(ByVal vKey As Variant) As Boolean
Dim iPanel As Long
   iPanel = PanelIndex(vKey)
   If (iPanel > 0) Then
      PanelFitToContents = m_tPanels(iPanel).bFit
   End If
End Property
Public Property Get PanelIcon(ByVal vKey As Variant) As Long
Attribute PanelIcon.VB_Description = "Gets/sets the 0 based index of an image in the associated image list to display in a panel."
Dim iPanel As Long
   iPanel = PanelIndex(vKey)
   If (iPanel > 0) Then
      PanelIcon = m_tPanels(iPanel).iImgIndex
   End If
End Property
Public Property Get PanelhIcon(ByVal vKey As Variant) As Long
Attribute PanelhIcon.VB_Description = "Gets/sets an icon handle to draw in a panel.  If reading this property, do not call DestroyIcon the returned hIcon - it is a handle to the actual icon used by the control, not a copy.  If setting this property, the hIcon will be automatically destroyed w"
Dim iPanel As Long
   iPanel = PanelIndex(vKey)
   If (iPanel > 0) Then
      ' Returns a hIcon if any:
      PanelhIcon = m_tPanels(iPanel).hIcon
   End If
End Property
Public Property Let PanelIcon(ByVal vKey As Variant, ByVal iImgIndex As Long)
Dim iPanel As Long
   iPanel = PanelIndex(vKey)
   If (iPanel > 0) Then
      If (m_tPanels(iPanel).hIcon <> 0) Then
         DestroyIcon m_tPanels(iPanel).hIcon
      End If
      m_tPanels(iPanel).hIcon = 0
      m_tPanels(iPanel).iImgIndex = iImgIndex
      If (iImgIndex > -1) Then
         ' extract a copy of the icon and add to sbar:
         m_tPanels(iPanel).hIcon = ImageList_GetIcon(m_hIml, iImgIndex, 0)
      End If
      SendMessageLong m_hWnd, SB_SETICON, iPanel - 1, m_tPanels(iPanel).hIcon
      pEvaluateIdealSize iPanel, iPanel
      pResizeStatus
   End If
End Property
Public Property Let PanelhIcon(ByVal vKey As Variant, ByVal hIcon As Long)
Dim iPanel As Long
   iPanel = PanelIndex(vKey)
   If (iPanel > 0) Then
      ' Destroy existing hIcon:
      If (m_tPanels(iPanel).hIcon <> 0) Then
         DestroyIcon m_tPanels(iPanel).hIcon
      End If
      m_tPanels(iPanel).hIcon = hIcon
      SendMessageLong m_hWnd, SB_SETICON, iPanel - 1, m_tPanels(iPanel).hIcon
      pEvaluateIdealSize iPanel, iPanel
      pResizeStatus
   End If
End Property
Public Property Let PanelStyle(ByVal vKey As Variant, ByVal eStyle As ESTBRPanelStyle)
Attribute PanelStyle.VB_Description = "Gets/sets the style used to draw a panel."
Dim iPanel As Long
   iPanel = PanelIndex(vKey)
   If (iPanel > 0) Then
      iPanel = iPanel - 1
      If (eStyle <> estbrOwnerDraw) Then
         SendMessageString m_hWnd, SB_SETTEXT, ((iPanel And &HFF) Or (eStyle And &HFF00)), m_tPanels(iPanel + 1).sText
      Else
         SendMessageLong m_hWnd, SB_SETTEXT, ((iPanel And &HFF) Or (eStyle And &HFF00)), m_tPanels(iPanel + 1).lItemData
      End If
   End If
End Property
Public Property Get PanelStyle(ByVal vKey As Variant) As ESTBRPanelStyle
Dim iPanel As Long
   iPanel = PanelIndex(vKey)
   If (iPanel > 0) Then
      PanelStyle = m_tPanels(iPanel).eStyle
   End If
End Property
Public Property Get PanelMinWidth(ByVal vKey As Variant) As Long
Attribute PanelMinWidth.VB_Description = "Gets/sets the minimum allowable width for a panel.  Note that icons and the sizing grip add additional width to the minimum."
Dim iPanel As Long
   iPanel = PanelIndex(vKey)
   If (iPanel > 0) Then
      PanelMinWidth = m_tPanels(iPanel).lMinWidth
   End If
End Property
Public Property Get PanelIdealWidth(ByVal vKey As Variant) As Long
Attribute PanelIdealWidth.VB_Description = "Gets/sets the width of the panel calculated for auto-sizing panels."
Dim iPanel As Long
   iPanel = PanelIndex(vKey)
   If (iPanel > 0) Then
      PanelIdealWidth = m_tPanels(iPanel).lIdealWidth
   End If
End Property
Public Property Let PanelIdealWidth(ByVal vKey As Variant, ByVal lWidth As Long)
Dim iPanel As Long
   iPanel = PanelIndex(vKey)
   If (iPanel > 0) Then
      m_tPanels(iPanel).lIdealWidth = lWidth
      pResizeStatus
   End If
End Property

Public Property Get PanelCount() As Long
Attribute PanelCount.VB_Description = "Gets the number of panels in the status bar."
   PanelCount = m_iPanelCount
End Property
Public Sub GetPanelRect( _
      ByVal vKey As Variant, _
      Optional ByRef iLeftPixels As Long, _
      Optional ByRef iTopPixels As Long, _
      Optional ByRef iRightPixels As Long, _
      Optional ByRef iBottomPixels As Long _
   )
Attribute GetPanelRect.VB_Description = "Gets the outside bounding rectangle of a panel in the control."
Dim iPanel As Long
Dim tR As RECT
   iPanel = PanelIndex(vKey)
   If (iPanel > 0) Then
      SendMessage m_hWnd, SB_GETRECT, iPanel - 1, tR
      iLeftPixels = tR.Left
      iTopPixels = tR.Top
      iRightPixels = tR.Right
      iBottomPixels = tR.Bottom
   End If
End Sub

Property Get Font() As StdFont
Attribute Font.VB_Description = "Gets/sets the font used to draw the status bar panels."
   ' Get the control's default font:
   Set Font = UserControl.Font
End Property
Property Set Font(fntThis As StdFont)
Dim hUFnt As Long
Dim tULF As LOGFONT
   ' Set the control's default font:
   Set UserControl.Font = fntThis
   ' Store a log font structure for this font:
   pOLEFontToLogFont fntThis, UserControl.hdc, tULF
   ' Store old font handle:
   hUFnt = m_hUFnt
   ' Create a new version of the font:
   m_hUFnt = CreateFontIndirect(tULF)
   ' Ensure the edit portion has the correct font:
   If (m_hWnd <> 0) Then
      SendMessage m_hWnd, WM_SETFONT, m_hUFnt, 1
      pEvaluateIdealSize 1, m_iPanelCount
      pResizeStatus
   End If
   ' Delete previous version, if we had one:
   If (hUFnt <> 0) Then
       DeleteObject hUFnt
   End If
   PropertyChanged "Font"
    
End Property
Private Sub pOLEFontToLogFont(fntThis As StdFont, hdc As Long, tLF As LOGFONT)
Dim sFont As String
Dim iChar As Integer
Dim b() As Byte

   ' Convert an OLE StdFont to a LOGFONT structure:
   With tLF
       sFont = fntThis.Name
       ' There is a quicker way involving StrConv and CopyMemory, but
       ' this is simpler!:
       b = StrConv(sFont, vbFromUnicode)
       For iChar = 1 To Len(sFont)
           .lfFaceName(iChar - 1) = b(iChar - 1)
       Next iChar
       ' Based on the Win32SDK documentation:
       .lfHeight = -MulDiv((fntThis.Size), (GetDeviceCaps(hdc, LOGPIXELSY)), 72)
       .lfItalic = fntThis.Italic
       If (fntThis.Bold) Then
           .lfWeight = FW_BOLD
       Else
           .lfWeight = FW_NORMAL
       End If
       .lfUnderline = fntThis.Underline
       .lfStrikeOut = fntThis.Strikethrough
       .lfCharSet = fntThis.Charset
   End With

End Sub
Private Sub pEvaluateIdealSize( _
      ByVal iStartPanel As Long, _
      Optional ByVal iEndPanel As Long = -1 _
   )
Dim i As Long
Dim tR As RECT
Dim lHDC As Long

   If (m_iPanelCount > 0) Then
      If (iEndPanel < iStartPanel) Then
         iEndPanel = iStartPanel
      End If
      lHDC = UserControl.hdc
      For i = iStartPanel To iEndPanel
         DrawText lHDC, m_tPanels(i).sText, Len(m_tPanels(i).sText), tR, DT_CALCRECT
         m_tPanels(i).lIdealWidth = tR.Right - tR.Left + 12
         If (m_tPanels(i).lIdealWidth < m_tPanels(i).lMinWidth) Then
            m_tPanels(i).lIdealWidth = m_tPanels(i).lMinWidth
         End If
      Next i
   End If
End Sub
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Gets/sets whether the control is enabled."
   Enabled = UserControl.Enabled
End Property
Public Property Let Enabled(ByVal bState As Boolean)
   UserControl.Enabled = bState
   If (m_hWnd <> 0) Then
      EnableWindow m_hWnd, CLng(Abs(bState))
   End If
   PropertyChanged "Enabled"
End Property

Public Function AddPanel( _
      Optional ByVal eStyle As ESTBRPanelStyle = estbrStandard, _
      Optional ByVal sText As String = "", _
      Optional ByVal sToolTipText As String = "", _
      Optional ByVal iImgIndex As Long = -1, _
      Optional ByVal lMinWidth As Long = 64, _
      Optional ByVal bSpring As Boolean = False, _
      Optional ByVal bFitContents As Boolean = False, _
      Optional ByVal lItemData As Long = 0, _
      Optional ByVal sKey As String = "", _
      Optional ByVal vKeyBefore As Variant _
   ) As Long
Attribute AddPanel.VB_Description = "Adds or inserts a panel to the status bar control."
Dim iIndex As Long
Dim i As Long
Dim bEnabled As Boolean
Dim tR As RECT
   
   If (m_iPanelCount >= &HFF) Then
      Err.Raise vbObjectError + 1051, App.EXEName & ".vbalStatusBar", "Too many panels."
      Exit Function
   End If
   
   If (eStyle > estbrOwnerDraw) Then
      bFitContents = True
      pGetCustomItem eStyle, sText, bEnabled
      DrawText UserControl.hdc, sText, Len(sText), tR, DT_CALCRECT
      lMinWidth = tR.Right - tR.Left + 8
   End If
   
   If Not IsMissing(vKeyBefore) Then
      ' Determine if vKeyBefore is valid:
      iIndex = PanelIndex(vKeyBefore)
      If (iIndex > 0) Then
         ' ok. Insert a space:
         m_iPanelCount = m_iPanelCount + 1
         ReDim Preserve m_tPanels(1 To m_iPanelCount) As tStatusPanel
         For i = m_iPanelCount To iIndex + 1 Step -1
            LSet m_tPanels(i) = m_tPanels(i - 1)
         Next i
         m_tPanels(iIndex).hIcon = 0
      Else
         ' Failed
         Exit Function
      End If
   Else
      ' Insert a space at the end:
      m_iPanelCount = m_iPanelCount + 1
      ReDim Preserve m_tPanels(1 To m_iPanelCount) As tStatusPanel
      iIndex = m_iPanelCount
   End If
   
   ' Set up the info:
   If (bSpring) Then
      For i = 1 To m_iPanelCount
         If (i <> iIndex) Then
            m_tPanels(i).bSpring = False
         End If
      Next i
   End If
   
   With m_tPanels(iIndex)
      .bFit = bFitContents
      .bSpring = bSpring
      .eStyle = eStyle
      .iImgIndex = iImgIndex
      .lMinWidth = lMinWidth
      .lItemData = lItemData
      .sKey = sKey
      .sText = sText
      .sToolTipText = sToolTipText
   End With
   
   ' Add the information to the status bar:
   pEvaluateIdealSize iIndex
   pResizeStatus
   
   ' Now ensure the text, style, tooltip and icon are actually correct:
   PanelText(iIndex) = m_tPanels(iIndex).sText
   PanelToolTipText(iIndex) = m_tPanels(iIndex).sToolTipText
   PanelIcon(iIndex) = m_tPanels(iIndex).iImgIndex
   If (m_tPanels(iIndex).hIcon <> 0) Then
      ' Ensure size is correct taking account of icon:
      pEvaluateIdealSize iIndex
      pResizeStatus
   End If
   
   
   ' Check whether we need a timer:
   pCheckEnableTimer
   
End Function

Public Function RemovePanel( _
      ByVal vKey As Variant _
   )
Attribute RemovePanel.VB_Description = "Removes a panel from the control."
Dim iIndex As Long
Dim i As Long
   iIndex = PanelIndex(vKey)
   If (iIndex > 0) Then
      If (m_tPanels(iIndex).hIcon <> 0) Then
         DestroyIcon m_tPanels(iIndex).hIcon
      End If
      For i = iIndex To m_iPanelCount - 1
         LSet m_tPanels(i) = m_tPanels(i + 1)
      Next i
      m_iPanelCount = m_iPanelCount - 1
      If (m_iPanelCount > 0) Then
         ReDim Preserve m_tPanels(1 To m_iPanelCount) As tStatusPanel
      End If
      pResizeStatus
   End If
End Function

Private Sub pCheckEnableTimer()
Dim i As Long
Dim bTimer As Boolean
   For i = 1 To m_iPanelCount
      If (m_tPanels(i).eStyle > estbrOwnerDraw) Then
         bTimer = True
         Exit For
      End If
   Next i
   If (bTimer) Then
      If (m_tmr Is Nothing) Then
         Set m_tmr = New CTimer
         m_tmr.Interval = 250
      End If
   Else
      If Not (m_tmr Is Nothing) Then
         m_tmr.Interval = 0
         Set m_tmr = Nothing
      End If
   End If
End Sub

Public Property Let BackColor(ByVal oColor As OLE_COLOR)
Attribute BackColor.VB_Description = "Gets/sets the back colour of the status bar control."
Dim lColor As Long
   ' v4.71+ only
   UserControl.BackColor = oColor
   lColor = TranslateColor(oColor)
   If (m_hWnd <> 0) Then
      SendMessageLong m_hWnd, SB_SETBKCOLOR, 0, lColor
   End If
   
   PropertyChanged "BackColor"
End Property
Public Property Get BackColor() As OLE_COLOR
   ' v4.71+ only
   BackColor = UserControl.BackColor
End Property
Private Function TranslateColor(ByVal clr As OLE_COLOR, _
                        Optional hPal As Long = 0) As Long
   If OleTranslateColor(clr, hPal, TranslateColor) Then
      TranslateColor = -1
   End If
End Function



Public Property Let ImageList(vThis As Variant)
Attribute ImageList.VB_Description = "Associates an ImageList with the status bar. The ImageList can either be a COMCTL32.OCX image list, a vbAccelerator image list or a long hImageList handle to an image list created using the COMCTL32.DLL API."
Dim cy As Long, lR As Long
    
    ' Set the ImageList handle property either from a VB
    ' image list or directly:
    m_hIml = 0
    If TypeName(vThis) = "ImageList" Then
        ' VB ImageList control.  Note that unless
        ' some call has been made to an object within a
        ' VB ImageList the image list itself is not
        ' created.  Therefore hImageList returns error. So
        ' ensure that the ImageList has been initialised by
        ' drawing into nowhere:
        On Error Resume Next
        ' Get the image list initialised..
        vThis.ListImages(1).Draw 0, 0, 0, 1
        m_hIml = vThis.hImageList
        If (Err.Number <> 0) Then
            ' No images.
            m_hIml = 0
        Else
            ' Get the icon size:
            lR = ImageList_GetIconSize(m_hIml, m_lIconSize, cy)
        End If
        On Error GoTo 0
    ElseIf VarType(vThis) = vbLong Then
        ' Assume ImageList handle:
        m_hIml = vThis
        ' Get the icon size:
        lR = ImageList_GetIconSize(m_hIml, m_lIconSize, cy)
    Else
        Err.Raise vbObjectError + 1049, App.EXEName & ".vbalStatusBar", "ImageList property expects ImageList object or long hImageList handle."
    End If
       
End Property

Public Sub RedrawPanel(ByVal vKey As Variant)
Attribute RedrawPanel.VB_Description = "Forces a panel to redraw."
Dim iPanel As Long
Dim tR As RECT
   iPanel = PanelIndex(vKey)
   If (iPanel > 0) Then
      SendMessage m_hWnd, SB_GETRECT, iPanel - 1, tR
      InvalidateRect m_hWnd, tR, 0
      UpdateWindow m_hWnd
   End If
End Sub

Private Sub pResizeStatus()
Dim tR As RECT
Dim i As Long
Dim iSpringIndex As Long
Dim lpParts() As Long
   
   If (m_iPanelCount > 0) Then
      
      ' Initiallly set to minimum widths:
      ReDim lpParts(0 To m_iPanelCount - 1) As Long
      If (m_tPanels(1).bFit) Then
         lpParts(0) = m_tPanels(1).lIdealWidth
      Else
         lpParts(0) = m_tPanels(1).lMinWidth
      End If
      If (m_tPanels(1).hIcon) Then
         lpParts(0) = lpParts(0) + m_lIconSize
      End If
      If (m_tPanels(1).bSpring) Then
         iSpringIndex = 1
      End If
      For i = 2 To m_iPanelCount
         If (m_tPanels(i).bFit) Then
            lpParts(i - 1) = lpParts(i - 2) + m_tPanels(i).lIdealWidth
         Else
            lpParts(i - 1) = lpParts(i - 2) + m_tPanels(i).lMinWidth
         End If
         If (m_tPanels(i).bSpring) Then
            iSpringIndex = i
         End If
         If (m_tPanels(i).hIcon <> 0) Then
            ' Add space for the icon:
            lpParts(i - 1) = lpParts(i - 1) + m_lIconSize
         End If
         If (i = m_iPanelCount) Then
            lpParts(i - 1) = lpParts(i - 1) + (UserControl.ScaleHeight * 3) \ (Screen.TwipsPerPixelY * 4)
         End If
      Next i
      
      ' Will all bars fit in at maximum size?
      GetClientRect m_hWnd, tR
      If (lpParts(m_iPanelCount - 1) > tR.Right) Then
         ' Draw all panels at min width
      Else
         ' Spring the spring panel to fit:
         If (iSpringIndex = 0) Then
            iSpringIndex = m_iPanelCount
         End If
         lpParts(iSpringIndex - 1) = lpParts(iSpringIndex - 1) + (tR.Right - lpParts(m_iPanelCount - 1))
         For i = iSpringIndex + 1 To m_iPanelCount
            If (m_tPanels(i).bFit) Then
               lpParts(i - 1) = lpParts(i - 2) + m_tPanels(i).lIdealWidth
            Else
               lpParts(i - 1) = lpParts(i - 2) + m_tPanels(i).lMinWidth
            End If
            If (m_tPanels(i).hIcon <> 0) Then
               ' Add space for the icon:
               lpParts(i - 1) = lpParts(i - 1) + m_lIconSize
            End If
            If (i = m_iPanelCount) Then
               lpParts(i - 1) = lpParts(i - 1) + (UserControl.ScaleHeight * 3) \ (Screen.TwipsPerPixelY * 4)
            End If
         Next i
      End If
      
      m_tPanels(1).lSetWidth = lpParts(0)
      For i = 2 To m_iPanelCount
         m_tPanels(i).lSetWidth = lpParts(i - 1) - lpParts(i - 2)
      Next i
      
      ' Set the sizes:
      SendMessage m_hWnd, SB_SETPARTS, m_iPanelCount, lpParts(0)
      
      RaiseEvent Resize
      
   End If
   
End Sub


Public Property Get SizeGrip() As Boolean
Attribute SizeGrip.VB_Description = "Gets/sets whether a sizing grip will be displayed at the right-hand bottom corner of the status bar."
   SizeGrip = m_bSizeGrip
End Property
Public Property Let SizeGrip(ByVal bSizeGrip As Boolean)
Dim lStyle As Long
   m_bSizeGrip = bSizeGrip
   If (m_hWnd <> 0) Then
      lStyle = GetWindowLong(m_hWnd, GWL_STYLE)
      If (bSizeGrip) Then
         lStyle = lStyle And SBARS_SIZEGRIP
      Else
         lStyle = lStyle And Not SBARS_SIZEGRIP
      End If
      SetWindowLong m_hWnd, GWL_STYLE, lStyle
   End If
End Property

Private Sub pInitialise()
   ' Ensure no status bar:
   pDestroy

   ' Create status bar:
   If (pbCreate()) Then
      ' Start subclassing:
      pAttachMessages
   End If
End Sub
Private Sub pDestroy()
   ' Clear up subclassing if any
   pDetachMessages
   ' Clear up status bar:
   pTerminate
End Sub
Private Sub pAttachMessages()
   ' If we have a status bar, start subclassing:
   If (m_hWnd <> 0) Then
      AttachMessage Me, UserControl.hwnd, WM_DRAWITEM
      AttachMessage Me, UserControl.hwnd, WM_NOTIFY
      AttachMessage Me, UserControl.hwnd, WM_WININICHANGE
      m_bSubClassing = True
   End If
End Sub
Private Sub pDetachMessages()
   ' If we have a status bar:
   If (m_hWnd <> 0) Then
      ' If we have started subclassing it:
      If (m_bSubClassing) Then
         ' Clear up messages:
         DetachMessage Me, UserControl.hwnd, WM_DRAWITEM
         DetachMessage Me, UserControl.hwnd, WM_NOTIFY
         DetachMessage Me, UserControl.hwnd, WM_WININICHANGE
      End If
      m_bSubClassing = False
   End If
End Sub

Private Function pbCreate() As Boolean
Dim lHwnd As Long
Dim lID As Long
Dim lStyle As Long
Dim szNull As String
Dim tR As RECT
   
   If (UserControl.Ambient.UserMode) Then
   
      ' Ensure common controls:
      InitCommonControls

      szNull = Chr$(0)
      lID = 0
      If (m_bSizeGrip) Then
         lStyle = SBARS_SIZEGRIP
      End If
      lStyle = lStyle Or WS_CHILD Or SBT_TOOLTIPS
      
      '// Create the status bar.
      lHwnd = CreateWindowEx( _
        0, _
        STATUSCLASSNAME, _
        "", _
        lStyle, _
        0, 0, 0, 0, _
        UserControl.hwnd, _
        lID, _
        App.hInstance, _
        ByVal 0&)
      'lhWnd = CreateStatusWindow(lStyle, szNull, UserControl.hwnd, lID)
      If (lHwnd <> 0) Then
         m_hWnd = lHwnd
         GetWindowRect lHwnd, tR
         UserControl.Height = (tR.Bottom - tR.Top) * Screen.TwipsPerPixelY
         MoveWindow m_hWnd, 0, 0, UserControl.ScaleWidth \ Screen.TwipsPerPixelX, UserControl.ScaleHeight \ Screen.TwipsPerPixelY, 1
         ShowWindow m_hWnd, SW_SHOW
         pbCreate = True
      End If
   End If
   
End Function

Private Sub pTerminate()
Dim i As Long
   
   ' Stop the timer if any:
   If Not (m_tmr Is Nothing) Then
      m_tmr.Interval = 0
      Set m_tmr = Nothing
   End If
   ' Destroy the status bar:
   If (m_hWnd <> 0) Then
      ShowWindow m_hWnd, SW_HIDE
      SetParent m_hWnd, 0
      DestroyWindow m_hWnd
      m_hWnd = 0
   End If
   ' Delete the font selected into the control
   ' (if we had one):
   If (m_hUFnt <> 0) Then
       DeleteObject m_hUFnt
   End If
   ' Delete any icons owned by the sbar:
   For i = 1 To m_iPanelCount
      If (m_tPanels(i).hIcon <> 0) Then
         DestroyIcon m_tPanels(i).hIcon
      End If
   Next i
   
End Sub

Private Property Let ISubclass_MsgResponse(ByVal RHS As SSubTimer.EMsgResponse)
   '
End Property

Private Property Get ISubclass_MsgResponse() As SSubTimer.EMsgResponse
   ISubclass_MsgResponse = emrPostProcess
End Property

Private Function ISubclass_WindowProc(ByVal hwnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Dim tDis As DRAWITEMSTRUCT
Dim tNMH As NMHDR
Dim eButton As MouseButtonConstants
Dim iPanel As Long, x As Single, y As Single
Dim eStyle As ESTBRPanelStyle

   Select Case iMsg
   Case WM_DRAWITEM
      CopyMemory tDis, ByVal lParam, Len(tDis)
      If tDis.hwndItem = m_hWnd Then
         eStyle = PanelStyle(tDis.itemID + 1)
         If (eStyle > estbrOwnerDraw) Then
            ' preset style:
            pDefaultDrawItem eStyle, tDis.hdc, tDis.itemID + 1, tDis.rcItem
         Else
            ' owner draw style:
            RaiseEvent DrawItem(tDis.hdc, tDis.itemID + 1, tDis.rcItem.Left, tDis.rcItem.Top, tDis.rcItem.Right, tDis.rcItem.Bottom)
         End If
      End If
      
   Case WM_NOTIFY
      CopyMemory tNMH, ByVal lParam, Len(tNMH)
      If (tNMH.hwndFrom = m_hWnd) Then
         Select Case tNMH.code
         Case NM_CLICK, NM_RCLICK
            If (tNMH.code = NM_CLICK) Then
               eButton = vbLeftButton
            Else
               eButton = vbRightButton
            End If
            pGetClickPosition iPanel, x, y
            RaiseEvent Click(iPanel, x, y, eButton)
         Case NM_DBLCLK, NM_RDBLCLK
            If (tNMH.code = NM_DBLCLK) Then
               eButton = vbLeftButton
            Else
               eButton = vbRightButton
            End If
            pGetClickPosition iPanel, x, y
            RaiseEvent DblClick(iPanel, x, y, eButton)
         End Select
      End If
   Case WM_WININICHANGE
   
   End Select
   
End Function
Private Sub pDefaultDrawItem( _
      ByVal eStyle As ESTBRPanelStyle, _
      ByVal lHDC As Long, _
      ByVal iPanel As Long, _
      ByRef tR As RECT _
   )
Dim bEnabled As Boolean
Dim sText As String
Dim lFlags As Long
Dim b(0 To 255) As Byte
Dim tTR As RECT

   pGetCustomItem eStyle, sText, bEnabled
   tR.Right = tR.Left + m_tPanels(iPanel).lSetWidth
   LSet tTR = tR
   DrawText lHDC, sText, Len(sText), tTR, DT_CALCRECT
   tR.Left = tR.Left + ((tR.Right - tR.Left - 4) - (tTR.Right - tTR.Left)) \ 2
   tR.Top = tR.Top + ((tR.Bottom - tR.Top) - (tTR.Bottom - tTR.Top)) - 2
   If Not (bEnabled) Then
      lFlags = DSS_DISABLED
   End If
   lFlags = lFlags Or DST_TEXT
   SetBkMode lHDC, TRANSPARENT
   DrawStateString lHDC, 0, 0, sText, Len(sText), tR.Left, tR.Top, tR.Right - tR.Left, tR.Bottom - tR.Top, lFlags
   
   m_tPanels(iPanel).sText = sText
   m_tPanels(iPanel).bState = bEnabled
   
End Sub
Private Sub pGetCustomItem( _
      ByVal eStyle As ESTBRPanelStyle, _
      ByRef sText As String, _
      ByRef bEnabled As Boolean _
   )
Dim b(0 To 255) As Byte
   
   bEnabled = True
   sText = ""
   Select Case eStyle
   Case estbrTime
      sText = Format$(Now, "short time")
   Case estbrScrl
      sText = "SCRL"
      GetKeyboardState b(0)
      bEnabled = (b(vbKeyScrollLock) <> 0)
   Case estbrNum
      sText = "NUM"
      GetKeyboardState b(0)
      bEnabled = (b(vbKeyNumlock) <> 0)
   Case estbrIns
      sText = "OVR"
      GetKeyboardState b(0)
      bEnabled = (b(vbKeyInsert) <> 0)
   Case estbrDateTime
      sText = Format$(Now, "medium date") & " " & Format$(Now, "short time")
   Case estbrDate
      sText = Format$(Now, "medium date")
   Case estbrCaps
      sText = "CAPS"
      GetKeyboardState b(0)
      bEnabled = (b(vbKeyCapital) <> 0)
   End Select
   
End Sub
Private Sub pGetClickPosition(ByRef iPanel As Long, ByRef x As Single, ByRef y As Single)
Dim tp As POINTAPI
Dim tR As RECT
Dim i As Long
   GetCursorPos tp
   ScreenToClient m_hWnd, tp
   ' Evaluate the panel:
   x = tp.x * Screen.TwipsPerPixelY
   y = tp.y * Screen.TwipsPerPixelY
   For i = 1 To m_iPanelCount
      SendMessage m_hWnd, SB_GETRECT, i - 1, tR
      If PtInRect(tR, tp.x, tp.y) Then
         iPanel = i
         Exit For
      End If
   Next i
End Sub

Private Sub m_tmr_ThatTime()
Dim i As Long
Dim bUpdate As Boolean
Dim tR As RECT
Dim sText As String
Dim bState As Boolean

   For i = 1 To m_iPanelCount
      If (m_tPanels(i).eStyle > estbrOwnerDraw) Then
         ' Update if required:
         pGetCustomItem m_tPanels(i).eStyle, sText, bState
         If (sText <> m_tPanels(i).sText) Or (bState <> m_tPanels(i).bState) Then
            SendMessage m_hWnd, SB_GETRECT, i - 1, tR
            InvalidateRect m_hWnd, tR, 0
            bUpdate = True
         End If
      End If
   Next i
   If (bUpdate) Then
      UpdateWindow m_hWnd
   End If
End Sub


Private Sub UserControl_Initialize()
   m_bSizeGrip = True
End Sub

Private Sub UserControl_InitProperties()
   pInitialise
   UserControl.Extender.Align = 2
End Sub

Private Sub UserControl_Paint()
Dim tR As RECT
   If Not (UserControl.Ambient.UserMode) Then
      GetClientRect UserControl.hwnd, tR
      DrawStatusText UserControl.hdc, tR, "vbAccelerator Status Bar", 0
   End If
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
Dim sFnt As New StdFont
   SizeGrip = PropBag.ReadProperty("SizeGrip", True)
   pInitialise
   sFnt.Name = "MS Sans Serif"
   sFnt.Size = 8
   Set Font = PropBag.ReadProperty("Font", sFnt)
   BackColor = PropBag.ReadProperty("BackColor", vbButtonFace)
   SimpleText = PropBag.ReadProperty("SimpleText", "")
   SimpleStyle = PropBag.ReadProperty("SimpleStyle", estbrsNoBorders)
   SimpleMode = PropBag.ReadProperty("SimpleMode", False)
End Sub

Private Sub UserControl_Resize()
Dim tR As RECT
Dim bInHere As Boolean
   If (UserControl.Ambient.UserMode) Then
      If Not (bInHere) Then
         bInHere = True
         ' Resize the status bar:
         SendMessageLong m_hWnd, WM_SIZE, 0, 0
         ' Is the UserControl the correct height?
         GetClientRect m_hWnd, tR
         If (UserControl.Height <> (tR.Bottom - tR.Top) * Screen.TwipsPerPixelY) Then
            UserControl.Height = (tR.Bottom - tR.Top) * Screen.TwipsPerPixelY
         End If
         ' Resize the panels:
         pResizeStatus
         bInHere = False
      End If
   End If
End Sub

Private Sub UserControl_Terminate()
   pTerminate
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
Dim sFnt As New StdFont
   PropBag.WriteProperty "SizeGrip", SizeGrip, True
   sFnt.Name = "MS Sans Serif"
   sFnt.Size = 8
   PropBag.WriteProperty "Font", Font, sFnt
   PropBag.WriteProperty "BackColor", BackColor
   PropBag.WriteProperty "SimpleText", SimpleText, ""
   PropBag.WriteProperty "SimpleStyle", SimpleStyle, estbrsNoBorders
   PropBag.WriteProperty "SimpleMode", SimpleMode, False
End Sub




