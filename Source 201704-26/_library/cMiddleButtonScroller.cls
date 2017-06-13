VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMiddleButtonScroller"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type POINTAPI
   x As Long
   y As Long
End Type

Private Type RECT
   left As Long
   top As Long
   right As Long
   bottom As Long
End Type

Private Type SCROLLINFO
    cbSize As Long
    fMask As Long
    nMin As Long
    nMax As Long
    nPage As Long
    nPos As Long
    nTrackPos As Long
End Type

Private Type SCROLLBARINFO
    cbSize As Long
    rcScrollBar As RECT
    dxyLineButton As Long
    xyThumbTop As Long
    xyThumbBottom As Long
    reserved As Long
    rgstate(0 To 5) As Long
End Type

Private Declare Function timeGetTime Lib "winmm.dll" () As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetScrollInfo Lib "user32" (ByVal hwnd As Long, ByVal n As Long, lpcScrollInfo As SCROLLINFO, ByVal BOOL As Boolean) As Long
Private Declare Function GetScrollInfo Lib "user32" (ByVal hwnd As Long, ByVal n As Long, LPSCROLLINFO As SCROLLINFO) As Long
Private Declare Function GetScrollBarInfo Lib "user32" (ByVal hwnd As Long, ByVal idObject As Long, psbi As SCROLLBARINFO) As Long
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long

'Private Const SW_SHOWNOACTIVATE = 4
Private Const GWL_STYLE = (-16)
Private Const WS_HSCROLL = &H100000
Private Const WS_VSCROLL = &H200000
Private Const WM_ACTIVATEAPP = &H1C
Private Const SIF_RANGE = &H1
Private Const SIF_PAGE = &H2
Private Const SIF_POS = &H4
Private Const SIF_DISABLENOSCROLL = &H8
Private Const SIF_TRACKPOS = &H10
Private Const SIF_ALL = (SIF_RANGE Or SIF_PAGE Or SIF_POS Or SIF_TRACKPOS)
Private Const SB_HORZ = 0
Private Const SB_VERT = 1
Private Const WM_VSCROLL = &H115
Private Const WM_HSCROLL = &H114
Private Const SB_THUMBPOSITION = 4
Private Const SB_THUMBTRACK = 5
Private Const OBJID_VSCROLL = &HFFFFFFFB
Private Const OBJID_HSCROLL = &HFFFFFFFA
Private Const STATE_SYSTEM_UNAVAILABLE = &H1&         '// Disabled

Private Enum EMiddleButtonScrollerEndMode
   eAnyButtonPress
   eMiddleButtonRelease
End Enum

Public Enum EMiddleButtonScrollerMode
   eNone
   ePixelBased
   eLineBased
End Enum

Private m_hWnd As Long
Private m_hWndParent As Long
Private m_tP As POINTAPI
Private m_bInMiddleScroll As Boolean
Private m_bInitial As Boolean
Private m_eMode As EMiddleButtonScrollerEndMode
Private m_bHasHorizontalScroll As Boolean
Private m_bHasVerticalScroll As Boolean
Private m_eHorizontalMode As EMiddleButtonScrollerMode
Private m_eVerticalMode As EMiddleButtonScrollerMode
Private m_lLastScreenCursor As Long
Private m_lLastVertTime As Long
Private m_lLastHorzTime As Long

Private WithEvents m_tmr As CTimer
Attribute m_tmr.VB_VarHelpID = -1

Private m_frmRefPoint As frmRefPoint

Implements IWindowsHook
Implements ISubclass

Public Property Get VerticalMode() As EMiddleButtonScrollerMode
   VerticalMode = m_eVerticalMode
End Property
Public Property Let VerticalMode(ByVal eMode As EMiddleButtonScrollerMode)
   m_eVerticalMode = eMode
End Property

Public Property Get HorizontalMode() As EMiddleButtonScrollerMode
   HorizontalMode = m_eHorizontalMode
End Property
Public Property Let HorizontalMode(ByVal eMode As EMiddleButtonScrollerMode)
   m_eHorizontalMode = eMode
End Property

Public Sub StartMiddleScroll(ByVal hwnd As Long)
   '
   EndMiddleScroll

   If (m_eHorizontalMode = eNone) And (m_eVerticalMode = eNone) Then
      ' nothing to do
      Exit Sub
   End If
      
   ' Determine which scroll bars we have:
   Dim lStyle As Long
   lStyle = GetWindowLong(hwnd, GWL_STYLE)
   m_bHasHorizontalScroll = ((lStyle And WS_HSCROLL) = WS_HSCROLL)
   m_bHasVerticalScroll = ((lStyle And WS_VSCROLL) = WS_VSCROLL)
   Dim tSBI As SCROLLBARINFO
   tSBI.cbSize = Len(tSBI)
   GetScrollBarInfo hwnd, OBJID_HSCROLL, tSBI
   If (tSBI.rgstate(0) And STATE_SYSTEM_UNAVAILABLE) = STATE_SYSTEM_UNAVAILABLE Then
      m_bHasHorizontalScroll = False
   End If
   GetScrollBarInfo hwnd, OBJID_VSCROLL, tSBI
   If (tSBI.rgstate(0) And STATE_SYSTEM_UNAVAILABLE) = STATE_SYSTEM_UNAVAILABLE Then
      m_bHasVerticalScroll = False
   End If
   
   ' Get the current cursor position
   GetCursorPos m_tP
   
   m_hWnd = hwnd
   
   ' Attach subclass for WM_ACTIVATEAPP to the
   ' parent window:
   m_hWndParent = GetParentFormhWNd(m_hWnd)
   AttachMessage Me, m_hWndParent, WM_ACTIVATEAPP
   
   ' Attach mouse hook
   InstallHook Me, WH_MOUSE
   m_bInMiddleScroll = True
   
   ' Show the middle scroller form:
   Set m_frmRefPoint = New frmRefPoint
   m_frmRefPoint.Init ((Not (m_eHorizontalMode = eNone)) And m_bHasHorizontalScroll), _
      ((Not (m_eVerticalMode = eNone)) And m_bHasVerticalScroll)
   Load m_frmRefPoint
   ShowWindow m_frmRefPoint.hwnd, SW_SHOWNOACTIVATE
   MoveWindow m_frmRefPoint.hwnd, m_tP.x - 11, m_tP.y - 11, 23, 23, 1
   
   Set m_tmr = New CTimer
   m_tmr.Interval = 25
   
   m_bInitial = True
      
   '
End Sub

Private Sub EndMiddleScroll()
   
   If Not (m_hWndParent = 0) Then
      DetachMessage Me, m_hWndParent, WM_ACTIVATEAPP
   End If
   
   If (m_bInMiddleScroll) Then
      RemoveHook Me, WH_MOUSE
      m_bInMiddleScroll = False
   End If
   
   If Not (m_frmRefPoint Is Nothing) Then
      Unload m_frmRefPoint
      Set m_frmRefPoint = Nothing
   End If
   
   Screen.MousePointer = vbDefault
End Sub

Private Function GetParentFormhWNd(ByVal lHWnd As Long) As Long
Dim lhWndParent As Long
   lhWndParent = GetParent(lHWnd)
   Do While Not (lhWndParent = 0) And Not (IsWindowVisible(lhWndParent) = 0)
      lHWnd = lhWndParent
      lhWndParent = GetParent(lHWnd)
   Loop
   GetParentFormhWNd = lHWnd
End Function

Private Function getScrollWParam(ByVal lCode As Long, ByVal lPos As Long) As Long
Dim wParam As Long
Dim lHIWord As Long
   wParam = lCode
   lHIWord = (lPos And &H7FFF&) * &H10000
   If (lPos And &H8000&) Then
      lHIWord = lHIWord Or &H80000000
   End If
   wParam = wParam Or lHIWord
   getScrollWParam = wParam
End Function

Private Sub MiddleScroll()
Dim lXOffset As Long
Dim lYOffset As Long
Dim tP As POINTAPI
Dim tSIHorz As SCROLLINFO
Dim tSIVert As SCROLLINFO
Dim lVertOffset As Long
Dim lHorzOffset As Long
Dim lRes As Long
Dim lNewPos As Long
Dim bDoIt As Boolean

   GetCursorPos tP
   lXOffset = tP.x - m_tP.x
   lYOffset = tP.y - m_tP.y
   
   If (m_bHasHorizontalScroll) Then
      tSIHorz.cbSize = Len(tSIHorz)
      tSIHorz.fMask = SIF_RANGE Or SIF_PAGE Or SIF_POS Or SIF_TRACKPOS
      GetScrollInfo m_hWnd, SB_HORZ, tSIHorz
      lHorzOffset = deltaToScrollAmount(lXOffset, m_eHorizontalMode, tSIHorz)
      If Not (lHorzOffset = 0) Then
         bDoIt = True
         If (m_eVerticalMode = eLineBased) Then
            If (lHorzOffset < 32) Then
               If (timeGetTime() - m_lLastHorzTime) < 100 Then
                  bDoIt = False
               Else
                  m_lLastHorzTime = timeGetTime()
               End If
            End If
         End If
         If (bDoIt) Then
            tSIHorz.fMask = SIF_POS Or SIF_TRACKPOS
            lNewPos = tSIHorz.nPos + lHorzOffset
            If (lNewPos < 0) Then lNewPos = 0
            If (lNewPos > tSIHorz.nMax + tSIHorz.nPage) Then lNewPos = tSIHorz.nMax + tSIHorz.nPage
            tSIHorz.nPos = lNewPos
            tSIHorz.nTrackPos = lNewPos
            SetScrollInfo m_hWnd, SB_HORZ, tSIHorz, True
            SendMessageLong m_hWnd, WM_HSCROLL, getScrollWParam(SB_THUMBTRACK, lNewPos), 0
         End If
      End If
   End If
   
   If (m_bHasVerticalScroll) Then
      tSIVert.cbSize = Len(tSIVert)
      tSIVert.fMask = SIF_RANGE Or SIF_PAGE Or SIF_POS Or SIF_TRACKPOS
      GetScrollInfo m_hWnd, SB_VERT, tSIVert
      lVertOffset = deltaToScrollAmount(lYOffset, m_eVerticalMode, tSIVert)
      If Not (lVertOffset = 0) Then
         bDoIt = True
         If (m_eVerticalMode = eLineBased) Then
            If (lVertOffset < 32) Then
               If (timeGetTime() - m_lLastVertTime) < 100 Then
                  bDoIt = False
               Else
                  m_lLastVertTime = timeGetTime()
               End If
            End If
         End If
         If (bDoIt) Then
            tSIVert.fMask = SIF_POS Or SIF_TRACKPOS
            lNewPos = tSIVert.nPos + lVertOffset
            If (lNewPos < 0) Then lNewPos = 0
            If (lNewPos > tSIVert.nMax + tSIVert.nPage) Then lNewPos = tSIVert.nMax + tSIVert.nPage
            tSIVert.nPos = lNewPos
            tSIVert.nTrackPos = lNewPos
            SetScrollInfo m_hWnd, SB_VERT, tSIVert, True
            SendMessageLong m_hWnd, WM_VSCROLL, getScrollWParam(SB_THUMBTRACK, lNewPos), 0
         End If
      End If
   End If
      
   If lVertOffset = lHorzOffset And lVertOffset = 0 Then
      If (m_eHorizontalMode = eNone) Or Not (m_bHasHorizontalScroll) Then
         lRes = 107
      ElseIf (m_eVerticalMode = eNone) Or Not (m_bHasVerticalScroll) Then
         lRes = 106
      Else
         lRes = 105
      End If
   ElseIf Abs(lVertOffset) > Abs(lHorzOffset) And Not (m_eVerticalMode = eNone) Then
      If (lVertOffset > 0) Then
         lRes = 102
      Else
         lRes = 104
      End If
   Else
      If (lHorzOffset > 0) Then
         lRes = 103
      Else
         lRes = 101
      End If
   End If
   
'   If Not (lRes = m_lLastScreenCursor) Then
'      Set Screen.MouseIcon = LoadResPicture(lRes, vbResCursor)
'      Screen.MousePointer = 99
'      m_lLastScreenCursor = lRes
'   End If
   
End Sub

Private Function deltaToScrollAmount(ByVal lDelta As Long, ByVal eMode As EMiddleButtonScrollerMode, tSI As SCROLLINFO) As Long
   
   If (eMode = ePixelBased) Then
   
      If Abs(lDelta) < 12 Then
         deltaToScrollAmount = 0
      ElseIf Abs(lDelta) < 32 Then
         deltaToScrollAmount = 2 * Sgn(lDelta)
      ElseIf Abs(lDelta) < 56 Then
         deltaToScrollAmount = 4 * Sgn(lDelta)
      ElseIf Abs(lDelta) < 80 Then
         deltaToScrollAmount = 8 * Sgn(lDelta)
      ElseIf Abs(lDelta) < 104 Then
         deltaToScrollAmount = 32 * Sgn(lDelta)
      ElseIf Abs(lDelta) < 128 Then
         deltaToScrollAmount = 128 * Sgn(lDelta)
      Else
         deltaToScrollAmount = 512 * Sgn(lDelta)
      End If
            
   ElseIf (eMode = eLineBased) Then
      
      If Abs(lDelta) < 12 Then
         deltaToScrollAmount = 0
      ElseIf Abs(lDelta) < 32 Then
         deltaToScrollAmount = Sgn(lDelta)
      ElseIf Abs(lDelta) < 56 Then
         deltaToScrollAmount = 2 * Sgn(lDelta)
      ElseIf Abs(lDelta) < 80 Then
         deltaToScrollAmount = 4 * Sgn(lDelta)
      ElseIf Abs(lDelta) < 104 Then
         deltaToScrollAmount = 8 * Sgn(lDelta)
      ElseIf Abs(lDelta) < 128 Then
         deltaToScrollAmount = 32 * Sgn(lDelta)
      Else
         deltaToScrollAmount = 80 * Sgn(lDelta)
      End If
      
   End If
      
End Function

Private Sub Class_Initialize()
   m_eHorizontalMode = ePixelBased
   m_eVerticalMode = ePixelBased
End Sub

Private Sub Class_Terminate()
   
   If Not (m_tmr Is Nothing) Then
      m_tmr.Interval = 0
      Set m_tmr = Nothing
   End If
   
   If Not (m_frmRefPoint Is Nothing) Then
      Unload m_frmRefPoint
      Set m_frmRefPoint = Nothing
   End If
   
   EndMiddleScroll
   
End Sub

Private Property Let ISubclass_MsgResponse(ByVal RHS As EMsgResponse)
   '
End Property

Private Property Get ISubclass_MsgResponse() As EMsgResponse
   '
   ISubclass_MsgResponse = emrPreprocess
   '
End Property

Private Function ISubclass_WindowProc(ByVal hwnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
   '
   If (iMsg = WM_ACTIVATEAPP) Then
      If (wParam = 0) Then
         EndMiddleScroll
      End If
   End If
   '
End Function

Private Function IWindowsHook_HookProc(ByVal eType As EHTHookTypeConstants, ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long, bConsume As Boolean) As Long
   '
   Dim tP As POINTAPI
   Dim iTmp As Long
   
   GetCursorPos tP
   If (m_bInitial) Then
      'Debug.Print Abs(tP.X - m_tP.X), Abs(tP.Y - m_tP.Y),
      iTmp = GetAsyncKeyState(vbKeyMButton)
      If Abs(tP.x - m_tP.x) > 4 Or Abs(tP.y - m_tP.y) > 4 Then
         If Not (GetAsyncKeyState(vbKeyMButton) = 0) Then
            m_eMode = eMiddleButtonRelease
         Else
            m_eMode = eAnyButtonPress
         End If
         m_bInitial = False
      End If
   Else
      If (m_eMode = eAnyButtonPress) Then
         iTmp = GetAsyncKeyState(vbKeyLButton)
         iTmp = GetAsyncKeyState(vbKeyRButton)
         iTmp = GetAsyncKeyState(vbKeyMButton)
         If Not (GetAsyncKeyState(vbKeyLButton) = 0) Or _
            Not (GetAsyncKeyState(vbKeyRButton) = 0) Or _
            Not (GetAsyncKeyState(vbKeyMButton) = 0) Then
            EndMiddleScroll
         End If
      Else
         If (GetAsyncKeyState(vbKeyMButton) = 0) Then
            EndMiddleScroll
         End If
      End If
   End If
   
End Function

Private Sub m_tmr_ThatTime()
   If (m_bInMiddleScroll) Then
      MiddleScroll
   Else
      m_tmr.Interval = 0
   End If
End Sub
