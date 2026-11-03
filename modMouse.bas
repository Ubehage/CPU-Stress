Attribute VB_Name = "modMouse"
Option Explicit

Public Enum MouseEventKind
  meEnter = 1
  meLeave = 2
  meMove = 3
  meDown = 4
  meUp = 5
  meClick = 6
  meCancel = 7
End Enum

Public Enum MouseCancelReason
  mcNone = 0
  mcCaptureLost = 1
  mcDisabled = 2
  mcHidden = 3
  mcDeactivated = 4
  mcCancelMode = 5
  mcRequested = 6
End Enum

Private Type MousePoint
  X As Long
  Y As Long
End Type
Private Type MouseRect
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type
Private Type MouseTracking
  cbSize As Long
  dwFlags As Long
  hwndTrack As Long
  dwHoverTime As Long
End Type

Private Declare Function SetWindowSubclass Lib "comctl32" (ByVal hWnd As Long, ByVal Callback As Long, ByVal Id As Long, ByVal RefData As Long) As Long
Private Declare Function RemoveWindowSubclass Lib "comctl32" (ByVal hWnd As Long, ByVal Callback As Long, ByVal Id As Long) As Long
Private Declare Function DefSubclassProc Lib "comctl32" (ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function TrackMouseEvent Lib "user32" (ByRef Tracking As MouseTracking) As Long
Private Declare Function GetCurrentThreadId Lib "kernel32" () As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, ByRef ProcessId As Long) As Long
Private Declare Function IsWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function IsWindowEnabled Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetAncestor Lib "user32" (ByVal hWnd As Long, ByVal Flags As Long) As Long
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function GetCursorPos Lib "user32" (ByRef Point As MousePoint) As Long
Private Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, ByRef Point As MousePoint) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, ByRef Point As MousePoint) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, ByRef Rect As MouseRect) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function GetCapture Lib "user32" () As Long
Private Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function GetKeyState Lib "user32" (ByVal Key As Long) As Integer

Private Const WM_NCDESTROY As Long = &H82
Private Const WM_CANCELMODE As Long = &H1F
Private Const WM_ENABLE As Long = &HA
Private Const WM_SHOWWINDOW As Long = &H18
Private Const WM_WINDOWPOSCHANGED As Long = &H47
Private Const WM_MOUSEMOVE As Long = &H200
Private Const WM_LBUTTONDOWN As Long = &H201
Private Const WM_LBUTTONUP As Long = &H202
Private Const WM_LBUTTONDBLCLK As Long = &H203
Private Const WM_RBUTTONDOWN As Long = &H204
Private Const WM_RBUTTONUP As Long = &H205
Private Const WM_RBUTTONDBLCLK As Long = &H206
Private Const WM_MBUTTONDOWN As Long = &H207
Private Const WM_MBUTTONUP As Long = &H208
Private Const WM_MBUTTONDBLCLK As Long = &H209
Private Const WM_CAPTURECHANGED As Long = &H215
Private Const WM_MOUSELEAVE As Long = &H2A3
Private Const TME_LEAVE As Long = &H2
Private Const TME_CANCEL As Long = &H80000000
Private Const GA_ROOT As Long = 2

Private Targets As Collection
Private NextToken As Long
Private LastErrorText As String
Private ShuttingDown As Boolean

Public Function MouseLastError() As String
  MouseLastError = LastErrorText
End Function

Public Function MouseNativeAvailable() As Boolean
  Dim InIDE As Boolean
  Debug.Assert MarkIDE(InIDE)
  MouseNativeAvailable = Not InIDE
End Function

Private Function MarkIDE(ByRef InIDE As Boolean) As Boolean
  InIDE = True
  MarkIDE = True
End Function

Public Function MouseRegister(ByVal hWnd As Long, ByVal Sink As IMouseSink) As Long
  Dim Target As CMouseTarget, Other As CMouseTarget, ProcessId As Long
  On Error GoTo Failed
  LastErrorText = vbNullString
  If ShuttingDown Then Exit Function
  If Not MouseNativeAvailable() Then
    LastErrorText = "Native mouse tracking is disabled in the VB6 IDE."
    Exit Function
  End If
  If Sink Is Nothing Then Exit Function
  If IsWindow(hWnd) = 0 Then Exit Function
  If GetWindowThreadProcessId(hWnd, ProcessId) <> GetCurrentThreadId() Then Exit Function
  If Targets Is Nothing Then Set Targets = New Collection
  For Each Other In Targets
    If Other.hWnd = hWnd Then Exit Function
  Next
  If NextToken = &H7FFFFFFF Then Exit Function
  NextToken = NextToken + 1
  Set Target = New CMouseTarget
  Target.Token = NextToken
  Target.hWnd = hWnd
  Set Target.Sink = Sink
  Target.Enabled = True
  Target.Active = True
  Targets.Add Target, CStr(Target.Token)
  If SetWindowSubclass(hWnd, AddressOf MouseWindowProc, Target.Token, Target.Token) = 0 Then
    LastErrorText = "SetWindowSubclass failed."
    MouseUnregister Target.Token
    Exit Function
  End If
  MouseRegister = Target.Token
  Exit Function
Failed:
  LastErrorText = "Register: " & CStr(Err.Number) & " - " & Err.Description
  If Not Target Is Nothing Then MouseUnregister Target.Token
End Function

Private Function FindTarget(ByVal Token As Long) As CMouseTarget
  If Targets Is Nothing Then Exit Function
  On Error Resume Next
  Set FindTarget = Targets(CStr(Token))
End Function

Public Sub MouseUnregister(ByVal Token As Long)
  Dim Target As CMouseTarget
  Set Target = FindTarget(Token)
  If Target Is Nothing Then Exit Sub
  ' Invalidate before calls into Windows: releasing capture can reenter us.
  Target.Active = False
  Target.PressRegion = 0
  Target.HoverRegion = 0
  Targets.Remove CStr(Token)
  On Error Resume Next
  Call RemoveWindowSubclass(Target.hWnd, AddressOf MouseWindowProc, Token)
  StopTracking Target
  If GetCapture() = Target.hWnd Then Call ReleaseCapture
  Set Target.Sink = Nothing
End Sub

Public Sub MouseShutdown()
  If Targets Is Nothing Then Exit Sub
  If ShuttingDown Then Exit Sub
  ShuttingDown = True
  Do While Targets.Count > 0
    MouseUnregister Targets(1).Token
  Loop
  ShuttingDown = False
End Sub

Public Sub MouseSetEnabled(ByVal Token As Long, ByVal Enabled As Boolean)
  Dim Target As CMouseTarget
  Set Target = FindTarget(Token)
  If Target Is Nothing Then Exit Sub
  If Target.Enabled = Enabled Then Exit Sub
  Target.Enabled = Enabled
  If Not Enabled Then
    MouseCancel Token, mcDisabled
  Else
    Target.Revision = Target.Revision + 1
    MouseRefresh Token
  End If
End Sub

Public Sub MouseCancel(ByVal Token As Long, ByVal Reason As MouseCancelReason)
  Dim Target As CMouseTarget, OldHover As Long, OldPress As Long, Revision As Long
  Set Target = FindTarget(Token)
  If Target Is Nothing Then Exit Sub
  OldHover = Target.HoverRegion
  OldPress = Target.PressRegion
  Target.HoverRegion = 0
  Target.PressRegion = 0
  Target.Revision = Target.Revision + 1
  Revision = Target.Revision
  StopTracking Target
  If GetCapture() = Target.hWnd Then Call ReleaseCapture
  If Not Current(Target, Revision) Then Exit Sub
  If Reason = mcNone Then Reason = mcRequested
  Notify Target, meCancel, OldPress, 0, Reason
  If Not Current(Target, Revision) Then Exit Sub
  If OldHover <> 0 Then Notify Target, meLeave, OldHover
End Sub

' Call from the owner form's activation/layout/lifecycle handlers. Take a
' snapshot because a recipient may register/unregister another recipient.
Public Sub MouseCancelAll(ByVal Reason As MouseCancelReason)
  Dim Snapshot As Collection, Target As CMouseTarget
  Set Snapshot = SnapshotTargets()
  For Each Target In Snapshot
    If Target.Active Then MouseCancel Target.Token, Reason
  Next
End Sub

Public Sub MouseRefreshAll()
  Dim Snapshot As Collection, Target As CMouseTarget
  Set Snapshot = SnapshotTargets()
  For Each Target In Snapshot
    If Target.Active Then MouseRefresh Target.Token
  Next
End Sub

Private Function SnapshotTargets() As Collection
  Dim Target As CMouseTarget
  Set SnapshotTargets = New Collection
  If Targets Is Nothing Then Exit Function
  For Each Target In Targets
    SnapshotTargets.Add Target
  Next
End Function

Public Sub MouseRefresh(ByVal Token As Long)
  Dim Target As CMouseTarget, Point As MousePoint
  Set Target = FindTarget(Token)
  If Target Is Nothing Then Exit Sub
  If GetCursorPos(Point) = 0 Then Exit Sub
  If ScreenToClient(Target.hWnd, Point) = 0 Then Exit Sub
  Target.Revision = Target.Revision + 1
  Target.X = Point.X
  Target.Y = Point.Y
  UpdateHover Target
End Sub

Private Function Current(ByVal Target As CMouseTarget, ByVal Revision As Long) As Boolean
  Current = Target.Active And Target.Revision = Revision
End Function

Private Function Eligible(ByVal Target As CMouseTarget) As Boolean
  Dim hParent As Long
  If Not Target.Active Or Not Target.Enabled Then Exit Function
  If IsWindowVisible(Target.hWnd) = 0 Then Exit Function
  If GetForegroundWindow() <> GetAncestor(Target.hWnd, GA_ROOT) Then Exit Function
  hParent = Target.hWnd
  Do While hParent <> 0
    If IsWindowEnabled(hParent) = 0 Then Exit Function
    hParent = GetParent(hParent)
  Loop
  Eligible = True
End Function

Private Function RegionAt(ByVal Target As CMouseTarget) As Long
  Dim Rect As MouseRect, Point As MousePoint, Sink As IMouseSink, Revision As Long, Region As Long
  On Error GoTo Failed
  If Target.HitTesting Then Exit Function
  If Not Eligible(Target) Then Exit Function
  If GetClientRect(Target.hWnd, Rect) = 0 Then Exit Function
  If Target.X < 0 Or Target.Y < 0 Or Target.X >= Rect.Right Or Target.Y >= Rect.Bottom Then Exit Function
  Point.X = Target.X
  Point.Y = Target.Y
  If ClientToScreen(Target.hWnd, Point) = 0 Then Exit Function
  If WindowFromPoint(Point.X, Point.Y) <> Target.hWnd Then Exit Function
  Revision = Target.Revision
  Set Sink = Target.Sink
  Target.HitTesting = True
  Region = Sink.HitTest(Target.Token, Target.X, Target.Y)
  Target.HitTesting = False
  If Current(Target, Revision) And Region > 0 Then RegionAt = Region
  Exit Function
Failed:
  Target.HitTesting = False
  FailTarget Target, "HitTest", Err.Number, Err.Description
End Function

Private Sub Notify(ByVal Target As CMouseTarget, ByVal Kind As MouseEventKind, ByVal Region As Long, _
    Optional ByVal Button As Integer = 0, Optional ByVal Reason As MouseCancelReason = mcNone)
  Dim Sink As IMouseSink
  On Error GoTo Failed
  If Not Target.Active Then Exit Sub
  Set Sink = Target.Sink
  Sink.OnMouse Target.Token, Kind, Region, Button, Target.Buttons, Target.Shift, Target.X, Target.Y, Reason
  Exit Sub
Failed:
  FailTarget Target, "OnMouse", Err.Number, Err.Description
End Sub

Private Sub FailTarget(ByVal Target As CMouseTarget, ByVal Operation As String, ByVal Number As Long, ByVal Description As String)
  LastErrorText = Operation & " (" & CStr(Target.Token) & "): " & CStr(Number) & " - " & Description
  MouseUnregister Target.Token
End Sub

Private Sub StopTracking(ByVal Target As CMouseTarget)
  Dim Tracking As MouseTracking
  If Not Target.Tracking Then Exit Sub
  Target.Tracking = False
  Tracking.cbSize = Len(Tracking)
  Tracking.hwndTrack = Target.hWnd
  Tracking.dwFlags = TME_CANCEL Or TME_LEAVE
  Call TrackMouseEvent(Tracking)
End Sub

Private Sub UpdateHover(ByVal Target As CMouseTarget)
  Dim Region As Long, OldRegion As Long, Revision As Long, Tracking As MouseTracking
  Revision = Target.Revision
  Region = RegionAt(Target)
  If Not Current(Target, Revision) Then Exit Sub
  If Region <> Target.HoverRegion Then
    OldRegion = Target.HoverRegion
    Target.HoverRegion = 0
    If OldRegion <> 0 Then Notify Target, meLeave, OldRegion
    If Not Current(Target, Revision) Then Exit Sub
    Target.HoverRegion = Region
    If Region <> 0 Then Notify Target, meEnter, Region
    If Not Current(Target, Revision) Then Exit Sub
  End If
  If Region <> 0 And Not Target.Tracking Then
    Tracking.cbSize = Len(Tracking)
    Tracking.hwndTrack = Target.hWnd
    Tracking.dwFlags = TME_LEAVE
    Target.Tracking = (TrackMouseEvent(Tracking) <> 0)
  End If
End Sub

Private Function SignedWord(ByVal Value As Long) As Long
  Value = Value And &HFFFF&
  If Value >= &H8000& Then Value = Value - &H10000
  SignedWord = Value
End Function

Private Sub DispatchMouse(ByVal Target As CMouseTarget, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long)
  Dim Revision As Long, Origin As Long, Region As Long, Button As Integer
  Target.Revision = Target.Revision + 1
  Target.X = SignedWord(lParam)
  Target.Y = SignedWord((lParam And &HFFFF0000) \ &H10000)
  Target.Buttons = 0
  If (wParam And 1) <> 0 Then Target.Buttons = Target.Buttons Or vbLeftButton
  If (wParam And 2) <> 0 Then Target.Buttons = Target.Buttons Or vbRightButton
  If (wParam And &H10) <> 0 Then Target.Buttons = Target.Buttons Or vbMiddleButton
  Target.Shift = 0
  If (wParam And 4) <> 0 Then Target.Shift = Target.Shift Or vbShiftMask
  If (wParam And 8) <> 0 Then Target.Shift = Target.Shift Or vbCtrlMask
  If GetKeyState(vbKeyMenu) < 0 Then Target.Shift = Target.Shift Or vbAltMask
  Revision = Target.Revision
  UpdateHover Target
  If Not Current(Target, Revision) Then Exit Sub
  If Not Eligible(Target) Then
    MouseCancel Target.Token, mcDeactivated
    Exit Sub
  End If
  Select Case Msg
    Case WM_MOUSEMOVE
      Notify Target, meMove, Target.HoverRegion
    Case WM_LBUTTONDOWN, WM_LBUTTONDBLCLK
      If Target.PressRegion <> 0 Or Target.HoverRegion = 0 Then Exit Sub
      Target.PressRegion = Target.HoverRegion
      ' Default processing may already own capture. Reasserting it can send
      ' capture-change notifications back through a native control's handler.
      If GetCapture() <> Target.hWnd Then Call SetCapture(Target.hWnd)
      If Not Current(Target, Revision) Then Exit Sub
      If GetCapture() <> Target.hWnd Then
        MouseCancel Target.Token, mcCaptureLost
        Exit Sub
      End If
      Notify Target, meDown, Target.PressRegion, vbLeftButton
    Case WM_LBUTTONUP
      Origin = Target.PressRegion
      Target.PressRegion = 0
      If GetCapture() = Target.hWnd Then Call ReleaseCapture
      If Not Current(Target, Revision) Then Exit Sub
      Notify Target, meUp, Target.HoverRegion, vbLeftButton
      If Not Current(Target, Revision) Or Origin = 0 Then Exit Sub
      ' Up can hide, disable, unregister or move the control without pumping messages.
      Region = RegionAt(Target)
      If Current(Target, Revision) And Region = Origin Then Notify Target, meClick, Origin, vbLeftButton
    Case WM_RBUTTONDOWN, WM_RBUTTONDBLCLK, WM_MBUTTONDOWN, WM_MBUTTONDBLCLK
      If Msg = WM_MBUTTONDOWN Or Msg = WM_MBUTTONDBLCLK Then Button = vbMiddleButton Else Button = vbRightButton
      Notify Target, meDown, Target.HoverRegion, Button
    Case WM_RBUTTONUP, WM_MBUTTONUP
      If Msg = WM_MBUTTONUP Then Button = vbMiddleButton Else Button = vbRightButton
      Notify Target, meUp, Target.HoverRegion, Button
  End Select
End Sub

' Windows entry point. Never allow a VB error to unwind into native code.
Public Function MouseWindowProc(ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, _
    ByVal lParam As Long, ByVal Id As Long, ByVal RefData As Long) As Long
  Dim Target As CMouseTarget, CalledDefault As Boolean, DefaultUp As Boolean, OldHover As Long
  On Error GoTo Failed
  Set Target = FindTarget(RefData)
  If Target Is Nothing Then
    MouseWindowProc = DefSubclassProc(hWnd, Msg, wParam, lParam)
    Exit Function
  End If
  If Msg = WM_NCDESTROY Then
    MouseUnregister Target.Token
  ElseIf Msg = WM_CAPTURECHANGED Then
    If Target.PressRegion <> 0 Then
      If Target.DefaultUpDepth = 0 Or lParam <> 0 Then MouseCancel Target.Token, mcCaptureLost
    End If
  ElseIf Msg = WM_CANCELMODE Then
    MouseCancel Target.Token, mcCancelMode
  End If
  ' VB6 may release its automatic capture inside default MouseUp handling.
  ' Preserve our gesture through that expected release, then dispatch Up/Click.
  DefaultUp = (Msg = WM_LBUTTONUP)
  If DefaultUp Then Target.DefaultUpDepth = Target.DefaultUpDepth + 1
  CalledDefault = True
  MouseWindowProc = DefSubclassProc(hWnd, Msg, wParam, lParam)
  If DefaultUp Then Target.DefaultUpDepth = Target.DefaultUpDepth - 1
  DefaultUp = False
  If Not Target.Active Then Exit Function
  Select Case Msg
    Case WM_MOUSEMOVE To WM_MBUTTONDBLCLK
      DispatchMouse Target, Msg, wParam, lParam
    Case WM_MOUSELEAVE
      Target.Tracking = False
      OldHover = Target.HoverRegion
      Target.HoverRegion = 0
      If OldHover <> 0 Then Notify Target, meLeave, OldHover
    Case WM_ENABLE
      If wParam = 0 Then MouseCancel Target.Token, mcDisabled
    Case WM_SHOWWINDOW
      If wParam = 0 Then MouseCancel Target.Token, mcHidden
    Case WM_WINDOWPOSCHANGED
      If IsWindowVisible(hWnd) = 0 Then
        MouseCancel Target.Token, mcHidden
      Else
        MouseRefresh Target.Token
      End If
  End Select
  Exit Function
Failed:
  If Not Target Is Nothing Then
    If DefaultUp Then Target.DefaultUpDepth = Target.DefaultUpDepth - 1
    FailTarget Target, "WindowProc", Err.Number, Err.Description
  End If
  On Error Resume Next
  If Not CalledDefault Then MouseWindowProc = DefSubclassProc(hWnd, Msg, wParam, lParam)
End Function
