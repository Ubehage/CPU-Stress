Attribute VB_Name = "modMouseTests"
Option Explicit

Private Type TestPoint
  X As Long
  Y As Long
End Type
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function GetCapture Lib "user32" () As Long
Public Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal Command As Long) As Long
Public Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function IsWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (ByRef Point As TestPoint) As Long
Private Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, ByRef Point As TestPoint) As Long
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal Buffer As String, ByVal Size As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal Milliseconds As Long)
Private Declare Function WindowFromPoint Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function UpdateWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal After As Long, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long, ByVal Flags As Long) As Long
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal ExStyle As Long, ByVal ClassName As String, ByVal Caption As String, ByVal Style As Long, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long, ByVal Parent As Long, ByVal Menu As Long, ByVal Instance As Long, ByVal Param As Long) As Long
Private Declare Sub ExitProcess Lib "kernel32" (ByVal Code As Long)

Private Const WM_MOUSEMOVE As Long = &H200
Private Const WM_LBUTTONDOWN As Long = &H201
Private Const WM_LBUTTONUP As Long = &H202
Private Const WM_MOUSELEAVE As Long = &H2A3
Private Const WM_CANCELMODE As Long = &H1F
Private Failures As Long
Private Checks As Long
Private LogFile As Integer
Private ModuleClicks As Long
Private ModuleToken As Long
Private PowerEvents As Long

Public Sub Main()
  Dim OldCursor As TestPoint, OldForeground As Long, Attempt As Long
  On Error GoTo Failed
  GetCursorPos OldCursor
  OldForeground = GetForegroundWindow()
  LogFile = FreeFile
  Open App.Path & "\test-results.log" For Output As #LogFile
  Load frmMouseTests
  frmMouseTests.Show
  Call ShowWindow(frmMouseTests.hWnd, 5)
  If LCase$(Trim$(Command$)) = "/manual" Then
    MsgBox "Click OK to run the isolated mouse tests. The pointer will move briefly and then return to its previous position.", vbOKOnly Or vbInformation, "Start mouse integration tests"
  End If
  Call SetForegroundWindow(frmMouseTests.hWnd)
  DoEvents
  ' Activation may still be queued while a hidden launcher/IDE is exiting.
  For Attempt = 1 To 20
    If GetForegroundWindow() = frmMouseTests.hWnd Then Exit For
    frmMouseTests.ZOrder 0
    Call SetForegroundWindow(frmMouseTests.hWnd)
    DoEvents
    Sleep 50
  Next
  Print #LogFile, "Foreground before hook: " & CStr(GetForegroundWindow()) & "; test: " & CStr(frmMouseTests.hWnd)
  HookWindowProc frmMouseTests.hWnd
  Print #LogFile, "Foreground after hook: " & CStr(GetForegroundWindow())
  RunTests
CleanUp:
  MouseShutdown
  UnhookWindowProc
  Unload frmMouseTests
  Unload frmMain
  Call SetCursorPos(OldCursor.X, OldCursor.Y)
  If OldForeground <> 0 Then Call SetForegroundWindow(OldForeground)
  Print #LogFile, CStr(Checks) & " checks; " & CStr(Failures) & " failures."
  Close #LogFile
  If Failures <> 0 Then ExitProcess 1
  Exit Sub
Failed:
  Failures = Failures + 1
  Print #LogFile, "UNEXPECTED: " & CStr(Err.Number) & " - " & Err.Description
  Resume CleanUp
End Sub

Public Sub Check(ByVal Condition As Boolean, ByVal Label As String)
  Checks = Checks + 1
  If Condition Then
    Print #LogFile, "PASS " & Label
  Else
    Failures = Failures + 1
    Print #LogFile, "FAIL " & Label & " [" & MouseLastError() & "]"
  End If
End Sub

Private Function Position(ByVal X As Long, ByVal Y As Long) As Long
  Position = (X And &HFFFF&) Or ((Y And &H7FFF&) * &H10000)
  If Y < 0 Then Position = Position Or &H80000000
End Function

Private Sub PointAt(ByVal hWnd As Long, ByVal X As Long, ByVal Y As Long)
  Dim Point As TestPoint
  Point.X = X
  Point.Y = Y
  ClientToScreen hWnd, Point
  SetCursorPos Point.X, Point.Y
End Sub

Public Sub Move(ByVal hWnd As Long, ByVal X As Long, ByVal Y As Long, Optional ByVal Held As Long = 0)
  PointAt hWnd, X, Y
  SendMessage hWnd, WM_MOUSEMOVE, Held, Position(X, Y)
End Sub

Public Sub Down(ByVal hWnd As Long, Optional ByVal X As Long = 10, Optional ByVal Y As Long = 10)
  PointAt hWnd, X, Y
  SendMessage hWnd, WM_LBUTTONDOWN, 1, Position(X, Y)
End Sub

Public Sub Up(ByVal hWnd As Long, Optional ByVal X As Long = 10, Optional ByVal Y As Long = 10)
  PointAt hWnd, X, Y
  SendMessage hWnd, WM_LBUTTONUP, 0, Position(X, Y)
End Sub

Private Sub RunTests()
  Dim A As TestMouseSink, B As TestMouseSink, Adapter As ModuleMouseSink
  Dim hA As Long, hB As Long, TokenA As Long, TokenB As Long, OldToken As Long
  Dim i As Long, hTemporary As Long, TemporaryToken As Long
  Dim BeforeClicks As Long, OldLeft As Single, OldTop As Single
  hA = frmMouseTests.TargetA.hWnd
  hB = frmMouseTests.TargetB.hWnd
  Set A = New TestMouseSink
  Set B = New TestMouseSink
  TokenA = MouseRegister(hA, A)
  Check TokenA <> 0, "register native VB6 control: " & MouseLastError()
  If TokenA = 0 Then Exit Sub
  TokenB = MouseRegister(hB, B)
  Check TokenB <> 0 And TokenA <> TokenB, "independent registration tokens"
  Check MouseRegister(hA, B) = 0, "reject duplicate window"
  Check MouseRegister(0, B) = 0, "reject invalid window"
  If GetForegroundWindow() <> frmMouseTests.hWnd Then Err.Raise 5, , "Test form could not become foreground; current window class: " & ForegroundClass()

  Move hA, 10, 10
  Check GetCapture() = 0, "passive hover does not capture"
  Check A.Trace = "1:1,3:1,", "enter then move: " & A.Trace
  A.Trace = vbNullString
  Down hA
  Check GetCapture() = hA, "left press owns capture"
  Up hA
  Check A.Trace = "4:1,5:1,6:1,", "VB6 default release preserves Up then Click: " & A.Trace
  Check A.Clicks = 1 And A.Cancels = 0, "ordinary click does not cancel"
  Check GetCapture() = 0, "release clears capture"

  A.Trace = vbNullString
  Down hA
  Move hA, -5, -7, 1
  Check A.LastX = -5 And A.LastY = -7, "signed captured coordinates"
  Up hA, -5, -7
  Check A.Clicks = 1, "release outside does not click"
  Down hA
  Move hA, -5, 10, 1
  Move hA, 10, 10, 1
  Up hA
  Check A.Clicks = 2, "drag out and back clicks original region"

  A.Trace = vbNullString
  Move hA, 70, 10
  Check A.Trace = "2:1,1:2,3:2,", "region transition ordering: " & A.Trace
  Down hA
  Up hA, 70, 10
  Check A.Clicks = 2, "release on another region does not click"
  Down hA, 130, 10
  Up hA, 130, 10
  Check A.Clicks = 2, "empty region never clicks zero"

  Down hA
  Call SetCapture(hB)
  Check A.LastReason = mcCaptureLost, "capture theft cancels original gesture"
  Check GetCapture() = hB, "cancellation does not release another window's capture"
  Call ReleaseCapture
  Up hA
  Check A.Clicks = 2, "capture theft prevents later click"

  Down hA
  MouseSetEnabled TokenA, False
  Check GetCapture() = 0 And A.LastReason = mcDisabled, "disable while pressed releases capture"
  Up hA
  Check A.Clicks = 2, "disabled target cannot click"
  MouseSetEnabled TokenA, True

  Down hA
  MouseCancelAll mcDeactivated
  Check GetCapture() = 0 And A.LastReason = mcDeactivated, "owner deactivation cancels"
  Up hA
  Check A.Clicks = 2, "deactivation prevents click"
  Down hA
  SendMessage hA, WM_CANCELMODE, 0, 0
  Check GetCapture() = 0 And A.LastReason = mcCancelMode, "native cancellation releases capture"

  Down hA
  frmMouseTests.TargetA.Visible = False
  Check GetCapture() = 0 And A.LastReason = mcHidden, "hiding a pressed window cancels"
  frmMouseTests.TargetA.Visible = True
  Move hA, 10, 10
  SendMessage hA, WM_MOUSELEAVE, 0, 0
  A.Trace = vbNullString
  Move hA, 10, 10
  Check A.Trace = "1:1,3:1,", "leave/reenter rearms hover"

  A.ActionKind = meUp
  A.Action = "disable"
  Down hA
  Up hA
  Check A.Clicks = 2 And GetCapture() = 0, "disable from Up suppresses Click"
  MouseSetEnabled TokenA, True
  A.ActionKind = meUp
  A.Action = "unregister"
  Down hA
  Up hA
  Check A.Clicks = 2 And GetCapture() = 0, "unregister from Up suppresses Click"
  OldToken = TokenA
  TokenA = MouseRegister(hA, A)
  Check TokenA <> 0 And TokenA <> OldToken, "re-registration gets a fresh token"
  MouseUnregister OldToken
  Down hA
  Up hA
  Check A.Clicks = 3, "stale token cannot remove replacement"

  A.ActionKind = meDown
  A.Action = "error"
  Down hA
  Check GetCapture() = 0, "recipient failure releases capture"
  Check InStr(MouseLastError(), "Deliberate recipient failure") > 0, "recipient failure is recorded"
  TokenA = MouseRegister(hA, A)
  Check TokenA <> 0, "recipient failure detached registration"
  A.FailHitTest = True
  Move hA, 10, 10
  Check InStr(MouseLastError(), "Deliberate hit-test failure") > 0, "hit-test failure is contained"
  A.FailHitTest = False
  TokenA = MouseRegister(hA, A)

  A.ActionKind = meUp
  A.Action = "reenter"
  A.TargetWindow = hA
  BeforeClicks = A.Clicks
  Down hA
  Up hA
  Check A.Clicks = BeforeClicks And GetCapture() = hA, "nested Down invalidates outer pending Click"
  Up hA
  Check A.Clicks = BeforeClicks + 1, "nested gesture completes independently"

  A.ActionKind = meUp
  A.Action = "move"
  BeforeClicks = A.Clicks
  Down hA
  Up hA
  Check A.Clicks = BeforeClicks, "moving target from Up invalidates pending Click"
  frmMouseTests.TargetA.Move 10, 10, 180, 80

  Move hA, 10, 10
  OldLeft = frmMouseTests.Left
  OldTop = frmMouseTests.Top
  A.Trace = vbNullString
  frmMouseTests.Move OldLeft + Screen.TwipsPerPixelX * 240, OldTop
  Check InStr(A.Trace, "2:1,") > 0, "parent movement updates stationary pointer through existing hook"
  frmMouseTests.Move OldLeft, OldTop
  A.Trace = vbNullString
  MouseRefreshAll
  Check GetCapture() = 0, "parent geometry refresh does not capture"
  SendMessage frmMouseTests.hWnd, &H218, &H4, 0
  SendMessage frmMouseTests.hWnd, &H218, &H12, 0
  Check PowerEvents = 2, "hibernation hook still dispatches power notifications"

  Down hA
  ' STATIC text is intentionally skipped by WindowFromPoint. Use a real
  ' input-receiving window and prove it is above the original target.
  hTemporary = CreateWindowEx(0, "BUTTON", "", &H50000000, 10, 10, 100, 40, frmMouseTests.hWnd, 0, App.hInstance, 0)
  SetWindowPos hTemporary, 0, 0, 0, 0, 0, &H13
  Check hTemporary <> 0 And WindowAt(hA, 10, 10) = hTemporary, "occlusion fixture owns pointer hit test"
  BeforeClicks = A.Clicks
  Up hA
  Check A.Clicks = BeforeClicks, "occluded target cannot click by geometry alone"
  DestroyWindow hTemporary

  MouseUnregister TokenB
  Set Adapter = New ModuleMouseSink
  ModuleToken = MouseRegister(hB, Adapter)
  Down hB
  Up hB
  Check ModuleClicks = 1, "class adapter delivers events to BAS module"
  MouseUnregister ModuleToken

  For i = 1 To 20
    TokenB = MouseRegister(hB, B)
    MouseUnregister TokenB
    MouseUnregister TokenB
  Next
  TokenB = MouseRegister(hB, B)
  Check TokenB <> 0, "repeated register/unregister is idempotent"

  hTemporary = CreateWindowEx(0, "BUTTON", "", &H50000000, 210, 100, 100, 40, frmMouseTests.hWnd, 0, App.hInstance, 0)
  SetWindowPos hTemporary, 0, 0, 0, 0, 0, &H13
  Set B = New TestMouseSink
  B.TargetWindow = hTemporary
  TemporaryToken = MouseRegister(hTemporary, B)
  Check TemporaryToken <> 0, "register temporary native child"
  Check WindowAt(hTemporary, 10, 10) = hTemporary, "destruction fixture owns pointer hit test"
  B.ActionKind = meUp
  B.Action = "destroy"
  Down hTemporary
  Check InStr(B.Trace, "4:1,") > 0 And GetCapture() = hTemporary, "destruction fixture starts a real gesture: " & B.Trace
  Up hTemporary
  Check InStr(B.Trace, "5:1,") > 0 And IsWindow(hTemporary) = 0, "Up callback destroys the registered window"
  Check B.Clicks = 0 And GetCapture() = 0, "destroy during Up safely prevents Click"
  MouseUnregister TemporaryToken
  If IsWindow(hTemporary) <> 0 Then DestroyWindow hTemporary

  TestRealButton
  TestCheckBoxes
  TestCPUViews
  TestMouseLifetime

  Down hA
  MouseShutdown
  MouseShutdown
  Check GetCapture() = 0, "shutdown releases capture and tolerates repetition"
End Sub

Private Function ForegroundClass() As String
  Dim Buffer As String, Length As Long
  Buffer = Space$(128)
  Length = GetClassName(GetForegroundWindow(), Buffer, Len(Buffer))
  ForegroundClass = Left$(Buffer, Length)
End Function

Private Function ButtonWindow() As Long
  Dim Point As TestPoint
  Point.X = frmMouseTests.TestButton.Left + 10
  Point.Y = frmMouseTests.TestButton.Top + 10
  ClientToScreen frmMouseTests.hWnd, Point
  ButtonWindow = WindowFromPoint(Point.X, Point.Y)
End Function

Public Function ControlWindow(ByVal Control As Object) As Long
  ' Harness containers use pixel ScaleMode; no production hWnd property needed.
  ControlWindow = WindowAt(Control.Container.hWnd, CLng(Control.Left) + 5, CLng(Control.Top) + 5)
End Function

Public Function PixelAt(ByVal hWnd As Long, ByVal X As Long, ByVal Y As Long) As Long
  Dim hDC As Long
  UpdateWindow hWnd
  hDC = GetDC(hWnd)
  PixelAt = GetPixel(hDC, X, Y)
  ReleaseDC hWnd, hDC
End Function

Private Function WindowAt(ByVal hWnd As Long, ByVal X As Long, ByVal Y As Long) As Long
  Dim Point As TestPoint
  Point.X = X
  Point.Y = Y
  ClientToScreen hWnd, Point
  WindowAt = WindowFromPoint(Point.X, Point.Y)
End Function

Private Function ButtonColor(ByVal hWnd As Long) As Long
  Dim hDC As Long
  UpdateWindow hWnd
  hDC = GetDC(hWnd)
  ButtonColor = GetPixel(hDC, 5, 5)
  ReleaseDC hWnd, hDC
End Function

Private Sub TestRealButton()
  Dim hButton As Long, Before As Long
  hButton = ButtonWindow()
  Check hButton <> 0 And hButton <> frmMouseTests.hWnd, "find real migrated Button window"
  Move hButton, 10, 10
  Check GetCapture() = 0, "real Button passive hover leaves capture free"
  Check ButtonColor(hButton) = COLOR_BUTTON_HOVER, "real Button paints hover"
  Before = frmMouseTests.ButtonClicks
  Down hButton
  Check ButtonColor(hButton) = COLOR_BUTTON_PRESSED, "real Button paints pressed"
  Move hButton, -10, 10, 1
  Check ButtonColor(hButton) = COLOR_BUTTON_PRESSED, "real Button retains pressed appearance outside"
  Move hButton, 10, 10, 1
  Up hButton
  Check frmMouseTests.ButtonClicks = Before + 1, "real Button raises exactly one Click"
  Check GetCapture() = 0, "real Button click leaves capture free"

  Down hButton
  frmMouseTests.TestButton.Enabled = False
  Check GetCapture() = 0, "real Button disable during press releases capture"
  Up hButton
  Check frmMouseTests.ButtonClicks = Before + 1, "real disabled Button does not click"
  Check ButtonColor(hButton) = COLOR_BACKGROUND_DISABLED, "real Button paints disabled"
  frmMouseTests.TestButton.Enabled = True
  frmMouseTests.DisableOnClick = True
  Down hButton
  Up hButton
  Check Not frmMouseTests.TestButton.Enabled And GetCapture() = 0, "Click handler can disable real Button without recapture"
  frmMouseTests.DisableOnClick = False
  frmMouseTests.TestButton.Enabled = True

  Before = frmMouseTests.ButtonClicks
  frmMouseTests.TestButton.SetFocus
  SendMessage hButton, &H100, vbKeySpace, 0
  SendMessage hButton, &H101, vbKeySpace, 0
  Check frmMouseTests.ButtonClicks = Before + 1, "real Button Space activation"
  SendMessage hButton, &H100, vbKeyReturn, 0
  SendMessage hButton, &H101, vbKeyReturn, 0
  Check frmMouseTests.ButtonClicks = Before + 2, "real Button Return activation"
  SendMessage hButton, &H100, vbKeySpace, 0
  MouseCancelAll mcDeactivated
  SendMessage hButton, &H101, vbKeySpace, 0
  Check frmMouseTests.ButtonClicks = Before + 2, "deactivation cancels pending keyboard activation"

  Down hButton
  frmMouseTests.TestButton.Visible = False
  Check GetCapture() = 0, "real Button hide releases capture"
  frmMouseTests.TestButton.Visible = True
  DoEvents
  hButton = ButtonWindow()
  Down hButton
  Up hButton
  Check frmMouseTests.ButtonClicks = Before + 3, "real Button re-registers after hide/show"
End Sub

Public Sub HandleModuleMouse(ByVal Token As Long, ByVal Kind As MouseEventKind)
  If Token = ModuleToken And Kind = meClick Then ModuleClicks = ModuleClicks + 1
End Sub

Public Sub StartNestedPress(ByVal hWnd As Long)
  Down hWnd
End Sub

Public Sub MoveDuringCallback(ByVal hWnd As Long)
  SetWindowPos hWnd, 0, 210, 100, 0, 0, &H15
End Sub

Public Sub RecordPowerEvent(ByVal Suspended As Boolean)
  PowerEvents = PowerEvents + 1
End Sub
