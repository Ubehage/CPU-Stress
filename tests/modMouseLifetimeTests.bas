Attribute VB_Name = "modMouseLifetimeTests"
Option Explicit

Private Type LifetimeRect
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type
Private Declare Function IsWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, ByRef Rect As LifetimeRect) As Long
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetForegroundWindow Lib "user32" () As Long

Private CheckClicks As Long
Private CPUClicks As Long
Private LastCPUIndex As Long

Public Sub RecordLifetimeCheckClick()
  CheckClicks = CheckClicks + 1
End Sub

Public Sub RecordLifetimeCPUClick(ByVal Index As Long)
  CPUClicks = CPUClicks + 1
  LastCPUIndex = Index
End Sub

Public Sub TestMouseLifetime()
  Dim hTarget As Long, Before As Long, X As Long, Y As Long, Iteration As Long
  Dim ErrorNumber As Long, ErrorDescription As String
  On Error GoTo Failed

  OpenLifetimeForm
  hTarget = modMouseTests.ControlWindow(frmMouseLifetime.CheckControl)
  Before = CheckClicks
  modMouseTests.Down hTarget
  modMouseTests.Check GetCapture() = hTarget, "lifetime CheckBox starts captured press"
  Unload frmMouseLifetime
  modMouseTests.Check IsWindow(hTarget) = 0 And GetCapture() = 0, "unloading pressed CheckBox destroys window and releases capture"
  SendMessage hTarget, &H202, 0, 0
  modMouseTests.Check CheckClicks = Before, "unloaded CheckBox cannot emit delayed Click"

  For Iteration = 1 To 2
    OpenLifetimeForm
    frmMouseLifetime.UnloadCheckOnClick = True
    hTarget = modMouseTests.ControlWindow(frmMouseLifetime.CheckControl)
    Before = CheckClicks
    modMouseTests.Down hTarget
    modMouseTests.Check GetCapture() = hTarget, "reloaded CheckBox registers for capture " & CStr(Iteration)
    modMouseTests.Up hTarget
    modMouseTests.Check CheckClicks = Before + 1, "CheckBox unload callback receives exactly one Click " & CStr(Iteration)
    modMouseTests.Check IsWindow(hTarget) = 0 And GetCapture() = 0, "CheckBox can unload its form from Click " & CStr(Iteration)
  Next

  OpenLifetimeForm
  hTarget = modMouseTests.ControlWindow(frmMouseLifetime.CPUControl)
  FindCPUPoint hTarget, X, Y
  Before = CPUClicks
  modMouseTests.Down hTarget, X, Y
  modMouseTests.Check GetCapture() = hTarget, "lifetime CPUView starts captured region press"
  Unload frmMouseLifetime
  modMouseTests.Check IsWindow(hTarget) = 0 And GetCapture() = 0, "unloading pressed CPUView destroys window and releases capture"
  SendMessage hTarget, &H202, 0, 0
  modMouseTests.Check CPUClicks = Before, "unloaded CPUView cannot emit delayed Click"

  For Iteration = 1 To 2
    OpenLifetimeForm
    frmMouseLifetime.UnloadCPUOnClick = True
    hTarget = modMouseTests.ControlWindow(frmMouseLifetime.CPUControl)
    FindCPUPoint hTarget, X, Y
    Before = CPUClicks
    modMouseTests.Down hTarget, X, Y
    modMouseTests.Check GetCapture() = hTarget, "reloaded CPUView registers for capture " & CStr(Iteration)
    modMouseTests.Up hTarget, X, Y
    modMouseTests.Check CPUClicks = Before + 1 And LastCPUIndex = 1, "CPUView unload callback receives one valid region " & CStr(Iteration)
    modMouseTests.Check IsWindow(hTarget) = 0 And GetCapture() = 0, "CPUView can unload its form from Click " & CStr(Iteration)
  Next

  RestoreMainForm
  hTarget = modMouseTests.ControlWindow(frmMouseTests.TestButton)
  Before = frmMouseTests.ButtonClicks
  modMouseTests.Down hTarget
  modMouseTests.Up hTarget
  modMouseTests.Check frmMouseTests.ButtonClicks = Before + 1 And GetCapture() = 0, "unloading secondary forms preserves main-form registration"
  Exit Sub
Failed:
  ErrorNumber = Err.Number
  ErrorDescription = Err.Description
  On Error Resume Next
  Unload frmMouseLifetime
  RestoreMainForm
  On Error GoTo 0
  Err.Raise ErrorNumber, "TestMouseLifetime", ErrorDescription
End Sub

Private Sub OpenLifetimeForm()
  Load frmMouseLifetime
  frmMouseLifetime.CPUControl.AutoUpdate = False
  frmMouseLifetime.Show
  Call ShowWindow(frmMouseLifetime.hWnd, 5)
  Call SetForegroundWindow(frmMouseLifetime.hWnd)
  DoEvents
  If GetForegroundWindow() <> frmMouseLifetime.hWnd Then Err.Raise 5, , "Lifetime test form could not become foreground"
End Sub

Private Sub RestoreMainForm()
  frmMouseTests.Show
  Call SetForegroundWindow(frmMouseTests.hWnd)
  DoEvents
  If GetForegroundWindow() <> frmMouseTests.hWnd Then Err.Raise 5, , "Main test form could not regain foreground"
  MouseRefreshAll
End Sub

Private Sub FindCPUPoint(ByVal hWnd As Long, ByRef X As Long, ByRef Y As Long)
  Dim Sink As IMouseSink, Bounds As LifetimeRect, View As CPUView
  Set View = frmMouseLifetime.CPUControl
  Set Sink = View
  If GetClientRect(hWnd, Bounds) = 0 Then Err.Raise 5, , "CPU lifetime fixture has no client rectangle"
  For Y = 0 To Bounds.Bottom - 1
    For X = 0 To Bounds.Right - 1
      If Sink.HitTest(0, X, Y) = 1 Then
        Set Sink = Nothing
        Exit Sub
      End If
    Next
  Next
  Err.Raise 5, , "CPU lifetime fixture has no visible positive region"
End Sub
