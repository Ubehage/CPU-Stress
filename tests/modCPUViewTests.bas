Attribute VB_Name = "modCPUViewTests"
Option Explicit

Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, ByRef Area As RECT) As Long
Private Declare Function EnableWindow Lib "user32" (ByVal hWnd As Long, ByVal Enabled As Long) As Long

Private Function CorePoint(ByVal View As CPUView, ByVal hWnd As Long, ByVal Region As Long, ByRef Point As POINTAPI) As Boolean
  Dim Sink As IMouseSink, Area As RECT, X As Long, Y As Long
  Set Sink = View
  GetClientRect hWnd, Area
  For Y = 0 To Area.Bottom - 1
    For X = 0 To Area.Right - 1
      If Sink.HitTest(0, X, Y) = Region Then
        ' Use an interior pixel so color assertions do not sample the border.
        Point.X = X + 3
        Point.Y = Y + 3
        CorePoint = (Sink.HitTest(0, Point.X, Point.Y) = Region)
        Exit Function
      End If
    Next
  Next
End Function

Public Sub TestCPUViews()
  Dim hCPU As Long, hCPU2 As Long, Sink As IMouseSink, View As CPUView
  Dim First As POINTAPI, Second As POINTAPI, Other As POINTAPI
  Dim Before As Long, OtherBefore As Long, X As Long, Y As Long
  Dim OldLeft As Single, OldWidth As Single, OldHeight As Single
  hCPU = modMouseTests.ControlWindow(frmMouseTests.TestCPU)
  hCPU2 = modMouseTests.ControlWindow(frmMouseTests.TestCPU2)
  modMouseTests.Check hCPU <> 0 And hCPU2 <> 0 And hCPU <> hCPU2 And hCPU <> frmMouseTests.hWnd, "find independent real CPUView windows"
  modMouseTests.Check CorePoint(frmMouseTests.TestCPU, hCPU, 1, First), "CPUView exposes first valid core region"
  modMouseTests.Check CorePoint(frmMouseTests.TestCPU, hCPU, 2, Second), "CPUView exposes second valid core region"
  modMouseTests.Check CorePoint(frmMouseTests.TestCPU2, hCPU2, 4, Other), "second CPUView exposes last core region"
  Set View = frmMouseTests.TestCPU
  Set Sink = View
  modMouseTests.Check Sink.HitTest(0, 10, 10) = 0, "CPUView heading is outside actionable regions"
  modMouseTests.Check Sink.HitTest(0, -1, First.Y) = 0, "CPUView rejects negative client coordinates"
  X = First.X
  Do While Sink.HitTest(0, X, First.Y) = 1
    X = X + 1
    If X > 500 Then Exit Do
  Loop
  modMouseTests.Check X < 500 And Sink.HitTest(0, X, First.Y) = 0, "CPUView right boundary leads to empty spacing"
  Y = First.Y
  Do While Sink.HitTest(0, First.X, Y) = 1
    Y = Y + 1
    If Y > 500 Then Exit Do
  Loop
  modMouseTests.Check Y < 500 And Sink.HitTest(0, First.X, Y) = 0, "CPUView bottom boundary leads to empty spacing"

  modMouseTests.Move hCPU, First.X, First.Y
  modMouseTests.Check GetCapture() = 0, "CPUView passive hover never captures"
  modMouseTests.Check modMouseTests.PixelAt(hCPU, First.X, First.Y) = COLOR_BUTTON_HOVER, "CPUView paints hovered core"
  modMouseTests.Move hCPU, Second.X, Second.Y
  modMouseTests.Check modMouseTests.PixelAt(hCPU, First.X, First.Y) = COLOR_BACKGROUND, "CPUView clears old core highlight on region transition"
  modMouseTests.Check modMouseTests.PixelAt(hCPU, Second.X, Second.Y) = COLOR_BUTTON_HOVER, "CPUView paints new core on region transition"

  Before = frmMouseTests.CPUClicks
  modMouseTests.Down hCPU, First.X, First.Y
  modMouseTests.Check GetCapture() = hCPU, "CPUView press captures its own window"
  modMouseTests.Move hCPU, Second.X, Second.Y, 1
  modMouseTests.Check modMouseTests.PixelAt(hCPU, First.X, First.Y) = COLOR_BUTTON_PRESSED, "CPUView retains original core pressed appearance"
  modMouseTests.Check modMouseTests.PixelAt(hCPU, Second.X, Second.Y) = COLOR_BUTTON_HOVER, "CPUView tracks hover independently of pressed core"
  modMouseTests.Up hCPU, Second.X, Second.Y
  modMouseTests.Check frmMouseTests.CPUClicks = Before, "CPUView release on another core does not click"
  modMouseTests.Check modMouseTests.PixelAt(hCPU, First.X, First.Y) = COLOR_BACKGROUND, "CPUView release clears original pressed appearance"

  modMouseTests.Down hCPU, First.X, First.Y
  modMouseTests.Move hCPU, -5, First.Y, 1
  modMouseTests.Check modMouseTests.PixelAt(hCPU, First.X, First.Y) = COLOR_BUTTON_PRESSED, "CPUView retains press outside its window"
  modMouseTests.Move hCPU, First.X, First.Y, 1
  modMouseTests.Up hCPU, First.X, First.Y
  modMouseTests.Check frmMouseTests.CPUClicks = Before + 1 And frmMouseTests.LastCPUIndex = 1, "CPUView drag out/back emits one indexed Click"
  modMouseTests.Check GetCapture() = 0, "CPUView releases capture before Click returns"
  Before = frmMouseTests.CPUClicks
  modMouseTests.Down hCPU, First.X, First.Y
  modMouseTests.Up hCPU, -5, First.Y
  modMouseTests.Check frmMouseTests.CPUClicks = Before, "CPUView release outside does not click"

  modMouseTests.Down hCPU, 10, 10
  modMouseTests.Up hCPU, 10, 10
  modMouseTests.Check frmMouseTests.CPUClicks = Before And frmMouseTests.InvalidCPUClicks = 0, "CPUView empty-space click never emits Click(0)"
  modMouseTests.Down hCPU, 10, 10
  modMouseTests.Up hCPU, First.X, First.Y
  modMouseTests.Check frmMouseTests.CPUClicks = Before, "CPUView press in empty space cannot arm a core"
  modMouseTests.Down hCPU, First.X, First.Y
  modMouseTests.Up hCPU, 10, 10
  modMouseTests.Check frmMouseTests.CPUClicks = Before, "CPUView core-to-empty release does not click"

  modMouseTests.Down hCPU, First.X, First.Y
  SetCapture hCPU2
  modMouseTests.Check GetCapture() = hCPU2, "CPUView capture loss preserves the new owner"
  modMouseTests.Check modMouseTests.PixelAt(hCPU, First.X, First.Y) = COLOR_BACKGROUND, "CPUView capture loss clears pressed and hover drawing"
  ReleaseCapture
  modMouseTests.Up hCPU, First.X, First.Y
  modMouseTests.Check frmMouseTests.CPUClicks = Before, "CPUView capture loss cancels pending Click"

  modMouseTests.Down hCPU, First.X, First.Y
  EnableWindow hCPU, 0
  modMouseTests.Check GetCapture() = 0, "CPUView native disable releases capture"
  modMouseTests.Up hCPU, First.X, First.Y
  modMouseTests.Check frmMouseTests.CPUClicks = Before, "CPUView native disable suppresses Click"
  EnableWindow hCPU, 1
  modMouseTests.Down hCPU, First.X, First.Y
  MouseCancelAll mcDeactivated
  modMouseTests.Up hCPU, First.X, First.Y
  modMouseTests.Check frmMouseTests.CPUClicks = Before And GetCapture() = 0, "CPUView deactivation cancels gesture"

  OtherBefore = frmMouseTests.CPU2Clicks
  modMouseTests.Down hCPU2, Other.X, Other.Y
  modMouseTests.Up hCPU2, Other.X, Other.Y
  modMouseTests.Check frmMouseTests.CPU2Clicks = OtherBefore + 1 And frmMouseTests.LastCPU2Index = 4 And frmMouseTests.CPUClicks = Before, "CPUView instances deliver independent indexed clicks"

  modMouseTests.Down hCPU, First.X, First.Y
  frmMouseTests.TestCPU.Visible = False
  modMouseTests.Check GetCapture() = 0, "CPUView hide during press releases capture"
  frmMouseTests.TestCPU.Visible = True
  DoEvents
  hCPU = modMouseTests.ControlWindow(frmMouseTests.TestCPU)
  modMouseTests.Up hCPU, First.X, First.Y
  modMouseTests.Check frmMouseTests.CPUClicks = Before, "CPUView hide/show does not revive an old press"
  modMouseTests.Down hCPU, First.X, First.Y
  modMouseTests.Up hCPU, First.X, First.Y
  modMouseTests.Check frmMouseTests.CPUClicks = Before + 1, "CPUView re-registers after hide/show"

  frmMouseTests.CPUAction = "hide"
  Before = frmMouseTests.CPUClicks
  modMouseTests.Down hCPU, Second.X, Second.Y
  modMouseTests.Up hCPU, Second.X, Second.Y
  modMouseTests.Check frmMouseTests.CPUClicks = Before + 1 And Not frmMouseTests.TestCPU.Visible And GetCapture() = 0, "CPUView Click callback can hide its source without recapture"
  frmMouseTests.CPUAction = vbNullString
  frmMouseTests.TestCPU.Visible = True
  DoEvents
  hCPU = modMouseTests.ControlWindow(frmMouseTests.TestCPU)

  modMouseTests.Move hCPU, First.X, First.Y
  OldLeft = frmMouseTests.TestCPU.Left
  frmMouseTests.TestCPU.Left = OldLeft + 15
  modMouseTests.Check modMouseTests.PixelAt(hCPU, First.X, First.Y) = COLOR_BACKGROUND, "moving CPUView under stationary pointer clears hover"
  frmMouseTests.TestCPU.Left = OldLeft
  OldWidth = frmMouseTests.TestCPU.Width
  frmMouseTests.TestCPU.Width = OldWidth - 20
  modMouseTests.Check CorePoint(frmMouseTests.TestCPU, hCPU, 2, Second), "CPUView rebuilds region geometry after resize"
  Before = frmMouseTests.CPUClicks
  modMouseTests.Down hCPU, Second.X, Second.Y
  modMouseTests.Up hCPU, Second.X, Second.Y
  modMouseTests.Check frmMouseTests.CPUClicks = Before + 1 And frmMouseTests.LastCPUIndex = 2, "CPUView clicks follow resized geometry"
  frmMouseTests.TestCPU.Width = OldWidth
  modMouseTests.Check CorePoint(frmMouseTests.TestCPU, hCPU, 1, First), "CPUView restores resized region geometry"
  modMouseTests.Move hCPU, First.X, First.Y
  frmMouseTests.TestCPU.Refresh True
  modMouseTests.Check modMouseTests.PixelAt(hCPU, First.X, First.Y) = COLOR_BUTTON_HOVER, "CPUView full refresh preserves current hover drawing"

  OldHeight = frmMouseTests.TestCPU.Height
  frmMouseTests.TestCPU.AutoSizeView = NoAutoSize
  modMouseTests.Check frmMouseTests.TestCPU.Width = OldWidth And frmMouseTests.TestCPU.Height = OldHeight, "CPUView NoAutoSize preserves control dimensions"
  modMouseTests.Check CorePoint(frmMouseTests.TestCPU, hCPU, 1, First) And Sink.HitTest(0, First.X + 20, First.Y) = 0, "CPUView NoAutoSize uses fixed-width core regions"
  Before = frmMouseTests.CPUClicks
  modMouseTests.Down hCPU, First.X, First.Y
  modMouseTests.Up hCPU, First.X, First.Y
  modMouseTests.Check frmMouseTests.CPUClicks = Before + 1 And frmMouseTests.LastCPUIndex = 1, "CPUView NoAutoSize click matches its drawing"

  frmMouseTests.TestCPU.AutoSizeView = FixedSize
  modMouseTests.Check frmMouseTests.TestCPU.Width < OldWidth And frmMouseTests.TestCPU.Height <= OldHeight, "CPUView FixedSize fits the fixed core layout"
  modMouseTests.Check CorePoint(frmMouseTests.TestCPU, hCPU, 4, Other), "CPUView FixedSize keeps last core inside visible client area"
  Before = frmMouseTests.CPUClicks
  modMouseTests.Down hCPU, Other.X, Other.Y
  modMouseTests.Up hCPU, Other.X, Other.Y
  modMouseTests.Check frmMouseTests.CPUClicks = Before + 1 And frmMouseTests.LastCPUIndex = 4, "CPUView FixedSize delivers final core index"
  frmMouseTests.TestCPU.AutoSizeView = ViewAutoSize
  frmMouseTests.TestCPU.Width = OldWidth
  frmMouseTests.TestCPU.Height = OldHeight
  modMouseTests.Check frmMouseTests.InvalidCPUClicks = 0, "all CPUView notifications use valid positive indices"
End Sub
