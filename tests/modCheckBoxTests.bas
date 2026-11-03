Attribute VB_Name = "modCheckBoxTests"
Option Explicit

Public Sub TestCheckBoxes()
  Dim hCheck As Long, hSecond As Long, Before As Long, SecondBefore As Long
  frmMouseTests.CheckAction = vbNullString
  frmMouseTests.TestCheck.Enabled = True
  frmMouseTests.TestCheck2.Enabled = True
  ChangedByCode = True
  frmMouseTests.TestCheck.Value = vbUnchecked
  frmMouseTests.TestCheck2.Value = vbUnchecked
  ChangedByCode = False
  hCheck = ControlWindow(frmMouseTests.TestCheck)
  hSecond = ControlWindow(frmMouseTests.TestCheck2)
  Check hCheck <> 0 And hSecond <> 0 And hCheck <> hSecond, "distinct real CheckBox windows"
  Before = frmMouseTests.CheckClicks
  SecondBefore = frmMouseTests.Check2Clicks

  modMouseTests.Move hCheck, 10, 10
  Check GetCapture() = 0, "CheckBox hover does not capture"
  Check PixelAt(hCheck, 5, 5) = COLOR_BUTTON_HOVER, "unchecked CheckBox paints hover"
  Down hCheck
  Check GetCapture() = hCheck, "CheckBox press owns capture"
  Check PixelAt(hCheck, 5, 5) = COLOR_BUTTON_PRESSED, "unchecked CheckBox paints pressed"
  modMouseTests.Move hCheck, -10, 10, 1
  Check PixelAt(hCheck, 5, 5) = COLOR_BUTTON_PRESSED, "CheckBox retains pressed appearance outside"
  Up hCheck, -10, 10
  Check frmMouseTests.TestCheck.Value = vbUnchecked And frmMouseTests.CheckClicks = Before, "CheckBox release outside does not toggle or click"
  Check GetCapture() = 0, "CheckBox release outside clears capture"

  Down hCheck
  Up hCheck
  Check frmMouseTests.TestCheck.Value = vbChecked And frmMouseTests.CheckClicks = Before + 1, "CheckBox toggles and raises exactly one Click"
  Check GetCapture() = 0, "CheckBox successful click leaves capture free"
  Check PixelAt(hCheck, 5, 5) = COLOR_GREEN_HOVER, "checked CheckBox paints its hovered mark"
  Down hCheck
  modMouseTests.Move hCheck, -10, 10, 1
  modMouseTests.Move hCheck, 10, 10, 1
  Up hCheck
  Check frmMouseTests.TestCheck.Value = vbUnchecked And frmMouseTests.CheckClicks = Before + 2, "CheckBox drag out and back toggles once"

  Before = frmMouseTests.CheckClicks
  Down hCheck
  SetCapture hSecond
  Check GetCapture() = hSecond, "CheckBox capture loss does not release the new owner"
  ReleaseCapture
  Up hCheck
  Check frmMouseTests.TestCheck.Value = vbUnchecked And frmMouseTests.CheckClicks = Before, "CheckBox capture theft cancels the pending toggle"

  Down hCheck
  frmMouseTests.TestCheck.Enabled = False
  Check GetCapture() = 0, "disabling pressed CheckBox releases capture"
  Check PixelAt(hCheck, 5, 5) = COLOR_BACKGROUND_DISABLED, "unchecked disabled CheckBox paints disabled"
  Up hCheck
  Check frmMouseTests.TestCheck.Value = vbUnchecked And frmMouseTests.CheckClicks = Before, "disabled CheckBox cannot complete a toggle"
  frmMouseTests.TestCheck.Enabled = True

  Down hCheck
  MouseCancelAll mcDeactivated
  Check GetCapture() = 0, "deactivation releases CheckBox capture"
  Up hCheck
  Check frmMouseTests.TestCheck.Value = vbUnchecked And frmMouseTests.CheckClicks = Before, "deactivation cancels CheckBox toggle"

  frmMouseTests.CheckAction = "disable"
  Down hCheck
  Up hCheck
  Check frmMouseTests.CheckClicks = Before + 1 And frmMouseTests.TestCheck.Value = vbChecked, "CheckBox value is updated before its Click callback"
  Check Not frmMouseTests.TestCheck.Enabled And GetCapture() = 0, "CheckBox Click callback can disable without recapture"
  frmMouseTests.CheckAction = vbNullString
  frmMouseTests.TestCheck.Enabled = True

  Before = frmMouseTests.CheckClicks
  frmMouseTests.CheckAction = "hide"
  Down hCheck
  Up hCheck
  Check Not frmMouseTests.TestCheck.Visible And GetCapture() = 0, "CheckBox Click callback can hide without recapture"
  Check frmMouseTests.CheckClicks = Before + 1 And frmMouseTests.TestCheck.Value = vbUnchecked, "hidden CheckBox receives one completed toggle"
  frmMouseTests.CheckAction = vbNullString
  frmMouseTests.TestCheck.Visible = True
  DoEvents
  hCheck = ControlWindow(frmMouseTests.TestCheck)
  Down hCheck
  Up hCheck
  Check frmMouseTests.CheckClicks = Before + 2 And frmMouseTests.TestCheck.Value = vbChecked, "CheckBox re-registers after hide/show"

  ' Reset only the first value without raising a Click; then exercise instance two.
  ChangedByCode = True
  frmMouseTests.TestCheck.Value = vbUnchecked
  ChangedByCode = False
  Before = frmMouseTests.CheckClicks
  Down hSecond
  Up hSecond
  Check frmMouseTests.TestCheck2.Value = vbChecked And frmMouseTests.Check2Clicks = SecondBefore + 1, "second CheckBox toggles independently"
  Check frmMouseTests.TestCheck.Value = vbUnchecked And frmMouseTests.CheckClicks = Before, "second CheckBox leaves first instance untouched"

  frmMouseTests.TestCheck.Value = vbChecked
  Check frmMouseTests.CheckClicks = Before + 1, "programmatic CheckBox Value change retains Click behavior"
  frmMouseTests.TestCheck.Value = vbChecked
  Check frmMouseTests.CheckClicks = Before + 1, "unchanged CheckBox Value does not raise Click"
  ChangedByCode = True
  frmMouseTests.TestCheck.Value = vbUnchecked
  Check frmMouseTests.CheckClicks = Before + 1 And ChangedByCode, "ChangedByCode suppresses programmatic CheckBox Click"
  ChangedByCode = False
  frmMouseTests.TestCheck.Value = vbChecked
  Check frmMouseTests.CheckClicks = Before + 2, "CheckBox Value notifications resume after suppression"
End Sub
