Attribute VB_Name = "modMonitorTests"
Option Explicit
Private Declare Sub Sleep Lib "kernel32" (ByVal Milliseconds As Long)
Private Declare Function QueryPerformanceCounter Lib "kernel32" (ByRef Counter As Currency) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (ByRef Frequency As Currency) As Long
Private Declare Sub ExitProcess Lib "kernel32" (ByVal ExitCode As Long)
Private LogFile As Integer
Private Checks As Long
Private Failures As Long

Public Sub Main()
  Dim View As CPUView, i As Long
  On Error GoTo Failed
  LogFile = FreeFile
  Open App.Path & "\monitor-results.log" For Output As #LogFile
  ' Load the existing hidden fixture; no window activation or pointer movement.
  Load frmMouseTests
  Set View = frmMouseTests.TestCPU
  Check Not View.MonitorTestHasData(), "new control has no stale measurements"
  Sample View, "first"
  Check View.MonitorTestBaseline(), "first OS snapshot establishes baseline"
  Check Not View.MonitorTestHasData(), "first snapshot is not a percentage yet"
  Check Not View.MonitorTestUnavailable(), "real query resolves and succeeds"
  Sleep 120
  Sample View, "second"
  Check View.MonitorTestHasData(), "second OS snapshot produces load data"
  Check Not View.MonitorTestUnavailable(), "second snapshot remains available"
  CheckLoads View

  View.MonitorTestReset
  Check Not View.MonitorTestHasData() And Not View.MonitorTestBaseline(), "monitor reset discards old interval"
  Sample View, "reset-first"
  Check Not View.MonitorTestHasData() And View.MonitorTestBaseline(), "reset requires fresh baseline"
  Sleep 120
  Sample View, "reset-second"
  Check View.MonitorTestHasData(), "reset recovers after second fresh sample"
  CheckLoads View

  View.MonitorTestDisableQuery
  Sample View, "unavailable"
  Check View.MonitorTestUnavailable(), "missing native API marks load unavailable"
  Check Not View.MonitorTestHasData() And Not View.MonitorTestBaseline(), "missing API clears data and baseline"
  For i = 0 To TotalCores - 1
    Check View.MonitorTestLoad(i) = 0, "missing API clears core " & CStr(i)
  Next i
  View.MonitorTestRestoreQuery
  Sample View, "recovery-first"
  Check Not View.MonitorTestUnavailable() And Not View.MonitorTestHasData(), "query recovery starts with a fresh baseline"
  Sleep 120
  Sample View, "recovery-second"
  Check View.MonitorTestHasData(), "query recovery produces fresh load"
  CheckLoads View
  GoTo Cleanup
Failed:
  Failures = Failures + 1
  Print #LogFile, "UNEXPECTED: " & CStr(Err.Number) & " " & Err.Description
Cleanup:
  On Error Resume Next
  Set View = Nothing
  Unload frmMouseTests
  Print #LogFile, CStr(Checks) & " checks; " & CStr(Failures) & " failures."
  Close #LogFile
  If Failures <> 0 Then ExitProcess 1
  ExitProcess 0
End Sub

Private Sub Sample(ByVal View As CPUView, ByVal Label As String)
  Dim Frequency As Currency, Started As Currency, Finished As Currency, Elapsed As Double
  QueryPerformanceFrequency Frequency
  QueryPerformanceCounter Started
  View.MonitorTestSample
  QueryPerformanceCounter Finished
  Elapsed = (CDbl(Finished) - CDbl(Started)) * 1000# / CDbl(Frequency)
  Print #LogFile, "TIMING " & Label & "=" & Trim$(Str$(Elapsed)) & " ms"
  Check Elapsed < 250#, Label & " sample stays below 250 ms"
End Sub

Private Sub CheckLoads(ByVal View As CPUView)
  Dim i As Long, Value As Double
  For i = 0 To TotalCores - 1
    Value = View.MonitorTestLoad(i)
    Print #LogFile, "LOAD " & CStr(i) & "=" & Trim$(Str$(Value))
    Check Value >= 0# And Value <= 100#, "core " & CStr(i) & " has finite 0..100 load"
  Next i
End Sub

Private Sub Check(ByVal Condition As Boolean, ByVal Label As String)
  Checks = Checks + 1
  If Condition Then
    Print #LogFile, "PASS " & Label
  Else
    Failures = Failures + 1
    Print #LogFile, "FAIL " & Label
  End If
End Sub
