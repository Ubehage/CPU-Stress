VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "TestMouseSink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IMouseSink

Public Trace As String
Public Clicks As Long
Public Cancels As Long
Public LastX As Long
Public LastY As Long
Public LastReason As MouseCancelReason
Public ActionKind As MouseEventKind
Public Action As String
Public TargetWindow As Long
Public FailHitTest As Boolean

Private Function IMouseSink_HitTest(ByVal Token As Long, ByVal X As Long, ByVal Y As Long) As Long
  If FailHitTest Then Err.Raise 5, , "Deliberate hit-test failure"
  If X < 50 Then
    IMouseSink_HitTest = 1
  ElseIf X < 100 Then
    IMouseSink_HitTest = 2
  End If
End Function

Private Sub IMouseSink_OnMouse(ByVal Token As Long, ByVal Kind As MouseEventKind, _
    ByVal Region As Long, ByVal Button As Integer, ByVal Buttons As Integer, _
    ByVal Shift As Integer, ByVal X As Long, ByVal Y As Long, ByVal Reason As MouseCancelReason)
  Dim PendingAction As String
  Trace = Trace & CStr(Kind) & ":" & CStr(Region) & ","
  LastX = X
  LastY = Y
  If Kind = meClick Then Clicks = Clicks + 1
  If Kind = meCancel Then
    Cancels = Cancels + 1
    LastReason = Reason
  End If
  If Kind <> ActionKind Then Exit Sub
  PendingAction = Action
  Action = vbNullString
  Select Case PendingAction
    Case "disable": MouseSetEnabled Token, False
    Case "unregister": MouseUnregister Token
    Case "cancel": MouseCancel Token, mcRequested
    Case "hide": Call ShowWindow(TargetWindow, 0)
    Case "destroy": Call DestroyWindow(TargetWindow)
    Case "reenter": StartNestedPress TargetWindow
    Case "move": MoveDuringCallback TargetWindow
    Case "error": Err.Raise 5, , "Deliberate recipient failure"
  End Select
End Sub
