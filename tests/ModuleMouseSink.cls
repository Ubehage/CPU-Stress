VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "ModuleMouseSink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IMouseSink

Private Function IMouseSink_HitTest(ByVal Token As Long, ByVal X As Long, ByVal Y As Long) As Long
  IMouseSink_HitTest = 1
End Function

Private Sub IMouseSink_OnMouse(ByVal Token As Long, ByVal Kind As MouseEventKind, _
    ByVal Region As Long, ByVal Button As Integer, ByVal Buttons As Integer, _
    ByVal Shift As Integer, ByVal X As Long, ByVal Y As Long, ByVal Reason As MouseCancelReason)
  modMouseTests.HandleModuleMouse Token, Kind
End Sub
