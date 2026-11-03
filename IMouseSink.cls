VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IMouseSink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Coordinates are signed client pixels. Zero means no actionable region.
Public Function HitTest(ByVal Token As Long, ByVal X As Long, ByVal Y As Long) As Long
End Function

Public Sub OnMouse(ByVal Token As Long, ByVal Kind As MouseEventKind, _
    ByVal Region As Long, ByVal Button As Integer, ByVal Buttons As Integer, _
    ByVal Shift As Integer, ByVal X As Long, ByVal Y As Long, _
    ByVal Reason As MouseCancelReason)
End Sub
