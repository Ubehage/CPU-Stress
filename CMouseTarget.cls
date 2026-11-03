VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CMouseTarget"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Internal record. A dispatch keeps this object alive after unregistering;
' Active and Revision prevent it from notifying a removed or changed target.
Public Token As Long
Public hWnd As Long
Public Sink As IMouseSink
Public Active As Boolean
Public Enabled As Boolean
Public Revision As Long
Public HoverRegion As Long
Public PressRegion As Long
Public Tracking As Boolean
Public HitTesting As Boolean
Public DefaultUpDepth As Long
Public X As Long
Public Y As Long
Public Buttons As Integer
Public Shift As Integer
