VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPUTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SetTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function KillTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long

Private m_TmrID As Long
Private m_hWnd As Long
Private m_Enabled As Boolean
Private m_Interval As Long

Private Const defEnabled As Boolean = False
Private Const defInterval As Long = 0

Public Event Timer()

Private Sub Class_Initialize()
   ' Set default values for members
   m_Enabled = defEnabled
   m_Interval = defInterval
   ' Get handle for main application window.
   m_hWnd = hWndMain()
End Sub

Private Sub Class_Terminate()
   ' Make sure timer is wiped out.
   Call KillTimer(m_hWnd, m_TmrID)
End Sub

' ****************************************************
'  Public Properties
' ****************************************************
Public Property Get Enabled() As Boolean
   ' Return cached setting.
   Enabled = m_Enabled
End Property

Public Property Let Enabled(ByVal NewEnabled As Boolean)
   ' Stop current timer, if already running.
   If m_Enabled Then
      Call KillTimer(m_hWnd, m_TmrID)
      m_TmrID = 0
   End If
   
   ' Cache passed value.
   m_Enabled = NewEnabled
   
   ' An interval of zero implicitly disables timer.
   If m_Interval > 0 Then
      ' Set new timer, if requested.
      If m_Enabled Then
         ' Pass pointer to Me so we can return event to this instance.
         m_TmrID = SetTimer(m_hWnd, ObjPtr(Me), m_Interval, AddressOf TimerProc)
      End If
   End If
End Property

Public Property Get Interval() As Long
   Interval = m_Interval
End Property

Public Property Let Interval(ByVal NewInterval As Long)
   If NewInterval >= 0 Then
      m_Interval = NewInterval
      Me.Enabled = m_Enabled
   End If
End Property

Public Sub RaiseTimer()
   RaiseEvent Timer
End Sub
