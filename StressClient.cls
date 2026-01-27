VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StressClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Enum StressStatus
  sIdle = 1
  sBusy = 2
End Enum

Dim WithEvents IdleTimer As CPUTimer
Attribute IdleTimer.VB_VarHelpID = -1

Public Sub Start()
  If IsRunningInIDE = False Then StartIdleTimer
End Sub

Private Sub StartIdleTimer()
  KillIdleTimer
  Set IdleTimer = New CPUTimer
  IdleTimer.Interval = 200
  IdleTimer.Enabled = True
End Sub

Private Sub KillIdleTimer()
  If IdleTimer Is Nothing Then Exit Sub
  IdleTimer.Enabled = False
  Set IdleTimer = Nothing
End Sub

Private Sub CheckAppMessage()
  Call ReadFromSharedMemory(False, SharedMemOffset)
  
End Sub
