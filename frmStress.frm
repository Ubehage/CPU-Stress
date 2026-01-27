VERSION 5.00
Begin VB.Form frmStress 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmStress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum StressStatus
  sIdle = 1
  sBusy = 2
End Enum

Dim WithEvents IdleTimer As CPUTimer
Attribute IdleTimer.VB_VarHelpID = -1

Dim InStressLoop As Boolean
Dim ExitStressLoop As Boolean

Public Sub Start()
  LockProcessToCPU (SharedMemOffset - 1)
  SetProcessPriority BELOW_NORMAL_PRIORITY_CLASS
  StartIdleTimer
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

Private Sub RunStressTest()
  InStressLoop = True
  Do
    StressLoop
    CheckAppMessage
  Loop Until ExitStressLoop = True
  InStressLoop = False
  ExitStressLoop = False
End Sub

Private Sub CheckAppMessage()
  Call ReadFromSharedMemory(False, SharedMemOffset)
  With SharedMemory.Instances(SharedMemOffset)
    Select Case .mCommand
      Case MEMMSG_RUNSTRESS
        .mCommand = 0
        .mStatus = MEMSTATUS_RUNNING
        Call WriteToSharedMemory(False)
        If InStressLoop = False Then
          RunStressTest
        End If
      Case MEMMSG_STOPSTRESS
        .mCommand = 0
        .mStatus = MEMSTATUS_IDLE
        Call WriteToSharedMemory(False)
        If InStressLoop = True Then ExitStressLoop = True
      Case MEMMSG_EXIT
        ExitNow = True
        If InStressLoop Then ExitStressLoop = True
        .mCommand = 0
        .mStatus = MEMSTATUS_EXITING
        Call WriteToSharedMemory(False)
    End Select
  End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
  KillIdleTimer
  UnloadAll
End Sub

Private Sub IdleTimer_Timer()
  IdleTimer.Enabled = False
  CheckAppMessage
  If ExitNow = True Then Unload Me
  If Not IdleTimer Is Nothing Then IdleTimer.Enabled = True
End Sub
