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

Dim WithEvents IdleTimer As CPUTimer
Attribute IdleTimer.VB_VarHelpID = -1

Dim InStressLoop As Boolean
Dim ExitStressLoop As Boolean

Dim HibernateTimer As Boolean
Dim doStartStress As Boolean
Private ResumeAfterHibernate As Boolean
Private StartupFailed As Boolean

Public Sub Start()
  On Error GoTo Failed
  LockProcessToCPU (SharedMemOffset - 1)
  StartIdleTimer
  Exit Sub
Failed:
  StartupFailed = True
  PublishWorkerStatus MEMSTATUS_FAILED
  StartIdleTimer
End Sub

Private Sub HibernationTrigger(IsHibernating As Boolean)
  If IsHibernating = True Then
    ResumeAfterHibernate = (InStressLoop Or doStartStress)
    doStartStress = False
    If InStressLoop Then ExitStressLoop = True
    HibernateTimer = True
    PublishWorkerStatus MEMSTATUS_HIBERNATING
  ElseIf IsHibernating = False Then
    doStartStress = ResumeAfterHibernate
    ResumeAfterHibernate = False
    HibernateTimer = False
    PublishWorkerStatus MEMSTATUS_IDLE
  End If
End Sub

Private Sub StartIdleTimer()
  KillIdleTimer
  Set IdleTimer = New CPUTimer
  IdleTimer.Interval = 100
  IdleTimer.Enabled = True
End Sub

Private Sub KillIdleTimer()
  If IdleTimer Is Nothing Then Exit Sub
  IdleTimer.Enabled = False
  Set IdleTimer = Nothing
End Sub

Private Sub RunStress()
  On Error GoTo Failed
  InStressLoop = True
  ExitStressLoop = False
  PublishWorkerStatus MEMSTATUS_RUNNING
  Do
    CheckAppMessage
    If ExitStressLoop Or ExitNow Then Exit Do
    Call StressLoop
    DoEvents2 Me.hWnd
    CheckAppMessage
    If IsProcessAlive(SharedMemory.Instances(0).mProcessID) = False Then
      ExitNow = True
      ExitStressLoop = True
    End If
  Loop Until ExitStressLoop = True
  InStressLoop = False
  ExitStressLoop = False
  If ExitNow Then
    PublishWorkerStatus MEMSTATUS_EXITING
  ElseIf HibernateTimer Then
    PublishWorkerStatus MEMSTATUS_HIBERNATING
  Else
    PublishWorkerStatus MEMSTATUS_IDLE
  End If
  Exit Sub
Failed:
  ' A worker reports an initialization/runtime error without a hidden modal
  ' dialog. The main window remains usable and shows a failed worker count.
  StressShutdown
  InStressLoop = False
  ExitStressLoop = False
  doStartStress = False
  PublishWorkerStatus MEMSTATUS_FAILED
End Sub

Private Sub CheckAppMessage()
    Select Case TakeWorkerCommand()
      Case MEMMSG_RUNSTRESS
        If Not StartupFailed Then
          HibernateTimer = False
          ResumeAfterHibernate = False
          If Not InStressLoop Then doStartStress = True
        End If
      Case MEMMSG_STOPSTRESS
        doStartStress = False
        ResumeAfterHibernate = False
        HibernateTimer = False
        PublishWorkerStatus MEMSTATUS_IDLE
        If InStressLoop = True Then ExitStressLoop = True
      Case MEMMSG_EXIT
        doStartStress = False
        PublishWorkerStatus MEMSTATUS_EXITING
        ExitNow = True
        If InStressLoop Then ExitStressLoop = True
      Case MEMMSG_HIBERNATE
        HibernationTrigger True
      Case MEMMSG_RESUME
        HibernationTrigger False
    End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
  If InStressLoop Then
    Cancel = True
    ExitNow = True
    ExitStressLoop = True
    Exit Sub
  End If
  KillIdleTimer
  StressShutdown
  ClearSharedMemoryIndex SharedMemOffset
  UnloadAll
End Sub

Private Sub IdleTimer_Timer()
  IdleTimer.Enabled = False
  CheckAppMessage
  If Not IsProcessAlive(SharedMemory.Instances(0).mProcessID) Then ExitNow = True
  If ExitNow = True Then
    Unload Me
    Exit Sub
  End If
  If HibernateTimer = False Then
    If doStartStress = True Then
      doStartStress = False
      RunStress
    End If
  End If
  If ExitNow Then
    Unload Me
    Exit Sub
  End If
  If Not IdleTimer Is Nothing Then IdleTimer.Enabled = True
End Sub
