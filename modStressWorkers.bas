Attribute VB_Name = "modStressWorkers"
Option Explicit

Private Type STARTUPINFO
  cb As Long
  Reserved As Long
  Desktop As Long
  Title As Long
  X As Long
  Y As Long
  Width As Long
  Height As Long
  XChars As Long
  YChars As Long
  Fill As Long
  Flags As Long
  Show As Integer
  ReservedBytes As Integer
  ReservedPointer As Long
  StdInput As Long
  StdOutput As Long
  StdError As Long
End Type

Private Type PROCESS_INFORMATION
  Process As Long
  Thread As Long
  ProcessID As Long
  ThreadID As Long
End Type

Private Declare Function CreateProcessW Lib "kernel32" (ByVal Application As Long, ByVal CommandLine As Long, ByVal ProcessAttributes As Long, ByVal ThreadAttributes As Long, ByVal Inherit As Long, ByVal Flags As Long, ByVal Environment As Long, ByVal Directory As Long, ByRef Startup As STARTUPINFO, ByRef Info As PROCESS_INFORMATION) As Long
Private Declare Function CreateJobObjectW Lib "kernel32" (ByVal Attributes As Long, ByVal Name As Long) As Long
Private Declare Function SetInformationJobObject Lib "kernel32" (ByVal Job As Long, ByVal InfoClass As Long, Info As Any, ByVal Size As Long) As Long
Private Declare Function AssignProcessToJobObject Lib "kernel32" (ByVal Job As Long, ByVal Process As Long) As Long
Private Declare Function ResumeThread Lib "kernel32" (ByVal Thread As Long) As Long
Private Declare Function TerminateProcess Lib "kernel32" (ByVal Process As Long, ByVal Code As Long) As Long

Private m_WorkerJob As Long

' The main process owns the job. Even an abnormal UI exit stops its workers.
' Start suspended so no worker can run before ownership and slot publication.
Public Sub LaunchOwnedStressWorker(ByVal Executable As String, ByVal CommandLine As String, ByVal Index As Long)
  Dim Startup As STARTUPINFO, Child As PROCESS_INFORMATION
  Dim Limits(0 To 111) As Byte, Flags As Long, Detail As String
  On Error GoTo Failed
  If m_WorkerJob = 0 Then
    m_WorkerJob = CreateJobObjectW(0, 0)
    If m_WorkerJob = 0 Then Err.Raise 5, , "Cannot create the worker job."
    Flags = &H2000& ' JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE
    CopyMemory Limits(16), Flags, 4
    If SetInformationJobObject(m_WorkerJob, 9, Limits(0), 112) = 0 Then
      StressWorkersShutdown
      Err.Raise 5, , "Cannot configure worker lifetime management."
    End If
  End If
  Startup.cb = LenB(Startup)
  Startup.Flags = 1 ' STARTF_USESHOWWINDOW: keep worker windows hidden
  ' CREATE_NO_WINDOW | CREATE_SUSPENDED | NORMAL_PRIORITY_CLASS
  If CreateProcessW(StrPtr(Executable), StrPtr(CommandLine), 0, 0, 0, &H8000024, 0, 0, Startup, Child) = 0 Then Err.Raise 5, , "Cannot create the stress worker."
  If AssignProcessToJobObject(m_WorkerJob, Child.Process) = 0 Then Err.Raise 5, , "Cannot attach the stress worker to its owner."
  With SharedMemory.Instances(Index)
    .mProcessID = Child.ProcessID
    .mAssignedCore = Index
    .mCommand = MEMMSG_RUNSTRESS
    .mStatus = 0
  End With
  If Not WriteToSharedMemory(False, Index) Then Err.Raise 5, , "Cannot publish the stress worker control slot."
  If ResumeThread(Child.Thread) = -1 Then Err.Raise 5, , "Cannot start the stress worker."
  Call CloseHandle(Child.Thread)
  Call CloseHandle(Child.Process)
  Exit Sub
Failed:
  Detail = Err.Description
  If Child.Process <> 0 Then
    Call TerminateProcess(Child.Process, 2)
    Call WaitForSingleObject(Child.Process, 1000)
  End If
  If Child.Thread <> 0 Then Call CloseHandle(Child.Thread)
  If Child.Process <> 0 Then Call CloseHandle(Child.Process)
  ClearSharedMemoryIndex Index
  Err.Raise vbObjectError + 2402, "LaunchOwnedStressWorker", Detail
End Sub

Public Sub StressWorkersShutdown()
  If m_WorkerJob <> 0 Then
    Call CloseHandle(m_WorkerJob)
    m_WorkerJob = 0
  End If
End Sub
