Attribute VB_Name = "modWorkerTests"
Option Explicit

Public Type TestSlot
    mProcessID As Long
    mAssignedCore As Long
    mCommand As Long
    mStatus As Long
End Type
Public Type TestSlots
    Instances(0 To 3) As TestSlot
End Type
Public SharedMemory As TestSlots
Public Const MEMMSG_RUNSTRESS As Long = 5
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Count As Long)
Public Declare Function CloseHandle Lib "kernel32" (ByVal Handle As Long) As Long
Public Declare Function WaitForSingleObject Lib "kernel32" (ByVal Handle As Long, ByVal Milliseconds As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal Access As Long, ByVal Inherit As Long, ByVal ProcessID As Long) As Long
Private Declare Function GetPriorityClass Lib "kernel32" (ByVal Process As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal Milliseconds As Long)
Private Declare Sub ExitProcess Lib "kernel32" (ByVal Code As Long)
Private Checks As Long
Private Failures As Long
Private LogFile As Integer
Private Published As Boolean

Public Sub Main()
    Dim Executable As String, Child As Long, Index As Long, Raised As Boolean, Detail As String, ErrorFile As Integer
    On Error GoTo Failed
    If Trim$(Command$) = "/idle" Then
        Sleep 10000
        ExitProcess 0
    End If
    Executable = App.Path & "\" & App.EXEName & ".exe"
    If Trim$(Command$) = "/owner" Then
        LaunchOwnedStressWorker Executable, Chr$(34) & Executable & Chr$(34) & " /idle", 1
        LogFile = FreeFile
        Open App.Path & "\owner-child.pid" For Output As #LogFile
        Print #LogFile, CStr(SharedMemory.Instances(1).mProcessID)
        Close #LogFile
        Sleep 10000
        StressWorkersShutdown
        ExitProcess 0
    End If
    LogFile = FreeFile
    Open App.Path & "\worker-results.log" For Output As #LogFile
    For Index = 1 To 2
        Published = False
        LaunchOwnedStressWorker Executable, Chr$(34) & Executable & Chr$(34) & " /idle", Index
        Check Published, "worker slot published"
        Check SharedMemory.Instances(Index).mProcessID > 0, "worker PID published"
        Check SharedMemory.Instances(Index).mAssignedCore = Index, "worker assignment published"
        Check SharedMemory.Instances(Index).mCommand = MEMMSG_RUNSTRESS, "run command published"
        Child = OpenProcess(&H100400, 0, SharedMemory.Instances(Index).mProcessID)
        Check Child <> 0, "owned child is accessible"
        If Child <> 0 Then
            Check GetPriorityClass(Child) = &H20&, "owned child starts at Normal priority"
            Check WaitForSingleObject(Child, 0) = 258, "owned child remains alive until shutdown"
            StressWorkersShutdown
            Check WaitForSingleObject(Child, 1500) = 0, "job shutdown terminates the child"
            CloseHandle Child
            Child = 0
        End If
        StressWorkersShutdown
    Next Index
    SharedMemory.Instances(3).mProcessID = -1
    On Error Resume Next
    LaunchOwnedStressWorker App.Path & "\missing-worker.exe", "missing-worker.exe", 3
    Raised = (Err.Number <> 0)
    Err.Clear
    On Error GoTo Failed
    Check Raised, "failed launch reports an error"
    Check SharedMemory.Instances(3).mProcessID = 0, "failed launch releases its slot"
Cleanup:
    On Error Resume Next
    StressWorkersShutdown
    If Child <> 0 Then CloseHandle Child
    Print #LogFile, CStr(Checks) & " checks; " & CStr(Failures) & " failures."
    Close #LogFile
    If Failures > 0 Then ExitProcess 1
    ExitProcess 0
Failed:
    Detail = Err.Description
    Failures = Failures + 1
    On Error Resume Next
    ErrorFile = FreeFile
    Open App.Path & "\worker-error.log" For Append As #ErrorFile
    Print #ErrorFile, "Command=" & Command$ & "; error=" & Detail
    Close #ErrorFile
    If LogFile <> 0 Then Print #LogFile, "UNEXPECTED: " & Detail
    Resume Cleanup
End Sub

Public Function WriteToSharedMemory(Optional ByVal All As Boolean = False, Optional ByVal Index As Long = -1) As Boolean
    Published = True
    WriteToSharedMemory = True
End Function

Public Sub ClearSharedMemoryIndex(ByVal Index As Long)
    With SharedMemory.Instances(Index)
        .mProcessID = 0
        .mAssignedCore = 0
        .mCommand = 0
        .mStatus = 0
    End With
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
