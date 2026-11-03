Attribute VB_Name = "modLabController"
Option Explicit

Private Type STARTUPINFO
    cb As Long
    lpReserved As Long
    lpDesktop As Long
    lpTitle As Long
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As Long
    hStdInput As Long
    hStdOutput As Long
    hStdError As Long
End Type

Private Type PROCESS_INFORMATION
    hProcess As Long
    hThread As Long
    dwProcessId As Long
    dwThreadId As Long
End Type

Private Declare Function CreateProcessW Lib "kernel32" (ByVal ApplicationName As Long, ByVal CommandLine As Long, ByVal ProcessAttributes As Long, ByVal ThreadAttributes As Long, ByVal InheritHandles As Long, ByVal CreationFlags As Long, ByVal Environment As Long, ByVal CurrentDirectory As Long, ByRef Startup As STARTUPINFO, ByRef ProcessInfo As PROCESS_INFORMATION) As Long
Private Declare Function CreateJobObjectW Lib "kernel32" (ByVal Attributes As Long, ByVal Name As Long) As Long
Private Declare Function SetInformationJobObject Lib "kernel32" (ByVal Job As Long, ByVal InformationClass As Long, Information As Any, ByVal Size As Long) As Long
Private Declare Function AssignProcessToJobObject Lib "kernel32" (ByVal Job As Long, ByVal Process As Long) As Long
Private Declare Function ResumeThread Lib "kernel32" (ByVal Thread As Long) As Long
Private Declare Function TerminateProcess Lib "kernel32" (ByVal Process As Long, ByVal ExitCode As Long) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal Handle As Long, ByVal Milliseconds As Long) As Long
Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal Process As Long, ByRef ExitCode As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal Handle As Long) As Long
Private Declare Function QueryPerformanceCounter Lib "kernel32" (ByRef Value As Currency) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (ByRef Value As Currency) As Long
Private Declare Function GetCommandLineW Lib "kernel32" () As Long
Private Declare Function CommandLineToArgvW Lib "shell32" (ByVal CommandLine As Long, ByRef Count As Long) As Long
Private Declare Function LocalFree Lib "kernel32" (ByVal Memory As Long) As Long
Private Declare Function lstrlenW Lib "kernel32" (ByVal Text As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Size As Long)
Private Declare Function GetFullPathNameW Lib "kernel32" (ByVal Path As Long, ByVal BufferSize As Long, ByVal Buffer As Long, ByVal FilePart As Long) As Long
Private Declare Function GetFileAttributesW Lib "kernel32" (ByVal Path As Long) As Long
Private Declare Function CreateDirectoryW Lib "kernel32" (ByVal Path As Long, ByVal Attributes As Long) As Long
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Private Declare Function SetErrorMode Lib "kernel32" (ByVal Mode As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal Milliseconds As Long)
Private Declare Sub ExitProcess Lib "kernel32" (ByVal Code As Long)

Private m_Process As Long
Private m_Job As Long
Private m_Started As Currency
Private m_Frequency As Currency
Private m_Timeout As Long
Private m_ResultPath As String
Private m_Result As String
Private m_Mode As String
Private m_Iterations As Long
Private m_FolderSerial As Long

Public Sub Main()
    Dim Args() As String, Count As Long, Mode As String, Code As Long
    Dim Detail As String, InIDE As Boolean
    On Error GoTo Failed
    Debug.Assert MarkIDE(InIDE)
    If InIDE Then
        frmRoutineLab.Show
        Exit Sub
    End If
    Call SetErrorMode(&H8003&)
    Count = ReadArguments(Args)
    If Count = 1 Then
        frmRoutineLab.Show
        Exit Sub
    End If
    Select Case LCase$(Args(1))
        Case "/worker32", "/workerwow64"
            If Count >= 5 Then m_ResultPath = Args(4)
            If Count <> 6 Then Err.Raise 5, , "Usage: /worker32|/workerwow64 hex iterations result requirement"
            Mode = "x86"
            If LCase$(Args(1)) = "/workerwow64" Then Mode = "wow64"
            Code = RunX86Worker(Args(2), ParsePositiveLong(Args(3)), Args(4), Args(5), Mode)
            ExitProcess Code
        Case "/run"
            If Count >= 7 Then m_ResultPath = Args(6)
            If Count <> 8 Then Err.Raise 5, , "Usage: /run mode hex iterations timeout result requirement"
            StartLabRun Args(2), Args(3), ParsePositiveLong(Args(4)), ParsePositiveLong(Args(5)), Args(6), Args(7)
            Do While PollLabRun()
                Sleep 10
            Loop
            If ResultField(m_Result, "status") = "ok" Then ExitProcess 0
            ExitProcess 2
        Case Else
            Err.Raise 5, , "Unknown command-line mode."
    End Select
    Exit Sub
Failed:
    Detail = Err.Description
    On Error Resume Next
    ReleaseWorker True
    SaveSynthetic "error", Detail
    If InIDE Then
        Debug.Print Detail
        Exit Sub
    End If
    ExitProcess 2
End Sub

Public Sub StartLabRun(ByVal Mode As String, ByVal HexPath As String, ByVal Iterations As Long, ByVal TimeoutMs As Long, ByVal ResultPath As String, ByVal Requirement As String)
    Dim Startup As STARTUPINFO, Child As PROCESS_INFORMATION
    Dim Limits(0 To 111) As Byte, LimitFlags As Long
    Dim Executable As String, Arguments As String, WorkingDirectory As String
    Dim Detail As String
    If LabIsRunning() Then Err.Raise 5, , "A routine is already running."
    On Error GoTo Failed
    m_Result = vbNullString
    m_ResultPath = vbNullString
    Mode = LCase$(Mode)
    Requirement = LCase$(Requirement)
    HexPath = AbsolutePath(HexPath)
    ResultPath = AbsolutePath(ResultPath)
    If StrComp(HexPath, ResultPath, vbTextCompare) = 0 Then Err.Raise 5, , "Input and result paths must be different."
    m_ResultPath = ResultPath
    m_Mode = Mode
    m_Iterations = Iterations
    If Mode <> "x86" And Mode <> "wow64" And Mode <> "x64" Then Err.Raise 5, , "Mode must be x86, wow64, or x64."
    Select Case Requirement
        Case "base", "sse2", "avx", "fma"
        Case Else: Err.Raise 5, , "Requirement must be base, sse2, avx, or fma."
    End Select
    If Iterations < 1 Or TimeoutMs < 1 Then Err.Raise 5, , "Iterations and timeout must be positive integers."
    If GetFileAttributesW(StrPtr(HexPath)) = -1 Then Err.Raise 53, , "Hex input file does not exist."
    m_Timeout = TimeoutMs
    ' Clear a prior report before launch; never accept stale successful output.
    WriteLabFile ResultPath, vbNullString
    WorkingDirectory = App.Path
    If Mode = "x64" Then
        Executable = App.Path & "\RoutineWorker64.exe"
        Arguments = QuoteArgument(Executable)
    Else
        Executable = App.Path & "\" & App.EXEName & ".exe"
        Arguments = QuoteArgument(Executable)
        If Mode = "wow64" Then
            Arguments = Arguments & " /workerwow64"
        Else
            Arguments = Arguments & " /worker32"
        End If
    End If
    Arguments = Arguments & " " & QuoteArgument(HexPath) & " " & CStr(Iterations) & _
        " " & QuoteArgument(ResultPath) & " " & Requirement
    If QueryPerformanceFrequency(m_Frequency) = 0 Or m_Frequency <= 0 Then Err.Raise 5, , "Performance counter is unavailable."
    m_Job = CreateJobObjectW(0, 0)
    If m_Job = 0 Then RaiseAPI "CreateJobObject"
    ' Win32 JOBOBJECT_EXTENDED_LIMIT_INFORMATION is 112 bytes. LimitFlags is
    ' at offset 16; explicit bytes avoid VB6/C structure alignment differences.
    LimitFlags = &H2000& ' JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE
    CopyMemory Limits(16), LimitFlags, 4
    If SetInformationJobObject(m_Job, 9, Limits(0), 112) = 0 Then RaiseAPI "SetInformationJobObject"
    Startup.cb = LenB(Startup)
    Startup.dwFlags = 1 ' STARTF_USESHOWWINDOW
    Startup.wShowWindow = 0
    ' No inherited handles, no console; hold worker suspended until job-owned.
    If CreateProcessW(StrPtr(Executable), StrPtr(Arguments), 0, 0, 0, &H8000004, 0, StrPtr(WorkingDirectory), Startup, Child) = 0 Then RaiseAPI "CreateProcess"
    m_Process = Child.hProcess
    If AssignProcessToJobObject(m_Job, m_Process) = 0 Then RaiseAPI "AssignProcessToJobObject"
    If QueryPerformanceCounter(m_Started) = 0 Then Err.Raise 5, , "Cannot start timeout counter."
    If ResumeThread(Child.hThread) = -1 Then RaiseAPI "ResumeThread"
    Call CloseHandle(Child.hThread)
    Exit Sub
Failed:
    Detail = Err.Description
    On Error Resume Next
    ReleaseWorker True
    If Child.hThread <> 0 Then Call CloseHandle(Child.hThread)
    SaveSynthetic "error", Detail
    On Error GoTo 0
    Err.Raise vbObjectError + 2200, "StartLabRun", Detail
End Sub

Public Function PollLabRun() As Boolean
    Dim State As Long, ExitCode As Long, NowCounter As Currency
    Dim Detail As String
    On Error GoTo Failed
    If m_Process = 0 Then Exit Function
    State = WaitForSingleObject(m_Process, 0)
    If State = 0 Then
        If GetExitCodeProcess(m_Process, ExitCode) = 0 Then RaiseAPI "GetExitCodeProcess"
        ReleaseWorker False
        m_Result = ReadCompletedReport()
        If Not CompleteReport(m_Result) Then
            SaveSynthetic "crash", "Worker exited without a complete report; exit code 0x" & Right$("00000000" & Hex$(ExitCode), 8) & "."
        ElseIf ExitCode <> 0 And ResultField(m_Result, "status") = "ok" Then
            SaveSynthetic "crash", "Worker reported success but exited with code 0x" & Right$("00000000" & Hex$(ExitCode), 8) & "."
        End If
        Exit Function
    End If
    If State <> 258 Then RaiseAPI "WaitForSingleObject"
    If QueryPerformanceCounter(NowCounter) = 0 Then Err.Raise 5, , "Cannot read timeout counter."
    If (CDbl(NowCounter) - CDbl(m_Started)) * 1000# / CDbl(m_Frequency) >= CDbl(m_Timeout) Then
        ReleaseWorker True
        SaveSynthetic "timeout", "Worker exceeded " & CStr(m_Timeout) & " ms and was terminated."
        Exit Function
    End If
    PollLabRun = True
    Exit Function
Failed:
    Detail = Err.Description
    On Error Resume Next
    ReleaseWorker True
    SaveSynthetic "error", "Worker monitoring failed: " & Detail
End Function

Private Function ReadCompletedReport() As String
    ' Missing, inaccessible, or oversized reports all count as a failed worker.
    On Error Resume Next
    ReadCompletedReport = ReadLabFile(m_ResultPath)
End Function

Private Function CompleteReport(ByVal Report As String) As Boolean
    Select Case ResultField(Report, "status")
        Case "ok", "error", "unsupported", "exception"
        Case Else: Exit Function
    End Select
    If Len(ResultField(Report, "architecture")) = 0 Then Exit Function
    If Len(ResultField(Report, "iterations")) = 0 Then Exit Function
    If Len(ResultField(Report, "elapsed_ms")) = 0 Then Exit Function
    If Len(ResultField(Report, "return_hex")) = 0 Then Exit Function
    If Len(ResultField(Report, "scratch_hex")) <> 32 Then Exit Function
    CompleteReport = True
End Function

Public Sub StopLabRun()
    If m_Process = 0 Then Exit Sub
    ' Reap a finished child before cancellation can replace its report.
    If Not PollLabRun() Then Exit Sub
    ReleaseWorker True
    SaveSynthetic "cancelled", "Worker was stopped by the user."
End Sub

Public Function LabIsRunning() As Boolean
    LabIsRunning = (m_Process <> 0)
End Function

Public Function LabResult() As String
    LabResult = m_Result
End Function

Private Sub ReleaseWorker(ByVal Terminate As Boolean)
    ' TerminateProcess covers the suspended child if job assignment failed.
    If Terminate And m_Process <> 0 Then Call TerminateProcess(m_Process, 5)
    If m_Job <> 0 Then
        Call CloseHandle(m_Job)
        m_Job = 0
    End If
    If m_Process <> 0 Then
        If Terminate Then Call WaitForSingleObject(m_Process, 1000)
        Call CloseHandle(m_Process)
        m_Process = 0
    End If
End Sub

Private Sub SaveSynthetic(ByVal Status As String, ByVal Detail As String)
    Dim Architecture As String, ReturnLength As Long
    Architecture = m_Mode
    If Architecture = "wow64" Then Architecture = "wow64-x64"
    If Len(Architecture) = 0 Then Architecture = "controller"
    ReturnLength = 16
    If m_Mode = "x86" Then ReturnLength = 8
    Detail = Replace$(Replace$(Detail, vbCr, " "), vbLf, " ")
    m_Result = "status=" & Status & vbCrLf & "architecture=" & Architecture & vbCrLf & _
        "iterations=" & CStr(m_Iterations) & vbCrLf & "elapsed_ms=0" & vbCrLf & _
        "return_hex=" & String$(ReturnLength, "0") & vbCrLf & _
        "scratch_hex=" & String$(32, "0") & vbCrLf & "error=" & Detail & vbCrLf
    If Len(m_ResultPath) > 0 Then WriteLabFile m_ResultPath, m_Result
End Sub

Private Function ResultField(ByVal Report As String, ByVal Name As String) As String
    Dim Lines() As String, i As Long, Prefix As String
    Lines = Split(Replace$(Report, vbCr, vbNullString), vbLf)
    Prefix = Name & "="
    For i = 0 To UBound(Lines)
        If Left$(Lines(i), Len(Prefix)) = Prefix Then
            ResultField = Mid$(Lines(i), Len(Prefix) + 1)
            Exit Function
        End If
    Next i
End Function

Public Function ReadLabFile(ByVal Path As String) As String
    Dim FileNumber As Integer, Size As Long, Contents As String, Detail As String
    On Error GoTo Failed
    If GetFileAttributesW(StrPtr(Path)) = -1 Then Err.Raise 53, , "File does not exist: " & Path
    FileNumber = FreeFile
    Open Path For Binary Access Read As #FileNumber
    Size = LOF(FileNumber)
    If Size > 262144 Then Err.Raise 5, , "Text file exceeds 256 KiB."
    If Size > 0 Then
        Contents = Space$(Size)
        Get #FileNumber, , Contents
    End If
    Close #FileNumber
    ReadLabFile = Contents
    Exit Function
Failed:
    Detail = Err.Description
    On Error Resume Next
    If FileNumber <> 0 Then Close #FileNumber
    On Error GoTo 0
    Err.Raise vbObjectError + 2201, "ReadLabFile", Detail
End Function

Public Sub WriteLabFile(ByVal Path As String, ByVal Text As String)
    Dim FileNumber As Integer, Detail As String
    On Error GoTo Failed
    FileNumber = FreeFile
    Open Path For Output Access Write As #FileNumber
    Print #FileNumber, Text;
    Close #FileNumber
    Exit Sub
Failed:
    Detail = Err.Description
    On Error Resume Next
    If FileNumber <> 0 Then Close #FileNumber
    On Error GoTo 0
    Err.Raise vbObjectError + 2202, "WriteLabFile", Detail
End Sub

Public Function LabRunFolder() As String
    Dim Parent As String, Candidate As String, Attempt As Long, Attributes As Long
    Parent = App.Path & "\runs"
    Attributes = GetFileAttributesW(StrPtr(Parent))
    If Attributes = -1 Then
        If CreateDirectoryW(StrPtr(Parent), 0) = 0 Then RaiseAPI "Create runs directory"
    ElseIf (Attributes And vbDirectory) = 0 Then
        Err.Raise 5, , "The runs path is not a directory."
    End If
    For Attempt = 1 To 1000
        m_FolderSerial = m_FolderSerial + 1
        Candidate = Parent & "\" & Format$(Now, "yyyymmdd-hhnnss") & "-" & _
            CStr(GetCurrentProcessId()) & "-" & CStr(m_FolderSerial)
        If CreateDirectoryW(StrPtr(Candidate), 0) <> 0 Then
            LabRunFolder = Candidate
            Exit Function
        End If
        If Err.LastDllError <> 183 Then RaiseAPI "Create run directory"
    Next Attempt
    Err.Raise 5, , "Cannot allocate a unique run directory."
End Function

Private Function ParsePositiveLong(ByVal Text As String) As Long
    Dim i As Long, Digit As Long, Value As Double
    If Len(Text) = 0 Then Err.Raise 5, , "Expected a positive decimal integer."
    For i = 1 To Len(Text)
        Digit = Asc(Mid$(Text, i, 1)) - 48
        If Digit < 0 Or Digit > 9 Then Err.Raise 5, , "Expected decimal digits only."
        Value = Value * 10# + Digit
        If Value > 2147483647# Then Err.Raise 5, , "Integer exceeds 2147483647."
    Next i
    If Value < 1 Then Err.Raise 5, , "Integer must be positive."
    ParsePositiveLong = CLng(Value)
End Function

Private Function AbsolutePath(ByVal Path As String) As String
    Dim Buffer As String, Length As Long
    If Len(Path) = 0 Or InStr(Path, Chr$(0)) <> 0 Or InStr(Path, Chr$(34)) <> 0 Then Err.Raise 5, , "Invalid file path."
    Buffer = Space$(32768)
    Length = GetFullPathNameW(StrPtr(Path), Len(Buffer), StrPtr(Buffer), 0)
    If Length = 0 Or Length >= Len(Buffer) Then RaiseAPI "GetFullPathName"
    AbsolutePath = Left$(Buffer, Length)
End Function

Private Function QuoteArgument(ByVal Text As String) As String
    ' File paths cannot contain quotes. Double trailing slashes before the
    ' closing quote according to CommandLineToArgvW/CRT parsing conventions.
    Dim i As Long, Tail As String
    If InStr(Text, Chr$(34)) <> 0 Then Err.Raise 5, , "A command argument contains a quote."
    For i = Len(Text) To 1 Step -1
        If Mid$(Text, i, 1) <> "\" Then Exit For
        Tail = Tail & "\"
    Next i
    QuoteArgument = Chr$(34) & Text & Tail & Chr$(34)
End Function

Private Function ReadArguments(ByRef Args() As String) As Long
    Dim Vector As Long, Count As Long, i As Long, Pointer As Long, Length As Long
    Dim Detail As String
    On Error GoTo Failed
    Vector = CommandLineToArgvW(GetCommandLineW(), Count)
    If Vector = 0 Or Count < 1 Then RaiseAPI "CommandLineToArgvW"
    ReDim Args(0 To Count - 1)
    For i = 0 To Count - 1
        CopyMemory Pointer, ByVal (Vector + i * 4), 4
        Length = lstrlenW(Pointer)
        If Length > 0 Then
            Args(i) = Space$(Length)
            CopyMemory ByVal StrPtr(Args(i)), ByVal Pointer, Length * 2
        End If
    Next i
    Call LocalFree(Vector)
    ReadArguments = Count
    Exit Function
Failed:
    Detail = Err.Description
    If Vector <> 0 Then Call LocalFree(Vector)
    Err.Raise vbObjectError + 2203, "ReadArguments", Detail
End Function

Private Sub RaiseAPI(ByVal Operation As String)
    Dim Code As Long
    Code = Err.LastDllError
    Err.Raise vbObjectError + 2204, "RoutineLab", Operation & " failed (Windows error " & CStr(Code) & ")."
End Sub

Private Function MarkIDE(ByRef InIDE As Boolean) As Boolean
    InIDE = True
    MarkIDE = True
End Function
