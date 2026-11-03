Attribute VB_Name = "modLabX86"
Option Explicit

' This module executes only in the disposable /worker32 child process.
' The routine uses the x86 stdcall WindowProc ABI and must return with RET 16.
Private Declare Function VirtualAlloc Lib "kernel32" (ByVal Address As Long, ByVal Size As Long, ByVal AllocationType As Long, ByVal Protection As Long) As Long
Private Declare Function VirtualProtect Lib "kernel32" (ByVal Address As Long, ByVal Size As Long, ByVal Protection As Long, ByRef OldProtection As Long) As Long
Private Declare Function VirtualFree Lib "kernel32" (ByVal Address As Long, ByVal Size As Long, ByVal FreeType As Long) As Long
Private Declare Function FlushInstructionCache Lib "kernel32" (ByVal Process As Long, ByVal Address As Long, ByVal Size As Long) As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal Code As Long, ByVal Arg1 As Long, ByVal Arg2 As Long, ByVal Arg3 As Long, ByVal Arg4 As Long) As Long
Private Declare Function QueryPerformanceCounter Lib "kernel32" (ByRef Counter As Currency) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (ByRef Frequency As Currency) As Long
Private Declare Function SetErrorMode Lib "kernel32" (ByVal Mode As Long) As Long
Private Declare Function AddVectoredExceptionHandler Lib "kernel32" (ByVal First As Long, ByVal Handler As Long) As Long
Private Declare Function RemoveVectoredExceptionHandler Lib "kernel32" (ByVal Handler As Long) As Long
Private Declare Sub ExitProcess Lib "kernel32" (ByVal ExitCode As Long)

Private Const MEM_COMMIT_RESERVE As Long = &H3000&
Private Const MEM_RELEASE As Long = &H8000&
Private Const PAGE_READWRITE As Long = &H4&
Private Const PAGE_EXECUTE_READ As Long = &H20&
Private Const MAX_CODE As Long = 65536
Private Const MAX_INPUT As Long = 262144
Private Const SCRATCH_SIZE As Long = 4096

Private Type CPURegisters
    EAX As Long
    EBX As Long
    ECX As Long
    EDX As Long
End Type

Private m_ResultPath As String
Private m_Iterations As Long
Private m_Architecture As String
Private m_InIDE As Boolean
Private m_HandlingFault As Boolean
Private m_CodeSegment As Integer

Public Function RunX86Worker(ByVal HexPath As String, ByVal Iterations As Long, ByVal ResultPath As String, ByVal Requirement As String, Optional ByVal Mode As String = "x86") As Long
    Dim CodeBytes() As Byte, ScratchBytes(0 To 15) As Byte
    Dim CodeAddress As Long, ScratchAddress As Long, Handler As Long
    Dim Started As Currency, Finished As Currency, Frequency As Currency
    Dim ReturnValue As Long, ReturnHigh As Long, ReturnHex As String, ScratchHex As String, i As Long
    Dim Detail As String, ResultStatus As String, FailureCode As Long
    On Error GoTo Failed

    m_ResultPath = ResultPath
    m_Iterations = Iterations
    m_Architecture = "x86"
    m_HandlingFault = False
    m_CodeSegment = 0
    Debug.Assert MarkIDE()
    If m_InIDE Then Err.Raise vbObjectError + 2100, , "Worker execution is disabled in the VB6 IDE. Compile and launch the lab."
    Call SetErrorMode(&H1& Or &H2& Or &H8000&)
    Mode = LCase$(Mode)
    If Mode = "wow64" Then
        m_Architecture = "wow64-x64"
    ElseIf Mode <> "x86" Then
        Err.Raise vbObjectError + 2109, , "Unknown worker mode."
    End If
    If Iterations < 1 Then Err.Raise vbObjectError + 2101, , "Iterations must be between 1 and 2147483647."
    Requirement = LCase$(Requirement)
    Select Case Requirement
        Case "base", "sse2", "avx", "fma"
        Case Else: Err.Raise vbObjectError + 2102, , "Unknown instruction requirement."
    End Select
    CodeBytes = ReadCode(HexPath)
    If Mode = "wow64" Then CodeBytes = BuildWOW64Code(CodeBytes)
    Handler = AddVectoredExceptionHandler(1, AddressOf WorkerException)
    If Handler = 0 Then Err.Raise vbObjectError + 2103, , "Cannot install native exception handler."
    If Not SupportsRequirement(Requirement) Then
        ResultStatus = "unsupported"
        Detail = "CPU or operating system does not support the requested " & Requirement & " state."
        FailureCode = 3
        GoTo ReportFailure
    End If
    CodeAddress = AllocateCode(CodeBytes)
    ScratchAddress = VirtualAlloc(0, SCRATCH_SIZE, MEM_COMMIT_RESERVE, PAGE_READWRITE)
    If ScratchAddress = 0 Then Err.Raise vbObjectError + 2104, , "Cannot allocate scratch memory."
    If QueryPerformanceFrequency(Frequency) = 0 Then Err.Raise vbObjectError + 2105, , "Performance counter unavailable."
    If Frequency <= 0 Then Err.Raise vbObjectError + 2106, , "Invalid performance counter frequency."
    If QueryPerformanceCounter(Started) = 0 Then Err.Raise vbObjectError + 2107, , "Cannot read performance counter."
    ReturnValue = CallWindowProc(CodeAddress, 0, ScratchAddress, Iterations, SCRATCH_SIZE)
    If QueryPerformanceCounter(Finished) = 0 Then Err.Raise vbObjectError + 2108, , "Cannot read performance counter."
    CopyMemory ScratchBytes(0), ByVal ScratchAddress, 16
    For i = 0 To 15
        ScratchHex = ScratchHex & Right$("0" & Hex$(ScratchBytes(i)), 2)
    Next i
    If Mode = "wow64" Then
        CopyMemory ReturnValue, ByVal (ScratchAddress + 4080), 4
        CopyMemory ReturnHigh, ByVal (ScratchAddress + 4084), 4
        CopyMemory m_CodeSegment, ByVal (ScratchAddress + 4088), 2
        ReturnHex = Right$("00000000" & Hex$(ReturnHigh), 8) & Right$("00000000" & Hex$(ReturnValue), 8)
    Else
        ReturnHex = Right$("00000000" & Hex$(ReturnValue), 8)
    End If
    WriteResult "ok", "", Trim$(Str$((CDbl(Finished) - CDbl(Started)) / CDbl(Frequency) * 1000#)), ReturnHex, ScratchHex
    RunX86Worker = 0
    GoTo Cleanup

Failed:
    Detail = Err.Description
    FailureCode = 2
    ResultStatus = "error"
ReportFailure:
    On Error Resume Next
    WriteResult ResultStatus, Detail, "0", EmptyReturnHex(), String$(32, "0")
    RunX86Worker = FailureCode
Cleanup:
    On Error Resume Next
    If Handler <> 0 Then Call RemoveVectoredExceptionHandler(Handler)
    If CodeAddress <> 0 Then Call VirtualFree(CodeAddress, 0, MEM_RELEASE)
    If ScratchAddress <> 0 Then Call VirtualFree(ScratchAddress, 0, MEM_RELEASE)
End Function

Private Function MarkIDE() As Boolean
    m_InIDE = True
    MarkIDE = True
End Function

Private Function ReadCode(ByVal Path As String) As Byte()
    Dim FileNumber As Integer, Contents As String, Size As Long
    Dim Description As String
    On Error GoTo Failed
    FileNumber = FreeFile
    Open Path For Binary Access Read As #FileNumber
    Size = LOF(FileNumber)
    If Size < 1 Or Size > MAX_INPUT Then Err.Raise vbObjectError + 2110, , "Hex input must contain 1 to 262144 bytes of text."
    Contents = Space$(Size)
    Get #FileNumber, , Contents
    Close #FileNumber
    FileNumber = 0
    ReadCode = ParseHex(Contents)
    Exit Function
Failed:
    Description = Err.Description
    On Error Resume Next
    If FileNumber <> 0 Then Close #FileNumber
    On Error GoTo 0
    Err.Raise vbObjectError + 2111, , Description
End Function

Private Function ParseHex(ByVal Contents As String) As Byte()
    Dim Bytes() As Byte, Count As Long, Position As Long
    Dim Token As String, Character As String, Comment As Boolean
    ReDim Bytes(0 To MAX_CODE - 1)
    For Position = 1 To Len(Contents) + 1
        If Position > Len(Contents) Then
            Character = " "
        Else
            Character = Mid$(Contents, Position, 1)
        End If
        If Character = vbCr Or Character = vbLf Then Comment = False
        If Not Comment Then
            Select Case Character
                Case ";", " ", vbTab, vbCr, vbLf, vbVerticalTab, vbFormFeed
                    If Len(Token) <> 0 Then
                        If Len(Token) <> 2 Then Err.Raise vbObjectError + 2112, , "Each hex byte must contain exactly two hexadecimal digits."
                        If Count >= MAX_CODE Then Err.Raise vbObjectError + 2113, , "Routine exceeds 65536 bytes."
                        Bytes(Count) = CByte("&H" & Token)
                        Count = Count + 1
                        Token = ""
                    End If
                    If Character = ";" Then Comment = True
                Case "0" To "9", "a" To "f", "A" To "F"
                    Token = Token & Character
                    If Len(Token) > 2 Then Err.Raise vbObjectError + 2114, , "Hex bytes must be separated by whitespace."
                Case Else
                    Err.Raise vbObjectError + 2115, , "Invalid character in hex input at position " & CStr(Position) & "."
            End Select
        End If
    Next Position
    If Count = 0 Then Err.Raise vbObjectError + 2116, , "Routine contains no hex bytes."
    ReDim Preserve Bytes(0 To Count - 1)
    ParseHex = Bytes
End Function

Private Function AllocateCode(ByRef Bytes() As Byte) As Long
    Dim Address As Long, Size As Long, PreviousProtection As Long, Description As String
    On Error GoTo Failed
    Size = UBound(Bytes) - LBound(Bytes) + 1
    Address = VirtualAlloc(0, Size, MEM_COMMIT_RESERVE, PAGE_READWRITE)
    If Address = 0 Then Err.Raise vbObjectError + 2120, , "Cannot allocate code memory."
    CopyMemory ByVal Address, Bytes(LBound(Bytes)), Size
    If VirtualProtect(Address, Size, PAGE_EXECUTE_READ, PreviousProtection) = 0 Then Err.Raise vbObjectError + 2121, , "Cannot make code memory executable."
    If FlushInstructionCache(GetCurrentProcess(), Address, Size) = 0 Then Err.Raise vbObjectError + 2122, , "Cannot flush instruction cache."
    AllocateCode = Address
    Exit Function
Failed:
    Description = Err.Description
    If Address <> 0 Then Call VirtualFree(Address, 0, MEM_RELEASE)
    Err.Raise vbObjectError + 2123, , Description
End Function

Private Function SupportsRequirement(ByVal Requirement As String) As Boolean
    Dim Registers As CPURegisters, Probe() As Byte, Address As Long
    Dim XCR0(0 To 1) As Long, Description As String
    On Error GoTo Failed
    If Requirement = "base" Then
        SupportsRequirement = True
        Exit Function
    End If
    ' CPUID(1,0), with EBX and EDI preserved. ESI and EBP are untouched/preserved.
    Probe = ParseHex("55 89 E5 53 57 8B 45 08 8B 7D 0C 31 C9 0F A2 89 07 89 5F 04 89 4F 08 89 57 0C 5F 5B 5D C2 10 00")
    Address = AllocateCode(Probe)
    Call CallWindowProc(Address, 1, VarPtr(Registers), 0, 0)
    Call VirtualFree(Address, 0, MEM_RELEASE)
    Address = 0
    If Requirement = "sse2" Then
        SupportsRequirement = ((Registers.EDX And &H4000000) <> 0)
        Exit Function
    End If
    ' AVX requires the CPU's AVX + XSAVE + OSXSAVE bits, then XMM/YMM in XCR0.
    If (Registers.ECX And &H1C000000) <> &H1C000000 Then Exit Function
    If Requirement = "fma" And (Registers.ECX And &H1000&) = 0 Then Exit Function
    Probe = ParseHex("55 89 E5 57 8B 7D 0C 31 C9 0F 01 D0 89 07 89 57 04 5F 5D C2 10 00")
    Address = AllocateCode(Probe)
    Call CallWindowProc(Address, 0, VarPtr(XCR0(0)), 0, 0)
    Call VirtualFree(Address, 0, MEM_RELEASE)
    Address = 0
    SupportsRequirement = ((XCR0(0) And &H6&) = &H6&)
    Exit Function
Failed:
    Description = Err.Description
    If Address <> 0 Then Call VirtualFree(Address, 0, MEM_RELEASE)
    Err.Raise vbObjectError + 2124, , Description
End Function

Private Sub WriteResult(ByVal Status As String, ByVal Detail As String, ByVal Elapsed As String, ByVal ReturnHex As String, ByVal ScratchHex As String)
    Dim FileNumber As Integer, Description As String
    On Error GoTo Failed
    FileNumber = FreeFile
    Open m_ResultPath For Output As #FileNumber
    Print #FileNumber, "status=" & Status
    Print #FileNumber, "architecture=" & m_Architecture
    Print #FileNumber, "process_bits=32"
    If m_Architecture = "wow64-x64" Then Print #FileNumber, "code_segment=" & Right$("0000" & Hex$(m_CodeSegment), 4)
    Print #FileNumber, "iterations=" & Trim$(Str$(m_Iterations))
    Print #FileNumber, "elapsed_ms=" & Elapsed
    Print #FileNumber, "return_hex=" & ReturnHex
    Print #FileNumber, "scratch_hex=" & ScratchHex
    If Len(Detail) > 0 Then Print #FileNumber, "error=" & Replace$(Replace$(Detail, vbCr, " "), vbLf, " ")
    Close #FileNumber
    Exit Sub
Failed:
    Description = Err.Description
    On Error Resume Next
    If FileNumber <> 0 Then Close #FileNumber
    On Error GoTo 0
    Err.Raise vbObjectError + 2130, , Description
End Sub

Public Function WorkerException(ByVal ExceptionPointers As Long) As Long
    Dim Record As Long, ExceptionCode As Long
    ' Never unwind through untrusted machine code: report fatal hardware faults
    ' and terminate this child. Non-hardware exceptions remain with VB/Windows.
    CopyMemory Record, ByVal ExceptionPointers, 4
    CopyMemory ExceptionCode, ByVal Record, 4
    Select Case ExceptionCode
        Case &HC0000005, &HC000001D, &HC0000094, &HC0000095, &HC000008C, _
             &HC000008D, &HC000008E, &HC000008F, &HC0000090, &HC0000091, _
             &HC0000092, &HC0000093, &HC0000096, &HC00000FD, &H80000001, &H80000002
            If m_HandlingFault Then ExitProcess 4
            m_HandlingFault = True
            On Error Resume Next
            WriteResult "exception", "Native exception 0x" & Right$("00000000" & Hex$(ExceptionCode), 8), "0", EmptyReturnHex(), String$(32, "0")
            ExitProcess 4
    End Select
    WorkerException = 0
End Function

Private Function EmptyReturnHex() As String
    If m_Architecture = "wow64-x64" Then
        EmptyReturnHex = String$(16, "0")
    Else
        EmptyReturnHex = String$(8, "0")
    End If
End Function
