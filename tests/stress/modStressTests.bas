Attribute VB_Name = "modStressTests"
Option Explicit

Public IsRunningInIDE As Boolean
Private Declare Sub ExitProcess Lib "kernel32" (ByVal Code As Long)
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Count As Long)
Private Declare Function QueryPerformanceCounter Lib "kernel32" (ByRef Counter As Currency) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (ByRef Counter As Currency) As Long
Private Checks As Long
Private Failures As Long
Private LogFile As Integer

Public Sub Main()
    Dim Result As String, FirstScratch As String, SecondScratch As String
    Dim Kernel As String, i As Long, Started As Currency, Finished As Currency, Frequency As Currency
    Dim HighestBatchMs As Double
    On Error GoTo Failed
    LogFile = FreeFile
    Open App.Path & "\test-results.log" For Output As #LogFile
    IsRunningInIDE = False
    StressShutdown
    Check RejectsCount(0), "zero iterations rejected"
    Check RejectsCount(-1), "negative iterations rejected"
    Check RejectsCount(8000001), "iteration upper bound enforced"
    Result = StressRunBatch(1)
    Kernel = StressKernelName()
    Print #LogFile, "Kernel: " & Kernel
    Check Len(Kernel) > 0, "kernel identified"
    Check Right$(Result, 8) = "00000001", "one iteration returned"
    FirstScratch = StressReadScratch()
    Check Len(FirstScratch) = 64, "32 scratch bytes reported"
    CheckFinite FirstScratch, Kernel, "one iteration"
    CheckReference FirstScratch, Kernel, 1
    Result = StressRunBatch(2)
    SecondScratch = StressReadScratch()
    Check Right$(Result, 8) = "00000002", "two iterations returned"
    Check SecondScratch <> FirstScratch, "arithmetic evolves with iteration count"
    CheckFinite SecondScratch, Kernel, "two iterations"
    CheckReference SecondScratch, Kernel, 2
    Result = StressRunBatch(31)
    Check Right$(Result, 8) = "0000001F", "31 iterations returned"
    CheckFinite StressReadScratch(), Kernel, "31 iterations"
    CheckReference StressReadScratch(), Kernel, 31
    Result = StressRunBatch(257)
    Check Right$(Result, 8) = "00000101", "257 iterations returned across seed reset"
    CheckReference StressReadScratch(), Kernel, 257
    If InStr(1, Kernel, "WOW64", vbTextCompare) > 0 Then
        Check Len(Result) = 16, "WOW64 retains complete 64-bit return"
        Check StressCodeSegment() = &H33&, "production routine executes in 64-bit code segment"
    End If
    Check StressLastElapsedMs() >= 0, "routine duration reported"
    QueryPerformanceFrequency Frequency
    QueryPerformanceCounter Started
    For i = 1 To 12
        StressLoop
        Check StressBatchIterations() >= 1 And StressBatchIterations() <= 8000000, "adaptive batch count bounded"
        If StressLastElapsedMs() > HighestBatchMs Then HighestBatchMs = StressLastElapsedMs()
        Check StressLastElapsedMs() < 1000#, "individual batch returns within one second"
    Next i
    QueryPerformanceCounter Finished
    Check (CDbl(Finished) - CDbl(Started)) / CDbl(Frequency) < 5#, "adaptive batch test is bounded"
    Print #LogFile, "Largest measured batch: " & Format$(HighestBatchMs, "0.000") & " ms"
    CheckFinite StressReadScratch(), Kernel, "adaptive batch"
    StressShutdown
    StressShutdown
    Result = StressRunBatch(1)
    Check Right$(Result, 8) = "00000001", "engine initializes after repeated shutdown"
    Check StressReadScratch() = FirstScratch, "fresh initialization produces deterministic output"
Cleanup:
    On Error Resume Next
    StressShutdown
    Print #LogFile, CStr(Checks) & " checks; " & CStr(Failures) & " failures."
    Close #LogFile
    If Failures > 0 Then ExitProcess 1
    ExitProcess 0
Failed:
    Failures = Failures + 1
    Print #LogFile, "UNEXPECTED: " & CStr(Err.Number) & " - " & Err.Description
    Resume Cleanup
End Sub

Private Sub CheckReference(ByVal Actual As String, ByVal Kernel As String, ByVal Iterations As Long)
    Dim Fused As Boolean, Width As Long, Lane As Long, RegisterIndex As Long, LaneIndex As Long
    Dim Value As Single, FactorDown As Single, FactorUp As Single, BiasUp As Single, BiasDown As Single
    Dim Product As Double, StepIndex As Long, UpdateIndex As Long, Bytes(0 To 3) As Byte
    Dim Expected As String, i As Long, Count As Long
    ' x87 has a different precision/ordering contract; finite-state checks cover
    ' its fallback. SIMD reference uses explicit binary32 rounding boundaries.
    If InStr(1, Kernel, "legacy", vbTextCompare) > 0 Then Exit Sub
    Fused = (InStr(1, Kernel, "FMA", vbTextCompare) > 0)
    Width = 4
    If Fused Or InStr(1, Kernel, "AVX", vbTextCompare) > 0 Then Width = 8
    FactorDown = CSng(1# - 2# ^ -24)
    FactorUp = CSng(1# + 2# ^ -23)
    BiasUp = CSng(2# ^ -20)
    BiasDown = CSng(-(2# ^ -21))
    Count = ((Iterations - 1) Mod 256) + 1
    For Lane = 0 To 7
        RegisterIndex = Lane \ Width
        LaneIndex = Lane Mod Width
        Value = CSng(CDbl(RegisterIndex + 1) / 8# + CDbl(LaneIndex) / 64#)
        For StepIndex = 1 To Count
            For UpdateIndex = 1 To 2
                Product = CDbl(Value) * CDbl(FactorDown)
                If Not Fused Then Product = CDbl(CSng(Product))
                Value = CSng(Product + CDbl(BiasUp))
                Product = CDbl(Value) * CDbl(FactorUp)
                If Not Fused Then Product = CDbl(CSng(Product))
                Value = CSng(Product + CDbl(BiasDown))
            Next UpdateIndex
        Next StepIndex
        CopyMemory Bytes(0), Value, 4
        For i = 0 To 3
            Expected = Expected & Right$("0" & Hex$(Bytes(i)), 2)
        Next i
    Next Lane
    Check Actual = Expected, CStr(Iterations) & " iterations match independent scalar reference"
End Sub

Private Function RejectsCount(ByVal Count As Long) As Boolean
    Dim Result As String
    On Error GoTo Rejected
    Result = StressRunBatch(Count)
    Exit Function
Rejected:
    RejectsCount = True
End Function

Private Sub Check(ByVal Condition As Boolean, ByVal Label As String)
    Checks = Checks + 1
    If Condition Then
        Print #LogFile, "PASS " & Label
    Else
        Failures = Failures + 1
        Print #LogFile, "FAIL " & Label
    End If
End Sub

Private Sub CheckFinite(ByVal Hex As String, ByVal Kernel As String, ByVal Label As String)
    Dim Bytes(0 To 31) As Byte, i As Long, Bits As Long, LaneCount As Long, Nonzero As Boolean
    Check Len(Hex) = 64, Label & " scratch length"
    If Len(Hex) <> 64 Then Exit Sub
    For i = 0 To 31
        Bytes(i) = CByte("&H" & Mid$(Hex, i * 2 + 1, 2))
    Next i
    LaneCount = 8
    If InStr(1, Kernel, "legacy", vbTextCompare) > 0 Then LaneCount = 4
    For i = 0 To LaneCount - 1
        CopyMemory Bits, Bytes(i * 4), 4
        Check (Bits And &H7F800000) <> &H7F800000, Label & " finite lane " & CStr(i)
        If (Bits And &H7FFFFFFF) <> 0 Then Nonzero = True
    Next i
    Check Nonzero, Label & " writes nonzero arithmetic output"
End Sub
