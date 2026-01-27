Attribute VB_Name = "modMain"
Option Explicit

Private Const CMD_SEP_COMMAND As String = ";"
Private Const CMD_SEP_VALUE As String = ":"
Private Const CMD_RUN As String = "uigbfhs"
Private Const CMD_INDEX As String = "uygbrf"

Global Const APP_NAME As String = "Ubehage's CPU-Stress Tool (v1)"

Global Const FONT_MAIN As String = "Segoe UI"
Global Const FONT_SECONDARY As String = "Consolas"
Global Const FONTSIZE_MAIN As Integer = 11
Global Const FONTSIZE_SECONDARY As Integer = 9

Global Const COLOR_BACKGROUND As Long = 2105376
Global Const COLOR_CONTROLS As Long = 2763306
Global Const COLOR_BUTTON_HOVER As Long = 3684408
Global Const COLOR_BUTTON_PRESSED As Long = 3289650
Global Const COLOR_BACKGROUND_DISABLED As Long = 5263440

Global Const COLOR_TEXT As Long = 14737632
Global Const COLOR_TEXT_HOVER As Long = 15790320
Global Const COLOR_TEXT_DISABLED As Long = 7895160
Global Const COLOR_TEXT_ONGREEN As Long = 15463654
Global Const COLOR_TEXT_ONRED As Long = 15395579
Global Const COLOR_TEXT_DISABLED_ONGREEN As Long = 13355947
Global Const COLOR_TEXT_DISABLED_ONRED As Long = 10592542

Global Const COLOR_GREEN As Long = 5023791
Global Const COLOR_GREEN_HOVER As Long = 6339651
Global Const COLOR_GREEN_PRESSED As Long = 4033061
Global Const COLOR_GREEN_DISABLED As Long = 6455130
Global Const COLOR_YELLOW As Long = 4965861
Global Const COLOR_YELLOW_HOVER As Long = 6673645
Global Const COLOR_YELLOW_PRESSED As Long = 3710156
Global Const COLOR_YELLOW_DISABLED As Long = 7112080
Global Const COLOR_RED As Long = 4539862
Global Const COLOR_RED_HOVER As Long = 6513642
Global Const COLOR_RED_PRESSED As Long = 3223992
Global Const COLOR_RED_DISABLED As Long = 6776730
Global Const COLOR_OUTLINE As Long = 3815994
Global Const COLOR_OUTLINE_LIGHT As Long = 7368816

Public Enum XorY_Enum
  xyX = &H1
  xyY = &H2
End Enum

Dim DoRun As Boolean

Global ChangedByCode As Boolean
Global ExitNow As Boolean

Global IsRunningInIDE As Boolean

Sub Main()
  Dim s As Integer
  IsRunningInIDE = IsInIDE()
  Call InitCommonControls
  s = Start
  If s = 0 Then Exit Sub
  If s = 1 Then
    LoadMainForm
  ElseIf s = 2 Then
    FillCPUInfo
    LoadStressForm
  End If
  SharedMemory.Instances(SharedMemOffset).mProcessID = GetMypId
  Call WriteToSharedMemory(False)
End Sub

Private Function Start() As Integer
  If OpenSharedMemory() = False Then
    MsgBox "Could not open shared memory!" & vbCrLf & "This program cannot continue.", vbOKOnly Or vbCritical, "Error opening shared memory"
    Exit Function
  End If
  Call ReadCommandLine(Command)
  If DoRun = True Then
    If SharedMemOffset = 0 Then
      MsgBox "Invalid command parameters!" & vbCrLf & "This program will now exit.", vbOKOnly Or vbInformation, "Invalid command parameters"
      UnloadAll
      Exit Function
    End If
    Start = 2
  Else
    If SharedMemOffset <> 0 Then
      MsgBox "Invalid command parameters!" & vbCrLf & "This program will now exit.", vbOKOnly Or vbInformation, "Invalid command parameters"
      UnloadAll
      Exit Function
    ElseIf CheckPrevInstance() = False Then
      MsgBox "You can only run one instance of this program!", vbOKOnly Or vbInformation, APP_NAME
      UnloadAll
      Exit Function
    End If
    Start = 1
  End If
End Function

Private Sub LoadMainForm()
  Load frmMain
  frmMain.SetForm
End Sub

Private Sub LoadStressForm()
  Load frmStress
  frmStress.Start
End Sub

Public Sub UnloadAll()
  Call CloseSharedMemory
End Sub

Public Function IsInIDE() As Boolean
  Dim inIDE As Boolean
  inIDE = False
  On Error Resume Next
  Debug.Assert MakeIDECheck(inIDE)
  On Error GoTo 0
  IsInIDE = inIDE
End Function

Private Function MakeIDECheck(bSet As Boolean) As Boolean
  bSet = True
  MakeIDECheck = True
End Function

Private Sub ReadCommandLine(CommandLine As String)
  Dim i As Integer, cArr() As String, c As String, v As String, vArr() As String
  cArr() = Split(CommandLine, CMD_SEP_COMMAND)
  For i = LBound(cArr) To UBound(cArr)
    c = cArr(i)
    vArr = Split(c, CMD_SEP_VALUE)
    If UBound(vArr) = 1 Then
      c = vArr(0)
      v = vArr(1)
    End If
    Select Case c
      Case CMD_RUN
        DoRun = True
      Case CMD_INDEX
        If Not (v = "" Or Val(v) = 0) Then SharedMemOffset = CLng(v)
    End Select
  Next
End Sub

Private Function GetAppFile() As String
  Dim p As String
  p = App.Path
  If Right$(p, 1) <> "\" Then p = p & "\"
  GetAppFile = p & App.EXEName & ".exe"
End Function

Private Function GetNewCommandLine(CPUIndex As Long) As String
  GetNewCommandLine = """" & GetAppFile & """ " & GetNewCommandLineParameters(CPUIndex)
End Function

Private Function GetNewCommandLineParameters(CPUIndex As Long) As String
  GetNewCommandLineParameters = CMD_RUN & CMD_SEP_COMMAND & _
                                CMD_INDEX & CMD_SEP_VALUE & CStr(CPUIndex)
End Function

Public Sub LaunchNewStresser(Optional Index As Long = 0)
  Dim cL As String, i As Long
  If Index = 0 Then i = GetNextAvailableIndex() Else i = Index
  If i = 0 Then Exit Sub
  cL = GetNewCommandLine(i)
  Call ReadFromSharedMemory(False, i)
  With SharedMemory.Instances(i)
    .mProcessID = -1
    .mAssignedCore = i
    .mCommand = MEMMSG_RUNSTRESS
  End With
  Call WriteToSharedMemory(False, i)
  On Error GoTo ShellError
  Shell cL, vbNormal
ExitLaunch:
  On Error GoTo 0
  Exit Sub
ShellError:
  Select Case MsgBox("There was an error trying to launch a new process." & vbCrLf & Error, vbRetryCancel Or vbCritical, "Error - " & APP_NAME)
    Case vbRetry
      Resume
    Case Else
      Resume ExitLaunch
  End Select
End Sub

Public Function GetNextAvailableIndex() As Long
  Dim i As Long, r As Long
  For i = 1 To TotalCores
    With SharedMemory.Instances(i)
      If .mProcessID = 0 Then
        GetNextAvailableIndex = i
        Exit For
      End If
    End With
  Next
End Function
