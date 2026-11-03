VERSION 5.00
Begin VB.Form frmRoutineLab
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Routine Lab - CPU Stress"
   ClientHeight    =   9750
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13500
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9750
   ScaleWidth      =   13500
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lblTitle
      Caption         =   "Machine-code routine workbench"
      Height          =   375
      Left            =   240
      Top             =   180
      Width           =   12720
      BeginProperty Font
         Name            =   "Segoe UI"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblSample
      Caption         =   "Sample routine"
      Height          =   255
      Left            =   240
      Top             =   720
      Width           =   4800
   End
   Begin VB.ComboBox cboSample
      Height          =   315
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1020
      Width           =   4920
   End
   Begin VB.Label lblMode
      Caption         =   "Execution mode (each run uses a disposable worker)"
      Height          =   255
      Left            =   5520
      Top             =   720
      Width           =   7680
   End
   Begin VB.ComboBox cboMode
      Height          =   315
      Left            =   5520
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1020
      Width           =   7680
   End
   Begin VB.Label lblABI
      Caption         =   "Routine calling convention"
      Height          =   540
      Left            =   240
      Top             =   1500
      Width           =   12960
   End
   Begin VB.Label lblHex
      Caption         =   "Machine code: two-digit hex bytes separated by whitespace; semicolon (;) starts a line comment."
      Height          =   255
      Left            =   240
      Top             =   2160
      Width           =   12960
   End
   Begin VB.TextBox txtHex
      Height          =   2700
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   2
      Top             =   2460
      Width           =   12960
      BeginProperty Font
         Name            =   "Consolas"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblIterations
      Caption         =   "Iterations (1 .. 2147483647)"
      Height          =   255
      Left            =   240
      Top             =   5340
      Width           =   2940
   End
   Begin VB.TextBox txtIterations
      Height          =   345
      Left            =   240
      TabIndex        =   3
      Text            =   "1000000"
      Top             =   5640
      Width           =   2700
   End
   Begin VB.Label lblTimeout
      Caption         =   "Worker timeout (milliseconds)"
      Height          =   255
      Left            =   3300
      Top             =   5340
      Width           =   3000
   End
   Begin VB.TextBox txtTimeout
      Height          =   345
      Left            =   3300
      TabIndex        =   4
      Text            =   "2000"
      Top             =   5640
      Width           =   2700
   End
   Begin VB.Label lblRequirement
      Caption         =   "Required CPU/OS features"
      Height          =   255
      Left            =   6360
      Top             =   5340
      Width           =   2940
   End
   Begin VB.ComboBox cboRequirement
      Height          =   315
      Left            =   6360
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   5640
      Width           =   2700
   End
   Begin VB.CommandButton cmdRun
      Caption         =   "Run routine"
      Height          =   435
      Left            =   9660
      TabIndex        =   6
      Top             =   5580
      Width           =   1680
   End
   Begin VB.CommandButton cmdStop
      Caption         =   "Stop"
      Enabled         =   0   'False
      Height          =   435
      Left            =   11520
      TabIndex        =   7
      Top             =   5580
      Width           =   1680
   End
   Begin VB.Label lblResults
      Caption         =   "Result and diagnostics"
      Height          =   255
      Left            =   240
      Top             =   6240
      Width           =   12960
   End
   Begin VB.TextBox txtResults
      Height          =   2580
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   8
      Top             =   6540
      Width           =   12960
      BeginProperty Font
         Name            =   "Consolas"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblStatus
      Caption         =   "Ready. Run files are saved in the runs folder beside this program."
      Height          =   375
      Left            =   240
      Top             =   9300
      Width           =   12960
   End
   Begin VB.Timer tmrWorker
      Enabled         =   0   'False
      Interval        =   50
      Left            =   12600
      Top             =   120
   End
End
Attribute VB_Name = "frmRoutineLab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_Loading As Boolean
Private m_RunFolder As String

Private Sub Form_Load()
    On Error GoTo Failed
    m_Loading = True
    cboMode.AddItem "x86 - 32-bit worker"
    cboMode.AddItem "x64 inside 32-bit WOW64 - direct execution"
    cboMode.AddItem "x64 - native 64-bit comparison"
    cboRequirement.AddItem "base"
    cboRequirement.AddItem "sse2"
    cboRequirement.AddItem "avx"
    cboRequirement.AddItem "fma"
    cboSample.AddItem "return42-x86"
    cboSample.AddItem "proof64-x64"
    cboSample.AddItem "sse2-x86"
    cboSample.AddItem "sse2-x64"
    cboSample.AddItem "fma-x64"
    cboMode.ListIndex = 1
    cboRequirement.ListIndex = 0
    cboSample.ListIndex = 1
    m_Loading = False
    LoadSelectedSample
    Exit Sub
Failed:
    m_Loading = False
    txtResults.Text = "Cannot initialize workbench: " & Err.Description
End Sub

Private Sub cboSample_Click()
    If m_Loading Then Exit Sub
    LoadSelectedSample
End Sub

Private Sub LoadSelectedSample()
    On Error GoTo Failed
    If cboSample.ListIndex < 0 Then Exit Sub
    Select Case cboSample.ListIndex
        Case 0
            cboMode.ListIndex = 0
            cboRequirement.ListIndex = 0
        Case 1
            cboMode.ListIndex = 1
            cboRequirement.ListIndex = 0
        Case 2
            cboMode.ListIndex = 0
            cboRequirement.ListIndex = 1
        Case 3
            cboMode.ListIndex = 1
            cboRequirement.ListIndex = 1
        Case 4
            cboMode.ListIndex = 1
            cboRequirement.ListIndex = 3
    End Select
    UpdateABI
    txtHex.Text = ReadLabFile(App.Path & "\routines\" & cboSample.Text & ".hex")
    txtResults.Text = "Ready: " & cboSample.Text & vbCrLf & _
        "Edit the machine code or run the sample. Iterations are passed to the routine; its code controls the loop."
    Exit Sub
Failed:
    txtHex.Text = vbNullString
    txtResults.Text = "Cannot load sample: " & Err.Description
End Sub

Private Sub cboMode_Click()
    UpdateABI
End Sub

Private Sub UpdateABI()
    Select Case cboMode.ListIndex
        Case 0
            lblABI.Caption = "x86 stdcall: [esp+4]=0, [esp+8]=scratch, [esp+12]=iterations, [esp+16]=4096 scratch bytes at entry. Return EAX with RET 16; preserve nonvolatile registers."
        Case 1
            lblABI.Caption = "x64 inside the 32-bit WOW64 worker: RCX=0, RDX=scratch, R8=iterations, R9=4080 usable scratch bytes. Return RAX with RET; obey the Windows x64 ABI."
        Case 2
            lblABI.Caption = "Native 64-bit comparison worker: RCX=0, RDX=scratch, R8=iterations, R9=4096 scratch bytes. Return RAX with RET; obey the Windows x64 ABI."
    End Select
End Sub

Private Function SelectedMode() As String
    Select Case cboMode.ListIndex
        Case 0: SelectedMode = "x86"
        Case 1: SelectedMode = "wow64"
        Case 2: SelectedMode = "x64"
        Case Else: Err.Raise 5, , "Select an execution mode."
    End Select
End Function

Private Function PositiveLong(ByVal Text As String, ByVal Field As String) As Long
    Dim Index As Long, Value As Double, Digit As Long
    Text = Trim$(Text)
    If Len(Text) = 0 Then Err.Raise 5, , Field & " must contain a positive decimal integer."
    For Index = 1 To Len(Text)
        Digit = Asc(Mid$(Text, Index, 1)) - Asc("0")
        If Digit < 0 Or Digit > 9 Then Err.Raise 5, , Field & " must contain decimal digits only."
        Value = Value * 10# + Digit
        If Value > 2147483647# Then Err.Raise 5, , Field & " must be at most 2147483647."
    Next Index
    If Value < 1# Then Err.Raise 5, , Field & " must be at least 1."
    PositiveLong = CLng(Value)
End Function

Private Sub SetRunningUI(ByVal Running As Boolean)
    cboMode.Enabled = Not Running
    cboSample.Enabled = Not Running
    cboRequirement.Enabled = Not Running
    txtHex.Locked = Running
    txtIterations.Enabled = Not Running
    txtTimeout.Enabled = Not Running
    cmdRun.Enabled = Not Running
    cmdStop.Enabled = Running
    tmrWorker.Enabled = Running
End Sub

Private Sub cmdRun_Click()
    Dim Iterations As Long, TimeoutMs As Long, Mode As String
    On Error GoTo Failed
    If LabIsRunning() Then Exit Sub
    Iterations = PositiveLong(txtIterations.Text, "Iterations")
    TimeoutMs = PositiveLong(txtTimeout.Text, "Timeout")
    Mode = SelectedMode()
    If cboRequirement.ListIndex < 0 Then Err.Raise 5, , "Select a CPU/OS feature requirement."
    If Len(Trim$(txtHex.Text)) = 0 Then Err.Raise 5, , "Enter or load machine code first."
    m_RunFolder = LabRunFolder()
    WriteLabFile m_RunFolder & "\input.hex", txtHex.Text
    StartLabRun Mode, m_RunFolder & "\input.hex", Iterations, TimeoutMs, _
        m_RunFolder & "\result.txt", cboRequirement.Text
    txtResults.Text = "Running " & cboSample.Text & " (editable code)" & vbCrLf & _
        cboMode.Text & vbCrLf & "Iterations: " & CStr(Iterations) & _
        "; timeout: " & CStr(TimeoutMs) & " ms."
    lblStatus.Caption = "Run files: " & m_RunFolder
    SetRunningUI True
    Exit Sub
Failed:
    txtResults.Text = "Cannot start routine: " & Err.Description
    SetRunningUI LabIsRunning()
End Sub

Private Sub tmrWorker_Timer()
    On Error GoTo Failed
    If PollLabRun() Then Exit Sub
    SetRunningUI False
    txtResults.Text = LabResult()
    lblStatus.Caption = "Finished. Run files: " & m_RunFolder
    Exit Sub
Failed:
    txtResults.Text = "Worker monitoring failed: " & Err.Description
    tmrWorker.Enabled = False
    StopLabRun
    SetRunningUI False
End Sub

Private Sub cmdStop_Click()
    On Error GoTo Failed
    StopLabRun
    SetRunningUI False
    txtResults.Text = LabResult()
    lblStatus.Caption = "Stopped. Run files: " & m_RunFolder
    Exit Sub
Failed:
    txtResults.Text = "Cannot stop worker: " & Err.Description
    SetRunningUI LabIsRunning()
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    tmrWorker.Enabled = False
    StopLabRun
End Sub
