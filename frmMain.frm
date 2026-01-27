VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00202020&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Form1"
   ClientHeight    =   5865
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9060
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5865
   ScaleWidth      =   9060
   StartUpPosition =   3  'Windows Default
   Begin CPU_Stress.StatusBar Status1 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   9
      Top             =   5550
      Width           =   9060
      _ExtentX        =   15981
      _ExtentY        =   556
   End
   Begin CPU_Stress.Button cmdStartAll 
      Height          =   795
      Left            =   4560
      TabIndex        =   7
      Top             =   2850
      Width           =   2700
      _ExtentX        =   4763
      _ExtentY        =   1402
      Caption         =   "Engage All Cores"
      BackColor       =   4539862
      HoverColor      =   6513642
      PressedColor    =   3223992
      ForeColor       =   15395579
      DisabledBackColor=   6776730
      DisabledTextColor=   10592542
      ButtonStyle     =   1
      FontName        =   "Consolas"
      FontSize        =   11.25
      FontBold        =   -1  'True
   End
   Begin CPU_Stress.Frame frmOptions 
      Height          =   2205
      Left            =   4605
      TabIndex        =   1
      Top             =   345
      Width           =   4650
      _ExtentX        =   8202
      _ExtentY        =   3889
      Caption         =   "Options"
      Begin CPU_Stress.Button cmdInterval 
         Height          =   315
         Left            =   2925
         TabIndex        =   6
         Top             =   1575
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   556
         Caption         =   "Apply"
         FontName        =   "Consolas"
      End
      Begin VB.TextBox txtInterval 
         Appearance      =   0  'Flat
         BackColor       =   &H00202020&
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   315
         Left            =   2040
         TabIndex        =   5
         Text            =   "0"
         Top             =   1530
         Width           =   540
      End
      Begin CPU_Stress.CheckBox chkLiveUpdate 
         Height          =   300
         Left            =   420
         TabIndex        =   3
         Top             =   1050
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   529
         Caption         =   "Active Monitoring"
      End
      Begin CPU_Stress.CheckBox chkOnTop 
         Height          =   300
         Left            =   375
         TabIndex        =   2
         Top             =   555
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   529
         Caption         =   "Keep this window on top "
      End
      Begin VB.Label lblInterval 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Update interval:"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   210
         Left            =   315
         TabIndex        =   4
         Top             =   1545
         Width           =   1680
      End
   End
   Begin CPU_Stress.CPUView CPUView1 
      Height          =   4740
      Left            =   315
      TabIndex        =   0
      Top             =   195
      Width           =   4065
      _ExtentX        =   7170
      _ExtentY        =   8361
   End
   Begin CPU_Stress.Button cmdStopAll 
      Height          =   795
      Left            =   4605
      TabIndex        =   8
      Top             =   4170
      Width           =   2700
      _ExtentX        =   4763
      _ExtentY        =   1402
      Caption         =   "Stop all stressers"
      BackColor       =   5023791
      HoverColor      =   6339651
      PressedColor    =   4033061
      ForeColor       =   15463654
      DisabledBackColor=   6455130
      DisabledTextColor=   13355947
      ButtonStyle     =   2
      FontName        =   "Consolas"
      FontSize        =   11.25
      FontBold        =   -1  'True
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const STATUS_TEXT As String = "%p% active processes."

Private Const SETTINGS_APPNAME As String = "UbesCPUStress"
Private Const SETTINGS_SECTION As String = "Settings"
Private Const SETTINGS_ONTOP As String = "OnTop"
Private Const SETTINGS_LIVEUPDATE As String = "LiveUpdate"
Private Const SETTINGS_UPDATEINTERVAL As String = "UpdateInterval"
Private Const SETTINGS_LEFT As String = "WindowLeft"
Private Const SETTINGS_TOP As String = "WindowTop"
Private Const SETTINGS_WIDTH As String = "WindowWidth"
Private Const SETTINGS_HEIGHT As String = "WindowHeight"

Dim o_OnTop As Boolean
Dim o_LiveUpdate As Boolean
Dim o_UpdateInterval As Long

Dim LastGoodTextVal As Long

Dim WithEvents StatusTimer As CPUTimer
Attribute StatusTimer.VB_VarHelpID = -1

Friend Sub SetForm()
  SetInitialValues
  Me.Caption = APP_NAME
  Me.Show
  SetStatusTimer
  SetProcessPriority NORMAL_PRIORITY_CLASS
End Sub

Private Sub MoveObjects()
  On Error GoTo MoveError 'We don't check for invalid values or if the window is too small.
                          'We just ignore any errors and move on.
  CPUView1.Move 30, 30
  chkOnTop.Move 90, (Screen.TwipsPerPixelX * 27)
  chkLiveUpdate.Move chkOnTop.Left, ((chkOnTop.Top + chkOnTop.Height) + (Screen.TwipsPerPixelX * 7))
  txtInterval.Top = ((chkLiveUpdate.Top + chkLiveUpdate.Height) + (Screen.TwipsPerPixelY * 7))
  lblInterval.Move chkLiveUpdate.Left, (txtInterval.Top + ((txtInterval.Height - lblInterval.Height) \ 2))
  txtInterval.Left = ((lblInterval.Left + lblInterval.Width) + (Screen.TwipsPerPixelX * 3))
  cmdInterval.Move ((txtInterval.Left + txtInterval.Width) + (Screen.TwipsPerPixelX * 5)), txtInterval.Top
  frmOptions.Width = (chkOnTop.Width + (chkOnTop.Left * 2))
  frmOptions.Move (Me.ScaleWidth - (frmOptions.Width + CPUView1.Left)), CPUView1.Top, frmOptions.Width, ((cmdInterval.Top + cmdInterval.Height) + chkOnTop.Left)
  CPUView1.Width = (frmOptions.Left - (CPUView1.Left * 2))
  CPUView1.Height = (Status1.Top - (CPUView1.Top + Screen.TwipsPerPixelY))
  Dim t As Long, v As Long
  t = ((frmOptions.Top + frmOptions.Height) + (Screen.TwipsPerPixelY * 5))
  v = ((Me.ScaleHeight - t) - (cmdStartAll.Height * 2))
  v = (v \ 3)
  cmdStartAll.Top = (t + v)
  cmdStopAll.Top = ((cmdStartAll.Top + cmdStartAll.Height) + v)
  cmdStartAll.Left = (frmOptions.Left + ((frmOptions.Width - cmdStartAll.Width) \ 2))
  cmdStopAll.Left = cmdStartAll.Left
  On Error GoTo 0 'Don't forget to reset the error handler.
  Exit Sub
MoveError:
  Resume Next
End Sub

Private Sub SetInitialValues()
  LoadUISettings
  chkOnTop.Value = IIf(o_OnTop, vbChecked, vbUnchecked)
  chkLiveUpdate.Value = IIf(o_LiveUpdate, vbChecked, vbUnchecked)
  CPUView1.UpdateInterval = o_UpdateInterval
  txtInterval.Text = CStr(CPUView1.UpdateInterval)
  Status1.Text = ""
End Sub

Private Sub LoadUISettings()
  Dim tRect As RECT
  o_OnTop = GetSetting(SETTINGS_APPNAME, SETTINGS_SECTION, SETTINGS_ONTOP, True)
  o_LiveUpdate = GetSetting(SETTINGS_APPNAME, SETTINGS_SECTION, SETTINGS_LIVEUPDATE, True)
  o_UpdateInterval = GetSetting(SETTINGS_APPNAME, SETTINGS_SECTION, SETTINGS_UPDATEINTERVAL, 1000)
  With tRect
    .Left = GetSetting(SETTINGS_APPNAME, SETTINGS_SECTION, SETTINGS_LEFT, 0)
    .Top = GetSetting(SETTINGS_APPNAME, SETTINGS_SECTION, SETTINGS_TOP, 0)
    .Right = GetSetting(SETTINGS_APPNAME, SETTINGS_SECTION, SETTINGS_WIDTH, 0)
    .Bottom = GetSetting(SETTINGS_APPNAME, SETTINGS_SECTION, SETTINGS_HEIGHT, 0)
    If (.Left <> 0 And .Top <> 0) Then
      If (.Right <> 0 And .Bottom <> 0) Then
        Me.Move .Left, .Top, .Right, .Bottom
      End If
    End If
  End With
End Sub

Private Sub SaveUISettings()
  SaveSetting SETTINGS_APPNAME, SETTINGS_SECTION, SETTINGS_ONTOP, (chkOnTop.Value = vbChecked)
  SaveSetting SETTINGS_APPNAME, SETTINGS_SECTION, SETTINGS_LIVEUPDATE, (chkLiveUpdate.Value = vbChecked)
  SaveSetting SETTINGS_APPNAME, SETTINGS_SECTION, SETTINGS_UPDATEINTERVAL, CPUView1.UpdateInterval
  SaveSetting SETTINGS_APPNAME, SETTINGS_SECTION, SETTINGS_LEFT, Me.Left
  SaveSetting SETTINGS_APPNAME, SETTINGS_SECTION, SETTINGS_TOP, Me.Top
  SaveSetting SETTINGS_APPNAME, SETTINGS_SECTION, SETTINGS_WIDTH, Me.Width
  SaveSetting SETTINGS_APPNAME, SETTINGS_SECTION, SETTINGS_HEIGHT, Me.Height
End Sub

Private Sub CheckIntervalText()
  Dim v As Long, e As Boolean
  If txtInterval.Text <> "" Then v = CLng(txtInterval.Text)
  If v > 0 Then e = Not (v = CPUView1.UpdateInterval) Else e = False
  If (e = True And CPUView1.AutoUpdate = True) Then cmdInterval.Enabled = True Else cmdInterval.Enabled = False
  txtInterval.Enabled = CPUView1.AutoUpdate
End Sub

Private Sub SetStatusTimer()
  KillStatusTimer
  Set StatusTimer = New CPUTimer
  StatusTimer.Interval = 500
  StatusTimer.Enabled = True
End Sub

Private Sub KillStatusTimer()
  If StatusTimer Is Nothing Then Exit Sub
  StatusTimer.Enabled = False
  Set StatusTimer = Nothing
End Sub

Private Sub CheckActiveProcesses()
  Dim i As Long, p As Long
  Call ReadFromSharedMemory(True)
  For i = 1 To TotalCores
    With SharedMemory.Instances(i)
      Select Case .mProcessID
        Case 0, -1
          CPUView1.CPUIsBusy(i) = False
        Case Else
          If IsProcessAlive(.mProcessID) = False Then
            ClearSharedMemoryIndex i
            CPUView1.CPUIsBusy(i) = False
          Else
            If .mStatus = MEMSTATUS_RUNNING Then
              p = (p + 1)
              CPUView1.CPUIsBusy(i) = True
            Else
              CPUView1.CPUIsBusy(i) = False
            End If
          End If
      End Select
    End With
  Next
  Status1.Text = Replace$(STATUS_TEXT, "%p%", CStr(p))
  cmdStopAll.Enabled = Not (p = 0)
  cmdStartAll.Enabled = (p < TotalCores)
End Sub

Private Sub StartAllStressers()
  Dim i As Long
  Call ReadFromSharedMemory(True)
  For i = 1 To TotalCores
    With SharedMemory.Instances(i)
      If .mProcessID = -1 Then
        'reserved
      ElseIf .mProcessID = 0 Then
        LaunchNewStresser i
      Else
        If IsProcessAlive(.mProcessID) = True Then
          If .mStatus = MEMSTATUS_IDLE Then
            .mCommand = MEMMSG_RUNSTRESS
            Call WriteToSharedMemory(False, i)
          End If
        Else
          ClearSharedMemoryIndex i
          LaunchNewStresser i
        End If
      End If
    End With
  Next
End Sub

Private Sub CloseAllStressers(Optional StopOnly As Boolean = False)
  Dim i As Long
  Call ReadFromSharedMemory(True)
  For i = 1 To TotalCores
    With SharedMemory.Instances(i)
      If .mProcessID <> 0 Then
        If StopOnly = True Then .mCommand = MEMMSG_STOPSTRESS Else .mCommand = MEMMSG_EXIT
        Call WriteToSharedMemory(False, i)
      End If
    End With
  Next
End Sub

Private Sub chkLiveUpdate_Click()
  CPUView1.AutoUpdate = (chkLiveUpdate.Value = vbChecked)
  CheckIntervalText
End Sub

Private Sub chkOnTop_Click()
  WindowOnTop Me.hWnd, (chkOnTop.Value = vbChecked)
End Sub

Private Sub cmdInterval_Click()
  Dim v As Long
  If txtInterval.Text <> "" Then v = CLng(txtInterval.Text)
  If v > 0 Then CPUView1.UpdateInterval = CLng(txtInterval.Text)
  CheckIntervalText
End Sub

Private Sub cmdStartAll_Click()
  cmdStartAll.Enabled = False
  cmdStopAll.Enabled = False
  StartAllStressers
End Sub

Private Sub cmdStopAll_Click()
  CloseAllStressers True
End Sub

Private Sub CPUView1_Click(Index As Long)
  Call ReadFromSharedMemory(False, Index)
  With SharedMemory.Instances(Index)
    If .mProcessID = 0 Then
      LaunchNewStresser Index
    Else
      If .mProcessID <> -1 Then
        If IsProcessAlive(.mProcessID) = False Then
          ClearSharedMemoryIndex Index
          LaunchNewStresser Index
        Else
          If .mStatus = MEMSTATUS_RUNNING Then
            .mCommand = MEMMSG_STOPSTRESS
          ElseIf .mStatus = MEMSTATUS_IDLE Then
            .mCommand = MEMMSG_RUNSTRESS
          End If
          Call WriteToSharedMemory(False, Index)
        End If
      End If
    End If
  End With
End Sub

Private Sub Form_Resize()
  MoveObjects
End Sub

Private Sub Form_Unload(Cancel As Integer)
  KillStatusTimer
  ClearSharedMemoryIndex SharedMemOffset
  CloseAllStressers
  SaveUISettings
  UnloadAll
End Sub

Private Sub StatusTimer_Timer()
  StatusTimer.Enabled = False
  If ExitNow Then Unload Me
  CheckActiveProcesses
  StatusTimer.Enabled = True
End Sub

Private Sub txtInterval_Change()
  With txtInterval
    If (IsNumeric(.Text) = False And .Text <> "") Then
      ChangedByCode = True
      txtInterval.Text = LastGoodTextVal
      ChangedByCode = False
    Else
      LastGoodTextVal = .Text
    End If
  End With
  CheckIntervalText
End Sub

Private Sub txtInterval_GotFocus()
  With txtInterval
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub

Private Sub txtInterval_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyUp, vbKeyDown, vbKeyLeft, vbKeyRight, vbKeyHome, vbKeyEnd, vbKeyDelete, vbKeyBack
      'do nothing...
    Case Else
      KeyCode = 0
  End Select
End Sub

Private Sub txtInterval_KeyPress(KeyAscii As Integer)
  Select Case KeyAscii
    Case vbKeyBack, vbKeyDelete
      Exit Sub
  End Select
  If KeyAscii = vbKeyBack Then Exit Sub
  If (KeyAscii < 48 Or KeyAscii > 57) Then KeyAscii = 0
End Sub
