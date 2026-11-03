VERSION 5.00
Begin VB.UserControl CPUView 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00202020&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   EditAtDesignTime=   -1  'True
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   11.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00E0E0E0&
   ScaleHeight     =   3600
   ScaleWidth      =   4800
End
Attribute VB_Name = "CPUView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IMouseSink

Private Const PROPNAME_AUTOUPDATE As String = "AutoUpdate"
Private Const PROPNAME_UPDATEINTERVAL As String = "UpdateInterval"
Private Const PROPNAME_AUTOSIZE As String = "AutoSize"

Private Const CPU_UPDATE_INTERVAL As Long = 1000

Private Const CPUVIEW_TITLE As String = "CPU Overview"

Private Const TOP_INDENT As Long = 23
Private Const KERNEL_PIXEL_WIDTH As Long = 20
Private Const KERNEL_PIXEL_HEIGHT As Long = 40
Private Const KERNEL_PIXEL_SPACING As Long = 7
Private Const KERNEL_PIXEL_INDENT = 5

Private Const PROCESSOR_SAMPLE_BYTES As Long = 48
Private Const MAX_SAMPLE_BYTES As Long = 65536

Public Enum CPUView_Autosize
  NoAutoSize = &H0
  FixedSize = &H1
  ViewAutoSize = &H2
End Enum

Private Declare Function CPUQueryGetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal ModuleName As String) As Long
Private Declare Function CPUQueryGetProcAddress Lib "kernel32" Alias "GetProcAddress" (ByVal ModuleHandle As Long, ByVal ProcedureName As String) As Long
Private Declare Function CPUQueryCall Lib "user32" Alias "CallWindowProcA" (ByVal ProcedureAddress As Long, ByVal InformationClass As Long, ByVal Buffer As Long, ByVal BufferLength As Long, ByVal ReturnLength As Long) As Long
Private Declare Sub CPUQueryCopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Size As Long)
    
Dim m_AutoUpdate As Boolean
Dim m_UpdateInterval As Long
Dim m_AutoSize As CPUView_Autosize

Dim cpuRECTs() As RECT
Dim CPULoad() As Double
Dim oldCPULoad() As Double
Dim CPUBusy() As Boolean
Dim CPUHibernating() As Boolean

Dim ViewCols As Long
Dim ViewRows As Long
Dim ViewTop As Long

Dim m_MouseToken As Long
Dim m_IsHovering As Boolean
Dim m_HoverIndex As Long
Dim m_IsPressed As Boolean
Dim m_MouseIsDown As Boolean
Dim m_PressedIndex As Long

Dim m_DisplayCPUName As String

Dim FixedWindowSize As POINTAPI
Dim LoadTextSize As Long

Dim gHasData As Boolean
Dim m_QueryResolved As Boolean
Dim m_QueryAddress As Long
Dim m_SampleBuffer() As Byte
Dim m_SampleCPUCount As Long
Dim m_PreviousIdle() As Currency
Dim m_PreviousKernel() As Currency
Dim m_PreviousUser() As Currency
Dim m_HaveBaseline As Boolean
Dim m_LoadUnavailable As Boolean

Dim WithEvents UpdateTimer As CPUTimer
Attribute UpdateTimer.VB_VarHelpID = -1

Public Event Click(Index As Long)

Public Property Get AutoUpdate() As Boolean
  AutoUpdate = m_AutoUpdate
End Property
Public Property Get UpdateInterval() As Long
  UpdateInterval = m_UpdateInterval
End Property
Public Property Get AutoSizeView() As CPUView_Autosize
  AutoSizeView = m_AutoSize
End Property
Public Property Get CPUIsBusy(Index As Long) As Boolean
  If (Index < 1 Or Index > TotalCores) Then Exit Property
  CPUIsBusy = CPUBusy(Index)
End Property
Public Property Get CPUIsHibernating(Index As Long) As Boolean
  If (Index < 1 Or Index > TotalCores) Then Exit Property
  CPUIsHibernating = CPUHibernating(Index)
End Property
Public Property Let AutoUpdate(New_AutoUpdate As Boolean)
  If m_AutoUpdate = New_AutoUpdate Then Exit Property
  m_AutoUpdate = New_AutoUpdate
  UserControl.PropertyChanged PROPNAME_AUTOUPDATE
  ShiftAutoUpdate
End Property
Public Property Let UpdateInterval(New_UpdateInterval As Long)
  If (m_UpdateInterval = New_UpdateInterval Or New_UpdateInterval = 0) Then Exit Property
  m_UpdateInterval = New_UpdateInterval
  UserControl.PropertyChanged PROPNAME_UPDATEINTERVAL
  If Not UpdateTimer Is Nothing Then UpdateTimer.Interval = m_UpdateInterval
End Property
Public Property Let AutoSizeView(New_AutoSize As CPUView_Autosize)
  If m_AutoSize = New_AutoSize Then Exit Property
  m_AutoSize = New_AutoSize
  UserControl.PropertyChanged PROPNAME_AUTOSIZE
  DoAutoSize
End Property
Public Property Let CPUIsBusy(Index As Long, New_Busy As Boolean)
  If (Index < 1 Or Index > TotalCores) Then Exit Property
  If CPUBusy(Index) = New_Busy Then Exit Property
  CPUBusy(Index) = New_Busy
  DrawCPURect Index
End Property
Public Property Let CPUIsHibernating(Index As Long, New_Hibernating As Boolean)
  If (Index < 1 Or Index > TotalCores) Then Exit Property
  If CPUHibernating(Index) = New_Hibernating Then Exit Property
  CPUHibernating(Index) = New_Hibernating
  DrawCPURect Index
End Property

Public Sub Refresh(Optional FullRefresh As Boolean = False)
  UserControl.Cls
  DrawBorderAndTitle
  DrawCPUName
  If FullRefresh Then SetCPURECTs
  DrawCPURects True
  If FullRefresh And m_MouseToken <> 0 Then MouseRefresh m_MouseToken
End Sub

Public Sub UpdateView()
  Dim i As Long, WasUnavailable As Boolean, HadData As Boolean
  WasUnavailable = m_LoadUnavailable
  HadData = gHasData
  GetCPULoads
  For i = 0 To (TotalCores - 1)
    If CPULoad(i) <> oldCPULoad(i) Or WasUnavailable <> m_LoadUnavailable Or HadData <> gHasData Then
      DrawCPURect (i + 1), True
      oldCPULoad(i) = CPULoad(i)
    End If
  Next
End Sub

Private Sub DrawBorderAndTitle()
  DrawBorder
  DrawTitle
End Sub

Private Sub DrawBorder()
  Dim cY As Long
  With UserControl
    cY = (Screen.TwipsPerPixelY * 10)
    UserControl.Line (0, cY)-((.ScaleWidth - Screen.TwipsPerPixelX), (.ScaleHeight - Screen.TwipsPerPixelY)), COLOR_OUTLINE, B
    UserControl.Line (Screen.TwipsPerPixelX, (cY + Screen.TwipsPerPixelY))-((.ScaleWidth - (Screen.TwipsPerPixelX * 2)), (.ScaleHeight - (Screen.TwipsPerPixelY * 2))), COLOR_OUTLINE, B
  End With
End Sub

Private Sub DrawTitle()
  SetControlFont True, , True
  With UserControl
    .CurrentX = (Screen.TwipsPerPixelX * 10)
    .CurrentY = Screen.TwipsPerPixelY
  End With
  UserControl.Print CPUVIEW_TITLE
  ViewTop = (UserControl.CurrentY + Screen.TwipsPerPixelY)
End Sub

Private Sub DrawCPUName()
  Dim X As Long
  SetControlFont , True, True
  With UserControl
    .CurrentX = ((UserControl.ScaleWidth - UserControl.TextWidth(m_DisplayCPUName)) \ 2)
    .CurrentY = ViewTop
  End With
  UserControl.Print m_DisplayCPUName
  ViewTop = (UserControl.CurrentY + Screen.TwipsPerPixelY)
End Sub

Private Sub DrawCPURects(Optional FullRedraw As Boolean = False)
  Dim i As Long
  For i = 1 To TotalCores
    DrawCPURect i, FullRedraw
  Next
End Sub

Private Sub DrawCPURect(RectIndex As Long, Optional FullRedraw As Boolean = False)
  If (RectIndex <= 0 Or RectIndex > TotalCores) Then Exit Sub
  With cpuRECTs(RectIndex)
    If FullRedraw = True Then UserControl.Line (.Left, .Top)-(.Right, .Bottom), GetCPUBackColor(RectIndex), BF
    If gHasData Then DrawCPUFlood RectIndex
    If m_AutoUpdate = True Then DrawCPULoadText RectIndex
    UserControl.Line (.Left, .Top)-(.Right, .Bottom), GetCPUBorderColor(RectIndex), B
  End With
End Sub

Private Sub DrawCPUFlood(CPUIndex As Long)
  Dim h As Long
  h = GetCPUFloodHeight(CPUIndex)
  If h < 1 Then Exit Sub
  With cpuRECTs(CPUIndex)
    UserControl.Line (.Left, (.Bottom - h))-(.Right, .Bottom), GetCPUFloodColor(CPUIndex), BF
  End With
End Sub

Private Sub DrawCPULoadText(CPUIndex As Long)
  Dim t As String
  If LoadTextSize = 0 Then Exit Sub
  If m_LoadUnavailable Then
    t = "N/A"
  Else
    t = CStr(Int(CPULoad(CPUIndex - 1)) & "%")
  End If
  SetControlFont , True
  UserControl.ForeColor = COLOR_TEXT
  With cpuRECTs(CPUIndex)
    UserControl.CurrentX = (.Left + (((.Right - .Left) - UserControl.TextWidth(t)) \ 2))
    UserControl.CurrentY = (.Bottom - (UserControl.TextHeight(t) + (Screen.TwipsPerPixelY * 2)))
  End With
  UserControl.Print t
End Sub

Private Function GetCPUFloodHeight(CPUIndex As Long) As Long
  Dim h As Long, r As Long
  With cpuRECTs(CPUIndex)
    h = (.Bottom - .Top)
  End With
  r = ((h / 100) * CPULoad(CPUIndex - 1))
  If r > h Then r = h
  GetCPUFloodHeight = r
End Function

Private Function GetCPUBorderColor(CPUIndex As Long) As Long
  If CPUHibernating(CPUIndex) = True Then
    GetCPUBorderColor = COLOR_YELLOW
  ElseIf CPUBusy(CPUIndex) = True Then
    GetCPUBorderColor = COLOR_TEXT
  Else
    GetCPUBorderColor = COLOR_OUTLINE_LIGHT
  End If
End Function

Private Function GetCPUFloodColor(CPUIndex As Long) As Long
  Dim cNorm As Boolean, cHov As Boolean, cClick As Boolean
  If (m_MouseIsDown = True And m_IsPressed = True And m_PressedIndex = CPUIndex) Then
    cClick = True
  Else
    If (m_IsHovering = True And m_HoverIndex = CPUIndex) Then
      cHov = True
    Else
      cNorm = True
    End If
  End If
  Select Case CPULoad(CPUIndex - 1)
    Case Is >= 90
      If cClick = True Then
        GetCPUFloodColor = COLOR_RED_PRESSED
      ElseIf cHov = True Then
        GetCPUFloodColor = COLOR_RED_HOVER
      Else
        GetCPUFloodColor = COLOR_RED
      End If
    Case Is >= 70
      If cClick = True Then
        GetCPUFloodColor = COLOR_YELLOW_PRESSED
      ElseIf cHov = True Then
        GetCPUFloodColor = COLOR_YELLOW_HOVER
      Else
        GetCPUFloodColor = COLOR_YELLOW
      End If
    Case Else
      If cClick = True Then
        GetCPUFloodColor = COLOR_GREEN_PRESSED
      ElseIf cHov = True Then
        GetCPUFloodColor = COLOR_GREEN_HOVER
      Else
        GetCPUFloodColor = COLOR_GREEN
      End If
  End Select
End Function

Private Function GetCPUBackColor(CPUIndex As Long) As Long
  If (m_MouseIsDown = True And m_IsPressed = True And m_PressedIndex = CPUIndex) Then
    GetCPUBackColor = COLOR_BUTTON_PRESSED
  Else
    If (m_IsHovering = True And m_HoverIndex = CPUIndex) Then
      GetCPUBackColor = COLOR_BUTTON_HOVER
    Else
      GetCPUBackColor = COLOR_BACKGROUND
    End If
  End If
End Function

Private Sub SetControlFont(Optional MainFont As Boolean = False, Optional SecondaryFont As Boolean = False, Optional BoldFont As Boolean = False)
  With UserControl.Font
    If MainFont = True Then
      .Name = FONT_MAIN
      .Size = FONTSIZE_MAIN
    End If
    If SecondaryFont = True Then
      .Name = FONT_SECONDARY
      .Size = FONTSIZE_SECONDARY
    End If
    .Bold = BoldFont
  End With
End Sub

Private Sub GetCPULoads()
  Dim i As Long, Offset As Long, Slots As Long, Returned As Long, Status As Long
  Dim Idle As Currency, Kernel As Currency, User As Currency
  Dim IdleDelta As Double, TotalDelta As Double, UsagePercent As Double
  On Error GoTo Unavailable
  ' One OS snapshot replaces PDH's expensive first-use counter setup on the
  ' UI thread. Resolve dynamically because this native API can be unavailable.
  If Not m_QueryResolved Then
    m_QueryResolved = True
    m_QueryAddress = CPUQueryGetProcAddress(CPUQueryGetModuleHandle("ntdll.dll"), "NtQuerySystemInformation")
  End If
  If m_QueryAddress = 0 Then GoTo Unavailable
  If TotalCores < 1 Or TotalCores > MAX_SAMPLE_BYTES \ PROCESSOR_SAMPLE_BYTES Then GoTo Unavailable
  If m_SampleCPUCount <> TotalCores Then
    Slots = TotalCores
    If Slots < 64 Then Slots = 64
    ReDim m_SampleBuffer(0 To Slots * PROCESSOR_SAMPLE_BYTES - 1)
    ReDim m_PreviousIdle(0 To TotalCores - 1)
    ReDim m_PreviousKernel(0 To TotalCores - 1)
    ReDim m_PreviousUser(0 To TotalCores - 1)
    m_SampleCPUCount = TotalCores
    m_HaveBaseline = False
  End If
  ' Class 8: SYSTEM_PROCESSOR_PERFORMANCE_INFORMATION, 48-byte stride.
  ' This four-argument stdcall OS function matches CallWindowProc's ABI.
  Status = CPUQueryCall(m_QueryAddress, 8, VarPtr(m_SampleBuffer(0)), UBound(m_SampleBuffer) + 1, VarPtr(Returned))
  If Status = &HC0000004 Or Status = &HC0000023 Then
    ' At most one bounded resize/retry if Windows reports more processors.
    If Returned <= UBound(m_SampleBuffer) + 1 Or Returned > MAX_SAMPLE_BYTES Then GoTo Unavailable
    ReDim m_SampleBuffer(0 To Returned - 1)
    Returned = 0
    Status = CPUQueryCall(m_QueryAddress, 8, VarPtr(m_SampleBuffer(0)), UBound(m_SampleBuffer) + 1, VarPtr(Returned))
  End If
  If Status < 0 Then GoTo Unavailable
  If Returned < TotalCores * PROCESSOR_SAMPLE_BYTES Or Returned > UBound(m_SampleBuffer) + 1 Then GoTo Unavailable
  If Returned Mod PROCESSOR_SAMPLE_BYTES <> 0 Then GoTo Unavailable
  For i = 0 To TotalCores - 1
    Offset = i * PROCESSOR_SAMPLE_BYTES
    CPUQueryCopyMemory Idle, m_SampleBuffer(Offset), 8
    CPUQueryCopyMemory Kernel, m_SampleBuffer(Offset + 8), 8
    CPUQueryCopyMemory User, m_SampleBuffer(Offset + 16), 8
    If m_HaveBaseline Then
      ' Currency preserves the 64-bit counters; its scale cancels in the ratio.
      ' Kernel time includes idle time, so total = kernel + user.
      IdleDelta = CDbl(Idle) - CDbl(m_PreviousIdle(i))
      TotalDelta = (CDbl(Kernel) - CDbl(m_PreviousKernel(i))) + (CDbl(User) - CDbl(m_PreviousUser(i)))
      If IdleDelta < 0 Or TotalDelta < 0 Then GoTo Unavailable
      UsagePercent = 0
      If TotalDelta > 0 Then UsagePercent = (1# - IdleDelta / TotalDelta) * 100#
      If UsagePercent < 0 Then UsagePercent = 0
      If UsagePercent > 100 Then UsagePercent = 100
      CPULoad(i) = UsagePercent
    End If
    m_PreviousIdle(i) = Idle
    m_PreviousKernel(i) = Kernel
    m_PreviousUser(i) = User
  Next i
  gHasData = m_HaveBaseline
  m_HaveBaseline = True
  m_LoadUnavailable = False
  Exit Sub
Unavailable:
  ' Never keep stale percentages after a failed/incompatible snapshot.
  gHasData = False
  m_HaveBaseline = False
  m_LoadUnavailable = True
  ClearCPULoads
End Sub

Private Sub ClearCPULoads()
  Dim i As Long
  For i = LBound(CPULoad) To UBound(CPULoad)
    CPULoad(i) = 0
  Next
End Sub

Private Sub SetCPURECTs()
  Dim i As Long, j As Long, c As Long
  Dim X As Long, Y As Long, w As Long, h As Long
  SetRowsAndCols
  Y = ViewTop
  w = GetCPURectSize(xyX)
  h = GetCPURectSize(xyY)
  For i = 1 To ViewRows
    X = (Screen.TwipsPerPixelX * KERNEL_PIXEL_INDENT)
    For j = 1 To ViewCols
      c = (c + 1)
      If c > TotalCores Then Exit Sub
      With cpuRECTs(c)
        .Left = X
        .Top = Y
        .Right = (.Left + w)
        .Bottom = (.Top + h)
        X = (.Right + (Screen.TwipsPerPixelX * KERNEL_PIXEL_SPACING))
      End With
    Next
    Y = (cpuRECTs(c).Bottom + (Screen.TwipsPerPixelY * KERNEL_PIXEL_SPACING))
  Next
  SetCPUTextSize
End Sub

Private Sub SetRowsAndCols()
  Dim tCols As Long, tRows As Long
  ViewCols = TotalCores
  ViewRows = 1
  Do
    tCols = (ViewCols / 2)
    tRows = (ViewRows * 2)
    If tRows > tCols Then Exit Do
    ViewCols = tCols
    ViewRows = tRows
  Loop
  tCols = (ViewCols * ViewRows)
  If tCols < TotalCores Then ViewCols = (ViewCols + 1)
End Sub

Private Function GetCPURectSize(XorY As XorY_Enum) As Long
  If (m_AutoSize = NoAutoSize Or m_AutoSize = FixedSize) Then
    If XorY = xyX Then
      GetCPURectSize = (Screen.TwipsPerPixelX * KERNEL_PIXEL_WIDTH)
    Else
      GetCPURectSize = (Screen.TwipsPerPixelY * KERNEL_PIXEL_HEIGHT)
    End If
  ElseIf m_AutoSize = ViewAutoSize Then
    Dim t As Long
    If XorY = xyX Then
      t = (UserControl.ScaleWidth - (Screen.TwipsPerPixelX * (KERNEL_PIXEL_INDENT * 2)))
      t = (t - (Screen.TwipsPerPixelX * (KERNEL_PIXEL_SPACING * (ViewCols - 1))))
      GetCPURectSize = (t / ViewCols)
    ElseIf XorY = xyY Then
      t = (UserControl.ScaleHeight - ((Screen.TwipsPerPixelY * KERNEL_PIXEL_INDENT) + ViewTop))
      t = (t - (Screen.TwipsPerPixelY * (KERNEL_PIXEL_SPACING * (ViewRows - 1))))
      GetCPURectSize = (t / ViewRows)
    End If
  End If
End Function

Private Sub SetCPUTextSize()
  Dim sMin As Long, sMax As Long, sVal As Long
  Dim tW As Long, tH As Long, rW As Long, rH As Long
  LoadTextSize = 0
  With cpuRECTs(1)
    rW = ((.Right - .Left) - (Screen.TwipsPerPixelX * 4))
    rH = ((.Bottom - .Top) - (Screen.TwipsPerPixelY * 4))
  End With
  If (rW <= 60 Or rH <= 60) Then
    Exit Sub
  End If
  sMax = 120
  sMin = 1
  Do While sMin <= sMax
    sVal = (sMin + sMax) \ 2
    UserControl.Font.Size = sVal
    tW = UserControl.TextWidth("100%")
    tH = UserControl.TextHeight("100%")
    If (tW <= rW And tH <= rH) Then
      LoadTextSize = sVal
      sMin = (sVal + 1)
    Else
      sMax = (sVal - 1)
    End If
  Loop
End Sub

Private Sub InitCPUInfo()
  GetCPUInfo
  SetCPURECTs
End Sub

Private Sub GetCPUInfo()
  Dim cCap As CPU_Capabilities
  FillCPUInfo
  ReDim cpuRECTs(1 To TotalCores) As RECT
  ReDim CPULoad(TotalCores - 1) As Double
  ReDim oldCPULoad(TotalCores - 1) As Double
  ReDim CPUBusy(1 To TotalCores) As Boolean
  ReDim CPUHibernating(1 To TotalCores) As Boolean
  cCap = GetCPUCapabilities()
  m_DisplayCPUName = CPUInfo.Name & " (" & GetCPUTechString(cCap) & ")"
End Sub

Private Function GetCPUTechString(TechCapabilities As CPU_Capabilities) As String
  Select Case TechCapabilities
    Case CPU_Capabilities.ccAVX_YMM
      GetCPUTechString = "AVX/YMM"
    Case CPU_Capabilities.ccAVX
      GetCPUTechString = "AVX"
    Case CPU_Capabilities.ccSSE2
      GetCPUTechString = "SSE2"
    Case CPU_Capabilities.ccLegacy
      GetCPUTechString = "Legacy"
  End Select
End Function

Private Function GetCPURectsTotalSize() As POINTAPI
  Dim i As Long, w As Long, h As Long
  For i = 1 To TotalCores
    With cpuRECTs(i)
      If .Right > w Then w = .Right
      If .Bottom > h Then h = .Bottom
    End With
  Next
  With GetCPURectsTotalSize
    .X = w
    .Y = h
  End With
End Function

Private Sub SetFixedWindowSize()
  With GetCPURectsTotalSize
    FixedWindowSize.X = (.X + (Screen.TwipsPerPixelX * KERNEL_PIXEL_SPACING))
    FixedWindowSize.Y = (.Y + (Screen.TwipsPerPixelY * KERNEL_PIXEL_SPACING))
  End With
End Sub

Private Sub DoAutoSize()
  SetCPURECTs
  If m_AutoSize = FixedSize Then
    SetFixedWindowSize
    UserControl_Resize
  Else
    Refresh True
  End If
End Sub

Private Sub ShiftAutoUpdate()
  gHasData = False
  m_HaveBaseline = False
  ClearCPULoads
  If m_AutoUpdate = True Then
    StartUpdateTimer
  Else
    KillUpdateTimer
    DrawCPURects True
  End If
End Sub

Private Sub StartUpdateTimer()
  Set UpdateTimer = New CPUTimer
  UpdateTimer.Interval = m_UpdateInterval
  UpdateTimer.Enabled = True
End Sub

Private Sub KillUpdateTimer()
  If UpdateTimer Is Nothing Then Exit Sub
  UpdateTimer.Enabled = False
  Set UpdateTimer = Nothing
End Sub

Private Function CPUIndexAt(ByVal X As Single, ByVal Y As Single) As Long
  Dim i As Long
  For i = 1 To TotalCores
    With cpuRECTs(i)
      If (X >= .Left And X < .Right) Then
        If (Y >= .Top And Y < .Bottom) Then
          CPUIndexAt = i
          Exit Function
        End If
      End If
    End With
  Next
End Function

Private Function IMouseSink_HitTest(ByVal Token As Long, ByVal X As Long, ByVal Y As Long) As Long
  If UserControl.Enabled Then IMouseSink_HitTest = CPUIndexAt(CSng(X) * Screen.TwipsPerPixelX, CSng(Y) * Screen.TwipsPerPixelY)
End Function

Private Sub IMouseSink_OnMouse(ByVal Token As Long, ByVal Kind As MouseEventKind, _
    ByVal Region As Long, ByVal Button As Integer, ByVal Buttons As Integer, _
    ByVal Shift As Integer, ByVal X As Long, ByVal Y As Long, ByVal Reason As MouseCancelReason)
  Dim OldHover As Long, OldPress As Long
  If Token <> m_MouseToken Then Exit Sub
  Select Case Kind
    Case meEnter
      m_IsHovering = True
      m_HoverIndex = Region
      DrawCPURect Region, True
    Case meLeave
      m_IsHovering = False
      m_HoverIndex = 0
      DrawCPURect Region, True
    Case meDown
      If Button <> vbLeftButton Or Region < 1 Or Region > TotalCores Then Exit Sub
      m_MouseIsDown = True
      m_IsPressed = True
      m_PressedIndex = Region
      DrawCPURect Region, True
    Case meUp
      If Button <> vbLeftButton Then Exit Sub
      OldPress = m_PressedIndex
      m_MouseIsDown = False
      m_IsPressed = False
      m_PressedIndex = 0
      DrawCPURect OldPress, True
    Case meCancel
      OldHover = m_HoverIndex
      OldPress = m_PressedIndex
      ClearMouseState
      DrawCPURect OldHover, True
      If OldPress <> OldHover Then DrawCPURect OldPress, True
    Case meClick
      If Region >= 1 And Region <= TotalCores Then RaiseEvent Click(Region)
      ' A recipient may hide or unload this control; do not touch it afterward.
  End Select
End Sub

Private Sub ClearMouseState()
  m_IsHovering = False
  m_HoverIndex = 0
  m_MouseIsDown = False
  m_IsPressed = False
  m_PressedIndex = 0
End Sub

Private Sub DetachMouse()
  Dim Token As Long
  Token = m_MouseToken
  m_MouseToken = 0
  MouseUnregister Token
  ClearMouseState
End Sub

Private Sub UpdateTimer_Timer()
  UpdateTimer.Enabled = False
  UpdateView
  If m_AutoUpdate = True Then
    If Not UpdateTimer Is Nothing Then UpdateTimer.Enabled = True
  End If
End Sub

Private Sub UserControl_Initialize()
  InitCPUInfo
  ' Measure the headings before laying out and sizing the CPU regions.
  Refresh True
  SetFixedWindowSize
  ShiftAutoUpdate
End Sub

Private Sub UserControl_InitProperties()
  m_AutoUpdate = False
  m_UpdateInterval = CPU_UPDATE_INTERVAL
  m_AutoSize = CPUView_Autosize.ViewAutoSize
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' The IDE uses ordinary VB6 input; compiled input comes from IMouseSink.
  If MouseNativeAvailable() Or Not UserControl.Enabled Then Exit Sub
  If Button = vbLeftButton Then
    m_PressedIndex = CPUIndexAt(X, Y)
    If m_PressedIndex > 0 Then
      m_MouseIsDown = True
      m_IsPressed = True
      DrawCPURect m_PressedIndex, True
    End If
  End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim OldPress As Long
  If MouseNativeAvailable() Or Not UserControl.Enabled Then Exit Sub
  If (Button = vbLeftButton And m_MouseIsDown = True) Then
    OldPress = m_PressedIndex
    ClearMouseState
    DrawCPURect OldPress, True
    If OldPress > 0 And OldPress = CPUIndexAt(X, Y) Then
      If X >= 0 And Y >= 0 And X < UserControl.ScaleWidth And Y < UserControl.ScaleHeight Then RaiseEvent Click(OldPress)
    End If
  End If
End Sub

Private Sub UserControl_Show()
  Refresh True
  If UserControl.Ambient.UserMode And MouseNativeAvailable() Then
    If m_MouseToken = 0 Then
      m_MouseToken = MouseRegister(UserControl.hWnd, Me)
      If m_MouseToken = 0 Then Err.Raise vbObjectError + 2102, "CPUView", MouseLastError()
    End If
    MouseRefresh m_MouseToken
  End If
End Sub

Private Sub UserControl_Hide()
  DetachMouse
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  m_AutoUpdate = PropBag.ReadProperty(PROPNAME_AUTOUPDATE, False)
  m_UpdateInterval = PropBag.ReadProperty(PROPNAME_UPDATEINTERVAL, CPU_UPDATE_INTERVAL)
  m_AutoSize = PropBag.ReadProperty(PROPNAME_AUTOSIZE, CPUView_Autosize.ViewAutoSize)
  DoAutoSize
End Sub

Private Sub UserControl_Resize()
  If m_AutoSize = FixedSize Then
    If UserControl.ScaleWidth <> FixedWindowSize.X Then UserControl.Width = ((UserControl.Width - UserControl.ScaleWidth) + FixedWindowSize.X): Exit Sub
    If UserControl.ScaleHeight <> FixedWindowSize.Y Then UserControl.Height = ((UserControl.Height - UserControl.ScaleHeight) + FixedWindowSize.Y): Exit Sub
  ElseIf m_AutoSize = ViewAutoSize Then
    DoAutoSize
  End If
  Refresh True
End Sub

Private Sub UserControl_Terminate()
  DetachMouse
  KillUpdateTimer
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  PropBag.WriteProperty PROPNAME_AUTOUPDATE, m_AutoUpdate, False
  PropBag.WriteProperty PROPNAME_UPDATEINTERVAL, m_UpdateInterval, CPU_UPDATE_INTERVAL
  PropBag.WriteProperty PROPNAME_AUTOSIZE, m_AutoSize, CPUView_Autosize.ViewAutoSize
End Sub
