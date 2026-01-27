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

Private Const PROPNAME_AUTOUPDATE = "AutoUpdate"

Private Const CPUVIEW_TITLE = "CPU Overview"

Private Const KERNEL_PIXEL_WIDTH As Long = 25
Private Const KERNEL_PIXEL_HEIGHT As Long = KERNEL_PIXEL_WIDTH
Private Const KERNEL_PIXEL_SPACING As Long = 7

Private Const PDH_FMT_DOUBLE = &H200

Private Type PDH_FMT_COUNTERVALUE
  CStatus As Long
  Padding As Long
  DoubleValue As Double
End Type

Private Declare Function PdhOpenQuery Lib "pdh.dll" Alias "PdhOpenQueryA" (ByVal dataSource As String, ByVal userData As Long, query As Long) As Long
Private Declare Function PdhAddCounter Lib "pdh.dll" Alias "PdhAddEnglishCounterA" (ByVal query As Long, ByVal counterPath As String, ByVal userData As Long, ByRef counter As Long) As Long
Private Declare Function PdhCollectQueryData Lib "pdh.dll" (ByVal query As Long) As Long
Private Declare Function PdhGetFormattedCounterValue Lib "pdh.dll" (ByVal counter As Long, ByVal format As Long, lpdwType As Long, ByRef value As PDH_FMT_COUNTERVALUE) As Long
Private Declare Function PdhCloseQuery Lib "pdh.dll" (ByVal query As Long) As Long
    
Dim m_AutoUpdate As Boolean

Dim CPUInfo As CPU_Info
Dim TotalCores As Long
Dim cpuRECTs() As RECT
Dim CPULoad() As Double
Dim oldCPULoad() As Double

Dim m_IsCapturing As Boolean
Dim m_IsHovering As Boolean
Dim m_HoverIndex As Long
Dim m_IsPressed As Boolean
Dim m_MouseIsDown As Boolean

Dim FixedWindowSize As POINTAPI

Dim gQuery As Long
Dim gCounters() As Long
Dim gHasData As Boolean

Dim WithEvents UpdateTimer As CPUTimer
Attribute UpdateTimer.VB_VarHelpID = -1

Public Event Click(Index As Long)

Public Property Get AutoUpdate() As Boolean
  AutoUpdate = m_AutoUpdate
End Property
Public Property Let AutoUpdate(New_AutoUpdate As Boolean)
  If m_AutoUpdate = New_AutoUpdate Then Exit Property
  m_AutoUpdate = New_AutoUpdate
  ShiftAutoUpdate
End Property

Public Sub Refresh(Optional FullRefresh As Boolean = False)
  If FullRefresh = True Then
    UserControl.Cls
    DrawBorderAndTitle
  End If
  DrawCPURects
End Sub

Public Sub UpdateView()
  GetCPULoads
  Dim i As Long
  For i = 0 To (TotalCores - 1)
    If CPULoad(i) <> oldCPULoad(i) Then
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
  With UserControl
    .CurrentX = (Screen.TwipsPerPixelX * 10)
    .CurrentY = Screen.TwipsPerPixelY
  End With
  UserControl.Print CPUVIEW_TITLE
End Sub

Private Sub DrawCPURects()
  Dim i As Long
  For i = 1 To TotalCores
    DrawCPURect i
  Next
End Sub

Private Sub DrawCPURect(RectIndex As Long, Optional FullRedraw As Boolean = False)
  If (RectIndex <= 0 Or RectIndex > TotalCores) Then Exit Sub
  With cpuRECTs(RectIndex)
    If FullRedraw = True Then UserControl.Line (.Left, .Top)-(.Right, .Bottom), GetCPUBackColor(RectIndex), BF
    If gHasData Then DrawCPUFlood RectIndex
    UserControl.Line (.Left, .Top)-(.Right, .Bottom), COLOR_OUTLINE_LIGHT, B
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

Private Function GetCPUFloodHeight(CPUIndex As Long) As Long
  Dim r As Long
  r = ((KERNEL_PIXEL_HEIGHT / 100) * CPULoad(CPUIndex - 1))
  If r > KERNEL_PIXEL_HEIGHT Then r = KERNEL_PIXEL_HEIGHT
  GetCPUFloodHeight = (r * Screen.TwipsPerPixelY)
End Function

Private Function GetCPUFloodColor(CPUIndex As Long) As Long
  Dim r As Long
  Select Case CPULoad(CPUIndex - 1)
    Case Is >= 85
      If m_HoverIndex = CPUIndex Then
        If m_MouseIsDown Then r = COLOR_RED_PRESSED Else r = COLOR_RED_HOVER
      Else
        r = COLOR_RED
      End If
    Case Is >= 60
      If m_HoverIndex = CPUIndex Then
        If m_MouseIsDown Then r = COLOR_YELLOW_PRESSED Else r = COLOR_YELLOW_HOVER
      Else
        r = COLOR_YELLOW
      End If
    Case Else
      If m_HoverIndex = CPUIndex Then
        If m_MouseIsDown Then r = COLOR_GREEN_PRESSED Else r = COLOR_GREEN_HOVER
      Else
        r = COLOR_GREEN
      End If
  End Select
  GetCPUFloodColor = r
End Function

Private Function GetCPUBackColor(CPUIndex As Long) As Long
  If m_HoverIndex = CPUIndex Then
    If m_MouseIsDown Then GetCPUBackColor = COLOR_BUTTON_PRESSED Else GetCPUBackColor = COLOR_BUTTON_HOVER
  Else
    GetCPUBackColor = COLOR_BACKGROUND
  End If
End Function

Private Sub GetCPULoads()
  Dim i As Long, v As PDH_FMT_COUNTERVALUE
  If gHasData = False Then
    ReDim gCounters(TotalCores - 1) As Long
    Call PdhOpenQuery(vbNullString, 0, gQuery)
    For i = 0 To (TotalCores - 1)
      Call PdhAddCounter(gQuery, "\Processor(" & CStr(i) & ")\% Processor Time", 0, gCounters(i))
    Next
  End If
  Call PdhCollectQueryData(gQuery)
  If gHasData = False Then
    gHasData = True
  Else
    For i = 0 To (TotalCores - 1)
      Call PdhGetFormattedCounterValue(gCounters(i), PDH_FMT_DOUBLE, 0, v)
      If v.CStatus = 0 Then CPULoad(i) = CSng(v.DoubleValue)
    Next
  End If
End Sub

Private Sub SetCPURECTs()
  Dim i As Long, j As Long, c As Long
  Dim cCols As Long, cRows As Long, tC As Long, tR As Long
  Dim X As Long, Y As Long, w As Long, h As Long
  cCols = TotalCores
  cRows = 1
  Do
    tC = (cCols / 2)
    tR = (cRows * 2)
    If tR > tC Then Exit Do
    cCols = tC
    cRows = tR
  Loop
  Y = (Screen.TwipsPerPixelY * 23)
  w = (Screen.TwipsPerPixelX * KERNEL_PIXEL_WIDTH)
  h = (Screen.TwipsPerPixelX * KERNEL_PIXEL_HEIGHT)
  For i = 1 To cRows
    X = (Screen.TwipsPerPixelX * 5)
    For j = 1 To cCols
      c = (c + 1)
      If c > TotalCores Then Exit For
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
End Sub

Private Sub InitCPUInfo()
  GetCPUInfo
  SetCPURECTs
End Sub

Private Sub GetCPUInfo()
  CPUInfo = GetCPUCoreCount()
  TotalCores = CountAllCores(CPUInfo)
  ReDim cpuRECTs(1 To TotalCores) As RECT
  ReDim CPULoad(TotalCores - 1) As Double
  ReDim oldCPULoad(TotalCores - 1) As Double
End Sub

Private Function GetCPURectSize() As POINTAPI
  Dim i As Long, w As Long, h As Long
  For i = 1 To TotalCores
    With cpuRECTs(i)
      If .Right > w Then w = .Right
      If .Bottom > h Then h = .Bottom
    End With
  Next
  With GetCPURectSize
    .X = w
    .Y = h
  End With
End Function

Private Sub SetFixedWindowSize()
  With GetCPURectSize
    FixedWindowSize.X = (.X + (Screen.TwipsPerPixelX * KERNEL_PIXEL_SPACING))
    FixedWindowSize.Y = (.Y + (Screen.TwipsPerPixelY * KERNEL_PIXEL_SPACING))
  End With
End Sub

Private Sub ShiftAutoUpdate()
  If m_AutoUpdate = True Then StartUpdateTimer Else KillUpdateTimer
End Sub

Private Sub StartUpdateTimer()
  Set UpdateTimer = New CPUTimer
  UpdateTimer.Interval = 500
  UpdateTimer.Enabled = True
End Sub

Private Sub KillUpdateTimer()
  If UpdateTimer Is Nothing Then Exit Sub
  UpdateTimer.Enabled = False
  Set UpdateTimer = Nothing
End Sub

Private Function GetCPUIndexFromMousePos() As Long
  Dim i As Long, p As POINTAPI, r As Long
  Call GetCursorPos(p)
  Call ScreenToClient(UserControl.hWnd, p)
  p.X = (p.X * Screen.TwipsPerPixelX)
  p.Y = (p.Y * Screen.TwipsPerPixelY)
  For i = 1 To TotalCores
    With cpuRECTs(i)
      If (p.X >= .Left And p.X <= .Right) Then
        If (p.Y >= .Top And p.Y <= .Bottom) Then
          r = i
          Exit For
        End If
      End If
    End With
  Next
  GetCPUIndexFromMousePos = r
End Function

Private Sub StartHover(Optional DoNotRefresh As Boolean = False, Optional ForceRefresh As Boolean = False)
  Dim hIndex As Long, t As Long, r As Boolean
  hIndex = GetCPUIndexFromMousePos()
  If hIndex <> m_HoverIndex Then
    t = m_HoverIndex
    m_HoverIndex = hIndex
    If DoNotRefresh = False Then r = True
  End If
  If ForceRefresh = True Then r = True
  If r = True Then
    DrawCPURect t, True
    DrawCPURect m_HoverIndex, True
  End If
  If m_IsCapturing = True Then Exit Sub
  Call SetCapture(UserControl.hWnd)
  m_IsCapturing = True
End Sub

Private Sub EndHover(Optional DoNotRefresh As Boolean = False, Optional ForceRefresh As Boolean = False)
  Dim r As Boolean
  If m_IsHovering = True Then
    m_IsHovering = False
    If DoNotRefresh = False Then r = True
  End If
  If (r = True Or ForceRefresh = True) Then DrawCPURect m_HoverIndex
  If (m_IsCapturing = False Or m_MouseIsDown = True) Then Exit Sub
  EndCapture
End Sub

Private Sub EndCapture()
  If m_IsCapturing = True Then
    Call ReleaseCapture
    m_IsCapturing = False
  End If
End Sub

Private Sub UpdateTimer_Timer()
  UpdateTimer.Enabled = False
  UpdateView
  UpdateTimer.Enabled = True
End Sub

Private Sub UserControl_Initialize()
  InitCPUInfo
  SetFixedWindowSize
  ShiftAutoUpdate
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = vbLeftButton Then
    m_MouseIsDown = True
    m_IsPressed = True
    StartHover , True
  End If
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  StartHover
  If m_IsCapturing = False Then Exit Sub
  If IsCursorOnWindow(UserControl.hWnd, m_MouseIsDown) = False Then EndHover
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim HadCapture As Boolean
  If (Button = vbLeftButton And m_MouseIsDown = True) Then
    HadCapture = m_IsCapturing
    m_MouseIsDown = False
    m_IsPressed = False
    DrawCPURect m_HoverIndex, True
    EndHover True
    If IsCursorOnWindow(UserControl.hWnd, False) = True Then
      RaiseEvent Click(GetCPUIndexFromMousePos)
      If HadCapture = True Then StartHover True
    End If
  End If
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  m_AutoUpdate = PropBag.ReadProperty(PROPNAME_AUTOUPDATE, False)
End Sub

Private Sub UserControl_Resize()
  If UserControl.ScaleWidth <> FixedWindowSize.X Then UserControl.Width = ((UserControl.Width - UserControl.ScaleWidth) + FixedWindowSize.X): Exit Sub
  If UserControl.ScaleHeight <> FixedWindowSize.Y Then UserControl.Height = ((UserControl.Height - UserControl.ScaleHeight) + FixedWindowSize.Y): Exit Sub
  Refresh True
End Sub

Private Sub UserControl_Terminate()
  Call PdhCloseQuery(gQuery)
  KillUpdateTimer
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  PropBag.WriteProperty PROPNAME_AUTOUPDATE, m_AutoUpdate, False
End Sub
