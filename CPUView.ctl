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

Dim FixedWindowSize As POINTAPI

Dim gQuery As Long
Dim gCounters() As Long
Dim gHasData As Boolean

Dim WithEvents UpdateTimer As CPUTimer
Attribute UpdateTimer.VB_VarHelpID = -1

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
  With cpuRECTs(RectIndex)
    If FullRedraw = True Then UserControl.Line (.Left, .Top)-(.Right, .Bottom), COLOR_BACKGROUND, BF
    If gHasData Then DrawCPUFlood RectIndex
    UserControl.Line (.Left, .Top)-(.Right, .Bottom), COLOR_OUTLINE_LIGHT, B
  End With
End Sub

Private Sub DrawCPUFlood(CPUIndex As Long)
  Dim h As Long
  h = GetCPUFloodHeight(CPULoad((CPUIndex - 1)))
  If h < 1 Then Exit Sub
  With cpuRECTs(CPUIndex)
    UserControl.Line (.Left, (.Bottom - h))-(.Right, .Bottom), GetCPUFloodColor(CPULoad((CPUIndex - 1))), BF
  End With
End Sub

Private Function GetCPUFloodHeight(LoadPercent As Double) As Long
  Dim r As Long
  r = ((KERNEL_PIXEL_HEIGHT / 100) * LoadPercent)
  If r > KERNEL_PIXEL_HEIGHT Then r = KERNEL_PIXEL_HEIGHT
  GetCPUFloodHeight = (r * Screen.TwipsPerPixelY)
End Function

Private Function GetCPUFloodColor(LoadPercent As Double) As Long
  Select Case LoadPercent
    Case Is >= 85
      GetCPUFloodColor = COLOR_RED
    Case Is >= 60
      GetCPUFloodColor = COLOR_YELLOW
    Case Else
      GetCPUFloodColor = COLOR_GREEN
  End Select
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
  Dim x As Long, y As Long, w As Long, h As Long
  cCols = TotalCores
  cRows = 1
  Do
    tC = (cCols / 2)
    tR = (cRows * 2)
    If tR > tC Then Exit Do
    cCols = tC
    cRows = tR
  Loop
  y = (Screen.TwipsPerPixelY * 23)
  w = (Screen.TwipsPerPixelX * KERNEL_PIXEL_WIDTH)
  h = (Screen.TwipsPerPixelX * KERNEL_PIXEL_HEIGHT)
  For i = 1 To cRows
    x = (Screen.TwipsPerPixelX * 5)
    For j = 1 To cCols
      c = (c + 1)
      If c > TotalCores Then Exit For
      With cpuRECTs(c)
        .Left = x
        .Top = y
        .Right = (.Left + w)
        .Bottom = (.Top + h)
        x = (.Right + (Screen.TwipsPerPixelX * KERNEL_PIXEL_SPACING))
      End With
    Next
    y = (cpuRECTs(c).Bottom + (Screen.TwipsPerPixelY * KERNEL_PIXEL_SPACING))
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
    .x = w
    .y = h
  End With
End Function

Private Sub SetFixedWindowSize()
  With GetCPURectSize
    FixedWindowSize.x = (.x + (Screen.TwipsPerPixelX * KERNEL_PIXEL_SPACING))
    FixedWindowSize.y = (.y + (Screen.TwipsPerPixelY * KERNEL_PIXEL_SPACING))
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

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  m_AutoUpdate = PropBag.ReadProperty(PROPNAME_AUTOUPDATE, False)
End Sub

Private Sub UserControl_Resize()
  If UserControl.ScaleWidth <> FixedWindowSize.x Then UserControl.Width = ((UserControl.Width - UserControl.ScaleWidth) + FixedWindowSize.x): Exit Sub
  If UserControl.ScaleHeight <> FixedWindowSize.y Then UserControl.Height = ((UserControl.Height - UserControl.ScaleHeight) + FixedWindowSize.y): Exit Sub
  Refresh True
End Sub

Private Sub UserControl_Terminate()
  Call PdhCloseQuery(gQuery)
  KillUpdateTimer
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  PropBag.WriteProperty PROPNAME_AUTOUPDATE, m_AutoUpdate, False
End Sub
