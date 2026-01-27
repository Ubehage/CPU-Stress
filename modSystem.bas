Attribute VB_Name = "modSystem"
Option Explicit

Global Const INVALID_HANDLE_VALUE As Long = -1&

Private Const HWND_NOTOPMOST  As Long = -2
Private Const HWND_TOPMOST  As Long = -1
Private Const SWP_NOMOVE  As Long = &H2
Private Const SWP_NOSIZE  As Long = &H1
Private Const SWP_SETWINDOWPOS  As Long = SWP_NOSIZE Or SWP_NOMOVE

Private Const ICC_LISTVIEW_CLASSES  As Long = &H1
Private Const ICC_TREEVIEW_CLASSES  As Long = &H2
Private Const ICC_BAR_CLASSES  As Long = &H4
Private Const ICC_TAB_CLASSES  As Long = &H8
Private Const ICC_UPDOWN_CLASS  As Long = &H10
Private Const ICC_PROGRESS_CLASS  As Long = &H20
Private Const ICC_HOTKEY_CLASS  As Long = &H40
Private Const ICC_ANIMATE_CLASS  As Long = &H80
Private Const ICC_WIN95_CLASSES  As Long = &HFF
Private Const ICC_DATE_CLASSES  As Long = &H100
Private Const ICC_USEREX_CLASSES  As Long = &H200
Private Const ICC_COOL_CLASSES  As Long = &H400
Private Const ICC_INTERNET_CLASSES  As Long = &H800
Private Const ICC_PAGESCROLLER_CLASS  As Long = 1000
Private Const ICC_NATIVEFNTCTL_CLASS  As Long = 2000
Private Const ICC_STANDARD_CLASSES  As Long = 4000
Private Const ICC_LINK_CLASS  As Long = 8000

Public Enum COMMONCONTROLS_CLASSES
  ccListView_Classes = ICC_LISTVIEW_CLASSES
  ccTreeView_Classes = ICC_TREEVIEW_CLASSES
  ccToolBar_Classes = ICC_BAR_CLASSES
  ccTab_Classes = ICC_TAB_CLASSES
  ccUpDown_Classes = ICC_UPDOWN_CLASS
  ccProgress_Class = ICC_PROGRESS_CLASS
  ccHotkey_Class = ICC_HOTKEY_CLASS
  ccAnimate_Class = ICC_ANIMATE_CLASS
  ccWin95_Classes = ICC_WIN95_CLASSES
  ccCalendar_Classes = ICC_DATE_CLASSES
  ccComboEx_Classes = ICC_USEREX_CLASSES
  ccCoolBar_Classes = ICC_COOL_CLASSES
  ccInternet_Classes = ICC_INTERNET_CLASSES
  ccPageScroller_Class = ICC_PAGESCROLLER_CLASS
  ccNativeFont_Class = ICC_NATIVEFNTCTL_CLASS
  ccStandard_Classes = ICC_STANDARD_CLASSES
  ccLink_Class = ICC_LINK_CLASS
  ccAll_Classes = ccListView_Classes Or ccTreeView_Classes Or ccToolBar_Classes Or ccTab_Classes Or ccUpDown_Classes Or ccProgress_Class Or ccHotkey_Class Or ccAnimate_Class Or ccWin95_Classes Or ccCalendar_Classes Or ccComboEx_Classes Or ccCoolBar_Classes Or ccInternet_Classes Or ccPageScroller_Class Or ccNativeFont_Class Or ccStandard_Classes Or ccLink_Class
End Enum

Public Type POINTAPI
  X As Long
  Y As Long
End Type

Public Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

Public Type CPU_Info
  PhysicalCores As Long
  KernelsPerCore() As Long
End Type

Private Type tagINITCOMMONCONTROLSEX
  dwSize As Long
  dwICC As Long
End Type

Private Type SYSTEM_LOGICAL_PROCESSOR_INFORMATION
    ProcessorMask As Long
    Relationship As Long ' 0 = Core, 1 = NUMA, 2 = Cache, 3 = Package
    Reserved(1) As Currency
End Type

Private Declare Function SetWindowPos Lib "user32.dll" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cX As Long, ByVal cY As Long, ByVal wFlags As Long) As Long
Private Declare Sub InitCommonControls9x Lib "comctl32" Alias "InitCommonControls" ()
Private Declare Function InitCommonControlsEx Lib "comctl32" (lpInitCtrls As tagINITCOMMONCONTROLSEX) As Boolean

Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long

Public Declare Sub GetWindowRect Lib "user32" (ByVal hWnd As Long, ByRef WindowRect As RECT)
Public Declare Function WindowFromPoint Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Public Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, ByRef lpPoint As POINTAPI) As Long
Public Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long

Private Declare Function GetLogicalProcessorInformation Lib "kernel32" (ByRef Buffer As Any, ByRef ReturnLength As Long) As Long

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef hpvDest As Any, ByRef hpvSource As Any, ByVal cbCopy As Long)

Public Sub WindowOnTop(hWnd As Long, OnTop As Boolean)
  Dim wFlags As Long
  If OnTop Then
    wFlags = HWND_TOPMOST
  Else
    wFlags = HWND_NOTOPMOST
  End If
  SetWindowPos hWnd, wFlags, 0&, 0&, 0&, 0&, SWP_SETWINDOWPOS
End Sub

Public Function InitCommonControls(Optional ccFlags As COMMONCONTROLS_CLASSES = ccAll_Classes) As Boolean
  Dim icc As tagINITCOMMONCONTROLSEX
  On Error GoTo OldCC
  With icc
    .dwSize = Len(icc)
    .dwICC = ccFlags
  End With
  InitCommonControls = InitCommonControlsEx(icc)
ExitNow:
  On Error GoTo 0
  Exit Function
OldCC:
  InitCommonControls9x
  Resume ExitNow
End Function

Public Function IsPointInRect(pRect As RECT, pPoint As POINTAPI) As Boolean
  With pRect
    If (pPoint.X >= .Left And pPoint.X <= .Right) Then
      If (pPoint.Y >= .Top And pPoint.Y <= .Bottom) Then IsPointInRect = True
    End If
  End With
End Function

Public Function IsCursorOnWindow(hWnd As Long, Optional MouseIsDown As Boolean = False) As Boolean
  Dim mPos As POINTAPI, wRect As RECT, hTop As Long
  Call GetCursorPos(mPos)
  Call GetWindowRect(hWnd, wRect)
  If IsPointInRect(wRect, mPos) = False Then If MouseIsDown = False Then Exit Function
  hTop = WindowFromPoint(mPos.X, mPos.Y)
  If hTop <> hWnd Then If MouseIsDown = False Then Exit Function
  IsCursorOnWindow = True
End Function

Public Function GetCPUCoreCount() As CPU_Info
  Dim i As Long
  Dim bBuffer() As Byte, cbBuffer As Long, nEntries As Long, structSize As Long, info As SYSTEM_LOGICAL_PROCESSOR_INFORMATION
  Dim bOffset As Long
  Dim r As CPU_Info
  structSize = LenB(info)
  Call GetLogicalProcessorInformation(ByVal 0&, cbBuffer)
  If cbBuffer = 0 Then Exit Function
  ReDim bBuffer(cbBuffer - 1) As Byte
  If GetLogicalProcessorInformation(bBuffer(0), cbBuffer) = 0 Then Exit Function
  nEntries = (cbBuffer / structSize)
  For i = 0 To (nEntries - 1)
    bOffset = (i * structSize)
    CopyMemory info, bBuffer(bOffset), structSize
    Select Case info.Relationship
      Case 0
        If (GetCPUCoreCount.PhysicalCores Mod 5) = 0 Then ReDim Preserve GetCPUCoreCount.KernelsPerCore(1 To (GetCPUCoreCount.PhysicalCores + 5)) As Long
        GetCPUCoreCount.PhysicalCores = (GetCPUCoreCount.PhysicalCores + 1)
        GetCPUCoreCount.KernelsPerCore(GetCPUCoreCount.PhysicalCores) = CountKernelBits(info.ProcessorMask)
    End Select
  Next
End Function

Public Function CountAllCores(cInfo As CPU_Info) As Long
  Dim i As Long, r As Long
  For i = 1 To cInfo.PhysicalCores
    r = (r + cInfo.KernelsPerCore(i))
  Next
  CountAllCores = r
End Function

Private Function CountKernelBits(ByVal kMask As Long) As Long
  Dim i As Long, bMask As Long, r As Long
  bMask = 1
  Do
    If (kMask And bMask) <> 0 Then r = (r + 1)
    If i = 30 Then Exit Do
    bMask = (bMask * 2)
    i = (i + 1)
  Loop
  If (kMask And &H80000000) <> 0 Then r = (r + 1)
  CountKernelBits = r
End Function
