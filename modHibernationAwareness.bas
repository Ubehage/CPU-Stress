Attribute VB_Name = "modHibernationAwareness"
Option Explicit

Private Const GWL_WNDPROC = -4
Private Const WM_POWERBROADCAST = &H218 'Power state is changing
Private Const WM_WINDOWPOSCHANGED As Long = &H47

Private Const PBT_APMSUSPEND As Long = &H4 'Windows is hibernating
Private Const PBT_APMRESUMEAUTOMATIC = &H12 'Windows has resumed from suspended state.

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Dim OldWndProc As Long
Dim HookhWnd As Long
Dim IsWorker As Boolean

Public Sub HookWindowProc(hWnd As Long, Optional WorkerHook As Boolean = False)
  OldWndProc = SetWindowLong(hWnd, GWL_WNDPROC, AddressOf WndProc)
  HookhWnd = hWnd
  IsWorker = WorkerHook
End Sub

Public Sub UnhookWindowProc()
  If (OldWndProc <> 0 And HookhWnd <> 0) Then
    Call SetWindowLong(HookhWnd, GWL_WNDPROC, OldWndProc)
    OldWndProc = 0
    HookhWnd = 0
  End If
End Sub

Public Function WndProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  Dim PreviousProc As Long
  PreviousProc = OldWndProc
  If uMsg = WM_POWERBROADCAST Then
    Dim hVal As Boolean, DoCall As Boolean
    If wParam = PBT_APMSUSPEND Then
      hVal = True
      DoCall = True
    ElseIf wParam = PBT_APMRESUMEAUTOMATIC Then
      hVal = False
      DoCall = True
    End If
    If DoCall = True Then If IsWorker = False Then Call frmMain.HibernationTrigger(hVal)
  End If
  WndProc = CallWindowProc(PreviousProc, hWnd, uMsg, wParam, lParam)
  ' Child windows do not receive their parent's move message. Reevaluate
  ' pointer geometry after default layout processing, using our existing hook.
  If uMsg = WM_WINDOWPOSCHANGED And IsWorker = False Then MouseRefreshAll
End Function
