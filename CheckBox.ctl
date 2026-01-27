VERSION 5.00
Begin VB.UserControl CheckBox 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00202020&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   BeginProperty Font 
      Name            =   "Consolas"
      Size            =   11.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   3600
   ScaleWidth      =   4800
End
Attribute VB_Name = "CheckBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const PROPNAME_VALUE As String = "Value"
Private Const PROPNAME_CAPTION As String = "Caption"
Private Const PROPNAME_ENABLED As String = "Enabled"

Dim m_Value As CheckBoxConstants
Dim m_Caption As String
Dim m_Enabled As Boolean

Dim m_IsHovering As Boolean

Dim m_ScreenRect As RECT
Dim m_IsCapturing As Boolean
Dim m_Hovering As Boolean
Dim m_IsPressed As Boolean
Dim m_MouseIsDown As Boolean
Dim m_KeyIsDown As Boolean
Dim m_HasFocus As Boolean

Dim BoxRect As RECT
Dim TextRect As RECT

Public Event Click()

Public Property Get Value() As CheckBoxConstants
  Value = m_Value
End Property
Public Property Get Caption() As String
  Caption = m_Caption
End Property
Public Property Get Enabled() As Boolean
  Enabled = m_Enabled
End Property
Public Property Let Value(New_Value As CheckBoxConstants)
  If m_Value = New_Value Then Exit Property
  m_Value = New_Value
  UserControl.PropertyChanged PROPNAME_VALUE
  Refresh
  If ChangedByCode = False Then RaiseEvent Click
End Property
Public Property Let Caption(New_Caption As String)
  If m_Caption = New_Caption Then Exit Property
  m_Caption = New_Caption
  UserControl.PropertyChanged PROPNAME_CAPTION
  Refresh True
End Property
Public Property Let Enabled(New_Enabled As Boolean)
  If m_Enabled = New_Enabled Then Exit Property
  m_Enabled = New_Enabled
  UserControl.PropertyChanged PROPNAME_ENABLED
  If m_Enabled = False Then
    EndHover , True
    m_MouseIsDown = False
    m_KeyIsDown = False
    m_IsPressed = False
  Else
    Refresh
  End If
End Property

Public Sub Refresh(Optional FullRedraw As Boolean = False)
  UserControl.Cls
  If FullRedraw = True Then Call CheckWindowSize
  DrawBox
  DrawTitle
End Sub

Private Sub DrawTitle()
  With TextRect
    UserControl.CurrentX = .Left
    UserControl.CurrentY = .Top
  End With
  UserControl.ForeColor = GetTextColor()
  UserControl.Print m_Caption
End Sub

Private Sub DrawBox()
  With BoxRect
    UserControl.Line (.Left, .Top)-((.Right - Screen.TwipsPerPixelX), (.Bottom - Screen.TwipsPerPixelY)), GetBoxBackColor(), BF
  End With
  DrawBoxBorder
  If m_Value = vbChecked Then DrawBoxChecked
End Sub

Private Sub DrawBoxBorder()
  Dim c As Long
  c = GetBoxBorderColor()
  With BoxRect
    UserControl.Line (.Left, .Top)-((.Right - Screen.TwipsPerPixelX), (.Bottom - Screen.TwipsPerPixelY)), c, B
    UserControl.Line ((.Left + Screen.TwipsPerPixelX), (.Top + Screen.TwipsPerPixelY))-((.Right - (Screen.TwipsPerPixelX * 2)), (.Bottom - (Screen.TwipsPerPixelY * 2))), c, B
  End With
End Sub

Private Sub DrawBoxChecked()
  With BoxRect
    UserControl.Line ((.Left + (Screen.TwipsPerPixelX * 3)), (.Top + (Screen.TwipsPerPixelY * 3)))-((.Right - (Screen.TwipsPerPixelX * 4)), (.Bottom - (Screen.TwipsPerPixelY * 4))), GetBoxCheckColor(), BF
  End With
End Sub

Private Function GetTextColor() As Long
  If m_Enabled = False Then
    GetTextColor = COLOR_TEXT_DISABLED
  ElseIf m_Hovering Then
    GetTextColor = COLOR_TEXT_HOVER
  Else
    GetTextColor = COLOR_TEXT
  End If
End Function

Private Function GetBoxBackColor() As Long
  If m_Enabled = False Then
    GetBoxBackColor = COLOR_BACKGROUND_DISABLED
  ElseIf m_IsPressed Then
    GetBoxBackColor = COLOR_BUTTON_PRESSED
  ElseIf m_Hovering Then
    GetBoxBackColor = COLOR_BUTTON_HOVER
  Else
    GetBoxBackColor = COLOR_CONTROLS
  End If
End Function

Private Function GetBoxCheckColor() As Long
  If m_Enabled = False Then
    GetBoxCheckColor = COLOR_GREEN_DISABLED
  ElseIf m_IsPressed Then
    GetBoxCheckColor = COLOR_GREEN_PRESSED
  ElseIf m_Hovering Then
    GetBoxCheckColor = COLOR_GREEN_HOVER
  Else
    GetBoxCheckColor = COLOR_GREEN
  End If
End Function

Private Function GetBoxBorderColor() As Long
  If m_Enabled = False Then
    GetBoxBorderColor = COLOR_TEXT_DISABLED
  ElseIf m_Hovering Then
    GetBoxBorderColor = COLOR_OUTLINE_LIGHT
  Else
    GetBoxBorderColor = COLOR_OUTLINE
  End If
End Function

Private Sub SetRects()
  With BoxRect
    .Left = 0
    .Top = 0
    .Bottom = (.Top + (UserControl.TextHeight(m_Caption) + (Screen.TwipsPerPixelY * 2)))
    .Right = (.Left + (.Bottom - .Top))
    TextRect.Left = (.Right + (Screen.TwipsPerPixelX * 5))
    TextRect.Top = (.Top + Screen.TwipsPerPixelY)
  End With
  TextRect.Right = (TextRect.Left + UserControl.TextWidth(m_Caption))
End Sub

Private Function CheckWindowSize() As Boolean
  SetRects
  If UserControl.ScaleWidth <> TextRect.Right Then
    UserControl.Width = ((UserControl.Width - UserControl.ScaleWidth) + TextRect.Right)
    Exit Function
  End If
  If UserControl.ScaleHeight <> BoxRect.Bottom Then
    UserControl.Height = ((UserControl.Height - UserControl.ScaleHeight) + BoxRect.Bottom)
    Exit Function
  End If
  CheckWindowSize = True
End Function

Private Sub SetScreenRect()
  Dim r As RECT, p As POINTAPI
  Call GetClientRect(UserControl.hWnd, r)
  p.X = r.Left
  p.Y = r.Top
  Call ClientToScreen(UserControl.hWnd, p)
  With m_ScreenRect
    .Left = p.X
    .Top = p.Y
    .Right = (.Left + r.Right)
    .Bottom = (.Top + r.Bottom)
  End With
End Sub

Private Sub StartHover(Optional DoNotRefresh As Boolean = False, Optional ForceRefresh As Boolean = False)
  Dim r As Boolean
  If m_Hovering = False Then
    m_Hovering = True
    If DoNotRefresh = False Then r = True
  End If
  If (r = True Or ForceRefresh = True) Then
    SetScreenRect
    Refresh
  End If
  If m_IsCapturing = True Then Exit Sub
  Call SetCapture(UserControl.hWnd)
  m_IsCapturing = True
End Sub

Private Sub EndHover(Optional DoNotRefresh As Boolean = False, Optional ForceRefresh As Boolean = False)
  Dim r As Boolean
  If m_Hovering = True Then
    m_Hovering = False
    If DoNotRefresh = False Then r = True
  End If
  If (r = True Or ForceRefresh = True) Then Refresh
  If m_IsCapturing = False Or m_MouseIsDown = True Then Exit Sub
  EndCapture
End Sub

Private Sub EndCapture()
  If m_IsCapturing Then
    Call ReleaseCapture
    m_IsCapturing = False
  End If
End Sub

Private Function IsCursorOnButton() As Boolean
  Dim p As POINTAPI, hTop As Long
  Call GetCursorPos(p)
  If IsPointInRect(m_ScreenRect, p) = False Then If m_MouseIsDown = False Then Exit Function
  hTop = WindowFromPoint(p.X, p.Y)
  If hTop <> UserControl.hWnd Then If m_MouseIsDown = False Then Exit Function
  IsCursorOnButton = True
End Function

Private Function CanInteractNow() As Boolean
  CanInteractNow = UserControl.Ambient.UserMode And m_Enabled And UserControl.Extender.Visible And UserControl.hWnd <> 0
End Function

Private Sub DoTheClick()
  ChangedByCode = True
  If m_Value = vbChecked Then
    Value = vbUnchecked
  Else
    Value = vbChecked
  End If
  ChangedByCode = False
  RaiseEvent Click
End Sub

Private Sub UserControl_InitProperties()
  m_Caption = "CheckBox"
  m_Value = vbUnchecked
  m_Enabled = True
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If m_Enabled = False Then Exit Sub
  If (Button = vbLeftButton And m_MouseIsDown = False) Then
    m_MouseIsDown = True
    m_IsPressed = True
    StartHover , True
  End If
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If m_Enabled = False Then Exit Sub
  StartHover
  If m_IsCapturing = False Then Exit Sub
  If IsCursorOnButton() = False Then EndHover
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If m_Enabled = False Then Exit Sub
  Dim HadCapture As Boolean
  If (Button = vbLeftButton And m_MouseIsDown = True) Then
    HadCapture = m_IsCapturing
    m_MouseIsDown = False
    m_IsPressed = False
    EndHover True
    If IsCursorOnButton() = True Then
      DoTheClick
      If (CanInteractNow() And HadCapture = True) Then StartHover True
    End If
    Refresh
  End If
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  m_Caption = PropBag.ReadProperty(PROPNAME_CAPTION, "CheckBox")
  m_Value = PropBag.ReadProperty(PROPNAME_VALUE, vbUnchecked)
  m_Enabled = PropBag.ReadProperty(PROPNAME_ENABLED, True)
End Sub

Private Sub UserControl_Resize()
  If CheckWindowSize = True Then Refresh
End Sub

Private Sub UserControl_Show()
  Refresh
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  PropBag.WriteProperty PROPNAME_CAPTION, m_Caption, "CheckBox"
  PropBag.WriteProperty PROPNAME_VALUE, m_Value, vbUnchecked
  PropBag.WriteProperty PROPNAME_ENABLED, m_Enabled, True
End Sub
