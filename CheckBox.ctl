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
Implements IMouseSink

Private Const PROPNAME_VALUE As String = "Value"
Private Const PROPNAME_CAPTION As String = "Caption"
Private Const PROPNAME_ENABLED As String = "Enabled"

Dim m_Value As CheckBoxConstants
Dim m_Caption As String
Dim m_Enabled As Boolean

Dim m_MouseToken As Long
Dim m_Hovering As Boolean
Dim m_IsPressed As Boolean
Dim m_MouseIsDown As Boolean

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
    m_MouseIsDown = False
    m_IsPressed = False
    m_Hovering = False
  End If
  If m_MouseToken <> 0 Then MouseSetEnabled m_MouseToken, m_Enabled
  Refresh
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
  ElseIf m_Hovering Or m_MouseIsDown Then
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
  ElseIf m_Hovering Or m_MouseIsDown Then
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

Private Function IMouseSink_HitTest(ByVal Token As Long, ByVal X As Long, ByVal Y As Long) As Long
  If m_Enabled Then IMouseSink_HitTest = 1
End Function

Private Sub IMouseSink_OnMouse(ByVal Token As Long, ByVal Kind As MouseEventKind, _
    ByVal Region As Long, ByVal Button As Integer, ByVal Buttons As Integer, _
    ByVal Shift As Integer, ByVal X As Long, ByVal Y As Long, ByVal Reason As MouseCancelReason)
  If Token <> m_MouseToken Then Exit Sub
  Select Case Kind
    Case meEnter
      m_Hovering = True
    Case meLeave
      m_Hovering = False
    Case meDown
      If Button <> vbLeftButton Then Exit Sub
      m_MouseIsDown = True
      m_IsPressed = True
    Case meUp
      If Button <> vbLeftButton Then Exit Sub
      m_MouseIsDown = False
      m_IsPressed = False
    Case meCancel
      m_MouseIsDown = False
      m_IsPressed = False
      m_Hovering = False
    Case meClick
      If m_Enabled Then DoTheClick
      Exit Sub ' A Click handler may hide or unload the control.
    Case Else
      Exit Sub
  End Select
  Refresh
End Sub

Private Sub DetachMouse()
  Dim Token As Long
  Token = m_MouseToken
  m_MouseToken = 0
  MouseUnregister Token
  m_Hovering = False
  m_MouseIsDown = False
  m_IsPressed = False
End Sub

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
  ' Conventional VB6 input in the IDE; compiled input uses IMouseSink only.
  If MouseNativeAvailable() Then Exit Sub
  If m_Enabled = False Then Exit Sub
  If (Button = vbLeftButton And m_MouseIsDown = False) Then
    m_MouseIsDown = True
    m_IsPressed = True
    Refresh
  End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If MouseNativeAvailable() Then Exit Sub
  If m_Enabled = False Then Exit Sub
  If (Button = vbLeftButton And m_MouseIsDown = True) Then
    m_MouseIsDown = False
    m_IsPressed = False
    Refresh
    If X >= 0 And Y >= 0 And X < UserControl.ScaleWidth And Y < UserControl.ScaleHeight Then DoTheClick
  End If
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  m_Caption = PropBag.ReadProperty(PROPNAME_CAPTION, "CheckBox")
  m_Value = PropBag.ReadProperty(PROPNAME_VALUE, vbUnchecked)
  m_Enabled = PropBag.ReadProperty(PROPNAME_ENABLED, True)
End Sub

Private Sub UserControl_Resize()
  If CheckWindowSize = True Then Refresh
  If m_MouseToken <> 0 Then MouseRefresh m_MouseToken
End Sub

Private Sub UserControl_Show()
  If UserControl.Ambient.UserMode And MouseNativeAvailable() Then
    If m_MouseToken = 0 Then
      m_MouseToken = MouseRegister(UserControl.hWnd, Me)
      If m_MouseToken = 0 Then Err.Raise vbObjectError + 2101, "CheckBox", MouseLastError()
      MouseSetEnabled m_MouseToken, m_Enabled
    End If
    MouseRefresh m_MouseToken
  End If
  Refresh
End Sub

Private Sub UserControl_Hide()
  DetachMouse
End Sub

Private Sub UserControl_Terminate()
  DetachMouse
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  PropBag.WriteProperty PROPNAME_CAPTION, m_Caption, "CheckBox"
  PropBag.WriteProperty PROPNAME_VALUE, m_Value, vbUnchecked
  PropBag.WriteProperty PROPNAME_ENABLED, m_Enabled, True
End Sub
