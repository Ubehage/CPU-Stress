VERSION 5.00
Begin VB.UserControl Button 
   AutoRedraw      =   -1  'True
   BackColor       =   &H002A2A2A&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   EditAtDesignTime=   -1  'True
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
Attribute VB_Name = "Button"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum Button_Style
  bsNormalDark = &H0
  bsRed = &H1
  bsGreen = &H2
  bsCustom = &H4
End Enum

Private Const PROPNAME_CAPTION As String = "Caption"
Private Const PROPNAME_ENABLED As String = "Enabled"
Private Const PROPNAME_BACKCOLOR As String = "BackColor"
Private Const PROPNAME_HOVERCOLOR As String = "HoverColor"
Private Const PROPNAME_PRESSEDCOLOR As String = "PressedColor"
Private Const PROPNAME_FONTNAME As String = "FontName"
Private Const PROPNAME_FONTSIZE As String = "FontSize"
Private Const PROPNAME_FONTBOLD As String = "FontBold"
Private Const PROPNAME_FORECOLOR As String = "ForeColor"
Private Const PROPNAME_DISABLEDBACKCOLOR As String = "DisabledBackColor"
Private Const PROPNAME_DISABLEDTEXTCOLOR As String = "DisabledTextColor"
Private Const PROPNAME_BUTTONSTYLE As String = "ButtonStyle"

Dim m_Caption As String
Dim m_Enabled As Boolean

Dim m_BackColor As Long
Dim m_HoverColor As Long
Dim m_PressedColor As Long
Dim m_ForeColor As Long
Dim m_DisabledBackColor As Long
Dim m_DisabledTextColor As Long
Dim m_ButtonStyle As Button_Style

Dim m_ScreenRect As RECT
Dim m_IsCapturing As Boolean
Dim m_Hovering As Boolean
Dim m_IsPressed As Boolean
Dim m_MouseIsDown As Boolean
Dim m_KeyIsDown As Boolean

Dim ButtonRect As RECT
Dim TextPos As POINTAPI

Public Event Click()

Public Property Get Caption() As String
  Caption = m_Caption
End Property
Public Property Get Enabled() As Boolean
  Enabled = m_Enabled
End Property
Public Property Get BackColor() As OLE_COLOR
  BackColor = m_BackColor
End Property
Public Property Get HoverColor() As OLE_COLOR
  HoverColor = m_HoverColor
End Property
Public Property Get PressedColor() As OLE_COLOR
  PressedColor = m_PressedColor
End Property
Public Property Get ForeColor() As OLE_COLOR
  ForeColor = m_ForeColor
End Property
Public Property Get DisabledBackColor() As OLE_COLOR
  DisabledBackColor = m_DisabledBackColor
End Property
Public Property Get DisabledTextColor() As OLE_COLOR
  DisabledTextColor = m_DisabledTextColor
End Property
Public Property Get ButtonStyle() As Button_Style
  ButtonStyle = m_ButtonStyle
End Property
Public Property Get FontName() As String
  FontName = UserControl.Font.Name
End Property
Public Property Get FontSize() As Integer
  FontSize = UserControl.Font.Size
End Property
Public Property Get FontBold() As Boolean
  FontBold = UserControl.Font.Bold
End Property
Public Property Let Caption(New_Caption As String)
  m_Caption = New_Caption
  UserControl.PropertyChanged PROPNAME_CAPTION
  SetTextPosition
  Refresh
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
Public Property Let BackColor(New_BackColor As OLE_COLOR)
  If m_BackColor = New_BackColor Then Exit Property
  m_BackColor = New_BackColor
  m_ButtonStyle = bsCustom
  UserControl.PropertyChanged PROPNAME_BUTTONSTYLE
  UserControl.PropertyChanged PROPNAME_BACKCOLOR
  If (m_Hovering = False And m_IsPressed = False) Then Refresh
End Property
Public Property Let HoverColor(New_HoverColor As OLE_COLOR)
  If m_HoverColor = New_HoverColor Then Exit Property
  m_HoverColor = New_HoverColor
  m_ButtonStyle = bsCustom
  UserControl.PropertyChanged PROPNAME_BUTTONSTYLE
  UserControl.PropertyChanged PROPNAME_HOVERCOLOR
  If m_Hovering = True Then Refresh
End Property
Public Property Let PressedColor(New_PressedColor As OLE_COLOR)
  If m_PressedColor = New_PressedColor Then Exit Property
  m_PressedColor = New_PressedColor
  m_ButtonStyle = bsCustom
  UserControl.PropertyChanged PROPNAME_BUTTONSTYLE
  UserControl.PropertyChanged PROPNAME_PRESSEDCOLOR
  If m_IsPressed = True Then Refresh
End Property
Public Property Let ForeColor(New_ForeColor As OLE_COLOR)
  If m_ForeColor = New_ForeColor Then Exit Property
  m_ForeColor = New_ForeColor
  m_ButtonStyle = bsCustom
  UserControl.PropertyChanged PROPNAME_BUTTONSTYLE
  UserControl.PropertyChanged PROPNAME_FORECOLOR
  If m_Enabled Then Refresh
End Property
Public Property Let DisabledBackColor(New_DisabledBackColor As OLE_COLOR)
  If m_DisabledBackColor = New_DisabledBackColor Then Exit Property
  m_DisabledBackColor = New_DisabledBackColor
  m_ButtonStyle = bsCustom
  UserControl.PropertyChanged PROPNAME_BUTTONSTYLE
  UserControl.PropertyChanged PROPNAME_DISABLEDBACKCOLOR
  If m_Enabled = False Then Refresh
End Property
Public Property Let DisabledTextColor(New_DisabledTextColor As OLE_COLOR)
  If m_DisabledTextColor = New_DisabledTextColor Then Exit Property
  m_DisabledTextColor = New_DisabledTextColor
  m_ButtonStyle = bsCustom
  UserControl.PropertyChanged PROPNAME_BUTTONSTYLE
  UserControl.PropertyChanged PROPNAME_DISABLEDTEXTCOLOR
  If m_Enabled = False Then Refresh
End Property
Public Property Let ButtonStyle(New_ButtonStyle As Button_Style)
  If m_ButtonStyle = New_ButtonStyle Then Exit Property
  m_ButtonStyle = New_ButtonStyle
  UserControl.PropertyChanged PROPNAME_BUTTONSTYLE
  SetButtonStyle
  Refresh
End Property
Public Property Let FontName(New_FontName As String)
  If New_FontName = UserControl.Font.Name Then Exit Property
  UserControl.Font.Name = New_FontName
  UserControl.PropertyChanged PROPNAME_FONTNAME
  Refresh True
End Property
Public Property Let FontSize(New_FontSize As Integer)
  If New_FontSize = UserControl.Font.Size Then Exit Property
  UserControl.Font.Size = New_FontSize
  UserControl.PropertyChanged PROPNAME_FONTSIZE
  Refresh True
End Property
Public Property Let FontBold(New_FontBold As Boolean)
  If New_FontBold = UserControl.Font.Bold Then Exit Property
  UserControl.Font.Bold = New_FontBold
  UserControl.PropertyChanged PROPNAME_FONTBOLD
  Refresh True
End Property

Public Sub Refresh(Optional FullRedraw As Boolean = False)
  If FullRedraw Then SetTextPosition
  ClearBackground
  DrawTitle
  DrawBorder
End Sub

Private Sub ClearBackground()
  With ButtonRect
    UserControl.Line (.Left, .Top)-(.Right, .Bottom), GetBackColor, BF
  End With
End Sub

Private Sub DrawTitle()
  UserControl.CurrentX = TextPos.X
  UserControl.CurrentY = TextPos.Y
  UserControl.ForeColor = GetTextColor
  UserControl.Print m_Caption
End Sub

Private Sub DrawBorder()
  Dim bColor As Long
  With ButtonRect
    UserControl.Line (.Left, .Top)-(.Right, .Bottom), COLOR_OUTLINE, B
    bColor = GetBackColor()
    UserControl.Line ((.Left + Screen.TwipsPerPixelX), (.Top + Screen.TwipsPerPixelY))-((.Right - Screen.TwipsPerPixelX), (.Bottom - Screen.TwipsPerPixelY)), bColor, B
    UserControl.Line ((.Left + (Screen.TwipsPerPixelX * 2)), (.Top + (Screen.TwipsPerPixelY * 2)))-((.Right - (Screen.TwipsPerPixelX * 2)), (.Bottom - (Screen.TwipsPerPixelY * 2))), bColor, B
  End With
End Sub

Private Sub SetButtonStyle()
  Dim s As Boolean
  If m_ButtonStyle = bsNormalDark Then
    m_BackColor = COLOR_CONTROLS
    m_ForeColor = COLOR_TEXT
    m_DisabledBackColor = COLOR_BACKGROUND_DISABLED
    m_DisabledTextColor = COLOR_TEXT_DISABLED
    m_HoverColor = COLOR_BUTTON_HOVER
    m_PressedColor = COLOR_BUTTON_PRESSED
    s = True
  ElseIf m_ButtonStyle = bsGreen Then
    m_BackColor = COLOR_GREEN
    m_ForeColor = COLOR_TEXT_ONGREEN
    m_DisabledBackColor = COLOR_GREEN_DISABLED
    m_DisabledTextColor = COLOR_TEXT_DISABLED_ONGREEN
    m_HoverColor = COLOR_GREEN_HOVER
    m_PressedColor = COLOR_GREEN_PRESSED
    s = True
  ElseIf m_ButtonStyle = bsRed Then
    m_BackColor = COLOR_RED
    m_ForeColor = COLOR_TEXT_ONRED
    m_DisabledBackColor = COLOR_RED_DISABLED
    m_DisabledTextColor = COLOR_TEXT_DISABLED_ONRED
    m_HoverColor = COLOR_RED_HOVER
    m_PressedColor = COLOR_RED_PRESSED
    s = True
  End If
  If s = True Then
    UserControl.PropertyChanged PROPNAME_BACKCOLOR
    UserControl.PropertyChanged PROPNAME_FORECOLOR
    UserControl.PropertyChanged PROPNAME_DISABLEDBACKCOLOR
    UserControl.PropertyChanged PROPNAME_DISABLEDTEXTCOLOR
    UserControl.PropertyChanged PROPNAME_HOVERCOLOR
    UserControl.PropertyChanged PROPNAME_PRESSEDCOLOR
  End If
End Sub

Private Sub SetTextPosition()
  TextPos.X = (UserControl.ScaleWidth - UserControl.TextWidth(m_Caption)) \ 2
  TextPos.Y = (UserControl.ScaleHeight - UserControl.TextHeight(m_Caption)) \ 2
End Sub

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

Private Sub SetButtonRect()
  With ButtonRect
    .Left = 0
    .Top = 0
    .Right = (UserControl.ScaleWidth - Screen.TwipsPerPixelX)
    .Bottom = (UserControl.ScaleHeight - Screen.TwipsPerPixelY)
  End With
End Sub

Private Function GetBackColor() As Long
  If m_Enabled = False Then
    GetBackColor = m_DisabledBackColor
  ElseIf m_IsPressed Then
    GetBackColor = m_PressedColor
  ElseIf m_Hovering = True Then
    GetBackColor = m_HoverColor
  Else
    GetBackColor = m_BackColor
  End If
End Function

Private Function GetTextColor() As Long
  If m_Enabled Then
    GetTextColor = m_ForeColor
  Else
    GetTextColor = m_DisabledTextColor
  End If
End Function

Private Sub StartHover(Optional DoNotRefresh As Boolean = False, Optional ForceRefresh As Boolean = False)
  Dim r As Boolean
  If m_Hovering = False Then
    m_Hovering = True
    If DoNotRefresh = False Then r = True
  End If
  If (r = True Or ForceRefresh = True) Then
    Refresh
    SetScreenRect
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

Private Sub DoClickEvent()
  EndCapture
  If m_Enabled Then RaiseEvent Click
End Sub

Private Function CanInteractNow() As Boolean
  CanInteractNow = UserControl.Ambient.UserMode And m_Enabled And UserControl.Extender.Visible And UserControl.hWnd <> 0
End Function

Private Sub UserControl_InitProperties()
  m_Caption = "Button"
  m_Enabled = True
  m_ButtonStyle = bsNormalDark
  SetButtonStyle
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
  If m_Enabled = False Then Exit Sub
  If (KeyCode = vbKeySpace Or KeyCode = vbKeyReturn) Then
    m_IsPressed = True
    m_KeyIsDown = True
    Refresh
  End If
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
  If m_Enabled = False Then Exit Sub
  If (KeyCode = vbKeySpace Or KeyCode = vbKeyReturn) Then
    If m_IsPressed Then
      m_IsPressed = False
      Refresh
      If m_KeyIsDown Then
        m_KeyIsDown = False
        DoClickEvent
      End If
    ElseIf m_KeyIsDown Then
      m_KeyIsDown = False
    End If
  End If
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If m_Enabled = False Then Exit Sub
  If Button = vbLeftButton Then
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
    Refresh
    EndHover True
    If IsCursorOnButton Then
      DoClickEvent
      If CanInteractNow() = True Then If HadCapture = True Then StartHover True
    End If
    Refresh
  End If
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  m_Caption = PropBag.ReadProperty(PROPNAME_CAPTION, "Button")
  m_Enabled = PropBag.ReadProperty(PROPNAME_ENABLED, True)
  m_BackColor = PropBag.ReadProperty(PROPNAME_BACKCOLOR, COLOR_CONTROLS)
  m_HoverColor = PropBag.ReadProperty(PROPNAME_HOVERCOLOR, COLOR_BUTTON_HOVER)
  m_PressedColor = PropBag.ReadProperty(PROPNAME_PRESSEDCOLOR, COLOR_BUTTON_PRESSED)
  m_ForeColor = PropBag.ReadProperty(PROPNAME_FORECOLOR, COLOR_TEXT)
  m_DisabledBackColor = PropBag.ReadProperty(PROPNAME_DISABLEDBACKCOLOR, COLOR_BACKGROUND_DISABLED)
  m_DisabledTextColor = PropBag.ReadProperty(PROPNAME_DISABLEDTEXTCOLOR, COLOR_TEXT_DISABLED)
  m_ButtonStyle = PropBag.ReadProperty(PROPNAME_BUTTONSTYLE, Button_Style.bsNormalDark)
  FontName = PropBag.ReadProperty(PROPNAME_FONTNAME, UserControl.Ambient.Font.Name)
  FontSize = PropBag.ReadProperty(PROPNAME_FONTSIZE, UserControl.Ambient.Font.Size)
  FontBold = PropBag.ReadProperty(PROPNAME_FONTBOLD, UserControl.Ambient.Font.Bold)
  SetButtonStyle
End Sub

Private Sub UserControl_Resize()
  SetScreenRect
  SetButtonRect
  Refresh True
End Sub

Private Sub UserControl_Show()
  Refresh
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  PropBag.WriteProperty PROPNAME_CAPTION, m_Caption, "Button"
  PropBag.WriteProperty PROPNAME_ENABLED, m_Enabled, True
  PropBag.WriteProperty PROPNAME_BACKCOLOR, m_BackColor, COLOR_CONTROLS
  PropBag.WriteProperty PROPNAME_HOVERCOLOR, m_HoverColor, COLOR_BUTTON_HOVER
  PropBag.WriteProperty PROPNAME_PRESSEDCOLOR, m_PressedColor, COLOR_BUTTON_PRESSED
  PropBag.WriteProperty PROPNAME_FORECOLOR, m_ForeColor, COLOR_TEXT
  PropBag.WriteProperty PROPNAME_DISABLEDBACKCOLOR, m_DisabledBackColor, COLOR_BACKGROUND_DISABLED
  PropBag.WriteProperty PROPNAME_DISABLEDTEXTCOLOR, m_DisabledTextColor, COLOR_TEXT_DISABLED
  PropBag.WriteProperty PROPNAME_BUTTONSTYLE, m_ButtonStyle, Button_Style.bsNormalDark
  PropBag.WriteProperty PROPNAME_FONTNAME, UserControl.Font.Name, UserControl.Ambient.Font.Name
  PropBag.WriteProperty PROPNAME_FONTSIZE, UserControl.Font.Size, UserControl.Ambient.Font.Size
  PropBag.WriteProperty PROPNAME_FONTBOLD, UserControl.Font.Bold, UserControl.Ambient.Font.Bold
End Sub
