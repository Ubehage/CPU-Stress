VERSION 5.00
Begin VB.UserControl StatusBar 
   Alignable       =   -1  'True
   AutoRedraw      =   -1  'True
   BackColor       =   &H00202020&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   EditAtDesignTime=   -1  'True
   BeginProperty Font 
      Name            =   "Consolas"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00E0E0E0&
   ScaleHeight     =   3600
   ScaleWidth      =   4800
End
Attribute VB_Name = "StatusBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const PROPNAME_TEXT = "Text"
Private Const PROP_TEXT_DEFAULT = "StatusBar"

Dim m_Text As String

Public Property Get Text() As String
  Text = m_Text
End Property
Public Property Let Text(New_Text As String)
  If m_Text = New_Text Then Exit Property
  m_Text = New_Text
  Refresh
End Property

Public Sub Refresh()
  UserControl.Cls
  DrawSeperatorLine
  DrawStatusText
End Sub

Private Sub DrawSeperatorLine()
  UserControl.Line (0, 0)-(UserControl.ScaleWidth, 0), COLOR_OUTLINE
  UserControl.Line (0, Screen.TwipsPerPixelY)-(UserControl.ScaleWidth, Screen.TwipsPerPixelY), COLOR_OUTLINE_LIGHT
End Sub

Private Sub DrawStatusText()
  UserControl.CurrentX = (Screen.TwipsPerPixelX * 5)
  UserControl.CurrentY = (Screen.TwipsPerPixelY * 4)
  UserControl.Print m_Text
End Sub

Private Function GetStatusBarHeight() As Long
  Dim h As Long
  h = (UserControl.Height - UserControl.ScaleHeight)
  GetStatusBarHeight = h + (UserControl.TextHeight(m_Text) + (Screen.TwipsPerPixelY * 7))
End Function

Private Sub UserControl_InitProperties()
  m_Text = PROP_TEXT_DEFAULT
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  m_Text = PropBag.ReadProperty(PROPNAME_TEXT, PROP_TEXT_DEFAULT)
End Sub

Private Sub UserControl_Resize()
  If UserControl.Height <> GetStatusBarHeight Then UserControl.Height = GetStatusBarHeight: Exit Sub
  Refresh
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  PropBag.WriteProperty PROPNAME_TEXT, m_Text, PROP_TEXT_DEFAULT
End Sub
