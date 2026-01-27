VERSION 5.00
Begin VB.UserControl Frame 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00202020&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ControlContainer=   -1  'True
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
Attribute VB_Name = "Frame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const PROPNAME_CAPTION As String = "Caption"
Private Const PROP_DEFAULT_CAPTION As String = "Frame"

Dim m_Caption As String

Public Property Get Caption() As String
  Caption = m_Caption
End Property
Public Property Let Caption(New_Caption As String)
  If m_Caption = New_Caption Then Exit Property
  m_Caption = New_Caption
  Refresh
End Property

Public Sub Refresh()
  UserControl.Cls
  DrawBorder
  If Not m_Caption = "" Then DrawTitle
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
  UserControl.Print m_Caption
End Sub

Private Sub UserControl_InitProperties()
  m_Caption = PROP_DEFAULT_CAPTION
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  m_Caption = PropBag.ReadProperty(PROPNAME_CAPTION, PROP_DEFAULT_CAPTION)
End Sub

Private Sub UserControl_Resize()
  Refresh
End Sub

Private Sub UserControl_Show()
  Refresh
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  PropBag.WriteProperty PROPNAME_CAPTION, m_Caption, PROP_DEFAULT_CAPTION
End Sub
