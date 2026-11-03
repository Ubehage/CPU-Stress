VERSION 5.00
Begin VB.Form frmMouseLifetime
   Caption         =   "Mouse recipient lifetime tests"
   ClientHeight    =   3900
   ClientWidth     =   6000
   StartUpPosition =   2
   ScaleMode       =   3
   Begin MouseTests.CheckBox CheckControl
      Height          =   450
      Left            =   150
      Top             =   150
      Width           =   2400
      Caption         =   "Unload checkbox"
   End
   Begin MouseTests.CPUView CPUControl
      Height          =   2700
      Left            =   150
      Top             =   750
      Width           =   5400
      AutoUpdate      =   0
   End
End
Attribute VB_Name = "frmMouseLifetime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public UnloadCheckOnClick As Boolean
Public UnloadCPUOnClick As Boolean

Private Sub CheckControl_Click()
  RecordLifetimeCheckClick
  If UnloadCheckOnClick Then Unload Me
End Sub

Private Sub CPUControl_Click(Index As Long)
  RecordLifetimeCPUClick Index
  If UnloadCPUOnClick Then Unload Me
End Sub

' Deliberately no MouseShutdown here: unloading this form must detach only
' its children, preserving recipients registered by the main test form.
